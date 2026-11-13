/// Excel storage contract for formulas.
///
/// Formulas are shown to users in display form and written into .xlsx files
/// in storage form. The two differ in these ways:
///   1. Post-2007 functions carry `_xlfn.` (FILTER and SORT: `_xlfn._xlws.`).
///   2. LET names and LAMBDA parameters carry `_xlpm.` and never `_xlfn.`; an
///      optional parameter `[y]` is stored `_xlop.y`. A call `name(` is the
///      built-in of that name when there is one, even if a variable has the
///      same name: `LET(date,5,DATE(2020,1,date))` keeps `DATE(`.
///   3. The implicit-intersection operator `@x` is stored as
///      `_xlfn.SINGLE(x)`.
///   4. The spill operator `ref#` is stored as `_xlfn.ANCHORARRAY(ref)`.
///   5. A bare function name given to BYROW, BYCOL, MAP, REDUCE, SCAN,
///      MAKEARRAY, GROUPBY or PIVOTBY (or bound by LET and then called) is
///      stored `_xleta.SUM`.
///   6. The table shorthand `[@Col]` and `Table1[@[Col]]` is stored
///      `[[#This Row],[Col]]` and `Table1[[#This Row],[Col]]`.
///
/// Excel opens a file whose formulas miss these markers with #NAME? (or a
/// repair prompt), so this module is the one conversion point for the writers.
///
/// `toStorageFormula` is idempotent (to(to(f)) === to(f)) and leaves anything
/// it does not understand byte-identical. `fromStorageFormula` is the inverse
/// for the display grammar this module emits.
///
/// Parsing is a small tokenizer plus nesting, not an expression grammar: it
/// recognises strings, quoted sheet names, bracket groups, array constants,
/// identifiers, numbers, error literals, calls and parenthesised groups, and
/// keeps every other character as-is. Unbalanced input never throws; it is
/// echoed back unchanged, and so is input nested deeply enough to overflow the
/// call stack (reported as UNSUPPORTED_SYNTAX).

import {
  DEFAULT_FUNCTION_REGISTRY,
  type FunctionRegistry,
  type FunctionSpec,
} from "./function-registry.ts";

export interface StorageIssue {
  /// UNKNOWN_FUNCTION: call with no storage rule in the registry (a classic
  ///   function, a defined name or a named LAMBDA); left unchanged.
  /// NOT_EXCEL: a Google Sheets only function; left unchanged.
  /// UNVERIFIED_PREFIX: a function whose marker is not confirmed; unchanged.
  /// UNSUPPORTED_SYNTAX: a LET name or LAMBDA parameter that is not a plain
  ///   identifier or `[identifier]`, or input nested too deeply to convert
  ///   (name NESTING); unchanged.
  readonly code: "UNKNOWN_FUNCTION" | "NOT_EXCEL" | "UNVERIFIED_PREFIX" | "UNSUPPORTED_SYNTAX";
  readonly name: string;
}

export interface StorageResult {
  /// Body without the leading `=`, as written to `<f>`.
  readonly text: string;
  /// True when the whole formula is one call whose result is always an array.
  readonly dynamic: boolean;
  readonly issues: readonly StorageIssue[];
}

const XLPM = "_xlpm.";
const XLOP = "_xlop.";
const XLETA = "_xleta.";
const XLFN = "_xlfn.";
const XLFN_XLWS = "_xlfn._xlws.";

// ---------------------------------------------------------------------------
// Tree

type TokKind = "ws" | "op" | "num" | "str" | "err" | "arr" | "bracket" | "qsheet";

type Node =
  | { readonly t: "tok"; readonly k: TokKind; readonly s: string }
  | { readonly t: "name"; readonly s: string }
  | {
      readonly t: "call";
      readonly name: string;
      readonly ws: string;
      readonly args: readonly (readonly Node[])[];
      readonly closed: boolean;
      /// Synthetic call built from `@`/`#`: its name is final.
      readonly fixed: boolean;
    }
  | {
      readonly t: "group";
      readonly args: readonly (readonly Node[])[];
      readonly closed: boolean;
    }
  | { readonly t: "at" }
  | { readonly t: "spill" };

type CallNode = Extract<Node, { t: "call" }>;

const WS_RE = /\s+/y;
const NUM_RE = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const IDENT_RE = /[\p{L}_\\$][\p{L}\p{N}_.\\$]*/uy;
const ERR_RE =
  /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!|EXTERNAL!)/iy;
/// A LET name / LAMBDA parameter / bound reference: a plain identifier.
const SIMPLE_NAME_RE = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u;

function matchAt(re: RegExp, src: string, index: number): string | undefined {
  re.lastIndex = index;
  const m = re.exec(src);
  return m ? m[0] : undefined;
}

/// End (exclusive) of a `quote`-delimited run starting at `start`, where a
/// doubled quote is an escaped quote. Unterminated runs end at the input end.
function scanQuoted(src: string, start: number, quote: string): number {
  let j = start + 1;
  while (j < src.length) {
    if (src[j] === quote) {
      if (src[j + 1] === quote) {
        j += 2;
        continue;
      }
      return j + 1;
    }
    j++;
  }
  return src.length;
}

/// End (exclusive) of a (possibly nested) `[...]` group. Inside table
/// references a `'` escapes the next character.
function scanBracket(src: string, start: number): number {
  let depth = 0;
  let j = start;
  while (j < src.length) {
    const c = src[j];
    if (c === "'") {
      j += 2;
      continue;
    }
    if (c === "[") depth++;
    else if (c === "]") {
      depth--;
      if (depth === 0) return j + 1;
    }
    j++;
  }
  return src.length;
}

/// End (exclusive) of a `{...}` array constant, skipping string literals.
function scanArray(src: string, start: number): number {
  let j = start;
  while (j < src.length) {
    const c = src[j];
    if (c === '"') {
      j = scanQuoted(src, j, '"');
      continue;
    }
    if (c === "}") return j + 1;
    j++;
  }
  return src.length;
}

interface ParseState {
  readonly src: string;
  i: number;
}

function isOperandEnd(node: Node | undefined): boolean {
  if (!node) return false;
  return (
    node.t === "name" ||
    node.t === "call" ||
    node.t === "group" ||
    node.t === "spill" ||
    (node.t === "tok" && node.k === "bracket")
  );
}

/// Parses one level. Nested levels split on `,` and end at `)`; the top level
/// treats both as plain characters.
function parseLevel(st: ParseState, nested: boolean): { args: Node[][]; closed: boolean } {
  const { src } = st;
  const args: Node[][] = [[]];
  let current = args[0] as Node[];
  const push = (node: Node): void => {
    current.push(node);
  };
  while (st.i < src.length) {
    const c = src.charAt(st.i);
    if (nested && c === ")") {
      st.i++;
      return { args, closed: true };
    }
    if (nested && c === ",") {
      st.i++;
      current = [];
      args.push(current);
      continue;
    }
    const ws = matchAt(WS_RE, src, st.i);
    if (ws !== undefined) {
      push({ t: "tok", k: "ws", s: ws });
      st.i += ws.length;
      continue;
    }
    if (c === '"') {
      const end = scanQuoted(src, st.i, '"');
      push({ t: "tok", k: "str", s: src.slice(st.i, end) });
      st.i = end;
      continue;
    }
    if (c === "'") {
      const end = scanQuoted(src, st.i, "'");
      push({ t: "tok", k: "qsheet", s: src.slice(st.i, end) });
      st.i = end;
      continue;
    }
    if (c === "[") {
      const end = scanBracket(src, st.i);
      push({ t: "tok", k: "bracket", s: src.slice(st.i, end) });
      st.i = end;
      continue;
    }
    if (c === "{") {
      const end = scanArray(src, st.i);
      push({ t: "tok", k: "arr", s: src.slice(st.i, end) });
      st.i = end;
      continue;
    }
    if (c === "(") {
      st.i++;
      const inner = parseLevel(st, true);
      push({ t: "group", args: inner.args, closed: inner.closed });
      continue;
    }
    if (c === "@") {
      push({ t: "at" });
      st.i++;
      continue;
    }
    if (c === "#") {
      // An error literal wins: a spill marker is never followed by one.
      const err = matchAt(ERR_RE, src, st.i);
      if (err !== undefined) {
        push({ t: "tok", k: "err", s: err });
        st.i += err.length;
        continue;
      }
      if (isOperandEnd(current[current.length - 1])) {
        push({ t: "spill" });
        st.i++;
        continue;
      }
      push({ t: "tok", k: "op", s: "#" });
      st.i++;
      continue;
    }
    const num = matchAt(NUM_RE, src, st.i);
    if (num !== undefined) {
      push({ t: "tok", k: "num", s: num });
      st.i += num.length;
      continue;
    }
    const ident = matchAt(IDENT_RE, src, st.i);
    if (ident !== undefined) {
      let end = st.i + ident.length;
      let merged = false;
      // Table references: Table1[Col], Table1[[#This Row],[Col]].
      while (src.charAt(end) === "[") {
        end = scanBracket(src, end);
        merged = true;
      }
      if (!merged) {
        const gap = matchAt(WS_RE, src, end) ?? "";
        if (src.charAt(end + gap.length) === "(") {
          st.i = end + gap.length + 1;
          const inner = parseLevel(st, true);
          push({
            t: "call",
            name: ident,
            ws: gap,
            args: inner.args,
            closed: inner.closed,
            fixed: false,
          });
          continue;
        }
      }
      push({ t: "name", s: src.slice(st.i, end) });
      st.i = end;
      continue;
    }
    const cp = String.fromCodePoint(src.codePointAt(st.i) ?? 0);
    push({ t: "tok", k: "op", s: cp });
    st.i += cp.length;
  }
  return { args, closed: false };
}

function parseFormula(src: string): Node[] {
  return parseLevel({ src, i: 0 }, false).args[0] as Node[];
}

// ---------------------------------------------------------------------------
// Operand scanning shared by both directions

function isCore(node: Node | undefined): boolean {
  if (!node) return false;
  return (
    node.t === "name" ||
    node.t === "call" ||
    node.t === "group" ||
    (node.t === "tok" && node.k === "bracket")
  );
}

function isSheetName(node: Node | undefined): boolean {
  return !!node && (node.t === "name" || (node.t === "tok" && node.k === "qsheet"));
}

function isTok(node: Node | undefined, text: string): boolean {
  return !!node && node.t === "tok" && node.s === text;
}

function isBracketTok(node: Node | undefined): boolean {
  return !!node && node.t === "tok" && node.k === "bracket";
}

/// End (exclusive) of one reference primary starting at `i`: an optional
/// sheet prefix (`Sheet1!`, `'My Sheet'!`, `[1]Sheet1!`) and a name, call,
/// parenthesised group or table-bracket. Returns `i` when there is none.
function primaryEnd(list: readonly Node[], i: number): number {
  let j = i;
  if (isBracketTok(list[i]) && isSheetName(list[i + 1]) && isTok(list[i + 2], "!")) j = i + 3;
  else if (isSheetName(list[i]) && isTok(list[i + 1], "!")) j = i + 2;
  return isCore(list[j]) ? j + 1 : i;
}

/// End (exclusive) of the operand of `@`: a primary and any `:` chained
/// primaries. Returns `i` when there is no operand.
function rangeEnd(list: readonly Node[], i: number): number {
  let j = primaryEnd(list, i);
  if (j === i) return i;
  while (isTok(list[j], ":")) {
    const k = primaryEnd(list, j + 1);
    if (k === j + 1) break;
    j = k;
  }
  return j;
}

/// Start of the primary that ends right before index `spill`, or -1.
function primaryStart(list: readonly Node[], spill: number): number {
  const core = spill - 1;
  if (!isCore(list[core])) return -1;
  let start = core;
  if (isTok(list[core - 1], "!") && isSheetName(list[core - 2])) {
    start = core - 2;
    if (isBracketTok(list[core - 3])) start = core - 3;
  }
  return start;
}

function syntheticCall(name: string, args: readonly (readonly Node[])[]): CallNode {
  return { t: "call", name, ws: "", args, closed: true, fixed: true };
}

/// Rewrites the `#` and `@` operators of one level into synthetic calls.
/// Spills resolve first so `@D2#` becomes SINGLE(ANCHORARRAY(D2)). A stray
/// operator without an operand is kept as a plain character.
function resolveOperators(nodes: readonly Node[]): readonly Node[] {
  if (!nodes.some((n) => n.t === "at" || n.t === "spill")) return nodes;
  const list: Node[] = [...nodes];
  for (let j = 0; j < list.length; j++) {
    if (list[j]?.t !== "spill") continue;
    const start = primaryStart(list, j);
    if (start < 0) {
      list[j] = { t: "tok", k: "op", s: "#" };
      continue;
    }
    const operand = list.slice(start, j);
    list.splice(start, j - start + 1, syntheticCall(`${XLFN}ANCHORARRAY`, [operand]));
    j = start;
  }
  for (let j = list.length - 1; j >= 0; j--) {
    if (list[j]?.t !== "at") continue;
    const end = rangeEnd(list, j + 1);
    if (end <= j + 1) {
      list[j] = { t: "tok", k: "op", s: "@" };
      continue;
    }
    const operand = list.slice(j + 1, end);
    const only = operand.length === 1 ? operand[0] : undefined;
    // Excel stores @(a+b) as SINGLE((a+b)); SINGLE(a+b) loads to the same value
    // (checked in Excel), so the group's own parentheses are not repeated.
    const args =
      only?.t === "group" && only.closed && only.args.length === 1 ? only.args : [operand];
    list.splice(j, end - j, syntheticCall(`${XLFN}SINGLE`, args));
  }
  return list;
}

// ---------------------------------------------------------------------------
// Display -> storage

type Bound = ReadonlySet<string>;

interface Ctx {
  readonly registry: FunctionRegistry;
  readonly issues: StorageIssue[];
  readonly seen: Set<string>;
}

function report(ctx: Ctx, code: StorageIssue["code"], name: string): void {
  const key = `${code}:${name.toUpperCase()}`;
  if (ctx.seen.has(key)) return;
  ctx.seen.add(key);
  ctx.issues.push({ code, name });
}

function startsWithFold(text: string, prefix: string): boolean {
  return text.slice(0, prefix.length).toLowerCase() === prefix;
}

/// Splits a call name into its storage marker and bare name.
function splitFunctionPrefix(name: string): { prefix: string; base: string } {
  if (startsWithFold(name, XLFN_XLWS)) {
    return { prefix: name.slice(0, XLFN_XLWS.length), base: name.slice(XLFN_XLWS.length) };
  }
  if (startsWithFold(name, XLFN)) {
    return { prefix: name.slice(0, XLFN.length), base: name.slice(XLFN.length) };
  }
  return { prefix: "", base: name };
}

/// `[@Col]`, `[@[Col]]` and `[@[Col1]:[Col2]]` are display shorthand for the
/// current table row; Excel stores `[[#This Row],[Col]]` (`[#This Row]` for
/// `[@]`), with the table name in front when the reference has one.
function thisRowGroup(group: string): string {
  if (!group.startsWith("[@") || scanBracket(group, 0) !== group.length) return group;
  const rest = group.slice(2);
  if (rest === "]") return "[#This Row]";
  return rest.startsWith("[") ? `[[#This Row],${rest}` : `[[#This Row],[${rest}]`;
}

function rewriteTableRef(text: string): string {
  const at = text.indexOf("[");
  return at < 0 ? text : text.slice(0, at) + thisRowGroup(text.slice(at));
}

function emitName(text: string, skip: boolean, bound: Bound): string {
  if (skip || startsWithFold(text, XLPM)) return text;
  if (!SIMPLE_NAME_RE.test(text)) return rewriteTableRef(text);
  return bound.has(text.toUpperCase()) ? XLPM + text : text;
}

/// An argument made of one bare name and optional surrounding whitespace.
function loneName(arg: readonly Node[]): { lead: string; name: string; trail: string } | undefined {
  let lead = "";
  let trail = "";
  let name: string | undefined;
  for (const node of arg) {
    if (node.t === "tok" && node.k === "ws") {
      if (name === undefined) lead += node.s;
      else trail += node.s;
      continue;
    }
    if (node.t === "name" && name === undefined) {
      name = node.s;
      continue;
    }
    return undefined;
  }
  return name === undefined ? undefined : { lead, name, trail };
}

/// `_xleta.NAME` for a bare identifier that Excel reads as a function value: a
/// built-in function name that is not a bound variable and not the boolean
/// literal TRUE or FALSE. A defined name spelled like a function cannot be told
/// apart from the text alone, so callers only ask where a function is expected.
function etaName(text: string, bound: Bound, registry: FunctionRegistry): string | undefined {
  if (!SIMPLE_NAME_RE.test(text) || startsWithFold(text, "_xl")) return undefined;
  const upper = text.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE" || bound.has(upper)) return undefined;
  return registry.isBuiltin(upper) ? XLETA + upper : undefined;
}

/// Functions that build a vector of functions: `GROUPBY(r,v,HSTACK(SUM,COUNT))`.
const FUNCTION_VECTORS: ReadonlySet<string> = new Set(["HSTACK", "VSTACK"]);

/// An argument that takes a LAMBDA: a lone built-in function name becomes
/// `_xleta.NAME`, and so does each argument of an HSTACK/VSTACK of them.
function emitFunctionArg(arg: readonly Node[], bound: Bound, ctx: Ctx): string {
  const lone = loneName(arg);
  if (lone) {
    const eta = etaName(lone.name, bound, ctx.registry);
    return eta === undefined ? emitList(arg, bound, ctx) : lone.lead + eta + lone.trail;
  }
  const core = arg.filter((node) => !(node.t === "tok" && node.k === "ws"));
  const only = core.length === 1 ? core[0] : undefined;
  if (
    only?.t === "call" &&
    !only.fixed &&
    only.ws === "" &&
    FUNCTION_VECTORS.has(splitFunctionPrefix(only.name).base.toUpperCase())
  ) {
    const at = arg.indexOf(only);
    const pad = (nodes: readonly Node[]): string =>
      nodes.map((node) => (node.t === "tok" ? node.s : "")).join("");
    return pad(arg.slice(0, at)) + emitCall(only, false, bound, ctx, true) + pad(arg.slice(at + 1));
  }
  return emitList(arg, bound, ctx);
}

function emitSlotArgs(
  args: readonly (readonly Node[])[],
  slot: number | "last",
  bound: Bound,
  ctx: Ctx,
): string {
  const at = slot === "last" ? args.length - 1 : slot;
  return args
    .map((arg, index) =>
      index === at && index > 0 ? emitFunctionArg(arg, bound, ctx) : emitList(arg, bound, ctx),
    )
    .join(",");
}

/// For each name, the last argument index at which it is used as a function:
/// called (`f(x)`) or placed in the function slot of BYROW, MAP, ...
function functionUses(
  args: readonly (readonly Node[])[],
  registry: FunctionRegistry,
): Map<string, number> {
  const last = new Map<string, number>();
  const visit = (nodes: readonly Node[], index: number): void => {
    for (const node of nodes) {
      if (node.t === "group") {
        for (const inner of node.args) visit(inner, index);
      } else if (node.t === "call") {
        if (!node.fixed && SIMPLE_NAME_RE.test(node.name)) last.set(node.name.toUpperCase(), index);
        const slot = registry.get(splitFunctionPrefix(node.name).base)?.lambdaArg;
        const at = slot === "last" ? node.args.length - 1 : slot;
        node.args.forEach((inner, position) => {
          visit(inner, index);
          const lone = position === at && position > 0 ? loneName(inner) : undefined;
          if (lone) last.set(lone.name.toUpperCase(), index);
        });
      }
    }
  };
  args.forEach((arg, index) => visit(arg, index));
  return last;
}

function emitArgs(args: readonly (readonly Node[])[], bound: Bound, ctx: Ctx): string {
  return args.map((arg) => emitList(arg, bound, ctx)).join(",");
}

/// A LET name / LAMBDA parameter argument: one plain identifier, optionally
/// padded by whitespace.
function readBinding(
  arg: readonly Node[],
): { lead: string; name: string; trail: string; key: string } | undefined {
  const lone = loneName(arg);
  if (!lone) return undefined;
  const { lead, name, trail } = lone;
  // Already stored: `_xlpm.x` (LET, LAMBDA) or `_xlop.y` (optional parameter).
  const marker = startsWithFold(name, XLPM) ? XLPM : startsWithFold(name, XLOP) ? XLOP : "";
  const bare = name.slice(marker.length);
  if (!SIMPLE_NAME_RE.test(bare)) return undefined;
  return { lead, name: marker === "" ? XLPM + name : name, trail, key: bare.toUpperCase() };
}

/// An optional LAMBDA parameter `[y]`, stored `_xlop.y`; its uses stay `_xlpm.y`.
function readOptionalBinding(
  arg: readonly Node[],
): { lead: string; name: string; trail: string; key: string } | undefined {
  let lead = "";
  let trail = "";
  let group: string | undefined;
  for (const node of arg) {
    if (node.t === "tok" && node.k === "ws") {
      if (group === undefined) lead += node.s;
      else trail += node.s;
      continue;
    }
    if (node.t === "tok" && node.k === "bracket" && group === undefined) {
      group = node.s;
      continue;
    }
    return undefined;
  }
  if (group === undefined || group.length < 3 || !group.endsWith("]")) return undefined;
  const bare = group.slice(1, -1);
  if (!SIMPLE_NAME_RE.test(bare)) return undefined;
  return { lead, name: XLOP + bare, trail, key: bare.toUpperCase() };
}

function withName(bound: Bound, key: string): Bound {
  const next = new Set(bound);
  next.add(key);
  return next;
}

/// LET(n1, v1, n2, v2, ..., calc): each name is visible to the arguments
/// after it, never to its own value.
function emitLetArgs(args: readonly (readonly Node[])[], bound: Bound, ctx: Ctx): string {
  const parts: string[] = [];
  let scope = bound;
  // A name becomes visible after its own value has been emitted.
  let pending: string | undefined;
  let uses: Map<string, number> | undefined;
  args.forEach((arg, index) => {
    if (index % 2 === 0 && index < args.length - 1) {
      const binding = readBinding(arg);
      if (binding) {
        parts.push(binding.lead + binding.name + binding.trail);
        pending = binding.key;
        return;
      }
      report(ctx, "UNSUPPORTED_SYNTAX", "LET");
    }
    // `LET(f,SUM,f(x))`: a bound bare function name is stored `_xleta.SUM`. Only
    // when the name is then used as a function, so a defined name that happens
    // to be spelled like one (`LET(r,Rate,r*2)`) stays as written.
    let text: string | undefined;
    const lone = pending === undefined ? undefined : loneName(arg);
    if (pending !== undefined && lone) {
      const eta = etaName(lone.name, scope, ctx.registry);
      if (eta !== undefined) {
        uses ??= functionUses(args, ctx.registry);
        if ((uses.get(pending) ?? -1) > index) text = lone.lead + eta + lone.trail;
      }
    }
    parts.push(text ?? emitList(arg, scope, ctx));
    if (pending !== undefined) {
      scope = withName(scope, pending);
      pending = undefined;
    }
  });
  return parts.join(",");
}

/// LAMBDA(p1, p2, ..., body): the parameters are visible to the body only.
function emitLambdaArgs(args: readonly (readonly Node[])[], bound: Bound, ctx: Ctx): string {
  const parts: string[] = [];
  let scope = bound;
  const last = args.length - 1;
  args.forEach((arg, index) => {
    if (index === last) return;
    const binding = readBinding(arg) ?? readOptionalBinding(arg);
    if (binding) {
      parts.push(binding.lead + binding.name + binding.trail);
      scope = withName(scope, binding.key);
      return;
    }
    report(ctx, "UNSUPPORTED_SYNTAX", "LAMBDA");
    parts.push(emitList(arg, bound, ctx));
  });
  const body = args[last];
  parts.push(body ? emitList(body, scope, ctx) : "");
  return parts.join(",");
}

function emitCall(
  node: CallNode,
  afterBang: boolean,
  bound: Bound,
  ctx: Ctx,
  inFunctionSlot = false,
): string {
  const closer = node.closed ? ")" : "";
  const open = `${node.ws}(`;
  if (node.fixed || afterBang) {
    return node.name + open + emitArgs(node.args, bound, ctx) + closer;
  }
  const { name } = node;
  // A variable named like a built-in is still shadowed by that built-in in call
  // position `name(`; with a gap before the paren it is the variable followed
  // by an intersection.
  const variable = SIMPLE_NAME_RE.test(name) && bound.has(name.toUpperCase());
  if (
    startsWithFold(name, XLPM) ||
    (variable && !(node.ws === "" && ctx.registry.isBuiltin(name)))
  ) {
    const head = startsWithFold(name, XLPM) ? name : XLPM + name;
    return head + open + emitArgs(node.args, bound, ctx) + closer;
  }
  const { prefix, base } = splitFunctionPrefix(name);
  let head = name;
  let spec: FunctionSpec | undefined;
  if (prefix !== "") {
    spec = ctx.registry.get(base);
  } else {
    spec = ctx.registry.get(name);
    if (!spec) report(ctx, "UNKNOWN_FUNCTION", name);
    else if (spec.excel === "not-excel") report(ctx, "NOT_EXCEL", spec.name);
    else if (!spec.verified) report(ctx, "UNVERIFIED_PREFIX", spec.name);
    else head = spec.storage + spec.name;
  }
  let body: string;
  if (spec?.scope === "LET") body = emitLetArgs(node.args, bound, ctx);
  else if (spec?.scope === "LAMBDA") body = emitLambdaArgs(node.args, bound, ctx);
  else if (inFunctionSlot) {
    body = node.args.map((arg) => emitFunctionArg(arg, bound, ctx)).join(",");
  } else if (spec?.lambdaArg !== undefined) {
    body = emitSlotArgs(node.args, spec.lambdaArg, bound, ctx);
  } else body = emitArgs(node.args, bound, ctx);
  return head + open + body + closer;
}

function emitList(nodes: readonly Node[], bound: Bound, ctx: Ctx): string {
  const list = resolveOperators(nodes);
  let out = "";
  for (let i = 0; i < list.length; i++) {
    const node = list[i] as Node;
    const afterBang = isTok(list[i - 1], "!");
    switch (node.t) {
      case "tok":
        out += node.k === "bracket" ? rewriteTableRef(node.s) : node.s;
        break;
      case "name":
        out += emitName(node.s, afterBang || isTok(list[i + 1], "!"), bound);
        break;
      case "group":
        out += `(${emitArgs(node.args, bound, ctx)}${node.closed ? ")" : ""}`;
        break;
      case "call":
        out += emitCall(node, afterBang, bound, ctx);
        break;
      case "at":
        out += "@";
        break;
      case "spill":
        out += "#";
        break;
    }
  }
  return out;
}

function isDynamicTopLevel(nodes: readonly Node[], registry: FunctionRegistry): boolean {
  const top = resolveOperators(nodes).filter((n) => !(n.t === "tok" && n.k === "ws"));
  const only = top.length === 1 ? top[0] : undefined;
  if (only?.t !== "call") return false;
  return registry.get(splitFunctionPrefix(only.name).base)?.returnsArray === "always";
}

/// Converts a display formula (with or without the leading `=`) to the text
/// stored in `<f>`. Pure and idempotent.
export function toStorageFormula(
  display: string,
  registry: FunctionRegistry = DEFAULT_FUNCTION_REGISTRY,
): StorageResult {
  const body = display.startsWith("=") ? display.slice(1) : display;
  // No call, group or operator to rewrite: nothing can change.
  if (!/[(@#]/.test(body)) return { text: body, dynamic: false, issues: [] };
  try {
    const nodes = parseFormula(body);
    const ctx: Ctx = { registry, issues: [], seen: new Set() };
    const text = emitList(nodes, new Set(), ctx);
    return { text, dynamic: isDynamicTopLevel(nodes, registry), issues: ctx.issues };
  } catch (error) {
    // Nesting deep enough to overflow the call stack: echo the input back.
    const name = error instanceof RangeError ? "NESTING" : "FORMULA";
    return { text: body, dynamic: false, issues: [{ code: "UNSUPPORTED_SYNTAX", name }] };
  }
}

// ---------------------------------------------------------------------------
// Storage -> display

function stripMarkers(text: string): string {
  let out = text;
  for (;;) {
    if (startsWithFold(out, XLFN_XLWS)) out = out.slice(XLFN_XLWS.length);
    else if (startsWithFold(out, XLFN)) out = out.slice(XLFN.length);
    else if (startsWithFold(out, XLPM)) out = out.slice(XLPM.length);
    else if (startsWithFold(out, XLETA)) out = out.slice(XLETA.length);
    else if (startsWithFold(out, XLOP)) out = out.slice(XLOP.length);
    else return out;
  }
}

function emitFromArgs(args: readonly (readonly Node[])[]): string {
  return args.map((arg) => emitFromList(arg)).join(",");
}

/// A stored optional parameter `_xlop.y` is displayed as `[y]`.
function emitFromParameter(arg: readonly Node[]): string {
  const lone = loneName(arg);
  if (!lone || !startsWithFold(lone.name, XLOP)) return emitFromList(arg);
  return `${lone.lead}[${lone.name.slice(XLOP.length)}]${lone.trail}`;
}

function emitFromCall(node: CallNode): string {
  const closer = node.closed ? ")" : "";
  const { prefix, base } = splitFunctionPrefix(node.name);
  const upper = base.toUpperCase();
  if (upper === "LAMBDA" && node.args.length > 0) {
    const last = node.args.length - 1;
    const args = node.args.map((arg, index) =>
      index === last ? emitFromList(arg) : emitFromParameter(arg),
    );
    return `${stripMarkers(node.name)}${node.ws}(${args.join(",")}${closer}`;
  }
  if (prefix !== "" && (upper === "SINGLE" || upper === "ANCHORARRAY")) {
    const only = node.args.length === 1 ? (node.args[0] as readonly Node[]) : undefined;
    if (only && only.length > 0 && node.ws === "" && node.closed) {
      const text = emitFromList(only);
      if (upper === "SINGLE") return rangeEnd(only, 0) === only.length ? `@${text}` : `@(${text})`;
      if (primaryEnd(only, 0) === only.length) return `${text}#`;
    }
    // Not the single-operand shape: keep the stored call as it is.
    return `${node.name}${node.ws}(${emitFromArgs(node.args)}${closer}`;
  }
  return `${stripMarkers(node.name)}${node.ws}(${emitFromArgs(node.args)}${closer}`;
}

function emitFromList(nodes: readonly Node[]): string {
  let out = "";
  for (const node of nodes) {
    switch (node.t) {
      case "tok":
        out += node.s;
        break;
      case "name":
        out += stripMarkers(node.s);
        break;
      case "group":
        out += `(${emitFromArgs(node.args)}${node.closed ? ")" : ""}`;
        break;
      case "call":
        out += emitFromCall(node);
        break;
      case "at":
        out += "@";
        break;
      case "spill":
        out += "#";
        break;
    }
  }
  return out;
}

/// Converts stored formula text back to display form: drops `_xlfn.`,
/// `_xlfn._xlws.` and `_xlpm.` markers, turns `_xlfn.SINGLE(x)` into `@x` and
/// `_xlfn.ANCHORARRAY(ref)` into `ref#`. Pure and idempotent.
export function fromStorageFormula(stored: string): string {
  if (!/_xl/i.test(stored)) return stored;
  try {
    return emitFromList(parseFormula(stored));
  } catch {
    // Too deeply nested to walk: leave the stored text as it is.
    return stored;
  }
}

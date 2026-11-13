// Modified for lit-ai: new module. Deck-wide deterministic text tools (cross-run find/replace incl. notes, role selectors, font and color replacement).
/**
 * Deck-wide text tools for the canonical op layer: one call does the whole
 * deck (or a slide range) with an exact count, instead of the model sending
 * one op per element.
 *
 * - Find/replace matches across differently formatted runs of one paragraph
 *   (never across a paragraph, a soft break, a field or a raw equation run).
 *   The replacement takes the formatting of the run where the match starts;
 *   the consumed text of later runs is removed in place, so the run count and
 *   every untouched byte stay aligned with the original XML (lossless patch).
 *   Speaker notes are patched in their own XML the same way.
 * - Role selectors classify elements as title / subtitle / body / footer /
 *   table / text (non-placeholder shapes and text boxes) from the parsed
 *   placeholder type, so "all titles on slides 3-9" is one deterministic op.
 * - Font replacement rewrites typeface references (deck-wide: slides, theme,
 *   masters and layouts, so inherited text follows; slide-scoped: the runs
 *   on those slides); color replacement rewrites explicit sRGB colors and,
 *   deck-wide, matching theme slots and master/layout colors.
 *
 * Every function has a count-only mode so an op can refuse on an expected
 * count before it mutates anything.
 */
import {
  commitSaved,
  editTableCellText,
  patchGroupChildText,
  patchSlideXml,
  reparseDeck,
  setElementFont,
  setElementParagraphFormat,
  setGroupChildFont,
  setGroupChildParagraphFormat,
  type ElementFontPatch,
  type OpenedPptx,
  type ParagraphFormatPatch,
} from './index'
import { elementDurableId, groupChildDurableId } from './identity'
import { findNotesBodySp, notesPathForSlide, unescapeXml } from './notes'
import type { GroupElement, Paragraph, Slide, SlideElement, TableElement, TextElement, TextRun } from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

// ── Roles ───────────────────────────────────────────────────────────────

export type TextRole = 'title' | 'subtitle' | 'body' | 'footer' | 'table' | 'text' | 'other'

/** Roles a selector may name; "all" selects every text-bearing element. */
export const TEXT_ROLE_NAMES = ['title', 'subtitle', 'body', 'footer', 'table', 'text', 'all'] as const

/** Placeholder type → role (the parser stores a typeless <p:ph> as "body"). */
export function placeholderRole(placeholder: string | undefined): TextRole {
  if (!placeholder) return 'text'
  if (placeholder === 'title' || placeholder === 'ctrTitle') return 'title'
  if (placeholder === 'subTitle') return 'subtitle'
  if (placeholder === 'body' || placeholder === 'obj') return 'body'
  if (placeholder === 'dt' || placeholder === 'ftr' || placeholder === 'sldNum' || placeholder === 'hdr')
    return 'footer'
  return 'other'
}

/** Role of one element (groups have no role of their own; their children are classified). */
export function elementRole(el: SlideElement): TextRole {
  if (el.type === 'table') return 'table'
  if (el.type === 'text' || el.type === 'shape') return placeholderRole(el.placeholder)
  return 'other'
}

function roleSelected(roles: readonly string[], role: TextRole): boolean {
  if (role === 'other') return roles.includes('all')
  return roles.includes('all') || roles.includes(role)
}

/** Validate a role list from an op; throws a guided Error. */
export function normalizeRoles(value: unknown, opName: string): string[] {
  const list = typeof value === 'string' ? [value] : value
  if (!Array.isArray(list) || list.length === 0) {
    throw new Error(`op "${opName}" needs "roles": one or more of ${TEXT_ROLE_NAMES.join('/')}.`)
  }
  for (const r of list) {
    if (typeof r !== 'string' || !(TEXT_ROLE_NAMES as readonly string[]).includes(r)) {
      throw new Error(
        `op "${opName}": unknown role ${JSON.stringify(r)}. Roles: ${TEXT_ROLE_NAMES.join(', ')}.`,
      )
    }
  }
  return [...new Set(list as string[])]
}

/** One text-bearing target: a top-level element or a direct child of a top-level group. */
interface TextTarget {
  slide: Slide
  slideIndex: number
  el: SlideElement
  group?: GroupElement
  role: TextRole
  id: string
}

function hasParagraphs(el: SlideElement): boolean {
  if (el.type === 'text' || el.type === 'shape') return !!(el as TextElement).text?.paragraphs.length
  if (el.type === 'table')
    return (el as TableElement).rows.some((row) => row.some((c) => !c.merged && !!c.text?.paragraphs.length))
  return false
}

function* textTargets(opened: OpenedPptx, slides: number[] | undefined): Generator<TextTarget> {
  const deck = opened.deck.slides
  const indices = slides ?? deck.map((_, i) => i)
  for (const slideIndex of indices) {
    const slide = deck[slideIndex]
    if (!slide) continue
    for (const el of slide.elements) {
      if (el.type === 'group') {
        const grp = el as GroupElement
        for (const child of grp.children) {
          if (child.type !== 'text' && child.type !== 'shape') continue
          yield {
            slide,
            slideIndex,
            el: child,
            group: grp,
            role: placeholderRole(child.placeholder),
            id: groupChildDurableId(grp, child) ?? child.id,
          }
        }
        continue
      }
      if (el.type !== 'text' && el.type !== 'shape' && el.type !== 'table') continue
      yield { slide, slideIndex, el, role: elementRole(el), id: elementDurableId(el) ?? el.id }
    }
  }
}

// ── Pattern compilation ─────────────────────────────────────────────────

export interface FindSpec {
  find: string
  /** Treat find as a regular expression (bounded: no nested quantifiers, no lookbehind). */
  regex?: boolean
  matchCase?: boolean
  /** Match only whole words (letters/digits/underscore on neither side). */
  wholeWord?: boolean
}

const MAX_REGEX_CHARS = 200
const MAX_LITERAL_CHARS = 1000
const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/**
 * Compile a find spec into a global unicode RegExp. User regexes are refused
 * in the shapes that backtrack catastrophically on ordinary prose (nested
 * quantifiers) and when they use lookbehind; whole-word matching wraps the
 * pattern in boundary assertions this module adds itself.
 */
export function compileFindPattern(spec: FindSpec): RegExp {
  if (typeof spec.find !== 'string' || !spec.find) throw new Error('"find" must be a non-empty string.')
  let source: string
  if (spec.regex) {
    if (spec.find.length > MAX_REGEX_CHARS)
      throw new Error(`"find" regex is longer than ${MAX_REGEX_CHARS} characters.`)
    if (/\((?:\?:)?[^()]*[+*][^()]*\)\s*[+*{]/.test(spec.find) || /\([^()]*\)[+*]\)[+*]/.test(spec.find))
      throw new Error('"find" regex has nested quantifiers (like (a+)+); rewrite it without them.')
    if (/\(\?<[=!]/.test(spec.find)) throw new Error('"find" regex lookbehind is not supported.')
    source = spec.find
  } else {
    if (spec.find.length > MAX_LITERAL_CHARS)
      throw new Error(`"find" is longer than ${MAX_LITERAL_CHARS} characters.`)
    source = escapeRegExp(spec.find)
  }
  if (spec.wholeWord) source = `(?<![\\p{L}\\p{N}_])(?:${source})(?![\\p{L}\\p{N}_])`
  try {
    return new RegExp(source, spec.matchCase ? 'gu' : 'giu')
  } catch (e) {
    throw new Error(`"find" is not a valid regex: ${e instanceof Error ? e.message : String(e)}`)
  }
}

/** Expand $&, $1..$99, $<name> and $$ in a regex replacement. */
function expandReplacement(template: string, m: RegExpExecArray): string {
  return template.replace(/\$(\$|&|<([^>]*)>|(\d{1,2}))/g, (whole, token: string, name?: string, digits?: string) => {
    if (token === '$') return '$'
    if (token === '&') return m[0]
    if (name !== undefined) return m.groups?.[name] ?? ''
    if (digits !== undefined) {
      const two = Number(digits)
      if (digits.length === 2 && two < m.length) return m[two] ?? ''
      const one = Number(digits[0])
      if (one > 0 && one < m.length) return (m[one] ?? '') + (digits.length === 2 ? digits[1] : '')
      return whole
    }
    return whole
  })
}

// ── Cross-run replacement core ──────────────────────────────────────────

/** A text cell in document order; boundaries (breaks, fields, equations) split matching. */
export interface TextCell {
  text: string
  boundary: boolean
}

export interface ReplacePlan {
  /** New text per cell (same length as the input); unchanged cells keep their text */
  texts: string[]
  count: number
}

/**
 * Plan replacements over one paragraph's cells. Matches run over each
 * contiguous non-boundary segment, so one match may span several runs; the
 * replacement lands in the run where the match starts and the matched text
 * of the following runs is removed. `budget` caps the number of matches.
 */
export function planCellReplacements(
  cells: readonly TextCell[],
  re: RegExp,
  replace: string,
  isRegex: boolean,
  budget: number,
): ReplacePlan {
  const texts = cells.map((c) => c.text)
  let count = 0
  let i = 0
  while (i < cells.length && count < budget) {
    if (cells[i]!.boundary) {
      i++
      continue
    }
    let j = i
    while (j < cells.length && !cells[j]!.boundary) j++
    // Segment [i, j): concatenated text with per-cell offsets
    const offsets: number[] = []
    let joined = ''
    for (let k = i; k < j; k++) {
      offsets.push(joined.length)
      joined += cells[k]!.text
    }
    const matches: Array<{ s: number; e: number; r: string }> = []
    re.lastIndex = 0
    let m: RegExpExecArray | null
    while (count + matches.length < budget && (m = re.exec(joined)) !== null) {
      if (m[0].length === 0) {
        re.lastIndex = m.index + 1
        continue
      }
      matches.push({ s: m.index, e: m.index + m[0].length, r: isRegex ? expandReplacement(replace, m) : replace })
    }
    const cellAt = (pos: number): number => {
      let k = offsets.length - 1
      while (k > 0 && offsets[k]! > pos) k--
      return k
    }
    for (let x = matches.length - 1; x >= 0; x--) {
      const { s, e, r } = matches[x]!
      const a = cellAt(s)
      const b = cellAt(e - 1)
      const ta = texts[i + a]!
      const sa = s - offsets[a]!
      if (a === b) {
        texts[i + a] = ta.slice(0, sa) + r + ta.slice(e - offsets[a]!)
        continue
      }
      texts[i + a] = ta.slice(0, sa) + r
      for (let k = a + 1; k < b; k++) texts[i + k] = ''
      texts[i + b] = texts[i + b]!.slice(e - offsets[b]!)
    }
    count += matches.length
    i = j
  }
  return { texts, count }
}

const isBoundaryRun = (r: TextRun): boolean =>
  !!r.field || r.rawXml !== undefined || r.text === '\n' || !!r.paraMark

/** Replace inside model paragraphs; returns the match count (writes only when apply). */
function replaceInParagraphs(
  paragraphs: Paragraph[],
  re: RegExp,
  replace: string,
  isRegex: boolean,
  budget: { left: number },
  apply: boolean,
): number {
  let total = 0
  for (const p of paragraphs) {
    if (budget.left <= 0) break
    const plan = planCellReplacements(
      p.runs.map((r) => ({ text: r.text ?? '', boundary: isBoundaryRun(r) })),
      re,
      replace,
      isRegex,
      budget.left,
    )
    if (!plan.count) continue
    total += plan.count
    budget.left -= plan.count
    if (apply) p.runs.forEach((r, k) => { if (!isBoundaryRun(r)) r.text = plan.texts[k]! })
  }
  return total
}

const NOTES_PARA_RE = /<a:p>([\s\S]*?)<\/a:p>/g
const NOTES_TOKEN_RE = /<a:r>([\s\S]*?)<\/a:r>|<a:br\b[^>]*\/>|<a:br\b[^>]*>[\s\S]*?<\/a:br>|<a:fld\b[^>]*\/>|<a:fld\b[^>]*>[\s\S]*?<\/a:fld>/g
const RUN_T_RE = /<a:t>([\s\S]*?)<\/a:t>|<a:t\/>/

/** Replace inside a notes part's body placeholder, run by run (formatting bytes untouched). */
function replaceInNotesXml(
  xml: string,
  re: RegExp,
  replace: string,
  isRegex: boolean,
  budget: { left: number },
): { xml: string; count: number } {
  const body = findNotesBodySp(xml)
  if (!body) return { xml, count: 0 }
  let count = 0
  const patchedBody = body.xml.replace(NOTES_PARA_RE, (para: string, inner: string) => {
    if (budget.left <= 0) return para
    const tokens: Array<{ start: number; end: number; run: boolean; body: string; text: string }> = []
    for (const t of inner.matchAll(NOTES_TOKEN_RE)) {
      const isRun = t[1] !== undefined
      const tm = isRun ? RUN_T_RE.exec(t[1]!) : null
      tokens.push({
        start: t.index!,
        end: t.index! + t[0].length,
        run: isRun && !!tm,
        body: t[0],
        text: tm && tm[1] !== undefined ? unescapeXml(tm[1]) : '',
      })
    }
    if (!tokens.length) return para
    const plan = planCellReplacements(
      tokens.map((t) => ({ text: t.text, boundary: !t.run })),
      re,
      replace,
      isRegex,
      budget.left,
    )
    if (!plan.count) return para
    count += plan.count
    budget.left -= plan.count
    let out = ''
    let cursor = 0
    tokens.forEach((t, k) => {
      out += inner.slice(cursor, t.start)
      if (t.run && plan.texts[k] !== t.text) {
        out += t.body.replace(RUN_T_RE, () => `<a:t>${escapeXmlText(plan.texts[k]!)}</a:t>`)
      } else out += t.body
      cursor = t.end
    })
    out += inner.slice(cursor)
    return `<a:p>${out}</a:p>`
  })
  if (!count) return { xml, count: 0 }
  return { xml: xml.slice(0, body.start) + patchedBody + xml.slice(body.end), count }
}

export interface DeckReplaceOptions extends FindSpec {
  replace: string
  /** 0-based slide indices to cover; undefined = every slide */
  slides?: number[]
  /** Restrict to one element (durable or parse-time id; with a single slide) */
  elementId?: string
  /** Also replace in speaker notes (default false) */
  includeNotes?: boolean
  /** Replace only the first match */
  firstOnly?: boolean
}

export interface DeckReplaceResult {
  count: number
  slides: Array<{ slide: number; count: number; notes?: number }>
  elements: string[]
}

/**
 * Deck find/replace across runs (and optionally notes). apply=false counts
 * without mutating anything, so callers can enforce an expected count first.
 */
export function findReplaceInDeck(
  opened: OpenedPptx,
  opts: DeckReplaceOptions,
  apply: boolean,
): DeckReplaceResult {
  const re = compileFindPattern(opts)
  const isRegex = !!opts.regex
  const budget = { left: opts.firstOnly ? 1 : Number.POSITIVE_INFINITY }
  const perSlide = new Map<number, { slide: number; count: number; notes?: number }>()
  const elements: string[] = []
  const bump = (slide: number, n: number, notes = false) => {
    const row = perSlide.get(slide) ?? { slide, count: 0 }
    row.count += n
    if (notes) row.notes = (row.notes ?? 0) + n
    perSlide.set(slide, row)
  }
  for (const t of textTargets(opened, opts.slides)) {
    if (budget.left <= 0) break
    if (opts.elementId != null && t.id !== opts.elementId && t.el.id !== opts.elementId) continue
    let n = 0
    if (t.el.type === 'table') {
      const table = t.el as TableElement
      table.rows.forEach((row, ri) =>
        row.forEach((cell, ci) => {
          if (budget.left <= 0 || cell.merged || !cell.text?.paragraphs.length) return
          const hit = replaceInParagraphs(cell.text.paragraphs, re, opts.replace, isRegex, budget, apply)
          if (hit && apply) editTableCellText(t.slide, table.id, ri, ci, cell.text.paragraphs)
          n += hit
        }),
      )
    } else {
      const te = t.el as TextElement
      if (!te.text?.paragraphs.length) continue
      n = replaceInParagraphs(te.text.paragraphs, re, opts.replace, isRegex, budget, apply)
      if (n && apply) {
        if (t.group) patchGroupChildText(t.slide, t.group.id, te)
        else te.dirty = true
      }
    }
    if (n) {
      bump(t.slideIndex, n)
      elements.push(`${t.slideIndex}/${t.id}`)
    }
  }
  if (opts.includeNotes && opts.elementId == null) {
    const indices = opts.slides ?? opened.deck.slides.map((_, i) => i)
    for (const i of indices) {
      if (budget.left <= 0) break
      const slide = opened.deck.slides[i]
      if (!slide) continue
      const path = notesPathForSlide(opened.archive, slide.path)
      const xml = path ? opened.archive.readText(path) : null
      if (!path || !xml) continue
      const r = replaceInNotesXml(xml, re, opts.replace, isRegex, budget)
      if (!r.count) continue
      if (apply) opened.archive.entries.set(path, Buffer.from(r.xml, 'utf8'))
      bump(i, r.count, true)
    }
  }
  const slides = [...perSlide.values()].sort((a, b) => a.slide - b.slide)
  return { count: slides.reduce((s, r) => s + r.count, 0), slides, elements }
}

// ── Format by role ──────────────────────────────────────────────────────

export interface FormatByRoleOptions {
  roles: string[]
  slides?: number[]
  font?: ElementFontPatch
  paragraph?: ParagraphFormatPatch
}

export interface FormatByRoleResult {
  count: number
  byRole: Record<string, number>
  slides: Array<{ slide: number; count: number }>
  elements: string[]
}

/** Apply one font/paragraph patch to every element of the selected roles. */
export function formatByRole(
  opened: OpenedPptx,
  opts: FormatByRoleOptions,
  apply: boolean,
): FormatByRoleResult {
  const byRole: Record<string, number> = {}
  const perSlide = new Map<number, number>()
  const elements: string[] = []
  const font = opts.font && Object.keys(opts.font).length ? opts.font : undefined
  const para = opts.paragraph && Object.keys(opts.paragraph).length ? opts.paragraph : undefined
  for (const t of textTargets(opened, opts.slides)) {
    if (!roleSelected(opts.roles, t.role) || !hasParagraphs(t.el)) continue
    // Paragraph formats do not apply to table cells; a table selected only for them is skipped
    if (!font && t.el.type === 'table') continue
    if (apply) {
      if (t.group) {
        if (font) setGroupChildFont(t.slide, t.group.id, t.el.id, font)
        if (para) setGroupChildParagraphFormat(t.slide, t.group.id, t.el.id, para)
      } else {
        if (font) setElementFont(t.slide, t.el.id, font)
        if (para && t.el.type !== 'table') setElementParagraphFormat(t.slide, t.el.id, para)
      }
    }
    byRole[t.role] = (byRole[t.role] ?? 0) + 1
    perSlide.set(t.slideIndex, (perSlide.get(t.slideIndex) ?? 0) + 1)
    elements.push(`${t.slideIndex}/${t.id}`)
  }
  return {
    count: elements.length,
    byRole,
    slides: [...perSlide].sort((a, b) => a[0] - b[0]).map(([slide, count]) => ({ slide, count })),
    elements,
  }
}
// ── Font replacement ────────────────────────────────────────────────────

export interface FontReplaceOptions {
  from: string
  to: string
  /** 0-based slide indices; undefined = whole deck (theme, masters and layouts included) */
  slides?: number[]
}

export interface FontReplaceResult {
  /** Typeface references rewritten in slide parts (deck-wide) or runs given the new font (slide-scoped) */
  slideRefs: number
  /** Theme heading/body latin fonts rewritten */
  themeFonts: number
  /** Typeface references rewritten in masters, layouts and presentation defaults */
  inherited: number
  slides: Array<{ slide: number; count: number }>
  /** Deck-wide apply only: text runs still resolving to `from` after the rewrite */
  remainingRuns?: number
}

const sameFont = (a: string | undefined, b: string) =>
  !!a && a.trim().toLowerCase() === b.trim().toLowerCase()

const LAYOUT_PART_RE = /^ppt\/(?:slideMasters\/slideMaster|slideLayouts\/slideLayout)\d+\.xml$/
const THEME_PART_RE = /^ppt\/theme\/theme\d+\.xml$/

/** Rewrite <a:latin|ea|cs typeface="from"> attributes; returns [xml, count]. */
function replaceTypefaces(xml: string, from: string, to: string): [string, number] {
  let n = 0
  const out = xml.replace(
    /(<a:(?:latin|ea|cs)\b[^>]*\btypeface=")([^"]*)(")/g,
    (whole, a: string, face: string, z: string) => {
      if (!sameFont(unescapeXml(face), from)) return whole
      n++
      return a + escapeXmlAttr(to) + z
    },
  )
  return [out, n]
}

/** Theme major/minor latin typefaces equal to `from` → `to`. */
function replaceThemeFonts(xml: string, from: string, to: string): [string, number] {
  let n = 0
  const out = xml.replace(
    /(<a:(?:major|minor)Font>[\s\S]*?<a:latin\b[^>]*\btypeface=")([^"]*)(")/g,
    (whole, a: string, face: string, z: string) => {
      if (!sameFont(unescapeXml(face), from)) return whole
      n++
      return a + escapeXmlAttr(to) + z
    },
  )
  return [out, n]
}

function countRunsWithFont(opened: OpenedPptx, font: string): number {
  let n = 0
  const visit = (paragraphs: Paragraph[]) => {
    for (const p of paragraphs)
      for (const r of p.runs) if (r.text && !r.paraMark && sameFont(r.fontFamily, font)) n++
  }
  for (const t of textTargets(opened, undefined)) {
    if (t.el.type === 'table') {
      for (const row of (t.el as TableElement).rows)
        for (const c of row) if (!c.merged && c.text) visit(c.text.paragraphs)
    } else if ((t.el as TextElement).text) visit((t.el as TextElement).text!.paragraphs)
  }
  return n
}

/**
 * Replace font `from` with `to`.
 *
 * Deck-wide (no slides): byte-level rewrite of every typeface reference equal
 * to `from` (latin, east-asian and complex-script slots independently) in the
 * slides, masters, layouts and presentation defaults, plus the theme heading/
 * body fonts, then a reparse. Inherited and theme-linked text follows without
 * being baked in; a dual-font run keeps its other script's typeface.
 *
 * Slide-scoped: every run on those slides whose resolved font is `from` gets
 * `to` explicitly (all three script slots, like setFont); theme stays.
 */
export function replaceFontInDeck(
  opened: OpenedPptx,
  opts: FontReplaceOptions,
  apply: boolean,
): FontReplaceResult {
  if (!opts.from.trim() || !opts.to.trim()) throw new Error('"from" and "to" must be font names.')
  const perSlide: Array<{ slide: number; count: number }> = []
  if (opts.slides !== undefined) {
    let total = 0
    const visit = (paragraphs: Paragraph[]): number => {
      let n = 0
      for (const p of paragraphs)
        for (const r of p.runs) {
          if (r.paraMark || r.rawXml !== undefined || !r.text || !sameFont(r.fontFamily, opts.from))
            continue
          n++
          if (apply) {
            r.fontFamily = opts.to
            delete r.latinFont
            delete r.eaFont
            delete r.csFont
            delete r.fontImplicit
            delete r.latinFamily
          }
        }
      return n
    }
    const counts = new Map<number, number>()
    for (const t of textTargets(opened, opts.slides)) {
      let n = 0
      if (t.el.type === 'table') {
        const table = t.el as TableElement
        table.rows.forEach((row, ri) =>
          row.forEach((cell, ci) => {
            if (cell.merged || !cell.text?.paragraphs.length) return
            const hit = visit(cell.text.paragraphs)
            if (hit && apply) editTableCellText(t.slide, table.id, ri, ci, cell.text.paragraphs)
            n += hit
          }),
        )
      } else {
        const te = t.el as TextElement
        if (!te.text?.paragraphs.length) continue
        n = visit(te.text.paragraphs)
        if (n && apply) {
          if (t.group) patchGroupChildText(t.slide, t.group.id, te)
          else te.dirty = true
        }
      }
      if (n) counts.set(t.slideIndex, (counts.get(t.slideIndex) ?? 0) + n)
      total += n
    }
    for (const [slide, count] of [...counts].sort((a, b) => a[0] - b[0])) perSlide.push({ slide, count })
    return { slideRefs: total, themeFonts: 0, inherited: 0, slides: perSlide }
  }
  const { archive, deck } = opened
  const edits: Array<[string, string]> = []
  let slideRefs = 0
  deck.slides.forEach((slide, i) => {
    // Current bytes including unsaved model edits (the apply path bakes them first)
    const [next, n] = replaceTypefaces(patchSlideXml(slide), opts.from, opts.to)
    if (!n) return
    slideRefs += n
    perSlide.push({ slide: i, count: n })
    edits.push([slide.path, next])
  })
  let themeFonts = 0
  let inherited = 0
  for (const path of archive.entries.keys()) {
    const isTheme = THEME_PART_RE.test(path)
    if (!isTheme && !LAYOUT_PART_RE.test(path) && path !== 'ppt/presentation.xml') continue
    const xml = archive.readText(path)
    if (!xml) continue
    const [next, n] = isTheme
      ? replaceThemeFonts(xml, opts.from, opts.to)
      : replaceTypefaces(xml, opts.from, opts.to)
    if (!n) continue
    if (isTheme) themeFonts += n
    else inherited += n
    edits.push([path, next])
  }
  if (!apply || !edits.length) return { slideRefs, themeFonts, inherited, slides: perSlide }
  commitSaved(opened)
  for (const [path, xml] of edits) archive.entries.set(path, Buffer.from(xml, 'utf8'))
  opened.deck.slides = reparseDeck(opened).deck.slides
  return {
    slideRefs,
    themeFonts,
    inherited,
    slides: perSlide,
    remainingRuns: countRunsWithFont(opened, opts.from),
  }
}

// ── Color replacement ───────────────────────────────────────────────────

export interface ColorReplaceOptions {
  /** from → to pairs, "#RRGGBB" */
  map: Array<{ from: string; to: string }>
  /** 0-based slide indices; undefined = whole deck (theme slots, masters and layouts included) */
  slides?: number[]
}

export interface ColorReplaceResult {
  /** Explicit sRGB color occurrences rewritten in slide parts */
  replaced: number
  /** Theme scheme slots rewritten (e.g. "accent1") */
  themeSlots: string[]
  /** Explicit colors rewritten in masters and layouts */
  inherited: number
  slides: Array<{ slide: number; count: number }>
}

const hexKey = (c: string) => c.replace(/^#/, '').toUpperCase()
const SRGB_VAL_RE = /(<a:srgbClr\b[^>]*\bval=")([0-9A-Fa-f]{6})(")/g
const SCHEME_SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]

function recolor(xml: string, map: Map<string, string>): [string, number] {
  let n = 0
  const out = xml.replace(SRGB_VAL_RE, (whole, a: string, hex: string, z: string) => {
    const to = map.get(hex.toUpperCase())
    if (!to) return whole
    n++
    return a + to + z
  })
  return [out, n]
}

function recolorThemeSlots(xml: string, map: Map<string, string>, slots: Set<string>): string {
  let out = xml
  for (const key of SCHEME_SLOTS) {
    out = out.replace(new RegExp(`<a:${key}>([\\s\\S]*?)</a:${key}>`), (whole, inner: string) => {
      const m =
        /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"|<a:sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(
          inner,
        )
      const current = (m?.[1] ?? m?.[2])?.toUpperCase()
      const to = current ? map.get(current) : undefined
      if (!to) return whole
      slots.add(key)
      return `<a:${key}><a:srgbClr val="${to}"/></a:${key}>`
    })
  }
  return out
}

/**
 * Replace explicit sRGB colors on slides; deck-wide also the theme scheme
 * slots and master/layout colors equal to a "from" color, so theme-linked
 * content follows. Pending edits are baked first and the deck is reparsed so
 * resolved colors refresh. Chart parts keep their own colors.
 */
export function replaceColorInDeck(
  opened: OpenedPptx,
  opts: ColorReplaceOptions,
  apply: boolean,
): ColorReplaceResult {
  const map = new Map(opts.map.map(({ from, to }) => [hexKey(from), hexKey(to)]))
  const deckWide = opts.slides === undefined
  const { archive, deck } = opened
  const edits: Array<[string, string]> = []
  const perSlide: Array<{ slide: number; count: number }> = []
  let replaced = 0
  const indices = opts.slides ?? deck.slides.map((_, i) => i)
  for (const i of indices) {
    const slide = deck.slides[i]
    if (!slide) continue
    const [next, n] = recolor(patchSlideXml(slide), map)
    if (!n) continue
    replaced += n
    perSlide.push({ slide: i, count: n })
    edits.push([slide.path, next])
  }
  const themeSlots = new Set<string>()
  let inherited = 0
  if (deckWide) {
    for (const path of archive.entries.keys()) {
      if (!THEME_PART_RE.test(path) && !LAYOUT_PART_RE.test(path)) continue
      const xml = archive.readText(path)
      if (!xml) continue
      if (THEME_PART_RE.test(path)) {
        const next = recolorThemeSlots(xml, map, themeSlots)
        if (next !== xml) edits.push([path, next])
      } else {
        const [next, n] = recolor(xml, map)
        inherited += n
        if (n) edits.push([path, next])
      }
    }
  }
  if (apply && edits.length) {
    commitSaved(opened)
    for (const [path, xml] of edits) archive.entries.set(path, Buffer.from(xml, 'utf8'))
    opened.deck.slides = reparseDeck(opened).deck.slides
  }
  return { replaced, themeSlots: [...themeSlots], inherited, slides: perSlide }
}

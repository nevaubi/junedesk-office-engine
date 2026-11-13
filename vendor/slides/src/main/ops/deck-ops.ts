// Modified for lit-ai: new module. Deck-wide selector ops (formatText, replaceFont, replaceColor) and geometry matching (matchElements, matchAcrossSlides).
/**
 * Selector-driven deck ops: one op formats every element of a role across a
 * slide range, replaces a font or a color deck-wide, or matches size and
 * position to a reference element (on one slide or across slides). Each op
 * counts first, refuses on an `expectedCount` mismatch before mutating, and
 * returns a bounded receipt in its record's `after`.
 */
import {
  elementDurableId,
  elementRole,
  formatByRole,
  normalizeRoles,
  replaceColorInDeck,
  replaceFontInDeck,
  resizeTable,
  updateConnectorsForMoved,
  type ElementFontPatch,
  type ParagraphFormatPatch,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  matchesElementRef,
  register,
  requireFinite,
  requireHexColor,
  resolveElement,
  resolveSlideScope,
  type Op,
  type OpContext,
  type OpRecord,
} from './registry'

/** Receipts list at most this many element refs (the count is always exact). */
const RECEIPT_ELEMENTS = 50

function boundedList<T>(list: T[]): { list: T[]; total: number } {
  return { list: list.slice(0, RECEIPT_ELEMENTS), total: list.length }
}

function guided<T>(opName: string, fn: () => T): T {
  try {
    return fn()
  } catch (e) {
    if (e instanceof GuidedError) throw e
    const msg = e instanceof Error ? e.message : String(e)
    throw new GuidedError(msg.startsWith('op "') ? msg : `op "${opName}": ${msg}`)
  }
}

/** Optional exact-count guard shared by the selector ops. */
function expectedCount(op: Op): number | undefined {
  const v = op.expectedCount
  if (v === undefined) return undefined
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 0) {
    throw new GuidedError(`op "${op.op}": "expectedCount" must be a non-negative integer.`)
  }
  return v
}

function requireCount(op: Op, actual: number, detail: string): void {
  const want = expectedCount(op)
  if (want !== undefined && want !== actual) {
    throw new GuidedError(
      `op "${op.op}": expectedCount is ${want} but the selector matches ${actual}${detail ? ` (${detail})` : ''}. Nothing was changed; inspect with search_deck/read_slide and resend with the right scope or count.`,
    )
  }
}

const perSlideDetail = (slides: Array<{ slide: number; count: number }>) =>
  slides
    .slice(0, 20)
    .map((s) => `slide ${s.slide}: ${s.count}`)
    .join(', ')

// ── font / paragraph patch validation (same rules as setFont / setParagraphFormat) ──

function fontPatch(op: Op): ElementFontPatch | undefined {
  const raw = op.font
  if (raw === undefined) return undefined
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new GuidedError(`op "${op.op}": "font" must be an object {fontFamily?, fontSizePt?, bold?, italic?, underline?, strike?, color?}.`)
  }
  const f = raw as Record<string, unknown>
  const out: ElementFontPatch = {}
  for (const k of Object.keys(f)) {
    if (!['fontFamily', 'fontSizePt', 'bold', 'italic', 'underline', 'strike', 'color'].includes(k)) {
      throw new GuidedError(`op "${op.op}": unknown font field "${k}".`)
    }
  }
  if (f.fontFamily !== undefined) {
    if (typeof f.fontFamily !== 'string' || !f.fontFamily.trim() || f.fontFamily.length > 64) {
      throw new GuidedError(`op "${op.op}": "font.fontFamily" must be a font name.`)
    }
    out.fontFamily = f.fontFamily.trim()
  }
  if (f.fontSizePt !== undefined) {
    requireFinite(f.fontSizePt, op.op, 'font.fontSizePt')
    if (f.fontSizePt <= 0 || f.fontSizePt > 400) {
      throw new GuidedError(`op "${op.op}": "font.fontSizePt" must be between 0 and 400 points.`)
    }
    out.fontSizePt = f.fontSizePt
  }
  for (const k of ['bold', 'italic', 'underline', 'strike'] as const) {
    if (f[k] === undefined) continue
    if (typeof f[k] !== 'boolean') throw new GuidedError(`op "${op.op}": "font.${k}" must be a boolean.`)
    out[k] = f[k] as boolean
  }
  if (f.color !== undefined) {
    requireHexColor(f.color, op.op, 'font.color')
    out.color = `#${String(f.color).replace(/^#/, '').slice(0, 6).toUpperCase()}`
  }
  return Object.keys(out).length ? out : undefined
}

function paragraphPatch(op: Op): ParagraphFormatPatch | undefined {
  const raw = op.paragraph
  if (raw === undefined) return undefined
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new GuidedError(`op "${op.op}": "paragraph" must be an object {align?, lineSpacingPct?, spaceBeforePt?, spaceAfterPt?}.`)
  }
  const p = raw as Record<string, unknown>
  const out: ParagraphFormatPatch = {}
  for (const k of Object.keys(p)) {
    if (!['align', 'lineSpacingPct', 'spaceBeforePt', 'spaceAfterPt'].includes(k)) {
      throw new GuidedError(`op "${op.op}": unknown paragraph field "${k}" (bullets and indents use setParagraphFormat per element).`)
    }
  }
  if (p.align !== undefined) {
    if (!['left', 'center', 'right', 'justify'].includes(String(p.align))) {
      throw new GuidedError(`op "${op.op}": "paragraph.align" must be left/center/right/justify.`)
    }
    out.align = p.align as ParagraphFormatPatch['align']
  }
  for (const k of ['lineSpacingPct', 'spaceBeforePt', 'spaceAfterPt'] as const) {
    if (p[k] === undefined) continue
    requireFinite(p[k], op.op, `paragraph.${k}`)
    if ((p[k] as number) < 0) throw new GuidedError(`op "${op.op}": "paragraph.${k}" must be >= 0.`)
    out[k] = p[k] as number
  }
  return Object.keys(out).length ? out : undefined
}

// ── formatText ──────────────────────────────────────────────────────────

function formatSpec(op: Op, ctx: OpContext) {
  const roles = guided(op.op, () => normalizeRoles(op.roles, op.op))
  const font = fontPatch(op)
  const paragraph = paragraphPatch(op)
  if (!font && !paragraph) {
    throw new GuidedError(`op "formatText" needs "font" and/or "paragraph" with at least one field.`)
  }
  return { roles, font, paragraph, slides: resolveSlideScope(ctx, op) }
}

register({
  name: 'formatText',
  validate(op, ctx) {
    formatSpec(op, ctx)
    expectedCount(op)
  },
  apply(op, ctx): OpRecord {
    const spec = formatSpec(op, ctx)
    const dry = formatByRole(ctx.opened, spec, false)
    if (dry.count === 0) {
      throw new GuidedError(
        `op "formatText": no element with text matches roles [${spec.roles.join(', ')}]${spec.slides ? ` on slides [${spec.slides.join(', ')}]` : ''}. Check the roles with search_deck or read_slide.`,
      )
    }
    requireCount(op, dry.count, perSlideDetail(dry.slides))
    const done = formatByRole(ctx.opened, spec, true)
    const { list, total } = boundedList(done.elements)
    return {
      op,
      after: { count: done.count, byRole: done.byRole, slides: done.slides, elements: list, elementsTotal: total },
    }
  },
})

// ── replaceFont ─────────────────────────────────────────────────────────

function fontNames(op: Op): { from: string; to: string } {
  const ok = (v: unknown) => typeof v === 'string' && !!v.trim() && v.length <= 64
  if (!ok(op.from) || !ok(op.to)) {
    throw new GuidedError('op "replaceFont" needs "from" and "to": font family names (at most 64 characters).')
  }
  if (String(op.from).trim().toLowerCase() === String(op.to).trim().toLowerCase()) {
    throw new GuidedError('op "replaceFont": "from" and "to" are the same font.')
  }
  return { from: String(op.from).trim(), to: String(op.to).trim() }
}

register({
  name: 'replaceFont',
  validate(op, ctx) {
    fontNames(op)
    resolveSlideScope(ctx, op)
    expectedCount(op)
  },
  apply(op, ctx): OpRecord {
    const spec = { ...fontNames(op), slides: resolveSlideScope(ctx, op) }
    const dry = guided(op.op, () => replaceFontInDeck(ctx.opened, spec, false))
    const found = dry.slideRefs + dry.themeFonts + dry.inherited
    if (found === 0) {
      throw new GuidedError(`op "replaceFont": the font "${spec.from}" is not used${spec.slides ? ' on those slides' : ' in the deck'}. Nothing was changed.`)
    }
    requireCount(op, dry.slideRefs, perSlideDetail(dry.slides))
    const done = guided(op.op, () => replaceFontInDeck(ctx.opened, spec, true))
    return { op, after: { ...done, scope: spec.slides ? 'slides' : 'deck' } }
  },
})

// ── replaceColor ────────────────────────────────────────────────────────

function colorMap(op: Op): Array<{ from: string; to: string }> {
  const pairs = op.map !== undefined ? op.map : [{ from: op.from, to: op.to }]
  if (!Array.isArray(pairs) || pairs.length === 0 || pairs.length > 24) {
    throw new GuidedError('op "replaceColor" needs "from"/"to" or "map": 1-24 {from, to} "#RRGGBB" pairs.')
  }
  const seen = new Set<string>()
  return pairs.map((p, i) => {
    const pair = p as { from?: unknown; to?: unknown }
    for (const k of ['from', 'to'] as const) {
      if (typeof pair?.[k] !== 'string' || !/^#?[0-9A-Fa-f]{6}$/.test(pair[k] as string)) {
        throw new GuidedError(`op "replaceColor": map[${i}].${k} must be "#RRGGBB" (six hex digits, no alpha).`)
      }
    }
    const from = String(pair.from).replace(/^#/, '').toUpperCase()
    if (seen.has(from)) throw new GuidedError(`op "replaceColor": "#${from}" appears twice in "map".`)
    seen.add(from)
    return { from: `#${from}`, to: `#${String(pair.to).replace(/^#/, '').toUpperCase()}` }
  })
}

register({
  name: 'replaceColor',
  validate(op, ctx) {
    colorMap(op)
    resolveSlideScope(ctx, op)
    expectedCount(op)
  },
  apply(op, ctx): OpRecord {
    const spec = { map: colorMap(op), slides: resolveSlideScope(ctx, op) }
    const dry = replaceColorInDeck(ctx.opened, spec, false)
    if (dry.replaced + dry.themeSlots.length + dry.inherited === 0) {
      throw new GuidedError(`op "replaceColor": none of ${spec.map.map((m) => m.from).join(', ')} is used${spec.slides ? ' on those slides' : ' in the deck'}. Nothing was changed.`)
    }
    requireCount(op, dry.replaced, perSlideDetail(dry.slides))
    const done = replaceColorInDeck(ctx.opened, spec, true)
    return { op, after: { ...done, scope: spec.slides ? 'slides' : 'deck' } }
  },
})

// ── geometry matching ───────────────────────────────────────────────────

const MATCH_FIELDS = ['width', 'height', 'left', 'top', 'right', 'bottom', 'centerH', 'centerV'] as const
type MatchField = (typeof MATCH_FIELDS)[number]

function matchFields(op: Op): MatchField[] {
  const raw = typeof op.match === 'string' ? [op.match] : op.match
  if (!Array.isArray(raw) || raw.length === 0) {
    throw new GuidedError(`op "${op.op}" needs "match": any of ${MATCH_FIELDS.join('/')} or the shorthands size/position.`)
  }
  const out = new Set<MatchField>()
  for (const f of raw) {
    if (f === 'size') {
      out.add('width')
      out.add('height')
    } else if (f === 'position') {
      out.add('left')
      out.add('top')
    } else if ((MATCH_FIELDS as readonly string[]).includes(String(f))) out.add(f as MatchField)
    else throw new GuidedError(`op "${op.op}": unknown match field ${JSON.stringify(f)}. Fields: ${MATCH_FIELDS.join(', ')}, size, position.`)
  }
  if (out.has('left') && (out.has('right') || out.has('centerH')))
    throw new GuidedError(`op "${op.op}": match at most one of left/right/centerH.`)
  if (out.has('right') && out.has('centerH')) throw new GuidedError(`op "${op.op}": match at most one of left/right/centerH.`)
  if ((out.has('top') && (out.has('bottom') || out.has('centerV'))) || (out.has('bottom') && out.has('centerV')))
    throw new GuidedError(`op "${op.op}": match at most one of top/bottom/centerV.`)
  return [...out]
}

type Box = { x: number; y: number; cx: number; cy: number }

function matchedBox(el: SlideElement, ref: Box, fields: MatchField[]): Box {
  const o = el.transform.offset
  const b: Box = { x: o.x, y: o.y, cx: o.cx, cy: o.cy }
  if (fields.includes('width')) b.cx = ref.cx
  if (fields.includes('height')) b.cy = ref.cy
  if (fields.includes('left')) b.x = ref.x
  if (fields.includes('right')) b.x = ref.x + ref.cx - b.cx
  if (fields.includes('centerH')) b.x = Math.round(ref.x + ref.cx / 2 - b.cx / 2)
  if (fields.includes('top')) b.y = ref.y
  if (fields.includes('bottom')) b.y = ref.y + ref.cy - b.cy
  if (fields.includes('centerV')) b.y = Math.round(ref.y + ref.cy / 2 - b.cy / 2)
  return b
}

/** Move/resize one top-level element; returns true when its frame changed. */
function applyBox(slide: Slide, el: SlideElement, box: Box): boolean {
  const o = el.transform.offset
  if (box.x === o.x && box.y === o.y && box.cx === o.cx && box.cy === o.cy) return false
  if (el.type === 'table' && (box.cx !== o.cx || box.cy !== o.cy)) {
    // Tables resize through their grid (the frame follows the redistributed sums)
    resizeTable(slide, el.id, box.cx, box.cy)
    const t = el.transform.offset
    el.transform = { ...el.transform, offset: { ...t, x: box.x, y: box.y } }
  } else {
    el.transform = { ...el.transform, offset: { x: box.x, y: box.y, cx: box.cx, cy: box.cy } }
  }
  el.dirtyTransform = true
  updateConnectorsForMoved(slide, [el.id])
  return true
}

const refBox = (el: SlideElement): Box => ({ ...el.transform.offset })

function topLevelEls(op: Op, slide: Slide, index: number): SlideElement[] {
  const ids = op.els
  if (!Array.isArray(ids) || ids.length < 1) throw new GuidedError(`op "${op.op}" needs "els": element ids to change.`)
  const els = ids.map((raw) => {
    const id = String(raw)
    const el = slide.elements.find((x) => matchesElementRef(x, id))
    if (!el) {
      const inGroup = slide.elements.some((g) => g.type === 'group' && (g as { children: SlideElement[] }).children.some((c) => matchesElementRef(c, id)))
      throw new GuidedError(
        inGroup
          ? `op "${op.op}": "${id}" is inside a group; match the group itself, or ungroupElement first.`
          : `op "${op.op}": no element "${id}" on slide ${index}. Available: [${slide.elements.map((x) => elementDurableId(x) ?? x.id).join(', ')}].`,
      )
    }
    return el
  })
  if (new Set(els).size !== els.length) throw new GuidedError(`op "${op.op}": "els" lists the same element twice.`)
  return els
}

// matchElements: target = the reference element; els = the elements to change on the same slide.
register({
  name: 'matchElements',
  validate(op, ctx) {
    matchFields(op)
    const { index, slide, el } = resolveElement(ctx, op)
    const els = topLevelEls(op, slide, index)
    if (els.includes(el)) throw new GuidedError('op "matchElements": "els" must not include the reference element (target.el).')
  },
  apply(op, ctx): OpRecord {
    const fields = matchFields(op)
    const { index, slide, el: ref } = resolveElement(ctx, op)
    const els = topLevelEls(op, slide, index)
    const box = refBox(ref)
    const before = els.map((el) => ({ id: elementDurableId(el) ?? el.id, ...el.transform.offset }))
    let changed = 0
    for (const el of els) if (applyBox(slide, el, matchedBox(el, box, fields))) changed++
    return { op, before, after: { fields, changed, unchanged: els.length - changed } }
  },
})

// matchAcrossSlides: target = the reference element; every slide in scope gets its
// single element of the same role matched to it.
function roleOf(op: Op, ref: SlideElement): string {
  if (op.role === undefined) {
    const r = elementRole(ref)
    if (r === 'other') throw new GuidedError('op "matchAcrossSlides": the reference element has no text role; pass "role" explicitly.')
    return r
  }
  const roles = guided(op.op, () => normalizeRoles(op.role, op.op))
  if (roles.length !== 1 || roles[0] === 'all') throw new GuidedError('op "matchAcrossSlides": "role" must be one role (title/subtitle/body/footer/table/text).')
  return roles[0]!
}

function plan(op: Op, ctx: OpContext) {
  const fields = matchFields(op)
  const { index: refIndex, el: ref } = resolveElement(ctx, op)
  const role = roleOf(op, ref)
  const scope = resolveSlideScope(ctx, op) ?? ctx.opened.deck.slides.map((_, i) => i)
  const box = refBox(ref)
  const targets: Array<{ slide: number; el: SlideElement }> = []
  const skipped: Array<{ slide: number; reason: string }> = []
  for (const i of scope) {
    if (i === refIndex) continue
    const slide = ctx.opened.deck.slides[i]!
    const hits = slide.elements.filter((el) => el.type !== 'group' && elementRole(el) === role)
    if (hits.length === 1) targets.push({ slide: i, el: hits[0]! })
    else skipped.push({ slide: i, reason: hits.length ? `${hits.length} ${role} elements (ambiguous)` : `no ${role} element` })
  }
  return { fields, role, box, targets, skipped }
}

register({
  name: 'matchAcrossSlides',
  validate(op, ctx) {
    plan(op, ctx)
    expectedCount(op)
  },
  apply(op, ctx): OpRecord {
    const p = plan(op, ctx)
    if (!p.targets.length) {
      throw new GuidedError(`op "matchAcrossSlides": no slide in scope has exactly one ${p.role} element. Skipped: ${p.skipped.slice(0, 20).map((s) => `slide ${s.slide} (${s.reason})`).join(', ')}.`)
    }
    requireCount(op, p.targets.length, `${p.skipped.length} slide(s) skipped`)
    let changed = 0
    const slides: number[] = []
    for (const t of p.targets) {
      const slide = ctx.opened.deck.slides[t.slide]!
      if (applyBox(slide, t.el, matchedBox(t.el, p.box, p.fields))) {
        changed++
        slides.push(t.slide)
      }
    }
    return {
      op,
      after: { role: p.role, fields: p.fields, matched: p.targets.length, changed, slides: slides.slice(0, 100), skipped: p.skipped.slice(0, 50) },
    }
  },
})


import {
  builtinLayoutInfos, elementDurableId, getSlideComments, getSlideNotes, listSlideLayouts,
  shouldOfferBuiltinLayouts, slideDurableId, type OpenedPptx,
} from '@genoffice/pptx-engine'
import { listSlideAnimations } from './ops/animation-ops'

/** Read package identities, never infer timeline positions from the UI's filtered list. */
export function readNativeSlideDetails(opened: OpenedPptx, slideIndex: number) {
  if (!Number.isSafeInteger(slideIndex) || slideIndex < 0) throw new Error('slideIndex must be a non-negative integer.')
  const slide = opened.deck.slides[slideIndex]
  if (!slide) throw new Error('The requested slide does not exist.')
  const layouts = listSlideLayouts(opened.archive)
  if (shouldOfferBuiltinLayouts(layouts)) {
    layouts.push(...builtinLayoutInfos(opened.deck.size, new Set(layouts.map((l) => l.name))))
  }
  const animations = listSlideAnimations(slide)
  const comments = getSlideComments(opened.archive, slide.path)
  const notes = getSlideNotes(opened.archive, slide.path)
  return {
    slideId: slideDurableId(slide), layoutPath: slide.layoutPath,
    layouts: layouts.slice(0, 200).map((layout, index) => ({ index, path: layout.path, name: layout.name, placeholders: layout.placeholders.length })),
    layoutCount: layouts.length,
    animations: animations.slice(0, 500), animationCount: animations.length,
    tables: slide.elements.filter((el) => el.type === 'table').slice(0, 100).map((el) => {
      const xml = el.anchor.originalXml
      const attributes = /<a:tblPr\b([^>]*)>/.exec(xml)?.[1] ?? ''
      return { el: elementDurableId(el) ?? el.id,
        styleId: /<a:tableStyleId>([^<]*)<\/a:tableStyleId>/.exec(xml)?.[1] ?? null,
        flags: Object.fromEntries(['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol', 'rtl'].map((name) => [name, new RegExp(`\\b${name}="(?:1|true)"`).test(attributes)])),
      }
    }),
    notes: notes.slice(0, 12000), notesTruncated: notes.length > 12000,
    comments: comments.slice(0, 100), commentCount: comments.length,
  }
}

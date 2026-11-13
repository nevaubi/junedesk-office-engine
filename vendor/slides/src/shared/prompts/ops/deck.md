# Deck ops

> Presentation-wide operations with no target: page size, find and replace, header/footer, theme, sections.

Deck ops carry no `target`. They affect every slide (or the presentation's
metadata), so read the deck outline first and confirm the scope with the user
when the request is ambiguous.

### setSlideSize

`{cx,cy} — EMU page size, whole deck`

Changes the page size of the whole deck. Element positions are not rescaled;
follow up with `execute_slide_script` when content must be re-laid out.
Common sizes: 16:9 is 12192000 x 6858000, 4:3 is 9144000 x 6858000.

```json
{ "op": "setSlideSize", "cx": 12192000, "cy": 6858000 }
```

### findReplace

`{find,replace,regex?,matchCase?,wholeWord?,slides?,includeNotes?,expectedCount?} — whole deck unless slides`

Replaces every match in text boxes, shapes, table cells and direct group
children, including matches that span differently formatted runs of one
paragraph (the replacement takes the format of the run where the match
starts). Matches never cross a paragraph, a line break or a field. Notes are
searched only with `includeNotes`. Find the hits first with `search_deck`
(same find/regex/matchCase/wholeWord), then send its total as
`expectedCount`: a different count fails before anything changes. The
receipt lists the count per slide.

| Field         | Type                        | Notes                                                                                   |
| ------------- | --------------------------- | --------------------------------------------------------------------------------------- |
| find          | non-empty string            | Literal text, or a regex with `regex: true` (max 200 chars, no nested quantifiers)      |
| replace       | string                      | May be empty to delete; with `regex`, `$1`/`$<name>`/`$&` insert captures, `$$` a dollar |
| regex         | boolean                     | Default false                                                                           |
| matchCase     | boolean                     | Default false                                                                           |
| wholeWord     | boolean                     | No letter, digit or underscore on either side                                           |
| slides        | index, id, array or {from,to} | 0-based, inclusive; omit for the whole deck                                           |
| includeNotes  | boolean                     | Also replace in speaker notes (formatting kept)                                         |
| expectedCount | positive integer            | Refuse unless exactly this many matches                                                 |
| firstOnly     | boolean                     | Replace only the first match                                                            |
| slideIndex    | integer                     | Legacy single page (use `slides`); `elementId` restricts to one element on it          |

```json
{ "op": "findReplace", "find": "2023", "replace": "2024", "wholeWord": true, "expectedCount": 3 }
```

```json
{
  "op": "findReplace",
  "find": "(\\d{1,2})/(\\d{1,2})/(\\d{4})",
  "replace": "$3-$1-$2",
  "regex": true,
  "slides": { "from": 0, "to": 1 },
  "includeNotes": true
}
```

Common mistakes

- Guessing `expectedCount`: take it from `search_deck` with the same options, or omit it.
- Expecting a report of zero changes: the op fails when nothing matched, so a "not found" error is normal feedback, not a fault.
- One findReplace per slide: a single op with `slides` (or none) covers every page.


### formatText

`{roles:[title|subtitle|body|footer|table|text|all],slides?,font?:{setFont fields},paragraph?:{align?,lineSpacingPct?,spaceBeforePt?,spaceAfterPt?},expectedCount?} — every element of those roles`

Formats every element of the selected roles in one op: "all titles", "body
text on slides 3-9". Roles come from the placeholder type: `title` (title and
centered title), `subtitle`, `body` (content placeholders), `footer` (date,
footer, slide number), `table` (all cell text), `text` (text boxes and shapes
that are not placeholders, including direct group children); `all` is every
element with text. `font` applies to every run of each element like
`setFont`; `paragraph` applies to every paragraph (not to tables).
`expectedCount` is the number of elements.

| Field         | Type                          | Notes                                           |
| ------------- | ----------------------------- | ----------------------------------------------- |
| roles         | role or array of roles        | See above                                       |
| slides        | index, id, array or {from,to} | 0-based, inclusive; omit for the whole deck     |
| font          | object                        | Same fields as `setFont`; `fontSizePt` in points |
| paragraph     | object                        | `align` left/center/right/justify, spacing in pt |
| expectedCount | integer                       | Refuse unless exactly this many elements match  |

```json
{
  "op": "formatText",
  "roles": ["title"],
  "slides": { "from": 0, "to": 1 },
  "font": { "fontSizePt": 32, "bold": true, "color": "#172E4C" }
}
```

Common mistakes

- One setFont per element across many slides: one formatText does them all with an exact count.
- Passing 1-based page numbers in `slides`: indices are 0-based like every slide reference.


### replaceFont

`{from,to,slides?,expectedCount?} — whole deck incl. theme, masters and layouts unless slides`

Replaces one font with another. Deck-wide, every typeface reference equal to
`from` (Latin, East Asian and complex-script slots independently) is
rewritten in the slides, masters, layouts and presentation defaults, and a
theme heading/body font equal to `from` becomes `to`, so inherited text
follows. With `slides`, every run on those slides whose font is `from` gets
`to` explicitly. `expectedCount` counts the slide references (deck-wide) or
the runs (with slides). The receipt reports runs still using `from`.

```json
{ "op": "replaceFont", "from": "Arial", "to": "Calibri" }
```

Common mistakes

- Using setFont on every element to change a font family deck-wide: replaceFont keeps other formatting and theme links.


### replaceColor

`{from,to}|{map:[{from,to}]},slides?,expectedCount? — "#RRGGBB"; whole deck incl. theme slots unless slides`

Replaces explicit colors (text, fills, lines) equal to `from` with `to`.
Deck-wide, a theme color slot (dk1, lt1, accent1..6, ...) whose color equals
a `from` is rewritten too, so content linked to that theme color follows,
and masters and layouts are recolored. With `slides`, only those slides'
explicit colors change. Charts keep their own colors. `expectedCount` counts
the explicit occurrences on slides.

```json
{ "op": "replaceColor", "map": [{ "from": "#1A73E8", "to": "#172E4C" }] }
```

Common mistakes

- Expecting a chart to recolor: use edit_chart for chart series colors.
- Alpha in colors: six hex digits only.

### applyHeaderFooter

`{settings:{footer?,slideNum?,date?,dateAuto?}} — every slide`

Footer text, slide numbers and date on every page (through the master's
placeholders).

| Field             | Type             | Notes                                         |
| ----------------- | ---------------- | --------------------------------------------- |
| settings.footer   | string or `null` | `null` or empty hides the footer              |
| settings.slideNum | boolean          | Show page numbers                             |
| settings.date     | string or `null` | Fixed date text; `null` or empty hides it     |
| settings.dateAuto | boolean          | Dynamic date field that updates in PowerPoint |

```json
{
  "op": "applyHeaderFooter",
  "settings": { "footer": "Acme Corporation - Confidential", "slideNum": true }
}
```

Common mistakes

- Re-sending identical settings: the op reports "nothing changed" as an error; skip it when the values are already in place.

### setSections

`{sections:[{id,name,slideIndices},…]} — full replace`

Replaces the whole section list. Each section lists the 0-based indices of its
slides in order; every slide should belong to exactly one section. Ids are
GUIDs in braces; reuse existing ids to keep a section's identity, mint a new
GUID for a new one.

```json
{
  "op": "setSections",
  "sections": [
    { "id": "SECTION_ID", "name": "Introduction", "slideIndices": [0] },
    { "id": "{6F1D2B0A-4C3E-4F58-9A21-0D7E5B3C9A11}", "name": "Details", "slideIndices": [1] }
  ]
}
```

Related: `addSection`, `renameSection`, `removeSection`, `moveSection` for single changes.

### applyTheme

`{name,colors:{dk1?,lt1?,dk2?,lt2?,accent1?..accent6?,hlink?,folHlink?},majorFont?,minorFont?} — "#RRGGBB" slots, whole deck`

Rewrites the theme color scheme (and optionally the heading/body fonts) of the
whole deck, remapping explicit colors that matched the old scheme. Pages
without their own background get the new `lt1` color.

| Field     | Type                                  | Notes                                                                                                  |
| --------- | ------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| name      | non-empty string                      | Theme name written into the file                                                                       |
| colors    | object of scheme slots to `"#RRGGBB"` | Slots: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`, `folHlink`; six hex digits, no alpha |
| majorFont | string                                | Heading font                                                                                           |
| minorFont | string                                | Body font                                                                                              |

```json
{
  "op": "applyTheme",
  "name": "Ocean",
  "colors": {
    "dk1": "#0B1F3A",
    "lt1": "#FFFFFF",
    "dk2": "#1F3B5C",
    "lt2": "#E8F0FE",
    "accent1": "#1A73E8",
    "accent2": "#34A853",
    "accent3": "#FBBC04",
    "accent4": "#EA4335",
    "accent5": "#9AA0A6",
    "accent6": "#5F6368",
    "hlink": "#1A73E8",
    "folHlink": "#7B1FA2"
  },
  "majorFont": "Georgia",
  "minorFont": "Calibri"
}
```

Common mistakes

- Alpha in scheme colors (`"#RRGGBBAA"`): theme slots take six hex digits only.
- Passing `colors` as an array: it is a slot-to-color object.

### addSection

`{atSlideIndex,name}`

Starts a new section at the given slide (that slide and the following ones up to
the next section move into it).

```json
{ "op": "addSection", "atSlideIndex": 1, "name": "Appendix" }
```

### renameSection

`{id,name}`

```json
{ "op": "renameSection", "id": "SECTION_ID", "name": "Overview" }
```

### removeSection

`{id} — keeps the slides`

Removes the section header; its slides join the previous section.

```json
{ "op": "removeSection", "id": "SECTION_ID" }
```

### moveSection

`{id,dir:"up"|"down"}`

Swaps the section (with all its slides) with its neighbor.

```json
{ "op": "moveSection", "id": "SECTION_ID", "dir": "down" }
```

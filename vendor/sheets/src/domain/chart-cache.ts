/** OOXML represents a missing numeric point by omitting its index, while
 * ptCount retains the full vector length. Zero placeholders are internal only. */
export function numericChartCache(raw: readonly unknown[]): { values: number[]; blanks: number[] } | null {
  const values: number[] = []
  const blanks: number[] = []
  for (const [index, value] of raw.entries()) {
    // An Excel error is not a valid cache refresh. Keep the previous cache
    // until calculation succeeds; never turn an error into a plotted zero.
    if (typeof value === 'string' && /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|ERROR!)/.test(value)) return null
    const number = typeof value === 'number' ? value :
      typeof value === 'string' && value.trim() !== '' ? Number(value) : NaN
    if (Number.isFinite(number)) values.push(number)
    else { values.push(0); blanks.push(index) }
  }
  return { values, blanks }
}

export function numericChartPoints(values: readonly number[], blanks: readonly number[] = []): string {
  const missing = new Set(blanks)
  if (blanks.some((index) => !Number.isInteger(index) || index < 0 || index >= values.length)) {
    throw new Error('Chart gap indexes must be within the numeric vector.')
  }
  return values.map((value, index) => {
    if (!Number.isFinite(value)) throw new Error('Chart numeric points must be finite.')
    return missing.has(index) ? '' : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`
  }).join('')
}

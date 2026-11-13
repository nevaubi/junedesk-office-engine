import { randomUUID } from 'node:crypto'
import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { blankXlsxBuffer } from '../gateway/csv-import'
/** Each new workbook gets an isolated backing file; no existing document is replaced. */
export async function createEmptyWorkbookPath(root: string): Promise<string> {
 const folder=join(root,randomUUID())
 await mkdir(folder,{recursive:true})
 const path=join(folder,'Untitled.xlsx')
 await writeFile(path,await blankXlsxBuffer('Sheet1'),{flag:'wx'})
 return path
}

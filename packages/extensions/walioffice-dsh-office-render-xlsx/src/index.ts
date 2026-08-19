/**
 * XLSX renderer using exceljs.
 * Replaces WaLiOffice's Rust xlsx_render.rs (using rust_xlsxwriter crate).
 * 
 * @module @walioffice/dsh-office-render-xlsx
 */

import ExcelJS from 'exceljs'
import { resolve } from 'node:path'
import { mkdir } from 'node:fs/promises'
import type { Context } from '@deepseek-ai/cordis'

// ── Types ───────────────────────────────────────────────────────────────────

export interface SheetTable {
  title: string
  headers: string[]
  rows: (string | number)[][]
  summary: string
}

export interface SheetOutput {
  tables: SheetTable[]
}

// ── Render ──────────────────────────────────────────────────────────────────

export async function renderXlsx(ctx: Context, output: SheetOutput): Promise<string> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'WaLiOffice'
  workbook.created = new Date()

  for (let i = 0; i < output.tables.length; i++) {
    const table = output.tables[i]!
    const sheetName = table.title || `Sheet${i + 1}`
    const sheet = workbook.addWorksheet(sheetName.slice(0, 31)) // Excel sheet name limit

    // Add title row
    sheet.addRow([table.title])
    sheet.mergeCells(1, 1, 1, table.headers.length)
    const titleCell = sheet.getCell(1, 1)
    titleCell.font = { bold: true, size: 14 }
    titleCell.alignment = { horizontal: 'center' }

    // Add summary row
    if (table.summary) {
      sheet.addRow([table.summary])
      sheet.mergeCells(2, 1, 2, table.headers.length)
      const summaryCell = sheet.getCell(2, 1)
      summaryCell.font = { italic: true, color: { argb: 'FF64748B' } }
      summaryCell.alignment = { horizontal: 'left' }
    }

    // Empty row
    sheet.addRow([])

    // Header row
    const headerRow = sheet.addRow(table.headers)
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2563EB' },
      }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
      }
    })

    // Data rows
    for (const row of table.rows) {
      const dataRow = sheet.addRow(row)
      dataRow.eachCell(cell => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        }
      })
    }

    // Auto-fit column widths
    sheet.columns.forEach((col, idx) => {
      let maxLen = table!.headers[idx]?.length ?? 10
      for (const row of table!.rows) {
        const cellVal = String(row[idx] ?? '')
        if (cellVal.length > maxLen) {
          maxLen = cellVal.length
        }
      }
      col.width = Math.min(Math.max(maxLen + 4, 10), 50)
    })

    // Freeze header row
    sheet.views = [{ state: 'frozen', ySplit: 4 }]
  }

  // Write file
  const outputDir = resolve(process.cwd(), 'output')
  await mkdir(outputDir, { recursive: true })

  const timestamp = Date.now()
  const firstTitle = output.tables[0]?.title ?? 'spreadsheet'
  const filename = `${firstTitle.replace(/[^\w\u4e00-\u9fff]/g, '_')}_${timestamp}.xlsx`
  const filePath = resolve(outputDir, filename)

  await workbook.xlsx.writeFile(filePath)

  return filePath
}

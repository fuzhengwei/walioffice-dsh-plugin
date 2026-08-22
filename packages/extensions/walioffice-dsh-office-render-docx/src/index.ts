/**
 * DOCX renderer using the `docx` npm library.
 * Replaces WaLiOffice's Rust docx_render.rs (using docx-rs crate).
 * 
 * @module @walioffice/dsh-office-render-docx
 */

import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle,
} from 'docx'
import { resolve } from 'node:path'
import { mkdir, writeFile } from 'node:fs/promises'
import type { Context } from '@deepseek-ai/cordis'

// ── Types ───────────────────────────────────────────────────────────────────

export interface DocSection {
  heading: string
  heading_level: number
  paragraphs: string[]
  bullets: string[]
  table?: DocTable
}

export interface DocTable {
  headers: string[]
  rows: string[][]
}

export interface DocOutput {
  title: string
  sections: DocSection[]
  format: string
  markdown: string
}

// ── Heading level mapping ───────────────────────────────────────────────────

function headingLevel(level: number): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  switch (level) {
    case 1: return HeadingLevel.HEADING_1
    case 2: return HeadingLevel.HEADING_2
    case 3: return HeadingLevel.HEADING_3
    case 4: return HeadingLevel.HEADING_4
    case 5: return HeadingLevel.HEADING_5
    default: return HeadingLevel.HEADING_6
  }
}

// ── Build paragraphs ────────────────────────────────────────────────────────

function buildParagraphs(text: string): Paragraph[] {
  // Split by **bold** and *italic* markers
  const result: Paragraph[] = []
  // Simple inline parsing: detect **bold** and *italic*
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g)
  const runs: TextRun[] = []
  
  for (const part of parts) {
    if (part.startsWith('**') && part.endsWith('**')) {
      runs.push(new TextRun({ text: part.slice(2, -2), bold: true, font: 'Microsoft YaHei', size: 22 }))
    } else if (part.startsWith('*') && part.endsWith('*')) {
      runs.push(new TextRun({ text: part.slice(1, -1), italics: true, font: 'Microsoft YaHei', size: 22 }))
    } else if (part) {
      runs.push(new TextRun({ text: part, font: 'Microsoft YaHei', size: 22 }))
    }
  }
  
  if (runs.length > 0) {
    result.push(new Paragraph({ children: runs, spacing: { after: 180, line: 300 } }))
  }
  
  return result
}

// ── Build table ─────────────────────────────────────────────────────────────

function buildTable(table: DocTable): Table {
  const headerRow = new TableRow({
    tableHeader: true,
    children: table.headers.map(h => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: h, bold: true, color: 'FFFFFF', font: 'Microsoft YaHei', size: 20 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
      })],
      shading: { fill: '2563EB' },
    })),
  })

  const dataRows = table.rows.map(row => new TableRow({
    children: row.map(cell => new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: 'Microsoft YaHei', size: 20 })], spacing: { after: 0, line: 260 } })],
    })),
  }))

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    },
  })
}

// ── Render ──────────────────────────────────────────────────────────────────

export async function renderDocx(ctx: Context, doc: DocOutput): Promise<string> {
  const children: (Paragraph | Table)[] = []

  // Title
  children.push(new Paragraph({
    children: [new TextRun({ text: doc.title, bold: true, color: '1E3A8A', font: 'Microsoft YaHei', size: 32 })],
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { after: 420, line: 360 },
  }))

  // Sections
  for (const section of doc.sections) {
    // Heading
    children.push(new Paragraph({
      children: [new TextRun({ text: section.heading, bold: true, color: section.heading_level === 1 ? '1E3A8A' : '334155', font: 'Microsoft YaHei', size: section.heading_level === 1 ? 28 : 24 })],
      heading: headingLevel(section.heading_level),
      spacing: { before: 360, after: 180, line: 320 },
      keepNext: true,
    }))

    // Paragraphs
    for (const para of section.paragraphs) {
      children.push(...buildParagraphs(para))
    }

    // Bullets
    for (const bullet of section.bullets) {
      children.push(new Paragraph({
        children: [new TextRun({ text: bullet, font: 'Microsoft YaHei', size: 22 })],
        bullet: { level: 0 },
        spacing: { after: 100, line: 280 },
      }))
    }

    // Table
    if (section.table) {
      children.push(buildTable(section.table))
      children.push(new Paragraph({ text: '', spacing: { after: 200 } }))
    }
  }

  const document = new Document({
    creator: 'WaLiOffice',
    title: doc.title,
    description: doc.format,
    sections: [{
      properties: {
        page: {
          margin: { top: 1080, right: 1200, bottom: 1080, left: 1200 },
        },
      },
      children,
    }],
  })

  const buffer = await Packer.toBuffer(document)

  const outputDir = resolve(process.cwd(), 'output')
  await mkdir(outputDir, { recursive: true })
  const filename = `${doc.title.replace(/[^\w\u4e00-\u9fff]/g, '_')}_${Date.now()}.docx`
  const filePath = resolve(outputDir, filename)
  await writeFile(filePath, buffer)

  return filePath
}

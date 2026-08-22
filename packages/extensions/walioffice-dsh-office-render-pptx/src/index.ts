/**
 * PPTX renderer using pptxgenjs.
 * Replaces WaLiOffice's Rust pptx_render.rs (318 lines of hand-written OOXML)
 * with the pptxgenjs library for equivalent output.
 * 
 * @module @walioffice/dsh-office-render-pptx
 */

import PptxGenJS from 'pptxgenjs'
import { resolve } from 'node:path'
import { mkdir } from 'node:fs/promises'
import type { Context } from '@deepseek-ai/cordis'

// ── Types (mirror dsh-tool-ppt) ─────────────────────────────────────────────

export interface Palette {
  name: string
  bg: string
  card: string
  primary: string
  accent: string
  dark: string
}

export interface RenderSlide {
  layout: string
  background?: string
  elements: RenderSlideElement[]
}

export interface RenderSlideElement {
  shape?: {
    x: number; y: number; w: number; h: number
    fill: string
    shapeType: 'rect' | 'roundRect' | 'ellipse' | 'line'
    line?: { color: string; width: number }
  }
  text?: {
    content: string
    x: number; y: number; w: number; h: number
    fontSize: number
    color: string
    bold?: boolean
    align?: 'left' | 'center' | 'right'
    valign?: 'top' | 'middle' | 'bottom'
  }
  table?: {
    x: number; y: number; w: number
    headers: string[]
    rows: string[][]
  }
}

export interface RenderRequest {
  title: string
  slides: RenderSlide[]
  palette: Palette
}

// ── Shape type mapping ──────────────────────────────────────────────────────

function shapeType(pptx: PptxGenJS, type: string): PptxGenJS.ShapeType {
  switch (type) {
    case 'roundRect': return pptx.ShapeType.roundRect
    case 'ellipse': return pptx.ShapeType.ellipse
    case 'line': return pptx.ShapeType.line
    default: return pptx.ShapeType.rect
  }
}

function alignType(align?: string): PptxGenJS.HAlign {
  return (align ?? 'left') as PptxGenJS.HAlign
}

function valignType(valign?: string): PptxGenJS.VAlign {
  return (valign ?? 'top') as PptxGenJS.VAlign
}

function colorValue(color: string | undefined, fallback: string): string {
  const value = (color || fallback).replace('#', '').trim()
  return /^[0-9a-f]{6}$/i.test(value) ? value : fallback
}

// ── Render ──────────────────────────────────────────────────────────────────

/**
 * Render slides to a .pptx file.
 * Returns the absolute file path of the generated file.
 */
export async function renderPptx(ctx: Context, req: RenderRequest): Promise<string> {
  const pptx = new PptxGenJS()
  
  // Set 16:9 layout
  pptx.defineLayout({ name: 'WALI_WIDE', width: 13.33, height: 7.5 })
  pptx.layout = 'WALI_WIDE'

  // Set core properties
  pptx.title = req.title
  pptx.author = 'WaLiOffice'
  pptx.company = 'WaLiOffice'
  pptx.subject = req.title
  pptx.theme = {
    headFontFace: 'Microsoft YaHei',
    bodyFontFace: 'Microsoft YaHei',
  }

  for (const slide of req.slides) {
    const s = pptx.addSlide()

    if (slide.background) {
      s.background = { color: colorValue(slide.background, colorValue(req.palette.bg, 'F8FAFC')) }
    }

    for (const el of slide.elements) {
      // Shape
      if (el.shape) {
        const opts: PptxGenJS.ShapeProps = {
          x: el.shape.x,
          y: el.shape.y,
          w: el.shape.w,
          h: el.shape.h,
          fill: { color: colorValue(el.shape.fill, colorValue(req.palette.card, 'FFFFFF')) },
        }
        if (el.shape.shapeType === 'line') {
          opts.fill = { color: colorValue(el.shape.fill, 'FFFFFF'), transparency: 100 }
          opts.line = {
            color: colorValue(el.shape.line?.color, colorValue(req.palette.primary, '2563EB')),
            width: el.shape.line?.width ?? 1,
          }
        }
        if (el.shape.line) {
          opts.line = {
            color: colorValue(el.shape.line.color, colorValue(req.palette.primary, '2563EB')),
            width: el.shape.line.width,
          }
        }
        s.addShape(shapeType(pptx, el.shape.shapeType), opts)
      }

      // Text
      if (el.text) {
        s.addText(el.text.content, {
          x: el.text.x,
          y: el.text.y,
          w: el.text.w,
          h: el.text.h,
          fontSize: el.text.fontSize,
          color: colorValue(el.text.color, '0F172A'),
          bold: el.text.bold,
          align: alignType(el.text.align),
          valign: valignType(el.text.valign),
          margin: 0.04,
          breakLine: false,
          fit: 'shrink',
          paraSpaceAfter: 0,
          fontFace: 'Microsoft YaHei',
        })
      }

      // Table
      if (el.table) {
        // pptxgenjs TableRow = TableCell[], TableCell = { text?: string, options?: TableCellProps }
        const rows: PptxGenJS.TableRow[] = []
        
        // Header row
        rows.push(el.table.headers.map(h => ({
          text: h,
          options: { bold: true, fill: { color: req.palette.primary.replace('#', '') }, color: 'FFFFFF' },
        })))
        
        // Data rows
        for (const row of el.table.rows) {
          rows.push(row.map(cell => ({ text: cell })))
        }

        s.addTable(rows, {
          x: el.table.x,
          y: el.table.y,
          w: el.table.w,
          border: { type: 'solid', pt: 1, color: 'CBD5E1' },
          fontFace: 'Microsoft YaHei',
          fontSize: 12,
          color: '334155',
          margin: 0.04,
          valign: 'middle',
        })
      }
    }
  }

  // Write file
  const outputDir = resolve(process.cwd(), 'output')
  await mkdir(outputDir, { recursive: true })
  
  const filename = `${req.title.replace(/[^\w\u4e00-\u9fff]/g, '_')}_${Date.now()}.pptx`
  const filePath = resolve(outputDir, filename)
  
  await pptx.writeFile({ fileName: filePath })
  
  return filePath
}

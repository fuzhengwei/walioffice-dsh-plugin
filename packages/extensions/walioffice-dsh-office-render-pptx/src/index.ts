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

  for (const slide of req.slides) {
    const s = pptx.addSlide()

    for (const el of slide.elements) {
      // Shape
      if (el.shape) {
        const opts: PptxGenJS.ShapeProps = {
          x: el.shape.x,
          y: el.shape.y,
          w: el.shape.w,
          h: el.shape.h,
          fill: { color: el.shape.fill.replace('#', '') },
        }
        if (el.shape.line) {
          opts.line = {
            color: el.shape.line.color.replace('#', ''),
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
          color: el.text.color.replace('#', ''),
          bold: el.text.bold,
          align: alignType(el.text.align),
          fontFace: 'Arial',
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
          border: { type: 'solid', pt: 1, color: 'CCCCCC' },
          fontFace: 'Arial',
          fontSize: 12,
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

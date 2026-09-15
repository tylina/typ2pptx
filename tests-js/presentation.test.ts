import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, test } from 'vitest'

import { createEditablePptx, type PresentationExportModel } from '../src/index'

const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
const PNG_1X1_ALT =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

describe('browser-compatible editable presentation writer', () => {
  test('emits native SVG geometry, editable text, links, and residual fallback once', async () => {
    const model: PresentationExportModel = {
      pages: [{
        pageIndex: 0,
        width: 960,
        height: 540,
        nonTextSvg: `
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">
            <path d="M10 10 L30 10 L30 30 Z" fill="#112233"/>
            <path d="M40 10 L60 10 L60 30 Z" fill="#445566" filter="url(#shadow)"/>
          </svg>
        `,
        nonTextPngBase64: PNG_1X1,
        elements: [{
          kind: 'rectangle', x: 70, y: 10, width: 40, height: 20, fillColor: '#abcdef'
        }, {
          kind: 'text', x: 70, y: 40, width: 220, height: 40,
          text: 'Editable evidence', fontFamily: 'Aptos', fontSize: 20,
          color: '#123456', bold: true, italic: false
        }, {
          kind: 'link', x: 70, y: 40, width: 220, height: 40,
          url: 'https://example.com/evidence'
        }, {
          kind: 'slideLink', x: 70, y: 90, width: 100, height: 30, pageIndex: 1
        }],
        fallbackTextCount: 0,
        fallbackShapeCount: 2
      }, {
        pageIndex: 1,
        width: 960,
        height: 540,
        nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540"/>',
        nonTextPngBase64: PNG_1X1,
        elements: [],
        fallbackTextCount: 0,
        fallbackShapeCount: 0
      }],
      fonts: ['Aptos'],
      editableTextCount: 1,
      editableShapeCount: 1,
      fallbackTextCount: 0,
      fallbackShapeCount: 2,
      warnings: []
    }

    const rasterized: string[] = []
    const result = await createEditablePptx({
      title: 'Portable deck',
      model,
      notesByPageIndex: new Map([[0, 'Explain the evidence.']]),
      residualSvgRasterizer: async ({ svg, width, height }) => {
        rasterized.push(svg)
        expect({ width, height }).toEqual({ width: 960, height: 540 })
        return PNG_1X1_ALT
      }
    })
    const files = unzipSync(result.bytes)
    const slide = strFromU8(files['ppt/slides/slide1.xml'])
    const relationships = strFromU8(files['ppt/slides/_rels/slide1.xml.rels'])
    const residualEntry = Object.entries(files).find(([name]) =>
      name.startsWith('ppt/media/typ2pptx-residual-') && name.endsWith('.svg'))

    expect(slide).toContain('SVG path')
    expect(slide).toContain('<a:custGeom>')
    expect(slide).toContain('Editable evidence')
    expect(slide).toContain('Editable Typst rectangle')
    expect(slide).not.toContain('typ2pptx-vector-anchor')
    expect(relationships).toContain('https://example.com/evidence')
    expect(slide).toContain('ppaction://hlinksldjump')
    expect(residualEntry && strFromU8(residualEntry[1])).not.toContain('112233')
    expect(residualEntry && strFromU8(residualEntry[1])).toContain('filter="url(#shadow)"')
    expect(rasterized).toHaveLength(1)
    expect(rasterized[0]).not.toContain('112233')
    expect(rasterized[0]).toContain('filter="url(#shadow)"')
    expect(Object.values(files).some((bytes) =>
      Buffer.from(bytes).equals(Buffer.from(PNG_1X1_ALT, 'base64')))).toBe(true)
    expect(Object.values(files).some((bytes) =>
      Buffer.from(bytes).equals(Buffer.from(PNG_1X1, 'base64')))).toBe(false)
    expect(Object.entries(files).some(([name, bytes]) =>
      name.startsWith('ppt/notesSlides/notesSlide') &&
      strFromU8(bytes).includes('Explain the evidence.'))).toBe(true)
    expect(result).toMatchObject({
      pageCount: 2,
      editableTextCount: 1,
      editableShapeCount: 2,
      editableLinkCount: 2,
      fallbackTextCount: 0,
      fallbackShapeCount: 2,
      nativeVectorShapeCount: 1
    })
  })

  test('does not add a page-sized image when every SVG item became native', async () => {
    const result = await createEditablePptx({
      title: 'Native deck',
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg"><circle cx="20" cy="20" r="10"/></svg>',
          nonTextPngBase64: PNG_1X1,
          elements: [],
          fallbackTextCount: 0,
          fallbackShapeCount: 1
        }],
        fonts: [],
        editableTextCount: 0,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 1,
        warnings: []
      }
    })
    const files = unzipSync(result.bytes)
    const slide = strFromU8(files['ppt/slides/slide1.xml'])

    expect(slide).toContain('SVG ellipse')
    expect(slide).not.toContain('<p:pic>')
    expect(result.fallbackShapeCount).toBe(0)
  })

  test('keeps residual pictures and native geometry in original SVG paint order', async () => {
    const result = await createEditablePptx({
      title: 'Layered deck',
      residualSvgRasterizer: async () => PNG_1X1_ALT,
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          nonTextSvg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
            <path d="M0 0 L20 0 L20 20 Z" fill="#aa0000" filter="url(#shadow)"/>
            <path d="M30 0 L50 0 L50 20 Z" fill="#00aa00"/>
            <path d="M60 0 L80 0 L80 20 Z" fill="#0000aa" clip-path="url(#clip)"/>
          </svg>`,
          nonTextPngBase64: PNG_1X1,
          elements: [],
          fallbackTextCount: 0,
          fallbackShapeCount: 3
        }],
        fonts: [],
        editableTextCount: 0,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 3,
        warnings: []
      }
    })
    const files = unzipSync(result.bytes)
    const slide = strFromU8(files['ppt/slides/slide1.xml'])
    const firstResidual = slide.indexOf('typ2pptx residual layer 1.1')
    const native = slide.indexOf('SVG path')
    const secondResidual = slide.indexOf('typ2pptx residual layer 1.3')

    expect(firstResidual).toBeGreaterThan(0)
    expect(native).toBeGreaterThan(firstResidual)
    expect(secondResidual).toBeGreaterThan(native)
    expect(strFromU8(files['ppt/media/typ2pptx-residual-1-1.svg'])).toContain('#aa0000')
    expect(strFromU8(files['ppt/media/typ2pptx-residual-1-3.svg'])).toContain('#0000aa')
  })

  test('keeps compiler-rendered images as independently editable pictures', async () => {
    const result = await createEditablePptx({
      title: 'Image deck',
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg"/>',
          nonTextPngBase64: PNG_1X1,
          elements: [{
            kind: 'image', x: 10, y: 8, width: 30, height: 20,
            mediaType: 'image/png', dataBase64: PNG_1X1, altText: 'A measured result'
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: [],
        editableTextCount: 0,
        editableShapeCount: 1,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const files = unzipSync(result.bytes)
    const slide = strFromU8(files['ppt/slides/slide1.xml'])

    expect(slide).toContain('Editable Typst image')
    expect(slide).toContain('A measured result')
    expect(slide).toContain('<p:pic>')
    expect(Object.keys(files).some((name) =>
      name.startsWith('ppt/media/image-') && name.endsWith('.png'))).toBe(true)
    expect(result).toMatchObject({ editableShapeCount: 1, fallbackShapeCount: 0 })
  })

  test('preserves a compiler-provided JPEG without converting it to PNG', async () => {
    const result = await createEditablePptx({
      title: 'JPEG deck',
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg"/>',
          nonTextPngBase64: PNG_1X1,
          elements: [{
            kind: 'image', x: 10, y: 8, width: 30, height: 20,
            mediaType: 'image/jpeg', dataBase64: '/9j/2Q==', altText: 'Original JPEG'
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: [],
        editableTextCount: 0,
        editableShapeCount: 1,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })

    const files = unzipSync(result.bytes)
    expect(Object.keys(files).some((name) =>
      name.startsWith('ppt/media/image-') && name.endsWith('.jpeg'))).toBe(true)
    expect(Object.values(files).some((bytes) =>
      bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff)).toBe(true)
  })

  test('groups adjacent styled runs on one baseline into one editable textbox', async () => {
    const result = await createEditablePptx({
      title: 'Styled text',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg"/>',
          nonTextPngBase64: PNG_1X1,
          elements: [{
            kind: 'text', x: 20, y: 30, width: 58, height: 18, baseline: 44,
            text: 'Evidence ', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false
          }, {
            kind: 'text', x: 78, y: 30, width: 52, height: 18, baseline: 44,
            text: 'matters', fontFamily: 'Aptos', fontSize: 16,
            color: '#654321', bold: true, italic: true
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos'],
        editableTextCount: 2,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(1)
    expect(slide).toContain('Evidence ')
    expect(slide).toContain('matters')
    expect(slide).toContain('b="1"')
    expect(slide).toContain('i="1"')
    expect(slide).toContain('123456')
    expect(slide).toContain('654321')
    expect(result.editableTextCount).toBe(2)
  })

  test('keeps inline math scripts with prose and writes native baseline offsets', async () => {
    const result = await createEditablePptx({
      title: 'Inline math',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          nonTextSvg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 135">
            <defs><symbol id="glyph"><path d="M0 0 L6 0 L6 8 L0 8 Z"/></symbol></defs>
            <g transform="translate(170 80)"><use href="#glyph"/></g>
            <g transform="translate(180 80)"><use href="#glyph"/></g>
          </svg>`,
          nonTextPngBase64: PNG_1X1,
          elements: [{
            kind: 'text', x: 20, y: 30, width: 38, height: 18, baseline: 44,
            text: 'Area ', fontFamily: 'Aptos', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 58, y: 30, width: 9, height: 18, baseline: 44,
            text: '𝑟', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 67, y: 25, width: 6, height: 12, baseline: 36,
            text: '2', fontFamily: 'Cambria Math', fontSize: 11,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 76, y: 30, width: 9, height: 18, baseline: 44,
            text: '+', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 88, y: 30, width: 18, height: 18, baseline: 44,
            text: ' H', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 106, y: 37, width: 6, height: 12, baseline: 50,
            text: '2', fontFamily: 'Cambria Math', fontSize: 11,
            color: '#000000', bold: false, italic: false
          }, {
            kind: 'text', x: 112, y: 30, width: 42, height: 18, baseline: 44,
            text: ' done', fontFamily: 'Aptos', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }],
          vectorGroups: [{
            id: 'formula-1', kind: 'math', x: 168, y: 78, width: 30, height: 15
          }],
          fallbackTextCount: 2,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos', 'Cambria Math'],
        editableTextCount: 7,
        editableShapeCount: 0,
        fallbackTextCount: 2,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(1)
    expect(slide).toContain('Area ')
    expect(slide).toContain(' done')
    expect(slide).toContain('<a:t> +</a:t>')
    expect(slide).toMatch(/baseline="[1-9][0-9]{4}"/u)
    expect(slide).toMatch(/baseline="-[1-9][0-9]{4}"/u)
    expect(slide).toContain('Typst math formula-1')
    expect(slide).toContain('<p:grpSp>')
  })

  test('writes an explicit East Asian typeface for editable CJK text', async () => {
    const result = await createEditablePptx({
      title: 'CJK text',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          nonTextSvg: '<svg xmlns="http://www.w3.org/2000/svg"/>',
          nonTextPngBase64: PNG_1X1,
          elements: [{
            kind: 'text', x: 20, y: 30, width: 80, height: 20,
            text: '研究结论', fontFamily: 'Noto Sans CJK SC', fontSize: 16,
            color: '#000000', bold: false, italic: false
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Noto Sans CJK SC'],
        editableTextCount: 1,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide).toContain('<a:ea typeface="Noto Sans CJK SC"')
  })
})

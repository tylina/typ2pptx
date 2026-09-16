import { DOMParser } from '@xmldom/xmldom'
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
        fallbackLayers: [{ kind: 'fallback', paintOrder: 0, svg: `
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">
            <path d="M10 10 L30 10 L30 30 Z" fill="#112233"/>
            <path d="M40 10 L60 10 L60 30 Z" fill="#445566" filter="url(#shadow)"/>
          </svg>
        ` }],
        elements: [{
          kind: 'rectangle', paintOrder: 1,
          x: 70, y: 10, width: 40, height: 20, fillColor: '#abcdef'
        }, {
          kind: 'text', paintOrder: 2, x: 70, y: 40, width: 220, height: 40,
          text: 'Editable evidence', fontFamily: 'Aptos', fontSize: 20, baseline: 64,
          color: '#123456', bold: true, italic: false, rtl: false, textBox: null
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
        fallbackLayers: [],
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
    expect(() => new DOMParser().parseFromString(slide, 'application/xml')).not.toThrow()
    expect(slide.match(/xmlns:asvg=/g) ?? []).toHaveLength(1)
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
          fallbackLayers: [{
            kind: 'fallback', paintOrder: 0,
            svg: '<svg xmlns="http://www.w3.org/2000/svg"><circle cx="20" cy="20" r="10"/></svg>'
          }],
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

  test('emits compiler page paint as a non-selectable slide background', async () => {
    let rasterizations = 0
    const result = await createEditablePptx({
      title: 'Slide backgrounds',
      residualSvgRasterizer: async ({ svg }) => {
        rasterizations += 1
        expect(svg).toContain('linearGradient')
        return PNG_1X1_ALT
      },
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          background: { kind: 'solid', color: '#fefefe' },
          fallbackLayers: [],
          elements: [],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }, {
          pageIndex: 1,
          width: 100,
          height: 60,
          background: {
            kind: 'svg',
            svg: '<svg xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="g"/></defs></svg>'
          },
          fallbackLayers: [],
          elements: [],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: [],
        editableTextCount: 0,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const files = unzipSync(result.bytes)
    const solidSlide = strFromU8(files['ppt/slides/slide1.xml'])
    const vectorSlide = strFromU8(files['ppt/slides/slide2.xml'])

    expect(solidSlide).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FEFEFE"')
    expect(solidSlide).not.toContain('<p:sp>')
    expect(vectorSlide).toContain('<p:bg><p:bgPr><a:blipFill')
    expect(vectorSlide).not.toContain('<p:pic>')
    expect(rasterizations).toBe(1)
    expect(Object.entries(files).some(([name, bytes]) =>
      name.startsWith('ppt/media/') &&
      name.endsWith('.png') &&
      Buffer.from(bytes).equals(Buffer.from(PNG_1X1_ALT, 'base64')))).toBe(true)
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
          fallbackLayers: [{ kind: 'fallback', paintOrder: 0, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
            <path d="M0 0 L20 0 L20 20 Z" fill="#aa0000" filter="url(#shadow)"/>
            <path d="M30 0 L50 0 L50 20 Z" fill="#00aa00"/>
            <path d="M60 0 L80 0 L80 20 Z" fill="#0000aa" clip-path="url(#clip)"/>
          </svg>` }],
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

  test('interleaves compiler fallback layers and editable objects by paint order', async () => {
    const result = await createEditablePptx({
      title: 'Compiler paint order',
      residualSvgRasterizer: async () => PNG_1X1_ALT,
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          fallbackLayers: [{
            kind: 'fallback', paintOrder: 0,
            svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
              <path d="M0 0 L20 0 L20 20 Z" fill="#aa0000" filter="url(#shadow)"/>
            </svg>`
          }, {
            kind: 'fallback', paintOrder: 2,
            svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
              <path d="M30 0 L50 0 L50 20 Z" fill="#00aa00"/>
            </svg>`
          }],
          elements: [{
            kind: 'text', paintOrder: 1, x: 10, y: 25, width: 50, height: 15,
            text: 'Between layers', fontFamily: 'Aptos', fontSize: 12, baseline: 37,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: null
          }, {
            kind: 'rectangle', paintOrder: 3,
            x: 60, y: 25, width: 20, height: 15, fillColor: '#abcdef'
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 2
        }],
        fonts: ['Aptos'],
        editableTextCount: 1,
        editableShapeCount: 1,
        fallbackTextCount: 0,
        fallbackShapeCount: 2,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])
    const residual = slide.indexOf('typ2pptx residual layer 1.1')
    const text = slide.indexOf('Between layers')
    const native = slide.indexOf('SVG path')
    const rectangle = slide.indexOf('Editable Typst rectangle')

    expect(residual).toBeGreaterThan(0)
    expect(text).toBeGreaterThan(residual)
    expect(native).toBeGreaterThan(text)
    expect(rectangle).toBeGreaterThan(native)
  })

  test('keeps compiler-rendered images as independently editable pictures', async () => {
    const result = await createEditablePptx({
      title: 'Image deck',
      model: {
        pages: [{
          pageIndex: 0,
          width: 100,
          height: 60,
          fallbackLayers: [],
          elements: [{
            kind: 'image', paintOrder: 0, x: 10, y: 8, width: 30, height: 20,
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
          fallbackLayers: [],
          elements: [{
            kind: 'image', paintOrder: 0, x: 10, y: 8, width: 30, height: 20,
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

  test('groups compiler-owned styled runs into one editable text box', async () => {
    const textBox = {
      id: 'paragraph-styled-line-0', paragraphId: 'paragraph-styled', lineIndex: 0,
      x: 20, width: 110, alignment: 'left' as const, reflow: false
    }
    const result = await createEditablePptx({
      title: 'Styled text',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 20, y: 30, width: 58, height: 18, baseline: 44,
            text: 'Evidence ', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 1, x: 78, y: 30, width: 52, height: 18, baseline: 44,
            text: 'matters', fontFamily: 'Aptos', fontSize: 16,
            color: '#654321', bold: true, italic: true, rtl: false, textBox
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
    expect(slide.match(/<a:pPr /gu)).toHaveLength(1)
    expect(result.editableTextCount).toBe(2)
  })

  test('does not geometry-merge compiler-isolated table cell text', async () => {
    const result = await createEditablePptx({
      title: 'Isolated table cells',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 16, y: 14, width: 45, height: 13, baseline: 23,
            text: 'Alpha cell', fontFamily: 'Aptos', fontSize: 11,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: null
          }, {
            kind: 'text', paintOrder: 1, x: 77, y: 14, width: 38, height: 13, baseline: 23,
            text: 'Beta cell', fontFamily: 'Aptos', fontSize: 11,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: null
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

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(2)
    expect(slide).not.toContain('<a:spAutoFit/>')
  })

  test('uses compiler-owned line boxes and alignment instead of coordinate guessing', async () => {
    const textBox = {
      id: 'paragraph-a-line-0',
      paragraphId: 'paragraph-a',
      lineIndex: 0,
      x: 12,
      width: 216,
      alignment: 'center' as const,
      reflow: false
    }
    const result = await createEditablePptx({
      title: 'Compiler-owned text line',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 72, y: 30, width: 48, height: 18, baseline: 44,
            text: 'Owned', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 1, x: 130, y: 30, width: 48, height: 18, baseline: 44,
            text: 'line', fontFamily: 'Aptos', fontSize: 16,
            color: '#654321', bold: true, italic: false, rtl: false, textBox
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
    expect(slide).toContain('<a:pPr algn="ctr"')
    expect(slide).toContain('<a:off x="152400" y="365760"')
    expect(slide).toContain('<a:ext cx="2743200"')
    expect(slide).toContain('<a:t>line</a:t>')
    expect(slide).not.toContain('<a:t> line</a:t>')

    const justified = await createEditablePptx({
      title: 'Compiler-owned justified line',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 216, height: 18, baseline: 44,
            text: 'Words fill this line', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false,
            rtl: false,
            textBox: { ...textBox, alignment: 'justify' }
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos'],
        editableTextCount: 1,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const justifiedSlide = strFromU8(unzipSync(justified.bytes)['ppt/slides/slide1.xml'])
    expect(justifiedSlide).toContain('<a:pPr algn="just"')
    expect(justifiedSlide).not.toContain('<a:br/>')
  })

  test('keeps proven physical lines in one fixed-layout paragraph object', async () => {
    const firstLine = {
      id: 'paragraph-fixed-line-0', paragraphId: 'paragraph-fixed', lineIndex: 0,
      x: 12, width: 216, alignment: 'left' as const, reflow: false
    }
    const secondLine = { ...firstLine, id: 'paragraph-fixed-line-1', lineIndex: 1 }
    const result = await createEditablePptx({
      title: 'Compiler-owned fixed lines',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 42, height: 18, baseline: 44,
            text: 'First ', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox: firstLine
          }, {
            kind: 'text', paintOrder: 1, x: 54, y: 30, width: 70, height: 18, baseline: 44,
            text: 'styled line', fontFamily: 'Aptos', fontSize: 16,
            color: '#654321', bold: true, italic: false, rtl: false, textBox: firstLine
          }, {
            kind: 'text', paintOrder: 2, x: 12, y: 50, width: 150, height: 18, baseline: 64,
            text: 'Second compiler line', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox: secondLine
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos'], editableTextCount: 3, editableShapeCount: 0,
        fallbackTextCount: 0, fallbackShapeCount: 0, warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(1)
    expect(slide).toContain('<a:br/>')
    expect(slide.match(/<a:pPr /gu)).toHaveLength(1)
    expect(slide).toContain('<a:spcPts val="2000"')
    expect(slide).toContain('<a:bodyPr wrap="none"')
    expect(slide).toContain('<a:t>Second compiler line</a:t>')
  })

  test('preserves compiler-owned blank code lines inside one copyable object', async () => {
    const firstLine = {
      id: 'raw-code-line-0', paragraphId: 'raw-code', lineIndex: 0,
      x: 12, width: 216, alignment: 'left' as const, reflow: false
    }
    const thirdLine = { ...firstLine, id: 'raw-code-line-2', lineIndex: 2 }
    const result = await createEditablePptx({
      title: 'Code with a blank line',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 70, height: 18, baseline: 44,
            text: 'let x = 1', fontFamily: 'Aptos Mono', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox: firstLine
          }, {
            kind: 'text', paintOrder: 1, x: 12, y: 70, width: 70, height: 18, baseline: 84,
            text: 'x + 1', fontFamily: 'Aptos Mono', fontSize: 16,
            color: '#654321', bold: true, italic: false, rtl: false, textBox: thirdLine
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos Mono'], editableTextCount: 2, editableShapeCount: 0,
        fallbackTextCount: 0, fallbackShapeCount: 0, warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(1)
    expect(slide.match(/<a:br\/>/gu)).toHaveLength(2)
    expect(slide).toContain('<a:spcPts val="2000"')
    expect(slide).toContain('<a:t>let x = 1</a:t>')
    expect(slide).toContain('<a:t>x + 1</a:t>')
  })

  test('bounds blank-line expansion across the complete presentation', async () => {
    const rawLine = {
      id: 'raw-code-line-0', paragraphId: 'raw-code', lineIndex: 0,
      x: 12, width: 216, alignment: 'left' as const, reflow: false
    }
    await expect(createEditablePptx({
      title: 'Sparse raw lines',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 10, width: 20, height: 12,
            baseline: 20, text: 'A', fontFamily: 'Aptos', fontSize: 10,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: null
          }, {
            kind: 'text', paintOrder: 1, x: 12, y: 30, width: 30, height: 12,
            baseline: 40, text: 'B', fontFamily: 'Aptos Mono', fontSize: 10,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: rawLine
          }, {
            kind: 'text', paintOrder: 2, x: 12, y: 50, width: 30, height: 12,
            baseline: 60, text: 'C', fontFamily: 'Aptos Mono', fontSize: 10,
            color: '#000000', bold: false, italic: false, rtl: false,
            textBox: { ...rawLine, id: 'raw-code-line-99999', lineIndex: 99_999 }
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos', 'Aptos Mono'], editableTextCount: 3, editableShapeCount: 0,
        fallbackTextCount: 0, fallbackShapeCount: 0, warnings: []
      }
    })).rejects.toThrow('generated-text-run budget')
  })

  test('keeps unequal indented line boxes independent', async () => {
    const result = await createEditablePptx({
      title: 'Indented compiler lines',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 30, y: 30, width: 150, height: 18, baseline: 44,
            text: 'Indented first line', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false,
            rtl: false,
            textBox: {
              id: 'paragraph-indent-line-0', paragraphId: 'paragraph-indent', lineIndex: 0,
              x: 30, width: 198, alignment: 'left', reflow: false
            }
          }, {
            kind: 'text', paintOrder: 1, x: 12, y: 50, width: 170, height: 18, baseline: 64,
            text: 'Unindented second line', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false,
            rtl: false,
            textBox: {
              id: 'paragraph-indent-line-1', paragraphId: 'paragraph-indent', lineIndex: 1,
              x: 12, width: 216, alignment: 'left', reflow: false
            }
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos'], editableTextCount: 2, editableShapeCount: 0,
        fallbackTextCount: 0, fallbackShapeCount: 0, warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(2)
    expect(slide).not.toContain('<a:br/>')
  })

  test('keeps a compiler-owned paragraph coherent and natively reflowable', async () => {
    const result = await createEditablePptx({
      title: 'Compiler-owned native paragraph',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 216, height: 48,
            baseline: 44, text: 'Native paragraphs remain coherent when edited in PowerPoint.',
            fontFamily: 'Aptos', fontSize: 16, color: '#123456', bold: false, italic: false,
            rtl: false,
            textBox: {
              id: 'paragraph-a-paragraph', paragraphId: 'paragraph-a', lineIndex: 0,
              x: 12, width: 216, alignment: 'justify', reflow: true, lineSpacing: 14.4
            }
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos'], editableTextCount: 1, editableShapeCount: 0,
        fallbackTextCount: 0, fallbackShapeCount: 0, warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(1)
    expect(slide).toContain('<a:off x="152400" y="365760"')
    expect(slide).toContain('<a:pPr algn="just"')
    expect(slide).toContain('<a:spcPts val="1440"')
    expect(slide).toContain('Native paragraphs remain coherent')
    expect(slide).not.toContain('<a:br/>')
  })

  test('rejects a compiler-owned text line interrupted in paint order', async () => {
    const textBox = {
      id: 'paragraph-a-line-0',
      paragraphId: 'paragraph-a',
      lineIndex: 0,
      x: 12,
      width: 216,
      alignment: 'left' as const,
      reflow: false
    }
    const create = createEditablePptx({
      title: 'Interrupted compiler line',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 48, height: 18, baseline: 44,
            text: 'Not ', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'rectangle', paintOrder: 1, x: 70, y: 30, width: 10, height: 10,
            fillColor: '#abcdef'
          }, {
            kind: 'text', paintOrder: 2, x: 80, y: 30, width: 48, height: 18, baseline: 44,
            text: 'contiguous', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos'],
        editableTextCount: 2,
        editableShapeCount: 1,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })

    await expect(create).rejects.toThrow('must be contiguous')
  })

  test('rejects a compiler-owned text run outside its retained line box', async () => {
    const textBox = {
      id: 'paragraph-a-line-0',
      paragraphId: 'paragraph-a',
      lineIndex: 0,
      x: 12,
      width: 100,
      alignment: 'left' as const,
      reflow: false
    }
    const create = createEditablePptx({
      title: 'Invalid compiler line geometry',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 130, y: 30, width: 48, height: 18, baseline: 44,
            text: 'Outside', fontFamily: 'Aptos', fontSize: 16,
            color: '#123456', bold: false, italic: false, rtl: false, textBox
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos'],
        editableTextCount: 1,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })

    await expect(create).rejects.toThrow('outside its compiler-owned text box')
  })

  test('requires a finite exact baseline for compiler-positioned text', async () => {
    const create = createEditablePptx({
      title: 'Missing compiler baseline',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 12, y: 30, width: 48, height: 18,
            text: 'Invalid', fontFamily: 'Aptos', fontSize: 16, baseline: Number.NaN,
            color: '#123456', bold: false, italic: false,
            rtl: false,
            textBox: {
              id: 'paragraph-a-line-0', paragraphId: 'paragraph-a', lineIndex: 0,
              x: 12, width: 100, alignment: 'left', reflow: false
            }
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Aptos'],
        editableTextCount: 1,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })

    await expect(create).rejects.toThrow('requires an exact baseline')
  })

  test('keeps inline math scripts with prose and writes native baseline offsets', async () => {
    const textBox = {
      id: 'paragraph-math-line-0', paragraphId: 'paragraph-math', lineIndex: 0,
      x: 20, width: 134, alignment: 'left' as const, reflow: false
    }
    const result = await createEditablePptx({
      title: 'Inline math',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [{ kind: 'fallback', paintOrder: 7, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 135">
            <defs><symbol id="glyph"><path d="M0 0 L6 0 L6 8 L0 8 Z"/></symbol></defs>
            <g transform="translate(170 80)"><use href="#glyph"/></g>
            <g transform="translate(180 80)"><use href="#glyph"/></g>
          </svg>` }],
          elements: [{
            kind: 'text', paintOrder: 0, x: 20, y: 30, width: 38, height: 18, baseline: 44,
            text: 'Area ', fontFamily: 'Aptos', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 1, x: 58, y: 30, width: 9, height: 18, baseline: 44,
            text: '𝑟', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 2, x: 67, y: 25, width: 6, height: 12, baseline: 36,
            text: '2', fontFamily: 'Cambria Math', fontSize: 11,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 3, x: 76, y: 30, width: 9, height: 18, baseline: 44,
            text: '+', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 4, x: 88, y: 30, width: 18, height: 18, baseline: 44,
            text: ' H', fontFamily: 'Cambria Math', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 5, x: 106, y: 37, width: 6, height: 12, baseline: 50,
            text: '2', fontFamily: 'Cambria Math', fontSize: 11,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 6, x: 112, y: 30, width: 42, height: 18, baseline: 44,
            text: ' done', fontFamily: 'Aptos', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
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
    expect(slide).toContain('<a:t>+</a:t>')
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
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 20, y: 30, width: 80, height: 20, baseline: 46,
            text: '研究结论', fontFamily: 'Noto Sans CJK SC', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox: null
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

  test('writes compiler-provided RTL direction on an isolated editable run', async () => {
    const result = await createEditablePptx({
      title: 'RTL text',
      model: {
        pages: [{
          pageIndex: 0,
          width: 240,
          height: 135,
          fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 80, y: 30, width: 140, height: 20,
            baseline: 46, text: 'مرحبا بالعالم', fontFamily: 'Noto Sans Arabic',
            fontSize: 16, color: '#000000', bold: false, italic: false,
            rtl: true, textBox: null
          }],
          fallbackTextCount: 0,
          fallbackShapeCount: 0
        }],
        fonts: ['Noto Sans Arabic'],
        editableTextCount: 1,
        editableShapeCount: 0,
        fallbackTextCount: 0,
        fallbackShapeCount: 0,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide).toContain('<a:pPr rtl="1"')
    expect(slide).toContain('<a:t>مرحبا بالعالم</a:t>')
  })

  test('keeps opposite writing directions in separate editable text objects', async () => {
    const textBox = {
      id: 'mixed-direction-line', paragraphId: 'mixed-direction', lineIndex: 0,
      x: 20, width: 200, alignment: 'left' as const, reflow: false
    }
    const result = await createEditablePptx({
      title: 'Mixed direction',
      model: {
        pages: [{
          pageIndex: 0, width: 240, height: 135, fallbackLayers: [],
          elements: [{
            kind: 'text', paintOrder: 0, x: 20, y: 30, width: 50, height: 20,
            baseline: 46, text: 'Evidence', fontFamily: 'Aptos', fontSize: 16,
            color: '#000000', bold: false, italic: false, rtl: false, textBox
          }, {
            kind: 'text', paintOrder: 1, x: 80, y: 30, width: 120, height: 20,
            baseline: 46, text: 'مرحبا بالعالم', fontFamily: 'Noto Sans Arabic',
            fontSize: 16, color: '#000000', bold: false, italic: false,
            rtl: true, textBox
          }],
          fallbackTextCount: 0, fallbackShapeCount: 0
        }],
        fonts: ['Aptos', 'Noto Sans Arabic'], editableTextCount: 2,
        editableShapeCount: 0, fallbackTextCount: 0, fallbackShapeCount: 0,
        warnings: []
      }
    })
    const slide = strFromU8(unzipSync(result.bytes)['ppt/slides/slide1.xml'])

    expect(slide.match(/Editable Typst text/gu)).toHaveLength(2)
    expect(slide.match(/<a:pPr rtl="1"/gu)).toHaveLength(1)
  })
})

import { describe, expect, test } from 'vitest'

import { convertSvgToDrawingMl } from '../src/svg/converter'

describe('portable SVG to DrawingML conversion', () => {
  test('inherits supported paint and opacity from the SVG root', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg" fill="#123456" opacity="0.5">
        <path d="M0 0 L10 0 L10 10 Z"/>
      </svg>
    `)

    expect(result.drawingMl).toContain('123456')
    expect(result.drawingMl).toContain('<a:alpha val="50000"/>')
  })

  test('turns native geometry and glyph uses into editable DrawingML in source order', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg"
           xmlns:xlink="http://www.w3.org/1999/xlink"
           viewBox="0 0 100 60">
        <defs>
          <symbol id="glyph"><path d="M0 0 L6 0 L6 8 L0 8 Z"/></symbol>
          <linearGradient id="shade" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0" stop-color="#ff0000"/>
            <stop offset="1" stop-color="#0000ff" stop-opacity="0.5"/>
          </linearGradient>
        </defs>
        <rect x="2" y="3" width="20" height="10" rx="2" fill="url(#shade)"/>
        <g transform="translate(30 20) scale(2 -2)" opacity="0.75">
          <use xlink:href="#glyph" x="4" y="5" fill="#123456"/>
        </g>
        <circle cx="80" cy="20" r="8" fill="none" stroke="#00aa00" stroke-width="2"/>
      </svg>
    `, { emuPerUnit: 12_700 })

    expect(result.nativeShapeCount).toBe(3)
    expect(result.unsupportedElementCount).toBe(0)
    expect(result.drawingMl).toContain('<a:gradFill')
    expect(result.drawingMl).toContain('<a:alpha val="50000"/>')
    expect(result.drawingMl).toContain('<a:alpha val="75000"/>')
    expect(result.drawingMl).toContain('<a:prstGeom prst="ellipse"')
    expect(result.drawingMl).toContain('123456')
    expect(result.drawingMl.indexOf('SVG rectangle')).toBeLessThan(
      result.drawingMl.indexOf('SVG path')
    )
    expect(result.residualHasVisualContent).toBe(false)
  })

  test('keeps unsupported clipped or filtered content in a residual SVG fallback', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
        <defs><clipPath id="clip"><rect width="10" height="10"/></clipPath></defs>
        <g clip-path="url(#clip)"><path d="M0 0 L20 0 L20 20 Z"/></g>
        <path d="M40 0 L50 0 L50 10 Z" filter="url(#shadow)"/>
      </svg>
    `)

    expect(result.nativeShapeCount).toBe(0)
    expect(result.unsupportedElementCount).toBe(2)
    expect(result.residualHasVisualContent).toBe(true)
    expect(result.residualSvg).toContain('clip-path="url(#clip)"')
    expect(result.residualSvg).toContain('filter="url(#shadow)"')
  })

  test('preserves alternating residual and native SVG paint layers', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
        <path d="M0 0 L20 0 L20 20 Z" fill="#aa0000" filter="url(#shadow)"/>
        <path d="M30 0 L50 0 L50 20 Z" fill="#00aa00"/>
        <path d="M60 0 L80 0 L80 20 Z" fill="#0000aa" clip-path="url(#clip)"/>
      </svg>
    `)

    expect(result.layers.map((layer) => layer.kind)).toEqual([
      'residualSvg', 'drawingMl', 'residualSvg'
    ])
    const residuals = result.layers.filter((layer) => layer.kind === 'residualSvg')
    expect(residuals[0]?.svg).toContain('#aa0000')
    expect(residuals[0]?.svg).not.toContain('#0000aa')
    expect(residuals[1]?.svg).toContain('#0000aa')
    expect(residuals[1]?.svg).not.toContain('#aa0000')
    expect(result.layers[1]).toMatchObject({ kind: 'drawingMl', nativeShapeCount: 1 })
  })

  test('groups converted glyph outlines using compiler-provided formula bounds', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 60">
        <defs>
          <symbol id="glyph"><path d="M0 0 L6 0 L6 8 L0 8 Z"/></symbol>
        </defs>
        <g transform="translate(10 20)"><use href="#glyph"/></g>
        <g transform="translate(20 20)"><use href="#glyph"/></g>
        <rect x="70" y="10" width="10" height="10"/>
      </svg>
    `, {
      vectorGroups: [{
        id: 'formula-1', kind: 'math', x: 8, y: 18, width: 30, height: 15
      }]
    })

    expect(result.nativeShapeCount).toBe(3)
    expect(result.drawingMl).toContain('<p:grpSp>')
    expect(result.drawingMl).toContain('Typst math formula-1')
    expect(result.drawingMl.indexOf('<p:grpSp>')).toBeLessThan(
      result.drawingMl.indexOf('SVG rectangle')
    )
    expect(result.residualHasVisualContent).toBe(false)
  })

  test('keeps a referenced symbol intact when only part of it is convertible', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg">
        <defs><symbol id="mixed">
          <path d="M0 0 L10 0 L10 10 Z"/>
          <path d="M20 0 L30 0 L30 10 Z" filter="url(#shadow)"/>
        </symbol></defs>
        <use href="#mixed" x="5" y="5"/>
      </svg>
    `)

    expect(result.nativeShapeCount).toBe(0)
    expect(result.residualHasVisualContent).toBe(true)
    expect(result.residualSvg).toContain('<use href="#mixed"')
  })

  test('keeps unsupported paint residual and converts dashed strokes explicitly', () => {
    const result = convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg">
        <path d="M0 0 L10 0 L10 10 Z" fill="url(#missing)"/>
        <path d="M20 0 L30 0" fill="none" stroke="#000" stroke-dasharray="2 2"/>
      </svg>
    `)

    expect(result.nativeShapeCount).toBe(1)
    expect(result.unsupportedElementCount).toBe(1)
    expect(result.residualHasVisualContent).toBe(true)
    expect(result.residualSvg).toContain('url(#missing)')
    expect(result.residualSvg).not.toContain('stroke-dasharray="2 2"')
    expect(result.drawingMl).toContain('<a:prstDash val="sysDot"/>')
  })

  test('rejects XML with parser errors or external use references', () => {
    expect(() => convertSvgToDrawingMl('<svg><path></svg>')).toThrow('invalid SVG')
    expect(() => convertSvgToDrawingMl(`
      <svg xmlns="http://www.w3.org/2000/svg">
        <use href="https://example.com/art.svg#shape"/>
      </svg>
    `)).toThrow('external SVG references')
  })
})

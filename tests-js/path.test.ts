import { describe, expect, test } from 'vitest'

import {
  normalizeSvgPath,
  parseSvgPath,
  pathToDrawingMlGeometry
} from '../src/svg/path'

describe('portable SVG path conversion', () => {
  test('parses repeated coordinates and converts every relative command to absolute', () => {
    expect(normalizeSvgPath(parseSvgPath(
      'm 10 20 5 0 h 5 v 5 q 5 10 10 0 t 10 0 s 5 -5 10 0 z'
    ))).toEqual([
      { command: 'M', values: [10, 20] },
      { command: 'L', values: [15, 20] },
      { command: 'L', values: [20, 20] },
      { command: 'L', values: [20, 25] },
      { command: 'C', values: [23.333333333333332, 31.666666666666664,
        26.666666666666668, 31.666666666666664, 30, 25] },
      { command: 'C', values: [33.333333333333336, 18.333333333333336,
        36.666666666666664, 18.333333333333336, 40, 25] },
      { command: 'C', values: [40, 25, 45, 20, 50, 25] },
      { command: 'Z', values: [] }
    ])
  })

  test('normalizes elliptical arcs to cubic DrawingML paths', () => {
    const geometry = pathToDrawingMlGeometry('M 10 20 A 30 15 30 0 1 80 50 Z', {
      emuPerUnit: 12_700
    })

    expect(geometry.bounds.width).toBeGreaterThan(65)
    expect(geometry.bounds.height).toBeGreaterThan(25)
    expect(geometry.xml).toContain('<a:cubicBezTo>')
    expect(geometry.xml).toContain('<a:close/>')
    expect(geometry.xml).not.toContain('NaN')
  })

  test('applies an affine matrix before calculating exact geometry bounds', () => {
    const geometry = pathToDrawingMlGeometry('M 0 0 L 10 0 L 10 20 Z', {
      emuPerUnit: 12_700,
      transform: [0, 2, -3, 0, 100, 50]
    })

    expect(geometry.bounds).toEqual({ x: 40, y: 50, width: 60, height: 20 })
    expect(geometry.xml).toContain('w="762000" h="254000"')
  })
})

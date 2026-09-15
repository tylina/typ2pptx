import { describe, expect, test } from 'vitest'

import { rasterDimensions, rasterizeSvgToPngBase64 } from '../src/svg/rasterizer'

describe('portable SVG fallback rasterizer', () => {
  test('encodes a transparent page-sized PNG in Node without Python', async () => {
    const encoded = await rasterizeSvgToPngBase64({
      svg: '<svg xmlns="http://www.w3.org/2000/svg"><rect width="50" height="30" fill="red"/></svg>',
      width: 100,
      height: 60
    })
    const png = Buffer.from(encoded, 'base64')

    expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10])
    expect(png.readUInt32BE(16)).toBe(133)
    expect(png.readUInt32BE(20)).toBe(80)
  })

  test('caps unusually large pages and rejects network-backed images', async () => {
    const dimensions = rasterDimensions(20_000, 10_000)
    expect(dimensions.pixelWidth * dimensions.pixelHeight).toBeLessThanOrEqual(4_010_000)
    await expect(rasterizeSvgToPngBase64({
      svg: '<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/a.png"/></svg>',
      width: 100,
      height: 60
    })).rejects.toThrow('external image reference')
  })
})

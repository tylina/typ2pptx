import {
  initWasm,
  Resvg,
  type InitInput
} from '@resvg/resvg-wasm'

import type { ResidualSvgRasterizer } from '../contracts'
import { rasterizeSvgWithRenderer } from './rasterizer'

let initialization: Promise<void> | undefined

/** Create a deterministic Rust/WASM rasterizer for browser and Worker hosts. */
export function createResvgWasmRasterizer(wasm: InitInput): ResidualSvgRasterizer {
  return (request) => rasterizeSvgWithRenderer(
    request,
    async (svg, width, height) => {
      initialization ??= initWasm(wasm)
      await initialization
      const renderer = new Resvg(svg, {
        fitTo: { mode: 'original' },
        imageRendering: 0,
        shapeRendering: 2,
        textRendering: 2
      })
      try {
        const unresolved = renderer.imagesToResolve()
        if (unresolved.length > 0) {
          throw new Error('SVG fallback contains an unresolved image resource')
        }
        if (renderer.width !== width || renderer.height !== height) {
          throw new Error('SVG fallback rasterizer returned unexpected dimensions')
        }
        const rendered = renderer.render()
        try {
          return rendered.asPng().slice()
        } finally {
          rendered.free()
        }
      } finally {
        renderer.free()
      }
    }
  )
}

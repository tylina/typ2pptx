import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement
} from '@xmldom/xmldom'

import type { ResidualSvgRasterizationRequest } from '../contracts'

const PIXELS_PER_POINT = 96 / 72
const MAX_RASTER_PIXELS = 4_000_000
const MAX_SVG_BYTES = 16 * 1024 * 1024
const MAX_PNG_BYTES = 32 * 1024 * 1024

interface NodeCanvasModule {
  createCanvas(width: number, height: number): {
    getContext(type: '2d'): { drawImage(image: unknown, x: number, y: number, w: number, h: number): void }
    toBuffer(mimeType: 'image/png'): Uint8Array
  }
  loadImage(source: Uint8Array): Promise<unknown>
}

/** Rasterize a self-contained SVG without Python in a browser Worker, browser page, or Node host. */
export async function rasterizeSvgToPngBase64(
  request: ResidualSvgRasterizationRequest
): Promise<string> {
  return rasterizeSvgWithRenderer(request, async (svg, width, height) =>
    typeof OffscreenCanvas !== 'undefined' && typeof createImageBitmap === 'function'
      ? rasterizeWithOffscreenCanvas(svg, width, height)
      : typeof document !== 'undefined'
        ? rasterizeWithHtmlCanvas(svg, width, height)
        : rasterizeWithNodeCanvas(svg, width, height))
}

export async function rasterizeSvgWithRenderer(
  request: ResidualSvgRasterizationRequest,
  renderer: (svg: string, width: number, height: number) => Promise<Uint8Array>
): Promise<string> {
  const { pixelWidth, pixelHeight } = rasterDimensions(request.width, request.height)
  const svg = normalizedSvg(request.svg, request.width, request.height, pixelWidth, pixelHeight)
  const bytes = await renderer(svg, pixelWidth, pixelHeight)
  assertPng(bytes)
  if (bytes.byteLength > MAX_PNG_BYTES) {
    throw new Error('Rasterized SVG exceeds the 32 MB PNG budget')
  }
  return bytesToBase64(bytes)
}

export function rasterDimensions(width: number, height: number): {
  pixelWidth: number
  pixelHeight: number
} {
  if (![width, height].every((value) => Number.isFinite(value) && value > 0)) {
    throw new Error('SVG rasterization requires positive page dimensions')
  }
  const requestedPixels = width * PIXELS_PER_POINT * height * PIXELS_PER_POINT
  const scale = requestedPixels > MAX_RASTER_PIXELS
    ? PIXELS_PER_POINT * Math.sqrt(MAX_RASTER_PIXELS / requestedPixels)
    : PIXELS_PER_POINT
  return {
    pixelWidth: Math.max(1, Math.round(width * scale)),
    pixelHeight: Math.max(1, Math.round(height * scale))
  }
}

async function rasterizeWithOffscreenCanvas(
  svg: string,
  width: number,
  height: number
): Promise<Uint8Array<ArrayBuffer>> {
  const image = await createImageBitmap(new Blob([svg], { type: 'image/svg+xml' }))
  try {
    const canvas = new OffscreenCanvas(width, height)
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Unable to create a 2D canvas for the SVG fallback')
    context.drawImage(image, 0, 0, width, height)
    const blob = await canvas.convertToBlob({ type: 'image/png' })
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    image.close()
  }
}

async function rasterizeWithHtmlCanvas(
  svg: string,
  width: number,
  height: number
): Promise<Uint8Array<ArrayBuffer>> {
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }))
  try {
    const image = await loadHtmlImage(url)
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Unable to create a 2D canvas for the SVG fallback')
    context.drawImage(image, 0, 0, width, height)
    const blob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((value) => {
        if (value) resolve(value)
        else reject(new Error('Browser could not encode the SVG fallback as PNG'))
      }, 'image/png')
    })
    canvas.width = 1
    canvas.height = 1
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    URL.revokeObjectURL(url)
  }
}

function loadHtmlImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error('Browser could not decode the SVG fallback'))
    image.src = url
  })
}

async function rasterizeWithNodeCanvas(
  svg: string,
  width: number,
  height: number
): Promise<Uint8Array<ArrayBuffer>> {
  if (typeof process === 'undefined' || !process.versions?.node) {
    throw new Error('This runtime cannot rasterize an SVG fallback')
  }
  const moduleName = '@napi-rs/canvas'
  let canvasModule: NodeCanvasModule
  try {
    canvasModule = await import(/* @vite-ignore */ moduleName) as unknown as NodeCanvasModule
  } catch {
    throw new Error('Node SVG fallback rasterization requires the optional @napi-rs/canvas package')
  }
  const image = await canvasModule.loadImage(new TextEncoder().encode(svg))
  const canvas = canvasModule.createCanvas(width, height)
  const context = canvas.getContext('2d')
  context.drawImage(image, 0, 0, width, height)
  const encoded = canvas.toBuffer('image/png')
  const bytes = new Uint8Array(encoded.byteLength)
  bytes.set(encoded)
  return bytes
}

function normalizedSvg(
  source: string,
  width: number,
  height: number,
  pixelWidth: number,
  pixelHeight: number
): string {
  if (new TextEncoder().encode(source).byteLength > MAX_SVG_BYTES) {
    throw new Error('SVG exceeds the 16 MB rasterization budget')
  }
  if (/<!DOCTYPE|<!ENTITY/iu.test(source)) throw new Error('SVG declarations are not allowed')
  const document = parseSvg(source)
  const root = document.documentElement
  if (!root || localName(root) !== 'svg') throw new Error('Cannot rasterize an invalid SVG root')
  assertSelfContained(document)
  if (!root.getAttribute('viewBox')) root.setAttribute('viewBox', `0 0 ${width} ${height}`)
  root.setAttribute('width', String(pixelWidth))
  root.setAttribute('height', String(pixelHeight))
  return new XMLSerializer().serializeToString(document)
}

function parseSvg(source: string): XmlDocument {
  const errors: string[] = []
  const document = new DOMParser({
    onError: (level, message) => {
      if (level !== 'warning') errors.push(message)
    }
  }).parseFromString(source, 'image/svg+xml')
  if (errors.length > 0 || document.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Cannot rasterize invalid SVG XML')
  }
  return document
}

function assertSelfContained(document: XmlDocument): void {
  const elements = document.getElementsByTagName('*')
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements.item(index) as XmlElement | null
    if (!element) continue
    const tag = localName(element)
    if (tag === 'script' || tag === 'foreignObject') {
      throw new Error(`SVG fallback contains unsupported <${tag}> content`)
    }
    if (tag !== 'image' && tag !== 'use') continue
    const href = element.getAttribute('href') || element.getAttribute('xlink:href')
    if (href && !href.startsWith('#') && !href.startsWith('data:')) {
      throw new Error('SVG fallback contains an external image reference')
    }
  }
}

function localName(element: XmlElement): string {
  return element.localName || element.tagName.replace(/^.*:/u, '')
}

function assertPng(bytes: Uint8Array): void {
  const signature = [137, 80, 78, 71, 13, 10, 26, 10]
  if (signature.some((value, index) => bytes[index] !== value)) {
    throw new Error('SVG rasterizer returned an invalid PNG image')
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64')
  let binary = ''
  const chunkSize = 32_768
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
  }
  return btoa(binary)
}

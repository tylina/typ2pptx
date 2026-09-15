import PptxGenJS from 'pptxgenjs'

import type {
  EditablePresentationRequest,
  PresentationArtifact,
  PresentationExportModel,
  PresentationImageElement,
  PresentationLinkElement,
  PresentationPageModel,
  PresentationRectangleElement,
  PresentationSlideLinkElement,
  PresentationTextElement,
  VisualPresentationRequest
} from './contracts'
import { patchPresentationArchive, type SlideArchivePatch } from './ooxml/archive'
import { convertSvgToDrawingMl } from './svg/converter'
import { rasterizeSvgToPngBase64 } from './svg/rasterizer'

const POINTS_PER_INCH = 72
const MAX_PAGES = 1_000
const MAX_ELEMENTS = 100_000
const MAX_OUTPUT_BYTES = 64 * 1024 * 1024
const MAX_INTERMEDIATE_BYTES = 128 * 1024 * 1024

export async function createVisualPptx(
  request: VisualPresentationRequest
): Promise<PresentationArtifact> {
  const pages = normalizedPages(request.pages)
  const pptx = createPresentation(request.title, pages[0].width, pages[0].height)
  for (const page of pages) {
    assertPng(page.pngBase64)
    const slide = pptx.addSlide()
    slide.addImage({
      data: `data:image/png;base64,${page.pngBase64}`,
      x: 0,
      y: 0,
      w: toInches(page.width),
      h: toInches(page.height),
      altText: `Rendered document page ${page.pageIndex + 1}`
    })
    addNotes(slide, request.notesByPageIndex?.get(page.pageIndex))
  }
  const bytes = await writePresentation(pptx)
  return emptyArtifact(bytes, pages.length)
}

export async function createEditablePptx(
  request: EditablePresentationRequest
): Promise<PresentationArtifact> {
  const pages = normalizedPages(request.model.pages)
  assertModelBudget(request.model)
  const pptx = createPresentation(request.title, pages[0].width, pages[0].height)
  const slideNumbers = new Map(pages.map((page, index) => [page.pageIndex, index + 1]))
  const patches: SlideArchivePatch[] = []
  const warnings = new Set(request.model.warnings)
  let nativeVectorShapeCount = 0
  let remainingFallbackShapeCount = 0

  for (const [pageOffset, page] of pages.entries()) {
    const converted = convertSvgToDrawingMl(page.nonTextSvg, {
      emuPerUnit: 12_700,
      firstShapeId: 1_000_000 + pageOffset * 100_000,
      vectorGroups: page.vectorGroups
    })
    nativeVectorShapeCount += converted.nativeShapeCount
    // Compiler fallback items and SVG leaves are different units. Only claim that all
    // shape fallbacks disappeared when the residual layer is actually empty.
    if (converted.residualHasVisualContent) {
      remainingFallbackShapeCount += page.fallbackShapeCount
    }
    for (const warning of converted.warnings) warnings.add(warning)

    const slide = pptx.addSlide()
    const vectorLayers: SlideArchivePatch['vectorLayers'] = []
    const residualLayers: SlideArchivePatch['residualLayers'] = []
    for (const [layerOffset, layer] of converted.layers.entries()) {
      const layerNumber = layerOffset + 1
      if (layer.kind === 'drawingMl') {
        const anchorName = `typ2pptx-vector-anchor-${page.pageIndex + 1}-${layerNumber}`
        addVectorAnchor(pptx, slide, anchorName)
        vectorLayers.push({ anchorName, drawingMl: layer.drawingMl })
        continue
      }
      const fallbackPngBase64 = converted.nativeShapeCount > 0
        ? await rasterizeResidualSvg(request, layer.svg, page.width, page.height)
        : page.nonTextPngBase64
      assertPng(fallbackPngBase64)
      const imageName = `typ2pptx residual layer ${page.pageIndex + 1}.${layerNumber}`
      slide.addImage({
        data: `data:image/png;base64,${fallbackPngBase64}`,
        x: 0,
        y: 0,
        w: toInches(page.width),
        h: toInches(page.height),
        altText: `Unconverted Typst artwork for page ${page.pageIndex + 1}`,
        objectName: imageName
      })
      residualLayers.push({
        imageName,
        mediaName: `typ2pptx-residual-${pageOffset + 1}-${layerNumber}`,
        svg: layer.svg
      })
    }
    addEditableElements(pptx, slide, page, slideNumbers)
    addNotes(slide, request.notesByPageIndex?.get(page.pageIndex))
    patches.push({
      slideNumber: pageOffset + 1,
      vectorLayers,
      residualLayers
    })
  }

  const omittedSlideLinkCount = countOmittedSlideLinks(pages, slideNumbers)
  if (omittedSlideLinkCount > 0) warnings.add(omittedSlideLinkWarning(omittedSlideLinkCount))
  const source = await writePresentation(pptx, false, MAX_INTERMEDIATE_BYTES)
  const bytes = patchPresentationArchive(source, patches)
  assertOutputSize(bytes)
  return {
    bytes,
    pageCount: pages.length,
    editableTextCount: request.model.editableTextCount,
    editableShapeCount: request.model.editableShapeCount + nativeVectorShapeCount,
    editableLinkCount: countEditablePresentationLinks(pages),
    fallbackTextCount: request.model.fallbackTextCount,
    fallbackShapeCount: remainingFallbackShapeCount,
    nativeVectorShapeCount,
    warnings: [...warnings]
  }
}

async function rasterizeResidualSvg(
  request: EditablePresentationRequest,
  svg: string,
  width: number,
  height: number
): Promise<string> {
  const rasterizer = request.residualSvgRasterizer ?? rasterizeSvgToPngBase64
  return rasterizer({ svg, width, height })
}

function createPresentation(title: string, width: number, height: number): PptxGenJS {
  const pptx = new PptxGenJS()
  pptx.defineLayout({
    name: 'TYP2PPTX_DOCUMENT',
    width: toInches(width),
    height: toInches(height)
  })
  pptx.layout = 'TYP2PPTX_DOCUMENT'
  pptx.author = 'typ2pptx'
  pptx.company = 'typ2pptx'
  pptx.subject = 'Editable presentation exported from a Typst document'
  pptx.title = title
  return pptx
}

function addEditableElements(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  page: PresentationPageModel,
  slideNumbers: ReadonlyMap<number, number>
): void {
  for (let index = 0; index < page.elements.length; index += 1) {
    const element = page.elements[index]!
    if (element.kind === 'text') {
      const runs = [element]
      while (
        index + 1 < page.elements.length &&
        page.elements[index + 1]?.kind === 'text' &&
        canShareTextBox(runs, page.elements[index + 1] as PresentationTextElement)
      ) {
        runs.push(page.elements[index + 1] as PresentationTextElement)
        index += 1
      }
      addEditableText(slide, runs)
    }
    if (element.kind === 'rectangle') addEditableRectangle(pptx, slide, element)
    if (element.kind === 'image') addEditableImage(slide, element)
  }
  for (const element of page.elements) {
    if (element.kind === 'link') addLinkOverlay(pptx, slide, element)
    if (element.kind === 'slideLink') {
      addSlideLinkOverlay(pptx, slide, element, slideNumbers)
    }
  }
}

function addEditableImage(slide: PptxGenJS.Slide, element: PresentationImageElement): void {
  validateElementBounds(element)
  assertEmbeddedImage(element.mediaType, element.dataBase64)
  slide.addImage({
    data: `data:${element.mediaType};base64,${element.dataBase64}`,
    x: toInches(element.x),
    y: toInches(element.y),
    w: toInches(element.width),
    h: toInches(element.height),
    altText: element.altText || 'Typst image',
    objectName: 'Editable Typst image'
  })
}

function addEditableText(
  slide: PptxGenJS.Slide,
  elements: readonly PresentationTextElement[]
): void {
  for (const element of elements) validateElementBounds(element)
  const visible = elements.filter((element) => element.text)
  if (visible.length === 0) return
  const left = Math.min(...visible.map((element) => element.x))
  const top = Math.min(...visible.map((element) => element.y))
  const right = Math.max(...visible.map((element) => element.x + element.width))
  const bottom = Math.max(...visible.map((element) => element.y + element.height))
  const line = textLineMetrics(visible)
  const runs: PptxGenJS.TextProps[] = visible.map((element, index) => ({
    text: textWithLayoutGap(visible, index),
    options: {
      fontFace: element.fontFamily || 'Arial',
      fontSize: runFontSize(element, line),
      color: cleanHexColor(element.color),
      bold: element.bold,
      italic: element.italic,
      baseline: runBaseline(element, line),
      breakLine: false
    }
  }))
  slide.addText(runs, {
    x: toInches(left),
    y: toInches(top),
    w: toInches(right - left),
    h: toInches(bottom - top),
    margin: 0,
    fit: 'resize',
    valign: 'top',
    isTextBox: true,
    objectName: 'Editable Typst text'
  })
}

function textWithLayoutGap(
  elements: readonly PresentationTextElement[],
  index: number
): string {
  const element = elements[index]!
  if (index === 0) return element.text
  const previous = elements[index - 1]!
  const gap = element.x - (previous.x + previous.width)
  const averageFontSize = (element.fontSize + previous.fontSize) / 2
  if (
    gap > averageFontSize * 0.15 &&
    !previous.text.endsWith(' ') &&
    !element.text.startsWith(' ')
  ) return ` ${element.text}`
  return element.text
}

function canShareTextBox(
  current: readonly PresentationTextElement[],
  right: PresentationTextElement
): boolean {
  const proposed = [...current, right]
  const line = textLineMetrics(proposed)
  if (!proposed.every((element) => belongsToTextLine(element, line))) return false

  const currentRight = Math.max(...current.map((element) => element.x + element.width))
  const horizontalGap = right.x - currentRight
  const overlapTolerance = Math.max(0.5, line.fontSize * 0.08)
  const includesScript = proposed.some((element) => isScriptRun(element, line))
  if (horizontalGap < -overlapTolerance) {
    if (!includesScript || horizontalGap < -line.fontSize * 1.5) return false
  }
  return horizontalGap <= line.fontSize * 2
}

interface TextLineMetrics {
  baseline: number
  fontSize: number
}

function textLineMetrics(elements: readonly PresentationTextElement[]): TextLineMetrics {
  const fontSize = Math.max(...elements.map((element) =>
    positiveNumber(element.fontSize, 'font size')))
  const anchors = elements.filter((element) => element.fontSize >= fontSize * 0.9)
  const baseline = anchors.reduce((sum, element) => sum + elementBaseline(element), 0) /
    anchors.length
  return { baseline, fontSize }
}

function belongsToTextLine(
  element: PresentationTextElement,
  line: TextLineMetrics
): boolean {
  const difference = Math.abs(elementBaseline(element) - line.baseline)
  if (difference <= Math.max(0.75, line.fontSize * 0.08)) return true
  return isScriptRun(element, line) && difference < line.fontSize * 0.8
}

function isScriptRun(
  element: PresentationTextElement,
  line: TextLineMetrics
): boolean {
  return element.fontSize < line.fontSize * 0.85
}

function runFontSize(
  element: PresentationTextElement,
  line: TextLineMetrics
): number {
  return isScriptRun(element, line)
    ? line.fontSize
    : positiveNumber(element.fontSize, 'font size')
}

function runBaseline(
  element: PresentationTextElement,
  line: TextLineMetrics
): number | undefined {
  if (!isScriptRun(element, line)) return undefined
  const difference = line.baseline - elementBaseline(element)
  if (Math.abs(difference) <= line.fontSize * 0.08) return undefined
  const percentage = Math.max(20, Math.min(50, Math.round(
    Math.abs(difference) / line.fontSize * 100
  )))
  // PptxGenJS 4 serializes its baseline option in fiftieths of a percent.
  // DrawingML stores thousandths of a percent, so multiply the desired value by 20.
  return Math.sign(difference) * percentage * 20
}

function elementBaseline(element: PresentationTextElement): number {
  return element.baseline ?? element.y + element.height * 0.8
}

function addEditableRectangle(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: PresentationRectangleElement
): void {
  validateElementBounds(element)
  const color = cleanHexColor(element.fillColor)
  slide.addShape(pptx.ShapeType.rect, {
    x: toInches(element.x),
    y: toInches(element.y),
    w: toInches(element.width),
    h: toInches(element.height),
    fill: { color },
    line: { color, transparency: 100, width: 0.01 },
    objectName: 'Editable Typst rectangle'
  })
}

function addLinkOverlay(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: PresentationLinkElement
): void {
  validateElementBounds(element)
  const url = supportedPresentationExternalUrl(element.url)
  if (!url) return
  slide.addShape(pptx.ShapeType.rect, {
    x: toInches(element.x),
    y: toInches(element.y),
    w: toInches(element.width),
    h: toInches(element.height),
    fill: { color: 'FFFFFF', transparency: 100 },
    line: { color: 'FFFFFF', transparency: 100 },
    hyperlink: { url },
    objectName: 'Typst hyperlink'
  })
}

function addSlideLinkOverlay(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: PresentationSlideLinkElement,
  slideNumbers: ReadonlyMap<number, number>
): void {
  validateElementBounds(element)
  const target = slideNumbers.get(element.pageIndex)
  if (target === undefined) return
  slide.addShape(pptx.ShapeType.rect, {
    x: toInches(element.x),
    y: toInches(element.y),
    w: toInches(element.width),
    h: toInches(element.height),
    fill: { color: 'FFFFFF', transparency: 100 },
    line: { color: 'FFFFFF', transparency: 100 },
    hyperlink: { slide: target },
    objectName: 'Typst slide hyperlink'
  })
}

function addVectorAnchor(pptx: PptxGenJS, slide: PptxGenJS.Slide, name: string): void {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 0.01,
    h: 0.01,
    fill: { color: 'FFFFFF', transparency: 100 },
    line: { color: 'FFFFFF', transparency: 100 },
    objectName: name
  })
}

function normalizedPages<T extends { pageIndex: number; width: number; height: number }>(
  source: readonly T[]
): T[] {
  if (source.length === 0) throw new Error('PPTX export requires at least one page')
  if (source.length > MAX_PAGES) throw new Error(`PPTX export exceeds the ${MAX_PAGES}-page budget`)
  const pages = [...source].sort((left, right) => left.pageIndex - right.pageIndex)
  const first = pages[0]
  const indexes = new Set<number>()
  for (const page of pages) {
    if (
      !Number.isInteger(page.pageIndex) ||
      page.pageIndex < 0 ||
      !Number.isFinite(page.width) ||
      !Number.isFinite(page.height) ||
      page.width <= 0 ||
      page.height <= 0
    ) throw new Error('PPTX export received invalid page geometry')
    if (indexes.has(page.pageIndex)) throw new Error('PPTX export received duplicate page indexes')
    indexes.add(page.pageIndex)
    if (!approximatelyEqual(page.width, first.width) || !approximatelyEqual(page.height, first.height)) {
      throw new Error('PPTX export requires every slide to have the same page size')
    }
  }
  return pages
}

function assertModelBudget(model: PresentationExportModel): void {
  const count = model.pages.reduce((sum, page) => sum + page.elements.length, 0)
  if (count > MAX_ELEMENTS) {
    throw new Error(`Editable PPTX exceeds the ${MAX_ELEMENTS}-element budget`)
  }
  const imageBytes = model.pages.reduce((total, page) => total + page.elements.reduce(
    (pageTotal, element) => pageTotal + (element.kind === 'image'
      ? element.dataBase64.length
      : 0),
    0
  ), 0)
  if (imageBytes > MAX_OUTPUT_BYTES) {
    throw new Error('Editable PPTX images exceed the 64 MB input budget')
  }
}

export function countEditablePresentationLinks(pages: readonly PresentationPageModel[]): number {
  const pageIndexes = new Set(pages.map((page) => page.pageIndex))
  return pages.reduce((count, page) => count + page.elements.filter((element) =>
    element.kind === 'link'
      ? supportedPresentationExternalUrl(element.url) !== null
      : element.kind === 'slideLink' && pageIndexes.has(element.pageIndex)
  ).length, 0)
}

function countOmittedSlideLinks(
  pages: readonly PresentationPageModel[],
  slideNumbers: ReadonlyMap<number, number>
): number {
  return pages.reduce((count, page) => count + page.elements.filter((element) =>
    element.kind === 'slideLink' && !slideNumbers.has(element.pageIndex)
  ).length, 0)
}

export function supportedPresentationExternalUrl(value: string): string | null {
  try {
    const url = new URL(value)
    return ['http:', 'https:', 'mailto:'].includes(url.protocol) ? url.toString() : null
  } catch {
    return null
  }
}

export function omittedSlideLinkWarning(count: number): string {
  return `Omitted ${count} internal link(s) whose target pages are not included in this PPTX.`
}

function validateElementBounds(element: {
  x: number
  y: number
  width: number
  height: number
}): void {
  for (const [label, value] of Object.entries({
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height
  })) {
    if (!Number.isFinite(value)) throw new Error(`PPTX element has an invalid ${label}`)
  }
  if (element.width <= 0 || element.height <= 0) {
    throw new Error('PPTX element width and height must be positive')
  }
}

function addNotes(slide: PptxGenJS.Slide, notes: string | undefined): void {
  const normalized = notes?.trim()
  if (normalized) slide.addNotes(normalized)
}

async function writePresentation(
  pptx: PptxGenJS,
  compression = true,
  maxBytes = MAX_OUTPUT_BYTES
): Promise<Uint8Array<ArrayBuffer>> {
  const output = await pptx.write({ outputType: 'uint8array', compression })
  const bytes = output instanceof Uint8Array
    ? exactArrayBufferView(output)
    : output instanceof ArrayBuffer
      ? new Uint8Array(output)
      : output instanceof Blob
        ? new Uint8Array(await output.arrayBuffer())
        : null
  if (!bytes) throw new Error('PPTX writer returned an unsupported output type')
  if (bytes.byteLength > maxBytes) throw new Error('PPTX working data exceeds its memory budget')
  return bytes
}

function emptyArtifact(
  bytes: Uint8Array<ArrayBuffer>,
  pageCount: number
): PresentationArtifact {
  return {
    bytes,
    pageCount,
    editableTextCount: 0,
    editableShapeCount: 0,
    editableLinkCount: 0,
    fallbackTextCount: 0,
    fallbackShapeCount: 0,
    nativeVectorShapeCount: 0,
    warnings: []
  }
}

function exactArrayBufferView(value: Uint8Array): Uint8Array<ArrayBuffer> {
  if (
    value.buffer instanceof ArrayBuffer &&
    value.byteOffset === 0 &&
    value.byteLength === value.buffer.byteLength
  ) return value as Uint8Array<ArrayBuffer>
  return value.slice()
}

function assertPng(value: string): void {
  if (!value.startsWith('iVBOR')) throw new Error('PPTX background fallback is not a PNG image')
}

function assertEmbeddedImage(
  mediaType: PresentationImageElement['mediaType'],
  value: string
): void {
  const signature = {
    'image/png': 'iVBOR',
    'image/jpeg': '/9j/',
    'image/gif': 'R0lGOD'
  }[mediaType]
  if (!signature || !value.startsWith(signature)) {
    throw new Error(`PPTX embedded ${mediaType || 'image'} has an invalid signature`)
  }
}

function assertOutputSize(bytes: Uint8Array): void {
  if (bytes.byteLength > MAX_OUTPUT_BYTES) throw new Error('PPTX output exceeds the 64 MB budget')
}

function cleanHexColor(value: string): string {
  const color = value.startsWith('#') ? value.slice(1) : value
  return /^[0-9a-f]{6}$/iu.test(color) ? color.toLocaleUpperCase() : '000000'
}

function positiveNumber(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) throw new Error(`PPTX text has an invalid ${label}`)
  return value
}

function toInches(points: number): number {
  return points / POINTS_PER_INCH
}

function approximatelyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= Math.max(0.01, Math.abs(right) * 0.00001)
}

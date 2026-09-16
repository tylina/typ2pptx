import PptxGenJS from 'pptxgenjs'

import type {
  EditablePresentationRequest,
  PresentationArtifact,
  PresentationExportModel,
  PresentationFallbackLayer,
  PresentationImageElement,
  PresentationLinkElement,
  PresentationPageBackground,
  PresentationPageModel,
  PresentationRectangleElement,
  PresentationSlideLinkElement,
  PresentationTextElement,
  PresentationTextBox,
  VisualPresentationRequest
} from './contracts'
import { patchPresentationArchive, type SlideArchivePatch } from './ooxml/archive'
import { convertSvgToDrawingMl } from './svg/converter'
import { rasterizeSvgToPngBase64 } from './svg/rasterizer'

const POINTS_PER_INCH = 72
const MAX_PAGES = 1_000
const MAX_ELEMENTS = 100_000
const MAX_FALLBACK_LAYERS = 512
const MAX_VECTOR_INPUT_BYTES = 64 * 1024 * 1024
const MAX_OUTPUT_BYTES = 64 * 1024 * 1024
const MAX_INTERMEDIATE_BYTES = 128 * 1024 * 1024
// Desktop PowerPoint places the baseline 0.95 em below a zero-inset text body's
// top edge. Keep compiler-positioned text anchored to its Typst baseline.
const POWERPOINT_BASELINE_FROM_TOP_EM = 0.95

type PresentationPaintLayer =
  | PresentationFallbackLayer
  | PresentationTextElement
  | PresentationRectangleElement
  | PresentationImageElement

interface GeneratedTextRunBudget {
  remaining: number
}

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
  const generatedTextRunBudget = { remaining: MAX_ELEMENTS }
  let nativeVectorShapeCount = 0
  let remainingFallbackShapeCount = 0

  for (const [pageOffset, page] of pages.entries()) {
    const slide = pptx.addSlide()
    await addPageBackground(request, slide, page)
    const vectorLayers: SlideArchivePatch['vectorLayers'] = []
    const residualLayers: SlideArchivePatch['residualLayers'] = []
    const paintLayers = orderedPaintLayers(page)
    let pageHasResidual = false
    let archiveLayerNumber = 0
    let fallbackLayerNumber = 0
    for (let paintOffset = 0; paintOffset < paintLayers.length; paintOffset += 1) {
      const paint = paintLayers[paintOffset]!
      if (paint.kind === 'text') {
        const runs = [paint]
        while (
          paintOffset + 1 < paintLayers.length &&
          paintLayers[paintOffset + 1]?.kind === 'text' &&
          canShareTextContainer(runs, paintLayers[paintOffset + 1] as PresentationTextElement)
        ) {
          runs.push(paintLayers[paintOffset + 1] as PresentationTextElement)
          paintOffset += 1
        }
        addEditableText(slide, runs, page.width, generatedTextRunBudget)
        continue
      }
      if (paint.kind === 'rectangle') {
        addEditableRectangle(pptx, slide, paint)
        continue
      }
      if (paint.kind === 'image') {
        addEditableImage(slide, paint)
        continue
      }

      fallbackLayerNumber += 1
      const converted = convertSvgToDrawingMl(paint.svg, {
        emuPerUnit: 12_700,
        firstShapeId: 1_000_000 + fallbackLayerNumber * 100_001,
        vectorGroups: page.vectorGroups
      })
      nativeVectorShapeCount += converted.nativeShapeCount
      pageHasResidual ||= converted.residualHasVisualContent
      for (const warning of converted.warnings) warnings.add(warning)
      for (const layer of converted.layers) {
        archiveLayerNumber += 1
        if (layer.kind === 'drawingMl') {
          const anchorName = `typ2pptx-vector-anchor-${page.pageIndex + 1}-${archiveLayerNumber}`
          addVectorAnchor(pptx, slide, anchorName)
          vectorLayers.push({ anchorName, drawingMl: layer.drawingMl })
          continue
        }
        const fallbackPngBase64 = await rasterizeResidualSvg(
          request,
          layer.svg,
          page.width,
          page.height
        )
        assertPng(fallbackPngBase64)
        const imageName = `typ2pptx residual layer ${page.pageIndex + 1}.${archiveLayerNumber}`
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
          mediaName: `typ2pptx-residual-${pageOffset + 1}-${archiveLayerNumber}`,
          svg: layer.svg
        })
      }
    }
    // Compiler fallback items and SVG leaves are different units. Only claim that all
    // shape fallbacks disappeared when every ordered fallback layer has no residual.
    if (pageHasResidual) remainingFallbackShapeCount += page.fallbackShapeCount
    addEditableLinks(pptx, slide, page, slideNumbers)
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

async function addPageBackground(
  request: EditablePresentationRequest,
  slide: PptxGenJS.Slide,
  page: PresentationPageModel
): Promise<void> {
  const background = page.background
  if (!background) return
  if (background.kind === 'solid') {
    slide.background = { color: cleanHexColor(background.color) }
    return
  }
  const pngBase64 = await rasterizeBackgroundSvg(request, background, page.width, page.height)
  assertPng(pngBase64)
  slide.background = {
    data: `data:image/png;base64,${pngBase64}`,
    path: `typ2pptx-background-${page.pageIndex + 1}.png`
  }
}

async function rasterizeBackgroundSvg(
  request: EditablePresentationRequest,
  background: Extract<PresentationPageBackground, { kind: 'svg' }>,
  width: number,
  height: number
): Promise<string> {
  const rasterizer = request.residualSvgRasterizer ?? rasterizeSvgToPngBase64
  return rasterizer({ svg: background.svg, width, height })
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

function orderedPaintLayers(page: PresentationPageModel): PresentationPaintLayer[] {
  return [
    ...page.fallbackLayers,
    ...page.elements.filter((element): element is Exclude<PresentationPaintLayer, PresentationFallbackLayer> =>
      element.kind === 'text' || element.kind === 'rectangle' || element.kind === 'image')
  ].sort((left, right) => left.paintOrder - right.paintOrder)
}

function addEditableLinks(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  page: PresentationPageModel,
  slideNumbers: ReadonlyMap<number, number>
): void {
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
  elements: readonly PresentationTextElement[],
  pageWidth: number,
  budget: GeneratedTextRunBudget
): void {
  for (const element of elements) {
    validateElementBounds(element)
    if (typeof element.rtl !== 'boolean') {
      throw new Error('Compiler-positioned PPTX text requires an exact writing direction')
    }
  }
  const visible = elements.filter((element) => element.text)
  if (visible.length === 0) return
  const lines = compilerTextLines(visible)
  const fixed = lines.length > 1 ? fixedLineContainer(lines) : undefined
  if (lines.length > 1 && !fixed) {
    for (const line of lines) addEditableText(slide, line.elements, pageWidth, budget)
    return
  }
  const textBox = lines.length === 1 ? lines[0]!.textBox : undefined
  const container = fixed ?? textBox
  let left = container?.x ?? Math.min(...visible.map((element) => element.x))
  const firstLine = lines[0]!.metrics
  const top = firstLine.baseline - firstLine.fontSize * POWERPOINT_BASELINE_FROM_TOP_EM
  let right = container
    ? container.x + container.width
    : Math.max(...visible.map((element) => element.x + element.width))
  if (!container) {
    const metricHeadroom = Math.max(0.5, firstLine.fontSize * 0.12)
    if (visible[0]!.rtl) left = Math.max(0, left - metricHeadroom)
    else right = Math.min(pageWidth, right + metricHeadroom)
  }
  const bottom = Math.max(...visible.map((element) => element.y + element.height))
  const missingLineCount = lines.slice(1).reduce((count, line, index) => {
    const previous = lines[index]!.textBox
    return count + (previous && line.textBox
      ? line.textBox.lineIndex - previous.lineIndex - 1
      : 0)
  }, 0)
  consumeGeneratedTextRunBudget(budget, visible.length + missingLineCount)
  const runs: PptxGenJS.TextProps[] = lines.flatMap((line, lineOffset) => {
    const previous = lines[lineOffset - 1]?.textBox
    const missingLines = previous && line.textBox
      ? line.textBox.lineIndex - previous.lineIndex - 1
      : 0
    const first = line.elements[0]!
    const blankRuns = Array.from({ length: missingLines }, () => editableTextRun(
      first,
      line.metrics,
      '',
      true
    ))
    const lineRuns = line.elements.map((element, runIndex) => editableTextRun(
      element,
      line.metrics,
      element.text,
      lineOffset > 0 && runIndex === 0
    ))
    return [...blankRuns, ...lineRuns]
  })
  slide.addText(runs, {
    x: toInches(left),
    y: toInches(top),
    w: toInches(right - left),
    h: toInches(bottom - top),
    margin: 0,
    rtlMode: visible[0]!.rtl,
    fit: 'none',
    wrap: textBox?.reflow === true,
    ...(container ? { align: container.alignment } : {}),
    ...(fixed?.lineSpacing || textBox?.lineSpacing
      ? { lineSpacing: fixed?.lineSpacing ?? textBox?.lineSpacing }
      : {}),
    valign: 'top',
    isTextBox: true,
    objectName: 'Editable Typst text'
  })
}

function consumeGeneratedTextRunBudget(
  budget: GeneratedTextRunBudget,
  count: number
): void {
  if (!Number.isSafeInteger(count) || count < 0 || count > budget.remaining) {
    throw new Error(`Editable PPTX exceeds the ${MAX_ELEMENTS}-generated-text-run budget`)
  }
  budget.remaining -= count
}

function canShareTextContainer(
  current: readonly PresentationTextElement[],
  right: PresentationTextElement
): boolean {
  const first = current[0]?.textBox
  const previous = current[current.length - 1]?.textBox
  const next = right.textBox
  if (!first || !previous || !next) return false
  if (current.some((element) => element.rtl !== right.rtl)) return false
  if (next.id === previous.id) return true
  return !first.reflow &&
    !previous.reflow &&
    !next.reflow &&
    next.paragraphId === first.paragraphId &&
    next.lineIndex > previous.lineIndex &&
    approximatelyEqual(next.x, first.x) &&
    approximatelyEqual(next.width, first.width) &&
    compatibleFixedLineAlignments([...current.map((element) => element.textBox!.alignment),
      next.alignment])
}

function editableTextRun(
  element: PresentationTextElement,
  metrics: TextLineMetrics,
  text: string,
  softBreakBefore: boolean
): PptxGenJS.TextProps {
  return {
    text,
    options: {
      fontFace: element.fontFamily || 'Arial',
      fontSize: runFontSize(element, metrics),
      color: cleanHexColor(element.color),
      bold: element.bold,
      italic: element.italic,
      rtlMode: element.rtl,
      baseline: runBaseline(element, metrics),
      softBreakBefore,
      breakLine: false
    }
  }
}

interface CompilerTextLine {
  elements: PresentationTextElement[]
  textBox: PresentationTextBox | undefined
  metrics: TextLineMetrics
}

interface FixedLineContainer {
  x: number
  width: number
  alignment: PresentationTextBox['alignment']
  lineSpacing: number
}

function compilerTextLines(elements: readonly PresentationTextElement[]): CompilerTextLine[] {
  const groups: PresentationTextElement[][] = []
  for (const element of elements) {
    const previous = groups[groups.length - 1]
    if (
      previous &&
      element.textBox !== null &&
      previous[0]?.textBox?.id === element.textBox.id
    ) previous.push(element)
    else groups.push([element])
  }
  return groups.map((line) => ({
    elements: line,
    textBox: compilerOwnedTextBox(line),
    metrics: textLineMetrics(line)
  }))
}

function fixedLineContainer(lines: readonly CompilerTextLine[]): FixedLineContainer | undefined {
  const boxes = lines.map((line) => line.textBox)
  const first = boxes[0]
  if (
    !first ||
    first.reflow ||
    boxes.some((box, index) =>
      !box ||
      box.reflow ||
      box.paragraphId !== first.paragraphId ||
      (index > 0 && box.lineIndex <= boxes[index - 1]!.lineIndex) ||
      !approximatelyEqual(box.x, first.x) ||
      !approximatelyEqual(box.width, first.width))
  ) return undefined
  const alignments = boxes.map((box) => box!.alignment)
  if (!compatibleFixedLineAlignments(alignments)) return undefined
  const lineSpacing = consistentLineSpacing(lines)
  if (lineSpacing === undefined) return undefined
  return {
    x: first.x,
    width: first.width,
    alignment: alignments.includes('justify') ? 'justify' : alignments[0]!,
    lineSpacing
  }
}

function compatibleFixedLineAlignments(
  alignments: readonly PresentationTextBox['alignment'][]
): boolean {
  const first = alignments[0]
  return first !== undefined && (
    alignments.every((alignment) => alignment === first) ||
    alignments.every((alignment) => alignment === 'left' || alignment === 'justify')
  )
}

function consistentLineSpacing(lines: readonly CompilerTextLine[]): number | undefined {
  if (lines.length < 2) return undefined
  const deltas = lines.slice(1).map((line, index) => {
    const previous = lines[index]!
    const lineGap = line.textBox!.lineIndex - previous.textBox!.lineIndex
    return (line.metrics.baseline - previous.metrics.baseline) / lineGap
  })
  if (deltas.some((delta) => !Number.isFinite(delta) || delta <= 0)) return undefined
  const average = deltas.reduce((sum, delta) => sum + delta, 0) / deltas.length
  const tolerance = Math.max(0.05, average * 0.01)
  return deltas.every((delta) => Math.abs(delta - average) <= tolerance)
    ? average
    : undefined
}

function compilerOwnedTextBox(
  elements: readonly PresentationTextElement[]
): PresentationTextBox | undefined {
  if (elements.some((element) => !Number.isFinite(element.baseline))) {
    throw new Error('Compiler-positioned PPTX text requires an exact baseline')
  }
  const first = elements[0]?.textBox
  if (first === null) {
    if (elements.some((element) => element.textBox !== null)) {
      throw new Error('Compiler-isolated PPTX text runs cannot share a text box')
    }
    return undefined
  }
  if (first === undefined) {
    throw new Error('PPTX text requires compiler-owned or isolated text-box metadata')
  }
  if (!isValidCompilerTextBox(first)) {
    throw new Error('Compiler-owned PPTX text box has invalid geometry or identity')
  }
  if (elements.some((element) => !sameCompilerTextBox(first, element.textBox))) {
    throw new Error('Compiler-owned PPTX text box metadata must agree across its runs')
  }
  return first
}

function isValidCompilerTextBox(value: PresentationTextBox): boolean {
  return value.id.length > 0 &&
    value.id.length <= 256 &&
    value.paragraphId.length > 0 &&
    value.paragraphId.length <= 256 &&
    Number.isSafeInteger(value.lineIndex) &&
    value.lineIndex >= 0 &&
    value.lineIndex < MAX_ELEMENTS &&
    Number.isFinite(value.x) &&
    Number.isFinite(value.width) &&
    value.x >= 0 &&
    value.width > 0 &&
    ['left', 'center', 'right', 'justify'].includes(value.alignment) &&
    typeof value.reflow === 'boolean' &&
    (value.lineSpacing === undefined || (
      value.reflow && Number.isFinite(value.lineSpacing) && value.lineSpacing > 0
    ))
}

function sameCompilerTextBox(
  left: PresentationTextBox,
  right: PresentationTextBox | null | undefined
): boolean {
  return right !== null &&
    right !== undefined &&
    right.id === left.id &&
    right.paragraphId === left.paragraphId &&
    right.lineIndex === left.lineIndex &&
    approximatelyEqual(right.x, left.x) &&
    approximatelyEqual(right.width, left.width) &&
    right.alignment === left.alignment &&
    right.reflow === left.reflow &&
    (right.lineSpacing === left.lineSpacing || (
      right.lineSpacing !== undefined &&
      left.lineSpacing !== undefined &&
      approximatelyEqual(right.lineSpacing, left.lineSpacing)
    ))
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
  return element.baseline
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
  if (model.pages.some((page) => page.fallbackLayers.length > MAX_FALLBACK_LAYERS)) {
    throw new Error(`Editable PPTX exceeds the ${MAX_FALLBACK_LAYERS}-fallback-layer page budget`)
  }
  const vectorBytes = model.pages.reduce((total, page) => total + page.fallbackLayers.reduce(
    (pageTotal, layer) => pageTotal + new TextEncoder().encode(layer.svg).byteLength,
    0
  ), 0)
  if (vectorBytes > MAX_VECTOR_INPUT_BYTES) {
    throw new Error('Editable PPTX vector layers exceed the 64 MB input budget')
  }
  for (const page of model.pages) {
    assertPaintOrder(page)
    assertCompilerTextBoxOrder(page)
  }
}

function assertPaintOrder(page: PresentationPageModel): void {
  const orders = orderedPaintLayers(page).map((layer) => layer.paintOrder)
  if (
    orders.some((order) => !Number.isSafeInteger(order) || order < 0) ||
    orders.some((order, index) => order !== index)
  ) {
    throw new Error('Editable PPTX paint order must be unique and contiguous')
  }
}

function assertCompilerTextBoxOrder(page: PresentationPageModel): void {
  const definitions = new Map<string, PresentationTextBox>()
  const closed = new Set<string>()
  let active: string | undefined
  for (const layer of orderedPaintLayers(page)) {
    if (layer.kind !== 'text' || !layer.textBox) {
      if (active) closed.add(active)
      active = undefined
      continue
    }
    const textBox = layer.textBox
    if (
      !isValidCompilerTextBox(textBox) ||
      textBox.x + textBox.width > page.width + 0.01
    ) {
      throw new Error('Compiler-owned PPTX text box has invalid geometry or identity')
    }
    const opticalTolerance = Math.max(0.01, layer.fontSize * 0.5)
    if (
      layer.x < textBox.x - opticalTolerance ||
      layer.x + layer.width > textBox.x + textBox.width + opticalTolerance
    ) {
      throw new Error('PPTX text run is outside its compiler-owned text box')
    }
    const existing = definitions.get(textBox.id)
    if (existing && !sameCompilerTextBox(existing, textBox)) {
      throw new Error('Compiler-owned PPTX text box metadata must agree across its runs')
    }
    definitions.set(textBox.id, textBox)
    if (textBox.id !== active) {
      if (closed.has(textBox.id)) {
        throw new Error('Compiler-owned PPTX text box runs must be contiguous')
      }
      if (active) closed.add(active)
      active = textBox.id
    }
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

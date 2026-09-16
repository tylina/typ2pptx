export interface PresentationTextElement {
  kind: 'text'
  paintOrder: number
  x: number
  y: number
  width: number
  height: number
  text: string
  fontFamily: string
  fontSize: number
  baseline: number
  color: string
  bold: boolean
  italic: boolean
  /** Compiler-provided base writing direction. */
  rtl: boolean
  /**
   * A compiler-owned line or paragraph box, or `null` when the compiler
   * requires this run to stay isolated.
   */
  textBox: PresentationTextBox | null
}

export interface PresentationTextBox {
  id: string
  paragraphId: string
  lineIndex: number
  x: number
  width: number
  alignment: 'left' | 'center' | 'right' | 'justify'
  reflow: boolean
  lineSpacing?: number
}

export interface PresentationLinkElement {
  kind: 'link'
  x: number
  y: number
  width: number
  height: number
  url: string
}

export interface PresentationSlideLinkElement {
  kind: 'slideLink'
  x: number
  y: number
  width: number
  height: number
  pageIndex: number
}

export interface PresentationRectangleElement {
  kind: 'rectangle'
  paintOrder: number
  x: number
  y: number
  width: number
  height: number
  fillColor: string
}

export interface PresentationImageElement {
  kind: 'image'
  paintOrder: number
  x: number
  y: number
  width: number
  height: number
  mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  dataBase64: string
  altText: string
}

export type PresentationEditableElement =
  | PresentationTextElement
  | PresentationRectangleElement
  | PresentationImageElement
  | PresentationLinkElement
  | PresentationSlideLinkElement

export interface PresentationVectorGroup {
  id: string
  kind: 'math'
  x: number
  y: number
  width: number
  height: number
}

export interface PresentationFallbackLayer {
  kind: 'fallback'
  paintOrder: number
  svg: string
}

export interface PresentationPageModel {
  pageIndex: number
  width: number
  height: number
  fallbackLayers: PresentationFallbackLayer[]
  elements: PresentationEditableElement[]
  vectorGroups?: PresentationVectorGroup[]
  fallbackTextCount: number
  fallbackShapeCount: number
}

export interface PresentationExportModel {
  pages: PresentationPageModel[]
  fonts: string[]
  editableTextCount: number
  editableShapeCount: number
  fallbackTextCount: number
  fallbackShapeCount: number
  warnings: string[]
}

export interface VisualPresentationPage {
  pageIndex: number
  width: number
  height: number
  pngBase64: string
}

export interface PresentationArtifact {
  bytes: Uint8Array<ArrayBuffer>
  pageCount: number
  editableTextCount: number
  editableShapeCount: number
  editableLinkCount: number
  fallbackTextCount: number
  fallbackShapeCount: number
  nativeVectorShapeCount: number
  warnings: string[]
}

export interface VisualPresentationRequest {
  title: string
  pages: VisualPresentationPage[]
  notesByPageIndex?: ReadonlyMap<number, string>
}

export interface EditablePresentationRequest {
  title: string
  model: PresentationExportModel
  notesByPageIndex?: ReadonlyMap<number, string>
  residualSvgRasterizer?: ResidualSvgRasterizer
}

export interface ResidualSvgRasterizationRequest {
  svg: string
  width: number
  height: number
}

export type ResidualSvgRasterizer = (
  request: ResidualSvgRasterizationRequest
) => Promise<string>

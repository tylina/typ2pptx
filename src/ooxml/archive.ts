import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement
} from '@xmldom/xmldom'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const PRESENTATION_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const SVG_BLIP_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'

export interface SlideArchivePatch {
  slideNumber: number
  vectorLayers: Array<{ anchorName: string; drawingMl: string }>
  residualLayers: Array<{ imageName: string; mediaName: string; svg: string }>
}

/**
 * Add browser-generated DrawingML and Office SVG fallbacks to a PptxGenJS archive.
 *
 * PptxGenJS supplies package structure, editable text, links, and notes. This
 * final archive pass owns only the features that its public browser API cannot
 * express: arbitrary custom geometry and a compiler-supplied SVG image source.
 */
export function patchPresentationArchive(
  pptxBytes: Uint8Array<ArrayBuffer>,
  patches: readonly SlideArchivePatch[]
): Uint8Array<ArrayBuffer> {
  const files = unzipSync(pptxBytes)
  if (patches.some((patch) => patch.residualLayers.length > 0)) ensureSvgContentType(files)
  for (const patch of patches) patchSlide(files, patch)
  return exactArrayBufferView(zipSync(files, { level: 6 }))
}

function patchSlide(
  files: Record<string, Uint8Array>,
  patch: SlideArchivePatch
): void {
  const slidePath = `ppt/slides/slide${patch.slideNumber}.xml`
  const slide = parsePart(files, slidePath)
  for (const layer of patch.vectorLayers) {
    replaceVectorAnchor(
      slide,
      layer.anchorName,
      layer.drawingMl,
      slidePath
    )
  }
  for (const layer of patch.residualLayers) {
    embedResidualSvg(files, slide, slidePath, patch.slideNumber, layer)
  }
  files[slidePath] = strToU8(new XMLSerializer().serializeToString(slide))
}

function replaceVectorAnchor(
  slide: XmlDocument,
  anchorName: string,
  drawingMl: string,
  slidePath: string
): void {
  const anchor = findShapeByName(slide, anchorName)
  const parent = anchor.parentNode
  if (!parent) throw new Error(`PPTX vector anchor in ${slidePath} has no parent`)
  if (drawingMl) {
    const fragments = parseDrawingMlFragments(drawingMl)
    for (const fragment of fragments) {
      parent.insertBefore(slide.importNode(fragment, true), anchor)
    }
  }
  parent.removeChild(anchor)
}

function parseDrawingMlFragments(value: string): XmlElement[] {
  const wrapper = parseXml(
    `<typ2pptx xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">${value}</typ2pptx>`,
    'generated DrawingML'
  )
  const root = wrapper.documentElement
  if (!root) throw new Error('Generated DrawingML is empty')
  const result: XmlElement[] = []
  for (let index = 0; index < root.childNodes.length; index += 1) {
    const node = root.childNodes.item(index)
    if (node?.nodeType === 1) result.push(node as XmlElement)
  }
  return result
}

function embedResidualSvg(
  files: Record<string, Uint8Array>,
  slide: XmlDocument,
  slidePath: string,
  slideNumber: number,
  layer: { imageName: string; mediaName: string; svg: string }
): void {
  const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`
  const rels = parsePart(files, relsPath)
  const picture = findPictureByName(slide, layer.imageName)
  const blip = requireOnlyElement(picture, DRAWING_NS, 'blip', slidePath)
  const pngRelationshipId = blip.getAttributeNS(OFFICE_REL_NS, 'embed')
  if (!pngRelationshipId) throw new Error(`PPTX residual image in ${slidePath} has no PNG relationship`)
  if (blip.getElementsByTagNameNS(DRAWING_NS, 'extLst').length > 0) {
    throw new Error(`PPTX residual image in ${slidePath} already has an extension list`)
  }

  const svgRelationshipId = nextRelationshipId(rels)
  const relationship = rels.createElementNS(PACKAGE_REL_NS, 'Relationship')
  relationship.setAttribute('Id', svgRelationshipId)
  relationship.setAttribute(
    'Type',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  )
  relationship.setAttribute('Target', `../media/${layer.mediaName}.svg`)
  const relationshipRoot = rels.documentElement
  if (!relationshipRoot) throw new Error(`PPTX archive contains an empty XML part: ${relsPath}`)
  relationshipRoot.appendChild(relationship)

  const extensionList = slide.createElementNS(DRAWING_NS, 'a:extLst')
  const extension = slide.createElementNS(DRAWING_NS, 'a:ext')
  extension.setAttribute('uri', SVG_EXTENSION_URI)
  const svgBlip = slide.createElementNS(SVG_BLIP_NS, 'asvg:svgBlip')
  svgBlip.setAttribute('xmlns:asvg', SVG_BLIP_NS)
  svgBlip.setAttributeNS(OFFICE_REL_NS, 'r:embed', svgRelationshipId)
  extension.appendChild(svgBlip)
  extensionList.appendChild(extension)
  blip.appendChild(extensionList)

  files[relsPath] = strToU8(new XMLSerializer().serializeToString(rels))
  files[`ppt/media/${layer.mediaName}.svg`] = strToU8(layer.svg)
}

function ensureSvgContentType(files: Record<string, Uint8Array>): void {
  const path = '[Content_Types].xml'
  const document = parsePart(files, path)
  const defaults = document.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Default')
  for (let index = 0; index < defaults.length; index += 1) {
    if (defaults.item(index)?.getAttribute('Extension')?.toLocaleLowerCase() === 'svg') return
  }
  const entry = document.createElementNS(CONTENT_TYPES_NS, 'Default')
  entry.setAttribute('Extension', 'svg')
  entry.setAttribute('ContentType', 'image/svg+xml')
  const root = document.documentElement
  if (!root) throw new Error(`PPTX archive contains an empty XML part: ${path}`)
  root.appendChild(entry)
  files[path] = strToU8(new XMLSerializer().serializeToString(document))
}

function findShapeByName(document: XmlDocument, name: string): XmlElement {
  return findObjectByName(document, 'sp', name, 'vector anchor')
}

function findPictureByName(document: XmlDocument, name: string): XmlElement {
  return findObjectByName(document, 'pic', name, 'residual image')
}

function findObjectByName(
  document: XmlDocument,
  localName: string,
  name: string,
  label: string
): XmlElement {
  const objects = document.getElementsByTagNameNS(PRESENTATION_NS, localName)
  for (let index = 0; index < objects.length; index += 1) {
    const object = objects.item(index)
    if (!object) continue
    const properties = object.getElementsByTagNameNS(PRESENTATION_NS, 'cNvPr')
    for (let propertyIndex = 0; propertyIndex < properties.length; propertyIndex += 1) {
      if (properties.item(propertyIndex)?.getAttribute('name') === name) return object
    }
  }
  throw new Error(`PPTX slide is missing its ${label}: ${name}`)
}

function nextRelationshipId(document: XmlDocument): string {
  const relationships = document.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship')
  let maximum = 0
  for (let index = 0; index < relationships.length; index += 1) {
    const id = relationships.item(index)?.getAttribute('Id') ?? ''
    if (!id.startsWith('rId')) continue
    const value = Number(id.slice(3))
    if (Number.isSafeInteger(value) && value > maximum) maximum = value
  }
  return `rId${maximum + 1}`
}

function requireOnlyElement(
  parent: XmlElement,
  namespace: string,
  localName: string,
  part: string
): XmlElement {
  const elements = parent.getElementsByTagNameNS(namespace, localName)
  if (elements.length !== 1 || !elements.item(0)) {
    throw new Error(`PPTX ${part} has an invalid ${localName} element`)
  }
  return elements.item(0)!
}

function parsePart(files: Record<string, Uint8Array>, path: string): XmlDocument {
  const bytes = files[path]
  if (!bytes) throw new Error(`PPTX archive is missing ${path}`)
  return parseXml(strFromU8(bytes), path)
}

function parseXml(value: string, label: string): XmlDocument {
  const errors: string[] = []
  const document = new DOMParser({
    onError: (level, message) => {
      if (level !== 'warning') errors.push(message)
    }
  }).parseFromString(value, 'application/xml')
  if (errors.length > 0 || document.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`PPTX archive contains invalid XML in ${label}`)
  }
  return document
}

function exactArrayBufferView(value: Uint8Array): Uint8Array<ArrayBuffer> {
  if (
    value.buffer instanceof ArrayBuffer &&
    value.byteOffset === 0 &&
    value.byteLength === value.buffer.byteLength
  ) return value as Uint8Array<ArrayBuffer>
  return value.slice()
}

export type AffineMatrix = readonly [
  a: number,
  b: number,
  c: number,
  d: number,
  e: number,
  f: number
]

export const IDENTITY_MATRIX: AffineMatrix = [1, 0, 0, 1, 0, 0]

const TRANSFORM_PATTERN = /([a-z]+)\s*\(([^)]*)\)/giu
const NUMBER_PATTERN = /[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/giu

export function multiplyMatrices(left: AffineMatrix, right: AffineMatrix): AffineMatrix {
  const [a1, b1, c1, d1, e1, f1] = left
  const [a2, b2, c2, d2, e2, f2] = right
  return [
    a1 * a2 + c1 * b2,
    b1 * a2 + d1 * b2,
    a1 * c2 + c1 * d2,
    b1 * c2 + d1 * d2,
    a1 * e2 + c1 * f2 + e1,
    b1 * e2 + d1 * f2 + f1
  ]
}

export function applyMatrix(
  matrix: AffineMatrix,
  x: number,
  y: number
): readonly [number, number] {
  const [a, b, c, d, e, f] = matrix
  return [a * x + c * y + e, b * x + d * y + f]
}

export function parseTransformList(value: string | null | undefined): AffineMatrix {
  if (!value?.trim()) return IDENTITY_MATRIX
  let matrix = IDENTITY_MATRIX
  let consumed = ''
  for (const match of value.matchAll(TRANSFORM_PATTERN)) {
    consumed += match[0]
    const operation = match[1]?.toLocaleLowerCase()
    const values = [...(match[2] ?? '').matchAll(NUMBER_PATTERN)].map((item) => Number(item[0]))
    if (values.some((item) => !Number.isFinite(item))) {
      throw new Error('SVG transform contains a non-finite number')
    }
    matrix = multiplyMatrices(matrix, operationMatrix(operation ?? '', values))
  }
  const unmatched = value.replace(TRANSFORM_PATTERN, '').replace(/[\s,]+/gu, '')
  if (unmatched || !consumed) throw new Error(`Unsupported SVG transform: ${value}`)
  return matrix
}

function operationMatrix(operation: string, values: number[]): AffineMatrix {
  if (operation === 'matrix' && values.length === 6) {
    return values as unknown as AffineMatrix
  }
  if (operation === 'translate' && (values.length === 1 || values.length === 2)) {
    return [1, 0, 0, 1, values[0] ?? 0, values[1] ?? 0]
  }
  if (operation === 'scale' && (values.length === 1 || values.length === 2)) {
    return [values[0] ?? 1, 0, 0, values[1] ?? values[0] ?? 1, 0, 0]
  }
  if (operation === 'rotate' && (values.length === 1 || values.length === 3)) {
    const radians = ((values[0] ?? 0) * Math.PI) / 180
    const rotation: AffineMatrix = [
      Math.cos(radians),
      Math.sin(radians),
      -Math.sin(radians),
      Math.cos(radians),
      0,
      0
    ]
    if (values.length === 1) return rotation
    const x = values[1] ?? 0
    const y = values[2] ?? 0
    return multiplyMatrices(
      multiplyMatrices([1, 0, 0, 1, x, y], rotation),
      [1, 0, 0, 1, -x, -y]
    )
  }
  if (operation === 'skewx' && values.length === 1) {
    return [1, 0, Math.tan(((values[0] ?? 0) * Math.PI) / 180), 1, 0, 0]
  }
  if (operation === 'skewy' && values.length === 1) {
    return [1, Math.tan(((values[0] ?? 0) * Math.PI) / 180), 0, 1, 0, 0]
  }
  throw new Error(`Unsupported SVG transform operation: ${operation}`)
}

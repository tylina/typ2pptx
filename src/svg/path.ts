import { applyMatrix, IDENTITY_MATRIX, type AffineMatrix } from './matrix'

export interface SvgPathCommand {
  command: string
  values: number[]
}

export interface DrawingMlPathGeometry {
  xml: string
  bounds: { x: number; y: number; width: number; height: number }
}

export interface DrawingMlPathOptions {
  emuPerUnit?: number
  transform?: AffineMatrix
}

const ARGUMENT_COUNTS: Readonly<Record<string, number>> = {
  M: 2, m: 2, L: 2, l: 2, H: 1, h: 1, V: 1, v: 1,
  C: 6, c: 6, S: 4, s: 4, Q: 4, q: 4, T: 2, t: 2,
  A: 7, a: 7, Z: 0, z: 0
}
const TOKEN_PATTERN = /[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gu
const COMMAND_PATTERN = /^[MmLlHhVvCcSsQqTtAaZz]$/u
const MAX_PATH_COMMANDS = 200_000

export function parseSvgPath(value: string): SvgPathCommand[] {
  if (!value.trim()) return []
  const tokens = value.match(TOKEN_PATTERN) ?? []
  const ignored = value.replace(TOKEN_PATTERN, '').replace(/[\s,]+/gu, '')
  if (ignored) throw new Error('SVG path contains unsupported syntax')
  const commands: SvgPathCommand[] = []
  let command: string | undefined
  let values: number[] = []

  const flush = () => {
    if (!command) return
    const count = ARGUMENT_COUNTS[command]
    if (count === undefined) throw new Error(`Unsupported SVG path command: ${command}`)
    if (count === 0) {
      if (values.length > 0) throw new Error(`SVG ${command} command cannot contain coordinates`)
      commands.push({ command, values: [] })
      return
    }
    if (values.length === 0 || values.length % count !== 0) {
      throw new Error(`SVG ${command} command has an invalid coordinate count`)
    }
    for (let index = 0; index < values.length; index += count) {
      const current = command === 'M' && index > 0
        ? 'L'
        : command === 'm' && index > 0
          ? 'l'
          : command
      commands.push({ command: current, values: values.slice(index, index + count) })
      if (commands.length > MAX_PATH_COMMANDS) throw new Error('SVG path exceeds the command budget')
    }
    values = []
  }

  for (const token of tokens) {
    if (COMMAND_PATTERN.test(token)) {
      flush()
      command = token
      values = []
      if ((ARGUMENT_COUNTS[command] ?? -1) === 0) {
        flush()
        command = undefined
      }
      continue
    }
    if (!command) throw new Error('SVG path coordinate appears before a command')
    const number = Number(token)
    if (!Number.isFinite(number)) throw new Error('SVG path contains a non-finite number')
    values.push(number)
  }
  flush()
  return commands
}

export function normalizeSvgPath(source: readonly SvgPathCommand[]): SvgPathCommand[] {
  const absolute = absolutize(source)
  const result: SvgPathCommand[] = []
  let x = 0
  let y = 0
  let subpathX = 0
  let subpathY = 0
  let cubicControlX = 0
  let cubicControlY = 0
  let quadraticControlX = 0
  let quadraticControlY = 0
  let previous = ''

  for (const item of absolute) {
    const values = item.values
    if (item.command === 'M') {
      x = values[0] ?? 0
      y = values[1] ?? 0
      subpathX = x
      subpathY = y
      result.push(item)
    } else if (item.command === 'L') {
      x = values[0] ?? 0
      y = values[1] ?? 0
      result.push(item)
    } else if (item.command === 'C') {
      cubicControlX = values[2] ?? x
      cubicControlY = values[3] ?? y
      x = values[4] ?? x
      y = values[5] ?? y
      result.push(item)
    } else if (item.command === 'S') {
      const firstX = previous === 'C' || previous === 'S' ? 2 * x - cubicControlX : x
      const firstY = previous === 'C' || previous === 'S' ? 2 * y - cubicControlY : y
      cubicControlX = values[0] ?? x
      cubicControlY = values[1] ?? y
      x = values[2] ?? x
      y = values[3] ?? y
      result.push({
        command: 'C',
        values: [firstX, firstY, cubicControlX, cubicControlY, x, y]
      })
    } else if (item.command === 'Q') {
      quadraticControlX = values[0] ?? x
      quadraticControlY = values[1] ?? y
      const endX = values[2] ?? x
      const endY = values[3] ?? y
      result.push({
        command: 'C',
        values: quadraticToCubic(x, y, quadraticControlX, quadraticControlY, endX, endY)
      })
      x = endX
      y = endY
    } else if (item.command === 'T') {
      quadraticControlX = previous === 'Q' || previous === 'T'
        ? 2 * x - quadraticControlX
        : x
      quadraticControlY = previous === 'Q' || previous === 'T'
        ? 2 * y - quadraticControlY
        : y
      const endX = values[0] ?? x
      const endY = values[1] ?? y
      result.push({
        command: 'C',
        values: quadraticToCubic(x, y, quadraticControlX, quadraticControlY, endX, endY)
      })
      x = endX
      y = endY
    } else if (item.command === 'A') {
      const endX = values[5] ?? x
      const endY = values[6] ?? y
      result.push(...arcToCubics(
        x,
        y,
        Math.abs(values[0] ?? 0),
        Math.abs(values[1] ?? 0),
        values[2] ?? 0,
        (values[3] ?? 0) !== 0,
        (values[4] ?? 0) !== 0,
        endX,
        endY
      ))
      x = endX
      y = endY
    } else if (item.command === 'Z') {
      result.push(item)
      x = subpathX
      y = subpathY
    } else {
      throw new Error(`Unsupported normalized SVG path command: ${item.command}`)
    }
    if (item.command !== 'C' && item.command !== 'S') {
      cubicControlX = x
      cubicControlY = y
    }
    if (item.command !== 'Q' && item.command !== 'T') {
      quadraticControlX = x
      quadraticControlY = y
    }
    previous = item.command
  }
  return result
}

export function pathToDrawingMlGeometry(
  value: string,
  options: DrawingMlPathOptions = {}
): DrawingMlPathGeometry {
  const commands = normalizeSvgPath(parseSvgPath(value))
  if (commands.length === 0) throw new Error('SVG path has no drawable commands')
  const matrix = options.transform ?? IDENTITY_MATRIX
  const transformed = commands.map((item) => ({
    command: item.command,
    values: transformCommand(item, matrix)
  }))
  const points = transformed.flatMap(commandPoints)
  if (points.length === 0) throw new Error('SVG path has no finite geometry')
  const xs = points.map(([x]) => x)
  const ys = points.map(([, y]) => y)
  const x = Math.min(...xs)
  const y = Math.min(...ys)
  const width = Math.max(Math.max(...xs) - x, 1 / (options.emuPerUnit ?? 12_700))
  const height = Math.max(Math.max(...ys) - y, 1 / (options.emuPerUnit ?? 12_700))
  if (![x, y, width, height].every(Number.isFinite)) {
    throw new Error('SVG path produced non-finite geometry')
  }
  const emuPerUnit = positiveFinite(options.emuPerUnit ?? 12_700, 'EMU scale')
  const widthEmu = toEmu(width, emuPerUnit)
  const heightEmu = toEmu(height, emuPerUnit)
  const body = transformed.map((item) => drawingMlCommand(item, x, y, emuPerUnit)).join('')
  return {
    xml: `<a:path w="${widthEmu}" h="${heightEmu}">${body}</a:path>`,
    bounds: { x, y, width, height }
  }
}

function absolutize(source: readonly SvgPathCommand[]): SvgPathCommand[] {
  const result: SvgPathCommand[] = []
  let x = 0
  let y = 0
  let subpathX = 0
  let subpathY = 0
  for (const item of source) {
    const relative = item.command === item.command.toLocaleLowerCase()
    const command = item.command.toLocaleUpperCase()
    const values = [...item.values]
    const point = (index: number) => {
      if (!relative) return
      values[index] = (values[index] ?? 0) + x
      values[index + 1] = (values[index + 1] ?? 0) + y
    }
    if (command === 'M' || command === 'L' || command === 'T') point(0)
    if (command === 'C') {
      point(0)
      point(2)
      point(4)
    }
    if (command === 'S' || command === 'Q') {
      point(0)
      point(2)
    }
    if (command === 'A' && relative) {
      values[5] = (values[5] ?? 0) + x
      values[6] = (values[6] ?? 0) + y
    }
    if (command === 'H') {
      const next = (values[0] ?? 0) + (relative ? x : 0)
      result.push({ command: 'L', values: [next, y] })
      x = next
      continue
    }
    if (command === 'V') {
      const next = (values[0] ?? 0) + (relative ? y : 0)
      result.push({ command: 'L', values: [x, next] })
      y = next
      continue
    }
    result.push({ command, values })
    if (command === 'M') {
      x = values[0] ?? x
      y = values[1] ?? y
      subpathX = x
      subpathY = y
    } else if (command === 'L' || command === 'T') {
      x = values[0] ?? x
      y = values[1] ?? y
    } else if (command === 'C') {
      x = values[4] ?? x
      y = values[5] ?? y
    } else if (command === 'S' || command === 'Q') {
      x = values[2] ?? x
      y = values[3] ?? y
    } else if (command === 'A') {
      x = values[5] ?? x
      y = values[6] ?? y
    } else if (command === 'Z') {
      x = subpathX
      y = subpathY
    }
  }
  return result
}

function quadraticToCubic(
  x0: number,
  y0: number,
  controlX: number,
  controlY: number,
  x3: number,
  y3: number
): number[] {
  return [
    x0 + (2 / 3) * (controlX - x0),
    y0 + (2 / 3) * (controlY - y0),
    x3 + (2 / 3) * (controlX - x3),
    y3 + (2 / 3) * (controlY - y3),
    x3,
    y3
  ]
}

function arcToCubics(
  x1: number,
  y1: number,
  sourceRx: number,
  sourceRy: number,
  angle: number,
  largeArc: boolean,
  sweep: boolean,
  x2: number,
  y2: number
): SvgPathCommand[] {
  if (Math.abs(x1 - x2) < 1e-10 && Math.abs(y1 - y2) < 1e-10) return []
  if (sourceRx < 1e-10 || sourceRy < 1e-10) {
    return [{ command: 'L', values: [x2, y2] }]
  }
  const phi = (angle * Math.PI) / 180
  const cosPhi = Math.cos(phi)
  const sinPhi = Math.sin(phi)
  const dx = (x1 - x2) / 2
  const dy = (y1 - y2) / 2
  const x1Prime = cosPhi * dx + sinPhi * dy
  const y1Prime = -sinPhi * dx + cosPhi * dy
  let rx = sourceRx
  let ry = sourceRy
  const scale = (x1Prime * x1Prime) / (rx * rx) + (y1Prime * y1Prime) / (ry * ry)
  if (scale > 1) {
    const factor = Math.sqrt(scale)
    rx *= factor
    ry *= factor
  }
  const rx2 = rx * rx
  const ry2 = ry * ry
  const xPrime2 = x1Prime * x1Prime
  const yPrime2 = y1Prime * y1Prime
  const numerator = Math.max(rx2 * ry2 - rx2 * yPrime2 - ry2 * xPrime2, 0)
  const denominator = rx2 * yPrime2 + ry2 * xPrime2
  let factor = denominator > 1e-10 ? Math.sqrt(numerator / denominator) : 0
  if (largeArc === sweep) factor = -factor
  const centerXPrime = factor * rx * y1Prime / ry
  const centerYPrime = -factor * ry * x1Prime / rx
  const centerX = cosPhi * centerXPrime - sinPhi * centerYPrime + (x1 + x2) / 2
  const centerY = sinPhi * centerXPrime + cosPhi * centerYPrime + (y1 + y2) / 2
  const vectorAngle = (ux: number, uy: number, vx: number, vy: number) => {
    const denominator = Math.hypot(ux, uy) * Math.hypot(vx, vy)
    if (denominator < 1e-10) return 0
    const radians = Math.acos(Math.max(-1, Math.min(1, (ux * vx + uy * vy) / denominator)))
    return ux * vy - uy * vx < 0 ? -radians : radians
  }
  const start = vectorAngle(
    1,
    0,
    (x1Prime - centerXPrime) / rx,
    (y1Prime - centerYPrime) / ry
  )
  let delta = vectorAngle(
    (x1Prime - centerXPrime) / rx,
    (y1Prime - centerYPrime) / ry,
    (-x1Prime - centerXPrime) / rx,
    (-y1Prime - centerYPrime) / ry
  )
  if (!sweep && delta > 0) delta -= 2 * Math.PI
  if (sweep && delta < 0) delta += 2 * Math.PI
  const segments = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2)))
  const step = delta / segments
  const alpha = (4 / 3) * Math.tan(step / 4)
  const transform = (x: number, y: number): readonly [number, number] => [
    cosPhi * rx * x - sinPhi * ry * y + centerX,
    sinPhi * rx * x + cosPhi * ry * y + centerY
  ]
  return Array.from({ length: segments }, (_, index) => {
    const first = start + index * step
    const second = first + step
    const c1 = transform(Math.cos(first) - alpha * Math.sin(first),
      Math.sin(first) + alpha * Math.cos(first))
    const c2 = transform(Math.cos(second) + alpha * Math.sin(second),
      Math.sin(second) - alpha * Math.cos(second))
    const end = transform(Math.cos(second), Math.sin(second))
    return { command: 'C', values: [...c1, ...c2, ...end] }
  })
}

function transformCommand(item: SvgPathCommand, matrix: AffineMatrix): number[] {
  const values: number[] = []
  for (let index = 0; index < item.values.length; index += 2) {
    const [x, y] = applyMatrix(matrix, item.values[index] ?? 0, item.values[index + 1] ?? 0)
    values.push(x, y)
  }
  return values
}

function commandPoints(item: SvgPathCommand): Array<readonly [number, number]> {
  const result: Array<readonly [number, number]> = []
  for (let index = 0; index < item.values.length; index += 2) {
    result.push([item.values[index] ?? 0, item.values[index + 1] ?? 0])
  }
  return result
}

function drawingMlCommand(
  item: SvgPathCommand,
  originX: number,
  originY: number,
  emuPerUnit: number
): string {
  const point = (index: number) => `<a:pt x="${toEmu(
    (item.values[index] ?? 0) - originX,
    emuPerUnit
  )}" y="${toEmu((item.values[index + 1] ?? 0) - originY, emuPerUnit)}"/>`
  if (item.command === 'M') return `<a:moveTo>${point(0)}</a:moveTo>`
  if (item.command === 'L') return `<a:lnTo>${point(0)}</a:lnTo>`
  if (item.command === 'C') {
    return `<a:cubicBezTo>${point(0)}${point(2)}${point(4)}</a:cubicBezTo>`
  }
  if (item.command === 'Z') return '<a:close/>'
  throw new Error(`Unsupported DrawingML path command: ${item.command}`)
}

export function toEmu(value: number, emuPerUnit = 12_700): number {
  return Math.round(value * emuPerUnit)
}

function positiveFinite(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid ${label}`)
  return value
}

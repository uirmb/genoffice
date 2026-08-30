export interface InlineImageDimensions {
  width: number
  height: number
}

const JPEG_HEADER_LIMIT = 256 * 1024

function decodeBase64Prefix(payload: string, maxBytes: number): Uint8Array | null {
  const requestedChars = Math.ceil((maxBytes * 4) / 3)
  const availableChars = Math.min(payload.length, requestedChars)
  const alignedChars =
    availableChars === payload.length
      ? availableChars
      : Math.max(4, availableChars - (availableChars % 4))

  let chunk = payload.slice(0, alignedChars).replace(/-/g, '+').replace(/_/g, '/')
  if (chunk.length % 4 !== 0) chunk += '='.repeat(4 - (chunk.length % 4))

  try {
    const binary = atob(chunk)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
    return bytes
  } catch {
    return null
  }
}

function validSize(width: number, height: number): InlineImageDimensions | null {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null
  return { width, height }
}

function readUint32Be(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]! * 0x1000000 +
    bytes[offset + 1]! * 0x10000 +
    bytes[offset + 2]! * 0x100 +
    bytes[offset + 3]!
  )
}

function parsePng(payload: string): InlineImageDimensions | null {
  const bytes = decodeBase64Prefix(payload, 24)
  if (!bytes || bytes.length < 24) return null
  if (
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47 ||
    bytes[4] !== 0x0d ||
    bytes[5] !== 0x0a ||
    bytes[6] !== 0x1a ||
    bytes[7] !== 0x0a
  ) {
    return null
  }
  return validSize(readUint32Be(bytes, 16), readUint32Be(bytes, 20))
}

function parseGif(payload: string): InlineImageDimensions | null {
  const bytes = decodeBase64Prefix(payload, 10)
  if (!bytes || bytes.length < 10) return null
  const signature = String.fromCharCode(...bytes.subarray(0, 6))
  if (signature !== 'GIF87a' && signature !== 'GIF89a') return null
  return validSize(bytes[6]! | (bytes[7]! << 8), bytes[8]! | (bytes[9]! << 8))
}

function isJpegSof(marker: number): boolean {
  return (
    (marker >= 0xc0 && marker <= 0xc3) ||
    (marker >= 0xc5 && marker <= 0xc7) ||
    (marker >= 0xc9 && marker <= 0xcb) ||
    (marker >= 0xcd && marker <= 0xcf)
  )
}

function parseJpeg(payload: string): InlineImageDimensions | null {
  const bytes = decodeBase64Prefix(payload, JPEG_HEADER_LIMIT)
  if (!bytes || bytes.length < 10 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null

  let offset = 2
  while (offset + 3 < bytes.length) {
    while (offset < bytes.length && bytes[offset] !== 0xff) offset += 1
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1
    if (offset >= bytes.length) return null

    const marker = bytes[offset]!
    offset += 1

    if (marker === 0xd9 || marker === 0xda) return null
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue
    if (offset + 1 >= bytes.length) return null

    const segmentLength = (bytes[offset]! << 8) | bytes[offset + 1]!
    if (segmentLength < 2 || offset + segmentLength > bytes.length) return null

    if (isJpegSof(marker) && segmentLength >= 7) {
      const height = (bytes[offset + 3]! << 8) | bytes[offset + 4]!
      const width = (bytes[offset + 5]! << 8) | bytes[offset + 6]!
      return validSize(width, height)
    }

    offset += segmentLength
  }

  return null
}

function ascii(bytes: Uint8Array, offset: number, length: number): string {
  return String.fromCharCode(...bytes.subarray(offset, offset + length))
}

function readUint24Le(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8) | (bytes[offset + 2]! << 16)
}

function parseWebp(payload: string): InlineImageDimensions | null {
  const bytes = decodeBase64Prefix(payload, 64)
  if (!bytes || bytes.length < 30) return null
  if (ascii(bytes, 0, 4) !== 'RIFF' || ascii(bytes, 8, 4) !== 'WEBP') return null

  const chunk = ascii(bytes, 12, 4)
  if (chunk === 'VP8X') {
    return validSize(1 + readUint24Le(bytes, 24), 1 + readUint24Le(bytes, 27))
  }

  if (chunk === 'VP8 ' && bytes.length >= 30) {
    if (bytes[23] !== 0x9d || bytes[24] !== 0x01 || bytes[25] !== 0x2a) return null
    const width = (bytes[26]! | (bytes[27]! << 8)) & 0x3fff
    const height = (bytes[28]! | (bytes[29]! << 8)) & 0x3fff
    return validSize(width, height)
  }

  if (chunk === 'VP8L' && bytes.length >= 25 && bytes[20] === 0x2f) {
    const b1 = bytes[21]!
    const b2 = bytes[22]!
    const b3 = bytes[23]!
    const b4 = bytes[24]!
    const width = 1 + (b1 | ((b2 & 0x3f) << 8))
    const height = 1 + ((b2 >> 6) | (b3 << 2) | ((b4 & 0x0f) << 10))
    return validSize(width, height)
  }

  return null
}

/**
 * Reads only the small image header needed for intrinsic dimensions. The full
 * Base64 payload is never decoded, so a 40+ MB inline image stays cheap during
 * first paint.
 */
export function readInlineImageDimensions(src: string): InlineImageDimensions | null {
  if (!src.startsWith('data:image/')) return null
  const marker = ';base64,'
  const markerAt = src.indexOf(marker)
  if (markerAt < 0 || markerAt > 96) return null

  const mime = src.slice('data:'.length, markerAt).toLowerCase()
  const payload = src.slice(markerAt + marker.length)
  if (!payload) return null

  if (mime === 'image/png') return parsePng(payload)
  if (mime === 'image/jpeg' || mime === 'image/jpg') return parseJpeg(payload)
  if (mime === 'image/gif') return parseGif(payload)
  if (mime === 'image/webp') return parseWebp(payload)
  return null
}

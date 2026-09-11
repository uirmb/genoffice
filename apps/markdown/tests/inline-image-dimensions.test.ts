import { describe, expect, it } from 'vitest'
import { readInlineImageDimensions } from '../src/renderer/editor/inlineImageDimensions'

function asDataUrl(mime: string, bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return `data:${mime};base64,${btoa(binary)}`
}

function writeUint32Be(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = (value >>> 24) & 0xff
  bytes[offset + 1] = (value >>> 16) & 0xff
  bytes[offset + 2] = (value >>> 8) & 0xff
  bytes[offset + 3] = value & 0xff
}

describe('inline image intrinsic dimensions', () => {
  it('reads PNG dimensions from the first 24 bytes', () => {
    const bytes = new Uint8Array(24)
    bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0)
    bytes.set([0x49, 0x48, 0x44, 0x52], 12)
    writeUint32Be(bytes, 16, 3840)
    writeUint32Be(bytes, 20, 2160)

    expect(readInlineImageDimensions(asDataUrl('image/png', bytes))).toEqual({
      width: 3840,
      height: 2160,
    })
  })

  it('reads JPEG SOF dimensions without decoding the image body', () => {
    const bytes = new Uint8Array(27)
    bytes.set([0xff, 0xd8], 0)
    bytes.set([0xff, 0xe0, 0x00, 0x04, 0x00, 0x00], 2)
    bytes.set([0xff, 0xc0, 0x00, 0x11, 0x08, 0x04, 0x38, 0x07, 0x80], 8)

    expect(readInlineImageDimensions(asDataUrl('image/jpeg', bytes))).toEqual({
      width: 1920,
      height: 1080,
    })
  })

  it('reads GIF logical-screen dimensions', () => {
    const bytes = new Uint8Array(10)
    bytes.set([...new TextEncoder().encode('GIF89a'), 0x20, 0x03, 0x58, 0x02])

    expect(readInlineImageDimensions(asDataUrl('image/gif', bytes))).toEqual({
      width: 800,
      height: 600,
    })
  })

  it('reads WebP VP8X dimensions', () => {
    const bytes = new Uint8Array(30)
    bytes.set(new TextEncoder().encode('RIFF'), 0)
    bytes.set(new TextEncoder().encode('WEBP'), 8)
    bytes.set(new TextEncoder().encode('VP8X'), 12)
    const widthMinusOne = 1279
    const heightMinusOne = 719
    bytes.set(
      [
        widthMinusOne & 0xff,
        (widthMinusOne >>> 8) & 0xff,
        (widthMinusOne >>> 16) & 0xff,
        heightMinusOne & 0xff,
        (heightMinusOne >>> 8) & 0xff,
        (heightMinusOne >>> 16) & 0xff,
      ],
      24,
    )

    expect(readInlineImageDimensions(asDataUrl('image/webp', bytes))).toEqual({
      width: 1280,
      height: 720,
    })
  })

  it('returns null for unsupported or malformed sources', () => {
    expect(readInlineImageDimensions('https://example.com/image.png')).toBeNull()
    expect(readInlineImageDimensions('data:image/png;base64,AAAA')).toBeNull()
  })
})

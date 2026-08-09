const utf8Encoder = new TextEncoder()
const utf8Decoder = new TextDecoder('utf-8')

export function encodeUtf8(value: string): Uint8Array {
  return utf8Encoder.encode(value)
}

export function decodeUtf8(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes)
}

export function encodeAscii(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length)
  for (let i = 0; i < value.length; i++) bytes[i] = value.charCodeAt(i) & 0x7f
  return bytes
}

export function decodeBase64(value: string): Uint8Array {
  if (typeof atob === 'function') {
    const binary = atob(value)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
    return bytes
  }

  const nodeBuffer = (globalThis as { Buffer?: { from(value: string, encoding: string): Uint8Array } })
    .Buffer
  if (nodeBuffer) return new Uint8Array(nodeBuffer.from(value, 'base64'))
  throw new Error('Base64 decoding is unavailable in this runtime.')
}

export function encodeBase64(bytes: Uint8Array): string {
  if (typeof btoa === 'function') {
    let binary = ''
    const chunk = 0x8000
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, Math.min(i + chunk, bytes.length)))
    }
    return btoa(binary)
  }

  const nodeBuffer = (globalThis as {
    Buffer?: { from(value: Uint8Array): { toString(encoding: string): string } }
  }).Buffer
  if (nodeBuffer) return nodeBuffer.from(bytes).toString('base64')
  throw new Error('Base64 encoding is unavailable in this runtime.')
}

export function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const length = parts.reduce((sum, part) => sum + part.length, 0)
  const output = new Uint8Array(length)
  let offset = 0
  for (const part of parts) {
    output.set(part, offset)
    offset += part.length
  }
  return output
}

export function writeUint32Be(target: Uint8Array, offset: number, value: number): void {
  new DataView(target.buffer, target.byteOffset, target.byteLength).setUint32(offset, value >>> 0, false)
}

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle
  if (!subtle) throw new Error('Web Crypto SHA-256 is unavailable in this runtime.')
  const data = new Uint8Array(bytes)
  const digest = new Uint8Array(await subtle.digest('SHA-256', data))
  return [...digest].map((value) => value.toString(16).padStart(2, '0')).join('')
}

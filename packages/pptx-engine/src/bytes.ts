const utf8Encoder = new TextEncoder()
const utf8Decoder = new TextDecoder('utf-8')

export function encodeUtf8(value: string): Uint8Array {
  return utf8Encoder.encode(value)
}

export function decodeUtf8(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes)
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

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle
  if (subtle) {
    const data = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
    const digest = new Uint8Array(await subtle.digest('SHA-256', data))
    return [...digest].map((value) => value.toString(16).padStart(2, '0')).join('')
  }

  const { createHash } = await import('node:crypto')
  return createHash('sha256').update(bytes).digest('hex')
}

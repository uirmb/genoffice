import { afterEach, describe, expect, it, vi } from 'vitest'
import { sha256Hex } from '../src/bytes'

const encoder = new TextEncoder()

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('sha256Hex', () => {
  it('falls back when Web Crypto is unavailable', async () => {
    vi.stubGlobal('crypto', undefined)

    await expect(sha256Hex(encoder.encode('abc'))).resolves.toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    )
  })

  it('hashes empty input without Web Crypto', async () => {
    vi.stubGlobal('crypto', undefined)

    await expect(sha256Hex(new Uint8Array())).resolves.toBe(
      'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    )
  })
})

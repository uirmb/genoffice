import { describe, expect, it } from 'vitest'
import { sha256Fallback } from '../src/web/sha256'

function bytes(value: string): ArrayBuffer {
  return new TextEncoder().encode(value).buffer
}

describe('Docx Web SHA-256 fallback', () => {
  it('matches the SHA-256 empty-string vector', () => {
    expect(sha256Fallback(bytes(''))).toBe(
      'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    )
  })

  it('matches the SHA-256 abc vector', () => {
    expect(sha256Fallback(bytes('abc'))).toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    )
  })

  it('matches the SHA-256 hello-world vector', () => {
    expect(sha256Fallback(bytes('hello world'))).toBe(
      'b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9',
    )
  })
})

import type { OfficeHostApi, SelectedOfficeFile } from '@genoffice/office-host-api'
import {
  localImageResultSchema,
  type LocalImageRequest,
  type LocalImageResult,
} from '../shared/desktop-api'

const MAX_LOCAL_IMAGE_BYTES = 20 * 1024 * 1024
const IMAGE_ACCEPT = [
  'image/png',
  'image/jpeg',
  'image/gif',
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
]

function sniffImageType(bytes: Uint8Array): LocalImageResult['mediaType'] | null {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  ) {
    return 'image/png'
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'image/jpeg'
  }
  if (bytes.length >= 6) {
    const signature = String.fromCharCode(...bytes.subarray(0, 6))
    if (signature === 'GIF87a' || signature === 'GIF89a') return 'image/gif'
  }
  return null
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunkSize = 0x8000
  for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
  }
  return btoa(binary)
}

async function selectedBytes(
  host: OfficeHostApi,
  selected: SelectedOfficeFile,
): Promise<ArrayBuffer> {
  if ((selected.size ?? 0) > MAX_LOCAL_IMAGE_BYTES) {
    throw new Error('Image exceeds 20MB and cannot be inserted.')
  }
  if (selected.transport === 'buffer' && selected.bytes) return selected.bytes
  return (await host.readFile(selected.id)).bytes
}

/**
 * Web cannot dereference an Electron absolute path. It delegates image choice
 * to the Office Host instead: standalone uses the browser picker, while an
 * embedded UC/Web OS host can return a platform file or token through the same
 * office:pick-file / office:read-file contract.
 */
export async function readLocalImageViaHost(
  host: OfficeHostApi,
  _request: LocalImageRequest,
): Promise<LocalImageResult> {
  const selected = await host.pickFile({
    multiple: false,
    accept: IMAGE_ACCEPT,
    mode: 'file',
  })
  if (!selected?.[0]) throw new Error('Image selection was cancelled.')

  const buffer = await selectedBytes(host, selected[0])
  if (buffer.byteLength > MAX_LOCAL_IMAGE_BYTES) {
    throw new Error('Image exceeds 20MB and cannot be inserted.')
  }

  const bytes = new Uint8Array(buffer)
  const mediaType = sniffImageType(bytes)
  if (!mediaType) throw new Error('The selected file is not a PNG/JPEG/GIF image.')

  return localImageResultSchema.parse({
    mediaType,
    base64: bytesToBase64(bytes),
  })
}

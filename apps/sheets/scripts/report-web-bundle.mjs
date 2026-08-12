import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { extname, relative, resolve } from 'node:path'
import { brotliCompressSync, constants as zlibConstants, gzipSync } from 'node:zlib'

const root = resolve(process.argv[2] || 'dist-web')
const reportPath = resolve(process.argv[3] || 'reports/web-bundle.json')
const include = new Set(['.js', '.css'])

async function filesUnder(directory) {
  const result = []
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name)
    if (entry.isDirectory()) result.push(...(await filesUnder(path)))
    else if (entry.isFile() && include.has(extname(entry.name))) result.push(path)
  }
  return result
}

function kib(value) {
  return Math.round((value / 1024) * 10) / 10
}

const files = []
for (const path of await filesUnder(root)) {
  const bytes = await readFile(path)
  const rawBytes = (await stat(path)).size
  const gzipBytes = gzipSync(bytes, { level: 9 }).byteLength
  const brotliBytes = brotliCompressSync(bytes, {
    params: { [zlibConstants.BROTLI_PARAM_QUALITY]: 11 },
  }).byteLength
  files.push({
    path: relative(root, path).replaceAll('\\', '/'),
    type: extname(path).slice(1),
    rawBytes,
    gzipBytes,
    brotliBytes,
  })
}

files.sort((left, right) => right.brotliBytes - left.brotliBytes)
const totals = files.reduce(
  (sum, item) => ({
    rawBytes: sum.rawBytes + item.rawBytes,
    gzipBytes: sum.gzipBytes + item.gzipBytes,
    brotliBytes: sum.brotliBytes + item.brotliBytes,
  }),
  { rawBytes: 0, gzipBytes: 0, brotliBytes: 0 },
)
const largestJavaScript = files.find((item) => item.type === 'js') ?? null
const largestCss = files.find((item) => item.type === 'css') ?? null
const summary = {
  bundleRoot: relative(process.cwd(), root).replaceAll('\\', '/'),
  assetCount: files.length,
  totals,
  totalsKiB: {
    raw: kib(totals.rawBytes),
    gzip: kib(totals.gzipBytes),
    brotli: kib(totals.brotliBytes),
  },
  largestJavaScript,
  largestCss,
  topAssets: files.slice(0, 15),
  assets: files,
}

await mkdir(resolve(reportPath, '..'), { recursive: true })
await writeFile(reportPath, `${JSON.stringify(summary, null, 2)}\n`)

console.log(
  `Sheets Web bundle: ${summary.assetCount} JS/CSS assets, ` +
    `${summary.totalsKiB.raw} KiB raw, ${summary.totalsKiB.gzip} KiB gzip, ` +
    `${summary.totalsKiB.brotli} KiB brotli`,
)
if (largestJavaScript) {
  console.log(
    `Largest JS: ${largestJavaScript.path} ` +
      `(${kib(largestJavaScript.rawBytes)} KiB raw / ` +
      `${kib(largestJavaScript.gzipBytes)} KiB gzip / ` +
      `${kib(largestJavaScript.brotliBytes)} KiB brotli)`,
  )
}
if (largestCss) {
  console.log(
    `Largest CSS: ${largestCss.path} ` +
      `(${kib(largestCss.rawBytes)} KiB raw / ` +
      `${kib(largestCss.gzipBytes)} KiB gzip / ` +
      `${kib(largestCss.brotliBytes)} KiB brotli)`,
  )
}

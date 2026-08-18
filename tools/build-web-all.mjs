import { cpSync, mkdirSync, rmSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const releaseRoot = resolve(repoRoot, 'dist', 'web')
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm'

const apps = [
  {
    label: 'DOCX',
    script: 'build:web:docs',
    base: '/docx/',
    source: 'apps/docs/dist-web',
    target: 'docx',
  },
  {
    label: 'XLSX',
    script: 'build:web:sheets',
    base: '/xlsx/',
    source: 'apps/sheets/dist-web',
    target: 'xlsx',
  },
  {
    label: 'PPTX',
    script: 'build:web:slides',
    base: '/pptx/',
    source: 'apps/slides/dist-web',
    target: 'pptx',
  },
  {
    label: 'Markdown',
    script: 'build:web:markdown',
    base: '/md/',
    source: 'apps/markdown/dist-web',
    target: 'md',
  },
  {
    label: 'PDF',
    script: 'build:web:pdf',
    base: '/pdf/',
    source: 'apps/pdf/dist-web',
    target: 'pdf',
  },
]

rmSync(releaseRoot, { recursive: true, force: true })
mkdirSync(releaseRoot, { recursive: true })

for (const app of apps) {
  console.log(`\n[genoffice-web] Building ${app.label} for ${app.base}`)
  const result = spawnSync(npmCommand, ['run', app.script], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GENOFFICE_WEB_BASE: app.base,
    },
    stdio: 'inherit',
  })

  if (result.error) throw result.error
  if (result.status !== 0) process.exit(result.status ?? 1)

  const source = resolve(repoRoot, app.source)
  const target = resolve(releaseRoot, app.target)
  cpSync(source, target, { recursive: true })
}

console.log(`\n[genoffice-web] Release bundle ready: ${releaseRoot}`)

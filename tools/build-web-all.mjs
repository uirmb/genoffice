import { spawnSync } from 'node:child_process'
import { cpSync, mkdirSync, rmSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const releaseRoot = resolve(repoRoot, 'dist', 'web')

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

function runNpmScript(script, env) {
  const npmExecPath = process.env.npm_execpath

  if (npmExecPath) {
    return spawnSync(process.execPath, [npmExecPath, 'run', script], {
      cwd: repoRoot,
      env,
      stdio: 'inherit',
    })
  }

  return spawnSync('npm', ['run', script], {
    cwd: repoRoot,
    env,
    stdio: 'inherit',
    shell: process.platform === 'win32',
  })
}

rmSync(releaseRoot, { recursive: true, force: true })
mkdirSync(releaseRoot, { recursive: true })

for (const app of apps) {
  console.log(`\n[genoffice-web] Building ${app.label} for ${app.base}`)
  const result = runNpmScript(app.script, {
    ...process.env,
    GENOFFICE_WEB_BASE: app.base,
  })

  if (result.error) throw result.error
  if (result.status !== 0) process.exit(result.status ?? 1)

  const source = resolve(repoRoot, app.source)
  const target = resolve(releaseRoot, app.target)
  cpSync(source, target, { recursive: true })
}

console.log(`\n[genoffice-web] Release bundle ready: ${releaseRoot}`)

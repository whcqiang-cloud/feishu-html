import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
import packageJson from '../package.json' with { type: 'json' }

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const distDir = path.resolve(__dirname, '../dist')
const releaseDir = path.resolve(__dirname, '../release')

const main = async () => {
  await fs.mkdir(releaseDir, { recursive: true })

  const version = packageJson.version
  const zipName = `cloud-document-html-${version}.zip`
  const zipPath = path.join(releaseDir, zipName)

  console.log(`Packaging release: ${zipName}`)

  await fs.rm(zipPath, { force: true })

  const result = spawnSync(
    'zip',
    ['-r', '-q', zipPath, '.', '-x', '.DS_Store', '*/.DS_Store'],
    {
      cwd: distDir,
      stdio: 'inherit',
    },
  )

  if (result.status !== 0) {
    const errorMessage = result.error ? result.error.message : 'unknown error'
    console.error(`Failed to create ZIP: ${errorMessage}`)
    process.exit(1)
  }

  console.log(`Release ready: ${zipPath}`)
}

await main()

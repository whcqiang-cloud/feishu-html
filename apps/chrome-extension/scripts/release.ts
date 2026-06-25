import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import zlib from 'node:zlib'
import packageJson from '../package.json' with { type: 'json' }

interface FileEntry {
  name: string
  fullPath: string
}

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const distDir = path.resolve(__dirname, '../dist')
const releaseDir = path.resolve(__dirname, '../release')

const crcTable = new Array<number>(256)
for (let n = 0; n < 256; n++) {
  let c = n
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  }
  crcTable[n] = c
}

const computeCrc32 = (buf: Buffer) => {
  let crc = 0xffffffff
  for (const byte of buf) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}

const writeUint32 = (buf: Buffer, offset: number, value: number) => {
  buf.writeUInt32LE(value >>> 0, offset)
}

const writeUint16 = (buf: Buffer, offset: number, value: number) => {
  buf.writeUInt16LE(value >>> 0, offset)
}

const createLocalFileHeader = (fileName: string, data: Buffer) => {
  const nameBuf = Buffer.from(fileName, 'utf8')
  const header = Buffer.alloc(30 + nameBuf.length)
  writeUint32(header, 0, 0x04034b50)
  writeUint16(header, 4, 20)
  writeUint16(header, 6, 0)
  writeUint16(header, 8, 0) // stored
  writeUint16(header, 10, 0)
  writeUint16(header, 12, 0)
  writeUint32(header, 14, computeCrc32(data))
  writeUint32(header, 18, data.length)
  writeUint32(header, 22, data.length)
  writeUint16(header, 26, nameBuf.length)
  writeUint16(header, 28, 0)
  nameBuf.copy(header, 30)
  return header
}

const createCentralDirectoryEntry = (
  fileName: string,
  data: Buffer,
  localHeaderOffset: number,
) => {
  const nameBuf = Buffer.from(fileName, 'utf8')
  const header = Buffer.alloc(46 + nameBuf.length)
  writeUint32(header, 0, 0x02014b50)
  writeUint16(header, 4, 20)
  writeUint16(header, 6, 20)
  writeUint16(header, 8, 0)
  writeUint16(header, 10, 0)
  writeUint16(header, 12, 0)
  writeUint16(header, 14, 0)
  writeUint32(header, 16, computeCrc32(data))
  writeUint32(header, 20, data.length)
  writeUint32(header, 24, data.length)
  writeUint16(header, 28, nameBuf.length)
  writeUint16(header, 30, 0)
  writeUint16(header, 32, 0)
  writeUint16(header, 34, 0)
  writeUint16(header, 36, 0)
  writeUint32(header, 38, 0)
  writeUint32(header, 42, localHeaderOffset)
  nameBuf.copy(header, 46)
  return header
}

const createEndOfCentralDirectory = (
  cdSize: number,
  cdOffset: number,
  entryCount: number,
) => {
  const header = Buffer.alloc(22)
  writeUint32(header, 0, 0x06054b50)
  writeUint16(header, 4, 0)
  writeUint16(header, 6, 0)
  writeUint16(header, 8, entryCount)
  writeUint16(header, 10, entryCount)
  writeUint32(header, 12, cdSize)
  writeUint32(header, 16, cdOffset)
  writeUint16(header, 20, 0)
  return header
}

const zipData = (data: Buffer) => {
  return zlib.deflateSync(data)
}

const collectFiles = async (dir: string, prefix = ''): Promise<FileEntry[]> => {
  const results: FileEntry[] = []
  const entries = await fs.readdir(dir, { withFileTypes: true })
  for (const entry of entries) {
    if (entry.name.startsWith('.')) continue
    const fullPath = path.join(dir, entry.name)
    const relativeName = prefix ? `${prefix}/${entry.name}` : entry.name
    if (entry.isDirectory()) {
      results.push(...(await collectFiles(fullPath, relativeName)))
    } else {
      results.push({ name: relativeName, fullPath })
    }
  }
  return results
}

const createZip = async (zipPath: string) => {
  const files = await collectFiles(distDir)
  const allBuffers: Buffer[] = []
  const localHeaderOffsets: number[] = []

  // First pass: write local file headers + compressed data
  for (const file of files) {
    const data = await fs.readFile(file.fullPath)
    const compressed = zipData(data)
    const header = createLocalFileHeader(file.name, compressed)
    localHeaderOffsets.push(allBuffers.length)
    allBuffers.push(header, compressed)
  }

  // Second pass: write central directory
  const cdStartOffset = allBuffers.reduce((sum, b) => sum + b.length, 0)
  const cdEntries: Buffer[] = []
  for (let i = 0; i < files.length; i++) {
    const file = files[i]
    const data = await fs.readFile(file.fullPath)
    const compressed = zipData(data)
    const entry = createCentralDirectoryEntry(
      file.name,
      compressed,
      localHeaderOffsets[i],
    )
    cdEntries.push(entry)
  }

  allBuffers.push(...cdEntries)

  // End of central directory
  const cdSize = cdEntries.reduce((sum, b) => sum + b.length, 0)
  allBuffers.push(
    createEndOfCentralDirectory(cdSize, cdStartOffset, files.length),
  )

  const zipBuffer = Buffer.concat(allBuffers)
  await fs.writeFile(zipPath, zipBuffer)
  console.log(`Created: ${zipPath}`)
}

const main = async () => {
  await fs.mkdir(releaseDir, { recursive: true })

  const version = packageJson.version
  const zipName = `cloud-document-html-${version}.zip`
  const zipPath = path.join(releaseDir, zipName)

  console.log(`Packaging release: ${zipName}`)
  await createZip(zipPath)

  console.log(`Release ready: ${zipPath}`)
}

await main()

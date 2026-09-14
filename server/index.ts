import 'dotenv/config'
import { ZipArchive } from 'archiver'
import cors from 'cors'
import express from 'express'
import mongoose, { Schema } from 'mongoose'
import multer from 'multer'
import PDFDocument from 'pdfkit'
import * as XLSX from 'xlsx'

const app = express()
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } })
app.use(cors())
app.use(express.json({ limit: '20mb' }))

type TemplatePayload = {
  name: string; title: string; type: string; primaryField: string; body: string
  signatory: string; organization: string; accent: string; logo?: string
  signatureImage?: string; certNumberPrefix?: string
}

const templateSchema = new Schema<TemplatePayload>({
  name: { type: String, required: true }, title: { type: String, required: true },
  type: { type: String, required: true }, primaryField: { type: String, required: true },
  body: { type: String, required: true }, signatory: { type: String, required: true },
  organization: { type: String, required: true }, accent: { type: String, required: true },
  logo: { type: String, default: '' },
  signatureImage: { type: String, default: '' },
  certNumberPrefix: { type: String, default: 'CERT' },
}, { timestamps: true })
const CertificateTemplate = mongoose.model<TemplatePayload>('CertificateTemplate', templateSchema)

let databaseState: 'connecting' | 'connected' | 'offline' = 'connecting'
let cachedPromise: Promise<typeof mongoose> | null = null

async function connectToDatabase() {
  const uri = process.env.MONGODB_URI
  if (!uri) {
    databaseState = 'offline'
    return null
  }
  if (mongoose.connection.readyState >= 1) {
    databaseState = 'connected'
    return mongoose
  }
  if (!cachedPromise) {
    cachedPromise = mongoose.connect(uri, { dbName: 'certificate_generator', serverSelectionTimeoutMS: 8000 })
      .then((m) => {
        databaseState = 'connected'
        return m
      })
      .catch((err) => {
        databaseState = 'offline'
        cachedPromise = null
        throw err
      })
  }
  return cachedPromise
}

// Initial connection attempt
connectToDatabase().catch(() => {})

const value = (row: Record<string, unknown>, key: string) => String(row[key] ?? '')
const interpolate = (source: string, row: Record<string, unknown>) =>
  source.replace(/{{\s*([^}]+)\s*}}/g, (_, key: string) => value(row, key.trim()) || `{{${key.trim()}}}`)

app.use(async (_req, _res, next) => {
  try {
    await connectToDatabase()
  } catch {
    // handled by route handler checking databaseState
  }
  next()
})

app.get('/api/health', (_req, res) => res.json({ database: databaseState }))
app.get('/api/templates', async (_req, res) => {
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  res.json(await CertificateTemplate.find().sort({ updatedAt: -1 }).lean())
})
app.post('/api/templates', async (req, res) => {
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  const payload = req.body as TemplatePayload
  if (!payload.name || !payload.primaryField || !payload.body) return res.status(400).json({ message: 'Nama, kolom utama, dan isi sertifikat wajib diisi.' })
  res.status(201).json(await CertificateTemplate.create(payload))
})
app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'File Excel belum dipilih.' })
  try {
    const book = XLSX.read(req.file.buffer, { type: 'buffer', raw: false })
    const sheet = book.Sheets[book.SheetNames[0] ?? '']
    if (!sheet) throw new Error('Sheet tidak ditemukan')
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' })
    res.json({ headers: rows.length ? Object.keys(rows[0] ?? {}) : [], rows })
  } catch {
    res.status(400).json({ message: 'File tidak dapat dibaca. Gunakan .xlsx atau .xls dengan header pada baris pertama.' })
  }
})
const getCertificateNumber = (template: TemplatePayload, recipient: Record<string, unknown>) => {
  // If Excel has explicit 'nomor', 'no_sertifikat', 'certificate_number', etc.
  for (const key of ['nomor', 'no_sertifikat', 'nomor_sertifikat', 'certificate_number', 'no']) {
    if (recipient[key]) return String(recipient[key]).trim()
  }
  const prefix = (template.certNumberPrefix || 'CERT').trim().toUpperCase()
  const rawId = `${value(recipient, template.primaryField)}_${template.title}_${template.signatory}`
  let hash = 0
  for (let i = 0; i < rawId.length; i++) {
    hash = ((hash << 5) - hash) + rawId.charCodeAt(i)
    hash |= 0
  }
  const positive = Math.abs(hash)
  const hexPart = positive.toString(16).toUpperCase().padStart(6, '0').slice(-6)
  const numPart = String(positive % 9000 + 1000)
  const year = new Date().getFullYear()
  return `${prefix}/${year}/${numPart}-${hexPart}`
}

function renderCertificatePage(doc: PDFKit.PDFDocument, template: TemplatePayload, recipient: Record<string, unknown>) {
  const name = value(recipient, template.primaryField) || 'Penerima'
  const accentColor = template.accent || '#0f766e'
  const certNumber = getCertificateNumber(template, recipient)

  // Subtle background tone
  doc.rect(0, 0, 842, 595).fill('#fcfbf7')

  // Outer primary accent frame
  doc.rect(20, 20, 802, 555).lineWidth(4).stroke(accentColor)
  // Inner thin border
  doc.rect(28, 28, 786, 539).lineWidth(1).stroke('#d7c28a')

  // Modern ornamental geometric corner brackets
  const drawCorner = (x: number, y: number, dx: number, dy: number) => {
    doc.save()
    doc.lineWidth(2).strokeColor(accentColor)
    doc.moveTo(x, y + dy * 22).lineTo(x, y).lineTo(x + dx * 22, y).stroke()
    doc.circle(x + dx * 8, y + dy * 8, 2.5).fillColor('#d7c28a').fill()
    doc.restore()
  }
  drawCorner(34, 34, 1, 1)
  drawCorner(808, 34, -1, 1)
  drawCorner(34, 561, 1, -1)
  drawCorner(808, 561, -1, -1)

  // Header Layout: Centered prominent logo or emblem badge
  let contentStartY = 156
  let logoRendered = false
  if (template.logo && template.logo.includes('base64,')) {
    try {
      const base64 = template.logo.substring(template.logo.indexOf('base64,') + 7)
      if (base64) {
        const logoBuffer = Buffer.from(base64, 'base64')
        const logoMaxW = 120
        const logoMaxH = 80
        const logoX = (842 - logoMaxW) / 2
        doc.image(logoBuffer, logoX, 36, { fit: [logoMaxW, logoMaxH], align: 'center', valign: 'center' })
        contentStartY = 126
        logoRendered = true
      }
    } catch (err) {
      console.error('Failed to render logo in PDF:', err)
    }
  }

  if (!logoRendered) {
    doc.save()
    doc.circle(421, 80, 22).lineWidth(1.5).strokeColor('#d7c28a').stroke()
    doc.circle(421, 80, 18).lineWidth(1).strokeColor(accentColor).stroke()
    doc.fillColor(accentColor).fontSize(16).font('Helvetica-Bold').text('certify', 400, 73)
    doc.restore()
    contentStartY = 118
  }

  // Organization Header
  doc.fillColor(accentColor)
    .fontSize(12)
    .font('Helvetica-Bold')
    .text(template.organization.toUpperCase(), 60, contentStartY, { align: 'center', characterSpacing: 2 })

  // Certificate Main Title
  doc.fillColor('#172327')
    .fontSize(35)
    .font('Times-Bold')
    .text(template.title, 60, contentStartY + 23, { align: 'center' })

  // Badge pill type
  const typeText = `SERTIFIKAT ${template.type.toUpperCase()}`
  doc.fillColor('#606f7b')
    .fontSize(10)
    .font('Helvetica-Bold')
    .text(typeText, 60, contentStartY + 72, { align: 'center', characterSpacing: 2.5 })

  // Awarded text
  doc.fillColor('#4b5563')
    .fontSize(12)
    .font('Helvetica')
    .text('Diberikan dengan penuh kehormatan kepada', 60, contentStartY + 104, { align: 'center' })

  // Recipient Name
  doc.fillColor(accentColor)
    .fontSize(34)
    .font('Times-BoldItalic')
    .text(name, 50, contentStartY + 128, { align: 'center' })

  // Elegant divider under name with center jewel
  const dividerY = contentStartY + 172
  doc.moveTo(220, dividerY).lineTo(390, dividerY).lineWidth(1).strokeColor('#d7c28a').stroke()
  doc.polygon([421, dividerY - 4], [426, dividerY], [421, dividerY + 4], [416, dividerY]).fillColor(accentColor).fill()
  doc.moveTo(452, dividerY).lineTo(622, dividerY).lineWidth(1).strokeColor('#d7c28a').stroke()

  // Body Description
  doc.fillColor('#334155')
    .fontSize(12.5)
    .font('Helvetica')
    .text(interpolate(template.body, recipient), 95, contentStartY + 188, {
      align: 'center',
      width: 652,
      lineGap: 5
    })

  // Bottom Footer: Official Seal (left), Unique Cert Number (center bottom), and Signatory (right)
  doc.save()
  const sealX = 145
  const sealY = 502
  doc.circle(sealX, sealY, 28).lineWidth(1.5).strokeColor('#d7c28a').stroke()
  doc.circle(sealX, sealY, 24).lineWidth(1).strokeColor(accentColor).stroke()
  doc.fillColor(accentColor).fontSize(7).font('Helvetica-Bold').text('AUTHENTIC', sealX - 22, sealY - 10, { width: 44, align: 'center', characterSpacing: 1 })
  doc.fillColor('#d7c28a').fontSize(11).text('★', sealX - 4, sealY - 2)
  doc.fillColor(accentColor).fontSize(7).text('VERIFIED', sealX - 22, sealY + 8, { width: 44, align: 'center', characterSpacing: 1 })
  doc.restore()

  // Center bottom: Unique Certificate Number badge
  doc.save()
  doc.fillColor('#64748b')
    .fontSize(8.5)
    .font('Helvetica-Bold')
    .text('NO. SERTIFIKAT', 280, 498, { width: 282, align: 'center', characterSpacing: 1.5 })
  doc.fillColor('#1e293b')
    .fontSize(10)
    .font('Helvetica')
    .text(certNumber, 280, 514, { width: 282, align: 'center', characterSpacing: 1.2 })
  doc.restore()

  // Right: Signature block
  doc.fillColor('#64748b').fontSize(9.5).font('Helvetica').text('Ditetapkan secara resmi oleh', 525, 452, { width: 220, align: 'center' })

  // Render PNG Signature if uploaded
  if (template.signatureImage && template.signatureImage.includes('base64,')) {
    try {
      const sigBase64 = template.signatureImage.substring(template.signatureImage.indexOf('base64,') + 7)
      if (sigBase64) {
        const sigBuffer = Buffer.from(sigBase64, 'base64')
        // Center the signature image in the signature block (x=525, width=220 → center=635)
        doc.image(sigBuffer, 565, 458, { fit: [140, 52], align: 'center', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render signature in PDF:', err)
    }
  }

  doc.fillColor('#172327').font('Times-Bold').fontSize(18).text(template.signatory, 525, 504, { width: 220, align: 'center' })
  doc.moveTo(555, 526).lineTo(715, 526).lineWidth(1).strokeColor('#cbd5e1').stroke()
  doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9.5).text(template.organization, 525, 532, { width: 220, align: 'center' })
}

function generateSinglePDFBuffer(template: TemplatePayload, recipient: Record<string, unknown>): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 42 })
    const chunks: Buffer[] = []
    doc.on('data', (chunk) => chunks.push(chunk))
    doc.on('end', () => resolve(Buffer.concat(chunks)))
    doc.on('error', reject)
    renderCertificatePage(doc, template, recipient)
    doc.end()
  })
}

// Single Certificate PDF Download
app.post('/api/certificates/pdf', (req, res) => {
  const { template, recipient } = req.body as { template: TemplatePayload; recipient: Record<string, unknown> }
  if (!template || !recipient) return res.status(400).json({ message: 'Template dan penerima wajib ada.' })
  const name = value(recipient, template.primaryField) || 'Penerima'
  const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 42 })
  const filename = `sertifikat-${name.toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
  res.setHeader('Content-Type', 'application/pdf')
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
  doc.pipe(res)
  renderCertificatePage(doc, template, recipient)
  doc.end()
})

// Batch PDF Download: All certificates combined into one multi-page PDF
app.post('/api/certificates/batch-pdf', (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 42, autoFirstPage: false })
  const filename = `semua-sertifikat-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
  res.setHeader('Content-Type', 'application/pdf')
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
  doc.pipe(res)

  for (const recipient of recipients) {
    doc.addPage({ size: 'A4', layout: 'landscape', margin: 42 })
    renderCertificatePage(doc, template, recipient)
  }
  doc.end()
})

// Batch ZIP Download: Individual PDF for each recipient packaged in a ZIP
app.post('/api/certificates/zip', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }

  const archive = new ZipArchive({ zlib: { level: 6 } })
  archive.pipe(res)

  archive.on('error', (err: Error) => {
    console.error('Archive error:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file ZIP' })
  })

  const usedFilenames = new Map<string, number>()
  for (let i = 0; i < recipients.length; i++) {
    const r = recipients[i]!
    const recipientName = value(r, template.primaryField) || `Penerima-${i + 1}`
    const safeName = recipientName.toLowerCase().replace(/[^a-z0-9]+/gi, '-')
    const count = (usedFilenames.get(safeName) || 0) + 1
    usedFilenames.set(safeName, count)
    const entryName = count > 1
      ? `sertifikat-${safeName}-${count}.pdf`
      : `sertifikat-${safeName}.pdf`

    const pdfBuffer = await generateSinglePDFBuffer(template, r)
    archive.append(pdfBuffer, { name: entryName })
  }

  await archive.finalize()
})

if (!process.env.VERCEL) {
  app.listen(Number(process.env.PORT || 3001), () => console.log('Certificate API running on port 3001'))
}

export default app

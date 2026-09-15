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
  name: string; title: string; subtitle?: string; type: string; primaryField: string; body: string
  signatory: string; organization: string; accent: string; logo?: string
  logo2?: string
  signatureImage?: string; signatureImage2?: string
  signatory2?: string; signatoryTitle1?: string; signatoryTitle2?: string
  certNumberPrefix?: string
}

const templateSchema = new Schema<TemplatePayload>({
  name: { type: String, required: true }, title: { type: String, required: true },
  subtitle: { type: String, default: '' },
  type: { type: String, required: true }, primaryField: { type: String, required: true },
  body: { type: String, required: true }, signatory: { type: String, required: true },
  signatory2: { type: String, default: '' },
  signatoryTitle1: { type: String, default: '' },
  signatoryTitle2: { type: String, default: '' },
  organization: { type: String, required: true }, accent: { type: String, required: true },
  logo: { type: String, default: '' },
  logo2: { type: String, default: '' },
  signatureImage: { type: String, default: '' },
  signatureImage2: { type: String, default: '' },
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
    cachedPromise = mongoose.connect(uri, { dbName: 'certificate_generator', serverSelectionTimeoutMS: 5000 })
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

connectToDatabase().catch(() => {})

const value = (row: Record<string, unknown>, key: string) => String(row[key] ?? '')
const interpolate = (source: string, row: Record<string, unknown>) =>
  source.replace(/{{\s*([^}]+)\s*}}/g, (_, key: string) => value(row, key.trim()) || `{{${key.trim()}}}`)

app.get('/api/health', async (_req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  res.json({ database: databaseState })
})

app.get('/api/templates', async (_req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  res.json(await CertificateTemplate.find().sort({ updatedAt: -1 }).lean())
})

app.post('/api/templates', async (req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
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

  // Header Layout: Top Left Logo 2 and Top Right Logo 1
  const logoMaxW = 100
  const logoMaxH = 65
  const logoY = 40

  // Logo 2 (Kiri Atas)
  if (template.logo2 && template.logo2.includes('base64,')) {
    try {
      const base64 = template.logo2.substring(template.logo2.indexOf('base64,') + 7)
      if (base64) {
        const logo2Buffer = Buffer.from(base64, 'base64')
        doc.image(logo2Buffer, 50, logoY, { fit: [logoMaxW, logoMaxH], valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo2 in PDF:', err)
    }
  }

  // Logo 1 (Kanan Atas)
  if (template.logo && template.logo.includes('base64,')) {
    try {
      const base64 = template.logo.substring(template.logo.indexOf('base64,') + 7)
      if (base64) {
        const logo1Buffer = Buffer.from(base64, 'base64')
        doc.image(logo1Buffer, 842 - 50 - logoMaxW, logoY, { fit: [logoMaxW, logoMaxH], align: 'right', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo1 in PDF:', err)
    }
  }

  // If neither logo is provided, show elegant center emblem
  if (!template.logo && !template.logo2) {
    doc.save()
    doc.circle(421, 62, 18).lineWidth(1.5).strokeColor('#d7c28a').stroke()
    doc.circle(421, 62, 14).lineWidth(1).strokeColor(accentColor).stroke()
    doc.fillColor(accentColor).fontSize(14).font('Helvetica-Bold').text('✦', 415, 55)
    doc.restore()
  }

  const contentStartY = 64

  // Organization Header
  doc.fillColor(accentColor)
    .fontSize(12)
    .font('Helvetica-Bold')
    .text(template.organization.toUpperCase(), 160, contentStartY, { align: 'center', width: 522, characterSpacing: 2 })

  // Certificate Main Title
  doc.fillColor('#172327')
    .fontSize(32)
    .font('Times-Bold')
    .text(template.title, 60, contentStartY + 18, { align: 'center' })

  let currentY = contentStartY + 54

  // Certificate Subtitle (jika ada)
  if (template.subtitle && template.subtitle.trim()) {
    doc.fillColor('#4b5563')
      .fontSize(12)
      .font('Helvetica-Oblique')
      .text(template.subtitle.trim(), 60, currentY, { align: 'center' })
    currentY += 20
  }

  // Badge pill type
  const typeText = `SERTIFIKAT ${template.type.toUpperCase()}`
  doc.fillColor('#606f7b')
    .fontSize(10)
    .font('Helvetica-Bold')
    .text(typeText, 60, currentY, { align: 'center', characterSpacing: 2.5 })

  // Awarded text
  doc.fillColor('#4b5563')
    .fontSize(11.5)
    .font('Helvetica')
    .text('Diberikan dengan penuh kehormatan kepada', 60, currentY + 24, { align: 'center' })

  // Recipient Name
  doc.fillColor(accentColor)
    .fontSize(33)
    .font('Times-BoldItalic')
    .text(name, 50, currentY + 44, { align: 'center' })

  // Elegant divider under name with center jewel
  const dividerY = currentY + 86
  doc.moveTo(220, dividerY).lineTo(390, dividerY).lineWidth(1).strokeColor('#d7c28a').stroke()
  doc.polygon([421, dividerY - 4], [426, dividerY], [421, dividerY + 4], [416, dividerY]).fillColor(accentColor).fill()
  doc.moveTo(452, dividerY).lineTo(622, dividerY).lineWidth(1).strokeColor('#d7c28a').stroke()

  // Body Description
  doc.fillColor('#334155')
    .fontSize(12)
    .font('Helvetica')
    .text(interpolate(template.body, recipient), 95, currentY + 98, {
      align: 'center',
      width: 652,
      lineGap: 4
    })

  // Center bottom: Official Seal and Unique Certificate Number badge
  doc.save()
  const sealX = 421
  const sealY = 468
  doc.circle(sealX, sealY, 24).lineWidth(1.5).strokeColor('#d7c28a').stroke()
  doc.circle(sealX, sealY, 20).lineWidth(1).strokeColor(accentColor).stroke()
  doc.fillColor(accentColor).fontSize(6.5).font('Helvetica-Bold').text('AUTHENTIC', sealX - 22, sealY - 9, { width: 44, align: 'center', characterSpacing: 1 })
  doc.fillColor('#d7c28a').fontSize(9).text('★', sealX - 4, sealY - 2)
  doc.fillColor(accentColor).fontSize(6.5).text('VERIFIED', sealX - 22, sealY + 6, { width: 44, align: 'center', characterSpacing: 1 })
  doc.restore()

  doc.save()
  doc.fillColor('#64748b')
    .fontSize(8)
    .font('Helvetica-Bold')
    .text('NO. SERTIFIKAT', 280, 506, { width: 282, align: 'center', characterSpacing: 1.5 })
  doc.fillColor('#1e293b')
    .fontSize(9.5)
    .font('Helvetica')
    .text(certNumber, 280, 519, { width: 282, align: 'center', characterSpacing: 1.2 })
  doc.restore()

  // Left: Tanda Tangan 2 (TTD 2 kiri bawah)
  const sig2Label = template.signatoryTitle2 || 'Mengetahui'
  const sig2Name = template.signatory2 || template.organization
  doc.fillColor('#64748b').fontSize(9).font('Helvetica').text(sig2Label, 55, 436, { width: 210, align: 'center' })

  if (template.signatureImage2 && template.signatureImage2.includes('base64,')) {
    try {
      const sig2Base64 = template.signatureImage2.substring(template.signatureImage2.indexOf('base64,') + 7)
      if (sig2Base64) {
        const sig2Buffer = Buffer.from(sig2Base64, 'base64')
        doc.image(sig2Buffer, 90, 448, { fit: [140, 50], align: 'center', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render signature 2 in PDF:', err)
    }
  }

  doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig2Name, 55, 502, { width: 210, align: 'center' })
  doc.moveTo(75, 523).lineTo(245, 523).lineWidth(1).strokeColor('#cbd5e1').stroke()
  doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9).text(template.organization, 55, 528, { width: 210, align: 'center' })

  // Right: Tanda Tangan 1 (TTD 1 kanan bawah)
  const sig1Label = template.signatoryTitle1 || 'Ditetapkan secara resmi oleh'
  const sig1Name = template.signatory || 'Penandatangan'
  doc.fillColor('#64748b').fontSize(9).font('Helvetica').text(sig1Label, 577, 436, { width: 210, align: 'center' })

  if (template.signatureImage && template.signatureImage.includes('base64,')) {
    try {
      const sigBase64 = template.signatureImage.substring(template.signatureImage.indexOf('base64,') + 7)
      if (sigBase64) {
        const sigBuffer = Buffer.from(sigBase64, 'base64')
        doc.image(sigBuffer, 612, 448, { fit: [140, 50], align: 'center', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render signature 1 in PDF:', err)
    }
  }

  doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig1Name, 577, 502, { width: 210, align: 'center' })
  doc.moveTo(597, 523).lineTo(767, 523).lineWidth(1).strokeColor('#cbd5e1').stroke()
  doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9).text(template.organization, 577, 528, { width: 210, align: 'center' })
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
app.post('/api/certificates/pdf', async (req, res) => {
  const { template, recipient } = req.body as { template: TemplatePayload; recipient: Record<string, unknown> }
  if (!template || !recipient) return res.status(400).json({ message: 'Template dan penerima wajib ada.' })
  try {
    const name = value(recipient, template.primaryField) || 'Penerima'
    const pdfBuffer = await generateSinglePDFBuffer(template, recipient)
    const filename = `sertifikat-${name.toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.setHeader('Content-Length', pdfBuffer.length)
    res.end(pdfBuffer)
  } catch (err) {
    console.error('Error generating PDF:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file PDF' })
  }
})

// Batch PDF Download: All certificates combined into one multi-page PDF
app.post('/api/certificates/batch-pdf', (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  try {
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
  } catch (err) {
    console.error('Error generating batch PDF:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file batch PDF' })
  }
})

// Batch ZIP Download: Individual PDF for each recipient packaged in a ZIP
app.post('/api/certificates/zip', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }

  try {
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
  } catch (err) {
    console.error('Error generating ZIP:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file ZIP' })
  }
})

export default app

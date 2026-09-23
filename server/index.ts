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
  enableLogo2?: boolean
  enableSig2?: boolean
  signatureImage?: string; signatureImage2?: string
  signatory2?: string; signatoryTitle1?: string; signatoryTitle2?: string
  certNumberPrefix?: string
  certificateDate?: string
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
  enableLogo2: { type: Boolean, default: false },
  enableSig2: { type: Boolean, default: false },
  signatureImage: { type: String, default: '' },
  signatureImage2: { type: String, default: '' },
  certNumberPrefix: { type: String, default: 'CERT' },
  certificateDate: { type: String, default: '' },
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
const interpolate = (source: string, row: Record<string, unknown>, template?: TemplatePayload) =>
  source.replace(/{{\s*([^}]+)\s*}}/g, (_, key: string) => {
    const trimmed = key.trim()
    const val = value(row, trimmed)
    if (val) return val
    if ((trimmed.toLowerCase() === 'tanggal' || trimmed.toLowerCase() === 'date') && template?.certificateDate) {
      return template.certificateDate
    }
    return `{{${trimmed}}}`
  })

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

app.delete('/api/templates/:id', async (req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  try {
    const deleted = await CertificateTemplate.findByIdAndDelete(req.params.id)
    if (!deleted) return res.status(404).json({ message: 'Template tidak ditemukan.' })
    res.json({ message: 'Template berhasil dihapus.' })
  } catch {
    res.status(400).json({ message: 'ID template tidak valid.' })
  }
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

function scallopSealPath(cx: number, cy: number, baseR: number, lobes = 16, bump = 2.5): string {
  const steps = lobes * 2
  const parts: string[] = []
  for (let i = 0; i <= steps; i++) {
    const angle = (i / steps) * Math.PI * 2 - Math.PI / 2
    const radius = i % 2 === 0 ? baseR + bump : baseR - bump * 0.35
    const x = cx + radius * Math.cos(angle)
    const y = cy + radius * Math.sin(angle)
    parts.push(i === 0 ? `M ${x.toFixed(2)} ${y.toFixed(2)}` : `L ${x.toFixed(2)} ${y.toFixed(2)}`)
  }
  return `${parts.join(' ')} Z`
}

function drawSealArcText(
  doc: PDFKit.PDFDocument,
  cx: number,
  cy: number,
  radius: number,
  text: string,
  color: string,
  position: 'top' | 'bottom'
) {
  const chars = text.split('')
  const arcSpan = Math.PI * 0.78
  const startAngle = position === 'top' ? -Math.PI / 2 - arcSpan / 2 : Math.PI / 2 - arcSpan / 2
  const step = chars.length > 1 ? arcSpan / (chars.length - 1) : 0

  doc.font('Helvetica-Bold').fontSize(5).fillColor(color)

  chars.forEach((char, index) => {
    const angle = startAngle + step * index
    const x = cx + radius * Math.cos(angle)
    const y = cy + radius * Math.sin(angle)
    const rotation = (angle * 180) / Math.PI + (position === 'top' ? 90 : -90)

    doc.save()
    doc.translate(x, y)
    doc.rotate(rotation)
    doc.text(char, -1.5, -2, { lineBreak: false, width: 4 })
    doc.restore()
  })
}

function drawOfficialSeal(doc: PDFKit.PDFDocument, cx: number, cy: number, accentColor: string, size: 'md' | 'lg' = 'md') {
  const outerR = size === 'lg' ? 30 : 27
  const bump = size === 'lg' ? 2.8 : 2.5
  const scallop = scallopSealPath(cx, cy, outerR, 16, bump)

  doc.save()

  doc.path(scallop).fillOpacity(0.14).fillColor('#d7c28a').fill()
  doc.path(scallop).fillOpacity(1).lineWidth(1.1).strokeColor('#d7c28a').stroke()

  doc.circle(cx, cy, outerR - 5).fillOpacity(0.95).fillColor('#ffffff').fill()
  doc.circle(cx, cy, outerR - 7).lineWidth(0.8).strokeColor('#d7c28a').stroke()
  doc.circle(cx, cy, outerR - 10).lineWidth(0.7).strokeColor(accentColor).stroke()

  const dotR = outerR - 2.5
  for (const [dx, dy] of [[0, -1], [1, 0], [0, 1], [-1, 0]] as const) {
    doc.circle(cx + dx * dotR * 0.82, cy + dy * dotR * 0.82, 1.1).fillColor('#d7c28a').fill()
  }

  const starOuter = outerR - 14
  const starInner = outerR - 19
  const starPoints: [number, number][] = []
  for (let i = 0; i < 8; i++) {
    const outerAngle = (i / 8) * Math.PI * 2 - Math.PI / 2
    const innerAngle = outerAngle + Math.PI / 8
    starPoints.push([cx + starOuter * Math.cos(outerAngle), cy + starOuter * Math.sin(outerAngle)])
    starPoints.push([cx + starInner * Math.cos(innerAngle), cy + starInner * Math.sin(innerAngle)])
  }
  doc.polygon(...starPoints).fillColor(accentColor).fillOpacity(0.85).fill()

  drawSealArcText(doc, cx, cy, outerR - 11.5, 'AUTHENTIC', accentColor, 'top')
  drawSealArcText(doc, cx, cy, outerR - 11.5, 'VERIFIED', accentColor, 'bottom')

  doc.restore()
}

function renderCertificatePage(doc: PDFKit.PDFDocument, template: TemplatePayload, recipient: Record<string, unknown>) {
  const name = value(recipient, template.primaryField) || 'Penerima'
  const accentColor = template.accent || '#0f766e'
  const certNumber = getCertificateNumber(template, recipient)
  const pageW = 842
  const pageH = 595

  doc.rect(0, 0, pageW, pageH).fill('#ffffff')

  doc.rect(20, 20, 802, 555).lineWidth(4).stroke(accentColor)
  doc.rect(28, 28, 786, 539).lineWidth(1).stroke('#d7c28a')

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

  const logoMaxW = 120
  const logoMaxH = 72
  const logoY = 48
  const isLogo2Active = Boolean(template.enableLogo2)

  if (isLogo2Active && template.logo2 && template.logo2.includes('base64,')) {
    try {
      const base64 = template.logo2.substring(template.logo2.indexOf('base64,') + 7)
      if (base64) {
        const logo2Buffer = Buffer.from(base64, 'base64')
        doc.image(logo2Buffer, 48, logoY, { fit: [logoMaxW, logoMaxH], valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo2 in PDF:', err)
    }
  }

  if (template.logo && template.logo.includes('base64,')) {
    try {
      const base64 = template.logo.substring(template.logo.indexOf('base64,') + 7)
      if (base64) {
        const logo1Buffer = Buffer.from(base64, 'base64')
        doc.image(logo1Buffer, pageW - 48 - logoMaxW, logoY, { fit: [logoMaxW, logoMaxH], align: 'right', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo1 in PDF:', err)
    }
  }

  if (!template.logo && (!isLogo2Active || !template.logo2)) {
    doc.save()
    doc.circle(421, 68, 18).lineWidth(1.5).strokeColor('#d7c28a').stroke()
    doc.circle(421, 68, 14).lineWidth(1).strokeColor(accentColor).stroke()
    doc.fillColor(accentColor).fontSize(15).font('Helvetica-Bold').text('✦', 415, 61)
    doc.restore()
  }

  const bodyText = interpolate(template.body, recipient, template)
  const bodyFontSize = 13
  const bodyWidth = 662
  const bodyLineGap = 6

  doc.font('Helvetica').fontSize(bodyFontSize)
  const bodyHeight = doc.heightOfString(bodyText, {
    width: bodyWidth,
    align: 'center',
    lineGap: bodyLineGap
  })

  const footerReserveH = 150
  const minTopMargin = 118
  const maxContentBottomY = pageH - 40 - footerReserveH

  let currentY = minTopMargin

  doc.fillColor(accentColor)
    .fontSize(12)
    .font('Helvetica-Bold')
    .text(template.organization.toUpperCase(), 160, currentY, { align: 'center', width: 522, characterSpacing: 2 })

  currentY += 30

  doc.fillColor('#172327')
    .fontSize(36)
    .font('Times-Bold')
    .text(template.title, 60, currentY, { align: 'center' })

  currentY += 50

  if (template.subtitle && template.subtitle.trim()) {
    doc.fillColor('#4b5563')
      .fontSize(13)
      .font('Helvetica-Oblique')
      .text(template.subtitle.trim(), 60, currentY, { align: 'center' })
    currentY += 30
  }

  const typeText = `SERTIFIKAT ${template.type.toUpperCase()}`
  doc.fillColor('#606f7b')
    .fontSize(10.5)
    .font('Helvetica-Bold')
    .text(typeText, 60, currentY, { align: 'center', characterSpacing: 2.5 })

  currentY += 34

  doc.fillColor('#4b5563')
    .fontSize(12.5)
    .font('Helvetica')
    .text('Diberikan dengan penuh kehormatan kepada', 60, currentY, { align: 'center' })

  currentY += 32

  doc.fillColor(accentColor)
    .fontSize(36)
    .font('Times-BoldItalic')
    .text(name, 50, currentY, { align: 'center' })

  currentY += 54

  doc.moveTo(210, currentY).lineTo(390, currentY).lineWidth(1.2).strokeColor('#d7c28a').stroke()
  doc.polygon([421, currentY - 4.5], [427, currentY], [421, currentY + 4.5], [415, currentY]).fillColor(accentColor).fill()
  doc.moveTo(452, currentY).lineTo(632, currentY).lineWidth(1.2).strokeColor('#d7c28a').stroke()

  currentY += 26

  const bodyStartY = currentY
  const bodyEndY = bodyStartY + bodyHeight + 12
  const sigBoxY = Math.max(bodyEndY + 28, Math.min(462, maxContentBottomY - 4))

  doc.fillColor('#334155')
    .fontSize(bodyFontSize)
    .font('Helvetica')
    .text(bodyText, 90, bodyStartY, {
      align: 'center',
      width: bodyWidth,
      lineGap: bodyLineGap
    })

  const isSig2Active = Boolean(template.enableSig2)

  if (isSig2Active) {
    const sealCY = sigBoxY + 38
    drawOfficialSeal(doc, 421, sealCY, accentColor, 'md')

    doc.save()
    doc.fillColor('#64748b')
      .fontSize(8.5)
      .font('Helvetica-Bold')
      .text('NO. SERTIFIKAT', 280, sealCY + 36, { width: 282, align: 'center', characterSpacing: 1.5 })
    doc.fillColor('#1e293b')
      .fontSize(10)
      .font('Helvetica')
      .text(certNumber, 280, sealCY + 50, { width: 282, align: 'center', characterSpacing: 1.2 })
    doc.restore()

    const sig2Label = template.signatoryTitle2 || 'Mengetahui'
    const sig2Name = template.signatory2 || template.organization

    if (template.signatureImage2 && template.signatureImage2.includes('base64,')) {
      try {
        const sig2Base64 = template.signatureImage2.substring(template.signatureImage2.indexOf('base64,') + 7)
        if (sig2Base64) {
          const sig2Buffer = Buffer.from(sig2Base64, 'base64')
          doc.image(sig2Buffer, 60, sigBoxY, { fit: [220, 86], align: 'center', valign: 'center' })
        }
      } catch (err) {
        console.error('Failed to render signature 2 in PDF:', err)
      }
    }

    doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig2Name, 55, sigBoxY + 88, { width: 210, align: 'center' })
    doc.moveTo(75, sigBoxY + 110).lineTo(245, sigBoxY + 110).lineWidth(1).strokeColor('#cbd5e1').stroke()
    doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9.5).text(sig2Label, 55, sigBoxY + 115, { width: 210, align: 'center' })
  } else {
    const sealCY = sigBoxY + 38
    drawOfficialSeal(doc, 155, sealCY, accentColor, 'lg')

    doc.save()
    doc.fillColor('#64748b')
      .fontSize(8.5)
      .font('Helvetica-Bold')
      .text('NO. SERTIFIKAT', 250, sealCY + 32, { width: 342, align: 'center', characterSpacing: 1.5 })
    doc.fillColor('#1e293b')
      .fontSize(10)
      .font('Helvetica')
      .text(certNumber, 250, sealCY + 46, { width: 342, align: 'center', characterSpacing: 1.2 })
    doc.restore()
  }

  const sig1Label = template.signatoryTitle1 || 'Ditetapkan secara resmi oleh'
  const sig1Name = template.signatory || 'Penandatangan'

  if (template.signatureImage && template.signatureImage.includes('base64,')) {
    try {
      const sigBase64 = template.signatureImage.substring(template.signatureImage.indexOf('base64,') + 7)
      if (sigBase64) {
        const sigBuffer = Buffer.from(sigBase64, 'base64')
        doc.image(sigBuffer, 580, sigBoxY, { fit: [220, 86], align: 'center', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render signature 1 in PDF:', err)
    }
  }

  doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig1Name, 577, sigBoxY + 88, { width: 210, align: 'center' })
  doc.moveTo(597, sigBoxY + 110).lineTo(767, sigBoxY + 110).lineWidth(1).strokeColor('#cbd5e1').stroke()
  doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9.5).text(sig1Label, 577, sigBoxY + 115, { width: 210, align: 'center' })
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
    const filename = `sertifikat-lengkap-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.zip`
    res.setHeader('Content-Type', 'application/zip')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)

    const archive = new ZipArchive({ zlib: { level: 6 }, forceZip64: true })
    archive.pipe(res)

    archive.on('error', (err: Error) => {
      console.error('Archive error:', err)
      if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file ZIP' })
    })

    const usedFilenames = new Map<string, number>()
    const pdfBuffers: { buffer: Buffer; name: string }[] = []

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
      pdfBuffers.push({ buffer: pdfBuffer, name: entryName })
    }

    for (const entry of pdfBuffers) {
      archive.append(entry.buffer, { name: entry.name, date: new Date() })
    }

    await archive.finalize()
  } catch (err) {
    console.error('Error generating ZIP:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file ZIP' })
  }
})

export default app

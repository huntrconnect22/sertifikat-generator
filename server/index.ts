import 'dotenv/config'
import cors from 'cors'
import express from 'express'
import mongoose, { Schema } from 'mongoose'
import multer from 'multer'
import PDFDocument from 'pdfkit'
import * as XLSX from 'xlsx'

const app = express()
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } })
app.use(cors())
app.use(express.json({ limit: '6mb' }))

type TemplatePayload = {
  name: string; title: string; type: string; primaryField: string; body: string
  signatory: string; organization: string; accent: string; logo?: string
}

const templateSchema = new Schema<TemplatePayload>({
  name: { type: String, required: true }, title: { type: String, required: true },
  type: { type: String, required: true }, primaryField: { type: String, required: true },
  body: { type: String, required: true }, signatory: { type: String, required: true },
  organization: { type: String, required: true }, accent: { type: String, required: true },
  logo: { type: String, default: '' },
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
app.post('/api/certificates/pdf', (req, res) => {
  const { template, recipient } = req.body as { template: TemplatePayload; recipient: Record<string, unknown> }
  if (!template || !recipient) return res.status(400).json({ message: 'Template dan penerima wajib ada.' })
  const name = value(recipient, template.primaryField) || 'Penerima'
  const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 42 })
  const filename = `sertifikat-${name.toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
  res.setHeader('Content-Type', 'application/pdf')
  res.setHeader('Content-Disposition', `attachment; filename=\"${filename}\"`)
  doc.pipe(res)
  doc.rect(18, 18, 806, 559).lineWidth(3).stroke(template.accent || '#0f766e')
  doc.rect(29, 29, 784, 537).lineWidth(1).stroke('#d7c28a')
  if (template.logo?.startsWith('data:image/')) {
    try {
      const base64 = template.logo.split(',')[1]
      if (base64) doc.image(Buffer.from(base64, 'base64'), 52, 57, { fit: [54, 54], align: 'center', valign: 'center' })
    } catch { /* Ignore invalid image data while producing the certificate. */ }
  }
  doc.fillColor(template.accent || '#0f766e').fontSize(12).font('Helvetica-Bold').text(template.organization.toUpperCase(), 50, 75, { align: 'center' })
  doc.fillColor('#152a31').fontSize(35).font('Times-Bold').text(template.title, 50, 120, { align: 'center' })
  doc.fillColor('#59666a').fontSize(12).font('Helvetica').text(`SERTIFIKAT ${template.type.toUpperCase()}`, 50, 170, { align: 'center', characterSpacing: 2 })
  doc.fillColor('#39454a').fontSize(13).text('Dengan bangga diberikan kepada', 50, 220, { align: 'center' })
  doc.fillColor(template.accent || '#0f766e').fontSize(30).font('Times-BoldItalic').text(name, 50, 250, { align: 'center' })
  doc.moveTo(260, 294).lineTo(580, 294).lineWidth(1).stroke('#d7c28a')
  doc.fillColor('#39454a').fontSize(13).font('Helvetica').text(interpolate(template.body, recipient), 110, 320, { align: 'center', width: 620, lineGap: 5 })
  doc.fillColor('#39454a').fontSize(12).text('Ditetapkan dengan penuh penghargaan', 535, 445, { width: 190, align: 'center' })
  doc.fillColor('#152a31').font('Times-Bold').fontSize(18).text(template.signatory, 535, 495, { width: 190, align: 'center' })
  doc.fillColor('#59666a').font('Helvetica').fontSize(10).text(template.organization, 535, 520, { width: 190, align: 'center' })
  doc.end()
})

if (!process.env.VERCEL) {
  app.listen(Number(process.env.PORT || 3001), () => console.log('Certificate API running on port 3001'))
}

export default app

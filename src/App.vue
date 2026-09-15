<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import * as XLSX from 'xlsx'
type Recipient = Record<string, string | number>
type Template = {
  _id?: string
  name: string
  title: string
  subtitle?: string
  type: string
  primaryField: string
  body: string
  signatory: string
  signatory2?: string
  signatoryTitle1?: string
  signatoryTitle2?: string
  organization: string
  accent: string
  logo?: string
  logo2?: string
  enableLogo2?: boolean
  enableSig2?: boolean
  signatureImage?: string
  signatureImage2?: string
  certNumberPrefix?: string
  certificateDate?: string
}
const colorPresets = [
  { name: 'Emerald Teal', value: '#0f766e' },
  { name: 'Classic Navy', value: '#1e3a8a' },
  { name: 'Royal Indigo', value: '#3730a3' },
  { name: 'Deep Crimson', value: '#991b1b' },
  { name: 'Burgundy', value: '#831843' },
  { name: 'Forest Green', value: '#166534' },
  { name: 'Warm Amber', value: '#b45309' },
  { name: 'Dark Slate', value: '#1e293b' },
  { name: 'Luxury Gold', value: '#9a7b2c' },
  { name: 'Charcoal Black', value: '#171717' },
]
const template = ref<Template>({
  name: 'Sertifikat Penyelesaian Kelas',
  title: 'Certificate of Achievement',
  subtitle: 'Atas dedikasi, partisipasi aktif, dan pencapaian luar biasa',
  type: 'PENGHARGAAN',
  primaryField: 'nama',
  body: 'telah menyelesaikan {{kelas}} dengan hasil yang membanggakan pada {{tanggal}}.',
  signatory: 'Dewi Anggraini',
  signatory2: 'Dr. Hendra Gunawan',
  signatoryTitle1: 'Ditetapkan secara resmi oleh',
  signatoryTitle2: 'Mengetahui / Pimpinan',
  organization: 'AKADEMI NUSANTARA',
  accent: '#0f766e',
  logo: '',
  logo2: '',
  enableLogo2: false,
  enableSig2: false,
  signatureImage: '',
  signatureImage2: '',
  certNumberPrefix: 'CERT',
  certificateDate: '14 September 2026'
})
const recipients = ref<Recipient[]>([{ nama: 'Nadia Pratama', kelas: 'Kelas Desain Produk', tanggal: '14 September 2026' }, { nama: 'Raka Wijaya', kelas: 'Kelas Desain Produk', tanggal: '14 September 2026' }])
const headers = ref(['nama', 'kelas', 'tanggal']), selectedIndex = ref(0), status = ref('Siap memproses data'), connection = ref<'connecting' | 'connected' | 'offline'>('connecting'), savedTemplates = ref<Template[]>([]), fileInput = ref<HTMLInputElement>(), logoInput = ref<HTMLInputElement>(), logo2Input = ref<HTMLInputElement>(), sigInput = ref<HTMLInputElement>(), sig2Input = ref<HTMLInputElement>()
const activeNav = ref<'editor' | 'templates'>('editor')
const showTemplatesModal = ref(false)

function navigateTo(target: 'editor' | 'templates') {
  activeNav.value = target
  if (target === 'editor') {
    showTemplatesModal.value = false
    window.scrollTo({ top: 0, behavior: 'smooth' })
  } else if (target === 'templates') {
    showTemplatesModal.value = true
    loadTemplates()
  }
}

async function deleteTemplate(id?: string) {
  if (!id) return
  if (!confirm('Apakah Anda yakin ingin menghapus template ini dari database?')) return
  try {
    status.value = 'Menghapus template…'
    const res = await fetch(`/api/templates/${id}`, { method: 'DELETE' })
    if (!res.ok) {
      status.value = (await res.json()).message || 'Gagal menghapus template'
      return
    }
    status.value = 'Template berhasil dihapus'
    await loadTemplates()
  } catch {
    status.value = 'Gagal menghubungi server untuk menghapus template'
  }
}
const selected = computed(() => recipients.value[selectedIndex.value] ?? {})
const displayName = computed(() => String(selected.value[template.value.primaryField] || 'Nama Penerima'))
const certNumber = computed(() => {
  for (const key of ['nomor', 'no_sertifikat', 'nomor_sertifikat', 'certificate_number', 'no']) {
    if (selected.value[key]) return String(selected.value[key]).trim()
  }
  const prefix = (template.value.certNumberPrefix || 'CERT').trim().toUpperCase()
  const rawId = `${displayName.value}_${template.value.title}_${template.value.signatory}`
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
})
const previewBody = computed(() =>
  template.value.body.replace(/{{\s*([^}]+)\s*}}/g, (_, key) => {
    const trimmed = key.trim()
    if (selected.value[trimmed] !== undefined && selected.value[trimmed] !== '') {
      return String(selected.value[trimmed])
    }
    if ((trimmed.toLowerCase() === 'tanggal' || trimmed.toLowerCase() === 'date') && template.value.certificateDate) {
      return template.value.certificateDate
    }
    return `{{${trimmed}}}`
  })
)
const placeholders = computed(() => {
  const list = [...headers.value.map((name) => `{{${name}}}`)]
  if (!headers.value.some((h) => h.toLowerCase() === 'tanggal')) {
    list.push('{{tanggal}}')
  }
  return list
})

async function importExcel(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  status.value = 'Membaca file Excel…'

  try {
    // 1. Parsing langsung di browser (cepat & tidak bergantung koneksi server)
    const arrayBuffer = await file.arrayBuffer()
    const workbook = XLSX.read(arrayBuffer, { type: 'array' })
    const firstSheetName = workbook.SheetNames[0]
    if (!firstSheetName) throw new Error('Sheet Excel kosong')
    const sheet = workbook.Sheets[firstSheetName]
    if (!sheet) throw new Error('Sheet Excel tidak valid')

    const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' })
    if (!rawRows.length) {
      status.value = 'File Excel tidak memiliki baris data.'
      return
    }

    const importedHeaders = Object.keys(rawRows[0] || {})
    headers.value = importedHeaders
    recipients.value = rawRows as Recipient[]
    selectedIndex.value = 0

    // 2. Deteksi otomatis kolom Nama (misal: '1. Nama Lengkap', 'Nama', 'Full Name', dll.)
    const nameCandidate = importedHeaders.find((h) =>
      /^(1\.\s*)?nama(\s*lengkap)?$/i.test(h.trim()) ||
      /nama/i.test(h) ||
      /name/i.test(h)
    )

    if (nameCandidate) {
      template.value.primaryField = nameCandidate
      status.value = `Berhasil mengimpor ${rawRows.length} peserta. Kolom nama terdeteksi otomatis: "${nameCandidate}".`
    } else {
      if (!importedHeaders.includes(template.value.primaryField)) {
        template.value.primaryField = importedHeaders[0] || ''
      }
      status.value = `${rawRows.length} penerima berhasil diimpor`
    }
  } catch (err: any) {
    console.warn('Gagal membaca excel di client, mencoba via server API:', err)
    // Fallback: via server API jika client gagal
    try {
      const form = new FormData()
      form.append('file', file)
      const response = await fetch('/api/import', { method: 'POST', body: form })
      const data = await response.json()
      if (!response.ok) { status.value = data.message; return }
      headers.value = data.headers
      recipients.value = data.rows
      selectedIndex.value = 0
      const nameCandidate = data.headers.find((h: string) => /nama/i.test(h) || /name/i.test(h))
      if (nameCandidate) template.value.primaryField = nameCandidate
      status.value = `${data.rows.length} penerima berhasil diimpor`
    } catch {
      status.value = 'Gagal membaca file Excel. Pastikan format file .xlsx atau .xls valid.'
    }
  }
}
const DRAFT_KEY = 'certify_draft_v1'
function saveDraft() {
  try {
    const draft = { ...template.value }
    sessionStorage.setItem(DRAFT_KEY, JSON.stringify(draft))
  } catch { /* storage full – ignore */ }
}
function restoreDraft() {
  try {
    const raw = sessionStorage.getItem(DRAFT_KEY)
    if (!raw) return
    const draft = JSON.parse(raw) as Template
    // Only restore image fields and accent – keep other fields from defaults
    if (draft.logo) template.value.logo = draft.logo
    if (draft.logo2) template.value.logo2 = draft.logo2
    if (draft.enableLogo2 !== undefined) template.value.enableLogo2 = draft.enableLogo2
    if (draft.enableSig2 !== undefined) template.value.enableSig2 = draft.enableSig2
    if (draft.signatureImage) template.value.signatureImage = draft.signatureImage
    if (draft.signatureImage2) template.value.signatureImage2 = draft.signatureImage2
    if (draft.accent) template.value.accent = draft.accent
    if (draft.name) template.value.name = draft.name
    if (draft.title) template.value.title = draft.title
    if (draft.subtitle !== undefined) template.value.subtitle = draft.subtitle
    if (draft.organization) template.value.organization = draft.organization
    if (draft.signatory) template.value.signatory = draft.signatory
    if (draft.signatory2) template.value.signatory2 = draft.signatory2
    if (draft.signatoryTitle1) template.value.signatoryTitle1 = draft.signatoryTitle1
    if (draft.signatoryTitle2) template.value.signatoryTitle2 = draft.signatoryTitle2
    if (draft.body) template.value.body = draft.body
    if (draft.type) template.value.type = draft.type
    if (draft.certNumberPrefix) template.value.certNumberPrefix = draft.certNumberPrefix
    if (draft.primaryField) template.value.primaryField = draft.primaryField
    if (draft.certificateDate !== undefined) template.value.certificateDate = draft.certificateDate
  } catch { /* ignore */ }
}
watch(template, saveDraft, { deep: true })
function processImageFile(file: File, maxDimension = 1200): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const img = new Image()
      img.onload = () => {
        let { width, height } = img
        if (width > maxDimension || height > maxDimension) {
          const ratio = Math.min(maxDimension / width, maxDimension / height)
          width = Math.round(width * ratio)
          height = Math.round(height * ratio)
        }
        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        const ctx = canvas.getContext('2d')
        if (!ctx) {
          resolve(String(reader.result))
          return
        }
        ctx.drawImage(img, 0, 0, width, height)
        // Always export as clean PNG so PDFKit can reliably parse it
        resolve(canvas.toDataURL('image/png'))
      }
      img.onerror = () => resolve(String(reader.result))
      img.src = String(reader.result)
    }
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

async function uploadLogo(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  if (!file.type.startsWith('image/')) { status.value = 'Logo 1 harus berupa file gambar.'; return }
  if (file.size > 5 * 1024 * 1024) { status.value = 'Ukuran logo maksimal 5 MB.'; return }
  try {
    status.value = 'Memproses logo 1…'
    template.value.logo = await processImageFile(file, 800)
    status.value = 'Logo 1 (Kanan Atas) berhasil diunggah'
  } catch {
    status.value = 'Gagal memproses gambar logo 1'
  }
}
function removeLogo() { template.value.logo = ''; if (logoInput.value) logoInput.value.value = ''; status.value = 'Logo 1 dihapus dari template' }

async function uploadLogo2(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  if (!file.type.startsWith('image/')) { status.value = 'Logo 2 harus berupa file gambar.'; return }
  if (file.size > 5 * 1024 * 1024) { status.value = 'Ukuran logo maksimal 5 MB.'; return }
  try {
    status.value = 'Memproses logo 2…'
    template.value.logo2 = await processImageFile(file, 800)
    status.value = 'Logo 2 (Kiri Atas) berhasil diunggah'
  } catch {
    status.value = 'Gagal memproses gambar logo 2'
  }
}
function removeLogo2() { template.value.logo2 = ''; if (logo2Input.value) logo2Input.value.value = ''; status.value = 'Logo 2 dihapus dari template' }

async function uploadSignature(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  if (!file.type.startsWith('image/')) { status.value = 'Tanda tangan 1 harus berupa file gambar.'; return }
  if (file.size > 5 * 1024 * 1024) { status.value = 'Ukuran tanda tangan maksimal 5 MB.'; return }
  try {
    status.value = 'Memproses tanda tangan 1…'
    template.value.signatureImage = await processImageFile(file, 800)
    status.value = 'Tanda tangan 1 (Kanan Bawah) berhasil diunggah'
  } catch {
    status.value = 'Gagal memproses gambar tanda tangan 1'
  }
}
function removeSignature() { template.value.signatureImage = ''; if (sigInput.value) sigInput.value.value = ''; status.value = 'Tanda tangan 1 dihapus' }

async function uploadSignature2(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  if (!file.type.startsWith('image/')) { status.value = 'Tanda tangan 2 harus berupa file gambar.'; return }
  if (file.size > 5 * 1024 * 1024) { status.value = 'Ukuran tanda tangan maksimal 5 MB.'; return }
  try {
    status.value = 'Memproses tanda tangan 2…'
    template.value.signatureImage2 = await processImageFile(file, 800)
    status.value = 'Tanda tangan 2 (Kiri Bawah) berhasil diunggah'
  } catch {
    status.value = 'Gagal memproses gambar tanda tangan 2'
  }
}
function removeSignature2() { template.value.signatureImage2 = ''; if (sig2Input.value) sig2Input.value.value = ''; status.value = 'Tanda tangan 2 dihapus' }
async function checkConnection() {
  try {
    const response = await fetch('/api/health')
    if (!response.ok) { connection.value = 'offline'; return }
    connection.value = (await response.json()).database
    if (connection.value === 'connected') loadTemplates()
  } catch {
    connection.value = 'offline'
  }
}
async function loadTemplates() {
  try {
    const response = await fetch('/api/templates')
    if (response.ok) savedTemplates.value = await response.json()
  } catch {
    // Server offline or proxy error – silently skip template list refresh
  }
}
function clearTable() {
  if (recipients.value.length && !confirm('Hapus semua data penerima dari tabel?')) return
  recipients.value = []
  headers.value = []
  selectedIndex.value = 0
  if (fileInput.value) fileInput.value.value = ''
  status.value = 'Data penerima telah dihapus'
}
async function saveTemplate() {
  if (connection.value !== 'connected') { status.value = 'MongoDB belum tersambung. Gunakan Simpan Format saat tersambung.'; return }
  const payload = { ...template.value }
  const response = await fetch('/api/templates', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) })
  if (!response.ok) { status.value = (await response.json()).message; return }
  status.value = 'Template (termasuk logo & TTD) tersimpan di MongoDB Atlas'
  loadTemplates()
}
function useTemplate(item: Template) {
  template.value = { ...item }
  saveDraft()
  showTemplatesModal.value = false
  activeNav.value = 'editor'
  window.scrollTo({ top: 0, behavior: 'smooth' })
  status.value = `Template "${item.name}" berhasil diterapkan`
}
const isExporting = ref(false)
async function exportPDF() {
  if (!recipients.value.length) { status.value = 'Impor data penerima terlebih dahulu.'; return }
  isExporting.value = true
  status.value = 'Menyiapkan PDF…'
  try {
    const response = await fetch('/api/certificates/pdf', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template: template.value, recipient: selected.value })
    })
    if (!response.ok) { status.value = (await response.json()).message; return }
    const url = URL.createObjectURL(await response.blob())
    const a = document.createElement('a')
    a.href = url
    a.download = `sertifikat-${displayName.value}.pdf`
    a.click()
    URL.revokeObjectURL(url)
    status.value = 'PDF berhasil diunduh'
  } catch {
    status.value = 'Gagal mengunduh PDF'
  } finally {
    isExporting.value = false
  }
}

async function exportBatchPDF() {
  if (!recipients.value.length) { status.value = 'Impor data penerima terlebih dahulu.'; return }
  isExporting.value = true
  status.value = `Menyiapkan ${recipients.value.length} sertifikat dalam 1 file PDF…`
  try {
    const response = await fetch('/api/certificates/batch-pdf', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template: template.value, recipients: recipients.value })
    })
    if (!response.ok) { status.value = (await response.json()).message; return }
    const url = URL.createObjectURL(await response.blob())
    const a = document.createElement('a')
    a.href = url
    const safeName = (template.value.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')
    a.download = `semua-sertifikat-${safeName}.pdf`
    a.click()
    URL.revokeObjectURL(url)
    status.value = `Berhasil mengunduh semua (${recipients.value.length}) sertifikat dalam 1 file PDF!`
  } catch {
    status.value = 'Gagal mengunduh batch PDF'
  } finally {
    isExporting.value = false
  }
}

async function exportBatchZIP() {
  if (!recipients.value.length) { status.value = 'Impor data penerima terlebih dahulu.'; return }
  isExporting.value = true
  status.value = `Mengompres ${recipients.value.length} file PDF ke dalam ZIP…`
  try {
    const response = await fetch('/api/certificates/zip', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template: template.value, recipients: recipients.value })
    })
    if (!response.ok) { status.value = (await response.json()).message; return }
    const url = URL.createObjectURL(await response.blob())
    const a = document.createElement('a')
    a.href = url
    const safeName = (template.value.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')
    a.download = `sertifikat-lengkap-${safeName}.zip`
    a.click()
    URL.revokeObjectURL(url)
    status.value = `Berhasil mengunduh ${recipients.value.length} sertifikat dalam file ZIP!`
  } catch {
    status.value = 'Gagal mengunduh ZIP sertifikat'
  } finally {
    isExporting.value = false
  }
}
onMounted(() => { restoreDraft(); checkConnection() })
</script>
<template>
  <main class="app-shell">
    <aside class="sidebar">
      <div class="brand" role="button" style="cursor:pointer" @click="navigateTo('editor')">
        <span class="brand-mark">✦</span>
        <span>certify</span>
      </div>
      <nav>
        <a :class="{ active: activeNav === 'editor' }" href="#editor" @click.prevent="navigateTo('editor')">
          Buat sertifikat
        </a>
        <a :class="{ active: activeNav === 'templates' }" href="#saved-templates-section" @click.prevent="navigateTo('templates')">
          <span>Template tersimpan</span>
          <b>{{ savedTemplates.length }}</b>
        </a>
      </nav>
      <div class="side-bottom">
        <span class="dot" :class="connection"></span>
        <span>{{ connection === 'connected' ? 'MongoDB Atlas terhubung' : connection === 'connecting' ? 'Menghubungkan database…' : 'Mode offline' }}</span>
      </div>
    </aside>
    <section class="workspace"><header class="topbar"><div><p class="eyebrow">SERTIFIKAT GENERATOR</p><h1>Kelola & Terbitkan Sertifikat</h1></div><button class="outline" @click="saveTemplate">Simpan Format</button></header><p class="status">{{ status }}</p>
      <div class="content-grid"><section class="editor-card"><div class="section-title"><span class="step">FORMAT</span><div><h2>Pengaturan Konten</h2><p>Sesuaikan teks dan parameter sertifikat untuk setiap penerima.</p></div></div><label>Nama template<input v-model="template.name" /></label><div class="two-cols"><label>Judul sertifikat<input v-model="template.title" /></label><label>Subjudul sertifikat<input v-model="template.subtitle" placeholder="Contoh: Atas dedikasi dan pencapaian luar biasa" /></label></div><div class="two-cols"><label>Jenis<span class="select-wrap"><select v-model="template.type"><option>APRESIASI</option><option>PENGHARGAAN</option><option>KELULUSAN</option><option>PARTISIPASI</option></select></span></label><label>Organisasi<input v-model="template.organization" /></label></div><label>Kolom nama penerima<span class="select-wrap"><select v-model="template.primaryField"><option v-for="header in headers" :key="header">{{ header }}</option></select></span></label><label>Isi penghargaan<textarea v-model="template.body" rows="3" /></label><div class="tokens"><span>Masukkan kolom Excel:</span><button v-for="token in placeholders" :key="token" @click="template.body += ` ${token}`">{{ token }}</button></div>

<div class="form-subheading-wrap">
  <div class="form-subheading">Pengaturan Logo (Kiri & Kanan Atas)</div>
  <label class="toggle-switch">
    <input type="checkbox" v-model="template.enableLogo2" />
    <span class="toggle-slider"></span>
    <span class="toggle-label">{{ template.enableLogo2 ? 'Logo 2 (Kiri Atas) Aktif' : 'Logo 2 Dinonaktifkan' }}</span>
  </label>
</div>
<div class="two-cols">
  <div class="logo-field" :class="{ 'is-disabled': !template.enableLogo2 }">
    <div>
      <span>Logo 2 (Kiri Atas)</span>
      <small>{{ template.enableLogo2 ? 'PNG, JPG, maks. 5 MB.' : 'Aktifkan toggle di atas untuk menampilkan.' }}</small>
    </div>
    <button class="upload" :disabled="!template.enableLogo2" @click="logo2Input?.click()">{{ template.logo2 ? 'Ganti' : 'Unggah' }}</button>
    <input ref="logo2Input" hidden type="file" accept="image/png,image/jpeg,image/webp" @change="uploadLogo2" />
    <button v-if="template.logo2 && template.enableLogo2" class="remove-logo" @click="removeLogo2">Hapus</button>
  </div>
  <div class="logo-field">
    <div>
      <span>Logo 1 (Kanan Atas)</span>
      <small>PNG, JPG, maks. 5 MB.</small>
    </div>
    <button class="upload" @click="logoInput?.click()">{{ template.logo ? 'Ganti' : 'Unggah' }}</button>
    <input ref="logoInput" hidden type="file" accept="image/png,image/jpeg,image/webp" @change="uploadLogo" />
    <button v-if="template.logo" class="remove-logo" @click="removeLogo">Hapus</button>
  </div>
</div>

<div class="form-subheading-wrap">
  <div class="form-subheading">Pengaturan Tanda Tangan (Kiri & Kanan Bawah)</div>
  <label class="toggle-switch">
    <input type="checkbox" v-model="template.enableSig2" />
    <span class="toggle-slider"></span>
    <span class="toggle-label">{{ template.enableSig2 ? 'TTD 2 (Kiri Bawah) Aktif' : 'TTD 2 Dinonaktifkan' }}</span>
  </label>
</div>

<div v-if="template.enableSig2" class="sig2-section">
  <div class="two-cols">
    <label>Jabatan TTD 2 (Kiri Bawah)<input v-model="template.signatoryTitle2" placeholder="Mengetahui / Pimpinan" /></label>
    <label>Nama TTD 2 (Kiri Bawah)<input v-model="template.signatory2" placeholder="Nama Penandatangan 2" /></label>
  </div>
  <div class="logo-field">
    <div>
      <span>File Tanda Tangan 2 (Kiri Bawah)</span>
      <small>PNG Transparan disarankan, maks. 5 MB.</small>
    </div>
    <button class="upload" @click="sig2Input?.click()">{{ template.signatureImage2 ? 'Ganti' : 'Unggah' }}</button>
    <input ref="sig2Input" hidden type="file" accept="image/png,image/jpeg,image/webp" @change="uploadSignature2" />
    <button v-if="template.signatureImage2" class="remove-logo" @click="removeSignature2">Hapus</button>
  </div>
</div>
<div v-else class="disabled-notice">
  <span>Tanda Tangan 2 (Kiri Bawah) sedang dinonaktifkan. Hanya TTD 1 yang akan ditampilkan pada sertifikat.</span>
</div>

<div class="two-cols" style="margin-top: 10px;">
  <label>Jabatan TTD 1 (Kanan Bawah)<input v-model="template.signatoryTitle1" placeholder="Ditetapkan secara resmi oleh" /></label>
  <label>Nama TTD 1 (Kanan Bawah)<input v-model="template.signatory" placeholder="Nama Penandatangan 1" /></label>
</div>
<div class="logo-field">
  <div>
    <span>File Tanda Tangan 1 (Kanan Bawah)</span>
    <small>PNG Transparan disarankan, maks. 5 MB.</small>
  </div>
  <button class="upload" @click="sigInput?.click()">{{ template.signatureImage ? 'Ganti' : 'Unggah' }}</button>
  <input ref="sigInput" hidden type="file" accept="image/png,image/jpeg,image/webp" @change="uploadSignature" />
  <button v-if="template.signatureImage" class="remove-logo" @click="removeSignature">Hapus</button>
</div>

<div class="two-cols">
  <label>Prefix / Format No. Sertifikat
    <input v-model="template.certNumberPrefix" placeholder="Contoh: CERT, SK, BATCH1" />
  </label>
  <label>Tanggal Sertifikat (Konfigurasi)
    <input v-model="template.certificateDate" placeholder="Contoh: 14 September 2026" />
  </label>
</div>
<div class="color-palette-section"><label>Pilihan warna tema sertifikat</label><div class="swatches"><button v-for="c in colorPresets" :key="c.value" type="button" class="swatch-btn" :class="{ active: template.accent.toLowerCase() === c.value.toLowerCase() }" :style="{ backgroundColor: c.value }" :title="c.name" @click="template.accent = c.value"><span v-if="template.accent.toLowerCase() === c.value.toLowerCase()" class="check-icon">✓</span></button></div><label class="color-field"><span>Warna kustom:</span><input v-model="template.accent" type="color" /><code>{{ template.accent }}</code></label></div></section>
        <section class="preview-area"><div class="section-title"><span class="step">PREVIEW</span><div><h2>Hasil sertifikat</h2><p>Pratinjau dari data penerima aktif.</p></div></div><article class="certificate" :style="{ '--accent': template.accent }"><div class="corner c-tl"></div><div class="corner c-tr"></div><div class="corner c-bl"></div><div class="corner c-br"></div><div class="inner"><div class="cert-header">
          <div class="header-logos-row" :class="{ 'single-logo-mode': !template.enableLogo2 }">
            <div v-if="template.enableLogo2" class="logo-slot left-slot">
              <img v-if="template.logo2" class="certificate-logo" :src="template.logo2" alt="Logo 2 (Kiri)" />
              <div v-else class="logo-placeholder"><span>Logo 2 (Kiri)</span></div>
            </div>
            <div class="header-center-info">
              <span v-if="!template.logo && (!template.enableLogo2 || !template.logo2)" class="emblem-icon">✦</span>
              <p class="org">{{ template.organization }}</p>
            </div>
            <div class="logo-slot right-slot">
              <img v-if="template.logo" class="certificate-logo" :src="template.logo" alt="Logo 1 (Kanan)" />
              <div v-else class="logo-placeholder"><span>Logo 1 (Kanan)</span></div>
            </div>
          </div>
        </div><div class="cert-main"><h3>{{ template.title }}</h3><p v-if="template.subtitle" class="cert-subtitle">{{ template.subtitle }}</p><div class="cert-type-badge">SERTIFIKAT {{ template.type }}</div><p class="awarded">Diberikan dengan penuh kehormatan kepada</p><p class="recipient">{{ displayName }}</p><div class="rule"><span class="rule-diamond">◆</span></div><p class="body-copy">{{ previewBody }}</p></div><div class="cert-footer" :class="{ 'single-sig-footer': !template.enableSig2 }">
          <div v-if="template.enableSig2" class="signature left-sig">
            <div v-if="template.signatureImage2" class="sig-img-wrap"><img class="sig-img" :src="template.signatureImage2" alt="Tanda tangan 2" /></div>
            <div v-else class="sig-spacer"></div>
            <strong>{{ template.signatory2 || template.organization }}</strong>
            <div class="sig-line"></div>
            <small>{{ template.signatoryTitle2 || 'Mengetahui' }}</small>
            <small class="sig-org">{{ template.organization }}</small>
          </div>
          <div class="footer-center">
            <div class="official-seal"><div class="seal-inner"><span class="seal-top">AUTHENTIC</span><span class="seal-star">★</span><span class="seal-bot">VERIFIED</span></div></div>
            <div class="cert-bottom-id"><span class="cert-id-label">NO. SERTIFIKAT</span><span class="cert-id-val">{{ certNumber }}</span></div>
          </div>
          <div class="signature right-sig">
            <div v-if="template.signatureImage" class="sig-img-wrap"><img class="sig-img" :src="template.signatureImage" alt="Tanda tangan 1" /></div>
            <div v-else class="sig-spacer"></div>
            <strong>{{ template.signatory || 'Penandatangan' }}</strong>
            <div class="sig-line"></div>
            <small>{{ template.signatoryTitle1 || 'Ditetapkan secara resmi oleh' }}</small>
            <small class="sig-org">{{ template.organization }}</small>
          </div>
        </div></div></article>        <div class="export-actions">
          <button class="primary full" :disabled="isExporting" @click="exportPDF">
            {{ isExporting ? 'Memproses…' : `Unduh PDF untuk ${displayName}` }} <span>→</span>
          </button>
          <div class="batch-buttons">
            <button class="batch-btn" :disabled="isExporting || !recipients.length" @click="exportBatchPDF">
              <svg class="btn-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
              Unduh Semua (1 File PDF) <b>{{ recipients.length }}</b>
            </button>
            <button class="batch-btn zip" :disabled="isExporting || !recipients.length" @click="exportBatchZIP">
              <svg class="btn-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
              Unduh Arsip ZIP (.zip) <b>{{ recipients.length }}</b>
            </button>
          </div>
        </div>
      </section></div>
      <section class="data-card"><div class="section-title"><span class="step">PESERTA</span><div><h2>Data Penerima</h2><p>Daftar nama dan data peserta. Gunakan kolom Excel sebagai placeholder teks sertifikat.</p></div><div class="data-actions"><button v-if="recipients.length" class="clear-btn" title="Hapus semua data" @click="clearTable">✕ Hapus Data</button><button class="upload" @click="fileInput?.click()">↑ Impor Excel</button><input ref="fileInput" hidden type="file" accept=".xlsx,.xls" @change="importExcel" /></div></div><div v-if="recipients.length" class="table-wrap"><table><thead><tr><th>#</th><th v-for="header in headers" :key="header">{{ header }}</th><th></th></tr></thead><tbody><tr v-for="(row, index) in recipients" :key="index" :class="{ chosen: selectedIndex === index }"><td>{{ String(index + 1).padStart(2, '0') }}</td><td v-for="header in headers" :key="header">{{ row[header] }}</td><td><button class="select" @click="selectedIndex = index">{{ selectedIndex === index ? 'Dipilih' : 'Pilih' }}</button></td></tr></tbody></table></div><div v-else class="empty">Belum ada data. Impor file Excel dengan baris pertama sebagai header.</div></section>
      <section id="saved-templates-section" class="saved">
        <div class="saved-header">
          <div>
            <h2>Template Tersimpan ({{ savedTemplates.length }})</h2>
            <p class="saved-desc">Format sertifikat yang tersimpan di database MongoDB Atlas.</p>
          </div>
          <button class="outline" style="padding:6px 12px; font-size:12px;" @click="loadTemplates">↻ Segarkan</button>
        </div>
        <div v-if="!savedTemplates.length" class="empty" style="padding:18px; width:100%">
          Belum ada template tersimpan di MongoDB. Sesuaikan format sertifikat di atas, lalu klik <strong>"Simpan Format"</strong> di kanan atas untuk menyimpan.
        </div>
        <div v-else class="template-cards-grid">
          <div v-for="item in savedTemplates" :key="item._id" class="template-card">
            <div class="template-card-top">
              <span class="template-color-dot" :style="{ background: item.accent || '#0f766e' }"></span>
              <span class="template-type-tag">{{ item.type }}</span>
              <button class="delete-template-btn" title="Hapus template" @click.stop="deleteTemplate(item._id)">✕</button>
            </div>
            <h3 class="template-card-name">{{ item.name }}</h3>
            <p class="template-card-title">{{ item.title }}</p>
            <div class="template-card-meta">
              <span>{{ item.organization }}</span>
              <span v-if="item.logo || item.logo2" class="feature-badge">Logo</span>
              <span v-if="item.signatureImage || item.signatureImage2" class="feature-badge">TTD</span>
            </div>
            <button class="use-template-btn" @click="useTemplate(item)">Gunakan Template →</button>
          </div>
        </div>
      </section>
    </section>

    <!-- Modal Daftar Template Tersimpan -->
    <div v-if="showTemplatesModal" class="modal-backdrop" @click.self="showTemplatesModal = false">
      <div class="modal-card">
        <div class="modal-header">
          <div>
            <h2>Koleksi Template Tersimpan</h2>
            <p class="saved-desc">Pilih template untuk langsung dimuat ke editor sertifikat.</p>
          </div>
          <button class="modal-close-btn" @click="showTemplatesModal = false">✕</button>
        </div>
        <div class="modal-body">
          <div v-if="!savedTemplates.length" class="empty" style="padding:32px 16px;">
            <p style="font-size:15px; margin-bottom:6px;">Belum ada template yang tersimpan.</p>
            <p style="font-size:12px; color:#768480;">Ketik pengaturan sertifikat pada editor lalu klik tombol "Simpan Format" di sudut kanan atas.</p>
          </div>
          <div v-else class="template-cards-grid modal-grid">
            <div v-for="item in savedTemplates" :key="item._id" class="template-card">
              <div class="template-card-top">
                <span class="template-color-dot" :style="{ background: item.accent || '#0f766e' }"></span>
                <span class="template-type-tag">{{ item.type }}</span>
                <button class="delete-template-btn" title="Hapus template" @click.stop="deleteTemplate(item._id)">✕</button>
              </div>
              <h3 class="template-card-name">{{ item.name }}</h3>
              <p class="template-card-title">{{ item.title }}</p>
              <div class="template-card-meta">
                <span>{{ item.organization }}</span>
                <span v-if="item.logo || item.logo2" class="feature-badge">Logo</span>
                <span v-if="item.signatureImage || item.signatureImage2" class="feature-badge">TTD</span>
              </div>
              <button class="use-template-btn" @click="useTemplate(item)">Gunakan Template →</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  </main>
</template>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;500;600;700&family=Playfair+Display:ital,wght@0,600;0,700;1,600&display=swap');
*{box-sizing:border-box}body{margin:0;background:#f6f5f0;color:#1b292b;font-family:'DM Sans',sans-serif}.app-shell{min-height:100vh;display:flex}.sidebar{width:235px;background:#153538;color:#e8f2ed;padding:32px 20px;display:flex;flex-direction:column;position:fixed;inset:0 auto 0 0}.brand{display:flex;align-items:center;gap:10px;font-family:'Playfair Display';font-size:28px;font-weight:700;padding:0 10px}.brand-mark{display:grid;place-items:center;background:#d4e9a2;color:#153538;width:28px;height:28px;border-radius:9px;font-family:serif;font-size:18px}nav{margin-top:72px;display:grid;gap:7px}nav a{padding:12px 14px;border-radius:9px;color:#b9cfcb;font-size:14px;display:flex;justify-content:space-between;align-items:center;cursor:pointer;text-decoration:none;user-select:none;transition:all .18s ease}nav a:hover{background:#1d4245;color:#fff}nav .active{background:#244a4c;color:#fff;font-weight:600}.side-bottom{font-size:12px;color:#9cb9b4;margin-top:auto;display:flex;gap:8px;align-items:center;padding:9px}.dot{width:8px;height:8px;background:#d4a538;border-radius:50%}.dot.connected{background:#a9d477}.dot.offline{background:#df8979}.workspace{padding:42px 52px 70px;margin-left:235px;width:calc(100% - 235px);max-width:1600px}.topbar{display:flex;justify-content:space-between;align-items:flex-start}.eyebrow{font-family:'DM Mono';font-size:11px;color:#71817d;letter-spacing:1.2px;margin:0 0 8px}h1{font-family:'Playfair Display';font-size:36px;margin:0;letter-spacing:-.8px}h2{font-size:16px;margin:0 0 4px}p{margin:0}.status{font-size:13px;color:#71817d;margin:13px 0 24px}.outline,.upload,.select,.remove-logo{border:1px solid #b9c6c1;background:transparent;color:#234447;border-radius:7px;padding:10px 14px;font:600 13px 'DM Sans';cursor:pointer}.content-grid{display:grid;grid-template-columns:minmax(400px,1fr) minmax(420px,1.05fr);gap:24px}.editor-card,.data-card{background:#fff;border:1px solid #e1e5df;border-radius:14px;padding:26px;box-shadow:0 6px 22px #19332d08}.section-title{display:flex;align-items:flex-start;gap:11px;margin-bottom:22px}.section-title p{font-size:13px;color:#768480}.step{font:500 10px 'DM Mono';color:#48736e;border:1px solid #b9cec7;background:#edf5ef;padding:6px 7px;border-radius:5px;margin-top:1px}.editor-card label{display:block;font-size:12px;font-weight:600;color:#52615d;margin:15px 0 0}.editor-card input:not([type=color]),textarea,select{font:400 14px 'DM Sans';color:#1b292b;width:100%;border:1px solid #d5ded9;border-radius:7px;background:#fcfdfb;padding:10px;margin-top:7px;outline-color:#5c928a}textarea{resize:vertical;line-height:1.45}.two-cols{display:grid;grid-template-columns:1fr 1fr;gap:14px}.select-wrap{display:block;position:relative}.select-wrap:after{content:'⌄';position:absolute;right:12px;bottom:8px;color:#62716d;pointer-events:none}select{appearance:none}.tokens{display:flex;gap:7px;align-items:center;flex-wrap:wrap;margin-top:11px;font-size:11px;color:#7b8985}.tokens button{border:0;border-radius:4px;background:#e8f1ed;color:#37655f;padding:4px 6px;font:11px 'DM Mono';cursor:pointer}.logo-field{border:1px dashed #bfd1c8;background:#f7faf7;border-radius:8px;padding:10px 11px;margin-top:16px;display:flex;align-items:center;gap:9px}.logo-field div{margin-right:auto;display:grid;gap:2px;font-size:12px;font-weight:600;color:#52615d}.logo-field small{font-size:10px;font-weight:400;color:#7b8985}.logo-field .upload,.logo-field .remove-logo{padding:7px 9px;font-size:11px}.remove-logo{color:#9c3d3a;border-color:#e2b7b3}.color-palette-section{margin-top:16px}.color-palette-section>label{margin-bottom:8px}.swatches{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}.swatch-btn{width:28px;height:28px;border-radius:6px;border:2px solid transparent;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;transition:transform .15s ease,box-shadow .15s ease;padding:0}.swatch-btn:hover{transform:scale(1.12);box-shadow:0 3px 8px rgba(0,0,0,0.2)}.swatch-btn.active{border-color:#1b292b;box-shadow:0 0 0 2px #fff inset, 0 3px 8px rgba(0,0,0,0.25);transform:scale(1.08)}.check-icon{color:#fff;font-size:13px;font-weight:700;text-shadow:0 1px 2px rgba(0,0,0,0.6);line-height:1}.color-field{display:flex!important;align-items:center;gap:10px;margin-top:4px!important}.color-field input{width:28px;height:28px;padding:2px;border:1px solid #d5ded9;border-radius:5px;margin:0 0 0 4px;cursor:pointer}.color-field code{font:12px 'DM Mono';color:#73817e}.preview-area{min-width:0;overflow:hidden}.certificate{width:100%;background:#fff;padding:12px;border:3px solid var(--accent);box-shadow:0 12px 36px rgba(15,23,42,0.12);aspect-ratio:1.414;margin-bottom:18px;position:relative;display:flex;overflow:hidden}.corner{position:absolute;width:24px;height:24px;border:2px solid var(--accent);pointer-events:none;z-index:2}.c-tl{top:18px;left:18px;border-right:none;border-bottom:none}.c-tr{top:18px;right:18px;border-left:none;border-bottom:none}.c-bl{bottom:18px;left:18px;border-right:none;border-top:none}.c-br{bottom:18px;right:18px;border-left:none;border-top:none}.corner::after{content:'';position:absolute;width:4px;height:4px;background:#d7c28a;border-radius:50%}.c-tl::after{top:3px;left:3px}.c-tr::after{top:3px;right:3px}.c-bl::after{bottom:3px;left:3px}.c-br::after{bottom:3px;right:3px}.inner{height:100%;width:100%;border:1px solid #d7c28a;text-align:center;padding:5% 7%;position:relative;display:flex;flex-direction:column;justify-content:space-between;background:#fff}.cert-header{display:flex;flex-direction:column;align-items:center;gap:6px}.logo-wrapper{height:80px;flex-shrink:0;display:flex;align-items:center;justify-content:center}.certificate-logo{width:100%;height:100%;object-fit:contain;filter:drop-shadow(0 2px 6px rgba(0,0,0,0.08))}.emblem-icon{font-size:36px;color:var(--accent);line-height:1}.org{font:700 11px 'DM Mono';letter-spacing:2px;color:var(--accent);margin:2px 0 0;text-transform:uppercase}.cert-main{margin:auto 0}.certificate h3{font:700 clamp(20px,2.8vw,36px) 'Playfair Display';color:#172327;margin:0 0 4px;letter-spacing:-0.5px}.cert-type-badge{display:inline-block;font:600 9px 'DM Mono';letter-spacing:2.5px;color:#606f7b;background:rgba(0,0,0,0.03);border:1px solid #e2e8f0;padding:3px 10px;border-radius:20px;margin-bottom:6px}.cert-bottom-id{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:0 8px}.cert-id-label{font:700 8px 'DM Mono';color:#64748b;letter-spacing:1.5px}.cert-id-val{font:500 10px 'DM Mono';color:#1e293b;letter-spacing:1.2px;background:rgba(0,0,0,0.03);border:1px dashed #cbd5e1;padding:2px 8px;border-radius:4px}.awarded{font-size:11px;color:#4b5563;margin-top:6px;font-style:italic}.recipient{font:700 italic clamp(22px,3.2vw,40px) 'Playfair Display';color:var(--accent);margin:4px 0 6px;text-shadow:0 1px 1px rgba(0,0,0,0.05)}.rule{display:flex;align-items:center;justify-content:center;width:60%;margin:0 auto 10px;position:relative}.rule::before,.rule::after{content:'';flex:1;height:1px;background:#d7c28a}.rule-diamond{color:var(--accent);font-size:10px;padding:0 8px}.body-copy{font-size:11.5px;line-height:1.55;color:#334155;max-width:540px;margin:0 auto}.cert-footer{display:flex;align-items:flex-end;justify-content:space-between;padding:0 10px;margin-top:auto}.official-seal{width:56px;height:56px;border-radius:50%;border:1.5px solid #d7c28a;padding:2px;display:flex;align-items:center;justify-content:center}.seal-inner{width:100%;height:100%;border-radius:50%;border:1px solid var(--accent);display:flex;flex-direction:column;align-items:center;justify-content:center;background:rgba(255,255,255,0.7)}.seal-top,.seal-bot{font:700 5.5px 'DM Mono';color:var(--accent);letter-spacing:1px}.seal-star{font-size:10px;color:#d7c28a;line-height:1;margin:1px 0}.signature{display:grid;gap:2px;text-align:center;color:#64748b;font-size:9.5px;min-width:160px;position:relative}.sig-img-wrap{height:52px;flex-shrink:0;display:flex;align-items:center;justify-content:center;margin:2px 0}.sig-img{height:48px;max-width:150px;width:auto;object-fit:contain;filter:contrast(1.1)}.sig-spacer{height:52px;flex-shrink:0}.signature strong{font:600 17px 'Playfair Display';color:#172327;margin-top:2px}.sig-line{height:1px;background:#cbd5e1;width:100%;margin:4px 0 2px}.signature small{font:600 9px 'DM Mono';color:#64748b}.sig-org{font:400 8.5px 'DM Mono'!important;color:#94a3b8!important;margin-top:1px}.primary{background:#1e5a55;color:#fff;border:0;border-radius:7px;padding:13px 16px;font:600 14px 'DM Sans';cursor:pointer}.primary:disabled{opacity:0.6;cursor:not-allowed}.primary span{margin-left:12px;font-size:18px}.full{width:100%}.export-actions{display:flex;flex-direction:column;gap:10px}.batch-buttons{display:grid;grid-template-columns:1fr 1fr;gap:10px}.batch-btn{background:#fff;border:1.5px solid #1e5a55;color:#1e5a55;border-radius:8px;padding:11px 12px;font:600 12.5px 'DM Sans';cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:6px;transition:all .15s ease}.batch-btn:hover:not(:disabled){background:#edf5ef;transform:translateY(-1px);box-shadow:0 3px 8px rgba(30,90,85,0.12)}.batch-btn:disabled{opacity:0.5;cursor:not-allowed}.batch-btn.zip{border-color:#b45309;color:#b45309}.batch-btn.zip:hover:not(:disabled){background:#fef8ee;box-shadow:0 3px 8px rgba(180,83,9,0.12)}.batch-btn b{background:rgba(0,0,0,0.06);padding:2px 7px;border-radius:12px;font:600 11px 'DM Mono'}.btn-svg{width:16px;height:16px;flex-shrink:0}.data-card{margin-top:25px}.data-card .section-title{margin-bottom:14px}.upload{margin-left:auto}.table-wrap{overflow:auto;border:1px solid #e4e9e5;border-radius:8px}table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}th{text-align:left;background:#f3f6f3;color:#72807d;font:500 10px 'DM Mono';letter-spacing:.6px;text-transform:uppercase}th,td{padding:13px 14px;border-bottom:1px solid #e8ece8}tbody tr:last-child td{border:0}td:first-child{font:11px 'DM Mono';color:#87938f}.chosen{background:#f1f8f3}.select{padding:6px 10px;font-size:12px}.chosen .select{border-color:#7bac87;color:#246144;background:#f5fbf4}.empty{padding:28px;border:1px dashed #cbd6d0;color:#74817d;text-align:center;border-radius:8px;font-size:13px}.saved{margin-top:23px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}.saved h2{margin-right:8px}.saved button{border:1px solid #d9e2dc;background:#fff;border-radius:100px;padding:8px 12px;color:#315854;font:500 12px 'DM Sans';cursor:pointer}.saved button span{margin-left:6px}.cert-subtitle{font:400 italic clamp(11px,1.3vw,14px) 'Playfair Display';color:#4b5563;margin:0 0 6px;letter-spacing:0.3px}.form-subheading{font:600 12px 'DM Mono';color:#315854;text-transform:uppercase;letter-spacing:1px;margin:18px 0 6px;padding-bottom:4px;border-bottom:1px solid #e1e7e4}.header-logos-row{display:flex;align-items:center;justify-content:space-between;width:100%;min-height:76px}.logo-slot{width:120px;height:72px;display:flex;align-items:center;flex-shrink:0;overflow:hidden}.left-slot{justify-content:flex-start}.right-slot{justify-content:flex-end}.logo-placeholder{border:1px dashed #cbd5e1;border-radius:6px;width:100%;height:100%;display:flex;align-items:center;justify-content:center;font:500 10px 'DM Mono';color:#94a3b8;background:rgba(255,255,255,0.4)}.header-center-info{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:0 12px}.footer-center{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px}.left-sig{text-align:center;min-width:140px;max-width:180px}.right-sig{text-align:center;min-width:140px;max-width:180px}.form-subheading-wrap{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin:18px 0 6px;padding-bottom:4px;border-bottom:1px solid #e1e7e4}.form-subheading-wrap .form-subheading{margin:0;padding:0;border:none}.toggle-switch{display:inline-flex;align-items:center;gap:7px;font:500 11px 'DM Sans';color:#315854;cursor:pointer;user-select:none}.toggle-switch input{position:absolute;opacity:0;width:0;height:0}.toggle-slider{position:relative;display:inline-block;width:34px;height:18px;background:#cbd5e1;border-radius:20px;transition:.2s ease}.toggle-slider:before{position:absolute;content:'';height:14px;width:14px;left:2px;bottom:2px;background:#fff;border-radius:50%;transition:.2s ease}.toggle-switch input:checked + .toggle-slider{background:#1e5a55}.toggle-switch input:checked + .toggle-slider:before{transform:translateX(16px)}.toggle-label{font-weight:600}.logo-field.is-disabled{opacity:0.55;background:#f1f5f3;border-style:dotted}.disabled-notice{font:400 12px 'DM Sans';color:#71817d;background:#f8faf9;border:1px dashed #d1dcd6;border-radius:8px;padding:10px 14px;margin-top:10px}.header-logos-row.single-logo-mode{justify-content:space-between}.single-sig-footer{justify-content:space-between}.single-sig-footer .footer-center{align-items:flex-start}.single-sig-footer .official-seal{margin-left:20px}.single-sig-footer .cert-bottom-id{margin-left:14px}@media(max-width:980px){.sidebar{width:70px;padding:26px 10px}.brand span:last-child,nav a:not(.active),.side-bottom span:last-child{display:none}.brand{padding:0;justify-content:center}.workspace{margin-left:70px;width:calc(100% - 70px);padding:32px}.content-grid{grid-template-columns:1fr}}@media(max-width:620px){.sidebar{display:none}.workspace{margin:0;width:100%;padding:23px 16px}.topbar{gap:14px}.topbar h1{font-size:29px}.outline{white-space:nowrap}.content-grid,.two-cols{grid-template-columns:1fr}.certificate{margin-top:10px}.upload{margin-left:0}.data-card .section-title{flex-wrap:wrap}.logo-field{flex-wrap:wrap}.logo-field div{width:100%}}
.data-actions{display:flex;align-items:center;gap:8px;margin-left:auto}.clear-btn{border:1px solid #e2b7b3;background:transparent;color:#9c3d3a;border-radius:7px;padding:8px 11px;font:600 12px 'DM Sans';cursor:pointer;white-space:nowrap;transition:all .15s}.clear-btn:hover{background:#fee2e2;border-color:#dc2626;color:#dc2626}.saved-header{display:flex;justify-content:space-between;align-items:flex-start;width:100%;margin-bottom:14px}
.saved-desc{font-size:12px;color:#71817d;margin-top:2px}
.template-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px;width:100%;margin-top:10px}
.template-card{background:#fff;border:1px solid #dce4e0;border-radius:12px;padding:16px;display:flex;flex-direction:column;gap:8px;position:relative;box-shadow:0 3px 10px rgba(0,0,0,0.03);transition:all .18s ease}
.template-card:hover{border-color:#1e5a55;box-shadow:0 6px 18px rgba(30,90,85,0.09);transform:translateY(-2px)}
.template-card-top{display:flex;align-items:center;gap:8px}
.template-color-dot{width:12px;height:12px;border-radius:50%;display:inline-block;box-shadow:0 0 0 2px #fff,0 0 0 3px #cfded7}
.template-type-tag{font:600 10px 'DM Mono';background:#edf5ef;color:#29635b;padding:2px 7px;border-radius:4px;letter-spacing:0.8px}
.delete-template-btn{margin-left:auto;background:transparent;border:0;color:#94a3b8;font-size:14px;cursor:pointer;padding:2px 6px;border-radius:4px;transition:all .15s}
.delete-template-btn:hover{color:#dc2626;background:#fee2e2}
.template-card-name{font:700 15px 'DM Sans';color:#172327;margin:4px 0 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.template-card-title{font:500 13px 'Playfair Display';color:#4b5563;margin:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.template-card-meta{display:flex;align-items:center;gap:6px;font-size:11px;color:#788884;margin-top:4px}
.feature-badge{background:#f1f5f3;color:#335b54;font:600 9px 'DM Mono';padding:1px 5px;border-radius:4px;border:1px solid #d6e2dd}
.use-template-btn{margin-top:10px;background:#edf5ef;border:1px solid #c4ded6;color:#1e5a55;font:600 12px 'DM Sans';padding:8px 12px;border-radius:7px;cursor:pointer;transition:all .15s;text-align:center}
.use-template-btn:hover{background:#1e5a55;color:#fff;border-color:#1e5a55}
.modal-backdrop{position:fixed;inset:0;background:rgba(15,23,42,0.55);backdrop-filter:blur(3px);z-index:999;display:flex;align-items:center;justify-content:center;padding:20px}
.modal-card{background:#fff;border-radius:16px;max-width:860px;width:100%;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 20px 48px rgba(0,0,0,0.22);overflow:hidden;animation:popIn .2s ease-out}
@keyframes popIn{from{opacity:0;transform:scale(0.96)}to{opacity:1;transform:scale(1)}}
.modal-header{padding:22px 26px;border-bottom:1px solid #e5ece8;display:flex;justify-content:space-between;align-items:flex-start}
.modal-close-btn{background:#f1f5f3;border:0;font-size:16px;color:#64748b;width:32px;height:32px;border-radius:50%;cursor:pointer;display:grid;place-items:center;transition:all .15s}
.modal-close-btn:hover{background:#e2e8f0;color:#0f172a}
.modal-body{padding:24px 26px;overflow-y:auto}
.modal-grid{grid-template-columns:repeat(auto-fill,minmax(240px,1fr))}
</style>

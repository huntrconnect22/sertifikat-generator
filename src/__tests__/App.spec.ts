import { describe, it, expect } from 'vitest'

import { mount } from '@vue/test-utils'
import App from '../App.vue'

describe('App', () => {
  it('mounts renders properly', () => {
    const wrapper = mount(App)
    expect(wrapper.text()).toContain('SERTIFIKAT GENERATOR')
    expect(wrapper.text()).toContain('Logo 1 (Kanan Atas)')
    expect(wrapper.text()).toContain('Logo 2 (Kiri Atas)')
    expect(wrapper.text()).toContain('Pilihan warna tema sertifikat')
  })

  it('updates accent color when preset swatch is clicked', async () => {
    const wrapper = mount(App)
    const swatches = wrapper.findAll('.swatch-btn')
    expect(swatches.length).toBeGreaterThan(1)
    const secondSwatch = swatches[1]
    if (secondSwatch) {
      await secondSwatch.trigger('click')
      const certificate = wrapper.find('.certificate')
      expect(certificate.attributes('style')).toContain('#1e3a8a')
    }
  })

  it('renders batch download buttons for multiple certificates', () => {
    const wrapper = mount(App)
    expect(wrapper.text()).toContain('Unduh Semua (1 File PDF)')
    expect(wrapper.text()).toContain('Unduh Arsip ZIP (.zip)')
  })

  it('renders unique certificate registration number and signature upload input', () => {
    const wrapper = mount(App)
    expect(wrapper.text()).toContain('Tanda Tangan 1 (Kanan Bawah)')
    expect(wrapper.text()).toContain('Tanda Tangan 2 (Kiri Bawah)')
    expect(wrapper.text()).toContain('NO. SERTIFIKAT')
  })
})

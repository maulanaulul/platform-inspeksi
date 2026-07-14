import React from 'react'
import { createRoot } from 'react-dom/client'
import PlatformApp from './SidakFatigueApp.jsx'
import './styles.css'
import runtimeStyles from './styles.css?inline'

// Fallback CSS runtime.
// Ini menjaga tampilan tetap rapi ketika file CSS hasil build tidak ikut ter-upload,
// terkena cache lama, atau asset CSS diblokir oleh konfigurasi hosting.
function injectRuntimeStyles(){
  if (typeof document === 'undefined') return
  if (document.getElementById('srgs-runtime-styles')) return

  const style = document.createElement('style')
  style.id = 'srgs-runtime-styles'
  style.setAttribute('data-source', 'src/styles.css')
  style.textContent = runtimeStyles
  document.head.appendChild(style)
}

injectRuntimeStyles()

// Final all-in-one tidak memakai landing 3 kartu lagi.
// Pilihan aplikasi tetap di satu tempat: profile / Ganti Aplikasi.

createRoot(document.getElementById('root')).render(<PlatformApp />)

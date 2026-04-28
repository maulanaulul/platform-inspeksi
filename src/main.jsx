import React from 'react'
import { createRoot } from 'react-dom/client'
import PlatformApp from './SidakFatigueApp.jsx'
import './styles.css'

// Final all-in-one tidak memakai landing 3 kartu lagi.
// Pilihan aplikasi tetap di satu tempat: profile / Ganti Aplikasi.
createRoot(document.getElementById('root')).render(<PlatformApp />)

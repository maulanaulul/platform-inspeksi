export function exportToCsv(filename, rows) {
  if (!rows || rows.length === 0) return alert('Tidak ada data untuk diexport')
  const headers = Object.keys(rows[0])
  const csv = [
    headers.join(','),
    ...rows.map(row => headers.map(h => JSON.stringify(row[h] ?? '')).join(','))
  ].join('\n')
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

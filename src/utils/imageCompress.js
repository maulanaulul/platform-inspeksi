export async function compressImage(file, maxWidth = 1280, quality = 0.72) {
  if (!file || !file.type?.startsWith('image/')) return file
  const imageBitmap = await createImageBitmap(file)
  const ratio = Math.min(1, maxWidth / imageBitmap.width)
  const width = Math.round(imageBitmap.width * ratio)
  const height = Math.round(imageBitmap.height * ratio)
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext('2d')
  ctx.drawImage(imageBitmap, 0, 0, width, height)
  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', quality))
  return new File([blob], file.name.replace(/\.[^.]+$/, '.jpg'), { type: 'image/jpeg' })
}

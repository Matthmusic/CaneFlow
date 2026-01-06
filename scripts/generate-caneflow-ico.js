const toIco = require('to-ico')
const fs = require('fs')
const path = require('path')

async function generateIco() {
  try {
    console.log('Generating caneflow.ico from CaneFlow.png...')

    const pngPath = path.join(__dirname, '../img/CaneFlow.png')
    const icoPath = path.join(__dirname, '../electron/caneflow.ico')

    const pngBuffer = fs.readFileSync(pngPath)
    const icoBuffer = await toIco([pngBuffer], {
      sizes: [16, 24, 32, 48, 64, 128, 256],
      resize: true,
    })

    fs.writeFileSync(icoPath, icoBuffer)
    console.log(`Saved: ${icoPath}`)
  } catch (error) {
    console.error('Error generating ico:', error.message)
  }
}

generateIco()

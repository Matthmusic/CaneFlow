const fs = require('fs')
const path = require('path')
const os = require('os')

const MARGIN = 1.33
const EMBEDDED_CABLES_PATH = path.join(__dirname, 'cables.json')

const NEXUS_CONFIG_PATH = process.env.APPDATA
  ? path.join(process.env.APPDATA, 'nexus', 'config.json')
  : path.join(os.homedir(), 'AppData', 'Roaming', 'nexus', 'config.json')

const CABLE_CATEGORIES = new Set([
  'cables industriels rigides',
  'cables incendie',
  'moyenne tension',
  'fils et cables souples',
  'domestique rigide',
  'cables speciaux',
  'cables alarmes',
  'cables telephoniques',
])

// Normalisation pour la comparaison des noms
function norm(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/,/g, '.')
    .replace(/\s+/g, ' ')
}

// Normalisation pour les sections câbles (sans espaces)
function normSection(s) {
  return String(s || '').trim().toLowerCase().replace(/,/g, '.').replace(/\s+/g, '')
}

function getNexusDataDir() {
  try {
    if (!fs.existsSync(NEXUS_CONFIG_PATH)) return null
    const config = JSON.parse(fs.readFileSync(NEXUS_CONFIG_PATH, 'utf-8'))
    return config.dataDir || null
  } catch {
    return null
  }
}

// Prix matériel brut depuis prices.json NEXUS (filtrés par catégorie câble)
function loadNexusCableItems() {
  const dataDir = getNexusDataDir()
  if (!dataDir) return []
  const pricesPath = path.join(dataDir, 'prices.json')
  try {
    if (!fs.existsSync(pricesPath)) return []
    const db = JSON.parse(fs.readFileSync(pricesPath, 'utf-8'))
    const items = Array.isArray(db.items) ? db.items : []
    return items.filter((item) => CABLE_CATEGORIES.has(norm(item.category || '')))
  } catch {
    return []
  }
}

// Temps et coûts de pose fixes depuis cables.json embarqué
function loadCablesData() {
  try {
    if (!fs.existsSync(EMBEDDED_CABLES_PATH)) return []
    return JSON.parse(fs.readFileSync(EMBEDDED_CABLES_PATH, 'utf-8'))
  } catch {
    return []
  }
}

// Cherche le prix matériel brut dans prices.json NEXUS (type + section)
function matchMaterialPrice(cable, typeCable, items) {
  if (!cable && !typeCable) return undefined
  const normCable = norm(cable)
  const normType = norm(typeCable)
  const fullKey = [normType, normCable].filter(Boolean).join(' ')

  for (const item of items) {
    if (norm(item.name) === fullKey) return item.price
  }
  if (normType && normCable) {
    for (const item of items) {
      const n = norm(item.name)
      if (n.includes(normType) && n.includes(normCable)) return item.price
    }
  }
  if (normCable) {
    for (const item of items) {
      if (norm(item.name).includes(normCable)) return item.price
    }
  }
  return undefined
}

// Cherche le coût de pose dans cables.json par section
function matchPoseEntry(cable, cablesData) {
  const s = normSection(cable)
  if (!s) return undefined
  for (const entry of cablesData) {
    if (normSection(entry.designation) === s) return entry
  }
  return undefined
}

function lookupCablePrices(compositeKeys, margin = MARGIN) {
  const result = {}
  if (!Array.isArray(compositeKeys) || compositeKeys.length === 0) return result

  const nexusItems = loadNexusCableItems()
  const cablesData = loadCablesData()

  for (const key of compositeKeys) {
    const parts = key.split(' | ')
    const cable = parts[0] || ''
    const typeCable = parts[1] || ''

    const materialPrice = matchMaterialPrice(cable, typeCable, nexusItems)
    if (materialPrice === undefined) continue

    const poseEntry = matchPoseEntry(cable, cablesData)
    if (!poseEntry) continue

    const total = materialPrice * margin + poseEntry.pt_pose_value
    if (total > 0) result[key] = total
  }

  return result
}

module.exports = { lookupCablePrices }

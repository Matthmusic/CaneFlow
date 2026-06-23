const fs = require('fs')
const path = require('path')
const os = require('os')

const MARGIN = 1.33
const CABLE_KEY_SEP = ' | '
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

// cables.json est un asset statique embarqué — chargé une seule fois au démarrage
let _cablesData = null

// prices.json NEXUS est sur disque/réseau — pas de cache persistant
// pour permettre les retry si le réseau n'était pas accessible au premier appel
function norm(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/,/g, '.')
    .replace(/\s+/g, ' ')
}

function normSection(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/,/g, '.')
    .replace(/\s+/g, '')
}

// Inline — évite la dépendance externe vers transform.cjs dans le package ASAR
function parseCableKey(key) {
  const idx = key.indexOf(CABLE_KEY_SEP)
  if (idx === -1) return { cable: key, typeCable: '' }
  return { cable: key.slice(0, idx), typeCable: key.slice(idx + CABLE_KEY_SEP.length) }
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

function loadCablesData() {
  if (_cablesData !== null) return _cablesData
  try {
    _cablesData = JSON.parse(fs.readFileSync(EMBEDDED_CABLES_PATH, 'utf-8'))
  } catch {
    _cablesData = []
  }
  return _cablesData
}

function matchMaterialPrice(cable, typeCable, items, exactMap) {
  if (!cable && !typeCable) return undefined
  const normCable = norm(cable)
  const normType = norm(typeCable)
  const fullKey = [normType, normCable].filter(Boolean).join(' ')

  if (exactMap.has(fullKey)) return exactMap.get(fullKey)

  if (normType && normCable) {
    for (const item of items) {
      const n = norm(item.name)
      if (n.includes(normType) && n.includes(normCable)) return item.price
    }
  }

  // Pass section seule uniquement si pas de typeCable
  if (!normType && normCable) {
    for (const item of items) {
      if (norm(item.name).includes(normCable)) return item.price
    }
  }

  return undefined
}

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

  const exactMap = new Map()
  for (const item of nexusItems) {
    const k = norm(item.name)
    if (!exactMap.has(k)) exactMap.set(k, item.price)
  }

  for (const key of compositeKeys) {
    const { cable, typeCable } = parseCableKey(key)

    const materialPrice = matchMaterialPrice(cable, typeCable, nexusItems, exactMap)
    if (materialPrice === undefined) continue

    const poseEntry = matchPoseEntry(cable, cablesData)
    if (!poseEntry) continue

    const total = materialPrice * margin + poseEntry.pt_pose_value
    if (total > 0) result[key] = total
  }

  return result
}

module.exports = { lookupCablePrices }

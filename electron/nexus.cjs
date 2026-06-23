const fs = require('fs')
const path = require('path')
const os = require('os')

const MARGIN = 1.33
const CABLE_KEY_SEP = ' | '
const EMBEDDED_CABLES_PATH = path.join(__dirname, 'cables.json')

const NEXUS_PRICES_URL = 'https://pub-12c18956bbe54a0888d82fd2921c8aa4.r2.dev/prices.json'
const NEXUS_FALLBACK_DIR = 'Z:\\F - UTILITAIRES\\NEXUS'
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

let _cablesData = null

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

function parseCableKey(key) {
  const idx = key.indexOf(CABLE_KEY_SEP)
  if (idx === -1) return { cable: key, typeCable: '' }
  return { cable: key.slice(0, idx), typeCable: key.slice(idx + CABLE_KEY_SEP.length) }
}

function filterCableItems(db) {
  const items = Array.isArray(db.items) ? db.items : []
  return items.filter((item) => CABLE_CATEGORIES.has(norm(item.category || '')))
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

async function loadNexusCableItems() {
  // 1. R2 online (fonctionne partout, même hors réseau local)
  try {
    const res = await fetch(NEXUS_PRICES_URL, { signal: AbortSignal.timeout(5000) })
    if (res.ok) return filterCableItems(await res.json())
  } catch {}

  // 2. Fallback Z: réseau local
  const candidates = []
  const configDir = getNexusDataDir()
  if (configDir) candidates.push(path.join(configDir, 'prices.json'))
  candidates.push(path.join(NEXUS_FALLBACK_DIR, 'prices.json'))

  for (const pricesPath of candidates) {
    try {
      if (!fs.existsSync(pricesPath)) continue
      return filterCableItems(JSON.parse(fs.readFileSync(pricesPath, 'utf-8')))
    } catch {}
  }

  return []
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

async function lookupCablePrices(compositeKeys, margin = MARGIN) {
  const result = {}
  if (!Array.isArray(compositeKeys) || compositeKeys.length === 0) return result

  const nexusItems = await loadNexusCableItems()
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

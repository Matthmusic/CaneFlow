const fs = require('fs')
const path = require('path')
const os = require('os')
const { parseCableKey } = require('./transform.cjs')

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

// Module-scope caches — populated on first call, never re-read at runtime
let _nexusDataDir = null
let _nexusDataDirLoaded = false
let _nexusItems = null
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

function getNexusDataDir() {
  if (_nexusDataDirLoaded) return _nexusDataDir
  _nexusDataDirLoaded = true
  try {
    if (!fs.existsSync(NEXUS_CONFIG_PATH)) return null
    const config = JSON.parse(fs.readFileSync(NEXUS_CONFIG_PATH, 'utf-8'))
    _nexusDataDir = config.dataDir || null
  } catch {
    _nexusDataDir = null
  }
  return _nexusDataDir
}

function loadNexusCableItems() {
  if (_nexusItems !== null) return _nexusItems
  const dataDir = getNexusDataDir()
  if (!dataDir) { _nexusItems = []; return _nexusItems }
  const pricesPath = path.join(dataDir, 'prices.json')
  try {
    if (!fs.existsSync(pricesPath)) { _nexusItems = []; return _nexusItems }
    const db = JSON.parse(fs.readFileSync(pricesPath, 'utf-8'))
    const items = Array.isArray(db.items) ? db.items : []
    _nexusItems = items.filter((item) => CABLE_CATEGORIES.has(norm(item.category || '')))
  } catch {
    _nexusItems = []
  }
  return _nexusItems
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

  // Pass 1 — exact match via pre-built Map (O(1))
  if (exactMap.has(fullKey)) return exactMap.get(fullKey)

  // Pass 2 — partial: name contains both type and section
  if (normType && normCable) {
    for (const item of items) {
      const n = norm(item.name)
      if (n.includes(normType) && n.includes(normCable)) return item.price
    }
  }

  // Pass 3 — section only, only when there is no type to match against
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

  // Pre-build exact-match Map once for all keys
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

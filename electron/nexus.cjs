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
  'cables speciaux',
  'cables alarmes',
  'cables telephoniques',
  'cables coaxiaux et reseaux',
  'cables haute temperature',
  'cuivre nu',
  'domotique',
  'industrie et speciaux',
  'installations de securite',
])

const CACHE_TTL_MS = 5 * 60 * 1000

let _cablesData = null
let _nexusItemsCache = null
let _nexusItemsCachedAt = 0
let _nexusExactMap = null
let _nexusNormNames = null
let _nexusSource = 'none'

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
  return norm(s).replace(/\s/g, '')
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
  const now = Date.now()
  if (_nexusItemsCache !== null && now - _nexusItemsCachedAt < CACHE_TTL_MS) {
    return _nexusItemsCache
  }

  let items = null

  // 1. R2 online (fonctionne partout, même hors réseau local)
  try {
    const res = await fetch(NEXUS_PRICES_URL, { signal: AbortSignal.timeout(5000) })
    if (res.ok) items = filterCableItems(await res.json())
  } catch {}

  // 2. Fallback Z: réseau local puis config NEXUS locale (aussi si R2 a renvoyé 0 article)
  if (!items?.length) {
    const configDir = getNexusDataDir()
    for (const dir of [configDir, NEXUS_FALLBACK_DIR].filter(Boolean)) {
      const pricesPath = path.join(dir, 'prices.json')
      try {
        if (!fs.existsSync(pricesPath)) continue
        items = filterCableItems(JSON.parse(fs.readFileSync(pricesPath, 'utf-8')))
        break
      } catch {}
    }
    _nexusSource = items?.length ? 'fallback' : 'none'
  } else {
    _nexusSource = 'r2'
  }

  _nexusItemsCache = items ?? []
  _nexusItemsCachedAt = now
  _nexusExactMap = null
  _nexusNormNames = null
  return _nexusItemsCache
}

function compact(s) {
  return norm(s).replace(/\s/g, '')
}

function getExactMap(nexusItems) {
  if (_nexusExactMap !== null) return _nexusExactMap
  const map = new Map()
  for (const item of nexusItems) {
    const k = compact(item.name)
    if (!map.has(k)) map.set(k, item.price)
  }
  _nexusExactMap = map
  return map
}

function getNormNames(nexusItems) {
  if (_nexusNormNames !== null) return _nexusNormNames
  _nexusNormNames = nexusItems.map((item) => ({ n: compact(item.name), price: item.price }))
  return _nexusNormNames
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

function stripParens(s) {
  return String(s || '').replace(/\(.*?\)/g, '').trim()
}

function matchMaterialPrice(cable, typeCable, exactMap, normNames) {
  if (!cable && !typeCable) return undefined
  const nc = compact(cable)
  const nt = compact(stripParens(typeCable))
  const fullKey = nc && nt ? nt + nc : nc || nt

  if (exactMap.has(fullKey)) return exactMap.get(fullKey)

  for (const { n, price } of normNames) {
    if (nc && n.includes(nc) && (!nt || n.includes(nt))) return price
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
  const prices = {}
  const details = {}
  if (!Array.isArray(compositeKeys) || compositeKeys.length === 0) return { prices, details, source: _nexusSource }

  const nexusItems = await loadNexusCableItems()
  const cablesData = loadCablesData()
  const exactMap = getExactMap(nexusItems)
  const normNames = getNormNames(nexusItems)

  for (const key of compositeKeys) {
    const { cable, typeCable } = parseCableKey(key)
    const mat = matchMaterialPrice(cable, typeCable, exactMap, normNames)
    if (mat === undefined) continue
    const poseEntry = matchPoseEntry(cable, cablesData)
    if (!poseEntry) continue
    const total = mat * margin + poseEntry.pt_pose_value
    if (total <= 0) continue
    prices[key] = total
    details[key] = { mat, margin, tpsPose: poseEntry.tps_pose, puPose: poseEntry.pu_pose, ptPose: poseEntry.pt_pose_value }
  }

  return { prices, details, source: _nexusSource }
}

module.exports = { lookupCablePrices }

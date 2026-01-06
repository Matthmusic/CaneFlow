const DEFAULT_UNIT = 'ml'

const DEFAULT_COLUMN_INDEX = {
  amont: 0,
  repere: 1,
  longueur: 2,
  cable: 3,
  neutre: 4,
  pe: 5,
  typeCable: 6,
}

const HEADER_ALIASES = new Map([
  ['amont', 'amont'],
  ['repere', 'repere'],
  ['longueur', 'longueur'],
  ['cable', 'cable'],
  ['neutre', 'neutre'],
  ['pe ou pen', 'pe'],
  ['type de cable', 'typeCable'],
])

function normalizeText(value) {
  if (value === null || value === undefined) {
    return ''
  }
  return String(value).trim()
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === '') {
    return ''
  }
  if (typeof value === 'number') {
    return value
  }
  const normalized = String(value).replace(',', '.').trim()
  const numberValue = Number(normalized)
  return Number.isFinite(numberValue) ? numberValue : value
}

function normalizeHeaderLabel(value) {
  if (value === null || value === undefined) {
    return ''
  }
  return String(value)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
}

function hasHeaderRow(row) {
  if (!row || row.length === 0) {
    return false
  }
  const normalized = row.map(normalizeHeaderLabel)
  return normalized.includes('amont') && normalized.includes('repere')
}

function buildColumnIndex(headerRow) {
  const indices = { ...DEFAULT_COLUMN_INDEX }
  if (!headerRow) {
    return indices
  }
  headerRow.forEach((value, index) => {
    const normalized = normalizeHeaderLabel(value)
    const key = HEADER_ALIASES.get(normalized)
    if (key) {
      indices[key] = index
    }
  })
  return indices
}

function rowFromArray(row, indices) {
  return {
    amont: normalizeText(row[indices.amont]),
    repere: normalizeText(row[indices.repere]),
    longueur: normalizeNumber(row[indices.longueur]),
    cable: normalizeText(row[indices.cable]),
    neutre: normalizeText(row[indices.neutre]),
    pe: normalizeText(row[indices.pe]),
    typeCable: normalizeText(row[indices.typeCable]),
  }
}

function buildCableKey(row) {
  const cable = normalizeText(row.cable)
  const typeCable = normalizeText(row.typeCable)
  if (cable && typeCable) {
    return `${cable} | ${typeCable}`
  }
  return cable || typeCable || ''
}

function buildTitle(row) {
  let title = `Fourniture, pose et raccordement \"${row.repere}\" - en c\u00e2ble : ${row.cable}`

  const extraParts = []
  if (row.neutre) {
    extraParts.push(`${row.neutre}`)
  }
  if (row.pe) {
    extraParts.push(`PE ${row.pe}`)
  }
  if (extraParts.length > 0) {
    title += ` + ${extraParts.join(' + ')}`
  }
  if (row.typeCable) {
    title += ` - ${row.typeCable}`
  }

  return title
}

function isEmptyRow(row) {
  return !(row.repere || row.cable || row.neutre || row.pe || row.typeCable || row.longueur !== '')
}

function parseSheetRows(sheetRows) {
  if (!Array.isArray(sheetRows) || sheetRows.length === 0) {
    return []
  }

  const headerRow = sheetRows[0]
  const startIndex = hasHeaderRow(headerRow) ? 1 : 0
  const columnIndex = startIndex === 1 ? buildColumnIndex(headerRow) : { ...DEFAULT_COLUMN_INDEX }
  const rows = []

  for (let i = startIndex; i < sheetRows.length; i += 1) {
    const rowArray = sheetRows[i]
    if (!rowArray || rowArray.length === 0) {
      continue
    }

    const row = rowFromArray(rowArray, columnIndex)
    if (isEmptyRow(row)) {
      continue
    }
    if (row.longueur === 0) {
      continue
    }

    rows.push(row)
  }

  return rows
}

function buildPreviewRows(sheetRows, options = {}) {
  const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null
  const rows = parseSheetRows(sheetRows)
  const preview = []
  let lastPercent = -1

  rows.forEach((row, index) => {
    preview.push({
      lineNumber: index + 1,
      title: buildTitle(row),
      quantity: row.longueur,
      repere: row.repere,
      typeCable: buildCableKey(row),
    })

    if (onProgress) {
      const percent = Math.round(((index + 1) / rows.length) * 100)
      if (percent !== lastPercent) {
        lastPercent = percent
        onProgress({ current: index + 1, total: rows.length, percent })
      }
    }
  })

  return preview
}

function mapSheetRows(sheetRows, options = {}) {
  const defaultUnitPrice = options.defaultUnitPrice ?? ''
  const defaultTva = options.defaultTva ?? ''
  const defaultUnit = options.defaultUnit ?? DEFAULT_UNIT
  const includeHeaders = options.includeHeaders ?? true
  const unitPrices = Array.isArray(options.unitPrices) ? options.unitPrices : null
  const unitPricesByType =
    options.unitPricesByType && typeof options.unitPricesByType === 'object' ? options.unitPricesByType : null
  const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null

  const outputRows = []
  if (includeHeaders) {
    outputRows.push(['N\u00b0', 'Titre', 'Unit\u00e9', 'Quantit\u00e9', 'Prix unitaire', '', 'TVA', 'Descriptif'])
  }

  const rows = parseSheetRows(sheetRows)
  if (rows.length === 0) {
    return outputRows
  }

  let lineNumber = 1
  let lastPercent = -1
  rows.forEach((row, index) => {
    const typeKey = buildCableKey(row)
    const unitPriceByType =
      unitPricesByType && Object.prototype.hasOwnProperty.call(unitPricesByType, typeKey)
        ? unitPricesByType[typeKey]
        : undefined
    const unitPriceByIndex = unitPrices ? unitPrices[index] : undefined
    const unitPrice = normalizeNumber(unitPriceByType ?? unitPriceByIndex ?? row.unitPrice ?? defaultUnitPrice)
    const tva = normalizeNumber(row.tva ?? defaultTva)

    outputRows.push([
      lineNumber,
      buildTitle(row),
      defaultUnit,
      row.longueur,
      unitPrice,
      '',
      tva,
      '',
    ])

    lineNumber += 1

    if (onProgress) {
      const percent = Math.round(((index + 1) / rows.length) * 100)
      if (percent !== lastPercent) {
        lastPercent = percent
        onProgress({ current: index + 1, total: rows.length, percent })
      }
    }
  })

  return outputRows
}

module.exports = {
  buildTitle,
  buildPreviewRows,
  mapSheetRows,
  normalizeNumber,
  normalizeText,
}

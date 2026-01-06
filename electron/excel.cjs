const ExcelJS = require('exceljs')
const fs = require('fs')
const os = require('os')
const path = require('path')
const { spawn } = require('child_process')

const LOG_PREFIX = '[CaneFlow]'

function isXlsFile(filePath) {
  return path.extname(filePath).toLowerCase() === '.xls'
}

function createTempXlsxPath() {
  const stamp = `${Date.now()}-${Math.random().toString(16).slice(2)}`
  return path.join(os.tmpdir(), `caneflow-${stamp}.xlsx`)
}

function convertXlsToXlsx(inputPath, outputPath) {
  console.log(`${LOG_PREFIX} convertXlsToXlsx start`, { inputPath, outputPath })
  return new Promise((resolve, reject) => {
    const script = [
      '& {',
      '$ErrorActionPreference = "Stop";',
      '$inputPath = $env:CANEFLOW_INPUT_PATH;',
      '$outputPath = $env:CANEFLOW_OUTPUT_PATH;',
      'if (-not (Test-Path -LiteralPath $inputPath)) { throw "Fichier source introuvable: $inputPath"; }',
      'try {',
      '  $excel = New-Object -ComObject Excel.Application -ErrorAction Stop;',
      '} catch {',
      '  throw "Microsoft Excel n\'est pas installé ou accessible. Erreur: $($_.Exception.Message)";',
      '}',
      '$excel.Visible = $false;',
      '$excel.DisplayAlerts = $false;',
      '$workbook = $null;',
      'try {',
      '  $workbook = $excel.Workbooks.Open($inputPath, $null, $true);',
      '  if (-not $workbook) { throw "Impossible d\'ouvrir le fichier. Il est peut-être corrompu ou protégé par mot de passe."; }',
      '  $workbook.SaveAs($outputPath, 51);',
      '  $workbook.Close($false);',
      '} catch {',
      '  $errorMsg = $_.Exception.Message;',
      '  if ($errorMsg -match "password|mot de passe") {',
      '    throw "Le fichier est protégé par mot de passe.";',
      '  } elseif ($errorMsg -match "permission|access denied|accès refusé") {',
      '    throw "Accès refusé. Vérifiez les permissions du fichier.";',
      '  } elseif ($errorMsg -match "format|corrupt") {',
      '    throw "Le fichier est corrompu ou dans un format non supporté.";',
      '  } else {',
      '    throw "Erreur lors de la conversion: $errorMsg";',
      '  }',
      '} finally {',
      '  if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null; }',
      '  $excel.Quit();',
      '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;',
      '  [GC]::Collect();',
      '  [GC]::WaitForPendingFinalizers();',
      '}',
      'if (-not (Test-Path -LiteralPath $outputPath)) { throw "La conversion a échoué: fichier de sortie non créé."; }',
      '}',
    ].join(' ')

    const child = spawn(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script],
      {
        windowsHide: true,
        env: {
          ...process.env,
          CANEFLOW_INPUT_PATH: inputPath,
          CANEFLOW_OUTPUT_PATH: outputPath,
        },
      },
    )

    let stderr = ''
    let stdout = ''

    child.stdout.on('data', (data) => {
      stdout += data.toString()
    })

    child.stderr.on('data', (data) => {
      stderr += data.toString()
    })

    child.on('error', (error) => {
      console.error(`${LOG_PREFIX} spawn error`, { error: error.message })
      reject(new Error(`Erreur PowerShell: ${error.message}`))
    })

    child.on('close', (code) => {
      if (code === 0) {
        console.log(`${LOG_PREFIX} convertXlsToXlsx done`, { outputPath })
        resolve()
        return
      }

      const errorOutput = stderr.trim() || stdout.trim()
      console.error(`${LOG_PREFIX} convertXlsToXlsx failed`, {
        code,
        stderr: stderr.trim(),
        stdout: stdout.trim(),
        inputPath,
        outputPath
      })

      let message = 'Échec de conversion du fichier .xls'
      if (errorOutput) {
        message = errorOutput
      } else if (code === 1) {
        message = 'Microsoft Excel est requis pour ouvrir les fichiers .xls. Veuillez vérifier qu\'Excel est bien installé.'
      }

      reject(new Error(message))
    })
  })
}

function extractCellValue(cellValue) {
  if (cellValue === null || cellValue === undefined) {
    return ''
  }
  if (cellValue instanceof Date) {
    return cellValue
  }
  if (typeof cellValue === 'object') {
    if (cellValue.result !== undefined) {
      return cellValue.result ?? ''
    }
    if (cellValue.text !== undefined) {
      return cellValue.text ?? ''
    }
    if (Array.isArray(cellValue.richText)) {
      return cellValue.richText.map((part) => part.text).join('')
    }
    if (cellValue.hyperlink && cellValue.text !== undefined) {
      return cellValue.text ?? ''
    }
  }
  return cellValue
}

function isRowEmpty(row) {
  return !row.some((value) => value !== '' && value !== null && value !== undefined)
}

function buildRowValues(row, columnCount) {
  const values = []
  for (let col = 1; col <= columnCount; col += 1) {
    const cell = row.getCell(col)
    const value = cell ? extractCellValue(cell.value) : ''
    values.push(value ?? '')
  }
  return values
}

async function readSheetRows(inputPath, sheetName) {
  console.log(`${LOG_PREFIX} readSheetRows start`, { inputPath, sheetName: sheetName || '' })
  let sourcePath = inputPath
  let tempPath = ''
  if (isXlsFile(inputPath)) {
    if (!fs.existsSync(inputPath)) {
      throw new Error(`Input file not found: ${inputPath}`)
    }
    tempPath = createTempXlsxPath()
    console.log(`${LOG_PREFIX} converting xls to xlsx`, { inputPath, tempPath })
    await convertXlsToXlsx(inputPath, tempPath)
    if (!fs.existsSync(tempPath)) {
      throw new Error('Conversion .xls vers .xlsx impossible. Verifie Excel ou les droits d ecriture.')
    }
    console.log(`${LOG_PREFIX} xls conversion complete`, { tempPath })
    sourcePath = tempPath
  }

  console.log(`${LOG_PREFIX} loading workbook`, { sourcePath })
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.readFile(sourcePath)
  } catch (error) {
    console.error(`${LOG_PREFIX} failed to load workbook`, { error: error.message })
    if (tempPath && fs.existsSync(tempPath)) {
      fs.unlinkSync(tempPath)
    }
    throw error
  }
  console.log(`${LOG_PREFIX} workbook loaded`, { sheetCount: workbook.worksheets.length })

  const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0]
  if (!worksheet) {
    console.error(`${LOG_PREFIX} sheet not found`, { sheetName })
    if (tempPath && fs.existsSync(tempPath)) {
      fs.unlinkSync(tempPath)
    }
    throw new Error(`Sheet not found: ${sheetName}`)
  }
  console.log(`${LOG_PREFIX} processing sheet`, { name: worksheet.name })

  const columnCount = worksheet.actualColumnCount || worksheet.columnCount || 0
  const totalRows = worksheet.actualRowCount || worksheet.rowCount || 0
  console.log(`${LOG_PREFIX} extracting rows`, { columnCount, totalRows })
  const rows = []
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    rows.push(buildRowValues(row, columnCount))
  })

  while (rows.length && isRowEmpty(rows[0])) {
    rows.shift()
  }
  while (rows.length && isRowEmpty(rows[rows.length - 1])) {
    rows.pop()
  }

  if (tempPath && fs.existsSync(tempPath)) {
    fs.unlinkSync(tempPath)
  }
  console.log(`${LOG_PREFIX} readSheetRows done`, { rowCount: rows.length })
  return rows
}

async function writeSheetRows(outputPath, sheetName, rows) {
  console.log(`${LOG_PREFIX} writeSheetRows start`, { outputPath, sheetName, rowCount: rows.length })
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet(sheetName)
  console.log(`${LOG_PREFIX} adding rows to worksheet`)
  rows.forEach((row) => {
    worksheet.addRow(row)
  })
  console.log(`${LOG_PREFIX} saving workbook to file`, { outputPath })
  await workbook.xlsx.writeFile(outputPath)
  console.log(`${LOG_PREFIX} writeSheetRows done`, { outputPath })
}

module.exports = {
  readSheetRows,
  writeSheetRows,
}

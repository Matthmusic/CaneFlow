import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import { CheckCircle2, FileSpreadsheet, FolderOpen, Info, Maximize2, Minus, RefreshCw, X as Close } from 'lucide-react'
import './App.css'
import logo from '../public/logo.svg'

type UpdateStatus = {
  state: 'idle' | 'available' | 'downloading' | 'downloaded' | 'error'
  version?: string
  progress?: number
  message?: string
}

type ConversionResult = {
  outputPath: string
  rowCount: number
}

type PreviewRow = {
  lineNumber: number
  title: string
  quantity: string | number
  repere: string
  typeCable: string
}

const hasApi = () => typeof window !== 'undefined' && typeof (window as any).api !== 'undefined'

function App() {
  const currentVersion = import.meta.env.VITE_APP_VERSION || '0.0.0'
  const [inputPath, setInputPath] = useState('')
  const [outputPath, setOutputPath] = useState('')
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([])
  const [cableTypes, setCableTypes] = useState<string[]>([])
  const [typePrices, setTypePrices] = useState<Record<string, string>>({})
  const [unitPrices, setUnitPrices] = useState<string[]>([])
  const [priceMode, setPriceMode] = useState<'perCable' | 'perLine'>('perCable')
  const [defaultUnitPrice, setDefaultUnitPrice] = useState('0')
  const [tva, setTva] = useState('0')
  const [status, setStatus] = useState('En attente de selection.')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [loadingPreview, setLoadingPreview] = useState(false)
  const [lastExport, setLastExport] = useState<ConversionResult | null>(null)
  const [showInfo, setShowInfo] = useState(false)
  const [toast, setToast] = useState<{ message: string; type: 'info' | 'error' | 'update' } | null>(null)
  const [toastTimeout, setToastTimeout] = useState<ReturnType<typeof setTimeout> | null>(null)
  const [updateStatus, setUpdateStatus] = useState<UpdateStatus>({ state: 'idle' })
  const [isDragging, setIsDragging] = useState(false)
  const noDragStyle: CSSProperties = { WebkitAppRegion: 'no-drag' }

  const outputLabel = useMemo(() => {
    if (outputPath) return outputPath
    if (lastExport?.outputPath) return lastExport.outputPath
    return ''
  }, [outputPath, lastExport])

  const typeCounts = useMemo(() => {
    const counts: Record<string, number> = {}
    previewRows.forEach((row) => {
      const typeKey = row.typeCable || ''
      counts[typeKey] = (counts[typeKey] || 0) + 1
    })
    return counts
  }, [previewRows])

  const formatCableType = (typeCable: string) => (typeCable ? typeCable : 'Cable/type vide')

  const showToast = (message: string, type: 'info' | 'error' | 'update' = 'info') => {
    if (toastTimeout) clearTimeout(toastTimeout)
    setToast({ message, type })
    const t = setTimeout(() => setToast(null), 3200)
    setToastTimeout(t)
  }

  useEffect(() => {
    if (!hasApi()) return
    const unsubscribe = window.api.onUpdateEvent((data: any) => {
      switch (data?.type) {
        case 'available':
          setUpdateStatus({ state: 'available', version: data.info?.version })
          showToast(`Mise a jour disponible (${data.info?.version}).`, 'update')
          break
        case 'downloaded':
          setUpdateStatus({ state: 'downloaded', version: data.info?.version })
          showToast('Mise a jour telechargee. Clique pour installer.', 'update')
          break
        case 'progress':
          setUpdateStatus((prev) => ({
            state: 'downloading',
            version: prev.version || data.progress?.version,
            progress: Math.round(data.progress?.percent || 0),
          }))
          break
        case 'error':
          setUpdateStatus({ state: 'error', message: data.message })
          showToast('Erreur de mise a jour.', 'error')
          break
        case 'not-available':
          setUpdateStatus({ state: 'idle' })
          break
        default:
          break
      }
    })
    const unsubscribeLog = window.api.onAppLog((logData) => {
      const style = 'color: #636EFF; font-weight: bold;'
      const prefix = '[CaneFlow]'
      if (logData.level === 'error') {
        console.error(`%c${prefix}`, style, logData.message, logData.data)
      } else if (logData.level === 'warn') {
        console.warn(`%c${prefix}`, style, logData.message, logData.data)
      } else {
        console.log(`%c${prefix}`, style, logData.message, logData.data)
      }
    })
    window.api.checkUpdates()
    return () => {
      if (unsubscribe) unsubscribe()
      if (unsubscribeLog) unsubscribeLog()
    }
  }, [])

  const shortenPath = (value: string) => {
    if (!value) return ''
    if (value.length <= 70) return value
    return `${value.slice(0, 32)} ... ${value.slice(-28)}`
  }

  const applyDefaultPrices = () => {
    if (priceMode === 'perLine') {
      if (previewRows.length === 0) return
      setUnitPrices(previewRows.map(() => defaultUnitPrice))
      return
    }

    if (cableTypes.length === 0) return
    setTypePrices((prev) => {
      const next: Record<string, string> = { ...prev }
      cableTypes.forEach((type) => {
        next[type] = defaultUnitPrice
      })
      return next
    })
  }

  const updateTypePrice = (typeCable: string, value: string) => {
    setTypePrices((prev) => ({ ...prev, [typeCable]: value }))
  }

  const updateUnitPriceAt = (index: number, value: string) => {
    setUnitPrices((prev) => {
      const next = [...prev]
      next[index] = value
      return next
    })
  }

  const loadPreview = async (selectedPath: string) => {
    if (!hasApi()) return
    setLoadingPreview(true)
    setError('')
    setStatus('Chargement des lignes...')
    console.log('[CaneFlow UI] preview start', { inputPath: selectedPath })
    try {
      const rows = await window.api.previewRows({ inputPath: selectedPath })
      const typeOrder: string[] = []
      const seen = new Set<string>()
      rows.forEach((row) => {
        const typeKey = row.typeCable || ''
        if (!seen.has(typeKey)) {
          seen.add(typeKey)
          typeOrder.push(typeKey)
        }
      })
      typeOrder.sort((a, b) => {
        if (a === '' && b !== '') return 1
        if (b === '' && a !== '') return -1
        return a.localeCompare(b)
      })
      setPreviewRows(rows)
      setUnitPrices((prev) => rows.map((_, index) => prev[index] ?? defaultUnitPrice))
      setCableTypes(typeOrder)
      setTypePrices((prev) => {
        const next: Record<string, string> = {}
        typeOrder.forEach((typeKey) => {
          next[typeKey] = prev[typeKey] ?? defaultUnitPrice
        })
        return next
      })
      setStatus(rows.length ? `Lignes chargees (${rows.length}).` : 'Aucune ligne detectee.')
      console.log('[CaneFlow UI] preview done', { count: rows.length })
    } catch (err) {
      console.error('previewRows', err)
      const errorMessage = err instanceof Error ? err.message : 'Impossible de lire le fichier Excel.'
      setError(errorMessage)
      setStatus('')
      setPreviewRows([])
      setCableTypes([])
      setTypePrices({})
      setUnitPrices([])
    } finally {
      setLoadingPreview(false)
    }
  }

  const pickInput = async () => {
    setError('')
    if (!hasApi()) {
      setError("Lance l'app Electron (npm run electron:dev) pour ouvrir le selecteur.")
      return
    }
    try {
      const selected = await window.api.pickInputFile()
      if (!selected) return
      setInputPath(selected)
      setOutputPath('')
      setLastExport(null)
      setStatus('Fichier charge. Pret a convertir.')
      await loadPreview(selected)
    } catch (err) {
      console.error('pickInput', err)
      setError('Impossible de selectionner le fichier.')
      setStatus('Erreur de selection.')
    }
  }

  const pickOutput = async () => {
    setError('')
    if (!hasApi()) return
    try {
      const selected = await window.api.pickOutputFile(inputPath)
      if (!selected) return
      setOutputPath(selected)
    } catch (err) {
      console.error('pickOutput', err)
      setError('Impossible de selectionner le fichier de sortie.')
    }
  }

  const convert = async () => {
    setError('')
    if (!inputPath) {
      setError('Selectionne un fichier Excel en entree.')
      return
    }
    if (!hasApi()) {
      setError('Conversion disponible uniquement dans l app Electron.')
      return
    }

    // Si pas de dossier de sortie, demander à l'utilisateur d'en choisir un
    let finalOutputPath = outputPath
    if (!finalOutputPath) {
      const selected = await window.api.pickOutputFile(inputPath)
      if (!selected) {
        // L'utilisateur a annulé la sélection
        return
      }
      finalOutputPath = selected
      setOutputPath(selected)
    }

    setBusy(true)
    setStatus('Conversion en cours...')
    try {
      const payload: {
        inputPath: string
        outputPath?: string
        unitPrice?: string
        unitPrices?: string[]
        unitPricesByType?: Record<string, string>
        tva?: string
        includeHeaders: boolean
      } = {
        inputPath,
        outputPath: finalOutputPath || undefined,
        unitPrice: defaultUnitPrice || '',
        tva: tva || '',
        includeHeaders: true,
      }

      if (priceMode === 'perLine') {
        payload.unitPrices = unitPrices
      } else {
        payload.unitPricesByType = typePrices
      }

      const result = await window.api.convertFile(payload)
      setLastExport(result)
      setOutputPath(result.outputPath)
      setStatus(`Conversion terminee (${result.rowCount} lignes).`)
      showToast('Export Multidoc genere.', 'info')
    } catch (err) {
      console.error('convert', err)
      const errorMessage = err instanceof Error ? err.message : String(err)
      setError(errorMessage)
      setStatus('')
      showToast('Erreur de conversion.', 'error')
    } finally {
      setBusy(false)
    }
  }

  const revealExport = async () => {
    if (!hasApi() || !outputLabel) return
    await window.api.revealPath(outputLabel)
  }

  const downloadUpdate = async () => {
    if (!hasApi()) return
    await window.api.downloadUpdate()
    setUpdateStatus((prev) => ({ ...prev, state: 'downloading' }))
  }

  const installUpdate = async () => {
    if (!hasApi()) return
    await window.api.installUpdate()
  }

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(true)
  }

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
  }

  const handleDrop = async (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    if (busy || !hasApi()) return

    const files = Array.from(e.dataTransfer.files)
    const excelFile = files.find((file) => {
      const ext = file.name.toLowerCase()
      return ext.endsWith('.xls') || ext.endsWith('.xlsx')
    })

    if (excelFile) {
      setError('')
      const filePath = (excelFile as any).path
      if (filePath) {
        setInputPath(filePath)
        setOutputPath('')
        setLastExport(null)
        setStatus('Fichier charge. Pret a convertir.')
        await loadPreview(filePath)
      }
    } else {
      setError('Veuillez glisser un fichier Excel (.xls ou .xlsx)')
    }
  }

  return (
    <div className="app-shell">
      <div className="titlebar">
        <div className="window-title">
          <img src={logo} alt="CaneFlow" className="title-logo" />
          <span className="window-title-text">CaneFlow</span>
        </div>
        <div className="window-controls">
          <button className="btn-icon info" aria-label="Infos" onClick={() => setShowInfo(true)}>
            <Info size={14} />
          </button>
          <button className="btn-icon" aria-label="Minimiser" onClick={() => hasApi() && window.api.windowMinimize()}>
            <Minus size={14} />
          </button>
          <button className="btn-icon" aria-label="Agrandir" onClick={() => hasApi() && window.api.windowToggleMaximize()}>
            <Maximize2 size={14} />
          </button>
          <button className="btn-icon close" aria-label="Fermer" onClick={() => hasApi() && window.api.windowClose()}>
            <Close size={14} strokeWidth={2.2} />
          </button>
        </div>
      </div>

      <div className={`toast ${toast ? 'visible' : ''} ${toast?.type === 'error' ? 'error' : ''} ${toast?.type === 'update' ? 'update' : ''}`}>
        {toast?.message}
      </div>
      {showInfo ? (
        <div className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal">
            <div className="modal-head">
              <div>
                <p className="eyebrow subtle">Infos</p>
                <h3>Marche a suivre</h3>
              </div>
              <button className="btn-icon close" aria-label="Fermer" onClick={() => setShowInfo(false)}>
                <Close size={14} />
              </button>
            </div>

            <div className="modal-body">
              <p className="modal-text">
                Exporte le carnet de cables Caneco avec les colonnes suivantes, puis charge le fichier dans CaneFlow.
              </p>
              <div className="info-table caneco">
                <div className="info-row header">
                  <div>Colonnes Caneco</div>
                </div>
                <div className="info-row caneco-columns">
                  <span>Amont</span>
                  <span>Repere</span>
                  <span>Longueur</span>
                  <span>Cable</span>
                  <span>Neutre</span>
                  <span>PE ou PEN</span>
                  <span>Type de cable</span>
                </div>
              </div>

              <div className="modal-section">
                <p className="label">Etapes</p>
                <ol className="modal-list">
                  <li>Cliquer sur "Choisir un Excel", puis "Choisir la sortie".</li>
                  <li>Etape 2: saisir les prix unitaires par ligne ou par cable + type, et la TVA globale.</li>
                  <li>Cliquer sur "Convertir" pour generer l Excel Multidoc.</li>
                </ol>
              </div>

              <div className="modal-section">
                <p className="label">Reglages Multidoc</p>
                <p className="modal-text">Dans Multidoc, regler les numeros de colonnes comme suit :</p>
                <div className="info-table mapping">
                  <div className="info-row header two-col">
                    <span>Champ</span>
                    <span>Numero</span>
                  </div>
                  <div className="info-row">
                    <span>Numeros</span>
                    <span>1</span>
                  </div>
                  <div className="info-row">
                    <span>Titres</span>
                    <span>2</span>
                  </div>
                  <div className="info-row">
                    <span>Unites</span>
                    <span>3</span>
                  </div>
                  <div className="info-row">
                    <span>Quantites</span>
                    <span>4</span>
                  </div>
                  <div className="info-row">
                    <span>Prix unitaires</span>
                    <span>5</span>
                  </div>
                  <div className="info-row">
                    <span>Colonne vide</span>
                    <span>6</span>
                  </div>
                  <div className="info-row">
                    <span>TVA</span>
                    <span>7</span>
                  </div>
                  <div className="info-row">
                    <span>Descriptif</span>
                    <span>8</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : null}
      <div className="bg-grid" />
      <div className="bg-glow" />
      <div className="bg-hex" />

      <div className="content">
        <header className="hero">
          <div className="hero-left">
            <div className="logo-wrap">
              <img src={logo} alt="CaneFlow" className="logo" />
              <div className="logo-text">
                <p className="eyebrow">CaneFlow</p>
                <p className="subtext">Conversion rapide vers Multidoc</p>
              </div>
            </div>
            <h1>Transforme un carnet de cables Caneco en Excel Multidoc, en un clic.</h1>
            <p className="lede">
              Importe ton fichier Caneco, controle le tarif global, puis exporte un Excel pret a importer dans Multidoc.
            </p>
          </div>
          <div className="right-stack">
            <div className="stats" style={noDragStyle}>
              <div className="stat">
                <span className="stat-label">Fichier</span>
                <span className="stat-value">{inputPath ? 'Selectionne' : 'En attente'}</span>
              </div>
              <div className="stat">
                <span className="stat-label">Etat</span>
                <div className="stat-value-row">
                  <span className={`stat-value ${error ? 'error-text' : ''}`}>
                    {error || (busy ? 'Conversion...' : loadingPreview ? 'Chargement...' : status)}
                  </span>
                  {(busy || loadingPreview) && (
                    <RefreshCw size={16} className="spinner" />
                  )}
                </div>
              </div>
              <div className="stat">
                <span className="stat-label">Lignes</span>
                <span className="stat-value">{lastExport ? lastExport.rowCount : '-'}</span>
              </div>
            </div>
            {updateStatus.state !== 'idle' ? (
              <div className="update-box" style={noDragStyle}>
                <div>
                  <p className="label">Mise a jour</p>
                  <p className="tiny">
                    {updateStatus.state === 'available' && `Version ${updateStatus.version} disponible.`}
                    {updateStatus.state === 'downloading' && `Telechargement... ${updateStatus.progress ?? 0}%`}
                    {updateStatus.state === 'downloaded' && `Version ${updateStatus.version} telechargee.`}
                    {updateStatus.state === 'error' && updateStatus.message}
                  </p>
                </div>
                {updateStatus.state === 'available' && (
                  <button className="btn secondary small" onClick={downloadUpdate}>
                    Telecharger
                  </button>
                )}
                {updateStatus.state === 'downloading' && <span className="pill">Telechargement...</span>}
                {updateStatus.state === 'downloaded' && (
                  <button className="btn primary small" onClick={installUpdate}>
                    Installer
                  </button>
                )}
              </div>
            ) : null}
          </div>
        </header>

        <main className="layout">
          <section className="panel">
            <div className="panel-head">
              <div>
                <p className="eyebrow subtle">Etape 1</p>
                <h2>Choisir le fichier source</h2>
                <p className="hint">Selectionne l'export Caneco (.xls/.xlsx). Les .xls sont convertis via Excel.</p>
              </div>
              {inputPath ? <span className="pill">OK</span> : <span className="pill muted">En attente</span>}
            </div>

            <div
              className={`path-box ${isDragging ? 'dragging' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <p className="label">Fichier source</p>
              <p className="path">
                {isDragging
                  ? 'Deposer le fichier Excel ici...'
                  : inputPath
                    ? shortenPath(inputPath)
                    : 'Aucun fichier selectionne (ou glisse un fichier ici)'}
              </p>
            </div>

            <div className="buttons" style={noDragStyle}>
              <button className="btn primary" onClick={pickInput} disabled={busy}>
                <FileSpreadsheet size={16} />
                {busy ? 'Patiente...' : 'Choisir un Excel'}
              </button>
              <button
                className="btn ghost"
                onClick={() => inputPath && loadPreview(inputPath)}
                disabled={!inputPath || loadingPreview || busy}
              >
                <RefreshCw size={16} />
                {loadingPreview ? 'Chargement...' : 'Recharger les lignes'}
              </button>
              <button className="btn ghost" onClick={() => inputPath && pickOutput()} disabled={!inputPath || busy}>
                <FolderOpen size={16} />
                Choisir la sortie
              </button>
            </div>
          </section>

          <section className="panel">
            <div className="panel-head">
              <div>
                <p className="eyebrow subtle">Etape 2</p>
                <h2>Definir les prix</h2>
                <p className="hint">
                  {priceMode === 'perLine'
                    ? 'Saisis un prix pour chaque ligne.'
                    : 'Saisis un prix par cable (colonne D) et type (colonne G).'}
                </p>
              </div>
              <span className="pill muted">
                {loadingPreview
                  ? 'Chargement'
                  : priceMode === 'perLine'
                    ? previewRows.length
                      ? `${previewRows.length} lignes`
                      : 'Aucune ligne'
                    : cableTypes.length
                      ? `${cableTypes.length} cables + types`
                      : 'Aucun cable + type'}
              </span>
            </div>

            <div className="price-mode" style={noDragStyle}>
              <button
                className={`btn ghost small ${priceMode === 'perLine' ? 'active' : ''}`}
                onClick={() => setPriceMode('perLine')}
              >
                Par ligne
              </button>
              <button
                className={`btn ghost small ${priceMode === 'perCable' ? 'active' : ''}`}
                onClick={() => setPriceMode('perCable')}
              >
                Par cable + type
              </button>
            </div>

            <div className="form-grid" style={noDragStyle}>
              <label className="field">
                <span className="label">Prix par defaut</span>
                <input
                  type="text"
                  inputMode="decimal"
                  placeholder="0"
                  value={defaultUnitPrice}
                  onChange={(event) => setDefaultUnitPrice(event.target.value)}
                />
              </label>
              <div className="field">
                <span className="label">Actions</span>
                <button
                  className="btn ghost small"
                  onClick={applyDefaultPrices}
                  disabled={priceMode === 'perLine' ? !previewRows.length : !cableTypes.length}
                >
                  Appliquer a toutes
                </button>
              </div>
              <label className="field">
                <span className="label">TVA</span>
                <input type="text" inputMode="decimal" placeholder="0" value={tva} onChange={(event) => setTva(event.target.value)} />
              </label>
            </div>

            {loadingPreview ? (
              <div className="price-table-empty">Chargement...</div>
            ) : priceMode === 'perLine' ? (
              previewRows.length ? (
                <div className="price-table" style={noDragStyle}>
                  <div className="price-row header lines">
                    <div className="price-cell">#</div>
                    <div className="price-cell title">Designation</div>
                    <div className="price-cell cable">Cable + type</div>
                    <div className="price-cell qty">Qt</div>
                    <div className="price-cell input">Prix</div>
                  </div>
                  {previewRows.map((row, index) => (
                    <div className="price-row lines" key={row.lineNumber}>
                      <div className="price-cell">{row.lineNumber}</div>
                      <div className="price-cell title">{row.title}</div>
                      <div className="price-cell cable">{row.typeCable || '-'}</div>
                      <div className="price-cell qty">{row.quantity}</div>
                      <div className="price-cell input">
                        <input
                          className="price-input"
                          type="text"
                          inputMode="decimal"
                          placeholder="0"
                          value={unitPrices[index] ?? ''}
                          onChange={(event) => updateUnitPriceAt(index, event.target.value)}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="price-table-empty">Charge un fichier pour voir les lignes.</div>
              )
            ) : cableTypes.length ? (
              <div className="price-table" style={noDragStyle}>
                <div className="price-row header types">
                  <div className="price-cell title">Cable + type</div>
                  <div className="price-cell qty">Lignes</div>
                  <div className="price-cell input">Prix</div>
                </div>
                {cableTypes.map((typeKey) => (
                  <div className="price-row types" key={typeKey || 'type-vide'}>
                    <div className="price-cell title">{formatCableType(typeKey)}</div>
                    <div className="price-cell qty">{typeCounts[typeKey] ?? 0}</div>
                    <div className="price-cell input">
                      <input
                        className="price-input"
                        type="text"
                        inputMode="decimal"
                        placeholder="0"
                        value={typePrices[typeKey] ?? ''}
                        onChange={(event) => updateTypePrice(typeKey, event.target.value)}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="price-table-empty">Charge un fichier pour voir les cables + types.</div>
            )}
          </section>

          <section className="panel">
            <div className="panel-head">
              <div>
                <p className="eyebrow subtle">Etape 3</p>
                <h2>Exporter vers Multidoc</h2>
                <p className="hint">Le fichier est genere avec les colonnes attendues.</p>
              </div>
              {lastExport ? <span className="pill">Pret</span> : <span className="pill muted">Conversion</span>}
            </div>

            <div className="path-box">
              <p className="label">Fichier Multidoc</p>
              <p className="path">{outputLabel ? shortenPath(outputLabel) : 'Aucune sortie pour le moment.'}</p>
            </div>

            <div className="buttons" style={noDragStyle}>
              <button className="btn primary" onClick={convert} disabled={busy || !inputPath || previewRows.length === 0}>
                <RefreshCw size={16} />
                {busy ? 'Conversion...' : 'Convertir'}
              </button>
              <button className="btn secondary" onClick={revealExport} disabled={!outputLabel}>
                <CheckCircle2 size={16} />
                Ouvrir le dossier
              </button>
            </div>

            {error ? <div className="error-box">{error}</div> : null}
          </section>
        </main>

        <footer className="footer">CaneFlow - v{currentVersion}</footer>
      </div>
    </div>
  )
}

export default App

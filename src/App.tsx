import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import { CheckCircle2, FileSpreadsheet, Info, Maximize2, Minus, RefreshCw, X as Close } from 'lucide-react'
import './App.css'
import logo from './assets/logo.svg'

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

type NexusDetail = { mat: number; margin: number; tpsPose: number; puPose: number; ptPose: number }

function NexusPill({ detail }: { detail: NexusDetail }) {
  const matMarged = detail.mat * detail.margin
  return (
    <div className="nexus-pill">
      <div className="nexus-pill-row">
        <span className="nexus-pill-label">Matériau</span>
        <span className="nexus-pill-calc">
          {detail.mat.toFixed(3)} €/ml <span className="nexus-pill-op">×</span> {detail.margin} <span className="nexus-pill-op">=</span> <strong>{matMarged.toFixed(3)} €/ml</strong>
        </span>
      </div>
      <div className="nexus-pill-row">
        <span className="nexus-pill-label">Pose</span>
        <span className="nexus-pill-calc">
          {detail.tpsPose} h <span className="nexus-pill-op">×</span> {detail.puPose.toFixed(2)} €/h <span className="nexus-pill-op">=</span> <strong>{detail.ptPose.toFixed(3)} €/ml</strong>
        </span>
      </div>
      <div className="nexus-pill-divider" />
      <div className="nexus-pill-row nexus-pill-total">
        <span className="nexus-pill-label">Total</span>
        <strong>{(matMarged + detail.ptPose).toFixed(2)} €/ml</strong>
      </div>
    </div>
  )
}

function App() {
  const currentVersion = import.meta.env.VITE_APP_VERSION || '0.0.0'
  const [inputPath, setInputPath] = useState('')
  const [outputPath, setOutputPath] = useState('')
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([])
  const [cableTypes, setCableTypes] = useState<string[]>([])
  const [typePrices, setTypePrices] = useState<Record<string, string>>({})
  const [unitPrices, setUnitPrices] = useState<string[]>([])
  const [priceMode, setPriceMode] = useState<'perCable' | 'perLine'>('perLine')
  const [margin, setMargin] = useState('1.33')
  const [appliedMargin, setAppliedMargin] = useState('1.33')
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
  const [currentStep, setCurrentStep] = useState<1 | 2 | 3>(1)
  const [expandedStep, setExpandedStep] = useState<1 | 2 | 3>(1)
  const [nexusFilledTypes, setNexusFilledTypes] = useState<Set<string>>(new Set())
  const [nexusFilledLines, setNexusFilledLines] = useState<Set<number>>(new Set())
  const [nexusSource, setNexusSource] = useState<'r2' | 'fallback' | 'none' | null>(null)
  const [nexusDetails, setNexusDetails] = useState<Record<string, { mat: number; margin: number; tpsPose: number; puPose: number; ptPose: number }>>({})

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

  const applyNexusPrices = async (typeOrder: string[], rows: PreviewRow[], marginValue: string) => {
    if (!hasApi() || rows.length === 0) return
    const parsedMargin = parseFloat(marginValue)
    if (isNaN(parsedMargin) || parsedMargin <= 0) return
    try {
      const { prices: nexusPrices, source, details } = await window.api.lookupCablePrices(typeOrder, parsedMargin)
      setNexusSource(source)
      setNexusDetails(details)
      if (Object.keys(nexusPrices).length === 0) return
      const filledTypes = new Set<string>()
      for (const key of typeOrder) {
        if (nexusPrices[key] !== undefined) filledTypes.add(key)
      }
      const filledLines = new Set<number>()
      rows.forEach((row, idx) => {
        if (nexusPrices[row.typeCable || ''] !== undefined) filledLines.add(idx)
      })
      setTypePrices((prev) => {
        const next = { ...prev }
        for (const key of typeOrder) {
          if (nexusPrices[key] !== undefined) next[key] = nexusPrices[key].toFixed(2)
        }
        return next
      })
      setUnitPrices((prev) => {
        const next = [...prev]
        rows.forEach((row, idx) => {
          const p = nexusPrices[row.typeCable || '']
          if (p !== undefined) next[idx] = p.toFixed(2)
        })
        return next
      })
      setNexusFilledTypes(filledTypes)
      setNexusFilledLines(filledLines)
      setAppliedMargin(marginValue)
    } catch (err) {
      console.error('NEXUS lookup error', err)
    }
  }


  const updateTypePrice = (typeCable: string, value: string) => {
    setNexusFilledTypes((prev) => { const s = new Set(prev); s.delete(typeCable); return s })
    setTypePrices((prev) => ({ ...prev, [typeCable]: value }))
  }

  const updateUnitPriceAt = (index: number, value: string) => {
    setNexusFilledLines((prev) => { const s = new Set(prev); s.delete(index); return s })
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
      setUnitPrices(rows.map(() => ''))
      setCableTypes(typeOrder)
      setTypePrices(Object.fromEntries(typeOrder.map((k) => [k, ''])))
      setNexusFilledTypes(new Set())
      setNexusFilledLines(new Set())
      setStatus(rows.length ? `Lignes chargees (${rows.length}).` : 'Aucune ligne detectee.')
      console.log('[CaneFlow UI] preview done', { count: rows.length })

      if (hasApi() && rows.length > 0) {
        await applyNexusPrices(typeOrder, rows, margin)
      }
    } catch (err) {
      console.error('previewRows', err)
      const errorMessage = err instanceof Error ? err.message : 'Impossible de lire le fichier Excel.'
      setError(errorMessage)
      setStatus('')
      setPreviewRows([])
      setCableTypes([])
      setTypePrices({})
      setUnitPrices([])
      setNexusFilledTypes(new Set())
      setNexusFilledLines(new Set())
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
      setCurrentStep(2)
      setExpandedStep(2)
      // Rester à l'étape 1 et afficher les boutons de choix
    } catch (err) {
      console.error('pickInput', err)
      setError('Impossible de selectionner le fichier.')
      setStatus('Erreur de selection.')
    }
  }

  const convert = async () => {
    setError('')
    if (!inputPath) {
      setError('Sélectionne un fichier Excel en entrée.')
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
    setUpdateStatus((prev) => ({ ...prev, state: 'downloading' }))
    window.api.downloadUpdate()
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
      const filePath = window.api.getFilePathFromDrop(excelFile)
      if (filePath) {
        setInputPath(filePath)
        setOutputPath('')
        setLastExport(null)
        setStatus('Fichier charge. Pret a convertir.')
        await loadPreview(filePath)
        setCurrentStep(2)
        setExpandedStep(2)
        // Rester à l'étape 1 et afficher les boutons de choix
      } else {
        setError('Impossible de lire le chemin du fichier')
      }
    } else {
      setError('Veuillez glisser un fichier Excel (.xls ou .xlsx)')
    }
  }

  const goToStep3 = () => {
    setCurrentStep(3)
    setExpandedStep(3)
  }

  const editStep = (step: 1 | 2 | 3) => {
    setExpandedStep(step)
  }

  const resetForNewExport = () => {
    setInputPath('')
    setOutputPath('')
    setPreviewRows([])
    setCableTypes([])
    setTypePrices({})
    setUnitPrices([])
    setMargin('1.33')
    setTva('0')
    setNexusFilledTypes(new Set())
    setNexusFilledLines(new Set())
    setNexusSource(null)
    setNexusDetails({})
    setAppliedMargin('1.33')
    setStatus('En attente de selection.')
    setError('')
    setLastExport(null)
    setCurrentStep(1)
    setExpandedStep(1)
  }

  return (
    <div className="app-shell">
      <div className="titlebar">
        <div className="window-title">
          <img src={logo} alt="CaneFlow" className="title-logo" />
          <span className="window-title-text">CaneFlow</span>
          {nexusSource !== null && (
            <span
              className={`nexus-dot nexus-dot--${nexusSource}`}
              title={nexusSource === 'r2' ? 'NEXUS connecté via R2' : nexusSource === 'fallback' ? 'NEXUS via réseau local (Z:)' : 'NEXUS non disponible'}
            />
          )}
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
                Exporte le carnet de câbles Caneco avec les colonnes suivantes, puis charge le fichier dans CaneFlow.
              </p>
              <div className="info-table caneco">
                <div className="info-row header">
                  <div>Colonnes Caneco</div>
                </div>
                <div className="info-row caneco-columns">
                  <span>Amont</span>
                  <span>Descriptif</span>
                  <span>Longueur</span>
                  <span>Câble</span>
                  <span>Neutre</span>
                  <span>PE ou PEN</span>
                  <span>Type de câble</span>
                </div>
              </div>

              <div className="modal-section">
                <p className="label">Étapes</p>
                <ol className="modal-list">
                  <li>Cliquer sur "Choisir un Excel", puis "Choisir la sortie".</li>
                  <li>Étape 2: saisir les prix unitaires par ligne ou par câble + type, et la TVA globale.</li>
                  <li>Cliquer sur "Convertir" pour générer l'Excel Multidoc.</li>
                </ol>
              </div>

              <div className="modal-section">
                <p className="label">Réglages Multidoc</p>
                <p className="modal-text">Dans Multidoc, régler les numéros de colonnes comme suit :</p>
                <div className="info-table mapping">
                  <div className="info-row header two-col">
                    <span>Champ</span>
                    <span>Numéro</span>
                  </div>
                  <div className="info-row">
                    <span>Numéros</span>
                    <span>1</span>
                  </div>
                  <div className="info-row">
                    <span>Titres</span>
                    <span>2</span>
                  </div>
                  <div className="info-row">
                    <span>Unités</span>
                    <span>3</span>
                  </div>
                  <div className="info-row">
                    <span>Quantités</span>
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
        <header className="app-bar">
          <div className="app-bar-brand" style={noDragStyle}>
            <img src={logo} alt="CaneFlow" className="app-bar-logo" />
            <span className="app-bar-name">CaneFlow</span>
            <span className="app-bar-version">v{currentVersion}</span>
          </div>

          <div className="app-bar-stats" style={noDragStyle}>
            <span className="app-bar-stat-label">Fichier</span>
            <span className={`app-bar-stat-value ${!inputPath ? 'muted' : ''}`}>
              {inputPath ? 'Sélectionné' : 'En attente'}
            </span>
            <span className="app-bar-divider" />
            <span className="app-bar-stat-label">État</span>
            <span className={`app-bar-stat-value ${error ? 'error' : ''}`}>
              {(busy || loadingPreview) && <RefreshCw size={11} className="spinner" style={{ marginRight: 4 }} />}
              {error || (busy ? 'Conversion…' : loadingPreview ? 'Chargement…' : status)}
            </span>
            {lastExport && (
              <>
                <span className="app-bar-divider" />
                <span className="app-bar-stat-label">Lignes</span>
                <span className="app-bar-stat-value">{lastExport.rowCount}</span>
              </>
            )}
          </div>

          <div className="app-bar-actions" style={noDragStyle}>
            {updateStatus.state !== 'idle' && (
              <div className="app-bar-update">
                <span className="app-bar-stat-label">
                  {updateStatus.state === 'available' && `v${updateStatus.version} disponible`}
                  {updateStatus.state === 'downloading' && `Téléchargement ${updateStatus.progress ?? 0}%`}
                  {updateStatus.state === 'downloaded' && `v${updateStatus.version} prête`}
                  {updateStatus.state === 'error' && 'Erreur de mise à jour'}
                </span>
                {updateStatus.state === 'available' && (
                  <button className="btn secondary small" onClick={downloadUpdate}>Télécharger</button>
                )}
                {updateStatus.state === 'downloading' && <RefreshCw size={13} className="spinner" />}
                {updateStatus.state === 'downloaded' && (
                  <button className="btn primary small" onClick={installUpdate}>Installer</button>
                )}
              </div>
            )}
            <button className="info-inline-btn" onClick={() => setShowInfo(true)} aria-label="Aide">
              <Info size={15} />
            </button>
          </div>
        </header>

        <main className="layout">
          {/* ÉTAPE 1 */}
          <section className={`panel ${expandedStep === 1 ? 'expanded' : 'collapsed'} ${currentStep >= 2 ? 'completed' : ''}`}>
            <div className="panel-head" onClick={() => currentStep >= 2 && editStep(1)} style={currentStep >= 2 ? { cursor: 'pointer' } : {}}>
              <div>
                <p className="eyebrow subtle">Etape 1</p>
                <h2>Choisir le fichier source</h2>
                {expandedStep === 1 && <p className="hint">Sélectionne l'export Caneco (.xls/.xlsx). Les .xls sont convertis via Excel.</p>}
                {expandedStep !== 1 && currentStep >= 2 && inputPath && (
                  <p className="summary">{shortenPath(inputPath)}</p>
                )}
              </div>
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                {currentStep >= 2 && <CheckCircle2 size={18} style={{ color: '#10b981' }} />}
                {currentStep >= 2 && expandedStep !== 1 && <button className="btn ghost small">Modifier</button>}
                {expandedStep === 1 && inputPath && <span className="pill">OK</span>}
                {expandedStep === 1 && !inputPath && <span className="pill muted">En attente</span>}
              </div>
            </div>

            {expandedStep === 1 && (
              <div className="panel-content">

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
                <RefreshCw size={16} className={loadingPreview ? 'spinner' : ''} />
                {loadingPreview ? 'Chargement...' : 'Recharger les lignes'}
              </button>
            </div>

            {inputPath && !loadingPreview && (
              <div className="step-actions" style={noDragStyle}>
                <button className="btn secondary" onClick={() => { setCurrentStep(2); setExpandedStep(2); }}>
                  Définir les prix →
                </button>
                <button className="btn primary" onClick={goToStep3}>
                  Exporter directement →
                </button>
              </div>
            )}
              </div>
            )}
          </section>

          {/* ÉTAPE 2 */}
          <section className={`panel price-panel ${expandedStep === 2 ? 'expanded' : 'collapsed'} ${currentStep >= 3 ? 'completed' : ''} ${currentStep < 2 ? 'locked' : ''}`}>
            <div className="panel-head" onClick={() => currentStep >= 2 && currentStep >= 3 && editStep(2)} style={currentStep >= 3 ? { cursor: 'pointer' } : {}}>
              <div>
                <p className="eyebrow subtle">Etape 2</p>
                <h2>Définir les prix</h2>
                {expandedStep === 2 && (
                  <p className="hint">
                    Les prix sont synchronisés depuis NEXUS — complète les champs manquants.
                  </p>
                )}
                {expandedStep !== 2 && currentStep >= 3 && (
                  <p className="summary">
                    Prix configuré - {priceMode === 'perLine' ? `${previewRows.length} lignes` : `${cableTypes.length} types de câble`}
                  </p>
                )}
                {currentStep < 2 && <p className="hint-muted">Complète l'étape 1 d'abord</p>}
              </div>
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                {currentStep >= 3 && <CheckCircle2 size={18} style={{ color: '#10b981' }} />}
                {currentStep >= 3 && expandedStep !== 2 && <button className="btn ghost small">Modifier</button>}
                {expandedStep === 2 && currentStep === 2 && (
                  <span className="pill muted">
                    {loadingPreview
                      ? 'Chargement'
                      : priceMode === 'perLine'
                        ? previewRows.length
                          ? `${previewRows.length} lignes`
                          : 'Aucune ligne'
                        : cableTypes.length
                          ? `${cableTypes.length} câbles + types`
                          : 'Aucun câble + type'}
                  </span>
                )}
              </div>
            </div>

            {expandedStep === 2 && currentStep >= 2 && (
              <div className="panel-content">

            <div className="table-toolbar" style={noDragStyle}>
              <div className="mode-tabs">
                <button
                  className={`mode-tab ${priceMode === 'perLine' ? 'active' : ''}`}
                  onClick={() => setPriceMode('perLine')}
                >
                  Par ligne
                </button>
                <button
                  className={`mode-tab ${priceMode === 'perCable' ? 'active' : ''}`}
                  onClick={() => setPriceMode('perCable')}
                >
                  Par câble + type
                </button>
              </div>

              <div className="toolbar-separator" />

              <div className="toolbar-param">
                <span className="toolbar-param-label">Marge ×</span>
                <input
                  className="toolbar-input"
                  type="text"
                  inputMode="decimal"
                  placeholder="1.33"
                  value={margin}
                  onChange={(event) => setMargin(event.target.value)}
                />
                <button
                  className={`btn small ${margin !== appliedMargin ? 'primary pulse-cta' : 'ghost'}`}
                  onClick={() => applyNexusPrices(cableTypes, previewRows, margin)}
                  disabled={previewRows.length === 0}
                >
                  Recalculer
                </button>
              </div>

              <div className="toolbar-separator" />

              <div className="toolbar-param">
                <span className="toolbar-param-label">TVA</span>
                <input
                  className="toolbar-input"
                  type="text"
                  inputMode="decimal"
                  placeholder="0"
                  value={tva}
                  onChange={(event) => setTva(event.target.value)}
                />
                <span className="toolbar-param-label">%</span>
              </div>

              <div className="toolbar-legend">
                <span className="legend-dot nexus" />
                <span className="toolbar-param-label">Prix NEXUS</span>
              </div>
            </div>

            {loadingPreview ? (
              <div className="price-table-empty">Chargement...</div>
            ) : priceMode === 'perLine' ? (
              previewRows.length ? (
                <div className="price-table" style={noDragStyle}>
                  <div className="price-row header lines">
                    <div className="price-cell">#</div>
                    <div className="price-cell title">Désignation</div>
                    <div className="price-cell cable">Câble + type</div>
                    <div className="price-cell qty">Qt</div>
                    <div className="price-cell input">Prix</div>
                  </div>
                  {previewRows.map((row, index) => {
                    const isMissing = !nexusFilledLines.has(index) && unitPrices[index] === ''
                    return (
                    <div className="price-row lines" key={row.lineNumber}>
                      <div className="price-cell">{row.lineNumber}</div>
                      <div className="price-cell title">{row.title}</div>
                      <div className="price-cell cable">{row.typeCable || '-'}</div>
                      <div className="price-cell qty">{row.quantity}</div>
                      <div className="price-cell input">
                        <div className="nexus-price-wrap">
                          <input
                            className={`price-input${nexusFilledLines.has(index) ? ' nexus-filled' : (isMissing ? ' price-missing' : '')}`}
                            style={isMissing ? { animationDelay: `${(index % 15) * 0.1}s` } : undefined}
                            type="text"
                            inputMode="decimal"
                            placeholder="0"
                            value={unitPrices[index] ?? ''}
                            onChange={(event) => updateUnitPriceAt(index, event.target.value)}
                          />
                          {nexusFilledLines.has(index) && nexusDetails[row.typeCable] && (
                            <NexusPill detail={nexusDetails[row.typeCable]} />
                          )}
                        </div>
                      </div>
                    </div>
                    )
                  })}
                </div>
              ) : (
                <div className="price-table-empty">Charge un fichier pour voir les lignes.</div>
              )
            ) : cableTypes.length ? (
              <div className="price-table" style={noDragStyle}>
                <div className="price-row header types">
                  <div className="price-cell title">Câble + type</div>
                  <div className="price-cell qty">Lignes</div>
                  <div className="price-cell input">Prix</div>
                </div>
                {cableTypes.map((typeKey, typeIdx) => {
                  const isMissing = !nexusFilledTypes.has(typeKey) && typePrices[typeKey] === ''
                  return (
                  <div className="price-row types" key={typeKey || 'type-vide'}>
                    <div className="price-cell title">{formatCableType(typeKey)}</div>
                    <div className="price-cell qty">{typeCounts[typeKey] ?? 0}</div>
                    <div className="price-cell input">
                      <div className="nexus-price-wrap">
                        <input
                          className={`price-input${nexusFilledTypes.has(typeKey) ? ' nexus-filled' : (isMissing ? ' price-missing' : '')}`}
                          style={isMissing ? { animationDelay: `${typeIdx * 0.1}s` } : undefined}
                          type="text"
                          inputMode="decimal"
                          placeholder="0"
                          value={typePrices[typeKey] ?? ''}
                          onChange={(event) => updateTypePrice(typeKey, event.target.value)}
                        />
                        {nexusFilledTypes.has(typeKey) && nexusDetails[typeKey] && (
                          <NexusPill detail={nexusDetails[typeKey]} />
                        )}
                      </div>
                    </div>
                  </div>
                  )
                })}
              </div>
            ) : (
              <div className="price-table-empty">Charge un fichier pour voir les câbles + types.</div>
            )}

            <div className="step-actions" style={noDragStyle}>
              <button className="btn primary" onClick={goToStep3}>
                Passer à l'étape 3 →
              </button>
            </div>
              </div>
            )}
          </section>

          {/* ÉTAPE 3 */}
          <section className={`panel ${expandedStep === 3 ? 'expanded' : 'collapsed'} ${currentStep < 3 ? 'locked' : ''}`}>
            <div className="panel-head">
              <div>
                <p className="eyebrow subtle">Etape 3</p>
                <h2>Exporter vers Multidoc</h2>
                {expandedStep === 3 && <p className="hint">Le fichier est genere avec les colonnes attendues.</p>}
                {currentStep < 3 && <p className="hint-muted">Complète l'étape 2 d'abord</p>}
              </div>
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                {lastExport && <CheckCircle2 size={18} style={{ color: '#10b981' }} />}
                {expandedStep === 3 && currentStep === 3 && (
                  lastExport ? <span className="pill">Pret</span> : <span className="pill muted">Conversion</span>
                )}
              </div>
            </div>

            {expandedStep === 3 && currentStep === 3 && (
              <div className="panel-content">
                <div className="path-box">
                  <p className="label">Fichier Multidoc</p>
                  <p className="path">{outputLabel ? shortenPath(outputLabel) : 'Aucune sortie pour le moment.'}</p>
                </div>

                <div className="buttons" style={noDragStyle}>
                  <button className="btn primary" onClick={convert} disabled={busy || !inputPath || previewRows.length === 0}>
                    <RefreshCw size={16} className={busy ? 'spinner' : ''} />
                    {busy ? 'Conversion...' : 'Convertir'}
                  </button>
                  <button className="btn secondary" onClick={revealExport} disabled={!lastExport}>
                    <CheckCircle2 size={18} />
                    Ouvrir le dossier
                  </button>
                  {lastExport && (
                    <button className="btn new-export" onClick={resetForNewExport}>
                      Nouvel export
                    </button>
                  )}
                </div>

                {error ? <div className="error-box">{error}</div> : null}
              </div>
            )}
          </section>
        </main>

        <footer className="footer">CaneFlow - v{currentVersion}</footer>
      </div>
    </div>
  )
}

export default App

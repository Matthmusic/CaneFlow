import 'react'

declare module 'react' {
  interface CSSProperties {
    WebkitAppRegion?: 'drag' | 'no-drag'
  }
}

export {}

declare global {
  interface Window {
    api: {
      pickInputFile: () => Promise<string | null>
      pickOutputFile: (inputPath?: string) => Promise<string | null>
      previewRows: (payload: { inputPath: string; sheetName?: string }) => Promise<
        { lineNumber: number; title: string; quantity: string | number; repere: string; typeCable: string }[]
      >
      convertFile: (payload: {
        inputPath: string
        outputPath?: string
        unitPrice?: string | number
        unitPrices?: Array<string | number>
        unitPricesByType?: Record<string, string | number>
        tva?: string | number
        includeHeaders?: boolean
        sheetName?: string
      }) => Promise<{ outputPath: string; rowCount: number }>
      getFilePathFromDrop: (file: File) => string | null
      onAppLog: (callback: (data: { level: string; message: string; data: any; timestamp: string }) => void) => () => void
      revealPath: (targetPath: string) => Promise<void>
      windowClose: () => Promise<void>
      windowMinimize: () => Promise<void>
      windowToggleMaximize: () => Promise<void>
      checkUpdates: () => Promise<{ status: string; version?: string }>
      downloadUpdate: () => Promise<void>
      installUpdate: () => Promise<void>
      onUpdateEvent: (callback: (data: any) => void) => () => void
    }
  }
}

interface ImportMetaEnv {
  readonly VITE_APP_VERSION: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}

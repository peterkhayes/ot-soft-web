import { createContext, useContext } from 'react'

export type BlobDownloadFn = (blob: Blob, filename: string) => void

function downloadBlobFile(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

export const BlobDownloadContext = createContext<BlobDownloadFn>(downloadBlobFile)
export const BlobDownloadProvider = BlobDownloadContext.Provider
export const useBlobDownload = () => useContext(BlobDownloadContext)

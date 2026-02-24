import { createContext, useContext } from 'react'

import { downloadTextFile } from '../utils'

export type DownloadFn = (content: string, filename: string) => void
export const DownloadContext = createContext<DownloadFn>(downloadTextFile)
export const DownloadProvider = DownloadContext.Provider
export const useDownload = () => useContext(DownloadContext)

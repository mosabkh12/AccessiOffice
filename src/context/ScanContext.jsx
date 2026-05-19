import { createContext, useContext, useState } from 'react'

const ScanContext = createContext(null)

export function ScanProvider({ children }) {
  const [scanData, setScanData] = useState(null)

  return (
    <ScanContext.Provider value={{ scanData, setScanData }}>
      {children}
    </ScanContext.Provider>
  )
}

export function useScan() {
  const ctx = useContext(ScanContext)
  if (!ctx) throw new Error('useScan must be used within ScanProvider')
  return ctx
}

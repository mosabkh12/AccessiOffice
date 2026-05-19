import { Routes, Route } from 'react-router-dom'
import Layout from './components/Layout.jsx'
import HomePage from './pages/HomePage.jsx'
import UploadPage from './pages/UploadPage.jsx'
import ScanProgressPage from './pages/ScanProgressPage.jsx'
import ResultsPage from './pages/ResultsPage.jsx'
import ReportPage from './pages/ReportPage.jsx'

export default function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/upload" element={<UploadPage />} />
        <Route path="/scan" element={<ScanProgressPage />} />
        <Route path="/results" element={<ResultsPage />} />
        <Route path="/report" element={<ReportPage />} />
      </Routes>
    </Layout>
  )
}

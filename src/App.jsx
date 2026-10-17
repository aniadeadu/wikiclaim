
import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import {
  ShieldCheck,
  Files,
  CheckCircle,
  CircleHelp,
  ExternalLink
} from 'lucide-react'

function App() {
  const [claims, setClaims] = useState([])
  const [selectedStatus, setSelectedStatus] = useState('All statuses')
  const [message, setMessage] = useState('Loading claims...')

  useEffect(() => {
    async function loadClaims() {
      try {
        const response = await fetch('/claims.xlsx')

        if (!response.ok) {
          throw new Error('Could not find claims.xlsx')
        }

        const data = await response.arrayBuffer()
        const workbook = XLSX.read(data, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })

        const importedClaims = rows
          .map((row, index) => ({
            id: index,
            text: row.Claim || row.claim || '',
            source: row.Source || row.source || '',
            status: row.Verdict || row.verdict || row.Status || row.status || 'Unclear',
            reason: row.Reason || row.reason || ''
          }))
          .filter((claim) => claim.text)

        setClaims(importedClaims)
        setMessage(`Loaded ${importedClaims.length} claims.`)
      } catch (error) {
        console.error(error)
        setMessage('Could not load claims.xlsx. Check that it is inside the public folder.')
      }
    }

    loadClaims()
  }, [])

  function handleStatusChange(id, newStatus) {
    setClaims((currentClaims) =>
      currentClaims.map((claim) =>
        claim.id === id ? { ...claim, status: newStatus } : claim
      )
    )
  }

  const statuses = [
    ...new Set(claims.map((claim) => claim.status).filter(Boolean))
  ]

  const filteredClaims = selectedStatus === 'All statuses'
    ? claims
    : claims.filter((claim) => claim.status === selectedStatus)

  return (
    <main>
      <h1 className="app-title">
  <ShieldCheck size={30} />
  WikiClaim Bench
</h1>
      <p>Wikipedia claim and source verification dashboard</p>

      <h2>Claims to review</h2>
      <p>{message}</p>

      <label htmlFor="status-filter">Filter by verdict: </label>
      <select
        id="status-filter"
        value={selectedStatus}
        onChange={(event) => setSelectedStatus(event.target.value)}
      >
        <option value="All statuses">All statuses</option>
        {statuses.map((status) => (
          <option key={status} value={status}>
            {status}
          </option>
        ))}
      </select>

      <div className="table-container">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Claim</th>
              <th>Cited source</th>
              <th>Verdict</th>
              <th>Reason</th>
            </tr>
          </thead>
          <tbody>
            {filteredClaims.map((claim, index) => (
              <tr key={claim.id}>
                <td>{index + 1}</td>
                <td>{claim.text}</td>
                <td>
  {claim.source ? (
    <a
      className="source-link"
      href={claim.source.trim()}
      target="_blank"
      rel="noopener noreferrer"
      title={claim.source}
    >
      {claim.source}
    </a>
  ) : (
    '—'
  )}
</td>

                <td>
                  <select
                    value={claim.status}
                    onChange={(event) =>
                      handleStatusChange(claim.id, event.target.value)
                    }
                    aria-label={`Verdict for claim ${index + 1}`}
                  >
                    <option value="Supported">Supported</option>
                    <option value="Not supported">Not supported</option>
                    <option value="Unclear">Unclear</option>
                  </select>
                </td>
                <td>{claim.reason || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </main>
  )
}

export default App


import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import {
  ShieldCheck,
  Files,
  CheckCircle,
  CircleHelp,
  ExternalLink,
  Layers3,
  ScanSearch,
  BookOpen
} from 'lucide-react'

function App() {
  const [claims, setClaims] = useState([])
  const [selectedStatus, setSelectedStatus] = useState('All statuses')
  const [selectedCategory, setSelectedCategory] = useState('All claims')
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
    article: row.Article || row.article || '',
    source: row.Source || row.source || '',
    evidence: row.Evidence || row.evidence || '',
    status: row.Verdict || row.verdict || row.Status || row.status || 'Unclear',
    reason: row.Reason || row.reason || '',
    category: row.Category || row.category || ''
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

 
const verdictOptions = [
  'Supported',
  'Partially supported',
  'Contradicted',
  'Insufficient evidence',
  'Unclear'
]

const summaryCards = [
  {
    label: 'Total claims',
    count: claims.length,
    color: 'blue'
  },
  ...verdictOptions.map((status) => ({
    label: status,
    count: claims.filter((claim) => claim.status === status).length,
    color: status === 'Supported'
      ? 'green'
      : status === 'Partially supported'
        ? 'amber'
        : status === 'Contradicted'
          ? 'red'
          : 'blue'
  }))
]

const statuses = verdictOptions


  {/*const filteredClaims = selectedStatus === 'All statuses'
    ? claims
    : claims.filter((claim) => claim.status === selectedStatus) */}
   

const filteredClaims = claims.filter((claim) => {
  const matchesCategory =
    selectedCategory === 'All claims' ||
    (claim.category || '').trim().toLowerCase() ===
      selectedCategory.toLowerCase()

  const matchesStatus =
    selectedStatus === 'All statuses' ||
    claim.status === selectedStatus

  return matchesCategory && matchesStatus
})


  return (
    <main>
      <h1 className="app-title">
  <ShieldCheck size={30} />
  WikiClaim Bench
</h1>
      <p>Wikipedia claim and source verification dashboard</p>

     {/*<h2>Claims to review</h2>
      <p>{message}</p>*/}

      
<section className="overview">
  <div className="overview-heading">
    <div>
      <h2>Claim Verification Overview</h2>
      <p>A snapshot of all claims in your dataset</p>
    </div>
 
  </div>


<div className="category-filters">
  <button
    type="button"
    className={`category-btn ${selectedCategory === 'All claims' ? 'active' : ''}`}
    onClick={() => setSelectedCategory('All claims')}
    aria-pressed={selectedCategory === 'All claims'}
  >
    <Layers3 size={19} />
    <span>All Claims</span>
    <span className="category-count">{claims.length}</span>
  </button>

  <button
    type="button"
    className={`category-btn ai-btn ${selectedCategory === 'AI flagged' ? 'active' : ''}`}
    onClick={() => setSelectedCategory('AI flagged')}
    aria-pressed={selectedCategory === 'AI flagged'}
  >
    <ScanSearch size={19} />
    <span>AI Flagged</span>
    <span className="category-count">
      {claims.filter(
        (claim) => (claim.category || '').trim().toLowerCase() === 'ai flagged'
      ).length}
    </span>
  </button>

  <button
    type="button"
    className={`category-btn non-ai-btn ${selectedCategory === 'Not AI flagged' ? 'active' : ''}`}
    onClick={() => setSelectedCategory('Not AI flagged')}
    aria-pressed={selectedCategory === 'Not AI flagged'}
  >
    <BookOpen size={19} />
    <span>Not AI Flagged</span>
    <span className="category-count">
      {claims.filter(
        (claim) => (claim.category || '').trim().toLowerCase() === 'not ai flagged'
      ).length}
    </span>
  </button>
</div>


  <div className="summary-grid">
    {summaryCards.map((card) => (
      <div className="summary-card" key={card.label}>
        <span className="summary-label">{card.label}</span>
        <strong className={`summary-count ${card.color}`}>
          {card.count}
        </strong>
      </div>
    ))}
  </div>

  <p className="load-message">{message}</p>
</section>

<h2>Claims to review</h2>


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
<th>Article</th>
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
  {claim.article ? (
    <a
      className="source-link"
      href={claim.article.trim()}
      target="_blank"
      rel="noopener noreferrer"
      title={claim.article}
    >
      Open article <ExternalLink size={14} />
    </a>
  ) : (
    '—'
  )}
</td>
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
<option value="Partially supported">Partially supported</option>
<option value="Contradicted">Contradicted</option>
<option value="Insufficient evidence">Insufficient evidence</option>
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

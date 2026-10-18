
import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import {
  ShieldCheck,
  ExternalLink,
  Layers3,
  ScanSearch,
  BookOpen
} from 'lucide-react'

// Get a usable URL from a cell value or an Excel hyperlink.
function getExternalUrl(value, hyperlink) {
  const target = hyperlink?.Target || hyperlink?.target

  const text = String(target || value || '').trim()

  if (!text) return null

  if (/^https?:\/\//i.test(text)) return text

  if (/^www\./i.test(text)) return `https://${text}`

  if (/^[\w.-]+\.[a-z]{2,}([/?#][^\s]*)?$/i.test(text)) {
    return `https://${text}`
  }

  return null
}

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
        const workbook = XLSX.read(data, {
          type: 'array',
          cellStyles: true
        })

        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })

        // Find the spreadsheet column for each field.
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1')
        const headers = {}

        for (let col = range.s.c; col <= range.e.c; col++) {
          const address = XLSX.utils.encode_cell({
            r: range.s.r,
            c: col
          })

          const header = String(sheet[address]?.v || '')
            .trim()
            .toLowerCase()

          if (header) headers[header] = col
        }

        function getCellLink(rowIndex, columnName) {
          const col = headers[columnName.toLowerCase()]
          if (col === undefined) return null

          const address = XLSX.utils.encode_cell({
            r: range.s.r + rowIndex + 1,
            c: col
          })

          return sheet[address]?.l || null
        }

        const importedClaims = rows
          .map((row, index) => {
            const article = row.Article || row.article || ''
            const source = row.Source || row.source || ''

            return {
              id: index,
              text: row.Claim || row.claim || '',
              article,
              articleUrl: getExternalUrl(
                article,
                getCellLink(index, 'article')
              ),
              source,
              sourceUrl: getExternalUrl(
                source,
                getCellLink(index, 'source')
              ),
              evidence: row.Evidence || row.evidence || '',
              status:
                row.Verdict ||
                row.verdict ||
                row.Status ||
                row.status ||
                'Unclear',
              reason: row.Reason || row.reason || '',
              category: row.Category || row.category || ''
            }
          })
          .filter((claim) => claim.text)

        setClaims(importedClaims)
        setMessage(`Loaded ${importedClaims.length} claims.`)
      } catch (error) {
        console.error('Error loading claims:', error)
        setMessage(
          'Could not load claims.xlsx. Check that it is inside the public folder.'
        )
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
      color:
        status === 'Supported'
          ? 'green'
          : status === 'Partially supported'
            ? 'amber'
            : status === 'Contradicted'
              ? 'red'
              : 'blue'
    }))
  ]

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
            className={`category-btn ${
              selectedCategory === 'All claims' ? 'active' : ''
            }`}
            onClick={() => setSelectedCategory('All claims')}
            aria-pressed={selectedCategory === 'All claims'}
          >
            <Layers3 size={19} />
            <span>All Claims</span>
            <span className="category-count">{claims.length}</span>
          </button>

          <button
            type="button"
            className={`category-btn ai-btn ${
              selectedCategory === 'AI flagged' ? 'active' : ''
            }`}
            onClick={() => setSelectedCategory('AI flagged')}
            aria-pressed={selectedCategory === 'AI flagged'}
          >
            <ScanSearch size={19} />
            <span>AI Flagged</span>
            <span className="category-count">
              {
                claims.filter(
                  (claim) =>
                    (claim.category || '').trim().toLowerCase() ===
                    'ai flagged'
                ).length
              }
            </span>
          </button>

          <button
            type="button"
            className={`category-btn non-ai-btn ${
              selectedCategory === 'Not AI flagged' ? 'active' : ''
            }`}
            onClick={() => setSelectedCategory('Not AI flagged')}
            aria-pressed={selectedCategory === 'Not AI flagged'}
          >
            <BookOpen size={19} />
            <span>Not AI Flagged</span>
            <span className="category-count">
              {
                claims.filter(
                  (claim) =>
                    (claim.category || '').trim().toLowerCase() ===
                    'not ai flagged'
                ).length
              }
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
        {verdictOptions.map((status) => (
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
                  {claim.articleUrl ? (
                    <a
                      className="source-link"
                      href={claim.articleUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      title={claim.article}
                    >
                      Open article <ExternalLink size={14} />
                    </a>
                  ) : (
                    claim.article || '—'
                  )}
                </td>

                <td>
                  {claim.sourceUrl ? (
                    <a
                      className="source-link"
                      href={claim.sourceUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      title={claim.source}
                    >
                      {claim.source} <ExternalLink size={14} />
                    </a>
                  ) : (
                    claim.source || '—'
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
                    {verdictOptions.map((status) => (
                      <option key={status} value={status}>
                        {status}
                      </option>
                    ))}
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

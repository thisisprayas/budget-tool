import { useState, useCallback } from 'react'
import ExcelJS from 'exceljs'

// ─── Australian tax (2025–26 FY) ─────────────────────────────────────────────

const SUPER_RATE = 0.12 // SGC rate 2025–26

function calcAusTax(taxable) {
  let tax = 0
  if (taxable <= 18200) {
    tax = 0
  } else if (taxable <= 45000) {
    tax = (taxable - 18200) * 0.16
  } else if (taxable <= 135000) {
    tax = 4288 + (taxable - 45000) * 0.30
  } else if (taxable <= 190000) {
    tax = 31288 + (taxable - 135000) * 0.37
  } else {
    tax = 51638 + (taxable - 190000) * 0.45
  }

  // Low Income Tax Offset (LITO)
  let lito = 0
  if (taxable <= 37500) {
    lito = 700
  } else if (taxable <= 45000) {
    lito = 700 - (taxable - 37500) * 0.05
  } else if (taxable <= 66667) {
    lito = 325 - (taxable - 45000) * 0.015
  }
  tax = Math.max(0, tax - lito)

  // Medicare Levy: 2% above ~$27,222 (singles, 2025–26)
  const medicare = taxable > 27222 ? taxable * 0.02 : 0

  const totalTax = tax + medicare
  const takeHome = taxable - totalTax

  return { incomeTax: tax, medicare, totalTax, takeHome }
}

// Given what the user typed and whether it includes super, derive all figures
function calcIncome(entered, superType) {
  if (superType === 'including') {
    // entered = taxable salary + super, so taxable = entered / 1.12
    const taxable = entered / (1 + SUPER_RATE)
    const superAmt = entered - taxable
    return { taxable, superAmt, ...calcAusTax(taxable) }
  } else {
    // entered = taxable salary; super is employer's extra 12% on top
    const superAmt = entered * SUPER_RATE
    return { taxable: entered, superAmt, ...calcAusTax(entered) }
  }
}

// ─── helpers ─────────────────────────────────────────────────────────────────

const PERIODS = ['Weekly', 'Monthly', 'Yearly']

function toWeekly(annual)  { return annual / 52 }
function toMonthly(annual) { return annual / 12 }
function toYearly(annual)  { return annual }

function forPeriod(annual, period) {
  if (period === 'Weekly')  return toWeekly(annual)
  if (period === 'Monthly') return toMonthly(annual)
  return toYearly(annual)
}

function fmt(n) {
  return new Intl.NumberFormat('en-AU', {
    style: 'currency',
    currency: 'AUD',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(n)
}

let nextId = 100
function uid() { return nextId++ }

const DEFAULT_EXPENSES = [
  { id: uid(), category: 'Rent / Mortgage', annual: 18000 },
  { id: uid(), category: 'Groceries',        annual: 6000  },
  { id: uid(), category: 'Transport',         annual: 3600  },
  { id: uid(), category: 'Utilities',         annual: 2400  },
  { id: uid(), category: 'Entertainment',     annual: 1800  },
]

// ─── sub-components ──────────────────────────────────────────────────────────

function SummaryCard({ label, value, sub, highlight }) {
  const color =
    highlight === 'positive' ? 'text-emerald-600' :
    highlight === 'negative' ? 'text-red-500' :
    'text-gray-800'

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 flex flex-col gap-1">
      <span className="text-xs font-semibold uppercase tracking-widest text-gray-400">{label}</span>
      <span className={`text-2xl font-bold tabular-nums ${color}`}>{value}</span>
      {sub && <span className="text-xs text-gray-400 mt-0.5">{sub}</span>}
    </div>
  )
}

function TaxLine({ label, value, muted, bold, green }) {
  const textColor = green ? 'text-emerald-700 font-bold' : bold ? 'text-gray-800 font-semibold' : muted ? 'text-gray-400' : 'text-gray-700'
  return (
    <div className={`flex justify-between text-sm ${textColor}`}>
      <span>{label}</span>
      <span className="tabular-nums">{value}</span>
    </div>
  )
}

function ExpenseRow({ row, onChange, onDelete }) {
  return (
    <div className="flex items-center gap-3">
      <input
        type="text"
        value={row.category}
        onChange={e => onChange(row.id, 'category', e.target.value)}
        placeholder="Category name"
        className="flex-1 rounded-lg border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300 bg-white"
      />
      <div className="relative w-40">
        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm">$</span>
        <input
          type="number"
          min="0"
          step="1"
          value={row.annual === 0 ? '' : row.annual}
          onChange={e => onChange(row.id, 'annual', parseFloat(e.target.value) || 0)}
          placeholder="0"
          className="w-full rounded-lg border border-gray-200 pl-7 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300 bg-white tabular-nums"
        />
      </div>
      <button
        onClick={() => onDelete(row.id)}
        aria-label="Delete row"
        className="w-8 h-8 flex items-center justify-center rounded-full text-gray-300 hover:text-red-400 hover:bg-red-50 transition-colors"
      >
        ×
      </button>
    </div>
  )
}

// ─── Excel export ─────────────────────────────────────────────────────────────

async function downloadExcel(enteredSalary, superType, expenses) {
  const { taxable, superAmt, incomeTax, medicare, takeHome } = calcIncome(enteredSalary, superType)
  const totalExpenses = expenses.reduce((s, r) => s + r.annual, 0)
  const leftover = takeHome - totalExpenses

  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Budget Planner'
  const ws = workbook.addWorksheet('Budget')

  ws.columns = [
    { key: 'category', width: 32 },
    { key: 'weekly',   width: 14 },
    { key: 'monthly',  width: 14 },
    { key: 'yearly',   width: 16 },
  ]

  const currencyFmt = '"$"#,##0.00'

  function styleRow(row, { bg, fontColor = 'FF1F2937', bold = false }) {
    row.eachCell(cell => {
      cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      cell.font   = { bold, color: { argb: fontColor }, size: 11 }
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } } }
    })
  }

  function applyNumberFmt(row, startCol = 2) {
    for (let c = startCol; c <= 4; c++) {
      row.getCell(c).numFmt = currencyFmt
    }
  }

  // header
  const headerRow = ws.addRow(['Category', 'Weekly', 'Monthly', 'Yearly'])
  styleRow(headerRow, { bg: 'FF4F46E5', fontColor: 'FFFFFFFF', bold: true })
  headerRow.height = 22

  // income section header
  const incDiv = ws.addRow(['Income', '', '', ''])
  styleRow(incDiv, { bg: 'FFD1FAE5', bold: true })
  incDiv.height = 18

  if (superType === 'including') {
    // total package row
    const pkgRow = ws.addRow(['Total Package (inc. super)', toWeekly(enteredSalary), toMonthly(enteredSalary), toYearly(enteredSalary)])
    styleRow(pkgRow, { bg: 'FFECFDF5' })
    applyNumberFmt(pkgRow)
    pkgRow.height = 18

    // super deducted
    const supRow = ws.addRow(['Super (12% deducted)', toWeekly(-superAmt), toMonthly(-superAmt), toYearly(-superAmt)])
    styleRow(supRow, { bg: 'FFECFDF5' })
    applyNumberFmt(supRow)
    supRow.height = 18

    // taxable salary
    const txRow = ws.addRow(['Taxable Salary', toWeekly(taxable), toMonthly(taxable), toYearly(taxable)])
    styleRow(txRow, { bg: 'FFECFDF5', bold: true })
    applyNumberFmt(txRow)
    txRow.height = 18
  } else {
    // salary row
    const salRow = ws.addRow(['Gross Salary (excl. super)', toWeekly(taxable), toMonthly(taxable), toYearly(taxable)])
    styleRow(salRow, { bg: 'FFECFDF5', bold: true })
    applyNumberFmt(salRow)
    salRow.height = 18

    // employer super (informational)
    const supRow = ws.addRow(['Super (employer 12%, on top)', toWeekly(superAmt), toMonthly(superAmt), toYearly(superAmt)])
    styleRow(supRow, { bg: 'FFD1FAE5' })
    applyNumberFmt(supRow)
    supRow.height = 18
  }

  // tax rows
  const taxRow = ws.addRow(['Income Tax (est.)', toWeekly(-incomeTax), toMonthly(-incomeTax), toYearly(-incomeTax)])
  styleRow(taxRow, { bg: 'FFFEF2F2' })
  applyNumberFmt(taxRow)
  taxRow.height = 18

  const medRow = ws.addRow(['Medicare Levy (est.)', toWeekly(-medicare), toMonthly(-medicare), toYearly(-medicare)])
  styleRow(medRow, { bg: 'FFFEF2F2' })
  applyNumberFmt(medRow)
  medRow.height = 18

  // take-home
  const thRow = ws.addRow(['Take-home Pay', toWeekly(takeHome), toMonthly(takeHome), toYearly(takeHome)])
  styleRow(thRow, { bg: 'FFD1FAE5', bold: true })
  applyNumberFmt(thRow)
  thRow.height = 20

  // expenses
  const divRow = ws.addRow(['Expenses', '', '', ''])
  styleRow(divRow, { bg: 'FFFEF3C7', bold: true })
  divRow.height = 18

  expenses.forEach(r => {
    const row = ws.addRow([r.category, toWeekly(r.annual), toMonthly(r.annual), toYearly(r.annual)])
    applyNumberFmt(row)
    row.height = 18
    row.eachCell(cell => { cell.border = { bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } } } })
  })

  // leftover + savings rate
  const totalsRow = ws.addRow(['Leftover (Take-home − Expenses)', toWeekly(leftover), toMonthly(leftover), toYearly(leftover)])
  styleRow(totalsRow, { bg: 'FFF3F4F6', bold: true })
  applyNumberFmt(totalsRow)
  totalsRow.height = 22

  const savingsRow = ws.addRow([
    'Savings Rate',
    takeHome > 0 ? `${((leftover / takeHome) * 100).toFixed(1)}%` : '—',
    '', '',
  ])
  styleRow(savingsRow, { bg: 'FFF3F4F6' })
  savingsRow.height = 18

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }]

  const buffer = await workbook.xlsx.writeBuffer()
  const blob   = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url    = URL.createObjectURL(blob)
  const link   = document.createElement('a')
  link.href     = url
  link.download = 'budget.xlsx'
  link.click()
  URL.revokeObjectURL(url)
}

// ─── main app ────────────────────────────────────────────────────────────────

export default function App() {
  const [enteredSalary, setEnteredSalary] = useState(75000)
  const [superType, setSuperType]         = useState('excluding')
  const [expenses, setExpenses]           = useState(DEFAULT_EXPENSES)
  const [period, setPeriod]               = useState('Monthly')

  const { taxable, superAmt, incomeTax, medicare, totalTax, takeHome } = calcIncome(enteredSalary, superType)
  const effectiveRate = taxable > 0 ? ((totalTax / taxable) * 100).toFixed(1) : '0.0'

  const totalExpAnnual = expenses.reduce((s, r) => s + r.annual, 0)
  const leftoverAnnual = takeHome - totalExpAnnual
  const savingsRate    = takeHome > 0 ? ((leftoverAnnual / takeHome) * 100).toFixed(1) : '0.0'
  const leftoverHighlight = leftoverAnnual >= 0 ? 'positive' : 'negative'

  const handleChange = useCallback((id, field, value) => {
    setExpenses(prev => prev.map(r => r.id === id ? { ...r, [field]: value } : r))
  }, [])

  const handleDelete = useCallback((id) => {
    setExpenses(prev => prev.filter(r => r.id !== id))
  }, [])

  const handleAdd = () => {
    setExpenses(prev => [...prev, { id: uid(), category: '', annual: 0 }])
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">

      {/* ── header ── */}
      <header className="bg-white border-b border-gray-100 px-6 py-5 shadow-sm">
        <div className="max-w-3xl mx-auto">
          <h1 className="text-2xl font-bold text-gray-800 tracking-tight">Personal Budget Planner</h1>
          <p className="text-sm text-gray-400 mt-0.5">Australian 2025–26 tax estimates included.</p>
        </div>
      </header>

      <main className="flex-1 max-w-3xl w-full mx-auto px-4 py-8 flex flex-col gap-8">

        {/* ── period toggle ── */}
        <div className="flex justify-end">
          <div className="inline-flex bg-white border border-gray-200 rounded-xl p-1 shadow-sm gap-1">
            {PERIODS.map(p => (
              <button
                key={p}
                onClick={() => setPeriod(p)}
                className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-colors ${
                  period === p
                    ? 'bg-indigo-600 text-white shadow'
                    : 'text-gray-500 hover:text-gray-800'
                }`}
              >
                {p}
              </button>
            ))}
          </div>
        </div>

        {/* ── income section ── */}
        <section className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6 flex flex-col gap-5">
          <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600">Income</h2>

          {/* salary input + super dropdown */}
          <div className="flex flex-col gap-3">
            <div className="flex flex-col sm:flex-row sm:items-center gap-3">
              <label htmlFor="income-input" className="text-sm text-gray-600 sm:w-48 shrink-0">
                Annual Salary
              </label>
              <div className="relative w-44">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400">$</span>
                <input
                  id="income-input"
                  type="number"
                  min="0"
                  step="1000"
                  value={enteredSalary === 0 ? '' : enteredSalary}
                  onChange={e => setEnteredSalary(parseFloat(e.target.value) || 0)}
                  placeholder="e.g. 75000"
                  className="w-full rounded-lg border border-gray-200 pl-7 pr-3 py-2.5 text-base focus:outline-none focus:ring-2 focus:ring-indigo-300 tabular-nums"
                />
              </div>
              <select
                value={superType}
                onChange={e => setSuperType(e.target.value)}
                className="rounded-lg border border-gray-200 px-3 py-2.5 text-sm text-gray-700 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-300 cursor-pointer"
              >
                <option value="excluding">Excluding super</option>
                <option value="including">Including super</option>
              </select>
            </div>
            <p className="text-xs text-gray-400">
              {superType === 'excluding'
                ? 'Your employer pays an additional 12% super on top of this salary.'
                : 'Your salary package includes super — we\'ll subtract 12% to find your taxable income.'}
            </p>
          </div>

          {/* tax & super breakdown */}
          <div className="bg-gray-50 rounded-xl p-4 flex flex-col gap-2 border border-gray-100">
            <p className="text-xs font-semibold uppercase tracking-widest text-gray-400 mb-1">
              Breakdown — {period}
            </p>

            {superType === 'including' ? (
              <>
                <TaxLine label="Total Package (inc. super)"  value={fmt(forPeriod(enteredSalary, period))} />
                <TaxLine label="Super (12% deducted)"        value={`− ${fmt(forPeriod(superAmt, period))}`} muted />
                <TaxLine label="Taxable Salary"              value={fmt(forPeriod(taxable, period))} bold />
              </>
            ) : (
              <>
                <TaxLine label="Gross Salary (excl. super)"  value={fmt(forPeriod(taxable, period))} />
                <TaxLine label="Super (employer, 12% extra)" value={fmt(forPeriod(superAmt, period))} muted />
              </>
            )}

            <div className="border-t border-gray-200 my-1" />
            <TaxLine label="Income Tax (est.)"    value={`− ${fmt(forPeriod(incomeTax, period))}`} muted />
            <TaxLine label="Medicare Levy (est.)" value={`− ${fmt(forPeriod(medicare, period))}`} muted />
            <div className="border-t border-gray-200 my-1" />
            <TaxLine label="Take-home Pay" value={fmt(forPeriod(takeHome, period))} green />

            <p className="text-xs text-gray-400 mt-1">
              Effective tax rate: {effectiveRate}% &nbsp;·&nbsp; Includes LITO &nbsp;·&nbsp; 2025–26 estimate only
            </p>
          </div>
        </section>

        {/* ── expenses section ── */}
        <section className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6 flex flex-col gap-4">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600">Expenses</h2>
            <span className="text-xs text-gray-400">All amounts are yearly totals.</span>
          </div>

          <div className="flex items-center gap-3 text-xs font-semibold uppercase tracking-wider text-gray-400">
            <span className="flex-1">Category</span>
            <span className="w-40 pl-7">Annual Amount</span>
            <span className="w-8" />
          </div>

          <div className="flex flex-col gap-3">
            {expenses.map(row => (
              <ExpenseRow
                key={row.id}
                row={row}
                onChange={handleChange}
                onDelete={handleDelete}
              />
            ))}
          </div>

          <button
            onClick={handleAdd}
            className="mt-1 self-start text-sm text-indigo-600 hover:text-indigo-800 font-medium flex items-center gap-1 transition-colors"
          >
            <span className="text-lg leading-none">+</span> Add expense
          </button>
        </section>

        {/* ── summary cards ── */}
        <section className="flex flex-col gap-3">
          <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600">
            Summary — {period}
          </h2>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <SummaryCard
              label="Take-home"
              value={fmt(forPeriod(takeHome, period))}
              sub="after tax & super"
            />
            <SummaryCard
              label="Expenses"
              value={fmt(forPeriod(totalExpAnnual, period))}
            />
            <SummaryCard
              label="Leftover"
              value={fmt(forPeriod(leftoverAnnual, period))}
              highlight={leftoverHighlight}
            />
            <SummaryCard
              label="Savings Rate"
              value={`${savingsRate}%`}
              sub="of take-home"
              highlight={leftoverHighlight}
            />
          </div>
        </section>

        {/* ── export ── */}
        <div className="flex justify-end">
          <button
            onClick={() => downloadExcel(enteredSalary, superType, expenses)}
            className="bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white text-sm font-semibold px-5 py-2.5 rounded-xl shadow transition-all flex items-center gap-2"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5 5-5M12 15V3" />
            </svg>
            Download as Excel
          </button>
        </div>

      </main>

      <footer className="text-center text-xs text-gray-400 py-6 border-t border-gray-100">
        For budgeting and illustration only — not financial advice. Tax estimates based on 2025–26 ATO rates for Australian residents.
      </footer>

    </div>
  )
}

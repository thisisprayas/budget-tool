import { useState, useCallback, useEffect } from 'react'
import ExcelJS from 'exceljs'
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, AreaChart, Area, XAxis, YAxis, CartesianGrid, ReferenceLine } from 'recharts'

// ─── Australian tax (2025–26 FY) ─────────────────────────────────────────────

const SUPER_RATE = 0.12

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

  let lito = 0
  if (taxable <= 37500) {
    lito = 700
  } else if (taxable <= 45000) {
    lito = 700 - (taxable - 37500) * 0.05
  } else if (taxable <= 66667) {
    lito = 325 - (taxable - 45000) * 0.015
  }
  tax = Math.max(0, tax - lito)

  const medicare = taxable > 27222 ? taxable * 0.02 : 0
  const totalTax = tax + medicare
  const takeHome = taxable - totalTax

  return { incomeTax: tax, medicare, totalTax, takeHome }
}

function calcIncome(entered, superType) {
  if (superType === 'including') {
    const taxable = entered / (1 + SUPER_RATE)
    const superAmt = entered - taxable
    return { taxable, superAmt, ...calcAusTax(taxable) }
  } else {
    const superAmt = entered * SUPER_RATE
    return { taxable: entered, superAmt, ...calcAusTax(entered) }
  }
}

// ─── helpers ─────────────────────────────────────────────────────────────────

const PERIODS = ['Weekly', 'Monthly', 'Yearly']
const periodDivisors = { Weekly: 52, Monthly: 12, Yearly: 1 }

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
  { id: uid(), category: 'Subscriptions',     annual: 600   },
  { id: uid(), category: 'Entertainment',     annual: 1800  },
  { id: uid(), category: 'Miscellaneous',     annual: 1200  },
]

// ─── sub-components ──────────────────────────────────────────────────────────

function SummaryCard({ label, value, sub, highlight }) {
  const color =
    highlight === 'positive' ? 'text-emerald-500' :
    highlight === 'negative' ? 'text-red-500' :
    'text-gray-800 dark:text-gray-100'

  return (
    <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700 p-5 flex flex-col gap-1">
      <span className="text-xs font-semibold uppercase tracking-widest text-gray-400">{label}</span>
      <span className={`text-2xl font-bold tabular-nums ${color}`}>{value}</span>
      {sub && <span className="text-xs text-gray-400 mt-0.5">{sub}</span>}
    </div>
  )
}

function TaxLine({ label, value, muted, bold, green }) {
  const textColor = green
    ? 'text-emerald-600 dark:text-emerald-400 font-bold'
    : bold
    ? 'text-gray-800 dark:text-gray-100 font-semibold'
    : muted
    ? 'text-gray-400 dark:text-gray-500'
    : 'text-gray-700 dark:text-gray-300'
  return (
    <div className={`flex justify-between text-sm ${textColor}`}>
      <span>{label}</span>
      <span className="tabular-nums">{value}</span>
    </div>
  )
}

function ExpenseRow({ row, onChange, onDelete, period, takeHome }) {
  const displayValue = row.annual === 0 ? '' : parseFloat((row.annual / periodDivisors[period]).toFixed(2))
  const pct = takeHome > 0 ? ((row.annual / takeHome) * 100).toFixed(1) : '0.0'

  return (
    <div className="flex items-center gap-3">
      <input
        type="text"
        value={row.category}
        onChange={e => onChange(row.id, 'category', e.target.value)}
        placeholder="Category name"
        className="flex-1 rounded-lg border border-gray-200 dark:border-gray-600 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300 bg-white dark:bg-gray-700 text-gray-800 dark:text-gray-100 placeholder-gray-400"
      />
      <div className="relative w-40">
        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm">$</span>
        <input
          type="number"
          min="0"
          step="1"
          value={displayValue}
          onChange={e => onChange(row.id, 'annual', (parseFloat(e.target.value) || 0) * periodDivisors[period])}
          placeholder="0"
          className="w-full rounded-lg border border-gray-200 dark:border-gray-600 pl-7 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300 bg-white dark:bg-gray-700 text-gray-800 dark:text-gray-100 tabular-nums"
        />
      </div>
      <span className="w-14 text-right text-xs tabular-nums text-gray-400">{pct}%</span>
      <button
        onClick={() => onDelete(row.id)}
        aria-label="Delete row"
        className="w-8 h-8 flex items-center justify-center rounded-full text-gray-300 dark:text-gray-600 hover:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"
      >
        ×
      </button>
    </div>
  )
}

// ─── Donut chart ─────────────────────────────────────────────────────────────

const CHART_COLORS = ['#6366F1', '#8B5CF6', '#A78BFA', '#EC4899', '#F59E0B', '#10B981', '#3B82F6', '#F97316']

function ExpenseDonut({ expenses, takeHome, period }) {
  const data = expenses.filter(r => r.annual > 0).map(r => ({
    name: r.category || 'Unnamed',
    value: r.annual,
  }))

  if (data.length === 0) return null

  return (
    <section className="bg-white dark:bg-gray-800 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700 p-6">
      <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600 dark:text-indigo-400 mb-6">Expense Breakdown</h2>
      <div className="flex flex-col sm:flex-row items-center gap-8">
        <div className="w-[200px] h-[200px] shrink-0">
          <ResponsiveContainer width="100%" height="100%">
            <PieChart>
              <Pie
                data={data}
                cx="50%"
                cy="50%"
                innerRadius={60}
                outerRadius={95}
                paddingAngle={2}
                dataKey="value"
              >
                {data.map((_, i) => (
                  <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />
                ))}
              </Pie>
              <Tooltip
                formatter={(value, name) => [fmt(forPeriod(value, period)), name]}
                contentStyle={{ borderRadius: '8px', border: '1px solid #E5E7EB', fontSize: '12px' }}
              />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div className="flex flex-col gap-2.5 flex-1 w-full">
          {data.map((d, i) => (
            <div key={i} className="flex items-center gap-2.5 text-sm">
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: CHART_COLORS[i % CHART_COLORS.length] }} />
              <span className="flex-1 text-gray-700 dark:text-gray-300 truncate">{d.name}</span>
              <span className="tabular-nums text-gray-400 text-xs">
                {takeHome > 0 ? `${((d.value / takeHome) * 100).toFixed(1)}%` : '—'}
              </span>
            </div>
          ))}
        </div>
      </div>
    </section>
  )
}

// ─── Savings Explorer ────────────────────────────────────────────────────────

function SavingsExplorer({ currentSalary, superType, totalExpAnnual, period }) {
  const [sliderSalary, setSliderSalary] = useState(currentSalary)
  const [expMultiplier, setExpMultiplier] = useState(100)

  useEffect(() => { setSliderSalary(currentSalary) }, [currentSalary])
  useEffect(() => { setExpMultiplier(100) }, [totalExpAnnual])

  const maxSalary = Math.max(250000, Math.ceil(currentSalary * 2 / 50000) * 50000)
  const scaledExpenses = totalExpAnnual * (expMultiplier / 100)

  const data = Array.from({ length: 51 }, (_, i) => {
    const salary = Math.round((i / 50) * maxSalary)
    const { takeHome } = calcIncome(salary, superType)
    return { salary, savings: forPeriod(takeHome - scaledExpenses, period) }
  })

  const { takeHome: sliderTakeHome } = calcIncome(sliderSalary, superType)
  const sliderSavings = forPeriod(sliderTakeHome - scaledExpenses, period)

  const { takeHome: actualTakeHome } = calcIncome(currentSalary, superType)
  const actualSavings = forPeriod(actualTakeHome - totalExpAnnual, period)
  const diff = sliderSavings - actualSavings

  return (
    <section className="bg-white dark:bg-gray-800 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700 p-6 flex flex-col gap-5">
      <div className="flex items-center justify-between">
        <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600 dark:text-indigo-400">Savings Explorer</h2>
        {(sliderSalary !== currentSalary || expMultiplier !== 100) && (
          <button
            onClick={() => { setSliderSalary(currentSalary); setExpMultiplier(100) }}
            className="text-xs text-gray-400 hover:text-indigo-600 dark:hover:text-indigo-400 transition-colors"
          >
            Reset to current
          </button>
        )}
      </div>

      <div className="flex flex-col gap-4">
        <div className="flex flex-col gap-1.5">
          <div className="flex justify-between text-sm">
            <span className="text-gray-600 dark:text-gray-300">Annual Salary</span>
            <span className="tabular-nums font-medium text-gray-800 dark:text-gray-100">{fmt(sliderSalary)}</span>
          </div>
          <input
            type="range" min={0} max={maxSalary} step={1000}
            value={sliderSalary}
            onChange={e => setSliderSalary(Number(e.target.value))}
            className="w-full accent-indigo-600"
          />
          <div className="flex justify-between text-xs text-gray-400">
            <span>$0</span><span>{fmt(maxSalary)}</span>
          </div>
        </div>

        <div className="flex flex-col gap-1.5">
          <div className="flex justify-between text-sm">
            <span className="text-gray-600 dark:text-gray-300">Expenses</span>
            <span className="tabular-nums font-medium text-gray-800 dark:text-gray-100">
              {fmt(forPeriod(scaledExpenses, period))}{' '}
              <span className="text-gray-400 font-normal">({expMultiplier}% of current)</span>
            </span>
          </div>
          <input
            type="range" min={50} max={200} step={1}
            value={expMultiplier}
            onChange={e => setExpMultiplier(Number(e.target.value))}
            className="w-full accent-indigo-600"
          />
          <div className="flex justify-between text-xs text-gray-400">
            <span>50% of current</span><span>200% of current</span>
          </div>
        </div>
      </div>

      <div className="h-[220px]">
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={data} margin={{ top: 8, right: 8, left: 8, bottom: 0 }}>
            <defs>
              <linearGradient id="savingsGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#6366F1" stopOpacity={0.2} />
                <stop offset="95%" stopColor="#6366F1" stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6" vertical={false} />
            <XAxis
              dataKey="salary"
              tickFormatter={v => `$${(v / 1000).toFixed(0)}k`}
              tick={{ fontSize: 11, fill: '#9CA3AF' }}
              axisLine={false} tickLine={false} tickCount={6}
            />
            <YAxis
              tickFormatter={v => `$${(v / 1000).toFixed(0)}k`}
              tick={{ fontSize: 11, fill: '#9CA3AF' }}
              axisLine={false} tickLine={false} width={48}
            />
            <Tooltip
              formatter={(value) => [fmt(value), `${period} Savings`]}
              labelFormatter={(salary) => `Salary: ${fmt(salary)}/yr`}
              contentStyle={{ borderRadius: '8px', border: '1px solid #E5E7EB', fontSize: '12px' }}
            />
            <ReferenceLine y={0} stroke="#E5E7EB" strokeWidth={1} />
            <ReferenceLine x={sliderSalary} stroke="#6366F1" strokeDasharray="4 4" strokeWidth={1.5} />
            <Area
              type="monotone" dataKey="savings"
              stroke="#6366F1" strokeWidth={2}
              fill="url(#savingsGrad)"
              dot={false} activeDot={{ r: 4, fill: '#6366F1' }}
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>

      <div className="bg-gray-50 dark:bg-gray-700/50 rounded-xl p-3 flex justify-between items-center text-sm gap-4">
        <span className="text-gray-500 dark:text-gray-400">
          {fmt(sliderSalary)}/yr salary · {expMultiplier}% expenses
        </span>
        <div className="flex items-center gap-2 shrink-0">
          <span className={`font-semibold tabular-nums ${sliderSavings >= 0 ? 'text-emerald-500' : 'text-red-500'}`}>
            {fmt(sliderSavings)} {period.toLowerCase()} savings
          </span>
          {diff !== 0 && (
            <span className={`text-xs tabular-nums px-1.5 py-0.5 rounded-md font-medium ${diff > 0 ? 'bg-emerald-50 text-emerald-600 dark:bg-emerald-900/20 dark:text-emerald-400' : 'bg-red-50 text-red-500 dark:bg-red-900/20 dark:text-red-400'}`}>
              {diff > 0 ? '+' : ''}{fmt(diff)}
            </span>
          )}
        </div>
      </div>
    </section>
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
    { key: 'category', width: 34 },
    { key: 'weekly',   width: 14 },
    { key: 'monthly',  width: 14 },
    { key: 'yearly',   width: 16 },
    { key: 'pct',      width: 14 },
  ]

  const currencyFmt = '"$"#,##0.00'
  const GREY_DARK   = 'FF374151'
  const GREY_MID    = 'FF6B7280'
  const GREY_BG     = 'FFF3F4F6'
  const GREY_ALT    = 'FFF9FAFB'
  const WHITE       = 'FFFFFFFF'
  const BORDER_COL  = 'FFE5E7EB'

  function applyBorder(row) {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = { bottom: { style: 'thin', color: { argb: BORDER_COL } } }
    })
  }

  function styleHeader(row) {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY_BG } }
      cell.font  = { bold: true, color: { argb: GREY_DARK }, size: 11 }
      cell.border = { bottom: { style: 'medium', color: { argb: BORDER_COL } } }
    })
    row.height = 22
  }

  function styleSection(row) {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY_BG } }
      cell.font = { bold: true, color: { argb: GREY_MID }, size: 10 }
      cell.border = { bottom: { style: 'thin', color: { argb: BORDER_COL } } }
    })
    row.height = 18
  }

  function styleData(row, { bold = false, bg = WHITE } = {}) {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      cell.font  = { bold, color: { argb: bold ? GREY_DARK : GREY_MID }, size: 11 }
      cell.border = { bottom: { style: 'thin', color: { argb: BORDER_COL } } }
    })
    row.height = 18
  }

  function applyCurrencyFmt(row, cols = [2, 3, 4]) {
    cols.forEach(c => { row.getCell(c).numFmt = currencyFmt })
  }

  // ── header ──
  const headerRow = ws.addRow(['', 'Weekly', 'Monthly', 'Yearly', '% of Take-home'])
  styleHeader(headerRow)
  headerRow.getCell(1).value = ''

  // ── income section ──
  const incSection = ws.addRow(['INCOME', '', '', '', ''])
  styleSection(incSection)

  if (superType === 'including') {
    const pkgRow = ws.addRow(['Total Package (inc. super)', toWeekly(enteredSalary), toMonthly(enteredSalary), toYearly(enteredSalary), ''])
    styleData(pkgRow)
    applyCurrencyFmt(pkgRow)

    const supRow = ws.addRow(['Superannuation', toWeekly(-superAmt), toMonthly(-superAmt), toYearly(-superAmt), ''])
    styleData(supRow, { bg: GREY_ALT })
    applyCurrencyFmt(supRow)

    const txRow = ws.addRow(['Taxable Income', toWeekly(taxable), toMonthly(taxable), toYearly(taxable), ''])
    styleData(txRow, { bold: true })
    applyCurrencyFmt(txRow)
  } else {
    const salRow = ws.addRow(['Gross Salary (excl. super)', toWeekly(taxable), toMonthly(taxable), toYearly(taxable), ''])
    styleData(salRow, { bold: true })
    applyCurrencyFmt(salRow)

    const supRow = ws.addRow(['Super (employer, 12% extra)', toWeekly(superAmt), toMonthly(superAmt), toYearly(superAmt), ''])
    styleData(supRow, { bg: GREY_ALT })
    applyCurrencyFmt(supRow)
  }

  const taxRow = ws.addRow(['Income Tax', toWeekly(-incomeTax), toMonthly(-incomeTax), toYearly(-incomeTax), ''])
  styleData(taxRow)
  applyCurrencyFmt(taxRow)

  const medRow = ws.addRow(['Medicare', toWeekly(-medicare), toMonthly(-medicare), toYearly(-medicare), ''])
  styleData(medRow, { bg: GREY_ALT })
  applyCurrencyFmt(medRow)

  const thRow = ws.addRow(['Take-home Pay', toWeekly(takeHome), toMonthly(takeHome), toYearly(takeHome), ''])
  styleData(thRow, { bold: true })
  applyCurrencyFmt(thRow)
  thRow.getCell(1).border = { top: { style: 'medium', color: { argb: BORDER_COL } }, bottom: { style: 'medium', color: { argb: BORDER_COL } } }

  // ── expenses section ──
  const expSection = ws.addRow(['EXPENSES', '', '', '', ''])
  styleSection(expSection)

  expenses.forEach((r, i) => {
    const pct = takeHome > 0 ? `${((r.annual / takeHome) * 100).toFixed(1)}%` : '—'
    const row = ws.addRow([r.category, toWeekly(r.annual), toMonthly(r.annual), toYearly(r.annual), pct])
    styleData(row, { bg: i % 2 === 1 ? GREY_ALT : WHITE })
    applyCurrencyFmt(row)
  })

  // ── summary ──
  const sumSection = ws.addRow(['SUMMARY', '', '', '', ''])
  styleSection(sumSection)

  const leftoverRow = ws.addRow(['Savings (Take-home − Expenses)', toWeekly(leftover), toMonthly(leftover), toYearly(leftover), ''])
  styleData(leftoverRow, { bold: true })
  applyCurrencyFmt(leftoverRow)

  const savingsRow = ws.addRow([
    'Savings Rate',
    takeHome > 0 ? `${((leftover / takeHome) * 100).toFixed(1)}%` : '—',
    '', '', '',
  ])
  styleData(savingsRow, { bg: GREY_ALT })

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
  const [darkMode, setDarkMode]           = useState(false)

  const { taxable, superAmt, incomeTax, medicare, totalTax, takeHome } = calcIncome(enteredSalary, superType)
  const effectiveRate  = taxable > 0 ? ((totalTax / taxable) * 100).toFixed(1) : '0.0'
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
    <div className={darkMode ? 'dark' : ''}>
      <div className="min-h-screen bg-gray-50 dark:bg-gray-900 flex flex-col transition-colors duration-200">

        {/* ── header ── */}
        <header className="bg-white dark:bg-gray-800 border-b border-gray-100 dark:border-gray-700 px-6 py-5 shadow-sm">
          <div className="max-w-3xl mx-auto flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold text-gray-800 dark:text-gray-100 tracking-tight">Personal Budget</h1>
              <p className="text-sm text-gray-400 mt-0.5">Happy budgeting :)</p>
            </div>
            <button
              onClick={() => setDarkMode(d => !d)}
              aria-label="Toggle dark mode"
              className="w-9 h-9 flex items-center justify-center rounded-full text-gray-400 hover:text-gray-600 dark:text-gray-400 dark:hover:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors"
            >
              {darkMode ? (
                // sun icon
                <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364-6.364l-.707.707M6.343 17.657l-.707.707M17.657 17.657l-.707-.707M6.343 6.343l-.707-.707M12 8a4 4 0 100 8 4 4 0 000-8z" />
                </svg>
              ) : (
                // moon icon
                <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" />
                </svg>
              )}
            </button>
          </div>
        </header>

        <main className="flex-1 max-w-3xl w-full mx-auto px-4 py-8 flex flex-col gap-8">

          {/* ── period toggle ── */}
          <div className="flex justify-end">
            <div className="inline-flex bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl p-1 shadow-sm gap-1">
              {PERIODS.map(p => (
                <button
                  key={p}
                  onClick={() => setPeriod(p)}
                  className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-colors ${
                    period === p
                      ? 'bg-indigo-600 text-white shadow'
                      : 'text-gray-500 dark:text-gray-400 hover:text-gray-800 dark:hover:text-gray-100'
                  }`}
                >
                  {p}
                </button>
              ))}
            </div>
          </div>

          {/* ── income section ── */}
          <section className="bg-white dark:bg-gray-800 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700 p-6 flex flex-col gap-5">
            <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600 dark:text-indigo-400">Income</h2>

            <div className="flex flex-col sm:flex-row sm:items-center gap-3">
              <label htmlFor="income-input" className="text-sm text-gray-600 dark:text-gray-300 sm:w-48 shrink-0">
                Annual Salary
              </label>
              <div className="relative w-44">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm">$</span>
                <input
                  id="income-input"
                  type="number"
                  min="0"
                  step="1000"
                  value={enteredSalary === 0 ? '' : enteredSalary}
                  onChange={e => setEnteredSalary(parseFloat(e.target.value) || 0)}
                  placeholder="e.g. 75000"
                  className="w-full rounded-lg border border-gray-200 dark:border-gray-600 pl-7 pr-3 py-2.5 text-base focus:outline-none focus:ring-2 focus:ring-indigo-300 tabular-nums bg-white dark:bg-gray-700 text-gray-800 dark:text-gray-100"
                />
              </div>
              <select
                value={superType}
                onChange={e => setSuperType(e.target.value)}
                className="rounded-lg border border-gray-200 dark:border-gray-600 px-3 py-2.5 text-sm text-gray-700 dark:text-gray-200 bg-white dark:bg-gray-700 focus:outline-none focus:ring-2 focus:ring-indigo-300 cursor-pointer"
              >
                <option value="excluding">Excluding super</option>
                <option value="including">Including super</option>
              </select>
            </div>

            {/* breakdown */}
            <div className="bg-gray-50 dark:bg-gray-700/50 rounded-xl p-4 flex flex-col gap-2 border border-gray-100 dark:border-gray-600">
              <p className="text-xs font-semibold uppercase tracking-widest text-gray-400 mb-1">
                Breakdown — {period}
              </p>

              {superType === 'including' ? (
                <>
                  <TaxLine label="Total Package (inc. super)" value={fmt(forPeriod(enteredSalary, period))} />
                  <TaxLine label="Superannuation"             value={`(${fmt(forPeriod(superAmt, period))})`} muted />
                  <TaxLine label="Taxable Income"             value={fmt(forPeriod(taxable, period))} bold />
                </>
              ) : (
                <>
                  <TaxLine label="Gross Salary (excl. super)"  value={fmt(forPeriod(taxable, period))} />
                  <TaxLine label="Super (employer, 12% extra)" value={fmt(forPeriod(superAmt, period))} muted />
                </>
              )}

              <div className="border-t border-gray-200 dark:border-gray-600 my-1" />
              <TaxLine label="Income Tax"    value={`(${fmt(forPeriod(incomeTax, period))})`} muted />
              <TaxLine label="Medicare" value={`(${fmt(forPeriod(medicare, period))})`} muted />
              <div className="border-t border-gray-200 dark:border-gray-600 my-1" />
              <TaxLine label="Take-home Pay" value={fmt(forPeriod(takeHome, period))} green />

              <p className="text-xs text-gray-400 mt-1">Effective tax rate: {effectiveRate}%</p>
            </div>
          </section>

          {/* ── expenses section ── */}
          <section className="bg-white dark:bg-gray-800 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700 p-6 flex flex-col gap-4">
            <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600 dark:text-indigo-400">Expenses</h2>

            <div className="flex items-center gap-3 text-xs font-semibold uppercase tracking-wider text-gray-400">
              <span className="flex-1">Category</span>
              <span className="w-40 pl-7">{period} Amount</span>
              <span className="w-14 text-right">% of Pay</span>
              <span className="w-8" />
            </div>

            <div className="flex flex-col gap-3">
              {expenses.map(row => (
                <ExpenseRow
                  key={row.id}
                  row={row}
                  period={period}
                  onChange={handleChange}
                  onDelete={handleDelete}
                  takeHome={takeHome}
                />
              ))}
            </div>

            <button
              onClick={handleAdd}
              className="mt-1 self-start text-sm text-indigo-600 dark:text-indigo-400 hover:text-indigo-800 dark:hover:text-indigo-300 font-medium flex items-center gap-1 transition-colors"
            >
              <span className="text-lg leading-none">+</span> Add expense
            </button>
          </section>

          {/* ── summary cards ── */}
          <section className="flex flex-col gap-3">
            <h2 className="text-sm font-semibold uppercase tracking-widest text-indigo-600 dark:text-indigo-400">
              Summary — {period}
            </h2>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <SummaryCard label="Take-home"   value={fmt(forPeriod(takeHome, period))}       sub="after tax & super" />
              <SummaryCard label="Expenses"    value={fmt(forPeriod(totalExpAnnual, period))} />
              <SummaryCard label="Savings"    value={fmt(forPeriod(leftoverAnnual, period))} highlight={leftoverHighlight} />
              <SummaryCard label="Savings Rate" value={`${savingsRate}%`}                     sub="of take-home" highlight={leftoverHighlight} />
            </div>
          </section>

          {/* ── donut chart ── */}
          <ExpenseDonut expenses={expenses} takeHome={takeHome} period={period} />

          {/* ── savings explorer ── */}
          <SavingsExplorer
            currentSalary={enteredSalary}
            superType={superType}
            totalExpAnnual={totalExpAnnual}
            period={period}
          />

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

        <footer className="text-center text-xs text-gray-400 py-6 border-t border-gray-100 dark:border-gray-700">
          For internal budgeting only — not financial advice. Tax estimates based on 2025–26 ATO rates for Australian residents.
        </footer>

      </div>
    </div>
  )
}

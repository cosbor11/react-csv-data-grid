// src/CsvDataGrid.tsx
'use client'

import React, {
  useState,
  useMemo,
  startTransition,
  useEffect,
  useCallback,
  useRef,
} from 'react'
import {
  ChevronDown,
  ChevronUp,
  X,
  Pencil,
  Download,
  Trash2,
  Save as SaveIcon,
  Check,
  GripVertical,
  Plus,
} from 'lucide-react'

export interface CsvTable {
  header: string[] | null
  rows: string[][]
}

interface RowWithIndex {
  index: number
  row: string[]
}

interface ActiveCell {
  rowIndex: number
  colIndex: number
}

interface FillDragState {
  startRowIndex: number
  endRowIndex: number
  colIndex: number
  visibleRowIndices: number[]
}

interface ResizeState {
  colIndex: number
  startX: number
  startWidth: number
}

interface ResizeIntent {
  colIndex: number
  isActive: boolean
}

interface SortSnapshotEntry {
  text: string
  numeric: number | null
}

interface RowContextMenuState {
  x: number
  y: number
  rowIndex: number | null
}

type ColumnTotal =
  | { type: 'number'; value: number }
  | { type: 'currency'; value: number; symbol: string | null }
  | null

type SaveResult = boolean | void | Promise<boolean | void>

export interface CsvDataGridProps {
  content?: string | null
  fileName?: string
  delimiter?: string
  dirty?: boolean
  columnStorageKey?: string | null
  // Use className for host typography/layout utilities and style for CSS custom
  // properties like --csv-surface-toolbar or --csv-text-primary.
  className?: string
  style?: React.CSSProperties
  onContentChange(nextContent: string): void
  onSave?(): SaveResult
  onEdit?(): void
  onDownload?(): void
  onClose?(): void
}

const DEFAULT_COLUMN_WIDTH = 160
const MIN_COLUMN_WIDTH = 80
const MAX_COLUMN_WIDTH = 520
const RESIZE_HIT_WIDTH = 6
const COUNT_COL_WIDTH = 56
const ACTION_COL_WIDTH = 52
// Hosts should override the public --csv-* tokens. The internal --csv-theme-*
// variables let the component keep stable implementation details while still
// inheriting colors and fonts from the parent app.
const CSV_THEME_VARS: Record<string, string> = {
  '--csv-theme-font-sans':
    'var(--csv-font-sans, var(--font-sans, Arial, Helvetica, sans-serif))',
  '--csv-theme-font-mono':
    'var(--csv-font-mono, var(--font-mono, ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace))',
  '--csv-theme-surface-toolbar': 'var(--csv-surface-toolbar, #181818)',
  '--csv-theme-surface-muted': 'var(--csv-surface-muted, #232323)',
  '--csv-theme-surface-muted-hover':
    'var(--csv-surface-muted-hover, #333333)',
  '--csv-theme-surface-row-even': 'var(--csv-surface-row-even, #252526)',
  '--csv-theme-surface-row-odd': 'var(--csv-surface-row-odd, #232323)',
  '--csv-theme-surface-selected': 'var(--csv-surface-selected, #313131)',
  '--csv-theme-surface-overlay':
    'var(--csv-surface-overlay, rgba(30, 30, 30, 0.8))',
  '--csv-theme-surface-accent': 'var(--csv-surface-accent, #3794ff)',
  '--csv-theme-surface-accent-soft':
    'var(--csv-surface-accent-soft, color-mix(in srgb, var(--csv-surface-accent, #3794ff) 20%, transparent))',
  '--csv-theme-border-primary': 'var(--csv-border-primary, #333333)',
  '--csv-theme-border-toolbar': 'var(--csv-border-toolbar, #2d2d2d)',
  '--csv-theme-border-secondary': 'var(--csv-border-secondary, #555555)',
  '--csv-theme-border-accent': 'var(--csv-border-accent, #3794ff)',
  '--csv-theme-text-primary': 'var(--csv-text-primary, #d4d4d4)',
  '--csv-theme-text-muted': 'var(--csv-text-muted, #9ca3af)',
  '--csv-theme-text-subtle': 'var(--csv-text-subtle, #d1d5db)',
  '--csv-theme-text-danger': 'var(--csv-text-danger, #fca5a5)',
  '--csv-theme-text-success': 'var(--csv-text-success, #4ade80)',
  '--csv-theme-text-accent': 'var(--csv-text-accent, #3794ff)',
  '--csv-theme-bg-checkbox': 'var(--csv-bg-checkbox, #1f1f1f)',
}

const normalizeColumnWidth = (width?: number) => {
  if (typeof width !== 'number' || !Number.isFinite(width)) {
    return DEFAULT_COLUMN_WIDTH
  }
  return Math.min(MAX_COLUMN_WIDTH, Math.max(MIN_COLUMN_WIDTH, width))
}

function readPersistedValue(key: string) {
  try {
    return typeof globalThis.localStorage?.getItem === 'function'
      ? globalThis.localStorage.getItem(key)
      : null
  } catch {
    return null
  }
}

function writePersistedValue(key: string, value: string) {
  try {
    if (typeof globalThis.localStorage?.setItem === 'function') {
      globalThis.localStorage.setItem(key, value)
    }
  } catch {}
}

export function detectCsvDelimiterFromFileName(fileName?: string) {
  return fileName?.toLowerCase().endsWith('.tsv') ? '\t' : ','
}

function parseCsvLine(line: string, delimiter: string) {
  const cells: string[] = []
  let current = ''
  let inQuotes = false

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i]
    if (inQuotes) {
      if (char === '"') {
        const nextChar = line[i + 1]
        if (nextChar === '"') {
          current += '"'
          i += 1
        } else {
          inQuotes = false
        }
      } else {
        current += char
      }
    } else if (char === '"') {
      inQuotes = true
    } else if (char === delimiter) {
      cells.push(current)
      current = ''
    } else {
      current += char
    }
  }

  cells.push(current)
  return cells
}

function parseRows(content: string, delimiter: string) {
  if (!content) return []
  return content.split('\n').map((row) =>
    parseCsvLine(row.replace(/\r$/, ''), delimiter),
  )
}

function escapeCsvValue(value: string, delimiter: string) {
  if (value.includes('"')) {
    const escaped = value.replace(/"/g, '""')
    return `"${escaped}"`
  }
  if (value.includes(delimiter) || value.includes('\n') || value.includes('\r')) {
    return `"${value}"`
  }
  return value
}

function serializeTable(table: CsvTable, delimiter: string) {
  const rows = table.header ? [table.header, ...table.rows] : table.rows
  return rows
    .map((row) => row.map((cell) => escapeCsvValue(cell, delimiter)).join(delimiter))
    .join('\n')
}

function buildHeader(maxCols: number) {
  return Array.from({ length: maxCols }, (_, i) =>
    String.fromCharCode(65 + (i % 26)),
  )
}

function buildRange(start: number, end: number) {
  const min = Math.min(start, end)
  const max = Math.max(start, end)
  return Array.from({ length: max - min + 1 }, (_, i) => min + i)
}

function parseNumericValue(
  raw: string | undefined,
): { value: number; currencySymbol: string | null } | null {
  if (raw === undefined || raw === null) return null
  const trimmed = raw.trim()
  if (!trimmed) return null
  const isLikelyDate = /^(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/.test(
    trimmed,
  )
  if (isLikelyDate) return null

  const symbolMatch = trimmed.match(/[$€£¥]/)
  const currencySymbol = symbolMatch ? symbolMatch[0] : null
  const numericPart = trimmed.replace(/[^\d+-.]/g, '')
  if (!numericPart || numericPart === '.' || numericPart === '-' || numericPart === '+') {
    return null
  }
  const value = Number(numericPart.replace(/,/g, ''))
  if (!Number.isFinite(value)) return null
  return { value, currencySymbol }
}

function formatNumericTotal(total: ColumnTotal) {
  if (!total) return ''
  const formatted = total.value.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  })
  if (total.type === 'currency') {
    return `${total.symbol ?? '$'}${formatted}`
  }
  return formatted
}

export default function CsvDataGrid({
  content,
  fileName,
  delimiter,
  dirty = false,
  columnStorageKey = null,
  className,
  style,
  onContentChange,
  onSave,
  onEdit,
  onDownload,
  onClose,
}: CsvDataGridProps) {
  const [search, setSearch] = useState('')
  const [columnFilters, setColumnFilters] = useState<string[]>([])
  const [sortCol, setSortCol] = useState<number | null>(null)
  const [sortAsc, setSortAsc] = useState(true)
  const [switching, setSwitching] = useState(false) // overlay flag
  const [table, setTable] = useState<CsvTable | null>(null)
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set())
  const [selectedCols, setSelectedCols] = useState<Set<number>>(new Set())
  const [activeCell, setActiveCell] = useState<ActiveCell | null>(null)
  const [fillDrag, setFillDrag] = useState<FillDragState | null>(null)
  const [rowSelectionAnchor, setRowSelectionAnchor] = useState<number | null>(
    null,
  )
  const [lastRowSelection, setLastRowSelection] = useState<number | null>(null)
  const [lastColSelection, setLastColSelection] = useState<number | null>(null)
  const [justSaved, setJustSaved] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [columnWidths, setColumnWidths] = useState<number[]>([])
  const [resizeState, setResizeState] = useState<ResizeState | null>(null)
  const [resizeIntent, setResizeIntent] = useState<ResizeIntent | null>(null)
  const [draggingCol, setDraggingCol] = useState<number | null>(null)
  const [tableContainerWidth, setTableContainerWidth] = useState(0)
  const [hasManualResize, setHasManualResize] = useState(false)
  const [rowContextMenu, setRowContextMenu] =
    useState<RowContextMenuState | null>(null)
  const tableContainerRef = useRef<HTMLDivElement | null>(null)
  const rowMenuRef = useRef<HTMLDivElement | null>(null)
  const historyRef = useRef<CsvTable[]>([])
  const historyIndexRef = useRef(-1)
  const skipParseRef = useRef(false)
  const resolvedDelimiter = delimiter ?? detectCsvDelimiterFromFileName(fileName)
  const formatLabel = resolvedDelimiter === '\t' ? 'TSV preview' : 'CSV preview'

  useEffect(() => {
    if (content === undefined || content === null) {
      setTable(null)
      historyRef.current = []
      historyIndexRef.current = -1
      return
    }
    if (skipParseRef.current) {
      skipParseRef.current = false
      return
    }
    const rows = parseRows(content, resolvedDelimiter)
    const header = rows.length > 0 ? rows[0] : null
    const dataRows = rows.length > 0 ? rows.slice(1) : rows
    const nextTable = { header, rows: dataRows }
    setTable(nextTable)
    historyRef.current = [
      {
        header: nextTable.header ? [...nextTable.header] : null,
        rows: nextTable.rows.map((row) => [...row]),
      },
    ]
    historyIndexRef.current = 0
    setSelectedRows(new Set())
    setSelectedCols(new Set())
    setActiveCell(null)
    setFillDrag(null)
    setRowSelectionAnchor(null)
    setLastRowSelection(null)
    setLastColSelection(null)
  }, [content, resolvedDelimiter])

  useEffect(() => {
    setSortCol(null)
    setSortAsc(true)
  }, [columnStorageKey, fileName])

  useEffect(() => {
    if (!dirty) {
      setSaveError(null)
    }
  }, [dirty])

  const maxCols = useMemo(() => {
    if (!table) return 0
    const rowLengths = table.rows.map((row) => row.length)
    const headerLength = table.header ? table.header.length : 0
    return Math.max(headerLength, ...rowLengths, 0)
  }, [table])

  const headerRow = useMemo(() => {
    if (!table) return []
    if (!table.header) return buildHeader(maxCols)
    return Array.from({ length: maxCols }, (_, index) => table.header?.[index] ?? '')
  }, [table, maxCols])

  /* filter + sort ------------------------------------------- */
  useEffect(() => {
    setColumnFilters((prev) => {
      if (prev.length === maxCols) return prev
      return Array.from({ length: maxCols }, (_, index) => prev[index] ?? '')
    })
  }, [maxCols])

  const filteredRows = useMemo(() => {
    if (!table) return [] as RowWithIndex[]

    const normalizedColumnFilters = columnFilters.map((value) =>
      value.trim().toLowerCase(),
    )
    const hasColumnFilters = normalizedColumnFilters.some(Boolean)
    const normalizedSearch = search.trim().toLowerCase()
    const hasGlobalSearch = normalizedSearch.length > 0

    if (!hasColumnFilters && !hasGlobalSearch) {
      return table.rows.map((row, index) => ({ row, index }))
    }

    return table.rows
      .map((row, index) => ({ row, index }))
      .filter((entry) => {
        if (hasGlobalSearch) {
          const matchesSearch = entry.row.some((cell) =>
            cell?.toLowerCase().includes(normalizedSearch),
          )
          if (!matchesSearch) return false
        }

        if (!hasColumnFilters) return true

        return normalizedColumnFilters.every((filterValue, colIndex) => {
          if (!filterValue) return true
          const cellValue = entry.row[colIndex] ?? ''
          return cellValue.toLowerCase().includes(filterValue)
        })
      })
  }, [columnFilters, search, table])

  const sortSnapshot = useMemo(() => {
    if (!table || sortCol === null) {
      return null
    }
    const next = new Map<number, SortSnapshotEntry>()
    table.rows.forEach((row, index) => {
      const raw = row[sortCol] ?? ''
      const text = raw.toLowerCase()
      const numeric = text.trim() === '' || Number.isNaN(Number(text)) ? null : Number(text)
      next.set(index, { text, numeric })
    })
    return next
  }, [sortCol, table])

  const sortedRows = useMemo(() => {
    if (!table || sortCol === null) return filteredRows
    if (!sortSnapshot) return filteredRows
    return [...filteredRows].sort((a, b) => {
      const aSnapshot = sortSnapshot.get(a.index)
      const bSnapshot = sortSnapshot.get(b.index)
      const aVal = aSnapshot?.text ?? (a.row[sortCol] ?? '').toLowerCase()
      const bVal = bSnapshot?.text ?? (b.row[sortCol] ?? '').toLowerCase()
      const aNumeric =
        aSnapshot?.numeric ??
        (aVal.trim() === '' || Number.isNaN(Number(aVal)) ? null : Number(aVal))
      const bNumeric =
        bSnapshot?.numeric ??
        (bVal.trim() === '' || Number.isNaN(Number(bVal)) ? null : Number(bVal))
      if (aNumeric !== null && bNumeric !== null) {
        return (aNumeric - bNumeric) * (sortAsc ? 1 : -1)
      }
      return aVal.localeCompare(bVal) * (sortAsc ? 1 : -1)
    })
  }, [filteredRows, sortCol, sortAsc, sortSnapshot, table])

  const rowOrder = useMemo(
    () => sortedRows.map((entry) => entry.index),
    [sortedRows],
  )

  const columnTotals = useMemo<ColumnTotal[]>(() => {
    if (!table || maxCols === 0) return [] as ColumnTotal[]
    return Array.from({ length: maxCols }, (_, colIndex) => {
      let total = 0
      let numericCount = 0
      let nonEmptyCount = 0
      let symbol: string | null = null

      filteredRows.forEach(({ row }) => {
        const value = row[colIndex]
        if (value === undefined || value === null || value.trim() === '') {
          return
        }
        nonEmptyCount += 1
        const parsed = parseNumericValue(value)
        if (!parsed) return
        numericCount += 1
        total += parsed.value
        if (parsed.currencySymbol && !symbol) {
          symbol = parsed.currencySymbol
        }
      })

      if (numericCount === 0 || nonEmptyCount !== numericCount) return null
      if (symbol) {
        return { type: 'currency', value: total, symbol }
      }
      return { type: 'number', value: total }
    })
  }, [filteredRows, maxCols, table])

  const isLoading = !table

  const cloneTable = useCallback((source: CsvTable): CsvTable => {
    return {
      header: source.header ? [...source.header] : null,
      rows: source.rows.map((row) => [...row]),
    }
  }, [])

  const applyTableUpdate = useCallback(
    (next: CsvTable, options?: { trackHistory?: boolean }) => {
      skipParseRef.current = true
      setTable(next)
      onContentChange(serializeTable(next, resolvedDelimiter))
      if (options?.trackHistory === false) return
      const nextHistory = historyRef.current.slice(
        0,
        historyIndexRef.current + 1,
      )
      nextHistory.push(cloneTable(next))
      historyRef.current = nextHistory
      historyIndexRef.current = nextHistory.length - 1
    },
    [cloneTable, onContentChange, resolvedDelimiter],
  )

  const ensureColumnWidths = useCallback(
    (widths: number[]) => {
      return Array.from({ length: maxCols }, (_, index) =>
        normalizeColumnWidth(widths[index]),
      )
    },
    [maxCols],
  )

  const visibleColumnWidths = useMemo(() => {
    return ensureColumnWidths(columnWidths)
  }, [columnWidths, ensureColumnWidths])

  const fillColumnWidths = useMemo(() => {
    if (hasManualResize) return visibleColumnWidths
    if (maxCols === 0) return visibleColumnWidths
    if (!tableContainerWidth) return visibleColumnWidths
    const fixedWidth = COUNT_COL_WIDTH + ACTION_COL_WIDTH
    const availableWidth = Math.max(0, tableContainerWidth - fixedWidth)
    const baseWidth = visibleColumnWidths.reduce((sum, width) => sum + width, 0)
    if (baseWidth >= availableWidth) return visibleColumnWidths
    const extra = (availableWidth - baseWidth) / maxCols
    return visibleColumnWidths.map((width) => width + extra)
  }, [maxCols, tableContainerWidth, visibleColumnWidths])

  const tableWidth = useMemo(() => {
    const dataWidth = fillColumnWidths.reduce((sum, width) => sum + width, 0)
    const totalWidth = dataWidth + COUNT_COL_WIDTH + ACTION_COL_WIDTH
    return Math.max(totalWidth, tableContainerWidth)
  }, [fillColumnWidths, tableContainerWidth])

  useEffect(() => {
    if (!columnStorageKey) {
      setColumnWidths(ensureColumnWidths([]))
      setHasManualResize(false)
      return
    }
    try {
      const stored = readPersistedValue(columnStorageKey)
      if (!stored) {
        setColumnWidths(ensureColumnWidths([]))
        setHasManualResize(false)
        return
      }
      const parsed = JSON.parse(stored) as { widths?: number[] } | null
      setColumnWidths(ensureColumnWidths(parsed?.widths ?? []))
      setHasManualResize(Boolean(parsed?.widths?.length))
    } catch {
      setColumnWidths(ensureColumnWidths([]))
      setHasManualResize(false)
    }
  }, [columnStorageKey, ensureColumnWidths])

  useEffect(() => {
    if (!columnStorageKey) return
    if (columnWidths.length === 0) return
    writePersistedValue(
      columnStorageKey,
      JSON.stringify({ widths: columnWidths }),
    )
  }, [columnStorageKey, columnWidths])

  useEffect(() => {
    setColumnWidths((prev) => {
      if (prev.length === maxCols) return prev
      return ensureColumnWidths(prev)
    })
  }, [ensureColumnWidths, maxCols])

  useEffect(() => {
    const container = tableContainerRef.current
    if (!container) return
    const updateWidth = () => {
      setTableContainerWidth(container.clientWidth)
    }
    updateWidth()
    if (typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => updateWidth())
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  /* handlers ------------------------------------------------- */
  const clearRowSelections = () => {
    setSelectedRows(new Set())
    setRowSelectionAnchor(null)
    setLastRowSelection(null)
  }

  const clearColumnSelections = () => {
    setSelectedCols(new Set())
    setLastColSelection(null)
  }

  const toggleRowSelection = (rowIndex: number, isRange: boolean) => {
    clearColumnSelections()
    if (!isRange || rowSelectionAnchor === null) {
      setSelectedRows((prev) => {
        const next = new Set(prev)
        if (next.has(rowIndex)) next.delete(rowIndex)
        else next.add(rowIndex)
        return next
      })
      setRowSelectionAnchor(rowIndex)
      setLastRowSelection(rowIndex)
      return
    }

    selectRowRange(rowSelectionAnchor, rowIndex)
    setLastRowSelection(rowIndex)
  }

  const selectRowRange = useCallback(
    (anchorIndex: number, targetIndex: number) => {
      clearColumnSelections()
      if (rowOrder.length === 0) return
      const anchorPos = rowOrder.indexOf(anchorIndex)
      const targetPos = rowOrder.indexOf(targetIndex)
      if (anchorPos === -1 || targetPos === -1) return
      const start = Math.min(anchorPos, targetPos)
      const end = Math.max(anchorPos, targetPos)
      const range = rowOrder.slice(start, end + 1)
      setSelectedRows(new Set(range))
    },
    [clearColumnSelections, rowOrder],
  )

  const toggleColumnSelection = (colIndex: number, isRange: boolean) => {
    clearRowSelections()
    if (!isRange || lastColSelection === null) {
      setSelectedCols((prev) => {
        const next = new Set(prev)
        if (next.has(colIndex)) next.delete(colIndex)
        else next.add(colIndex)
        return next
      })
      setLastColSelection(colIndex)
      return
    }

    const range = buildRange(lastColSelection, colIndex)
    setSelectedCols((prev) => {
      const next = new Set(prev)
      range.forEach((idx) => next.add(idx))
      return next
    })
  }

  const handleHeaderClick = (col: number) => {
    if (sortCol === col) setSortAsc((a) => !a)
    else {
      setSortCol(col)
      setSortAsc(true)
    }
  }

  const updateCell = (rowIndex: number, colIndex: number, value: string) => {
    if (!table) return
    const nextRows = table.rows.map((row, idx) => {
      if (idx !== rowIndex) return row
      const nextRow = [...row]
      nextRow[colIndex] = value
      return nextRow
    })
    applyTableUpdate({ ...table, rows: nextRows })
  }

  const updateHeaderCell = (colIndex: number, value: string) => {
    if (!table || !table.header) return
    const nextHeader = Array.from(
      { length: Math.max(table.header.length, colIndex + 1, maxCols) },
      (_, index) => table.header?.[index] ?? '',
    )
    nextHeader[colIndex] = value
    applyTableUpdate({ ...table, header: nextHeader })
  }

  const addColumn = () => {
    if (!table) return
    const nextRows = table.rows.map((row) => [...row, ''])
    const nextHeader = table.header ? [...table.header, ''] : null
    applyTableUpdate({ header: nextHeader, rows: nextRows })
    setColumnWidths((prev) => {
      const next = [...prev]
      next.push(DEFAULT_COLUMN_WIDTH)
      return next
    })
  }

  const createEmptyRow = useCallback(() => {
    const length = Math.max(maxCols, 1)
    return Array.from({ length }, () => '')
  }, [maxCols])

  const insertRowAt = useCallback(
    (index: number) => {
      if (!table) return
      const safeIndex = Math.min(Math.max(index, 0), table.rows.length)
      const nextRows = [...table.rows]
      nextRows.splice(safeIndex, 0, createEmptyRow())
      applyTableUpdate({ ...table, rows: nextRows })
      setSelectedRows(new Set([safeIndex]))
      setRowSelectionAnchor(safeIndex)
      setLastRowSelection(safeIndex)
      setActiveCell({ rowIndex: safeIndex, colIndex: 0 })
    },
    [applyTableUpdate, createEmptyRow, table],
  )

  const addRowAbove = () => {
    if (!table) return
    if (table.rows.length === 0) {
      insertRowAt(0)
      return
    }
    if (selectedRows.size > 0) {
      const target = Math.min(...selectedRows)
      insertRowAt(target)
      return
    }
    insertRowAt(0)
  }

  const addRowBelow = () => {
    if (!table) return
    if (table.rows.length === 0) {
      insertRowAt(0)
      return
    }
    if (selectedRows.size > 0) {
      const target = Math.max(...selectedRows) + 1
      insertRowAt(target)
      return
    }
    insertRowAt(table.rows.length)
  }

  const addRowBetween = () => {
    if (!table) return
    if (table.rows.length === 0) {
      insertRowAt(0)
      return
    }
    if (selectedRows.size === 1) {
      const [target] = Array.from(selectedRows)
      insertRowAt(Math.min(table.rows.length, target + 1))
      return
    }
    if (selectedRows.size >= 2) {
      const sorted = Array.from(selectedRows).sort((a, b) => a - b)
      let insertIndex = sorted[0] + 1
      for (let i = 0; i < sorted.length - 1; i += 1) {
        if (sorted[i + 1] - sorted[i] > 1) {
          insertIndex = sorted[i] + 1
          break
        }
      }
      insertRowAt(insertIndex)
      return
    }
    const middleIndex = Math.ceil(table.rows.length / 2)
    insertRowAt(middleIndex)
  }

  const deleteSelectedRows = () => {
    if (!table || selectedRows.size === 0) return
    const nextRows = table.rows.filter((_, idx) => !selectedRows.has(idx))
    applyTableUpdate({ ...table, rows: nextRows })
    setSelectedRows(new Set())
    setActiveCell(null)
  }

  const deleteRowAt = (rowIndex: number) => {
    if (!table) return
    if (rowIndex < 0 || rowIndex >= table.rows.length) return
    const nextRows = table.rows.filter((_, idx) => idx !== rowIndex)
    applyTableUpdate({ ...table, rows: nextRows })
    setSelectedRows(new Set())
    setActiveCell(null)
  }

  const deleteSelectedColumns = () => {
    if (!table || selectedCols.size === 0) return
    const colsToRemove = Array.from(selectedCols).sort((a, b) => a - b)
    const removeSet = new Set(colsToRemove)

    const nextRows = table.rows.map((row) =>
      row.filter((_, colIndex) => !removeSet.has(colIndex)),
    )

    const nextHeader = table.header
      ? table.header.filter((_, colIndex) => !removeSet.has(colIndex))
      : null

    applyTableUpdate({ header: nextHeader, rows: nextRows })
    setColumnWidths((prev) => {
      if (prev.length === 0) return prev
      return prev.filter((_, colIndex) => !removeSet.has(colIndex))
    })

    setSortCol((prev) => {
      if (prev === null) return null
      if (removeSet.has(prev)) return null
      const removedBefore = colsToRemove.filter((col) => col < prev).length
      return prev - removedBefore
    })
    setSelectedCols(new Set())
    setActiveCell((prev) => {
      if (!prev) return prev
      if (removeSet.has(prev.colIndex)) return null
      const removedBefore = colsToRemove.filter((col) => col < prev.colIndex).length
      return { ...prev, colIndex: prev.colIndex - removedBefore }
    })
  }

  const clearSelections = () => {
    clearColumnSelections()
    clearRowSelections()
  }

  const reorderColumns = (fromIndex: number, toIndex: number) => {
    if (!table || fromIndex === toIndex) return
    const normalizeRow = (row: string[]) => {
      const nextRow = [...row]
      while (nextRow.length < maxCols) nextRow.push('')
      return nextRow
    }

    const nextRows = table.rows.map((row) => {
      const nextRow = normalizeRow(row)
      const [moved] = nextRow.splice(fromIndex, 1)
      nextRow.splice(toIndex, 0, moved ?? '')
      return nextRow
    })

    const nextHeader = table.header
      ? (() => {
          const next = [...table.header]
          while (next.length < maxCols) next.push('')
          const [moved] = next.splice(fromIndex, 1)
          next.splice(toIndex, 0, moved ?? '')
          return next
        })()
      : null

    applyTableUpdate({ header: nextHeader, rows: nextRows })

    setSelectedCols((prev) => {
      if (prev.size === 0) return prev
      const next = new Set<number>()
      prev.forEach((col) => {
        if (col === fromIndex) {
          next.add(toIndex)
        } else if (fromIndex < toIndex && col > fromIndex && col <= toIndex) {
          next.add(col - 1)
        } else if (fromIndex > toIndex && col >= toIndex && col < fromIndex) {
          next.add(col + 1)
        } else {
          next.add(col)
        }
      })
      return next
    })

    setSortCol((prev) => {
      if (prev === null) return prev
      if (prev === fromIndex) return toIndex
      if (fromIndex < toIndex && prev > fromIndex && prev <= toIndex) return prev - 1
      if (fromIndex > toIndex && prev >= toIndex && prev < fromIndex) return prev + 1
      return prev
    })

    setColumnWidths((prev) => {
      if (prev.length === 0) return prev
      const next = [...prev]
      const [moved] = next.splice(fromIndex, 1)
      next.splice(toIndex, 0, moved ?? DEFAULT_COLUMN_WIDTH)
      return next
    })
  }

  const startFillDrag = (
    event: React.MouseEvent<HTMLButtonElement>,
    rowIndex: number,
    colIndex: number,
  ) => {
    event.preventDefault()
    event.stopPropagation()
    setFillDrag({
      startRowIndex: rowIndex,
      endRowIndex: rowIndex,
      colIndex,
      visibleRowIndices: [...rowOrder],
    })
  }

  const extendFillDragToRow = useCallback((rowIndex: number, colIndex: number) => {
    setFillDrag((prev) => {
      if (!prev || prev.colIndex !== colIndex || prev.endRowIndex === rowIndex) {
        return prev
      }
      return { ...prev, endRowIndex: rowIndex }
    })
  }, [])

  const isRowInFillRange = (rowIndex: number, colIndex: number) => {
    if (!fillDrag || fillDrag.colIndex !== colIndex) return false
    const startPosition = fillDrag.visibleRowIndices.indexOf(fillDrag.startRowIndex)
    const endPosition = fillDrag.visibleRowIndices.indexOf(fillDrag.endRowIndex)
    const rowPosition = fillDrag.visibleRowIndices.indexOf(rowIndex)
    if (startPosition === -1 || endPosition === -1 || rowPosition === -1) {
      return false
    }
    const start = Math.min(startPosition, endPosition)
    const end = Math.max(startPosition, endPosition)
    return rowPosition >= start && rowPosition <= end
  }

  const isFillStartCell = (rowIndex: number, colIndex: number) => {
    if (!fillDrag) return false
    return fillDrag.startRowIndex === rowIndex && fillDrag.colIndex === colIndex
  }

  useEffect(() => {
    if (!fillDrag) return

    const previousUserSelect = document.body.style.userSelect
    document.body.style.userSelect = 'none'

    const handleMove = (event: MouseEvent) => {
      const element = document.elementFromPoint(
        event.clientX,
        event.clientY,
      ) as HTMLElement | null
      const cell = element?.closest('[data-row-index]') as HTMLElement | null
      if (!cell) return
      const rowIndex = Number(cell.dataset.rowIndex)
      if (Number.isNaN(rowIndex)) return
      extendFillDragToRow(rowIndex, fillDrag.colIndex)
    }

    const handleUp = () => {
      setFillDrag((prev) => {
        if (!prev) return null
        if (!table) return null

        const startPosition = prev.visibleRowIndices.indexOf(prev.startRowIndex)
        const endPosition = prev.visibleRowIndices.indexOf(prev.endRowIndex)
        if (startPosition === -1 || endPosition === -1) return null
        const start = Math.min(startPosition, endPosition)
        const end = Math.max(startPosition, endPosition)
        const rowIndices = prev.visibleRowIndices.slice(start, end + 1)
        const rowIndexSet = new Set(rowIndices)

        const fillValue = table.rows[prev.startRowIndex]?.[prev.colIndex] ?? ''
        const nextRows = table.rows.map((row, idx) => {
          if (!rowIndexSet.has(idx)) return row
          const nextRow = [...row]
          nextRow[prev.colIndex] = fillValue
          return nextRow
        })

        applyTableUpdate({ ...table, rows: nextRows })
        return null
      })
    }

    document.addEventListener('mousemove', handleMove)
    document.addEventListener('mouseup', handleUp)

    return () => {
      document.body.style.userSelect = previousUserSelect
      document.removeEventListener('mousemove', handleMove)
      document.removeEventListener('mouseup', handleUp)
    }
  }, [fillDrag, applyTableUpdate, table, extendFillDragToRow])

  useEffect(() => {
    if (!resizeState) return

    const handleMove = (event: MouseEvent) => {
      const delta = event.clientX - resizeState.startX
      const nextWidth = Math.min(
        MAX_COLUMN_WIDTH,
        Math.max(MIN_COLUMN_WIDTH, resizeState.startWidth + delta),
      )
      setColumnWidths((prev) => {
        const next = [...prev]
        next[resizeState.colIndex] = nextWidth
        return next
      })
    }

    const handleUp = () => {
      setResizeState(null)
    }

    document.addEventListener('mousemove', handleMove)
    document.addEventListener('mouseup', handleUp)

    return () => {
      document.removeEventListener('mousemove', handleMove)
      document.removeEventListener('mouseup', handleUp)
    }
  }, [resizeState])

  const handleEditClick = () => {
    if (!onEdit) return
    startTransition(() => setSwitching(true))
    setTimeout(onEdit, 0)
  }

  const handleClose = () => {
    onClose?.()
  }

  const handleSave = async () => {
    if (!onSave) return
    const ok = (await onSave()) !== false
    if (ok) {
      setJustSaved(true)
      setTimeout(() => setJustSaved(false), 2000)
      setSaveError(null)
      return
    }
    setSaveError('Save failed')
  }

  const handleKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const key = event.key.toLowerCase()
    const hasModifier = event.metaKey || event.ctrlKey
    if (hasModifier && !event.altKey) {
      if (key === 'z') {
        event.preventDefault()
        if (event.shiftKey) {
          const redoIndex = historyIndexRef.current + 1
          const redoTable = historyRef.current[redoIndex]
          if (redoTable) {
            historyIndexRef.current = redoIndex
            applyTableUpdate(cloneTable(redoTable), { trackHistory: false })
          }
        } else {
          const undoIndex = historyIndexRef.current - 1
          const undoTable = historyRef.current[undoIndex]
          if (undoTable) {
            historyIndexRef.current = undoIndex
            applyTableUpdate(cloneTable(undoTable), { trackHistory: false })
          }
        }
        return
      }
      if (key === 'y') {
        event.preventDefault()
        const redoIndex = historyIndexRef.current + 1
        const redoTable = historyRef.current[redoIndex]
        if (redoTable) {
          historyIndexRef.current = redoIndex
          applyTableUpdate(cloneTable(redoTable), { trackHistory: false })
        }
        return
      }
    }
    if (!event.shiftKey) return
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return
    if (sortedRows.length === 0 || selectedRows.size === 0) return
    event.preventDefault()

    const currentRowIndex =
      lastRowSelection ?? rowOrder.find((row) => selectedRows.has(row))
    if (currentRowIndex === undefined) return
    const currentPos = rowOrder.indexOf(currentRowIndex)
    if (currentPos === -1) return

    const delta = event.key === 'ArrowUp' ? -1 : 1
    const nextPos = Math.min(rowOrder.length - 1, Math.max(0, currentPos + delta))
    const nextRowIndex = rowOrder[nextPos]

    if (lastRowSelection === null) {
      setLastRowSelection(currentRowIndex)
    }
    const anchor = rowSelectionAnchor ?? currentRowIndex
    if (rowSelectionAnchor === null) {
      setRowSelectionAnchor(anchor)
    }
    selectRowRange(anchor, nextRowIndex)
    setLastRowSelection(nextRowIndex)
    requestAnimationFrame(() => {
      const container = tableContainerRef.current
      if (!container) return
      const cell = container.querySelector(
        `[data-row-index="${nextRowIndex}"]`,
      ) as HTMLElement | null
      if (!cell) return
      const containerRect = container.getBoundingClientRect()
      const cellRect = cell.getBoundingClientRect()
      if (cellRect.bottom > containerRect.bottom) {
        container.scrollTop += cellRect.bottom - containerRect.bottom
      } else if (cellRect.top < containerRect.top) {
        container.scrollTop -= containerRect.top - cellRect.top
      }
    })
  }

  const canSave = dirty && Boolean(onSave)
  const rootStyle = useMemo(
    () =>
      ({
        // Parent components can theme the grid by passing style={{
        //   '--csv-surface-toolbar': '...',
        //   '--csv-text-primary': '...',
        //   '--csv-font-sans': '...',
        // }}. Explicit values passed in style win over these defaults.
        ...CSV_THEME_VARS,
        fontFamily: 'var(--csv-theme-font-sans)',
        color: 'var(--csv-theme-text-primary)',
        ...style,
      }) as React.CSSProperties,
    [style],
  )
  const monoFontStyle = useMemo(
    () =>
      ({
        fontFamily: 'var(--csv-theme-font-mono)',
      }) as React.CSSProperties,
    [],
  )

  const isNearColumnEdge = (
    event: React.MouseEvent<HTMLTableCellElement>,
  ) => {
    const rect = event.currentTarget.getBoundingClientRect()
    const offset = event.clientX - rect.left
    return rect.width - offset <= RESIZE_HIT_WIDTH
  }

  const startColumnResize = (
    event: React.MouseEvent<HTMLElement>,
    colIndex: number,
    width: number,
  ) => {
    event.preventDefault()
    event.stopPropagation()
    if (!hasManualResize) {
      setColumnWidths(ensureColumnWidths(fillColumnWidths))
      setHasManualResize(true)
    }
    setResizeState({
      colIndex,
      startX: event.clientX,
      startWidth: width,
    })
  }

  useEffect(() => {
    if (!rowContextMenu) return
    const handleClose = (event: MouseEvent) => {
      if (rowMenuRef.current?.contains(event.target as Node)) return
      setRowContextMenu(null)
    }
    const handleScroll = () => setRowContextMenu(null)
    window.addEventListener('mousedown', handleClose)
    window.addEventListener('scroll', handleScroll, true)
    return () => {
      window.removeEventListener('mousedown', handleClose)
      window.removeEventListener('scroll', handleScroll, true)
    }
  }, [rowContextMenu])

  const openRowContextMenu = (
    event: React.MouseEvent,
    rowIndex: number | null,
  ) => {
    event.preventDefault()
    event.stopPropagation()
    const container = tableContainerRef.current
    if (!container) return
    if (rowIndex !== null) {
      setSelectedRows(new Set([rowIndex]))
      setRowSelectionAnchor(rowIndex)
      setLastRowSelection(rowIndex)
      setActiveCell({ rowIndex, colIndex: 0 })
    }
    const rect = container.getBoundingClientRect()
    setRowContextMenu({
      x: event.clientX - rect.left + container.scrollLeft,
      y: event.clientY - rect.top + container.scrollTop,
      rowIndex,
    })
  }

  const handleEmptyBodyContextMenu = (event: React.MouseEvent) => {
    if (sortedRows.length > 0) return
    openRowContextMenu(event, null)
  }

  /* render --------------------------------------------------- */
  return (
    <div
      className={`csv-data-grid flex-1 flex flex-col min-h-0 relative ${className ?? ''}`}
      style={rootStyle}
    >
      {/* toolbar */}
      <div className="px-3 py-1 text-xs bg-[var(--csv-theme-surface-toolbar)] border-b border-[var(--csv-theme-border-toolbar)] flex items-center justify-between gap-2">
        <div className="flex items-center gap-3 flex-1">
          <span className="text-[var(--csv-theme-text-muted)] whitespace-nowrap flex items-center gap-2">
            {dirty && (
              <span
                className="inline-flex h-1.5 w-1.5 rounded-full bg-amber-400/60"
                aria-label="Unsaved changes"
                title="Unsaved changes"
              />
            )}
            {formatLabel} · {sortedRows.length} row
            {sortedRows.length === 1 ? '' : 's'}
          </span>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Find…"
            className="rounded px-2 py-0.5 text-xs bg-[var(--csv-theme-surface-muted)] border border-[var(--csv-theme-border-primary)] text-[var(--csv-theme-text-primary)] focus:outline-none"
            style={{ minWidth: 220, maxWidth: 350 }}
          />
          {(selectedRows.size > 0 || selectedCols.size > 0) && (
            <span className="text-[11px] text-[var(--csv-theme-text-muted)] whitespace-nowrap">
              Selected: {selectedRows.size} row
              {selectedRows.size === 1 ? '' : 's'} · {selectedCols.size} col
              {selectedCols.size === 1 ? '' : 's'}
            </span>
          )}
          {saveError && (
            <span className="text-[11px] text-[var(--csv-theme-text-danger)] whitespace-nowrap">
              {saveError}
            </span>
          )}
        </div>
        <div className="flex items-center gap-1">
          <button
            onClick={deleteSelectedRows}
            disabled={selectedRows.size === 0}
            className="px-2 py-0.5 bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs cursor-pointer flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
            type="button"
            title="Delete selected rows"
          >
            <Trash2 className="w-3 h-3" />
            Rows
          </button>
          <button
            onClick={deleteSelectedColumns}
            disabled={selectedCols.size === 0}
            className="px-2 py-0.5 bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs cursor-pointer flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
            type="button"
            title="Delete selected columns"
          >
            <Trash2 className="w-3 h-3" />
            Columns
          </button>
          <button
            onClick={clearSelections}
            disabled={selectedRows.size === 0 && selectedCols.size === 0}
            className="px-2 py-0.5 bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs cursor-pointer flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
            type="button"
            title="Clear selection"
          >
            Clear
          </button>
          {onSave && (
            <button
              onClick={handleSave}
              disabled={!canSave}
              className={`px-2 py-0.5 border rounded text-xs flex items-center gap-1 ${
                canSave
                  ? 'bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border-[var(--csv-theme-border-primary)] text-[var(--csv-theme-text-primary)] cursor-pointer'
                  : 'bg-[var(--csv-theme-surface-muted)] border-[var(--csv-theme-border-primary)] text-[var(--csv-theme-text-primary)] opacity-40 cursor-not-allowed'
              }`}
              type="button"
              title="Save changes"
            >
              <SaveIcon className="w-3 h-3" />
              Save
            </button>
          )}
          {justSaved && (
            <span className="flex items-center gap-1 text-[var(--csv-theme-text-success)] text-xs">
              <Check className="w-3 h-3" /> Saved
            </span>
          )}
          {onEdit && (
            <button
              onClick={handleEditClick}
              className="px-2 py-0.5 bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs cursor-pointer flex items-center gap-1"
              type="button"
            >
              <Pencil className="w-3 h-3" />
              Edit
            </button>
          )}
          {onDownload && (
            <button
              onClick={onDownload}
              className="px-2 py-0.5 bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs flex items-center gap-1"
              type="button"
            >
              <Download className="w-3 h-3" />
              Download
            </button>
          )}
          {onClose && (
            <button
              onClick={handleClose}
              className="px-1 py-0.5 bg-[var(--csv-theme-surface-muted)] cursor-pointer hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] rounded text-[var(--csv-theme-text-primary)] text-xs flex items-center"
              title="Close file"
              aria-label="Close file"
              tabIndex={0}
              type="button"
            >
              <X className="w-3 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* grid */}
      <div
        className="flex-1 min-h-0 overflow-auto bg-[var(--csv-theme-surface-toolbar)] relative"
        ref={tableContainerRef}
        tabIndex={0}
        onKeyDown={handleKeyDown}
        onContextMenu={handleEmptyBodyContextMenu}
      >
        {isLoading ? (
          <div className="absolute inset-0 bg-[var(--csv-theme-surface-overlay)] flex items-center justify-center z-30">
            <span className="text-[var(--csv-theme-text-subtle)] text-sm">Loading…</span>
          </div>
        ) : (
          <div className="min-w-max">
            <table
              className="border border-[var(--csv-theme-border-toolbar)] text-xs text-[var(--csv-theme-text-primary)] border-collapse table-fixed"
              style={{ ...monoFontStyle, width: tableWidth }}
            >
              <colgroup>
                <col style={{ width: COUNT_COL_WIDTH }} />
                {fillColumnWidths.map((width, index) => (
                  <col key={`csv-col-${index}`} style={{ width }} />
                ))}
                <col style={{ width: ACTION_COL_WIDTH }} />
              </colgroup>
              <thead>
                <tr>
                  <th
                    className="sticky top-0 z-20 bg-[var(--csv-theme-surface-toolbar)] border border-[var(--csv-theme-border-primary)] px-2 py-1 text-[11px] text-[var(--csv-theme-text-muted)] text-center"
                    style={{
                      width: COUNT_COL_WIDTH,
                      minWidth: COUNT_COL_WIDTH,
                      maxWidth: COUNT_COL_WIDTH,
                    }}
                  >
                    #
                  </th>
                  {headerRow.map((cell, i) => {
                    const isSelected = selectedCols.has(i)
                    const width = fillColumnWidths[i] ?? DEFAULT_COLUMN_WIDTH
                    return (
                      <th
                        key={i}
                        onDragOver={(event) => {
                          if (draggingCol === null) return
                          event.preventDefault()
                        }}
                        onDrop={(event) => {
                          event.preventDefault()
                          if (draggingCol === null) return
                          reorderColumns(draggingCol, i)
                          setDraggingCol(null)
                        }}
                        onMouseMove={(event) => {
                          if (resizeState) return
                          if (isNearColumnEdge(event)) {
                            setResizeIntent({ colIndex: i, isActive: true })
                          } else if (resizeIntent?.colIndex === i) {
                            setResizeIntent(null)
                          }
                        }}
                        onMouseLeave={() => {
                          if (resizeIntent?.colIndex === i) {
                            setResizeIntent(null)
                          }
                        }}
                        onMouseDown={(event) => {
                          if (!isNearColumnEdge(event)) return
                          startColumnResize(event, i, width)
                        }}
                        style={{ width, minWidth: width, maxWidth: width }}
                        className={`sticky top-0 z-20 border border-[var(--csv-theme-border-primary)] px-2 py-1 font-bold text-[var(--csv-theme-text-primary)] text-xs select-none transition-colors relative ${
                          isSelected
                            ? 'bg-[var(--csv-theme-surface-selected)]'
                            : 'bg-[var(--csv-theme-surface-toolbar)]'
                        } ${
                          resizeIntent?.colIndex === i && resizeIntent.isActive
                            ? 'cursor-col-resize'
                            : ''
                        }`}
                      >
                        <div className="flex items-center gap-1">
                          <input
                            type="checkbox"
                            aria-label="Select column"
                            className="h-3 w-3 m-0 cursor-pointer border border-[var(--csv-theme-border-secondary)] rounded"
                            style={{ accentColor: 'var(--csv-theme-bg-checkbox)' }}
                            checked={isSelected}
                            onClick={(event) => {
                              event.stopPropagation()
                              toggleColumnSelection(i, event.shiftKey)
                            }}
                            title="Select column (Shift for range)"
                            onChange={() => {}}
                          />
                          {table?.header ? (
                            <input
                              value={cell ?? ''}
                              onChange={(event) =>
                                updateHeaderCell(i, event.target.value)
                              }
                              placeholder="Add header…"
                              className="bg-transparent w-full text-[var(--csv-theme-text-primary)] outline-none"
                              onFocus={() =>
                                setActiveCell({ rowIndex: -1, colIndex: i })
                              }
                            />
                          ) : (
                            <span className="text-[var(--csv-theme-text-subtle)]">{cell}</span>
                          )}
                          {isSelected && (
                            <button
                              type="button"
                              draggable
                              aria-label="Drag to reorder column"
                              title="Drag to reorder column"
                              onClick={(event) => event.stopPropagation()}
                              onDragStart={(event) => {
                                event.stopPropagation()
                                setDraggingCol(i)
                              }}
                              onDragEnd={() => setDraggingCol(null)}
                              className="text-[var(--csv-theme-text-muted)] hover:text-[var(--csv-theme-text-subtle)] cursor-grab active:cursor-grabbing"
                            >
                              <GripVertical className="w-3 h-3" />
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation()
                              handleHeaderClick(i)
                            }}
                            aria-label="Sort column"
                            className="text-[var(--csv-theme-text-muted)] hover:text-[var(--csv-theme-text-subtle)]"
                            title="Sort column"
                          >
                            {sortCol === i ? (
                              sortAsc ? (
                                <ChevronUp
                                  size={13}
                                  className="inline -mt-0.5 opacity-60"
                                />
                              ) : (
                                <ChevronDown
                                  size={13}
                                  className="inline -mt-0.5 opacity-60"
                                />
                              )
                            ) : (
                              <ChevronUp
                                size={13}
                                className="inline -mt-0.5 opacity-25"
                              />
                            )}
                          </button>
                        </div>
                        <div
                          className={`absolute right-0 top-0 h-full w-1.5 ${
                            resizeIntent?.colIndex === i
                              ? 'bg-[var(--csv-theme-surface-accent)] opacity-60'
                              : 'bg-transparent'
                          }`}
                          onMouseDown={(event) => {
                            event.stopPropagation()
                            startColumnResize(event, i, width)
                          }}
                          role="presentation"
                          style={{ cursor: 'col-resize' }}
                        />
                      </th>
                    )
                  })}
                  <th
                    className="sticky top-0 z-20 bg-[var(--csv-theme-surface-toolbar)] border border-[var(--csv-theme-border-primary)] px-2 py-1 text-center"
                    style={{
                      width: ACTION_COL_WIDTH,
                      minWidth: ACTION_COL_WIDTH,
                      maxWidth: ACTION_COL_WIDTH,
                    }}
                  >
                    <button
                      type="button"
                      onClick={addColumn}
                      className="inline-flex items-center justify-center w-6 h-6 rounded bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)] border border-[var(--csv-theme-border-primary)] text-[var(--csv-theme-text-primary)]"
                      aria-label="Add column"
                      title="Add column"
                    >
                      <Plus className="w-3 h-3" />
                    </button>
                  </th>
                </tr>
                <tr>
                  <th
                    className="sticky top-[30px] z-20 bg-[var(--csv-theme-surface-toolbar)] border border-[var(--csv-theme-border-primary)] px-2 py-1 text-[11px] text-[var(--csv-theme-text-muted)] text-center"
                    style={{
                      width: COUNT_COL_WIDTH,
                      minWidth: COUNT_COL_WIDTH,
                      maxWidth: COUNT_COL_WIDTH,
                    }}
                  >
                    Filter
                  </th>
                  {Array.from({ length: maxCols }).map((_, i) => {
                    const width = fillColumnWidths[i] ?? DEFAULT_COLUMN_WIDTH
                    return (
                      <th
                        key={`filter-${i}`}
                        className="sticky top-[30px] z-20 bg-[var(--csv-theme-surface-toolbar)] border border-[var(--csv-theme-border-primary)] px-1 py-1"
                        style={{ width, minWidth: width, maxWidth: width }}
                      >
                        <input
                          value={columnFilters[i] ?? ''}
                          onChange={(event) => {
                            const nextValue = event.target.value
                            setColumnFilters((prev) => {
                              const next = [...prev]
                              next[i] = nextValue
                              return next
                            })
                          }}
                          placeholder="Contains…"
                          className="w-full rounded px-1.5 py-0.5 text-[11px] bg-[var(--csv-theme-surface-muted)] border border-[var(--csv-theme-border-primary)] text-[var(--csv-theme-text-primary)] focus:outline-none"
                        />
                      </th>
                    )
                  })}
                  <th
                    className="sticky top-[30px] z-20 bg-[var(--csv-theme-surface-toolbar)] border border-[var(--csv-theme-border-primary)] px-1 py-1 text-center"
                    style={{
                      width: ACTION_COL_WIDTH,
                      minWidth: ACTION_COL_WIDTH,
                      maxWidth: ACTION_COL_WIDTH,
                    }}
                  >
                    <button
                      type="button"
                      onClick={() =>
                        setColumnFilters(Array.from({ length: maxCols }, () => ''))
                      }
                      className="px-1 py-0.5 text-[10px] rounded border border-[var(--csv-theme-border-primary)] bg-[var(--csv-theme-surface-muted)] hover:bg-[var(--csv-theme-surface-muted-hover)]"
                      title="Clear column filters"
                    >
                      Clear
                    </button>
                  </th>
                </tr>
              </thead>
              <tbody>
                {sortedRows.map(({ row, index: rowIndex }, i) => {
                  const isRowSelected = selectedRows.has(rowIndex)
                  return (
                    <tr
                      key={rowIndex}
                      data-row-index={rowIndex}
                      className={
                        isRowSelected
                          ? 'bg-[var(--csv-theme-surface-selected)]'
                          : i % 2 === 0
                            ? 'bg-[var(--csv-theme-surface-row-even)]'
                            : 'bg-[var(--csv-theme-surface-row-odd)]'
                      }
                      onContextMenu={(event) =>
                        openRowContextMenu(event, rowIndex)
                      }
                    >
                      <td
                        style={{
                          width: COUNT_COL_WIDTH,
                          minWidth: COUNT_COL_WIDTH,
                          maxWidth: COUNT_COL_WIDTH,
                        }}
                        className={`border border-[var(--csv-theme-border-toolbar)] px-2 py-1 text-[11px] text-[var(--csv-theme-text-muted)] select-none bg-[var(--csv-theme-surface-toolbar)] ${
                          isRowSelected ? 'ring-1 ring-[var(--csv-theme-border-accent)]' : ''
                        }`}
                      >
                        <label className="flex items-center justify-center gap-1 cursor-pointer">
                          <input
                            type="checkbox"
                            className="h-3 w-3 m-0 cursor-pointer border border-[var(--csv-theme-border-secondary)] rounded"
                            style={{ accentColor: 'var(--csv-theme-bg-checkbox)' }}
                            checked={isRowSelected}
                            onClick={(event) => {
                              event.stopPropagation()
                              toggleRowSelection(rowIndex, event.shiftKey)
                              tableContainerRef.current?.focus()
                            }}
                            onChange={() => {}}
                          />
                          <span>{i + 1}</span>
                        </label>
                      </td>
                      {Array.from({ length: maxCols }).map((_, j) => {
                        const isColSelected = selectedCols.has(j)
                        const isCellSelected = isColSelected || isRowSelected
                        const isActive =
                          activeCell?.rowIndex === rowIndex &&
                          activeCell?.colIndex === j
                        const isFillRange = isRowInFillRange(rowIndex, j)
                        const isFillOrigin = isFillStartCell(rowIndex, j)
                        const width = fillColumnWidths[j] ?? DEFAULT_COLUMN_WIDTH
                        const fillRangeStyle = isFillRange
                          ? {
                              outline: '1px dotted var(--csv-theme-border-accent)',
                              outlineOffset: '-2px',
                            }
                          : undefined
                        return (
                          <td
                            key={`${rowIndex}-${j}`}
                            data-row-index={rowIndex}
                            onMouseEnter={() => extendFillDragToRow(rowIndex, j)}
                            className={`border border-[var(--csv-theme-border-toolbar)] px-2 py-1 whitespace-pre transition-colors duration-75 relative ${
                              isCellSelected
                                ? 'bg-[var(--csv-theme-surface-selected)]'
                                : ''
                            } ${
                              isFillRange
                                ? 'bg-[var(--csv-theme-surface-accent-soft)] ring-1 ring-[var(--csv-theme-border-accent)]'
                                : ''
                            } ${isFillOrigin ? 'ring-2 ring-[var(--csv-theme-border-accent)]' : ''}`}
                            style={{
                              width,
                              minWidth: width,
                              maxWidth: width,
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              ...fillRangeStyle,
                            }}
                          >
                            <input
                              value={row[j] ?? ''}
                              onChange={(event) =>
                                updateCell(rowIndex, j, event.target.value)
                              }
                              onFocus={() =>
                                setActiveCell({ rowIndex, colIndex: j })
                              }
                              className={`bg-transparent w-full outline-none text-[var(--csv-theme-text-primary)] ${
                                isActive ? 'ring-1 ring-[var(--csv-theme-border-accent)]' : ''
                              } ${isFillRange ? 'font-semibold' : ''}`}
                            />
                            {isActive && (
                              <button
                                type="button"
                                aria-label="Drag to fill"
                                title="Drag to fill"
                                onMouseDown={(event) =>
                                  startFillDrag(event, rowIndex, j)
                                }
                                className="absolute bottom-0 right-0 w-2 h-2 bg-[var(--csv-theme-surface-accent)] border border-[var(--csv-theme-border-primary)] rounded-sm cursor-crosshair hover:cursor-crosshair"
                              />
                            )}
                          </td>
                        )
                      })}
                    </tr>
                    )
                  })}
              </tbody>
            </table>

            {sortedRows.length === 0 && (
              <div className="text-xs text-[var(--csv-theme-text-muted)] p-4">
                No matching rows.
              </div>
            )}
          </div>
        )}
        <div className="sticky bottom-0 z-30 bg-[var(--csv-theme-surface-toolbar)] border-t border-[var(--csv-theme-border-primary)]">
          <div className="min-w-max">
            <table
              className="border border-[var(--csv-theme-border-toolbar)] text-xs text-[var(--csv-theme-text-primary)] border-collapse table-fixed"
              style={{ ...monoFontStyle, width: tableWidth }}
            >
              <colgroup>
                <col style={{ width: COUNT_COL_WIDTH }} />
                {fillColumnWidths.map((width, index) => (
                  <col key={`csv-total-col-${index}`} style={{ width }} />
                ))}
                <col style={{ width: ACTION_COL_WIDTH }} />
              </colgroup>
              <tbody>
                <tr className="bg-[var(--csv-theme-surface-toolbar)]">
                  <td
                    className="border border-[var(--csv-theme-border-toolbar)] px-2 py-1 text-[11px] text-[var(--csv-theme-text-muted)] text-center bg-[var(--csv-theme-surface-toolbar)]"
                    style={{
                      width: COUNT_COL_WIDTH,
                      minWidth: COUNT_COL_WIDTH,
                      maxWidth: COUNT_COL_WIDTH,
                    }}
                  >
                    Totals
                  </td>
                  {Array.from({ length: maxCols }).map((_, colIndex) => {
                    const total = columnTotals[colIndex] ?? null
                    const width = fillColumnWidths[colIndex] ?? DEFAULT_COLUMN_WIDTH
                    return (
                      <td
                        key={`totals-${colIndex}`}
                        className="border border-[var(--csv-theme-border-toolbar)] px-2 py-1 text-[var(--csv-theme-text-primary)] font-semibold bg-[var(--csv-theme-surface-toolbar)]"
                        style={{
                          width,
                          minWidth: width,
                          maxWidth: width,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {total ? formatNumericTotal(total) : ''}
                      </td>
                    )
                  })}
                  <td
                    className="border border-[var(--csv-theme-border-toolbar)] px-2 py-1 bg-[var(--csv-theme-surface-toolbar)]"
                    style={{
                      width: ACTION_COL_WIDTH,
                      minWidth: ACTION_COL_WIDTH,
                      maxWidth: ACTION_COL_WIDTH,
                    }}
                  />
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        {rowContextMenu && (
          <div
            className="absolute z-40 rounded border border-[var(--csv-theme-border-primary)] bg-[var(--csv-theme-surface-muted)] text-xs text-[var(--csv-theme-text-primary)] shadow-lg min-w-[160px]"
            style={{ left: rowContextMenu.x, top: rowContextMenu.y }}
            role="menu"
            ref={rowMenuRef}
          >
            {rowContextMenu.rowIndex === null ? (
              <button
                type="button"
                className="w-full px-3 py-2 text-left hover:bg-[var(--csv-theme-surface-muted-hover)] flex items-center gap-2"
                onClick={() => {
                  insertRowAt(0)
                  setRowContextMenu(null)
                }}
              >
                <Plus className="w-3 h-3" /> + Row
              </button>
            ) : (
              <>
                <button
                  type="button"
                  className="w-full px-3 py-2 text-left hover:bg-[var(--csv-theme-surface-muted-hover)] flex items-center gap-2"
                  onClick={() => {
                    addRowAbove()
                    setRowContextMenu(null)
                  }}
                >
                  <Plus className="w-3 h-3" /> + Row above
                </button>
                <button
                  type="button"
                  className="w-full px-3 py-2 text-left hover:bg-[var(--csv-theme-surface-muted-hover)] flex items-center gap-2"
                  onClick={() => {
                    addRowBelow()
                    setRowContextMenu(null)
                  }}
                >
                  <Plus className="w-3 h-3" /> + Row below
                </button>
                <div className="border-t border-[var(--csv-theme-border-primary)] opacity-40" />
                <button
                  type="button"
                  className="w-full px-3 py-2 text-left hover:bg-[var(--csv-theme-surface-muted-hover)] flex items-center gap-2 text-[var(--csv-theme-text-danger)]"
                  onClick={() => {
                    if (rowContextMenu.rowIndex !== null) {
                      deleteRowAt(rowContextMenu.rowIndex)
                    }
                    setRowContextMenu(null)
                  }}
                >
                  <Trash2 className="w-3 h-3" /> Delete row
                </button>
              </>
            )}
          </div>
        )}
      </div>

      {/* overlay while switching */}
      {switching && (
        <div className="absolute inset-0 bg-[var(--csv-theme-surface-overlay)] flex flex-col items-center justify-center z-50">
          <svg
            className="animate-spin h-6 w-6 text-[var(--csv-theme-text-accent)] mb-3"
            viewBox="0 0 24 24"
          >
            <circle
              className="opacity-25"
              cx="12"
              cy="12"
              r="10"
              stroke="currentColor"
              strokeWidth="4"
              fill="none"
            />
            <path
              className="opacity-75"
              fill="currentColor"
              d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z"
            />
          </svg>
          <span className="text-[var(--csv-theme-text-subtle)] text-sm">
            Opening editor…
          </span>
        </div>
      )}
    </div>
  )
}

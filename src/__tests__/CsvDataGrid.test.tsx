// src/__tests__/CsvDataGrid.test.tsx

import '@testing-library/jest-dom'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { beforeAll, beforeEach, expect, test, vi } from 'vitest'
import CsvDataGrid from '../CsvDataGrid'

class ResizeObserverMock {
  observe() {}
  disconnect() {}
  unobserve() {}
}

beforeAll(() => {
  ;(globalThis as typeof globalThis & { ResizeObserver: typeof ResizeObserverMock }).ResizeObserver =
    ResizeObserverMock
})

beforeEach(() => {
  window.localStorage.clear()
})

test('updates serialized content when a cell changes', () => {
  const onContentChange = vi.fn()

  render(
    <CsvDataGrid
      content={'name,age\nAlice,30\nBob,25'}
      fileName="people.csv"
      onContentChange={onContentChange}
    />,
  )

  fireEvent.change(screen.getByDisplayValue('30'), {
    target: { value: '31' },
  })

  expect(onContentChange).toHaveBeenLastCalledWith('name,age\nAlice,31\nBob,25')
})

test('supports toolbar callbacks without app-specific context', async () => {
  const onContentChange = vi.fn()
  const onSave = vi.fn().mockResolvedValue(true)
  const onEdit = vi.fn()
  const onDownload = vi.fn()
  const onClose = vi.fn()

  render(
    <CsvDataGrid
      content={'name,amount\nAlice,10\nBob,25'}
      fileName="report.csv"
      dirty
      onContentChange={onContentChange}
      onSave={onSave}
      onEdit={onEdit}
      onDownload={onDownload}
      onClose={onClose}
    />,
  )

  fireEvent.click(screen.getByRole('button', { name: /^save$/i }))
  await waitFor(() => expect(onSave).toHaveBeenCalledTimes(1))
  await waitFor(() =>
    expect(
      screen.getByText(
        (_, node) => node?.textContent?.replace(/\s+/g, ' ').trim() === 'Saved',
      ),
    ).toBeInTheDocument(),
  )

  fireEvent.click(screen.getByRole('button', { name: /^edit$/i }))
  await waitFor(() => expect(onEdit).toHaveBeenCalledTimes(1))

  fireEvent.click(screen.getByRole('button', { name: /^download$/i }))
  expect(onDownload).toHaveBeenCalledTimes(1)

  fireEvent.click(screen.getByRole('button', { name: /close file/i }))
  expect(onClose).toHaveBeenCalledTimes(1)
})

test('fills visible rows when dragging the active cell handle', async () => {
  const onContentChange = vi.fn()

  render(
    <CsvDataGrid
      content={'value\n3\n9\n10'}
      fileName="values.csv"
      onContentChange={onContentChange}
    />,
  )

  fireEvent.focus(screen.getByDisplayValue('3'))

  fireEvent.mouseDown(screen.getByRole('button', { name: /drag to fill/i }))

  const targetCell = screen.getByDisplayValue('10').closest('td')
  expect(targetCell).not.toBeNull()
  fireEvent.mouseEnter(targetCell as HTMLTableCellElement)
  fireEvent.mouseUp(document)

  await waitFor(() =>
    expect(onContentChange).toHaveBeenLastCalledWith('value\n3\n3\n3'),
  )
})

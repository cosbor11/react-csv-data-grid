# react-csv-data-grid

Reusable React CSV/TSV grid component with inline editing, filtering, sorting, column resizing, row and column selection, fill-drag, and host-controlled theming.

Package name: `@cosbor11/react-csv-data-grid`

## Install

```bash
npm install @cosbor11/react-csv-data-grid lucide-react
```

Peer dependencies:

- `react`
- `react-dom`

## Tailwind Requirement

This component uses Tailwind utility classes for its layout and interaction styling.

Your app should:

1. Have Tailwind CSS configured.
2. Include this package in your Tailwind content scan so the utility classes are emitted.

Example:

```js
// tailwind.config.js
export default {
  content: [
    './src/**/*.{js,ts,jsx,tsx}',
    './node_modules/@cosbor11/react-csv-data-grid/dist/**/*.{js,mjs,cjs}',
  ],
}
```

## Usage

```tsx
import { useState } from 'react'
import { CsvDataGrid } from '@cosbor11/react-csv-data-grid'

export function Example() {
  const [content, setContent] = useState('name,amount\nAlice,10\nBob,25')

  return (
    <CsvDataGrid
      content={content}
      fileName="report.csv"
      onContentChange={setContent}
      style={
        {
          '--csv-surface-toolbar': '#181818',
          '--csv-surface-muted': '#232323',
          '--csv-surface-selected': '#313131',
          '--csv-border-primary': '#333333',
          '--csv-text-primary': '#d4d4d4',
          '--csv-text-muted': '#9ca3af',
        } as React.CSSProperties
      }
    />
  )
}
```

## Theming

Override the public `--csv-*` tokens from the parent app:

- `--csv-font-sans`
- `--csv-font-mono`
- `--csv-surface-toolbar`
- `--csv-surface-muted`
- `--csv-surface-muted-hover`
- `--csv-surface-row-even`
- `--csv-surface-row-odd`
- `--csv-surface-selected`
- `--csv-surface-overlay`
- `--csv-surface-accent`
- `--csv-surface-accent-soft`
- `--csv-border-primary`
- `--csv-border-toolbar`
- `--csv-border-secondary`
- `--csv-border-accent`
- `--csv-text-primary`
- `--csv-text-muted`
- `--csv-text-subtle`
- `--csv-text-danger`
- `--csv-text-success`
- `--csv-text-accent`
- `--csv-bg-checkbox`

The component resolves these host-facing tokens into internal theme variables, so consumers should style `--csv-*` and not rely on `--csv-theme-*`.

## Features

- CSV and TSV delimiter support
- Inline cell editing
- Per-column filtering
- Sorting
- Column resizing
- Row and column selection
- Keyboard range selection
- Fill-drag from the active cell handle
- Optional save, edit, download, and close callbacks
- Optional persisted column widths

## Development

```bash
npm install
npm run build
npm test
```

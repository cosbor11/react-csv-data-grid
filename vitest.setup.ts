import '@testing-library/jest-dom/vitest'

const createStorageMock = () => {
  const store = new Map<string, string>()

  return {
    getItem(key: string) {
      return store.has(key) ? store.get(key)! : null
    },
    setItem(key: string, value: string) {
      store.set(key, String(value))
    },
    removeItem(key: string) {
      store.delete(key)
    },
    clear() {
      store.clear()
    },
    key(index: number) {
      return Array.from(store.keys())[index] ?? null
    },
    get length() {
      return store.size
    },
  }
}

const storage = createStorageMock()

if (typeof globalThis.localStorage?.getItem !== 'function') {
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: storage,
  })
}

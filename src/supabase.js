import { createClient } from '@supabase/supabase-js'

const clean = s => (s || '').replace(/^﻿/, '').trim()

function fetchWithTimeout(url, options = {}) {
  const controller = new AbortController()
  const id = setTimeout(() => controller.abort(), 15000)
  return fetch(url, { ...options, signal: controller.signal }).finally(() => clearTimeout(id))
}

export const supabase = createClient(
  clean(import.meta.env.VITE_SUPABASE_URL),
  clean(import.meta.env.VITE_SUPABASE_ANON_KEY),
  { global: { fetch: fetchWithTimeout } }
)

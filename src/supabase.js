import { createClient } from '@supabase/supabase-js'

const clean = s => (s || '').replace(/^﻿/, '').trim()

export const supabase = createClient(
  clean(import.meta.env.VITE_SUPABASE_URL),
  clean(import.meta.env.VITE_SUPABASE_ANON_KEY)
)

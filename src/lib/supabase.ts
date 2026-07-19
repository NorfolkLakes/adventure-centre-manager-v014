import { createClient } from '@supabase/supabase-js'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabasePublishableKey = import.meta.env.VITE_SUPABASE_ANON_KEY

if (!supabaseUrl || !supabasePublishableKey) {
  throw new Error(
    'Supabase is not configured. Add VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY to .env.local.',
  )
}

export const supabase = createClient(
  supabaseUrl,
  supabasePublishableKey,
)

import { createClient } from "@supabase/supabase-js";

// Settings -> API in your Supabase project dashboard.
// Using Vite-style env vars here; swap for process.env.NEXT_PUBLIC_... if you're on Next.js.
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

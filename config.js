// من Supabase: Project Settings > API
const SUPABASE_URL = 'https://uipueqnhwozzluqizjnl.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_wTlvamv1Zo6dl2NOdH3XXA_rrNsmRUp';   // الـ anon key آمن للنشر لأن الحماية في RLS
const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

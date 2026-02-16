import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://gufxlygktbhgmbcgdgxf.supabase.co";
const supabaseAnonKey = "sb_publishable_JkmN3v_yXgcoytXwd0xRTw_5E38-9ii";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type, x-app-token',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
};

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
  });
}

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

function getAdminClient() {
  return createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  );
}

async function checkAuth(req: Request): Promise<boolean> {
  const APP_PASSWORD = Deno.env.get('APP_PASSWORD') || '';
  if (!APP_PASSWORD) return true;

  const token = req.headers.get('x-app-token');
  if (!token) return false;

  const supabase = getAdminClient();
  const { data, error } = await supabase
    .from('sessions')
    .select('token, expires_at')
    .eq('token', token)
    .maybeSingle();

  if (error || !data) return false;
  if (new Date(data.expires_at).getTime() < Date.now()) return false;
  return true;
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });

  if (!(await checkAuth(req))) {
    return jsonResponse({ error: 'لازم تسجل دخول الأول' }, 401);
  }

  try {
    const MAX_ROWS = parseInt(Deno.env.get('MAX_ROWS') || '1000', 10);

    const body = await req.json().catch(() => ({}));
    const headers: unknown = body?.headers;
    const rows: unknown = body?.rows;
    const sheetName: string = typeof body?.sheetName === 'string' ? body.sheetName : 'Sheet1';

    if (!Array.isArray(headers) || headers.length === 0) {
      return jsonResponse({ error: 'لم يتم إرسال أعمدة صحيحة' }, 400);
    }
    if (!Array.isArray(rows)) {
      return jsonResponse({ error: 'لم يتم إرسال صفوف صحيحة' }, 400);
    }
    if (rows.length === 0) {
      return jsonResponse({ error: 'الملف فارغ' }, 400);
    }
    if (rows.length > MAX_ROWS) {
      return jsonResponse(
        { error: `عدد الصفوف (${rows.length}) أكبر من الحد المسموح (${MAX_ROWS} صف). قسّم الملف لأجزاء أصغر.` },
        400,
      );
    }

    const supabase = getAdminClient();
    const { data, error } = await supabase
      .from('jobs')
      .insert({
        status: 'uploaded',
        headers,
        rows,
        total: rows.length,
        processed: 0,
        sheet_name: sheetName,
      })
      .select('id')
      .single();

    if (error) throw error;

    return jsonResponse({
      jobId: data.id,
      headers,
      rowCount: rows.length,
      preview: rows.slice(0, 5),
    });
  } catch (e) {
    return jsonResponse({ error: 'فشل إنشاء الوظيفة: ' + String(e?.message || e) }, 500);
  }
});

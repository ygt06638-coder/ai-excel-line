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
import * as XLSX from 'https://esm.sh/xlsx@0.18.5';

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
    const url = new URL(req.url);
    const jobId = url.searchParams.get('jobId');
    if (!jobId) return jsonResponse({ error: 'jobId مفقود' }, 400);

    const supabase = getAdminClient();
    const { data: job, error } = await supabase
      .from('jobs')
      .select('status, rows, sheet_name')
      .eq('id', jobId)
      .maybeSingle();

    if (error) throw error;
    if (!job || job.status !== 'done') {
      return jsonResponse({ error: 'الملف لسه مش جاهز أو انتهت صلاحيته' }, 400);
    }

    const sheet = XLSX.utils.json_to_sheet(job.rows as Record<string, unknown>[]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, job.sheet_name || 'Sheet1');

    const buffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });

    return new Response(buffer, {
      status: 200,
      headers: {
        ...corsHeaders,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="result.xlsx"',
      },
    });
  } catch (e) {
    return jsonResponse({ error: 'فشل تجهيز الملف: ' + String(e?.message || e) }, 500);
  }
});

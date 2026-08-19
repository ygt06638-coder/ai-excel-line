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

const MAX_INSTRUCTION_LENGTH = 1000;

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });

  if (!(await checkAuth(req))) {
    return jsonResponse({ error: 'لازم تسجل دخول الأول' }, 401);
  }

  try {
    const body = await req.json().catch(() => ({}));
    const { jobId, sourceColumn, targetColumn, instruction } = body || {};

    if (!jobId || typeof jobId !== 'string') {
      return jsonResponse({ error: 'jobId مفقود أو غير صحيح' }, 400);
    }
    if (
      typeof sourceColumn !== 'string' ||
      typeof targetColumn !== 'string' ||
      typeof instruction !== 'string'
    ) {
      return jsonResponse({ error: 'لازم تحدد العمود المصدر، العمود الهدف، والتعليمات' }, 400);
    }

    const trimmedInstruction = instruction.trim();
    if (!sourceColumn.trim() || !targetColumn.trim() || !trimmedInstruction) {
      return jsonResponse({ error: 'لازم تحدد العمود المصدر، العمود الهدف، والتعليمات' }, 400);
    }
    if (trimmedInstruction.length > MAX_INSTRUCTION_LENGTH) {
      return jsonResponse(
        { error: `التعليمات طويلة جداً (أقصى حد ${MAX_INSTRUCTION_LENGTH} حرف)` },
        400,
      );
    }

    const GEMINI_API_KEY = Deno.env.get('GEMINI_API_KEY');
    if (!GEMINI_API_KEY) {
      return jsonResponse({ error: 'مفتاح Gemini API غير موجود (سيرفر)' }, 500);
    }

    const supabase = getAdminClient();
    const { data: job, error: fetchErr } = await supabase
      .from('jobs')
      .select('id, status, headers')
      .eq('id', jobId)
      .maybeSingle();

    if (fetchErr) throw fetchErr;
    if (!job) {
      return jsonResponse({ error: 'الوظيفة غير موجودة أو انتهت صلاحيتها، ارفع الملف تاني' }, 404);
    }
    if (job.status === 'processing') {
      return jsonResponse({ error: 'الوظيفة دي شغالة بالفعل، استنى لحد ما تخلص' }, 409);
    }

    const headers: string[] = job.headers || [];
    if (!headers.includes(sourceColumn)) {
      return jsonResponse({ error: `العمود المصدر "${sourceColumn}" مش موجود في الملف` }, 400);
    }
    if (!headers.includes(targetColumn)) {
      return jsonResponse({ error: `العمود الهدف "${targetColumn}" مش موجود في الملف` }, 400);
    }

    const { error: updateErr } = await supabase
      .from('jobs')
      .update({
        source_column: sourceColumn,
        target_column: targetColumn,
        instruction: trimmedInstruction,
        status: 'processing',
        processed: 0,
        error: null,
        last_processed_at: null,
      })
      .eq('id', jobId);

    if (updateErr) throw updateErr;

    return jsonResponse({ started: true });
  } catch (e) {
    return jsonResponse({ error: 'فشل بدء المعالجة: ' + String(e?.message || e) }, 500);
  }
});

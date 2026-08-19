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

const GEMINI_MODEL = Deno.env.get('GEMINI_MODEL') || 'gemini-2.5-flash-lite';
const MIN_DELAY_MS = parseInt(Deno.env.get('MIN_DELAY_MS') || '4200', 10);

async function callGemini(prompt: string): Promise<string> {
  const GEMINI_API_KEY = Deno.env.get('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Gemini API error (${res.status}): ${errText}`);
  }

  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('لم يرجع الـ AI أي نص');
  return text.trim();
}

async function callGeminiWithRetry(prompt: string, maxRetries = 2): Promise<string> {
  let delay = 2000;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await callGemini(prompt);
    } catch (err) {
      const isRateLimit = String(err?.message || err).includes('429');
      if (attempt === maxRetries || !isRateLimit) throw err;
      await new Promise((r) => setTimeout(r, delay));
      delay *= 2;
    }
  }
  throw new Error('فشلت كل المحاولات');
}

function statusResponse(job: any) {
  const etaSeconds =
    job.status === 'processing' ? Math.max(0, job.total - job.processed) * (MIN_DELAY_MS / 1000) : 0;
  return jsonResponse({
    status: job.status,
    total: job.total,
    processed: job.processed,
    error: job.error,
    etaSeconds,
    pollAfterMs: MIN_DELAY_MS,
  });
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });

  if (!(await checkAuth(req))) {
    return jsonResponse({ error: 'لازم تسجل دخول الأول' }, 401);
  }

  try {
    const body = await req.json().catch(() => ({}));
    const { jobId } = body || {};
    if (!jobId || typeof jobId !== 'string') {
      return jsonResponse({ error: 'jobId مفقود أو غير صحيح' }, 400);
    }

    const supabase = getAdminClient();
    const { data: job, error: fetchErr } = await supabase
      .from('jobs')
      .select('*')
      .eq('id', jobId)
      .maybeSingle();

    if (fetchErr) throw fetchErr;
    if (!job) return jsonResponse({ error: 'الوظيفة غير موجودة' }, 404);

    // خلصت أو فشلت بالفعل - رجّع الحالة الحالية من غير أي شغل
    if (job.status === 'done' || job.status === 'error') {
      return statusResponse(job);
    }

    if (job.status !== 'processing') {
      return jsonResponse({ error: 'لازم تبدأ المعالجة الأول' }, 400);
    }

    // شبكة أمان: احترام المهلة بين كل طلب وطلب حتى لو الكلاينت داس بسرعة أو فتح أكتر من تاب
    if (job.last_processed_at) {
      const elapsed = Date.now() - new Date(job.last_processed_at).getTime();
      if (elapsed < MIN_DELAY_MS) {
        return statusResponse(job);
      }
    }

    const rows = job.rows as Record<string, unknown>[];
    const i = job.processed;

    if (i >= job.total) {
      const { error: doneErr } = await supabase
        .from('jobs')
        .update({ status: 'done' })
        .eq('id', jobId);
      if (doneErr) throw doneErr;
      job.status = 'done';
      return statusResponse(job);
    }

    const row = rows[i];
    const sourceValue = row[job.source_column as string] ?? '';

    const prompt = `أنت خبير سيو (SEO) ومحتوى تسويقي.
التعليمات: ${job.instruction}

القيمة الأصلية للمنتج (العمود: ${job.source_column}): "${sourceValue}"

اكتب فقط النتيجة النهائية المطلوبة بدون أي شرح أو مقدمات أو علامات اقتباس.`;

    try {
      const result = await callGeminiWithRetry(prompt);
      row[job.target_column as string] = result;
    } catch (err) {
      row[job.target_column as string] = '⚠️ فشل المعالجة: ' + String(err?.message || err);
    }

    const processed = i + 1;
    const newStatus = processed >= job.total ? 'done' : 'processing';

    const { error: updateErr } = await supabase
      .from('jobs')
      .update({
        rows,
        processed,
        status: newStatus,
        last_processed_at: new Date().toISOString(),
      })
      .eq('id', jobId);

    if (updateErr) throw updateErr;

    job.processed = processed;
    job.status = newStatus;
    return statusResponse(job);
  } catch (e) {
    return jsonResponse({ error: 'خطأ أثناء المعالجة: ' + String(e?.message || e) }, 500);
  }
});

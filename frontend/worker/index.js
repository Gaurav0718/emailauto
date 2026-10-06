/**
 * Worker entry point. This project deploys as a Worker with static assets
 * (not classic Cloudflare Pages), so a single script routes /api/* requests
 * here and falls through to the static site (env.ASSETS) for everything
 * else — there is no functions/ auto-routing in this deploy model.
 */

import {
  prepareHtml,
  loadRecipientsFromBytes,
  bucketsFromRows,
  summarize,
  sendEmail,
  DEFAULT_SENDER,
  DEFAULT_SUBJECT,
} from "./_lib/emailService.js";
import defaultTestRows from "./_lib/defaultTestList.json";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

async function handleHealth(env) {
  return json({
    status: "ok",
    sender: DEFAULT_SENDER,
    subject: DEFAULT_SUBJECT,
    send_provider: "resend",
    default_test_list_exists: true,
    resend_configured: !!env.RESEND_API_KEY,
  });
}

async function handleAnalyze(request) {
  const form = await request.formData();
  const htmlFile = form.get("html");
  if (!htmlFile) {
    return json({ error: "No HTML file uploaded (field 'html')." }, 400);
  }

  const htmlIn = await htmlFile.text();
  const tagline = form.get("tagline");
  const [htmlOut, changes] = prepareHtml(htmlIn, tagline);

  let recipients = null;
  let recipientError = null;
  const excelFile = form.get("excel");
  if (excelFile && excelFile.size > 0) {
    try {
      const buf = await excelFile.arrayBuffer();
      const buckets = await loadRecipientsFromBytes(buf);
      recipients = summarize(buckets);
    } catch (e) {
      recipientError = `Could not read recipient sheet: ${e.message}`;
    }
  }

  return json({
    changes,
    transformed_html: htmlOut,
    recipients,
    recipient_error: recipientError,
  });
}

async function handleSend(request, env) {
  const form = await request.formData();
  const htmlFile = form.get("html");
  if (!htmlFile) {
    return json({ error: "No HTML file uploaded (field 'html')." }, 400);
  }

  const mode = (form.get("mode") || "test").toLowerCase();
  const subject = form.get("subject") || DEFAULT_SUBJECT;
  const sender = form.get("sender") || DEFAULT_SENDER;

  if (!env.RESEND_API_KEY) {
    return json({ error: "RESEND_API_KEY is not configured on the server." }, 500);
  }

  const htmlIn = await htmlFile.text();
  const tagline = form.get("tagline");
  const [htmlOut, changes] = prepareHtml(htmlIn, tagline);

  let buckets;
  try {
    if (mode === "real") {
      const excelFile = form.get("excel");
      if (!excelFile || excelFile.size === 0) {
        return json({ error: "Recipient sheet required for a real send." }, 400);
      }
      const buf = await excelFile.arrayBuffer();
      buckets = await loadRecipientsFromBytes(buf);
    } else {
      buckets = bucketsFromRows(defaultTestRows);
    }
  } catch (e) {
    return json({ error: `Could not read recipient sheet: ${e.message}` }, 400);
  }

  let delivered;
  try {
    delivered = await sendEmail({
      html: htmlOut,
      buckets,
      subject,
      sender,
      resendApiKey: env.RESEND_API_KEY,
    });
  } catch (e) {
    const status = e.status === 401 || e.status === 403 ? 401 : 502;
    return json({ error: e.message }, status);
  }

  return json({
    ok: true,
    mode,
    changes,
    subject,
    sender,
    recipients: summarize(buckets),
    delivered_count: delivered.length,
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    try {
      if (url.pathname === "/api/health" && request.method === "GET") {
        return await handleHealth(env);
      }
      if (url.pathname === "/api/analyze" && request.method === "POST") {
        return await handleAnalyze(request);
      }
      if (url.pathname === "/api/send" && request.method === "POST") {
        return await handleSend(request, env);
      }
    } catch (e) {
      return json({ error: `Unexpected server error: ${e.message}` }, 500);
    }

    // Not an API route — serve the static site.
    return env.ASSETS.fetch(request);
  },
};

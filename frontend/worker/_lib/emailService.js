/**
 * Cloudflare Pages Functions port of the Python email_service.py.
 * Same three jobs: HTML fit-to-screen transform, recipient-sheet parsing,
 * and sending (now via Resend's HTTP API instead of raw SMTP, since
 * Workers can't open TCP sockets to smtp.office365.com).
 */

import * as XLSX from "xlsx";

export const DEFAULT_SENDER = "nudge-app@indegene.com";
export const DEFAULT_SUBJECT = "Account Intelligence Update";

// Which columns hold the data (1-based, matching the sheet).
const EMAIL_COL = 3; // column 3 -> email address
const DISPOSITION_COL = 4; // column 4 -> to / cc / bcc
const VALID_DISPOSITIONS = new Set(["to", "cc", "bcc"]);

// =========================================================================
// 1. HTML transform: fit-to-screen, no mobile reflow
// =========================================================================

function stripMaxWidthMediaBlocks(html) {
  let out = "";
  let removed = 0;
  let i = 0;
  const lower = html.toLowerCase();

  while (true) {
    const at = lower.indexOf("@media", i);
    if (at === -1) {
      out += html.slice(i);
      break;
    }
    const brace = html.indexOf("{", at);
    if (brace === -1) {
      out += html.slice(i);
      break;
    }
    const query = html.slice(at, brace).toLowerCase();

    let depth = 0;
    let j = brace;
    while (j < html.length) {
      const c = html[j];
      if (c === "{") depth++;
      else if (c === "}") {
        depth--;
        if (depth === 0) {
          j++;
          break;
        }
      }
      j++;
    }

    if (query.includes("max-width")) {
      out += html.slice(i, at);
      removed++;
      i = j;
    } else {
      out += html.slice(i, j);
      i = j;
    }
  }

  return [out, removed];
}

export function addOutlookBgcolors(html) {
  let added = 0;
  const tagRe = /<(td|th|tr|table)\b([^>]*)>/gi;

  const out = html.replace(tagRe, (match, tag, attrs) => {
    if (/\bbgcolor\s*=/i.test(attrs)) return match; // already has one
    const colorMatch = attrs.match(/background(?:-color)?\s*:\s*(#[0-9a-fA-F]{3,6})/i);
    if (!colorMatch) return match;
    added++;
    return `<${tag} bgcolor="${colorMatch[1]}"${attrs}>`;
  });

  return [out, added];
}

export function makeFitToScreen(html) {
  const changes = [];

  // 1. Remove x-apple-disable-message-reformatting meta (blocks auto-scale).
  const before1 = html;
  html = html.replace(
    /[ \t]*<meta[^>]*x-apple-disable-message-reformatting[^>]*>\s*\n?/gi,
    ""
  );
  if (html !== before1) {
    changes.push("Removed x-apple-disable-message-reformatting meta");
  }

  // 2. Force viewport to width=640 (replace existing, or insert into head).
  const viewportTag = '<meta name="viewport" content="width=640">';
  if (/<meta[^>]+name=["']viewport["']/i.test(html)) {
    const newHtml = html.replace(/<meta[^>]+name=["']viewport["'][^>]*>/i, viewportTag);
    if (newHtml !== html) changes.push("Set viewport to fixed width=640");
    html = newHtml;
  } else if (/<head[^>]*>/i.test(html)) {
    html = html.replace(/(<head[^>]*>)/i, `$1\n${viewportTag}`);
    changes.push("Inserted fixed viewport width=640");
  } else {
    html = `<head>\n${viewportTag}\n</head>\n${html}`;
    changes.push("Inserted fixed viewport width=640");
  }

  // 3. Strip mobile reflow @media blocks.
  const [strippedHtml, removed] = stripMaxWidthMediaBlocks(html);
  html = strippedHtml;
  if (removed) {
    changes.push(`Removed ${removed} mobile reflow @media block${removed !== 1 ? "s" : ""}`);
  }

  // 4. Add Outlook bgcolor fallbacks so colored bands don't render white.
  const [bgHtml, added] = addOutlookBgcolors(html);
  html = bgHtml;
  if (added) {
    changes.push(`Added Outlook color fallbacks to ${added} element(s)`);
  }

  if (changes.length === 0) {
    changes.push("Already fit-to-screen — no changes needed");
  }

  return [html, changes];
}

export function applyTagline(html, tagline) {
  const bylineText = (tagline || "").trim();
  if (!bylineText) return [html, null];

  const safe = bylineText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  // 1. Replace an existing byline if present.
  let replaced = false;
  let newHtml = html.replace(/For\s+[^<>]{1,120}?portfolio/i, () => {
    replaced = true;
    return safe;
  });
  if (replaced) {
    return [newHtml, `Set portfolio tagline to “${bylineText}”`];
  }

  // 2. No byline yet — insert one after the banner paragraph.
  const bylineP =
    '<p style="margin:0 0 18px 0;font-size:18px;line-height:24px;' +
    'font-weight:800;color:#FFFFFF;">' + safe + "</p>";

  const bannerRe = /<p[^>]*>[^<]*?(?:portfolio intelligence|verified signals|signals)[^<]*?<\/p>/i;
  const bannerMatch = html.match(bannerRe);
  if (bannerMatch) {
    const at = bannerMatch.index + bannerMatch[0].length;
    html = html.slice(0, at) + "\n    " + bylineP + html.slice(at);
    return [html, `Inserted portfolio tagline “${bylineText}”`];
  }

  // 3. Fallback — put it just before the first headline.
  const h1Match = html.match(/<h1\b/i);
  if (h1Match) {
    const at = h1Match.index;
    html = html.slice(0, at) + bylineP + "\n    " + html.slice(at);
    return [html, `Inserted portfolio tagline “${bylineText}”`];
  }

  return [html, null];
}

export function prepareHtml(html, tagline) {
  const changes = [];
  const [taggedHtml, tagChange] = applyTagline(html, tagline);
  html = taggedHtml;
  if (tagChange) changes.push(tagChange);

  const [fitHtml, fitChanges] = makeFitToScreen(html);
  html = fitHtml;
  changes.push(...fitChanges);

  return [html, changes];
}

// =========================================================================
// 2. Recipient loading
// =========================================================================

function rowsToBuckets(rows) {
  const buckets = { to: [], cc: [], bcc: [] };
  for (const row of rows) {
    if (row.length < Math.max(EMAIL_COL, DISPOSITION_COL)) continue;
    let email = row[EMAIL_COL - 1];
    let disposition = row[DISPOSITION_COL - 1];
    if (email == null || disposition == null) continue;
    email = String(email).trim();
    disposition = String(disposition).trim().toLowerCase();
    if (!email.includes("@") || !VALID_DISPOSITIONS.has(disposition)) continue;
    buckets[disposition].push(email);
  }
  return buckets;
}

export function bucketsFromRows(rows) {
  return rowsToBuckets(rows);
}

export async function loadRecipientsFromBytes(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  return rowsToBuckets(rows);
}

export function summarize(buckets) {
  return {
    to: buckets.to,
    cc: buckets.cc,
    bcc: buckets.bcc,
    total: buckets.to.length + buckets.cc.length + buckets.bcc.length,
  };
}

// =========================================================================
// 3. Sending (via Resend HTTP API)
// =========================================================================

export async function sendEmail({ html, buckets, subject, sender, resendApiKey }) {
  const allRecipients = [...buckets.to, ...buckets.cc, ...buckets.bcc];
  if (allRecipients.length === 0) {
    throw new Error("No valid recipients found.");
  }
  if (buckets.to.length === 0) {
    throw new Error("Sheet has no 'to' recipient. At least one is required.");
  }

  const body = {
    from: sender,
    to: buckets.to,
    subject,
    html,
    text: "This message is best viewed in an HTML-capable email client.",
  };
  if (buckets.cc.length) body.cc = buckets.cc;
  if (buckets.bcc.length) body.bcc = buckets.bcc;

  const r = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${resendApiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!r.ok) {
    const errBody = await r.json().catch(() => ({}));
    const err = new Error(errBody.message || `Resend API error (${r.status})`);
    err.status = r.status;
    throw err;
  }

  return allRecipients;
}

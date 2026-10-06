import React, { useEffect, useRef, useState } from "react";

// In dev this is empty and the Vite proxy forwards /api to the backend.
// In production (Cloudflare Pages) set VITE_API_BASE to the backend URL.
const API = (import.meta.env.VITE_API_BASE || "").replace(/\/$/, "");

const DEFAULTS = {
  sender: "nudge-app@indegene.com",
  subject: "Account Intelligence Update",
};

export default function App() {
  const [htmlFile, setHtmlFile] = useState(null);
  const [excelFile, setExcelFile] = useState(null);

  const [sender, setSender] = useState(DEFAULTS.sender);
  const [subject, setSubject] = useState(DEFAULTS.subject);
  const [tagline, setTagline] = useState("");
  const [smtpUser, setSmtpUser] = useState(DEFAULTS.sender);
  const [smtpPass, setSmtpPass] = useState("");

  const [changes, setChanges] = useState([]);
  const [transformedHtml, setTransformedHtml] = useState("");
  const [recipients, setRecipients] = useState(null);
  const [recipientError, setRecipientError] = useState(null);

  const [previewMode, setPreviewMode] = useState("mobile");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState(null); // {type, text}
  const [health, setHealth] = useState(null);

  // Backend wake state. On a sleeping host (e.g. Render free tier) the first
  // request can take ~50s while the server cold-starts, so we poll /api/health
  // and show a live timer until it answers.
  const [backendStatus, setBackendStatus] = useState("connecting"); // connecting | ready | error
  const [wakeSeconds, setWakeSeconds] = useState(0);
  const [pingNonce, setPingNonce] = useState(0);

  const iframeRef = useRef(null);

  // Poll the backend until it responds (cold start friendly).
  useEffect(() => {
    let cancelled = false;
    const startedAt = Date.now();
    const MAX_WAIT_MS = 120000; // give up after 2 min

    setBackendStatus("connecting");
    setWakeSeconds(0);

    const tick = setInterval(() => {
      if (cancelled) return;
      const elapsed = Date.now() - startedAt;
      setWakeSeconds(Math.floor(Math.min(elapsed, MAX_WAIT_MS) / 1000));
    }, 1000);

    async function poll() {
      while (!cancelled) {
        try {
          const r = await fetch(`${API}/api/health`, { cache: "no-store" });
          if (r.ok) {
            const data = await r.json();
            if (!cancelled) {
              setHealth(data);
              setBackendStatus("ready");
              clearInterval(tick); // stop the timer once settled
            }
            return;
          }
        } catch (_) {
          // backend still waking or unreachable — keep trying
        }
        if (Date.now() - startedAt > MAX_WAIT_MS) {
          if (!cancelled) {
            setBackendStatus("error");
            clearInterval(tick); // stop the timer once settled
          }
          return;
        }
        await new Promise((res) => setTimeout(res, 2500));
      }
    }
    poll();

    return () => {
      cancelled = true;
      clearInterval(tick);
    };
  }, [pingNonce]);

  // Re-analyze whenever the HTML or Excel changes.
  useEffect(() => {
    if (!htmlFile) {
      setChanges([]);
      setTransformedHtml("");
      setRecipients(null);
      setRecipientError(null);
      return;
    }
    // Debounce so typing the tagline doesn't fire a request per keystroke.
    const timer = setTimeout(() => {
      const form = new FormData();
      form.append("html", htmlFile);
      if (excelFile) form.append("excel", excelFile);
      if (tagline.trim()) form.append("tagline", tagline.trim());

      setBusy(true);
      fetch(`${API}/api/analyze`, { method: "POST", body: form })
        .then((r) => r.json())
        .then((data) => {
          if (data.error) {
            setMessage({ type: "err", text: data.error });
            return;
          }
          setChanges(data.changes || []);
          setTransformedHtml(data.transformed_html || "");
          setRecipients(data.recipients || null);
          setRecipientError(data.recipient_error || null);
          setMessage(null);
        })
        .catch((e) => setMessage({ type: "err", text: String(e) }))
        .finally(() => setBusy(false));
    }, 350);

    return () => clearTimeout(timer);
  }, [htmlFile, excelFile, tagline]);

  // Write the transformed HTML into the preview iframe.
  useEffect(() => {
    const iframe = iframeRef.current;
    if (!iframe) return;
    const doc = iframe.contentDocument;
    if (!doc) return;
    doc.open();
    doc.write(transformedHtml || "<p style='padding:20px;color:#8894a6'>Upload an HTML file to preview.</p>");
    doc.close();
  }, [transformedHtml, previewMode]);

  async function doSend(mode) {
    if (!htmlFile) return;
    if (!smtpPass) {
      setMessage({ type: "err", text: "Enter the SMTP app password first." });
      return;
    }
    if (mode === "real") {
      const total = recipients?.total || 0;
      const ok = window.confirm(
        `Send "${subject}" to ${total} recipient(s) from the uploaded list?\n\nThis actually delivers email and cannot be undone.`
      );
      if (!ok) return;
    }

    const form = new FormData();
    form.append("html", htmlFile);
    if (excelFile) form.append("excel", excelFile);
    form.append("mode", mode);
    form.append("subject", subject);
    form.append("sender", sender);
    if (tagline.trim()) form.append("tagline", tagline.trim());
    form.append("smtp_user", smtpUser);
    form.append("smtp_pass", smtpPass);

    setBusy(true);
    setMessage({ type: "info", text: mode === "test" ? "Sending test run..." : "Sending..." });
    try {
      const r = await fetch(`${API}/api/send`, { method: "POST", body: form });
      const data = await r.json();
      if (!r.ok || data.error) {
        setMessage({ type: "err", text: data.error || "Send failed." });
      } else {
        const rc = data.recipients;
        setMessage({
          type: "ok",
          text:
            `${mode === "test" ? "Test run" : "Send"} complete — delivered to ` +
            `${data.delivered_count} recipient(s) ` +
            `(to ${rc.to.length}, cc ${rc.cc.length}, bcc ${rc.bcc.length}).`,
        });
      }
    } catch (e) {
      setMessage({ type: "err", text: String(e) });
    } finally {
      setBusy(false);
    }
  }

  const ready = backendStatus === "ready";
  const canTest = ready && !!htmlFile && !!smtpPass && !busy;
  const canSend = ready && !!htmlFile && !!excelFile && !!smtpPass && !busy && (recipients?.total || 0) > 0;

  return (
    <div className="app">
      <h1>NUDGE Email Sender</h1>
      <p className="sub">
        Upload an HTML email + recipient sheet. Every HTML is auto-adjusted to
        fit-to-screen (no mobile reflow) before sending — exactly like the
        hand-sent versions.
      </p>

      <BackendStatus
        status={backendStatus}
        seconds={wakeSeconds}
        onRetry={() => setPingNonce((n) => n + 1)}
      />

      <div className="grid">
        {/* ---- Left column: controls ---- */}
        <div>
          <div className="card">
            <h2>1. Email HTML</h2>
            <div className={"file-row" + (htmlFile ? " done" : "")}>
              <input
                type="file"
                accept=".html,.htm"
                onChange={(e) => setHtmlFile(e.target.files[0] || null)}
              />
            </div>
            {htmlFile && changes.length > 0 && (
              <ul className="changes">
                {changes.map((c, i) => (
                  <li key={i}>{c}</li>
                ))}
              </ul>
            )}
          </div>

          <div className="card">
            <h2>2. Recipient sheet (.xlsx)</h2>
            <p className="hint" style={{ marginTop: 0 }}>
              Column 3 = email, Column 4 = to / cc / bcc. Needed for a real
              send. Test runs use the default test list.
            </p>
            <div className={"file-row" + (excelFile ? " done" : "")}>
              <input
                type="file"
                accept=".xlsx"
                onChange={(e) => setExcelFile(e.target.files[0] || null)}
              />
            </div>
            {recipientError && <div className="msg err">{recipientError}</div>}
            {recipients && (
              <div style={{ marginTop: 10 }}>
                <RecipientChips label="To" list={recipients.to} kind="to" />
                <RecipientChips label="Cc" list={recipients.cc} kind="cc" />
                <RecipientChips label="Bcc" list={recipients.bcc} kind="bcc" collapse />
                <div className="hint">Total: {recipients.total} recipient(s)</div>
              </div>
            )}
          </div>

          <div className="card">
            <h2>3. Message + SMTP</h2>
            <label>Portfolio tagline (optional)</label>
            <input
              type="text"
              value={tagline}
              placeholder={"e.g. For Gobin Chandra’s portfolio"}
              onChange={(e) => setTagline(e.target.value)}
            />
            <p className="hint" style={{ marginTop: 4 }}>
              Used exactly as typed. Sets the byline in the hero band. Leave
              blank to keep the HTML as-is.
            </p>
            <label>Subject</label>
            <input type="text" value={subject} onChange={(e) => setSubject(e.target.value)} />
            <label>From (sender)</label>
            <input type="text" value={sender} onChange={(e) => setSender(e.target.value)} />
            <label>SMTP username</label>
            <input type="text" value={smtpUser} onChange={(e) => setSmtpUser(e.target.value)} />
            <label>SMTP app password</label>
            <input
              type="password"
              value={smtpPass}
              placeholder="Office365 app password"
              onChange={(e) => setSmtpPass(e.target.value)}
              autoComplete="off"
            />
            <p className="hint">
              Not stored — used only for this send. Office365 needs SMTP AUTH
              enabled + an app password if MFA is on.
            </p>
          </div>

          <div className="card">
            <h2>4. Send</h2>
            <button className="btn btn-test" disabled={!canTest} onClick={() => doSend("test")}>
              Test Run
              {health?.default_test_list && (
                <span className="count-badge">→ {health.default_test_list}</span>
              )}
            </button>
            <p className="hint">Sends to the default test list, regardless of the uploaded sheet.</p>

            <button className="btn btn-send" disabled={!canSend} onClick={() => doSend("real")}>
              Send Emails
              {recipients && <span className="count-badge" style={{ color: "#cbd5e6" }}>→ {recipients.total}</span>}
            </button>
            <p className="hint">Sends to the uploaded recipient list. Asks for confirmation first.</p>

            {message && <div className={"msg " + message.type}>{message.text}</div>}
          </div>
        </div>

        {/* ---- Right column: live preview ---- */}
        <div className="preview-wrap">
          <div className="card">
            <h2>Preview (transformed HTML)</h2>
            <div className="preview-toolbar">
              <button
                className={previewMode === "mobile" ? "active" : ""}
                onClick={() => setPreviewMode("mobile")}
              >
                Mobile
              </button>
              <button
                className={previewMode === "desktop" ? "active" : ""}
                onClick={() => setPreviewMode("desktop")}
              >
                Desktop
              </button>
            </div>
            <div style={{ display: "flex", justifyContent: "center", overflow: "auto" }}>
              <iframe
                ref={iframeRef}
                title="preview"
                className="preview-frame"
                style={{ width: previewMode === "mobile" ? "390px" : "100%", maxWidth: "100%" }}
              />
            </div>
            <p className="hint">
              This is the exact HTML that will be sent. Mobile view shows the
              fit-to-screen scaling recipients will see on a phone.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

function BackendStatus({ status, seconds, onRetry }) {
  if (status === "ready") {
    return (
      <div className="backend-status ready">
        <span className="dot" /> Backend connected
      </div>
    );
  }
  if (status === "error") {
    return (
      <div className="backend-status error">
        <span className="dot" /> Backend didn&apos;t respond after {seconds}s.
        <button onClick={onRetry}>Retry</button>
      </div>
    );
  }
  // connecting
  const slow = seconds >= 5;
  return (
    <div className="backend-status connecting">
      <span className="spinner" />
      {slow ? (
        <span>
          Waking the backend… <strong>{seconds}s</strong>
          <span className="sub"> (a sleeping server can take ~50s on the first request)</span>
        </span>
      ) : (
        <span>Connecting to backend… <strong>{seconds}s</strong></span>
      )}
    </div>
  );
}

function RecipientChips({ label, list, kind, collapse }) {
  const [open, setOpen] = useState(!collapse);
  if (!list || list.length === 0) return null;
  const shown = open ? list : list.slice(0, 0);
  return (
    <div className="chips">
      <strong style={{ fontSize: 12, color: "#1b365d" }}>
        {label} ({list.length})
      </strong>
      {collapse && (
        <button
          onClick={() => setOpen((o) => !o)}
          style={{
            marginLeft: 8,
            fontSize: 11,
            border: "1px solid #dfe4ec",
            borderRadius: 12,
            background: "#fff",
            cursor: "pointer",
            padding: "1px 8px",
          }}
        >
          {open ? "hide" : "show"}
        </button>
      )}
      <div>
        {shown.map((e, i) => (
          <span key={i} className={"tag " + kind}>
            {e}
          </span>
        ))}
      </div>
    </div>
  );
}

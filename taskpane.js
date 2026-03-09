/**
 * taskpane.js
 * Office.js add-in logic for AI Email Enrichment.
 *
 * FLOW:
 *   Office.onReady → read email metadata
 *   analyseEmail()  → POST to /api/analyse → render result cards
 */

// ── CONFIG ─────────────────────────────────────────────────────────────────
// Replace with the URL of your deployed backend (Azure Function, Express, etc.)
const API_URL = "https://devious-postmedian-makai.ngrok-free.dev/api/analyse";
// ───────────────────────────────────────────────────────────────────────────

let _emailData = null;   // cache the current email once read
let _startTime = null;

// ── OFFICE INIT ────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showError("This add-in only works inside Outlook.");
    return;
  }

  // Attempt auto-read of current item
  readCurrentEmail()
    .then((data) => {
      _emailData = data;
      setFooterMeta(`From: ${data.senderEmail}`);
      setHeaderSub(`${data.subject.slice(0, 42)}${data.subject.length > 42 ? "…" : ""}`);
    })
    .catch((err) => {
      console.warn("Email pre-read failed:", err);
      setHeaderSub("Open an email to begin");
    });
});

// ── READ EMAIL FROM OFFICE.JS ──────────────────────────────────────────────
function readCurrentEmail() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    if (!item) {
      return reject(new Error("No email item found."));
    }

    // 1. Get the full body (HTML or text)
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: "body" }, (bodyResult) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("Could not read email body: " + bodyResult.error.message));
      }

      const body = bodyResult.value;

      // 2. Collect synchronous fields
      const data = {
        subject:      item.subject            || "(no subject)",
        senderName:   item.from?.displayName  || "Unknown",
        senderEmail:  item.from?.emailAddress || "unknown@unknown.com",
        dateReceived: item.dateTimeCreated
                        ? new Date(item.dateTimeCreated).toISOString()
                        : new Date().toISOString(),
        bodyText:     body.trim(),
        // Trim for API call — first 4000 chars is more than enough for analysis
        bodyTrimmed:  body.trim().slice(0, 4000),
      };

      resolve(data);
    });
  });
}

// ── MAIN: ANALYSE EMAIL ────────────────────────────────────────────────────
async function analyseEmail() {
  const btn = document.getElementById("analyse-btn");

  // Re-read email in case user switched items
  try {
    showLoading("Reading email…");
    btn.disabled = true;
    _emailData = await readCurrentEmail();
  } catch (err) {
    showError("Could not read email: " + err.message);
    btn.disabled = false;
    return;
  }

  // Call backend
  try {
    setLoadingMsg("Analysing with AI…");
    _startTime = Date.now();

    const response = await fetch(API_URL, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject:      _emailData.subject,
        senderName:   _emailData.senderName,
        senderEmail:  _emailData.senderEmail,
        dateReceived: _emailData.dateReceived,
        bodyText:     _emailData.bodyTrimmed,
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => response.statusText);
      throw new Error(`Backend error ${response.status}: ${errText}`);
    }

    const result = await response.json();
    const elapsed = ((Date.now() - _startTime) / 1000).toFixed(1);

    renderResult(result, elapsed);

  } catch (err) {
    console.error("AI analysis failed:", err);
    showError(err.message || "Network error — check the backend is running.");
  } finally {
    btn.disabled = false;
  }
}

// ── RENDER RESULT ──────────────────────────────────────────────────────────
/**
 * Expected result shape from backend:
 * {
 *   summary:   string,
 *   category:  string,           // e.g. "COMPLAINT", "INQUIRY", "REQUEST", "INFO"
 *   urgency:   string,           // "CRITICAL" | "HIGH" | "MEDIUM" | "LOW"
 *   actions:   string[],         // 1–3 short action strings
 *   tone: {
 *     negative: number,          // 0–100
 *     neutral:  number,
 *     positive: number,
 *   }
 * }
 */
function renderResult(r, elapsed) {
  // Summary
  setText("res-summary", r.summary || "No summary available.");

  // Chips (category + urgency)
  const chipsEl = document.getElementById("res-chips");
  chipsEl.innerHTML = "";
  [r.category, r.urgency].filter(Boolean).forEach((label) => {
    const chip = document.createElement("span");
    const key  = label.toUpperCase().replace(/\s+/g, "_");
    chip.className = `badge badge-${key}`;
    chip.innerHTML = `<span class="badge-dot"></span>${label}`;
    chipsEl.appendChild(chip);
  });

  // Actions
  const actionsEl = document.getElementById("res-actions");
  actionsEl.innerHTML = "";
  const actions = r.actions || [];
  if (actions.length === 0) {
    actionsEl.textContent = "No specific actions suggested.";
  } else {
    actions.forEach((action, i) => {
      actionsEl.innerHTML += `
        <div class="action-item">
          <span class="action-num">${String(i + 1).padStart(2, "0")}</span>
          <span class="action-text">${escapeHtml(action)}</span>
        </div>`;
    });
  }

  // Tone bars
  const toneEl = document.getElementById("res-tone");
  toneEl.innerHTML = "";
  if (r.tone) {
    const tones = [
      { label: "Negative", key: "negative", cls: "negative" },
      { label: "Neutral",  key: "neutral",  cls: "neutral"  },
      { label: "Positive", key: "positive", cls: "positive" },
    ];
    tones.forEach(({ label, key, cls }) => {
      const pct = Math.round(r.tone[key] || 0);
      toneEl.innerHTML += `
        <div class="tone-row" style="margin-bottom:8px">
          <span class="tone-label">${label}</span>
          <div class="tone-bar-track">
            <div class="tone-bar-fill ${cls}" style="width:${pct}%"></div>
          </div>
          <span class="tone-pct">${pct}%</span>
        </div>`;
    });
  } else {
    toneEl.textContent = "Tone data unavailable.";
  }

  // Footer
  setFooterMeta(`Analysed in ${elapsed}s · ${_emailData.senderEmail}`);
  document.getElementById("retry-btn").style.display = "block";
  setHeaderSub("Analysis complete");

  showState("result");
}

// ── UI HELPERS ─────────────────────────────────────────────────────────────
function showState(state) {
  ["idle", "loading", "error", "result"].forEach((s) => {
    document.getElementById(`state-${s}`).style.display = s === state ? "flex" : "none";
  });
}

function showLoading(msg) {
  setLoadingMsg(msg || "Analysing…");
  showState("loading");
}

function setLoadingMsg(msg) {
  const el = document.getElementById("loading-msg");
  if (el) el.textContent = msg;
}

function showError(msg) {
  document.getElementById("error-msg").textContent = msg;
  showState("error");
}

function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function setFooterMeta(text) {
  const el = document.getElementById("footer-meta");
  if (el) el.textContent = text;
}

function setHeaderSub(text) {
  const el = document.getElementById("header-sub");
  if (el) el.textContent = text;
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

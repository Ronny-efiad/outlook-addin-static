/**
 * taskpane.js
 * Office.js add-in logic for AI Email Enrichment.
 *
 * FLOW:
 *   Office.onReady → read email metadata
 *   analyseEmail()  → POST to Power Automate → render product match
 */

// ── CONFIG ─────────────────────────────────────────────────────────────────
// Power Automate HTTP trigger URL
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

    item.body.getAsync(Office.CoercionType.Text, { asyncContext: "body" }, (bodyResult) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("Could not read email body: " + bodyResult.error.message));
      }

      const body = bodyResult.value;

      const data = {
        subject:      item.subject            || "(no subject)",
        senderName:   item.from?.displayName  || "Unknown",
        senderEmail:  item.from?.emailAddress || "unknown@unknown.com",
        dateReceived: item.dateTimeCreated
                        ? new Date(item.dateTimeCreated).toISOString()
                        : new Date().toISOString(),
        bodyText:     body.trim(),
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

  // Call Power Automate flow
  try {
    setLoadingMsg("Analysing with AI Builder…");
    _startTime = Date.now();

    const response = await fetch(API_URL, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject:     _emailData.subject,
        body:        _emailData.bodyTrimmed,
        sender:      `${_emailData.senderName} <${_emailData.senderEmail}>`,
        senderEmail: _emailData.senderEmail,
        senderName:  _emailData.senderName,
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => response.statusText);
      throw new Error(`Power Automate error ${response.status}: ${errText}`);
    }

    const result = await response.json();
    const elapsed = ((Date.now() - _startTime) / 1000).toFixed(1);

    renderResult(result, elapsed);

  } catch (err) {
    console.error("AI analysis failed:", err);
    showError(err.message || "Network error — check the Power Automate flow.");
  } finally {
    btn.disabled = false;
  }
}

// ── RENDER RESULT ──────────────────────────────────────────────────────────
/**
 * Expected result shape:
 * {
 *   productNumber: string,
 *   productName:   string,
 *   confidence:    "HIGH" | "MEDIUM" | "LOW",
 *   reasoning:     string
 * }
 */
function renderResult(raw, elapsed) {
  let r = extractProductMatch(raw);

  if (!r) {
    showError("Could not parse AI response. Check the flow output.");
    console.error("Raw response:", raw);
    return;
  }

  // Customer card
  const hasCustomer = r.customerName || r.customerNumber;
  document.getElementById("card-customer").style.display = hasCustomer ? "block" : "none";
  setText("res-customer-name",   r.customerName   || "—");
  setText("res-customer-number", r.customerNumber || "—");
  setText("res-domain",          r.domain         || "—");

  // Product
  setText("res-product-name",   r.productName   || "Unknown");
  setText("res-product-number", r.productNumber || "UNKNOWN");

  // Confidence badge
  const confidenceEl = document.getElementById("res-confidence");
  confidenceEl.innerHTML = "";
  if (r.confidence) {
    const chip = document.createElement("span");
    const key = r.confidence.toUpperCase();
    chip.className = `badge badge-${key}`;
    chip.innerHTML = `<span class="badge-dot"></span>${r.confidence}`;
    confidenceEl.appendChild(chip);
  }

  setText("res-reasoning", r.reasoning || "No reasoning provided.");

  setFooterMeta(`Analysed in ${elapsed}s · ${_emailData.senderEmail}`);
  document.getElementById("retry-btn").style.display = "block";
  setHeaderSub(r.customerName || "Product match found");

  showState("result");
}

function extractProductMatch(raw) {
  // Direct match
  if (raw && raw.productNumber) return raw;

  // Power Automate / AI Builder nested path:
  // responsev2 → predictionOutput → text (JSON string)
  try {
    const text = raw?.responsev2?.predictionOutput?.text;
    if (text) {
      const parsed = tryParseJSON(text);
      if (parsed && parsed.productNumber) return parsed;
    }
  } catch (_) {}

  // Generic deep search through all string values
  if (raw && typeof raw === "object") {
    for (const val of Object.values(raw)) {
      if (typeof val === "string") {
        const parsed = tryParseJSON(val);
        if (parsed && parsed.productNumber) return parsed;
      }
      if (typeof val === "object" && val !== null) {
        const nested = extractProductMatch(val);
        if (nested) return nested;
      }
    }
  }

  if (typeof raw === "string") {
    const parsed = tryParseJSON(raw);
    if (parsed && parsed.productNumber) return parsed;
  }

  return null;
}

function tryParseJSON(str) {
  try { return JSON.parse(str); } catch (_) {}
  const stripped = str.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();
  try { return JSON.parse(stripped); } catch (_) {}
  return null;
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

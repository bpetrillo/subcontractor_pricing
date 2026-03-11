import { useState, useEffect } from "react";

// ── CONFIG ────────────────────────────────────────────────────────────────────
const CLIENT_ID  = "68dbf196-5b89-4932-bb18-cc9b36cd19b8";
const TENANT_ID  = "ee3eca2c-e65d-4770-a009-1f49aa38b996";
const SITE_NAME  = "SunbeltUtilitiesCorp";
const CHANNEL    = "General";
const FOLDER     = "Project Trackers";
const WORKBOOK   = "Job Tracker.xlsx";
const SHEET      = "subs";

const SCOPES = [
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Files.Read.All",
].join(" ");

// ── AUTH ──────────────────────────────────────────────────────────────────────
function getAccessToken() {
  return new Promise((resolve, reject) => {
    const redirectUri = window.location.origin;
    const url =
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
      new URLSearchParams({
        client_id: CLIENT_ID,
        response_type: "token",
        redirect_uri: redirectUri,
        scope: SCOPES,
        response_mode: "fragment",
        nonce: Math.random().toString(36),
      });

    const popup = window.open(url, "msalPopup", "width=500,height=650");
    if (!popup) return reject(new Error("Popup blocked — please allow popups for this page."));

    const timer = setInterval(() => {
      try {
        const hash = popup.location.hash;
        if (hash && hash.includes("access_token")) {
          clearInterval(timer);
          popup.close();
          const p = new URLSearchParams(hash.slice(1));
          resolve(p.get("access_token"));
        }
        if (popup.closed) {
          clearInterval(timer);
          reject(new Error("Sign-in window closed before completing."));
        }
      } catch (_) {}
    }, 300);
  });
}

// ── GRAPH HELPERS ─────────────────────────────────────────────────────────────
async function graphGet(token, path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    throw new Error(e?.error?.message || `HTTP ${res.status}: ${path}`);
  }
  return res.json();
}

async function loadSubsFromExcel(token) {
  const siteRes = await graphGet(token, `/sites?search=${SITE_NAME}`);
  const site = siteRes.value?.[0];
  if (!site) throw new Error(`SharePoint site "${SITE_NAME}" not found.`);

  const drivesRes = await graphGet(token, `/sites/${site.id}/drives`);
  const drive = drivesRes.value?.find((d) => d.name === "Documents") || drivesRes.value?.[0];
  if (!drive) throw new Error("Could not find the Documents drive.");

  const filePath = `${CHANNEL}/${FOLDER}/${WORKBOOK}`;
  const fileRes = await graphGet(token, `/drives/${drive.id}/root:/${filePath}`);
  const fileId = fileRes.id;

  const rangeRes = await graphGet(
    token,
    `/drives/${drive.id}/items/${fileId}/workbook/worksheets/${SHEET}/usedRange`
  );

  const rows = rangeRes.values;
  if (!rows || rows.length < 2) throw new Error(`The "${SHEET}" sheet appears empty.`);

  const headers = rows[0].map((h) => String(h).trim().toLowerCase());
  const col = (name) => headers.findIndex((h) => h === name);
  const iName = col("name"), iCompany = col("company"), iEmail = col("email"), iTrade = col("trade");

  if (iEmail === -1) throw new Error(`Could not find an "email" column in the "${SHEET}" sheet.`);

  return rows.slice(1)
    .map((r) => ({
      name:    String(r[iName]    ?? "").trim(),
      company: String(r[iCompany] ?? "").trim(),
      email:   String(r[iEmail]   ?? "").trim(),
      trade:   String(r[iTrade]   ?? "").trim(),
    }))
    .filter((s) => s.email && s.email !== "undefined");
}

async function sendEmail(token, ccList, toEmail, subject, body) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: "Text", content: body },
        toRecipients: [{ emailAddress: { address: toEmail } }],
        ccRecipients: ccList.filter(Boolean).map((e) => ({ emailAddress: { address: e.trim() } })),
      },
      saveToSentItems: true,
    }),
  });
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    throw new Error(e?.error?.message || `HTTP ${res.status}`);
  }
}

// ── EMAIL TEMPLATE ────────────────────────────────────────────────────────────
function buildEmailBody(sub, project, fileLink, fileLabel, dueDate) {
  const firstName = (sub.name || "").split(" ")[0] || "there";
  const due = dueDate
    ? new Date(dueDate + "T12:00:00").toLocaleDateString("en-US", { month: "long", day: "numeric" })
    : "[due date TBD]";
  const label = fileLabel || project;
  return [
    `Hey ${firstName},`,
    "",
    `I have another project for you that's in my queue — ${project}. I wanted to get it into your hands as soon as possible.`,
    "",
    `When you have a chance, could you take a look at the plans and put together a quote for the usual?`,
    "",
    `You can access the plans here:\n${label} – ${fileLink}`,
    "",
    `If possible, could you have it back to me by ${due}?`,
    "",
    `Please let me know if you need anything else!`,
    "",
    `Thank you,`,
  ].join("\n");
}

// ── SHARED STYLES ─────────────────────────────────────────────────────────────
const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "10px 14px", borderRadius: 8,
  border: "1.5px solid #ddd", fontFamily: "'Lora', serif", fontSize: 14,
  color: "#1a1a2e", background: "#fdfcf9", outline: "none", transition: "border .2s",
};

const btnPrimary = (disabled) => ({
  padding: "11px 28px",
  background: disabled ? "#ccc" : "#1a1a2e",
  color: "#f5f0e8", border: "none", borderRadius: 9,
  fontFamily: "'DM Mono', monospace", fontSize: 13,
  cursor: disabled ? "not-allowed" : "pointer", transition: "all .2s",
});

const btnGhost = {
  padding: "11px 22px", background: "transparent", border: "1.5px solid #ddd",
  borderRadius: 9, fontFamily: "'DM Mono', monospace", fontSize: 13, cursor: "pointer",
};

const btnGreen = (disabled) => ({
  padding: "11px 28px", background: disabled ? "#aaa" : "#1a6b1a",
  color: "#fff", border: "none", borderRadius: 9,
  fontFamily: "'DM Mono', monospace", fontSize: 13,
  cursor: disabled ? "not-allowed" : "pointer", transition: "all .2s",
});

function Field({ label, hint, children }) {
  return (
    <div style={{ marginBottom: 20 }}>
      <label style={{ display: "block", fontFamily: "'DM Mono', monospace", fontSize: 11, color: "#888", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>
        {label}
      </label>
      {children}
      {hint && <p style={{ margin: "5px 0 0", fontSize: 12, color: "#aaa" }}>{hint}</p>}
    </div>
  );
}

function StepBar({ step }) {
  const labels = ["Project Info", "Select Subs", "Review & Send"];
  return (
    <div style={{ display: "flex", alignItems: "center", marginBottom: 36 }}>
      {labels.map((label, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", flex: i < labels.length - 1 ? 1 : 0 }}>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 5 }}>
            <div style={{
              width: 30, height: 30, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
              fontFamily: "'DM Mono', monospace", fontSize: 12, fontWeight: 600,
              background: step > i ? "#c8f5c8" : step === i ? "#1a1a2e" : "#e8e4dc",
              color: step > i ? "#1a6b1a" : step === i ? "#f5f0e8" : "#999",
              border: step === i ? "2px solid #1a1a2e" : "2px solid transparent",
            }}>{step > i ? "✓" : i + 1}</div>
            <span style={{ fontSize: 10, fontFamily: "'DM Mono', monospace", color: step === i ? "#1a1a2e" : "#bbb", whiteSpace: "nowrap" }}>{label}</span>
          </div>
          {i < labels.length - 1 && (
            <div style={{ flex: 1, height: 2, background: step > i ? "#1a1a2e" : "#e0e0e0", margin: "0 8px", marginBottom: 18, transition: "background .3s" }} />
          )}
        </div>
      ))}
    </div>
  );
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
export default function SubPricingTool() {
  const [step, setStep]               = useState(0);
  const [token, setToken]             = useState(null);
  const [subs, setSubs]               = useState([]);
  const [subsLoading, setSubsLoading] = useState(false);
  const [subsError, setSubsError]     = useState("");

  // Step 0
  const [project, setProject]   = useState("");
  const [fileLink, setFileLink] = useState("");
  const [fileLabel, setFileLabel] = useState("");
  const [dueDate, setDueDate]   = useState("");
  const [cc1, setCc1]           = useState("");
  const [cc2, setCc2]           = useState("");

  // Step 1
  const [tradeFilter, setTradeFilter] = useState("All");
  const [selected, setSelected]       = useState({});

  // Step 2
  const [emails, setEmails]   = useState([]);
  const [sending, setSending] = useState(false);
  const [results, setResults] = useState(null);

  const ccList = [cc1, cc2].filter(Boolean);
  const trades = ["All", ...Array.from(new Set(subs.map((s) => s.trade).filter(Boolean)))];
  const filteredSubs = tradeFilter === "All" ? subs : subs.filter((s) => s.trade === tradeFilter);
  const selectedCount = Object.values(selected).filter(Boolean).length;

  // Auto sign-in + load subs on mount
  useEffect(() => {
    (async () => {
      setSubsLoading(true);
      setSubsError("");
      try {
        const t = await getAccessToken();
        setToken(t);
        const loaded = await loadSubsFromExcel(t);
        setSubs(loaded);
        const sel = {};
        loaded.forEach((_, i) => (sel[i] = false));
        setSelected(sel);
      } catch (err) {
        setSubsError(err.message);
      } finally {
        setSubsLoading(false);
      }
    })();
  }, []);

  const buildReviewEmails = () =>
    subs
      .filter((_, i) => selected[i])
      .map((sub) => ({
        sub,
        subject: `Pricing Request – ${project}`,
        body: buildEmailBody(sub, project, fileLink, fileLabel, dueDate),
      }));

  const handleSend = async () => {
    setSending(true);
    const res = [];
    for (const em of emails) {
      try {
        await sendEmail(token, ccList, em.sub.email, em.subject, em.body);
        res.push({ name: em.sub.name || em.sub.email, ok: true });
      } catch (err) {
        res.push({ name: em.sub.name || em.sub.email, ok: false, error: err.message });
      }
    }
    setResults(res);
    setSending(false);
  };

  const reset = () => {
    setStep(0); setResults(null); setProject(""); setFileLink(""); setFileLabel("");
    setDueDate(""); setCc1(""); setCc2(""); setEmails([]);
    const sel = {}; subs.forEach((_, i) => (sel[i] = false)); setSelected(sel);
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#f5f0e8",
      backgroundImage: "radial-gradient(circle at 80% 10%, #e8dfc8 0%, transparent 50%), radial-gradient(circle at 10% 90%, #d8e8d8 0%, transparent 40%)",
      display: "flex", alignItems: "flex-start", justifyContent: "center",
      padding: "40px 20px", fontFamily: "'Lora', serif",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500;600&family=Lora:ital,wght@0,400;0,600;1,400&display=swap');
        * { box-sizing: border-box; }
        *:focus { outline: none; }
        input:focus, textarea:focus { border-color: #1a1a2e !important; box-shadow: 0 0 0 3px rgba(26,26,46,.08); }
        .sub-row:hover { background: #f0ece4 !important; }
        .chip:hover { background: #1a1a2e !important; color: #f5f0e8 !important; }
        .bp:hover:not([disabled]) { background: #2d2d4e !important; transform: translateY(-1px); }
        .bg:hover { background: #e8e4dc !important; }
        .bs:hover:not([disabled]) { background: #145214 !important; transform: translateY(-1px); }
      `}</style>

      <div style={{ width: "100%", maxWidth: 700 }}>
        {/* Header */}
        <div style={{ marginBottom: 28 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4 }}>
            <div style={{ width: 38, height: 38, background: "#1a1a2e", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>📬</div>
            <h1 style={{ margin: 0, fontSize: 22, fontWeight: 600, color: "#1a1a2e", fontFamily: "'DM Mono', monospace" }}>Sub Pricing Requests</h1>
          </div>
          <p style={{ margin: 0, color: "#999", fontSize: 13, paddingLeft: 50 }}>Sunbelt Utilities Corp · Estimating</p>
        </div>

        <div style={{ background: "#fff", borderRadius: 16, padding: "36px 40px", boxShadow: "0 4px 40px rgba(0,0,0,.08)" }}>

          {/* Loading */}
          {subsLoading && (
            <div style={{ textAlign: "center", padding: "48px 0", color: "#999" }}>
              <div style={{ fontSize: 30, marginBottom: 14 }}>⏳</div>
              <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 13, margin: 0 }}>
                Signing in & loading subcontractors from Job Tracker.xlsx…
              </p>
            </div>
          )}

          {/* Error */}
          {!subsLoading && subsError && (
            <div style={{ background: "#fff5f5", border: "1px solid #f5c8c8", borderRadius: 10, padding: "22px 24px" }}>
              <p style={{ margin: "0 0 6px", fontWeight: 600, color: "#8b1a1a", fontSize: 15 }}>⚠ Could not load subcontractors</p>
              <p style={{ margin: "0 0 18px", fontSize: 13, color: "#c0392b", fontFamily: "'DM Mono', monospace" }}>{subsError}</p>
              <button className="bp" onClick={() => window.location.reload()} style={btnPrimary(false)}>Retry</button>
            </div>
          )}

          {!subsLoading && !subsError && (
            <>
              <StepBar step={step} />

              {/* ── STEP 0 ── */}
              {step === 0 && (
                <div>
                  <h2 style={{ margin: "0 0 22px", fontSize: 17, fontWeight: 600, color: "#1a1a2e" }}>Project Details</h2>

                  <Field label="Project Name">
                    <input style={inputStyle} value={project} onChange={(e) => setProject(e.target.value)} placeholder="e.g. Wilgrove Subdivision" />
                  </Field>
                  <Field label="Plans Link (URL)">
                    <input style={inputStyle} value={fileLink} onChange={(e) => setFileLink(e.target.value)} placeholder="https://..." />
                  </Field>
                  <Field label="Link Label" hint='How it appears in the email, e.g. "Wilgrove Subdivision – Civils & CAD"'>
                    <input style={inputStyle} value={fileLabel} onChange={(e) => setFileLabel(e.target.value)} placeholder="Project name + file type" />
                  </Field>
                  <Field label="Quote Due Date">
                    <input type="date" style={inputStyle} value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
                  </Field>

                  <div style={{ borderTop: "1px solid #eee", paddingTop: 20, marginTop: 6 }}>
                    <p style={{ margin: "0 0 14px", fontFamily: "'DM Mono', monospace", fontSize: 11, color: "#888", textTransform: "uppercase", letterSpacing: ".06em" }}>
                      CC Recipients — sends from your personal M365 account
                    </p>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                      <Field label="CC Email 1">
                        <input style={inputStyle} value={cc1} onChange={(e) => setCc1(e.target.value)} placeholder="you@sunbeltutilities.com" />
                      </Field>
                      <Field label="CC Email 2">
                        <input style={inputStyle} value={cc2} onChange={(e) => setCc2(e.target.value)} placeholder="colleague@sunbeltutilities.com" />
                      </Field>
                    </div>
                  </div>

                  <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
                    <button className="bp" disabled={!project || !fileLink} onClick={() => setStep(1)} style={btnPrimary(!project || !fileLink)}>
                      Next → Select Subs
                    </button>
                  </div>
                </div>
              )}

              {/* ── STEP 1 ── */}
              {step === 1 && (
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
                    <div>
                      <h2 style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 600, color: "#1a1a2e" }}>Select Recipients</h2>
                      <p style={{ margin: 0, fontSize: 13, color: "#999" }}>
                        {selectedCount} selected · {subs.length} subs loaded from Job Tracker.xlsx
                      </p>
                    </div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button onClick={() => {
                        const s = { ...selected };
                        filteredSubs.forEach((sub) => (s[subs.indexOf(sub)] = true));
                        setSelected(s);
                      }} style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", padding: "5px 10px", borderRadius: 6, border: "1px solid #ddd", background: "#fff", cursor: "pointer", color: "#555" }}>
                        Select Visible
                      </button>
                      <button onClick={() => setSelected({})}
                        style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", padding: "5px 10px", borderRadius: 6, border: "1px solid #ddd", background: "#fff", cursor: "pointer", color: "#555" }}>
                        Clear All
                      </button>
                    </div>
                  </div>

                  {/* Trade chips */}
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
                    {trades.map((t) => (
                      <button key={t} className="chip" onClick={() => setTradeFilter(t)} style={{
                        padding: "5px 14px", borderRadius: 20, border: "1.5px solid #ddd",
                        fontFamily: "'DM Mono', monospace", fontSize: 11, cursor: "pointer", transition: "all .15s",
                        background: tradeFilter === t ? "#1a1a2e" : "#fdfcf9",
                        color: tradeFilter === t ? "#f5f0e8" : "#555",
                      }}>{t}</button>
                    ))}
                  </div>

                  {/* Sub list */}
                  <div style={{ border: "1.5px solid #eee", borderRadius: 10, overflow: "hidden", marginBottom: 22, maxHeight: 360, overflowY: "auto" }}>
                    {filteredSubs.length === 0 && (
                      <div style={{ padding: 24, textAlign: "center", color: "#bbb", fontSize: 13 }}>No subs match this trade filter.</div>
                    )}
                    {filteredSubs.map((sub) => {
                      const gi = subs.indexOf(sub);
                      return (
                        <div key={gi} className="sub-row" onClick={() => setSelected((s) => ({ ...s, [gi]: !s[gi] }))}
                          style={{ display: "flex", alignItems: "center", padding: "11px 16px", borderBottom: "1px solid #f0ece4", cursor: "pointer", transition: "background .12s", background: selected[gi] ? "#f5fdf5" : "#fff" }}>
                          <div style={{ width: 19, height: 19, borderRadius: 5, marginRight: 14, flexShrink: 0, border: `2px solid ${selected[gi] ? "#1a6b1a" : "#ccc"}`, background: selected[gi] ? "#c8f5c8" : "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11 }}>
                            {selected[gi] ? "✓" : ""}
                          </div>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600, fontSize: 14, color: "#1a1a2e" }}>{sub.name || "(no name)"}</div>
                            <div style={{ fontSize: 12, color: "#999" }}>{[sub.company, sub.email].filter(Boolean).join(" · ")}</div>
                          </div>
                          {sub.trade && (
                            <span style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", padding: "3px 10px", borderRadius: 12, background: "#e8e4dc", color: "#666", flexShrink: 0 }}>
                              {sub.trade}
                            </span>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <button className="bg" onClick={() => setStep(0)} style={btnGhost}>← Back</button>
                    <button className="bp" disabled={!selectedCount} onClick={() => { setEmails(buildReviewEmails()); setStep(2); }} style={btnPrimary(!selectedCount)}>
                      Review {selectedCount} Email{selectedCount !== 1 ? "s" : ""} →
                    </button>
                  </div>
                </div>
              )}

              {/* ── STEP 2: Review ── */}
              {step === 2 && !results && (
                <div>
                  <h2 style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 600, color: "#1a1a2e" }}>Review & Send</h2>
                  <p style={{ margin: "0 0 20px", fontSize: 13, color: "#999" }}>
                    {emails.length} individual email{emails.length !== 1 ? "s" : ""}, each sent separately · CC: {ccList.join(", ") || "none"}
                  </p>

                  <div style={{ display: "flex", flexDirection: "column", gap: 14, maxHeight: 440, overflowY: "auto", paddingRight: 2, marginBottom: 20 }}>
                    {emails.map((em, i) => (
                      <div key={i} style={{ border: "1.5px solid #e8e4dc", borderRadius: 10, overflow: "hidden" }}>
                        <div style={{ background: "#f8f5f0", padding: "9px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #e8e4dc" }}>
                          <div>
                            <span style={{ fontWeight: 600, fontSize: 13, color: "#1a1a2e" }}>{em.sub.name || em.sub.email}</span>
                            <span style={{ fontSize: 12, color: "#bbb", marginLeft: 8 }}>{em.sub.email}</span>
                            {em.sub.trade && <span style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", marginLeft: 8, color: "#aaa" }}>{em.sub.trade}</span>}
                          </div>
                          <span style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", color: "#aaa", flexShrink: 0 }}>Subj: {em.subject}</span>
                        </div>
                        <textarea
                          value={em.body}
                          onChange={(e) => setEmails((prev) => prev.map((x, j) => j === i ? { ...x, body: e.target.value } : x))}
                          style={{ width: "100%", border: "none", padding: "13px 16px", fontFamily: "'Lora', serif", fontSize: 13, color: "#333", resize: "vertical", minHeight: 190, background: "#fff", lineHeight: 1.65, display: "block" }}
                        />
                      </div>
                    ))}
                  </div>

                  <div style={{ background: "#f0fdf0", border: "1px solid #c8f5c8", borderRadius: 8, padding: "11px 16px", marginBottom: 18, fontSize: 13, color: "#1a5c1a" }}>
                    ✉ <strong>{emails.length} email{emails.length !== 1 ? "s" : ""}</strong> will be sent individually from your Microsoft 365 account.
                    {ccList.length > 0 && <> CC'd to: {ccList.join(" & ")}.</>}
                  </div>

                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <button className="bg" onClick={() => setStep(1)} style={btnGhost}>← Back</button>
                    <button className="bs" disabled={sending} onClick={handleSend} style={btnGreen(sending)}>
                      {sending ? "Sending…" : `✉ Send ${emails.length} Email${emails.length !== 1 ? "s" : ""}`}
                    </button>
                  </div>
                </div>
              )}

              {/* ── RESULTS ── */}
              {step === 2 && results && (
                <div>
                  <h2 style={{ margin: "0 0 20px", fontSize: 17, fontWeight: 600, color: "#1a1a2e" }}>
                    {results.every((r) => r.ok) ? "✅ All Emails Sent!" : "⚠ Completed with Errors"}
                  </h2>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 28 }}>
                    {results.map((r, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", borderRadius: 8, border: `1px solid ${r.ok ? "#c8f5c8" : "#f5c8c8"}`, background: r.ok ? "#f0fdf0" : "#fff5f5" }}>
                        <span style={{ fontSize: 16 }}>{r.ok ? "✓" : "✗"}</span>
                        <div>
                          <div style={{ fontWeight: 600, fontSize: 13, color: r.ok ? "#1a5c1a" : "#8b1a1a" }}>{r.name}</div>
                          {!r.ok && <div style={{ fontSize: 12, color: "#c0392b", fontFamily: "'DM Mono', monospace" }}>{r.error}</div>}
                        </div>
                      </div>
                    ))}
                  </div>
                  <button className="bp" onClick={reset} style={btnPrimary(false)}>Start New Request</button>
                </div>
              )}
            </>
          )}
        </div>

        <p style={{ textAlign: "center", marginTop: 18, fontSize: 11, color: "#ccc", fontFamily: "'DM Mono', monospace" }}>
          Subs auto-loaded from Job Tracker.xlsx · Sunbelt Utilities Corp
        </p>
      </div>
    </div>
  );
}

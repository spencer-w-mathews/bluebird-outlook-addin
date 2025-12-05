import React, { useState } from "react";

const API_BASE = "https://api.bluebird.ai";

export default function BluebirdPane() {
  const [tone, setTone] = useState("default");
  const [action, setAction] = useState("rewrite");
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);
  const [lastDraft, setLastDraft] = useState(null); // { originalHtml, rewrittenHtml, tone, action }

  async function getBodyHtml() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("html", (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error);
        }
      });
    });
  }

  async function setBodyHtml(html) {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(
        html,
        { coercionType: "html" },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(result.error);
          }
        }
      );
    });
  }

  async function handleRewrite() {
    try {
      setLoading(true);
      setStatus("Fetching draft from Bluebird…");

      const originalHtml = await getBodyHtml();

      const res = await fetch(`${API_BASE}/v1/rewrite`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          html: originalHtml,
          tone,
          action, // "rewrite" | "shorter" | "longer" | "more_formal" | ...
        }),
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const data = await res.json();
      const rewrittenHtml =
        data.rewrittenHtml || data.rewritten || originalHtml;

      await setBodyHtml(rewrittenHtml);

      setLastDraft({ originalHtml, rewrittenHtml, tone, action });
      setStatus("Draft updated by Bluebird.");
    } catch (err) {
      console.error("Bluebird rewrite error", err);
      setStatus("Error rewriting draft. Please try again.");
    } finally {
      setLoading(false);
    }
  }

  async function sendFeedback(vote) {
    if (!lastDraft) return;
    try {
      setStatus("Sending feedback…");

      await fetch(`${API_BASE}/v1/feedback`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          vote, // "up" | "down"
          ...lastDraft,
        }),
      });

      setStatus(
        vote === "up"
          ? "Thanks for the thumbs up 💙"
          : "Thanks, we’ll use this to improve Bluebird."
      );
    } catch (err) {
      console.error("Bluebird feedback error", err);
      setStatus("Could not send feedback.");
    }
  }

  return (
    <div
      style={{
        fontFamily:
          'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
        padding: "12px",
        fontSize: "13px",
      }}
    >
      <h3 style={{ margin: "0 0 8px", color: "#1A73E8" }}>Bluebird</h3>
      <p style={{ margin: "0 0 12px", color: "#5f6368" }}>
        Rewrite your draft, adjust tone, and tell us if Bluebird got it right.
      </p>

      <div style={{ marginBottom: "10px" }}>
        <label
          style={{
            fontSize: "11px",
            textTransform: "uppercase",
            letterSpacing: "0.03em",
            color: "#5f6368",
          }}
        >
          Tone
        </label>
        <select
          value={tone}
          onChange={(e) => setTone(e.target.value)}
          style={{
            width: "100%",
            padding: "4px 6px",
            borderRadius: "6px",
            border: "1px solid #dadce0",
            fontSize: "12px",
            marginTop: "4px",
          }}
        >
          <option value="default">My default tone</option>
          <option value="more_formal">More formal</option>
          <option value="more_casual">More casual</option>
          <option value="more_direct">More direct</option>
          <option value="more_warm">More warm</option>
        </select>
      </div>

      <div style={{ marginBottom: "10px" }}>
        <div
          style={{
            fontSize: "11px",
            textTransform: "uppercase",
            letterSpacing: "0.03em",
            color: "#5f6368",
            marginBottom: "4px",
          }}
        >
          Quick actions
        </div>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "4px",
            marginTop: 4,
          }}
        >
          {[
            { id: "rewrite", label: "Smart rewrite" },
            { id: "shorter", label: "Shorter" },
            { id: "longer", label: "Longer" },
            { id: "fix_grammar", label: "Fix grammar" },
            { id: "summarize", label: "Summarize" },
          ].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setAction(opt.id)}
              style={{
                flex: "1 1 48%",
                borderRadius: 999,
                padding: "4px 8px",
                fontSize: "11px",
                border:
                  action === opt.id ? "1px solid #1A73E8" : "1px solid #dadce0",
                background: action === opt.id ? "#E8F0FE" : "#f8f9fa",
                color: action === opt.id ? "#1A73E8" : "#202124",
                cursor: "pointer",
              }}
            >
              {opt.label}
            </button>
          ))}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginTop: "8px",
        }}
      >
        <button
          onClick={handleRewrite}
          disabled={loading}
          style={{
            borderRadius: 999,
            padding: "6px 14px",
            border: "none",
            fontSize: "12px",
            fontWeight: 500,
            cursor: loading ? "default" : "pointer",
            background: "#1A73E8",
            color: "white",
            opacity: loading ? 0.7 : 1,
          }}
        >
          {loading ? "Rewriting…" : "Rewrite with Bluebird"}
        </button>

        <div>
          <button
            onClick={() => sendFeedback("up")}
            disabled={!lastDraft}
            style={{
              border: "none",
              background: "transparent",
              fontSize: "18px",
              cursor: lastDraft ? "pointer" : "default",
              opacity: lastDraft ? 1 : 0.3,
              marginRight: 4,
            }}
          >
            👍
          </button>
          <button
            onClick={() => sendFeedback("down")}
            disabled={!lastDraft}
            style={{
              border: "none",
              background: "transparent",
              fontSize: "18px",
              cursor: lastDraft ? "pointer" : "default",
              opacity: lastDraft ? 1 : 0.3,
            }}
          >
            👎
          </button>
        </div>
      </div>

      {status && (
        <div
          style={{
            marginTop: "8px",
            fontSize: "11px",
            color: "#5f6368",
          }}
        >
          {status}
        </div>
      )}
    </div>
  );
}

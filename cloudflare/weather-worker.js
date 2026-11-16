// IRIS Road Weather dashboard — Cloudflare Worker
// Route: mindfreedom.co.tz/weather*
//
// Serves the dashboard and daily advisories from a private Workers KV namespace.
// Cloudflare Access sits in front of this route and admits only the advisory
// recipients. As a second, independent check the Worker verifies the signed
// Access token itself, so a mis-configured Access policy fails closed instead
// of exposing the advisories.
//
// Bindings / variables (set in the Worker's Settings):
//   WEATHER      KV namespace binding holding index.html, advisories/index.json, advisories/<island>/<date>.html
//   TEAM_DOMAIN  e.g. "https://<team>.cloudflareaccess.com"
//   POLICY_AUD   Access application "Application Audience (AUD) Tag"

const PREFIX = "/weather/";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/weather") {
      return Response.redirect(url.origin + PREFIX, 301);
    }
    if (!url.pathname.startsWith(PREFIX)) {
      return fetch(request); // not ours: pass through to the main site
    }

    const user = await verifyAccess(request, env);
    if (!user) {
      return new Response("Access denied. Sign in through the advisory link.", {
        status: 403,
        headers: { "content-type": "text/plain; charset=utf-8", "cache-control": "no-store" },
      });
    }

    let key = decodeURIComponent(url.pathname.slice(PREFIX.length)) || "index.html";
    if (key.endsWith("/")) key += "index.html";
    if (key.includes("..") || key.startsWith("/")) {
      return new Response("Bad request", { status: 400 });
    }

    const body = await env.WEATHER.get(key, "arrayBuffer");
    if (body === null) return new Response("Not found", { status: 404 });

    const type = key.endsWith(".json") ? "application/json; charset=utf-8"
      : key.endsWith(".html") ? "text/html; charset=utf-8"
      : "application/octet-stream";

    return new Response(body, {
      headers: {
        "content-type": type,
        "cache-control": "private, no-store",
        "x-robots-tag": "noindex, nofollow",
        "x-content-type-options": "nosniff",
        "referrer-policy": "same-origin",
      },
    });
  },
};

// ── Cloudflare Access JWT verification (RS256) ───────────────────────────────
let certCache = { at: 0, keys: [] };

async function verifyAccess(request, env) {
  if (!env.TEAM_DOMAIN || !env.POLICY_AUD) return null; // fail closed
  const token = request.headers.get("cf-access-jwt-assertion");
  if (!token) return null;
  const parts = token.split(".");
  if (parts.length !== 3) return null;

  let header, payload;
  try {
    header = JSON.parse(b64urlText(parts[0]));
    payload = JSON.parse(b64urlText(parts[1]));
  } catch {
    return null;
  }

  const now = Math.floor(Date.now() / 1000);
  const aud = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (!aud.includes(env.POLICY_AUD)) return null;
  if (typeof payload.exp !== "number" || payload.exp < now) return null;
  if (payload.iss && payload.iss !== env.TEAM_DOMAIN) return null;

  const jwk = (await getKeys(env.TEAM_DOMAIN)).find((k) => k.kid === header.kid);
  if (!jwk) return null;

  const key = await crypto.subtle.importKey(
    "jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
  const ok = await crypto.subtle.verify(
    "RSASSA-PKCS1-v1_5", key, b64urlBytes(parts[2]),
    new TextEncoder().encode(parts[0] + "." + parts[1]));
  return ok ? (payload.email || "user") : null;
}

async function getKeys(team) {
  if (Date.now() - certCache.at < 3600_000 && certCache.keys.length) return certCache.keys;
  const res = await fetch(team + "/cdn-cgi/access/certs");
  if (!res.ok) return certCache.keys;
  const data = await res.json();
  certCache = { at: Date.now(), keys: data.keys || [] };
  return certCache.keys;
}

function b64urlBytes(s) {
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function b64urlText(s) {
  return new TextDecoder().decode(b64urlBytes(s));
}

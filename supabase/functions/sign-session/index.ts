// ═══════════════════════════════════════════════════════════════════════
// One Dream Community · client e-sign — Supabase Edge Function
//
// Actions (all POST):
//   {action:"create", pin, page, coordinatorEmail, participantName, participantEmail, payload}
//        → {ok, id, expiresAt}                 (coordinator, PIN-protected)
//   {action:"get", id}
//        → {ok, page, status, participant:{name,email}, payload}   (participant's browser)
//   multipart/form-data  action=submit id=… page=… snapshot=… tz=… file0..fileN
//        → {ok, signedAt, emailedTo, files}    (participant's browser)
//
// On submit the signed PDFs are uploaded to SharePoint and e-mailed from the
// team mailbox, both through Microsoft Graph with application permissions
// (client-credentials). Nothing but metadata stays in Supabase.
//
// Deploy:  supabase functions deploy sign-session --no-verify-jwt
// Secrets: see README.md
// ═══════════════════════════════════════════════════════════════════════
import { createClient } from "npm:@supabase/supabase-js@2";
import { encodeBase64 } from "jsr:@std/encoding@1/base64";
import { encodeHex } from "jsr:@std/encoding@1/hex";

const env = (k: string, d = "") => Deno.env.get(k) ?? d;
const CFG = {
  supabaseUrl: env("SUPABASE_URL"),
  serviceKey: env("SUPABASE_SERVICE_ROLE_KEY"),
  tenant: env("MS_TENANT_ID"),
  clientId: env("MS_CLIENT_ID"),
  clientSecret: env("MS_CLIENT_SECRET"),
  spHost: env("SP_HOST", "onedreamcommunity.sharepoint.com"),
  spSitePath: env("SP_SITE_PATH", "/sites/AgedCareTeam622"),
  spFolder: env("SP_FOLDER", "Service Agreement Records/Client Signed"),
  mailFrom: env("MAIL_FROM", "agedcare@onedreamcommunity.com.au"),
  mailTeam: env("MAIL_TEAM", "agedcare@onedreamcommunity.com.au").split(",").map((s) => s.trim()).filter(Boolean),
  pin: env("COORDINATOR_PIN"),
  coordDomain: env("COORDINATOR_DOMAIN", "onedreamcommunity.com.au").toLowerCase(),
  origins: env("ALLOWED_ORIGINS", "https://onedream-hy.github.io").split(",").map((s) => s.trim()).filter(Boolean),
  linkDays: Number(env("LINK_DAYS", "30")),
  maxFileMb: Number(env("MAX_FILE_MB", "25")),
  maxFiles: 8,
  createPerHour: 40,
  officePhone: env("OFFICE_PHONE", "02 9161 7567"),
  orgName: "One Dream Community",
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const CHUNK = 3_276_800; // 10 × 320 KiB — Graph upload sessions want multiples of 320 KiB

const db = () => createClient(CFG.supabaseUrl, CFG.serviceKey, { auth: { persistSession: false } });

// ── CORS ──────────────────────────────────────────────────────────────────
function cors(req: Request): Record<string, string> {
  const origin = req.headers.get("origin") ?? "";
  const ok = CFG.origins.includes("*") || CFG.origins.includes(origin) || /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin);
  return {
    "Access-Control-Allow-Origin": ok ? origin : CFG.origins[0] ?? "",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}
const json = (req: Request, body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...cors(req), "Content-Type": "application/json; charset=utf-8" } });
class HttpError extends Error { constructor(public status: number, msg: string) { super(msg); } }
const bad = (status: number, msg: string) => new HttpError(status, msg);

const ipOf = (req: Request) => (req.headers.get("x-forwarded-for") ?? "").split(",")[0].trim() || req.headers.get("cf-connecting-ip") || "";
const isEmail = (s: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
const isUuid = (s: string) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
const esc = (s: unknown) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string));
const safeName = (s: string) => s.replace(/[\\/:*?"<>|#%&{}~]/g, "").replace(/\s+/g, " ").trim().slice(0, 80) || "Client";
function timingSafeEqual(a: string, b: string) {
  const A = new TextEncoder().encode(a), B = new TextEncoder().encode(b);
  let r = A.length ^ B.length;
  for (let i = 0; i < Math.max(A.length, B.length); i++) r |= (A[i] ?? 0) ^ (B[i] ?? 0);
  return r === 0;
}

// ── Microsoft Graph (client credentials) ──────────────────────────────────
let tokenCache: { token: string; exp: number } | null = null;
async function graphToken(): Promise<string> {
  if (!CFG.tenant || !CFG.clientId || !CFG.clientSecret) throw bad(500, "Microsoft Graph credentials are not configured (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).");
  if (tokenCache && tokenCache.exp > Date.now() + 60_000) return tokenCache.token;
  const res = await fetch(`https://login.microsoftonline.com/${CFG.tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, client_secret: CFG.clientSecret, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }),
  });
  const data = await res.json();
  if (!res.ok || !data.access_token) throw bad(502, "Microsoft sign-in failed: " + (data.error_description ?? res.status));
  tokenCache = { token: data.access_token, exp: Date.now() + (data.expires_in ?? 3600) * 1000 };
  return data.access_token;
}
async function graph(path: string, init: RequestInit = {}, raw = false): Promise<any> {
  const token = await graphToken();
  const headers = new Headers(init.headers ?? {});
  headers.set("Authorization", "Bearer " + token);
  if (!headers.has("Content-Type") && init.body && typeof init.body === "string") headers.set("Content-Type", "application/json");
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, { ...init, headers });
  if (raw) return res;
  if (res.status === 204) return null;
  const text = await res.text();
  let data: any = null;
  try { data = text ? JSON.parse(text) : null; } catch { data = text; }
  if (!res.ok) {
    const msg = data?.error?.message ?? (typeof data === "string" ? data.slice(0, 200) : res.statusText);
    throw bad(res.status === 404 ? 404 : 502, `Graph ${res.status}: ${msg}`);
  }
  return data;
}
const encPath = (p: string) => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");

let driveCache: string | null = null;
async function driveId(): Promise<string> {
  if (driveCache) return driveCache;
  const site = await graph(`/sites/${CFG.spHost}:${CFG.spSitePath}`);
  const drive = await graph(`/sites/${site.id}/drive`);
  driveCache = drive.id;
  return drive.id;
}
async function ensureFolder(drive: string, path: string) {
  const parts = path.split("/").filter(Boolean);
  let cur = "";
  for (const p of parts) {
    const next = cur ? cur + "/" + p : p;
    try { await graph(`/drives/${drive}/root:/${encPath(next)}`); }
    catch (e) {
      if (!(e instanceof HttpError) || e.status !== 404) throw e;
      const parent = cur ? `/drives/${drive}/root:/${encPath(cur)}:/children` : `/drives/${drive}/root/children`;
      try { await graph(parent, { method: "POST", body: JSON.stringify({ name: p, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }) }); }
      catch (e2) { if (!(e2 instanceof HttpError) || !/409|nameAlreadyExists/i.test(e2.message)) throw e2; }
    }
    cur = next;
  }
}
async function putChunks(uploadUrl: string, bytes: Uint8Array): Promise<any> {
  let last: any = null;
  for (let start = 0; start < bytes.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.length);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${bytes.length}` },
      body: bytes.slice(start, end),
    });
    const text = await res.text();
    if (!res.ok) throw bad(502, `Upload chunk failed (${res.status}): ${text.slice(0, 200)}`);
    try { last = text ? JSON.parse(text) : null; } catch { last = null; }
  }
  return last;
}
async function uploadToSharePoint(folder: string, name: string, bytes: Uint8Array) {
  const drive = await driveId();
  await ensureFolder(drive, folder);
  const session = await graph(`/drives/${drive}/root:/${encPath(folder + "/" + name)}:/createUploadSession`, {
    method: "POST",
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename", name } }),
  });
  const item = await putChunks(session.uploadUrl, bytes);
  return { id: item?.id ?? null, name: item?.name ?? name, webUrl: item?.webUrl ?? null };
}

// Draft → attachments (small inline, large via upload session) → send.
async function sendMailWithAttachments(opts: { to: string[]; cc: string[]; subject: string; html: string; files: { name: string; bytes: Uint8Array }[] }) {
  const u = `/users/${encodeURIComponent(CFG.mailFrom)}`;
  const msg = await graph(`${u}/messages`, {
    method: "POST",
    body: JSON.stringify({
      subject: opts.subject,
      body: { contentType: "HTML", content: opts.html },
      toRecipients: opts.to.map((a) => ({ emailAddress: { address: a } })),
      ccRecipients: opts.cc.map((a) => ({ emailAddress: { address: a } })),
    }),
  });
  for (const f of opts.files) {
    if (f.bytes.length < 3_000_000) {
      await graph(`${u}/messages/${msg.id}/attachments`, {
        method: "POST",
        body: JSON.stringify({ "@odata.type": "#microsoft.graph.fileAttachment", name: f.name, contentType: "application/pdf", contentBytes: encodeBase64(f.bytes) }),
      });
    } else {
      const s = await graph(`${u}/messages/${msg.id}/attachments/createUploadSession`, {
        method: "POST",
        body: JSON.stringify({ AttachmentItem: { attachmentType: "file", name: f.name, contentType: "application/pdf", size: f.bytes.length } }),
      });
      await putChunks(s.uploadUrl, f.bytes);
    }
  }
  await graph(`${u}/messages/${msg.id}/send`, { method: "POST" });
  return { messageId: msg.id as string };
}

// ── Handlers ──────────────────────────────────────────────────────────────
async function handleCreate(req: Request, b: any) {
  if (!CFG.pin) throw bad(500, "COORDINATOR_PIN is not set on the server.");
  if (typeof b.pin !== "string" || !timingSafeEqual(b.pin, CFG.pin)) throw bad(403, "Incorrect coordinator PIN.");
  const page = String(b.page ?? "");
  if (!["contract", "onboarding"].includes(page)) throw bad(400, "Unknown page.");
  const coord = String(b.coordinatorEmail ?? "").trim().toLowerCase();
  if (!isEmail(coord) || (CFG.coordDomain && !coord.endsWith("@" + CFG.coordDomain))) throw bad(400, `Coordinator e-mail must be an @${CFG.coordDomain} address.`);
  const pName = String(b.participantName ?? "").trim();
  if (!pName) throw bad(400, "Participant name is required.");
  const pEmail = String(b.participantEmail ?? "").trim().toLowerCase();
  if (pEmail && !isEmail(pEmail)) throw bad(400, "Participant e-mail is not valid.");
  if (!b.payload || typeof b.payload !== "object") throw bad(400, "Missing form payload.");
  const size = JSON.stringify(b.payload).length;
  if (size > 6_000_000) throw bad(413, "The form payload is too large.");

  const ip = ipOf(req);
  const sb = db();
  if (ip) {
    const since = new Date(Date.now() - 3600_000).toISOString();
    const { count } = await sb.from("sign_sessions").select("id", { count: "exact", head: true }).eq("created_ip", ip).gte("created_at", since);
    if ((count ?? 0) >= CFG.createPerHour) throw bad(429, "Too many links created from this network in the last hour. Please try again later.");
  }
  const expiresAt = new Date(Date.now() + CFG.linkDays * 86400_000).toISOString();
  const { data, error } = await sb.from("sign_sessions").insert({
    page, coordinator_email: coord, participant_name: pName, participant_email: pEmail || null,
    payload: b.payload, expires_at: expiresAt, created_ip: ip || null,
  }).select("id").single();
  if (error) throw bad(500, "Could not save the session: " + error.message);
  return { ok: true, id: data.id, expiresAt };
}

async function handleGet(req: Request, b: any) {
  const id = String(b.id ?? "").toLowerCase();
  if (!isUuid(id)) throw bad(400, "Invalid link.");
  const sb = db();
  const { data: s, error } = await sb.from("sign_sessions").select("id, page, status, expires_at, participant_name, participant_email, payload, signed_at, email_result").eq("id", id).maybeSingle();
  if (error) throw bad(500, error.message);
  if (!s) throw bad(404, "This link is not valid.");
  const participant = { name: s.participant_name, email: s.participant_email };
  if (s.status === "signed") return { ok: true, page: s.page, status: "signed", participant, signedAt: s.signed_at, emailedTo: s.email_result?.to ?? [] };
  if (new Date(s.expires_at).getTime() < Date.now()) return { ok: true, page: s.page, status: "expired", participant };
  if (s.status === "sent") {
    await sb.from("sign_sessions").update({ status: "opened", opened_at: new Date().toISOString(), opened_ip: ipOf(req) || null }).eq("id", id).eq("status", "sent");
  }
  if (s.status === "submitting") {
    // a submit is in flight — or crashed. Allow a retry once it is older than 5 minutes.
    const stale = !s.signed_at || new Date(s.signed_at).getTime() < Date.now() - 5 * 60_000;
    if (!stale) return { ok: true, page: s.page, status: "submitting", participant };
    await sb.from("sign_sessions").update({ status: "opened" }).eq("id", id).eq("status", "submitting");
  }
  return { ok: true, page: s.page, status: "opened", participant, payload: s.payload };
}

async function handleSubmit(req: Request, form: FormData) {
  const id = String(form.get("id") ?? "").toLowerCase();
  if (!isUuid(id)) throw bad(400, "Invalid link.");
  const page = String(form.get("page") ?? "");
  const tz = String(form.get("tz") ?? "");
  let snapshot: unknown = null;
  try { snapshot = JSON.parse(String(form.get("snapshot") ?? "null")); } catch { snapshot = null; }

  const files: { name: string; bytes: Uint8Array; sha256: string }[] = [];
  for (const [k, v] of form.entries()) {
    if (!/^file\d+$/.test(k) || !(v instanceof File)) continue;
    if (files.length >= CFG.maxFiles) throw bad(413, "Too many files.");
    const bytes = new Uint8Array(await v.arrayBuffer());
    if (bytes.length > CFG.maxFileMb * 1024 * 1024) throw bad(413, `${v.name} is larger than ${CFG.maxFileMb} MB.`);
    if (bytes.length < 100 || String.fromCharCode(...bytes.slice(0, 4)) !== "%PDF") throw bad(400, `${v.name} is not a PDF.`);
    const sha256 = encodeHex(await crypto.subtle.digest("SHA-256", bytes));
    files.push({ name: safeName(v.name.replace(/\.pdf$/i, "")) + ".pdf", bytes, sha256 });
  }
  if (!files.length) throw bad(400, "No PDF received.");

  const sb = db();
  const { data: s, error } = await sb.from("sign_sessions").select("*").eq("id", id).maybeSingle();
  if (error) throw bad(500, error.message);
  if (!s) throw bad(404, "This link is not valid.");
  if (s.status === "signed") throw bad(409, "These documents have already been submitted.");
  if (new Date(s.expires_at).getTime() < Date.now()) throw bad(410, "This link has expired. Please ask One Dream Community for a new one.");
  if (page && page !== s.page) throw bad(400, "Wrong document type for this link.");

  // claim the session (prevents a double submit)
  const now = new Date();
  const { data: claimed } = await sb.from("sign_sessions").update({ status: "submitting", signed_at: now.toISOString() })
    .eq("id", id).in("status", ["sent", "opened"]).select("id");
  if (!claimed || !claimed.length) throw bad(409, "This submission is already being processed. Please wait a minute and refresh.");

  const docLabel = s.page === "contract" ? "Service Agreement" : "Onboarding Forms";
  const whenSydney = now.toLocaleString("en-AU", { timeZone: "Australia/Sydney", dateStyle: "long", timeStyle: "short" });
  const stamp = now.toLocaleDateString("en-CA", { timeZone: "Australia/Sydney" }); // yyyy-mm-dd, Sydney date
  try {
    // 1. SharePoint
    const folder = `${CFG.spFolder}/${safeName(s.participant_name)}`;
    const uploaded: { name: string; size: number; sha256: string; webUrl: string | null }[] = [];
    for (const f of files) {
      const name = f.name.replace(/\.pdf$/i, "") + `_signed_${stamp}.pdf`;
      const item = await uploadToSharePoint(folder, name, f.bytes);
      uploaded.push({ name: item.name, size: f.bytes.length, sha256: f.sha256, webUrl: item.webUrl });
    }
    // 2. E-mail
    const to: string[] = [];
    if (s.participant_email && isEmail(s.participant_email)) to.push(s.participant_email);
    for (const t of CFG.mailTeam) if (!to.includes(t)) to.push(t);
    const cc: string[] = [];
    if (s.coordinator_email && isEmail(s.coordinator_email) && !to.includes(s.coordinator_email)) cc.push(s.coordinator_email);
    const first = esc(s.participant_name.split(" ")[0]);
    const html = `
      <div style="font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#243024;line-height:1.5;">
        <p>Dear ${first},</p>
        <p>Thank you — your signed <strong>${esc(docLabel)}</strong> with ${esc(CFG.orgName)} has been received.</p>
        <p>Signed copy attached:</p>
        <ul>${uploaded.map((u) => `<li>${esc(u.name)} (${(u.size / 1048576).toFixed(1)} MB)</li>`).join("")}</ul>
        <p style="font-size:13px;color:#555;">Signed electronically on ${esc(whenSydney)} (Sydney time). Reference: ${esc(id)}.</p>
        <p>If you have any questions, please call us on <strong>${esc(CFG.officePhone)}</strong> or reply to this e-mail.</p>
        <p>Kind regards,<br>${esc(CFG.orgName)}<br>${esc(CFG.mailFrom)}</p>
        <hr style="border:none;border-top:1px solid #ddd;margin:18px 0;">
        <p style="font-size:12px;color:#777;">For ODC staff: filed in SharePoint under <em>${esc(folder)}</em>${uploaded[0]?.webUrl ? ` — <a href="${esc(uploaded[0].webUrl)}">open</a>` : ""}. Coordinator: ${esc(s.coordinator_email)}.</p>
      </div>`;
    const mail = await sendMailWithAttachments({
      to, cc, subject: `${CFG.orgName} – signed ${docLabel} – ${s.participant_name}`, html,
      files: uploaded.map((u, i) => ({ name: u.name, bytes: files[i].bytes })),
    });
    // 3. Record
    await sb.from("sign_sessions").update({
      status: "signed", signed_at: now.toISOString(), signer_ip: ipOf(req) || null, signer_ua: req.headers.get("user-agent") ?? null,
      signer_tz: tz || null, signed_payload: snapshot, files: uploaded, email_result: { to, cc, messageId: mail.messageId }, last_error: null,
    }).eq("id", id);
    return { ok: true, signedAt: now.toISOString(), emailedTo: to, cc, files: uploaded.map((u) => ({ name: u.name, webUrl: u.webUrl })) };
  } catch (e) {
    await sb.from("sign_sessions").update({ status: "opened", last_error: String((e as Error).message ?? e).slice(0, 1000) }).eq("id", id);
    throw e;
  }
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(req) });
  if (req.method !== "POST") return json(req, { ok: false, error: "POST only" }, 405);
  try {
    const ct = req.headers.get("content-type") ?? "";
    if (ct.includes("multipart/form-data")) {
      const form = await req.formData();
      if (String(form.get("action")) !== "submit") throw bad(400, "Unknown action.");
      return json(req, await handleSubmit(req, form));
    }
    const b = await req.json().catch(() => ({}));
    switch (b.action) {
      case "create": return json(req, await handleCreate(req, b));
      case "get": return json(req, await handleGet(req, b));
      default: throw bad(400, "Unknown action.");
    }
  } catch (e) {
    const status = e instanceof HttpError ? e.status : 500;
    console.error("[sign-session]", status, (e as Error).message);
    return json(req, { ok: false, error: (e as Error).message ?? String(e) }, status);
  }
});

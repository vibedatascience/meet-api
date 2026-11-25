import { Hono } from "hono";
import { cors } from "hono/cors";
import { serve } from "@hono/node-server";
import { DateTime, Interval } from "luxon";
import crypto from "node:crypto";
import { config } from "./config.js";

const MATON_KEY = process.env.MATON_KEY;
const SECRET = process.env.SIGNING_SECRET || "dev-secret";
const MEETING_LINK = process.env.MEETING_LINK || "";
const SITE_URL = process.env.SITE_URL || "https://rahulch.site";
const GATEWAY = "https://gateway.maton.ai/outlook/";
const GCAL = "https://gateway.maton.ai/google-calendar/calendar/v3/";

// ---------- Google Calendar via Maton (Meet links + extra busy times) ----------
async function gcal(path, opts = {}) {
  const r = await fetch(GCAL + path, {
    ...opts,
    signal: AbortSignal.timeout(8000),
    headers: { Authorization: `Bearer ${MATON_KEY}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (r.status === 204) return null;
  const body = await r.json().catch(() => null);
  if (!r.ok) throw new Error(`gcal ${r.status}: ${body?.error?.message || ""}`);
  return body;
}

const meetKey = (startIso, email) => crypto.createHash("sha256").update(`${startIso}|${email}`).digest("hex").slice(0, 32);

// Placeholder event on Google Calendar that only exists to mint a Meet link.
async function createMeet({ startIso, endIso, title, email }) {
  const ev = await gcal("calendars/primary/events?conferenceDataVersion=1&sendUpdates=none", {
    method: "POST",
    body: JSON.stringify({
      summary: `Meet link: ${title}`,
      description: "Placeholder created by rahulch.site/meet. The real invite is on Outlook.",
      transparency: "transparent",
      start: { dateTime: startIso },
      end: { dateTime: endIso },
      extendedProperties: { private: { meetKey: meetKey(startIso, email) } },
      conferenceData: { createRequest: { requestId: crypto.randomUUID(), conferenceSolutionKey: { type: "hangoutsMeet" } } },
    }),
  });
  return ev.hangoutLink || null;
}

async function deleteMeet(startIso, email) {
  const r = await gcal(`calendars/primary/events?privateExtendedProperty=${encodeURIComponent(`meetKey=${meetKey(startIso, email)}`)}`);
  for (const e of r?.items || []) await gcal(`calendars/primary/events/${e.id}?sendUpdates=none`, { method: "DELETE" });
}

async function googleBusy(start, end) {
  if (!config.googleBusyCalendars.length) return [];
  const r = await gcal("freeBusy", {
    method: "POST",
    body: JSON.stringify({ timeMin: start, timeMax: end, items: config.googleBusyCalendars.map((id) => ({ id })) }),
  });
  return Object.values(r.calendars).flatMap((c) => c.busy || []).map((b) => ({
    subject: "google busy", showAs: "busy", isCancelled: false,
    start: { dateTime: b.start }, end: { dateTime: b.end },
  }));
}
const CATEGORY = "Booked via rahulch.site";

// ---------- Graph via Maton ----------
async function graph(path, opts = {}) {
  const url = path.startsWith("http")
    ? path.replace("https://graph.microsoft.com/", GATEWAY)
    : GATEWAY + path;
  const r = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${MATON_KEY}`,
      "Content-Type": "application/json",
      Prefer: 'outlook.timezone="UTC"',
      ...(opts.headers || {}),
    },
  });
  if (r.status === 204) return null;
  const text = await r.text();
  const body = text ? JSON.parse(text) : null;
  if (!r.ok) {
    if (opts.onFail) await opts.onFail();
    const err = new Error(`graph ${r.status}: ${body?.error?.message || text}`);
    err.status = r.status;
    throw err;
  }
  return body;
}

// Busy blocks from the default calendar, cached for 60 s.
let busyCache = { at: 0, events: [] };
async function getEvents(force = false) {
  if (!force && Date.now() - busyCache.at < 60_000) return busyCache.events;
  const start = DateTime.utc().minus({ days: 1 }).toISO();
  const end = DateTime.utc().plus({ days: config.maxDaysAhead + 2 }).toISO();
  let next =
    `v1.0/me/calendarView?startDateTime=${start}&endDateTime=${end}` +
    `&$select=subject,start,end,showAs,isAllDay,isCancelled,categories,attendees&$top=200`;
  const out = [];
  while (next) {
    const page = await graph(next);
    out.push(...page.value);
    next = page["@odata.nextLink"];
  }
  // Holiday calendars: block the whole day in the host timezone.
  const cals = (await graph("v1.0/me/calendars?$select=id,name")).value;
  for (const cal of cals.filter((x) => config.holidayCalendars.includes(x.name))) {
    const h = await graph(
      `v1.0/me/calendars/${encodeURIComponent(cal.id)}/calendarView?startDateTime=${start}&endDateTime=${end}&$select=subject,start&$top=100`
    );
    for (const e of h.value) {
      const day = DateTime.fromISO(e.start.dateTime.slice(0, 10), { zone: config.hostTz });
      out.push({
        subject: e.subject, showAs: "busy", isCancelled: false,
        start: { dateTime: day.toUTC().toISO({ includeOffset: false }) },
        end: { dateTime: day.plus({ days: 1 }).toUTC().toISO({ includeOffset: false }) },
      });
    }
  }
  try {
    out.push(...(await googleBusy(start, end)));
  } catch (e) {
    console.error("google busy skipped:", e.message);
  }
  busyCache = { at: Date.now(), events: out };
  return out;
}

const utc = (s) => DateTime.fromISO(s.endsWith("Z") ? s : s + "Z", { zone: "utc" });

function busyIntervals(events) {
  const pad = { minutes: config.bufferMinutes };
  return events
    .filter((e) => !e.isCancelled && config.blockingShowAs.includes(e.showAs))
    .map((e) => Interval.fromDateTimes(utc(e.start.dateTime).minus(pad), utc(e.end.dateTime).plus(pad)));
}

// ---------- slot math ----------
function computeSlots(minutes, events) {
  const tz = config.hostTz;
  const busy = busyIntervals(events);
  const earliest = DateTime.utc().plus({ hours: config.minNoticeHours });
  const step = Math.min(minutes, 30);
  const perDay = {};
  for (const e of events) {
    if (e.isCancelled || !(e.categories || []).includes(CATEGORY)) continue;
    const d = utc(e.start.dateTime).setZone(tz).toISODate();
    perDay[d] = (perDay[d] || 0) + 1;
  }

  const slots = [];
  const today = DateTime.now().setZone(tz).startOf("day");
  for (let i = 0; i <= config.maxDaysAhead; i++) {
    const day = today.plus({ days: i });
    if ((perDay[day.toISODate()] || 0) >= config.maxPerDay) continue;
    for (const [a, b] of config.hours[day.weekday] || []) {
      const [ah, am] = a.split(":").map(Number);
      const [bh, bm] = b.split(":").map(Number);
      const winEnd = day.set({ hour: bh, minute: bm });
      for (let t = day.set({ hour: ah, minute: am }); t.plus({ minutes }) <= winEnd; t = t.plus({ minutes: step })) {
        if (t < earliest) continue;
        const iv = Interval.fromDateTimes(t, t.plus({ minutes }));
        if (busy.some((bz) => bz.overlaps(iv))) continue;
        slots.push(t.toUTC().toISO({ suppressMilliseconds: true }));
      }
    }
  }
  return slots;
}

// ---------- signed cancel tokens ----------
const b64 = (s) => Buffer.from(s).toString("base64url");
const sign = (payload) => crypto.createHmac("sha256", SECRET).update(payload).digest("base64url").slice(0, 22);
function makeToken(start, email) {
  const p = b64(JSON.stringify({ s: start, e: email.toLowerCase() }));
  return `${p}.${sign(p)}`;
}
function readToken(tok) {
  const [p, sig] = String(tok || "").split(".");
  if (!p || !sig || sign(p) !== sig) return null;
  try { return JSON.parse(Buffer.from(p, "base64url").toString()); } catch { return null; }
}

// ---------- abuse limits (in memory, reset on deploy) ----------
const hits = new Map();
function limited(key, max, windowMs) {
  const now = Date.now();
  const arr = (hits.get(key) || []).filter((t) => now - t < windowMs);
  arr.push(now);
  hits.set(key, arr);
  return arr.length > max;
}
const locks = new Set();

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

// ---------- app ----------
const app = new Hono();
app.use("/api/*", cors({ origin: config.allowedOrigins, allowMethods: ["GET", "POST"] }));

app.get("/healthz", (c) => c.text("ok"));

app.get("/api/config", (c) =>
  c.json({
    host: config.hostName,
    hostTz: config.hostTz,
    eventTypes: config.eventTypes,
    maxDaysAhead: config.maxDaysAhead,
    hasMeetingLink: Boolean(MEETING_LINK),
  })
);

app.get("/api/slots", async (c) => {
  const type = config.eventTypes[c.req.query("type") || "30"];
  if (!type) return c.json({ error: "unknown type" }, 400);
  const events = await getEvents();
  return c.json({ type: type.minutes, slots: computeSlots(type.minutes, events) });
});

app.post("/api/book", async (c) => {
  const ip = c.req.header("x-forwarded-for")?.split(",")[0].trim() || "local";
  if (limited(`ip:${ip}`, 5, 3600_000)) return c.json({ error: "Too many requests. Try again later." }, 429);

  const b = await c.req.json().catch(() => ({}));
  if (b.website) return c.json({ ok: true }); // honeypot: pretend success
  const type = config.eventTypes[String(b.type)];
  const name = String(b.name || "").trim().slice(0, 80);
  const email = String(b.email || "").trim().toLowerCase().slice(0, 120);
  const notes = String(b.notes || "").trim().slice(0, 1500);
  const guestTz = String(b.tz || "UTC").slice(0, 60);
  const start = DateTime.fromISO(String(b.start || ""), { zone: "utc" });
  if (!type || !name || !EMAIL_RE.test(email) || !start.isValid)
    return c.json({ error: "Name, a valid email and a time are required." }, 400);
  if (limited(`em:${email}`, 2, 86400_000)) return c.json({ error: "You already booked today." }, 429);

  const startIso = start.toISO({ suppressMilliseconds: true });
  if (locks.has(startIso)) return c.json({ error: "That time was just taken. Pick another." }, 409);
  locks.add(startIso);
  try {
    const fresh = await getEvents(true);
    if (!computeSlots(type.minutes, fresh).includes(startIso))
      return c.json({ error: "That time is no longer available. Pick another." }, 409);

    const end = start.plus({ minutes: type.minutes });
    const endIso = end.toISO({ suppressMilliseconds: true });
    let link = MEETING_LINK;
    try {
      link = (await createMeet({ startIso, endIso, title: `${name} <> ${config.hostName}`, email })) || MEETING_LINK;
    } catch (e) {
      console.error("meet link skipped:", e.message);
    }
    const tok = makeToken(startIso, email);
    const cancelUrl = `${SITE_URL}/meet/cancel/?t=${tok}`;
    const rescheduleUrl = `${SITE_URL}/meet/?reschedule=${tok}`;
    const guestWhen = start.setZone(guestTz).toFormat("ccc d LLL yyyy, h:mm a ZZZZ");
    const body = `
      <p><b>${esc(type.title)}</b> with ${esc(config.hostName)} and ${esc(name)}.</p>
      <p>${link ? `Join Google Meet: <a href="${esc(link)}">${esc(link)}</a>` : "Video link to follow by email."}</p>
      ${config.hostPhone ? `<p>Can't join the video call? You can also call ${esc(config.hostName.split(" ")[0])} at <a href="tel:${config.hostPhone.replace(/[^\d+]/g, "")}">${esc(config.hostPhone)}</a>.</p>` : ""}
      ${notes ? `<p>Notes from ${esc(name)}:<br>${esc(notes).replace(/\n/g, "<br>")}</p>` : ""}
      <p>Need to reschedule? <a href="${rescheduleUrl}">${rescheduleUrl}</a></p>
      <p>Need to cancel? <a href="${cancelUrl}">${cancelUrl}</a></p>
      <p style="color:#64748b">Booked through rahulch.site/meet for ${esc(guestWhen)}.</p>`;

    const ev = await graph("v1.0/me/events", {
      onFail: () => link && link !== MEETING_LINK && deleteMeet(startIso, email).catch(() => {}),
      method: "POST",
      body: JSON.stringify({
        subject: `${name} <> ${config.hostName.split(" ")[0]} (${type.minutes} min)`,
        body: { contentType: "HTML", content: body },
        start: { dateTime: start.toFormat("yyyy-MM-dd'T'HH:mm:ss"), timeZone: "UTC" },
        end: { dateTime: end.toFormat("yyyy-MM-dd'T'HH:mm:ss"), timeZone: "UTC" },
        location: { displayName: link || "Video call" },
        attendees: [{ emailAddress: { address: email, name }, type: "required" }],
        categories: [CATEGORY],
        showAs: "busy",
        isReminderOn: true,
        reminderMinutesBeforeStart: 10,
        responseRequested: true,
        allowNewTimeProposals: true,
        transactionId: crypto.createHash("sha256").update(startIso + email).digest("hex").slice(0, 36),
      }),
    });
    busyCache.at = 0;
    return c.json({ ok: true, start: startIso, end: endIso, meet: link || null, id: ev.id.slice(-12) });
  } finally {
    locks.delete(startIso);
  }
});

// Cancel page (linked from the invite body)
const page = (msg, form = "") => `<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Cancel meeting</title><body style="font:13px/1.45 system-ui,sans-serif;margin:10px 12px;color:#0f172a">
<p>${msg}</p>${form}<p><a href="https://rahulch.site/meet/">rahulch.site/meet</a></p></body>`;

async function findBooking(tok) {
  const d = readToken(tok);
  if (!d) return null;
  const s = utc(d.s);
  const r = await graph(
    `v1.0/me/calendarView?startDateTime=${s.minus({ minutes: 1 }).toISO()}&endDateTime=${s.plus({ minutes: 1 }).toISO()}` +
      `&$select=subject,start,end,categories,attendees,isCancelled,body`
  );
  const ev = r.value.find(
    (e) => !e.isCancelled && (e.categories || []).includes(CATEGORY) &&
      (e.attendees || []).some((a) => a.emailAddress.address.toLowerCase() === d.e)
  );
  return ev ? { ev, d } : null;
}

// JSON endpoints used by rahulch.site/meet/cancel/
async function cancelBooking(f) {
  await graph(`v1.0/me/events/${encodeURIComponent(f.ev.id)}/cancel`, {
    method: "POST",
    body: JSON.stringify({ comment: "Cancelled by the guest through rahulch.site/meet." }),
  });
  try {
    await deleteMeet(f.d.s, f.d.e);
  } catch (e) {
    console.error("meet cleanup skipped:", e.message);
  }
  busyCache.at = 0;
}

app.get("/api/booking", async (c) => {
  const f = await findBooking(c.req.query("t"));
  if (!f) return c.json({ error: "not found" }, 404);
  const s = utc(f.ev.start.dateTime), e = utc(f.ev.end.dateTime);
  const guest = (f.ev.attendees || []).find((a) => a.emailAddress.address.toLowerCase() === f.d.e);
  return c.json({
    name: guest?.emailAddress.name || "",
    email: f.d.e,
    subject: f.ev.subject,
    start: s.toISO({ suppressMilliseconds: true }),
    minutes: Math.round(e.diff(s, "minutes").minutes),
  });
});

app.post("/api/reschedule", async (c) => {
  const ip = c.req.header("x-forwarded-for")?.split(",")[0].trim() || "local";
  if (limited(`rs:${ip}`, 6, 3600_000)) return c.json({ error: "Too many requests." }, 429);
  const b = await c.req.json().catch(() => ({}));
  const f = await findBooking(b.t);
  if (!f) return c.json({ error: "This meeting was not found. It may have been cancelled." }, 404);

  const oldStart = utc(f.ev.start.dateTime);
  const minutes = Math.round(utc(f.ev.end.dateTime).diff(oldStart, "minutes").minutes);
  const start = DateTime.fromISO(String(b.start || ""), { zone: "utc" });
  if (!start.isValid) return c.json({ error: "Pick a new time." }, 400);
  const startIso = start.toISO({ suppressMilliseconds: true });
  if (startIso === f.d.s) return c.json({ error: "That is the current time. Pick a different one." }, 400);
  if (locks.has(startIso)) return c.json({ error: "That time was just taken. Pick another." }, 409);
  locks.add(startIso);
  try {
    // The meeting being moved should not block its own new time.
    const fresh = (await getEvents(true)).filter((e) => e.id !== f.ev.id);
    if (!computeSlots(minutes, fresh).includes(startIso))
      return c.json({ error: "That time is no longer available. Pick another." }, 409);

    const end = start.plus({ minutes });
    const endIso = end.toISO({ suppressMilliseconds: true });
    const oldTok = String(b.t).trim();
    const newTok = makeToken(startIso, f.d.e);
    const guestTz = String(b.tz || "UTC").slice(0, 60);
    const guestWhen = start.setZone(guestTz).toFormat("ccc d LLL yyyy, h:mm a ZZZZ");
    const html = f.ev.body.content
      .split(oldTok).join(newTok)
      .replace(/(Booked through rahulch\.site\/meet for|Rescheduled through rahulch\.site\/meet to) [^<]+/,
        `Rescheduled through rahulch.site/meet to ${esc(guestWhen)}.`);

    await graph(`v1.0/me/events/${encodeURIComponent(f.ev.id)}`, {
      method: "PATCH",
      body: JSON.stringify({
        start: { dateTime: start.toFormat("yyyy-MM-dd'T'HH:mm:ss"), timeZone: "UTC" },
        end: { dateTime: end.toFormat("yyyy-MM-dd'T'HH:mm:ss"), timeZone: "UTC" },
        body: { contentType: "HTML", content: html },
      }),
    });

    // Move the Google placeholder so the Meet link stays tied to the new time.
    try {
      const r = await gcal(`calendars/primary/events?privateExtendedProperty=${encodeURIComponent(`meetKey=${meetKey(f.d.s, f.d.e)}`)}`);
      for (const ge of r?.items || [])
        await gcal(`calendars/primary/events/${ge.id}?sendUpdates=none`, {
          method: "PATCH",
          body: JSON.stringify({
            start: { dateTime: startIso }, end: { dateTime: endIso },
            extendedProperties: { private: { meetKey: meetKey(startIso, f.d.e) } },
          }),
        });
    } catch (e) {
      console.error("meet move skipped:", e.message);
    }
    busyCache.at = 0;
    return c.json({ ok: true, start: startIso, end: endIso });
  } finally {
    locks.delete(startIso);
  }
});

app.post("/api/cancel", async (c) => {
  const ip = c.req.header("x-forwarded-for")?.split(",")[0].trim() || "local";
  if (limited(`cx:${ip}`, 10, 3600_000)) return c.json({ error: "Too many requests." }, 429);
  const b = await c.req.json().catch(() => ({}));
  const f = await findBooking(b.t);
  if (!f) return c.json({ error: "This meeting was not found. It may already be cancelled." }, 404);
  await cancelBooking(f);
  return c.json({ ok: true });
});

// Older server-rendered cancel page, kept so earlier links still work.
app.get("/cancel", async (c) => {
  const t = c.req.query("t");
  const f = await findBooking(t);
  if (!f) return c.html(page("This meeting was not found. It may already be cancelled."));
  const when = utc(f.ev.start.dateTime).setZone(config.hostTz).toFormat("ccc d LLL, h:mm a ZZZZ");
  return c.html(page(`Cancel <b>${esc(f.ev.subject)}</b> on ${esc(when)}?`,
    `<form method="post"><input type="hidden" name="t" value="${esc(t)}"><button style="font:inherit;padding:4px 10px;background:#E60023;color:#fff;border:0">Cancel meeting</button></form>`));
});

app.post("/cancel", async (c) => {
  const form = await c.req.parseBody();
  const f = await findBooking(form.t);
  if (!f) return c.html(page("This meeting was not found. It may already be cancelled."));
  await cancelBooking(f);
  return c.html(page("Cancelled. Both calendars will update shortly."));
});

app.onError((err, c) => {
  console.error(err);
  return c.json({ error: "Something went wrong. Try again in a minute." }, 500);
});

const port = Number(process.env.PORT || 10000);
serve({ fetch: app.fetch, port });
console.log(`meet-api on :${port}`);

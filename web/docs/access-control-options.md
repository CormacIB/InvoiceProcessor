# Access Control Options

How to keep strangers who have the URL from loading the Invoice Processor web app.
Not implemented yet — this is a decision document.

## Context

- The app is fully client-side (Vite + React, deployed on Vercel). There is no
  server and no user accounts, and we want to keep it that way.
- Invoices never leave the browser, and each shop's category profiles live in
  their own browser's localStorage. So this is purely about gating **who can
  load the page**, not about protecting stored data.

## ⚠️ The trap: password check inside the React app

Any password check written into the client-side JavaScript ships to the
browser. Anyone can open dev tools, read the bundle, and find the password or
skip the check entirely. It stops casual visitors only — **not real security.
Don't do this.**

## Option 1 — HTTP Basic Auth via Vercel middleware (recommended)

A small `middleware.ts` at the project root runs on Vercel's edge before the
page is served. It checks the `Authorization` header against credentials
stored as environment variables; the browser shows its native login prompt.

- **Cost:** free (works on Vercel's hobby tier).
- **Effort:** ~20 lines of code + 2 env vars in the Vercel dashboard.
- **Security:** real — wrong credentials never receive the app at all, and the
  password lives in Vercel's dashboard, not in the repo or the bundle.
- **Model:** one shared username/password for everyone (per shop would mean
  multiple credential pairs in env vars — workable, but manual).
- **App impact:** none. The app stays 100% client-side; only page delivery is
  gated.

## Option 2 — Vercel built-in Password Protection

A toggle in the Vercel dashboard, zero code.

- **Cost:** paid — requires Vercel Pro (and Password Protection may be an
  add-on on top). Likely overkill for this project.
- **Effort:** none.
- **Security:** real.
- **Model:** one shared password.

## Option 3 — Cloudflare Access in front of the deployment

Put the site behind Cloudflare and enable Cloudflare Access (Zero Trust). Each
user logs in with a one-time email code, against an allowlist of addresses you
control.

- **Cost:** free for small teams (up to ~50 users).
- **Effort:** moderate — DNS through Cloudflare, configure an Access policy.
  No code changes.
- **Security:** real, and per-person.
- **Model:** individual logins you can grant/revoke per email address — the
  best fit if different shops should have separately revocable access.
- **Trade-off:** adds Cloudflare as a second service in front of Vercel; more
  moving parts to maintain.

## Recommendation

Start with **Option 1** (Basic Auth middleware): free, tiny, keeps the app
client-side, and matches the "give each shop the link plus a password" model.
Revisit **Option 3** later if shops ever need individually revocable access.

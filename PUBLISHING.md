# Publishing Delta to Microsoft AppSource

A step-by-step runbook, from where the project is today to a live store listing.
Written assuming no prior experience with AppSource.

## What already exists

| Thing | File | Status |
|---|---|---|
| Dev manifest (localhost, for your own Excel) | `manifest.xml` | working, sideloaded |
| Production manifest (for the store) | `manifest.prod.xml` | ready — needs your real domain |
| Privacy policy page | `site/privacy.html` | ready |
| Support page | `site/support.html` | ready |
| App code | `src/` | local-only, no network calls |

The dev and production manifests have **different Ids** on purpose: you can keep
the localhost version sideloaded for development even after the store version
is installed, and Excel treats them as two separate add-ins.

## Step 1 — Host the files publicly (HTTPS)

The store version of the manifest must point at a public HTTPS host serving
these folders exactly as they are in the repo: `src/`, `assets/`, `site/`.

Easiest free option: **GitHub Pages.**

1. Create a repo (public or private with Pages enabled) and push this project.
2. Repo → Settings → Pages → deploy from the main branch, root folder.
3. Your base URL becomes `https://<username>.github.io/<repo>`.
4. Verify in a browser that these load:
   - `<base>/src/taskpane.html`
   - `<base>/assets/icon-32.png`
   - `<base>/site/privacy.html` and `<base>/site/support.html`

(Custom domain like `delta.yourname.com` also works and looks nicer in the
listing, but is optional.)

Note: the vault server (`server/`) and `server.js` are **dev-only** — nothing
in production references them; don't deploy them anywhere.

## Step 2 — Point the production manifest at your host

In `manifest.prod.xml`, replace every `https://delta.example.com` with your
base URL from Step 1. That's the only edit it needs. Then sanity-check it:

```
xmllint --noout manifest.prod.xml
```

## Step 3 — Test the production build before submitting

1. Sideload `manifest.prod.xml` the same way as the dev one (copy it into the
   wef folder — see `sideload.js` — or temporarily point the script at it).
2. Confirm the pane loads from your public host, not localhost.
3. If at all possible, also test on **Excel for Windows** and **Excel on the
   web** (upload a workbook to OneDrive → open in browser → Insert → Add-ins →
   Upload My Add-in). AppSource validation tests all platforms, and this is
   where surprises live. Known thing to watch: the pane was developed on Mac —
   check fonts and the sticky bars render sensibly on Windows.

## Step 4 — Partner Center account

1. Go to https://partner.microsoft.com and register for the
   **Microsoft AppSource / Office Store program** (one-time fee, individual
   accounts are fine — you don't need a company).
2. Identity verification can take a few days; start this early.

## Step 5 — The listing

Prepare before you start filling the form:

- **Name:** "Delta" is short and may collide with existing trademarks/listings;
  have "Delta – Version History for Excel" ready as a fallback title.
- **Short description** (~100 chars): e.g. *"Version history for Excel — review
  every change cell by cell, save versions, revert safely."*
- **Long description:** what it does, the local-only privacy story is a selling
  point — lead with it.
- **Screenshots:** at least one, 1366×768 PNG. Good candidates: the diff view
  with green/amber/red changes; the grid with highlight frames + tinted tabs;
  the revert preview.
- **Links:** privacy = `<base>/site/privacy.html`, support =
  `<base>/site/support.html`, terms of use (can be a short page; can add later).
- **Pricing:** Free.

Upload `manifest.prod.xml`, fill in the above, submit.

## Step 6 — Validation

Microsoft reviews against their checklist (typically days to ~2 weeks).
Common rejection reasons to pre-empt:

- Add-in fails to load on one of the platforms (test Step 3!).
- Description promises features that aren't in the add-in.
- Missing/broken privacy or support URLs.
- Icons: the listing also wants a store logo (usually 300×300) — export a
  larger version of the check-in-square logo before submitting.

If rejected, the report says exactly what to fix; resubmitting is normal —
most add-ins take a round or two.

## Later: taking money

Ship free first. When ready, the usual path for Office add-ins is a freemium
license key: the add-in stays free in the store, an in-app "Pro" unlock talks
to your own payment page (Stripe/Paddle/LemonSqueezy). Note this reintroduces
a network call for license checking, so `site/privacy.html` must be updated
first — it currently promises zero network activity, and that promise is why
people will install it.

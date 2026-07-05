# Quinta da Aldeia — Standalone Quote Form
## Deployment Guide

---

## What this is

A completely separate public website containing only the quote request form.
Clients access it directly — they never see your management app.

---

## Files in this folder

| File | Purpose |
|---|---|
| `index.html` | The quote request form (opens directly as the homepage) |
| `logo.png` | Quinta da Aldeia logo |
| `config.js` | Your Apps Script URL (must be updated before deploying) |

---

## Step 1 — Update config.js

Open `config.js` and paste your Google Apps Script URL:

```js
window.SCRIPT_URL = 'https://script.google.com/macros/s/YOUR_ID/exec';
```

This is the same Apps Script URL used by your main app.

---

## Step 2 — Deploy to Netlify as a NEW site

1. Go to https://netlify.com and log in
2. On the dashboard, click **Add new site → Deploy manually**
3. Drag and drop this entire **quinta-form-site** folder
4. Netlify gives you a URL like `https://amazing-name-123.netlify.app`
5. Optional: rename it to something memorable like `quinta-orcamento`
   - Go to **Site settings → General → Site name** → change it
   - Your URL becomes `https://fcbasto-a11y.github.io/quinta-form-site`

This is a completely separate site from your management app.

---

## Step 3 — Update your Gmail auto-reply

In your Gmail auto-reply message, use this form URL:

```
https://fcbasto-a11y.github.io/quinta-form-site
```

(Replace with your actual Netlify URL)

---

## Step 4 — Update BOOKING_URL in Apps Script

In your `appsscript.gs`, update the BOOKING_URL to point to your main app
(not the form site):

```js
var BOOKING_URL = 'https://your-main-app.netlify.app';
```

The `QUOTE_FORM_URL` should point to the form site:
```js
var QUOTE_FORM_URL = 'https://fcbasto-a11y.github.io/quinta-form-site';
```

---

## How it works

```
Client receives auto-reply email
        ↓
Clicks link → opens https://fcbasto-a11y.github.io/quinta-form-site
(completely separate from your management app)
        ↓
Fills in the form and submits
        ↓
Apps Script saves to QuoteRequests sheet
Apps Script emails you a notification
Apps Script sends client a confirmation
        ↓
You log in to your management app to review the request
```

---

## Updating the form later

Made a change to `index.html`? Just go to your Netlify form site
dashboard → Deploys → drag and drop this folder again.

# Kartik Tungenwar — Portfolio (Next.js)

A rebuild of [karttiktungenwar.github.io/portfolio](https://karttiktungenwar.github.io/portfolio/)
as an interactive Next.js site, with a contact form backed by Firebase —
**without ever exposing Firebase credentials in the browser.**

Tech: Next.js 14 (App Router) · TypeScript · Tailwind CSS · Framer Motion · Firebase Admin SDK

---

## How the security actually works (read this first)

A lot of portfolio tutorials paste a Firebase **client** config (apiKey, authDomain,
etc.) straight into the frontend. That config isn't a secret by itself — Firebase
expects it to be public — but it's not what you want for a contact form either,
because it means anyone can call your Firestore directly from their own browser
console.

This project avoids that entirely:

- The contact form never talks to Firebase directly. It calls **your own**
  API route: `app/api/contact/route.ts`.
- That route runs only on the server (`export const runtime = "nodejs"`) and
  uses the **Firebase Admin SDK** with a service account key.
- The service account key lives only in environment variables
  (`.env.local` locally, your host's dashboard in production) — it is never
  imported by any `"use client"` component, never bundled into JavaScript
  sent to the browser, and never committed to git (`.gitignore` blocks it).
- The browser only ever sees your own domain's `/api/contact` endpoint.

So: no Firebase URL, key, or project ID is ever visible in browser dev tools,
"View Source", or your GitHub repo.

---

## Project structure

```
portfolio-nextjs/
├── app/
│   ├── api/contact/route.ts   # Secure server-side endpoint (Firebase Admin lives here)
│   ├── layout.tsx             # Fonts + global metadata
│   ├── page.tsx               # Assembles all sections
│   └── globals.css
├── components/                # Navbar, Hero, About, Experience, Projects, Testimonials, Contact, Footer
├── data/portfolio.ts          # All your text content — edit this to update the site
├── lib/firebaseAdmin.ts       # Server-only Firebase Admin singleton
├── .env.local.example         # Template — copy to .env.local
└── .gitignore                 # Excludes .env* and service account files
```

---

## Step 1 — Prerequisites

Install once, if you don't already have them:

- [Node.js](https://nodejs.org/) 18.18 or later (`node -v` to check)
- [Git](https://git-scm.com/)
- A free [Firebase](https://firebase.google.com/) account
- A free [Vercel](https://vercel.com/) account (recommended host — see Step 6 for why)

---

## Step 2 — Get the project onto your machine

If you're starting from this generated project folder, unzip it, then:

```bash
cd portfolio-nextjs
npm install
```

If you'd rather start the repo fresh and push this into
`github.com/karttiktungenwar/portfolio`:

```bash
git clone https://github.com/karttiktungenwar/portfolio.git
cd portfolio
# copy all files from this project into the repo folder, then:
npm install
```

> Note: your current repo serves a static HTML site directly from
> `main`/`gh-pages`. Consider building this Next.js version on a new
> branch (e.g. `nextjs-rebuild`) first, so your live site keeps working
> while you set things up.

---

## Step 3 — Run it locally

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000). You should see the site
with animations working. The contact form will show an error until you
finish Step 4 and 5 — that's expected.

---

## Step 4 — Create the Firebase project and service account

1. Go to the [Firebase Console](https://console.firebase.google.com/) → **Add project**
   → name it (e.g. `kartik-portfolio`) → finish the wizard.
2. In the left sidebar, go to **Build → Firestore Database** → **Create database**
   → start in **production mode** → choose a region close to your users.
3. Click the **gear icon → Project settings → Service accounts** tab.
4. Click **Generate new private key**. A JSON file downloads — treat this
   like a password. **Do not commit it to GitHub.**
5. Open that JSON file. You need three fields from it:
   - `project_id` → this is your `FIREBASE_PROJECT_ID`
   - `client_email` → this is your `FIREBASE_CLIENT_EMAIL`
   - `private_key` → this is your `FIREBASE_PRIVATE_KEY`

### Lock down Firestore rules

Since only your server (via the Admin SDK) writes to Firestore, you can
deny all client-side access. In **Firestore → Rules**:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if false;
    }
  }
}
```

The Admin SDK bypasses these rules entirely (it authenticates as an admin),
so this is safe — it just stops anyone else from reading or writing your
Firestore from a browser.

---

## Step 5 — Set your environment variables locally

```bash
cp .env.local.example .env.local
```

Open `.env.local` and paste in your three values:

```env
FIREBASE_PROJECT_ID=kartik-portfolio
FIREBASE_CLIENT_EMAIL=firebase-adminsdk-xxxxx@kartik-portfolio.iam.gserviceaccount.com
FIREBASE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nMIIEvQ...\n-----END PRIVATE KEY-----\n"
```

Important: keep the `private_key` value **exactly as it appears in the JSON
file**, including the `\n` characters and the surrounding quotes. The code
in `lib/firebaseAdmin.ts` converts those `\n` sequences back into real
line breaks.

Restart `npm run dev`, fill out the contact form, and submit. Check
**Firestore → Data** in the Firebase console — you should see a new
document in a `contactMessages` collection.

---

## Step 6 — Choose where to deploy

Your API route needs a server that runs Node.js on each request. **GitHub
Pages only serves static files and cannot run this route** — this is why
we recommend **Vercel** (made by the creators of Next.js, free tier is
generous, and it deploys API routes automatically).

You can still keep your source code in
`github.com/karttiktungenwar/portfolio` — Vercel deploys straight from
GitHub, you just won't use GitHub Pages as the *host*.

If you truly want to stay on GitHub Pages, see "Alternative: staying on
GitHub Pages" at the bottom of this file — it trades away the server-side
security model above.

### Deploy to Vercel

1. Push this project to your GitHub repo:
   ```bash
   git add .
   git commit -m "Rebuild portfolio in Next.js with secure contact form"
   git push origin nextjs-rebuild   # or your chosen branch
   ```
2. Go to [vercel.com/new](https://vercel.com/new) → **Import Git Repository**
   → select `karttiktungenwar/portfolio`.
3. Vercel auto-detects Next.js. Before clicking Deploy, expand
   **Environment Variables** and add the same three keys from your
   `.env.local`:
   - `FIREBASE_PROJECT_ID`
   - `FIREBASE_CLIENT_EMAIL`
   - `FIREBASE_PRIVATE_KEY` (paste with the `\n` literal characters, same as locally)
4. Click **Deploy**. Vercel gives you a URL like
   `portfolio-kartiktungenwar.vercel.app` immediately.
5. Optional: **Project Settings → Domains** → add a custom domain, or point
   it at a subdomain if you want to keep `karttiktungenwar.github.io` for
   something else.

Because the environment variables are stored in Vercel's dashboard (not
in your code), your Firebase credentials never touch GitHub at any point.

---

## Step 7 — Keep GitHub and Vercel in sync

Once connected, every `git push` to your chosen branch automatically
triggers a new Vercel deployment. Typical workflow:

```bash
# edit data/portfolio.ts or any component
git add .
git commit -m "Update project descriptions"
git push
```

Vercel rebuilds and redeploys within a minute or two — no manual steps.

---

## Editing your content

Almost everything on the page — your name, experience, project list,
testimonials, social links — lives in `data/portfolio.ts`. Update the
values there; you shouldn't need to touch the component files for routine
content changes.

Replace the placeholder profile image: add your photo to `/public` (e.g.
`/public/profile.jpg`) and reference it from a component with Next's
`<Image />` component if you'd like a photo in the hero section.

---

## Security checklist before you make the repo public

- [ ] `.env.local` is NOT committed (`git status` should not show it —
      `.gitignore` already excludes it)
- [ ] The downloaded service-account JSON file is deleted from your
      Downloads folder or stored somewhere outside the repo
- [ ] Firestore rules deny direct client read/write (Step 4)
- [ ] Environment variables are set in Vercel's dashboard, not hardcoded
      anywhere in the code
- [ ] Run `git log --all --full-history -- .env.local` — should return
      nothing, confirming it was never accidentally committed in an
      earlier commit

---

## Alternative: staying on GitHub Pages

If you need to keep hosting on GitHub Pages, you'd export a static site
(`next.config.js` → `output: "export"`) and the contact form would have to
call Firebase directly from the browser using the Firebase **client** SDK
(not Admin). That means your Firebase client config (apiKey, project ID,
etc.) is visible in the page source — this is normal for Firebase client
apps, but it means your protection now depends entirely on strict
Firestore Security Rules (e.g. only allow `create` on `contactMessages`,
never `read`, and validate field types/lengths in the rules themselves)
rather than a server hiding the credentials. This project is set up for
the server-side approach because it's the simpler security story; ask if
you'd like the static-export + security-rules version instead.

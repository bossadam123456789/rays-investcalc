# Rays InvestCalc — Local Setup

Premium fintech dashboard for project investment analysis.

## ⚡ Quick Start (3 steps)

### 1. Install Node.js (one-time, only if you don't have it)
Download from https://nodejs.org/ (pick the LTS version). Verify by opening Terminal/Command Prompt:
```bash
node --version
```
You should see something like `v20.11.0` or higher.

### 2. Install dependencies
Open Terminal/Command Prompt in this folder and run:
```bash
npm install
```
Takes about 1–2 minutes.

### 3. Run the app
```bash
npm run dev
```
Your browser will open automatically at **http://localhost:5173**. 🚀

The Kanyakwar Phase 3 project is pre-loaded — you'll see it immediately on the dashboard.

---

## 📦 Other commands

- `npm run build` — Build for production (creates a `dist/` folder)
- `npm run preview` — Preview the production build locally

---

## 🐛 Troubleshooting

**"npm: command not found"** → Install Node.js first (Step 1).

**"Port 5173 is already in use"** → Either close the other process or edit `vite.config.js` and change `port: 5173` to `port: 5174`.

**Page is blank** → Open the browser console (Right-click → Inspect → Console tab). Errors here tell you what's wrong. Most common: a missing dependency. Re-run `npm install`.

**Data doesn't save when I refresh** → That's expected in some browser modes. Data is saved in `localStorage` under the key `rays_projects_v7`. Open DevTools → Application → Local Storage to inspect.

---

## 🌐 Deploying later (when you're ready)

The easiest path is Vercel:
1. Create a free account at https://vercel.com
2. Install: `npm install -g vercel`
3. In this folder: `vercel`
4. Follow the prompts — first deploy is live in ~30 seconds.

Or push to GitHub and connect the repo on vercel.com — it auto-deploys on every push.

---

## 📁 Folder structure

```
rays-investcalc/
├── index.html                  ← App entry HTML
├── package.json                ← Dependencies & scripts
├── vite.config.js              ← Vite build config
├── tailwind.config.js          ← Tailwind setup
├── postcss.config.js           ← PostCSS pipeline
├── public/
│   └── favicon.svg             ← Blue calculator favicon
└── src/
    ├── main.jsx                ← React mount point
    ├── index.css               ← Global Tailwind imports
    └── investment_calculator.jsx  ← The whole app (~5000 lines)
```

To edit the app, change `src/investment_calculator.jsx`. Save the file — the page auto-reloads.

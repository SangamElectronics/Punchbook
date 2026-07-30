# Punchbook — Attendance Ledger

A multi-employee attendance tracking app: punch-in/out records, shifts,
late-in/early-out calculations, overtime, monthly reports, CSV/Excel
import-export, and a dashboard — all in the browser.

Data is stored in **your browser's `localStorage`**, so it's fully private
to whoever is using that specific browser on that specific device. There is
no backend server and no shared database (see "Sharing data across
devices" below if you need that).

---

## 1. Run it locally

You need [Node.js](https://nodejs.org) 18 or later installed.

```bash
npm install
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`). Changes to the
code hot-reload automatically.

---

## 2. Put it on GitHub

If you don't already have a repo:

```bash
cd punchbook
git init
git add .
git commit -m "Initial commit"
```

Then on GitHub.com:
1. Click **New repository**, name it (e.g. `punchbook`), don't initialize
   it with a README (you already have one).
2. Copy the commands GitHub shows you under "…or push an existing
   repository from the command line", something like:

```bash
git remote add origin https://github.com/YOUR_USERNAME/punchbook.git
git branch -M main
git push -u origin main
```

---

## 3. Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and sign in with your GitHub
   account.
2. Click **Add New… → Project**.
3. Select your `punchbook` repository and click **Import**.
4. Vercel auto-detects Vite — the defaults are already correct:
   - **Framework Preset:** Vite
   - **Build Command:** `npm run build` (auto-filled)
   - **Output Directory:** `dist` (auto-filled)
5. Click **Deploy**. In about a minute you'll get a live URL like
   `https://punchbook-yourname.vercel.app`.

From then on, every time you `git push` to `main`, Vercel automatically
rebuilds and redeploys the live site. No manual redeploy steps needed.

---

## 4. Using the app

- First screen asks for your name and role (Admin / Manager / Viewer) —
  this is a view-mode switch, not a secured login.
- **Employees** — add your team, their default shift, department.
- **Settings** — add/edit shifts (start/end time, break, full/half-day
  hour thresholds) and grace-period minutes.
- **Attendance** — pick an employee and month, click a day to log
  first-in / last-out / day type (working, week off, leave).
- **Reports / Dashboard** — auto-calculated late-in, early-out, worked
  hours, overtime, and attendance % for any employee or the whole company.
- Data can be exported to Excel/CSV and imported back in bulk from the
  Attendance view.

---

## Important: data storage and limitations

- **Per-browser only.** Data lives in `localStorage` on whatever browser/
  device you're using it in. Opening the site on a different browser or
  device starts with empty data. Clearing browser data/cache will erase it.
- **Not encrypted, not backed up.** Export to Excel/CSV regularly (built
  into the app) if you want a backup.
- **No multi-user sync.** If several people need to see the *same* shared
  data from different devices, this app in its current form can't do that
  — it would need a real backend/database (e.g. Supabase, Firebase, or a
  small custom API) to replace the `src/storage.js` layer. The rest of the
  app's logic wouldn't need to change, since all reads/writes already go
  through that one file.

---

## Project structure

```
punchbook/
├── index.html          # HTML entry point
├── package.json         # Dependencies & scripts
├── vite.config.js       # Build config
└── src/
    ├── main.jsx          # React entry point
    ├── storage.js        # localStorage-backed data layer
    └── App.jsx           # The entire application (UI + logic)
```

## Tech stack

React 18, Vite, [lucide-react](https://lucide.dev) (icons),
[recharts](https://recharts.org) (charts), [xlsx](https://sheetjs.com) and
[papaparse](https://www.papaparse.com/) (Excel/CSV import-export).

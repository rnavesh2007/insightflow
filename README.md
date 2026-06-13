# InsightFlow — AI-Powered Data Dashboard

> © 2026 Navesh R. All rights reserved.

An AI-powered interactive data analytics dashboard. Upload any spreadsheet and instantly get charts, statistics, and Claude AI-generated insights.

---

## Features

- 🤖 **AI Analysis** — Claude AI generates insights, anomalies & recommendations
- 📊 **6 Chart Types** — Bar, Area, Line, Pie, Scatter, Radar
- 📋 **Stats Engine** — Mean, Median, Q1, Q3, Std Dev and more
- 🗄️ **Data Table** — Searchable, sortable, paginated
- 💾 **CSV Export** — Download processed data
- ⚡ **No backend** — Pure React, runs in browser

---

## Setup

1. Clone this repo
2. Install dependencies:
```bash
npm install
```
3. Open `src/DataDashboard.jsx`
4. Replace `YOUR_API_KEY_HERE` with your Anthropic API key:
```js
"x-api-key": "sk-ant-..."
```
5. Run the app:
```bash
npm run dev
```

---

## Deployment (Vercel)

1. Push to GitHub
2. Go to vercel.com → Import repo
3. Framework: Vite
4. Deploy!

Live URL: https://insightflow-beta.vercel.app

---

## Supported File Types

| Format | Description |
|--------|-------------|
| `.csv` | Comma-separated values |
| `.xlsx` | Excel spreadsheet |
| `.xls` | Legacy Excel format |

---

## Tech Stack

- **Framework** — React + Vite
- **Charts** — Recharts
- **CSV Parsing** — PapaParse
- **Excel Parsing** — SheetJS
- **AI** — Anthropic Claude API
- **Hosting** — Vercel

---

## Project Structure

```
insightflow/
├── src/
│   ├── App.jsx
│   └── DataDashboard.jsx
├── public/
├── index.html
├── package.json
└── vite.config.js
```

---

## License

```
© 2025 Navesh R. All rights reserved.

Unauthorized reproduction or distribution of this software,
in whole or in part, is strictly prohibited.

Intellectual property of Navesh R
B.Tech AI & Data Science
Vel Tech High Tech Dr. Rangarajan Dr. Sakunthala Engineering College
Chennai, Tamil Nadu, India
```

---

*Built with ❤️ by Navesh R. | Powered by Anthropic Claude AI*

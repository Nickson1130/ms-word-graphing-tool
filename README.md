# 🧮 MathGraph VBA Designer (DSE Style)

Generate professional, DSE-style mathematical graphs and export them as native VBA scripts for Microsoft Word.

## ✨ Features

- **Multiple Curves** — Plot several equations on the same graph, each with its own line style, color, and weight.
- **Disjoint Domains & Ranges** — Set custom domain and range restrictions per curve, including multiple disjoint intervals (e.g. plot `y = √x` only on `[0, 2] ∪ [4, 6]`).
- **Implicit & Explicit Functions** — Supports both `y = f(x)` and implicit forms like `x² + y² = 9`.
- **DSE Style** — Tick marks on intercepts, italicized axis labels, and clean coordinate axes with arrowheads.
- **Custom Labels** — Add italic point labels with Cross (×) or Dot (●) markers anywhere on the graph.
- **VBA Export** — Generates editable Word shapes via VBA instead of blurry screenshots.

## 🛠️ How to Use

1. **Add equations** — Type your expressions and add as many curves as needed.
2. **Set intervals** — For each curve, optionally restrict the domain and/or range. Add multiple intervals for disjoint regions.
3. **Customize** — Adjust axis ranges, labels, arrow styles, line styles, and colors.
4. **Export** — Click **Generate VBA Script** and copy the code.
5. **In Word** — Go to **Developer → Visual Basic**, then **Insert → Module**, paste the code, and press **F5** to run the `DrawGraph` macro.

## 🚀 Deployment

### GitHub Pages (Free)
1. Go to repository **Settings → Pages**.
2. Set **Source** to `GitHub Actions`.
3. Add your `GEMINI_API_KEY` under **Settings → Secrets and variables → Actions**.
4. Push to `main` — the site deploys automatically.

### Vercel / Netlify (Recommended)
1. Connect your GitHub repo to [Vercel](https://vercel.com/) or [Netlify](https://www.netlify.com/).
2. Set build command to `npm run build` and output directory to `dist`.
3. Add `GEMINI_API_KEY` as an environment variable in the platform settings.
4. Deploy.

## 🏁 Local Development

```bash
npm install
npm run dev
```

Open `http://localhost:3000`.

## 📜 License

MIT

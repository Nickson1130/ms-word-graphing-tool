# 🧮 MathGraph VBA Designer (DSE Style)

Generate professional, DSE-style mathematical graphs and export them as native VBA scripts for Microsoft Word.

Built for HKDSE educators and students to create vector-quality graphs directly in Word — no blurry screenshots, no external dependencies.

## 🌐 Try It Live

**[https://nickson1130.github.io/ms-word-graphing-tool/](https://nickson1130.github.io/ms-word-graphing-tool/)**

No installation needed — just open the link and start designing.

---

## ✨ Features

- **Multiple Curves** — Plot several equations on the same graph, each with its own line style, color, and weight.
- **Disjoint Domains & Ranges** — Restrict each curve to custom intervals, including multiple disjoint regions (e.g. plot `y = √x` only on `[0, 2] ∪ [4, 6]`).
- **Implicit & Explicit Functions** — Supports both `y = f(x)` and implicit forms like `x² + y² = 9`.
- **DSE Style** — Tick marks on intercepts, italicized axis labels, and clean coordinate axes with arrowheads.
- **Custom Labels** — Add italic point labels with Cross (×) or Dot (●) markers anywhere on the graph.
- **VBA Export** — Generates editable Word shapes via VBA — no images, fully scalable.

## 🛠️ How to Use

1. **Add equations** — Type your expressions and add as many curves as needed.
2. **Set intervals** — Optionally restrict the domain and/or range per curve. Add multiple intervals for disjoint regions.
3. **Customize** — Adjust axis ranges, labels, arrow styles, line styles, and colors.
4. **Export** — Click **Generate VBA Script** and copy the code.
5. **In Word** — Go to **Developer → Visual Basic**, then **Insert → Module**, paste the code, and press **F5** to run the `DrawGraph` macro.

## 🏁 Local Development

```bash
npm install
npm run dev
```

Open `http://localhost:3000`.

## 🚀 Self-Hosting

### GitHub Pages (Free)
1. Go to repository **Settings → Pages**.
2. Set **Source** to `GitHub Actions`.
3. Push to `main` — the site deploys automatically.

### Vercel / Netlify
1. Connect your GitHub repo to [Vercel](https://vercel.com/) or [Netlify](https://www.netlify.com/).
2. Set build command to `npm run build` and output directory to `dist`.
3. Deploy.

## 📜 License

MIT

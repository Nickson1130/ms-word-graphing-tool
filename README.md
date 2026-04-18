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

## 📐 Sizing Your Graph

The **Unit Size** slider controls how large the graph will appear in your Word document. It represents the number of **points** (the standard Word measurement unit) per mathematical unit on the graph.

For reference: **72 points = 1 inch ≈ 2.54 cm**. A unit size of **40** means each grid square will be about 14 mm wide in Word. Increase it for larger, more detailed graphs; decrease it to fit more graphs on a page.

Adjust this value *before* generating the VBA script — the preview on screen reflects the final size in Word.

## ⚠️ Implicit Function Support — What Works and What Doesn't

The tool handles two categories of equations differently:

### ✅ Explicit functions `y = f(x)`
Always renders smoothly. Examples:
- `y = x^2 - 3x + 1`
- `y = sin(x)`
- `y = log(2, x)`
- `y = sqrt(x)`

### ✅ Implicit equations *linear in y*
These are automatically rearranged into `y = h(x)` internally, so they render just as smoothly as explicit functions. Examples:
- `2x + 3y + 7 = 0`
- `x^2 + 3x + 4 - y = 0`
- `y - sin(x) = 0`

### ⚠️ General implicit equations (nonlinear in y)
These fall back to a numerical contour-tracing algorithm (marching squares). They work, but come with caveats:

**Examples that render well:**
- `x^2 + y^2 = 9` (circles)
- `x^2/9 + y^2/4 = 1` (ellipses)
- `y^2 = x` (parabolas opening sideways)

**Known limitations:**
- **Sharp corners or cusps** may appear slightly rounded.
- **Self-intersections** (e.g. `y^2 = x^2(x+1)`) may show rendering artifacts near the crossing point.
- **Very thin or near-tangent regions** may break into small disconnected segments.
- **Equations whose solution is an entire axis** (e.g. `x^2 = 0`, `x·y = 0`) are detected and the axis-lying portion is hidden to avoid spurious tick marks.
- **Highly oscillatory curves** (e.g. `sin(x) = sin(y)`) may miss features between sample points.

**Tip:** If an implicit curve looks rough, try rewriting it to be linear in `y` whenever possible. For instance, `y^2 - 2y + x = 0` renders better as two curves: `y = 1 + sqrt(1 - x)` and `y = 1 - sqrt(1 - x)`.

## 📝 Expression Syntax

- **Arithmetic**: `+`, `-`, `*`, `/`, `^`
- **Implicit multiplication**: `2x`, `3xy` (no `*` needed)
- **Functions**: `sin`, `cos`, `tan`, `sqrt`, `abs`, `exp`, `ln`
- **Logarithms**: `log(base, value)` — e.g. `log(2, x)` for log base 2
- **Constants**: `pi`, `e`

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

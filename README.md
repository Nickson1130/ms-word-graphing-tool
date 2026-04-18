# 🧮 MathGraph VBA Designer (DSE Style)

**Generate professional, DSE-style mathematical graphs for Microsoft Word.**

Built for HKDSE educators and students to create vector-quality graphs directly in Word using VBA.

## ✨ Features

- **DSE Style**: Ticks on intercepts, italicized labels, and clean coordinate systems.
- **Implicit & Explicit**: Plot `y = f(x)` or implicit functions like `x^2 + y^2 = 9`.
- **Custom Labels**: Add italicized labels with **Cross (×)** or **Dot (●)** markers.
- **VBA Export**: Generate editable Word shapes instead of blurry screenshots.

## 🛠️ How to Use

1. **Design**: Type your equations and customize axis ranges.
2. **Export**: Click **Generate VBA Script** and copy the code.
3. **In Word**: Open **Developer > Visual Basic**, go to `Insert > Module`, and paste the code.
4. **Run**: Press **F5** or run the `DrawGraph` macro.

## 🚀 Deployment (Hosting)

### Option 1: GitHub Pages (Free) 
1. Go to repository **Settings > Pages**.
2. Set **Build and deployment > Source** to `GitHub Actions`.
3. Use a static site deployment workflow. (Note: I have pre-configured `vite.config.ts` for this).

### Option 2: Vercel / Netlify (Recommended)
1. Sign up for [Vercel](https://vercel.com/) or [Netlify](https://www.netlify.com/).
2. Connect your GitHub account and select this repository.
3. They will automatically detect Vite. Use these settings:
   - **Build Command**: `npm run build`
   - **Output Directory**: `dist`
4. Click **Deploy**. Your app will be live on a custom URL.

## 🏁 Quick Start (Self-Hosting)

```bash
npm install
npm run dev
```
Open `http://localhost:3000`.

## 📜 License
MIT

import React, { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Settings2, 
  Code2, 
  Maximize2, 
  Download, 
  RefreshCcw, 
  Plus, 
  Trash2,
  Info,
  FunctionSquare,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { create, all } from 'mathjs';
import { cn } from './lib/utils';

// Enable implicit multiplication globally by creating a custom instance
const math = create(all, { 
  implicit: 'multiply' 
} as any);

// --- Helper Functions ---
function mergeSegments(segments: { x: number; y: number }[][], threshold = 0.05): { x: number; y: number }[][] {
  if (segments.length === 0) return [];
  const result: { x: number; y: number }[][] = [];
  const points = [...segments];

  while (points.length > 0) {
    let current = [...points.shift()!];
    let found = true;
    while (found) {
      found = false;
      const tail = current[current.length - 1];
      const head = current[0];

      for (let i = 0; i < points.length; i++) {
        const next = points[i];
        const nHead = next[0];
        const nTail = next[next.length - 1];

        // Tail to Head
        if (Math.hypot(tail.x - nHead.x, tail.y - nHead.y) < threshold) {
          current.push(...next.slice(1));
          points.splice(i, 1);
          found = true;
          break;
        }
        // Tail to Tail
        if (Math.hypot(tail.x - nTail.x, tail.y - nTail.y) < threshold) {
          current.push(...[...next].reverse().slice(1));
          points.splice(i, 1);
          found = true;
          break;
        }
        // Head to Tail
        if (Math.hypot(head.x - nTail.x, head.y - nTail.y) < threshold) {
          current.unshift(...next.slice(0, -1));
          points.splice(i, 1);
          found = true;
          break;
        }
        // Head to Head
        if (Math.hypot(head.x - nHead.x, head.y - nHead.y) < threshold) {
          current.unshift(...[...next].reverse().slice(0, -1));
          points.splice(i, 1);
          found = true;
          break;
        }
      }
    }
    result.push(current);
  }
  return result;
}

// --- Types ---
interface CustomLabel {
  id: string;
  x: number;
  y: number;
  text: string;
  symbol?: 'dot' | 'cross';
}

interface Interval {
  id: string;
  xMin: string;
  xMax: string;
  yMin: string;
  yMax: string;
  useCustomDomain: boolean;
  useCustomRange: boolean;
}

interface Equation {
  id: string;
  expression: string;
  lineWidth: number;
  dashStyle: string; // msoLineDashStyle value
  color: string; // hex color
  intervals: Interval[];
  useCompression?: boolean; // opt-in: apply axis compression bands to this curve
}

interface AxisBreakBand {
  id: string;
  axis: 'x' | 'y';
  min: string;          // interval lower bound (data coords)
  max: string;          // interval upper bound (data coords)
  compression: number;  // 0..1 — fraction of original band width kept
}

interface GraphSettings {
  unitSize: number; 
  xMin: string;
  xMax: string;
  yMin: string;
  yMax: string;
  xAxisLabel: string;
  yAxisLabel: string;
  showOrigin: boolean;
  equations: Equation[];
  samplingStep: number;
  customLabels: CustomLabel[];
  showXIntercepts: boolean;
  showYIntercepts: boolean;
  showTicks: boolean;
  showXAxis: boolean;
  showYAxis: boolean;
  showXLabel: boolean;
  showYLabel: boolean;
  customXTicks: string; // comma separated
  customYTicks: string; // comma separated
  arrowStyle: string; // msoArrowheadStyle value
  tickWidth: number;
  axisLabelFontSize: number;   // font size for x, y, O labels
  customLabelFontSize: number; // font size for custom labels
  showGrid: boolean;
  gridSpacingX: number;   // units between vertical gridlines
  gridSpacingY: number;   // units between horizontal gridlines
  gridColor: string;      // hex
  gridOpacity: number;    // 0–1
  gridLineWidth: number;  // line weight in points
  xScaleMode: string;     // 'linear-0.5' | 'linear-1' | 'linear-2' | 'linear-3' | 'linear-5' | 'linear-10' | 'log10' | 'log2' | 'ln' | 'custom'
  yScaleMode: string;     // same as above
  xScaleMultiplier: number; // free-form multiplier (used when mode is 'linear-custom')
  yScaleMultiplier: number;
  xScaleCustom: string;     // custom expression, e.g. 'sqrt(x)', '1/x'
  yScaleCustom: string;
  axisBreakEnabled: boolean;          // master switch — turns all bands on/off
  axisBreakBands: AxisBreakBand[];    // disjoint compression bands; any mix of X and Y
}

// --- Constants ---
const ARROW_STYLES = [
  { label: 'Triangle', value: 'msoArrowheadTriangle' },
  { label: 'Stealth', value: 'msoArrowheadStealth' },
  { label: 'Open', value: 'msoArrowheadOpen' },
  { label: 'Diamond', value: 'msoArrowheadDiamond' },
  { label: 'Oval', value: 'msoArrowheadOval' },
  { label: 'None', value: 'msoArrowheadNone' },
];

const DASH_STYLES = [
  { label: 'Solid', value: 'msoLineSolid', dash: '' },
  { label: 'Square Dot', value: 'msoLineSquareDot', dash: '2,2' },
  { label: 'Round Dot', value: 'msoLineRoundDot', dash: '1,3' },
  { label: 'Dash', value: 'msoLineDash', dash: '4,4' },
  { label: 'Dash Dot', value: 'msoLineDashDot', dash: '4,2,1,2' },
  { label: 'Long Dash', value: 'msoLineLongDash', dash: '8,4' },
];

const SCALE_OPTIONS = [
  { label: 'Linear 0.5×', value: 'linear-0.5' },
  { label: 'Linear 1×',   value: 'linear-1' },
  { label: 'Linear 2×',   value: 'linear-2' },
  { label: 'Linear 3×',   value: 'linear-3' },
  { label: 'Linear 5×',   value: 'linear-5' },
  { label: 'Linear 10×',  value: 'linear-10' },
  { label: 'Linear custom×', value: 'linear-custom' },
  { label: 'Log₁₀',  value: 'log10' },
  { label: 'Log₂',   value: 'log2' },
  { label: 'Ln',     value: 'ln' },
  { label: 'Custom f(x)', value: 'custom' },
];

const DEFAULT_SETTINGS: GraphSettings = {
  unitSize: 40,
  xMin: '-5',
  xMax: '5',
  yMin: '-5',
  yMax: '5',
  xAxisLabel: 'x',
  yAxisLabel: 'y',
  showOrigin: true,
  equations: [
    {
      id: '1',
      expression: 'y = x^2',
      lineWidth: 0.75,
      dashStyle: 'msoLineSolid',
      color: '#000000',
      useCompression: false,
      intervals: [
        {
          id: 'int-1',
          xMin: '-5',
          xMax: '5',
          yMin: '-5',
          yMax: '5',
          useCustomDomain: false,
          useCustomRange: false
        }
      ]
    }
  ],
  samplingStep: 0.01,
  customLabels: [],
  showXIntercepts: true,
  showYIntercepts: true,
  showTicks: false,
  showXAxis: true,
  showYAxis: true,
  showXLabel: true,
  showYLabel: true,
  customXTicks: '',
  customYTicks: '',
  arrowStyle: 'msoArrowheadOpen',
  tickWidth: 0.75,
  axisLabelFontSize: 12,
  customLabelFontSize: 10,
  showGrid: false,
  gridSpacingX: 1,
  gridSpacingY: 1,
  gridColor: '#cccccc',
  gridOpacity: 0.5,
  gridLineWidth: 0.25,
  xScaleMode: 'linear-1',
  yScaleMode: 'linear-1',
  xScaleMultiplier: 1,
  yScaleMultiplier: 1,
  xScaleCustom: 'x',
  yScaleCustom: 'y',
  axisBreakEnabled: false,
  axisBreakBands: [],
};

export default function MathGraphDesigner() {
  const [settings, setSettings] = useState<GraphSettings>(DEFAULT_SETTINGS);
  const [vbaVisible, setVbaVisible] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // --- Derived Data ---
  const nXMin = parseFloat(settings.xMin) || 0;
  const nXMax = parseFloat(settings.xMax) || 0;
  const nYMin = parseFloat(settings.yMin) || 0;
  const nYMax = parseFloat(settings.yMax) || 0;

  // --- Axis transform system ---
  // Each axis has a transform t(v) that maps a data value to a "scaled value"
  // in the same units. Canvas coordinate = originX + tx(x) * unitSize.
  // Returns null for invalid values (e.g. log of non-positive).
  const buildTransform = (mode: string, multiplier: number, customExpr: string, varName: 'x' | 'y'): (v: number) => number | null => {
    if (mode.startsWith('linear-')) {
      let k = 1;
      if (mode === 'linear-custom') k = multiplier || 1;
      else k = parseFloat(mode.slice(7)) || 1;
      return (v: number) => v * k;
    }
    if (mode === 'log10') return (v: number) => (v > 0 ? Math.log10(v) : null);
    if (mode === 'log2')  return (v: number) => (v > 0 ? Math.log2(v)  : null);
    if (mode === 'ln')    return (v: number) => (v > 0 ? Math.log(v)   : null);
    if (mode === 'custom') {
      try {
        const compiled = math.parse(customExpr || varName).compile();
        return (v: number) => {
          try {
            const r = compiled.evaluate({ [varName]: v });
            if (typeof r === 'number' && isFinite(r)) return r;
            return null;
          } catch { return null; }
        };
      } catch {
        return (v: number) => v; // fallback to identity if expression is invalid
      }
    }
    return (v: number) => v;
  };

  const txBase = useMemo(
    () => buildTransform(settings.xScaleMode, settings.xScaleMultiplier, settings.xScaleCustom, 'x'),
    [settings.xScaleMode, settings.xScaleMultiplier, settings.xScaleCustom]
  );
  const tyBase = useMemo(
    () => buildTransform(settings.yScaleMode, settings.yScaleMultiplier, settings.yScaleCustom, 'y'),
    [settings.yScaleMode, settings.yScaleMultiplier, settings.yScaleCustom]
  );

  // --- Axis compression bands ---
  // Compute the per-axis list of bands in scaled space (after buildTransform).
  // Each band carries its post-scale interval and compression factor; the
  // curve transforms sum displacements across all bands on the relevant axis,
  // so multiple disjoint bands on X and/or Y compose naturally.
  type ScaledBand = { a: number; b: number; c: number };
  const scaledXBands = useMemo<ScaledBand[]>(() => {
    if (!settings.axisBreakEnabled) return [];
    const out: ScaledBand[] = [];
    for (const band of settings.axisBreakBands) {
      if (band.axis !== 'x') continue;
      const numMin = parseFloat(band.min);
      const numMax = parseFloat(band.max);
      if (!isFinite(numMin) || !isFinite(numMax) || numMax <= numMin) continue;
      const sa = txBase(numMin);
      const sb = txBase(numMax);
      if (sa === null || sb === null || sb <= sa) continue;
      out.push({ a: sa, b: sb, c: Math.max(0, band.compression) });
    }
    return out;
  }, [settings.axisBreakEnabled, settings.axisBreakBands, txBase]);

  // Y-axis bands: Range is an *x-interval* (not y!) where the curve's y
  // values are multiplied by the scale factor. So Y-band {min:1, max:5, c:1.6}
  // means "for curve points with x in [1, 5], stretch y by 1.6×."
  const scaledYBands = useMemo<ScaledBand[]>(() => {
    if (!settings.axisBreakEnabled) return [];
    const out: ScaledBand[] = [];
    for (const band of settings.axisBreakBands) {
      if (band.axis !== 'y') continue;
      const numMin = parseFloat(band.min);
      const numMax = parseFloat(band.max);
      if (!isFinite(numMin) || !isFinite(numMax) || numMax <= numMin) continue;
      const sa = txBase(numMin);
      const sb = txBase(numMax);
      if (sa === null || sb === null || sb <= sa) continue;
      out.push({ a: sa, b: sb, c: Math.max(0, band.compression) });
    }
    return out;
  }, [settings.axisBreakEnabled, settings.axisBreakBands, txBase]);

  // Gradual band compression / stretch that resizes the band's image width.
  //
  // Input band [a, b] (width h) is mapped to an image of width c·h centered on
  // the same midpoint. Outside the band, the curve is shifted by h·(1−c)/2 so
  // it joins continuously with the band image (inward for c<1, outward for c>1).
  //
  // Inside the band we use cos² ramps of width w·h at each boundary, with a
  // plateau of slope p in between, picked so the derivative is exactly 1 at
  // the band edges (no kink) and averages to c overall:
  //   w = min(c, 0.5),   p = (c − w)/(1 − w)
  //
  // For c=1 → identity. For c=0 the band collapses to a single point at the
  // band midpoint. For c>1 the band stretches (plateau slope p>1) — the rest
  // of the curve is pushed outward by h(c−1)/2 on each side.
  const compressGradual = (s: number, a: number, b: number, c: number): number => {
    const h = b - a;
    const cc = Math.max(0, c);
    const halfShift = h * (1 - cc) / 2;
    if (s <= a) return s + halfShift;
    if (s >= b) return s - halfShift;
    if (cc === 0) return a + h / 2; // entire band collapses to its midpoint

    const tau = (s - a) / h;
    const w = Math.min(cc, 0.5);
    const p = (cc - w) / (1 - w);

    let G: number;
    if (tau <= w) {
      G = p * tau + (1 - p) * (tau / 2 + (w / (2 * Math.PI)) * Math.sin(Math.PI * tau / w));
    } else if (tau >= 1 - w) {
      const tauR = 1 - tau;
      const GR = p * tauR + (1 - p) * (tauR / 2 + (w / (2 * Math.PI)) * Math.sin(Math.PI * tauR / w));
      G = cc - GR;
    } else {
      G = w * (1 + p) / 2 + p * (tau - w);
    }
    return a + halfShift + h * G;
  };

  // tx/ty are pure axis transforms — no compression applied here so axes,
  // grids, ticks and non-opted-in curves stay untouched.
  const tx = txBase;
  const ty = tyBase;

  // Per-curve compression — only opt-in curves (useCompression === true)
  // are affected.
  //
  // X-bands compress/stretch the x-coordinate inside the band; displacements
  // sum across disjoint bands.
  const compressXForCurve = (s: number): number => {
    if (scaledXBands.length === 0) return s;
    let total = 0;
    for (const { a, b, c } of scaledXBands) total += compressGradual(s, a, b, c) - s;
    return s + total;
  };

  // Y-band scale factor at a given x: smooth raised-cosine bump going 1
  // outside the band → c at the middle → 1 again, with slope 0 at the
  // band edges so the curve transitions in/out smoothly.
  const yFactorAtX = (sx: number, a: number, b: number, c: number): number => {
    if (sx <= a || sx >= b) return 1;
    const tau = (sx - a) / (b - a);
    const window = 0.5 - 0.5 * Math.cos(2 * Math.PI * tau);
    return 1 + (Math.max(0, c) - 1) * window;
  };

  // For each Y-band whose x-interval contains the curve's (pre-x-compression)
  // scaled x, multiply the curve's scaled y by the band's local factor.
  const compressYForCurve = (sy: number, sx: number): number => {
    if (scaledYBands.length === 0) return sy;
    let factor = 1;
    for (const { a, b, c } of scaledYBands) factor *= yFactorAtX(sx, a, b, c);
    return sy * factor;
  };

  // Is a log-like scale active? Used to hide origin label etc.
  const xIsLog = ['log10', 'log2', 'ln'].includes(settings.xScaleMode);
  const yIsLog = ['log10', 'log2', 'ln'].includes(settings.yScaleMode);

  // Transformed min/max for canvas sizing & origin placement
  const txMin = tx(nXMin); const txMax = tx(nXMax);
  const tyMin = ty(nYMin); const tyMax = ty(nYMax);
  const sxMin = txMin === null ? nXMin : txMin;
  const sxMax = txMax === null ? nXMax : txMax;
  const syMin = tyMin === null ? nYMin : tyMin;
  const syMax = tyMax === null ? nYMax : tyMax;

  // Override canvas dimensions & origin with transformed values
  const scaledCanvasWidth  = (sxMax - sxMin) * settings.unitSize;
  const scaledCanvasHeight = (syMax - syMin) * settings.unitSize;
  const scaledOriginX = -sxMin * settings.unitSize;
  const scaledOriginY =  syMax * settings.unitSize;

  // Coordinate conversion — takes data coords, applies scale, returns pixel coords.
  // Returns null for any coordinate that cannot be mapped (e.g. log of 0).
  const toPoints = (x: number, y: number): { x: number; y: number } => {
    const xs = tx(x); const ys = ty(y);
    const ex = xs === null ? x : xs;
    const ey = ys === null ? y : ys;
    return {
      x: scaledOriginX + ex * settings.unitSize,
      y: scaledOriginY - ey * settings.unitSize,
    };
  };
  // Version that returns null if either coordinate is invalid (for curve rendering)
  const toPointsStrict = (x: number, y: number): { x: number; y: number } | null => {
    const xs = tx(x); const ys = ty(y);
    if (xs === null || ys === null) return null;
    return {
      x: scaledOriginX + xs * settings.unitSize,
      y: scaledOriginY - ys * settings.unitSize,
    };
  };

  // Aliases for backward compatibility with existing UI code
  const canvasWidth = scaledCanvasWidth;
  const canvasHeight = scaledCanvasHeight;
  const originX = scaledOriginX;
  const originY = scaledOriginY;

  const hexToVbaRgb = (hex: string) => {
    const r = parseInt(hex.slice(1, 3), 16) || 0;
    const g = parseInt(hex.slice(3, 5), 16) || 0;
    const b = parseInt(hex.slice(5, 7), 16) || 0;
    return `RGB(${r}, ${g}, ${b})`;
  };

  // Custom math context for helpers like log(base, value)
  const mathScope = useMemo(() => ({
    // log(base, x) helper - user specifically requested log(a, b) style
    log: (a: any, b: any) => {
      if (b === undefined) return Math.log(a); // Natural log if 1 arg
      return Math.log(b) / Math.log(a); // log_a(b) if 2 args
    },
    ln: (x: number) => Math.log(x),
  }), []);

  // --- Helper: try to rearrange f(x,y)=g(x,y) into y=h(x) when y appears linearly ---
  // Returns an evaluator { evaluate: ({x}) => y } or null if not linear in y.
  const tryResolveForY = (lhs: string, rhs: string): { evaluate: (scope: { x: number }) => number } | null => {
    let compiled: ReturnType<ReturnType<typeof math.parse>['compile']> | any;
    try {
      compiled = math.parse(`(${lhs}) - (${rhs})`).compile();
    } catch {
      return null;
    }

    // F(x,y) = A(x)*y + B(x) if linear in y. Test linearity at several x values
    // by checking F(x, y2) - F(x, y1) is proportional to (y2 - y1) and that
    // the second difference in y is zero (i.e. y^2 term is absent).
    const testXs = [-1.3, 0.7, 1.5, 2.4];
    for (const tx of testXs) {
      try {
        const f0 = compiled.evaluate({ ...mathScope, x: tx, y: 0 });
        const f1 = compiled.evaluate({ ...mathScope, x: tx, y: 1 });
        const f2 = compiled.evaluate({ ...mathScope, x: tx, y: 2 });
        if (!isFinite(f0) || !isFinite(f1) || !isFinite(f2)) return null;

        const a = f1 - f0;                  // coefficient of y
        const predicted_f2 = 2 * a + f0;    // what F(x,2) should be if linear
        if (Math.abs(f2 - predicted_f2) > 1e-7) return null; // nonlinear in y
        if (Math.abs(a) < 1e-9) return null; // y doesn't appear (or cancels)
      } catch {
        return null; // couldn't evaluate at this test point (e.g. div by zero)
      }
    }

    // Confirmed linear-in-y. Build runtime evaluator: y = -B(x)/A(x)
    return {
      evaluate: (scope: { x: number }) => {
        try {
          const b = compiled.evaluate({ ...mathScope, ...scope, y: 0 });
          const a = compiled.evaluate({ ...mathScope, ...scope, y: 1 }) - b;
          if (Math.abs(a) < 1e-12) return NaN;
          return -b / a;
        } catch {
          return NaN;
        }
      }
    };
  };

  // Curve sampling for multiple equations
  const allCurvePoints = useMemo(() => {
    return settings.equations.map(eq => {
      try {
        let expression = eq.expression.trim();
        if (!expression) {
          return { 
            id: eq.id, 
            segments: [], 
            lineWidth: eq.lineWidth, 
            dashStyle: eq.dashStyle,
            color: eq.color || '#000000'
          };
        }
        
        // Check if it's a relation or a simple y=f(x)
        const parts = expression.split('=');
        const isSimpleFunction = parts.length === 2 && parts[0].trim() === 'y';
        
        const allSegments: { x: number; y: number }[][] = [];

        if (isSimpleFunction) {
          // Fast Path: y = f(x)
          const fnExpr = parts[1].trim();
          const compiled = math.parse(fnExpr).compile();
          
          eq.intervals.forEach(interval => {
            const iXMin = interval.useCustomDomain ? (parseFloat(interval.xMin) || 0) : nXMin;
            const iXMax = interval.useCustomDomain ? (parseFloat(interval.xMax) || 0) : nXMax;
            const iYMin = interval.useCustomRange ? (parseFloat(interval.yMin) || 0) : nYMin;
            const iYMax = interval.useCustomRange ? (parseFloat(interval.yMax) || 0) : nYMax;

            const points: { x: number; y: number }[] = [];
            const samples = Math.max(2, Math.ceil((iXMax - iXMin) / Math.max(settings.samplingStep, 0.0001)) + 1);
            const step = (iXMax - iXMin) / (samples - 1);

            for (let i = 0; i < samples; i++) {
              const x = iXMin + i * step;
              try {
                const y = compiled.evaluate({ ...mathScope, x });
                if (typeof y === 'number' && isFinite(y)) {
                  if (y >= (iYMin - 0.0001) && y <= (iYMax + 0.0001)) {
                    points.push({ x, y });
                  } else if (points.length > 0) {
                    allSegments.push([...points]);
                    points.length = 0;
                  }
                } else if (points.length > 0) {
                  allSegments.push([...points]);
                  points.length = 0;
                }
              } catch (e) { }
            }
            if (points.length > 0) allSegments.push([...points]);
          });
        } else {
          // Relation Path: f(x, y) = 0
          const lhs = parts[0] || '0';
          const rhs = parts.length >= 2 ? (parts[1] || '0') : '0';

          // Optimisation: if y appears linearly, rearrange to y=h(x) and use fast path
          const resolvedForY = tryResolveForY(lhs, rhs);
          if (resolvedForY) {
            eq.intervals.forEach(interval => {
              const iXMin = interval.useCustomDomain ? (parseFloat(interval.xMin) || 0) : nXMin;
              const iXMax = interval.useCustomDomain ? (parseFloat(interval.xMax) || 0) : nXMax;
              const iYMin = interval.useCustomRange ? (parseFloat(interval.yMin) || 0) : nYMin;
              const iYMax = interval.useCustomRange ? (parseFloat(interval.yMax) || 0) : nYMax;

              const points: { x: number; y: number }[] = [];
              const samples = Math.max(2, Math.ceil((iXMax - iXMin) / Math.max(settings.samplingStep, 0.0001)) + 1);
              const step = (iXMax - iXMin) / (samples - 1);

              for (let i = 0; i < samples; i++) {
                const x = iXMin + i * step;
                try {
                  const y = resolvedForY.evaluate({ x });
                  if (typeof y === 'number' && isFinite(y)) {
                    if (y >= (iYMin - 0.0001) && y <= (iYMax + 0.0001)) {
                      points.push({ x, y });
                    } else if (points.length > 0) {
                      allSegments.push([...points]);
                      points.length = 0;
                    }
                  } else if (points.length > 0) {
                    allSegments.push([...points]);
                    points.length = 0;
                  }
                } catch {}
              }
              if (points.length > 0) allSegments.push([...points]);
            });
          } else {
          
          eq.intervals.forEach(interval => {
            const iXMin = interval.useCustomDomain ? (parseFloat(interval.xMin) || 0) : nXMin;
            const iXMax = interval.useCustomDomain ? (parseFloat(interval.xMax) || 0) : nXMax;
            const iYMin = interval.useCustomRange ? (parseFloat(interval.yMin) || 0) : nYMin;
            const iYMax = interval.useCustomRange ? (parseFloat(interval.yMax) || 0) : nYMax;

            // Use a high-resolution grid for smooth implicit plotting
            const gridX = 400;
            const gridY = 400;
            const compiled = math.parse(`(${lhs}) - (${rhs})`).compile();

            // Sanity check: reject degenerate equations (e.g. y-y=0, 0=0, x-x=1)
            // by sampling F at a few points. If F is constant across all samples,
            // the equation is either always true (infinite solutions → would freeze
            // the browser) or always false (no solutions).
            let degenerate = true;
            let firstVal: number | null = null;
            const probeXs = [iXMin + (iXMax - iXMin) * 0.17, iXMin + (iXMax - iXMin) * 0.53, iXMin + (iXMax - iXMin) * 0.89];
            const probeYs = [iYMin + (iYMax - iYMin) * 0.23, iYMin + (iYMax - iYMin) * 0.61, iYMin + (iYMax - iYMin) * 0.97];
            for (const px of probeXs) {
              for (const py of probeYs) {
                try {
                  const v = compiled.evaluate({ ...mathScope, x: px, y: py });
                  if (typeof v !== 'number' || !isFinite(v)) continue;
                  if (firstVal === null) { firstVal = v; continue; }
                  if (Math.abs(v - firstVal) > 1e-9) { degenerate = false; break; }
                } catch {}
              }
              if (!degenerate) break;
            }
            if (degenerate) return; // skip this interval — no valid curve to draw

            const dx = (iXMax - iXMin) / Math.max(gridX, 1);
            const dy = (iYMax - iYMin) / Math.max(gridY, 1);
            const implicitSegments: { x: number; y: number }[][] = [];

            const values: number[][] = [];
            for (let i = 0; i <= gridX; i++) {
              values[i] = [];
              const x = iXMin + i * dx;
              for (let j = 0; j <= gridY; j++) {
                const y = iYMin + j * dy;
                try {
                  const val = compiled.evaluate({ ...mathScope, x, y });
                  values[i][j] = typeof val === 'number' ? val : NaN;
                } catch {
                  values[i][j] = NaN;
                }
              }
            }

            // Simple edge crossing segments
            for (let i = 0; i < gridX; i++) {
              for (let j = 0; j < gridY; j++) {
                const v1 = values[i][j];
                const v2 = values[i+1][j];
                const v3 = values[i+1][j+1];
                const v4 = values[i][j+1];

                if (isNaN(v1) || isNaN(v2) || isNaN(v3) || isNaN(v4)) continue;

                // Check edges and create micro-segments
                const edges = [[v1, v2, i, j, i+1, j], [v2, v3, i+1, j, i+1, j+1], [v3, v4, i+1, j+1, i, j+1], [v4, v1, i, j+1, i, j]];
                const crossings: {x: number, y: number}[] = [];
                
                edges.forEach(([va, vb, ax, ay, bx, by]) => {
                  if (va * vb <= 0) {
                    const t = Math.abs(va) + Math.abs(vb) < 0.000001 ? 0.5 : Math.abs(va) / (Math.abs(va) + Math.abs(vb));
                    crossings.push({
                      x: (iXMin + ax * dx) + t * (bx - ax) * dx,
                      y: (iYMin + ay * dy) + t * (by - ay) * dy
                    });
                  }
                });

                if (crossings.length >= 2) {
                  implicitSegments.push([crossings[0], crossings[1]]);
                }
              }
            }
            // Merge fragmented segments into continuous paths
            allSegments.push(...mergeSegments(implicitSegments, Math.max(dx, dy) * 1.5));
          });
          } // end else (marching squares)
        }

        return { 
          id: eq.id, 
          segments: allSegments, 
          lineWidth: eq.lineWidth, 
          dashStyle: eq.dashStyle,
          color: eq.color || '#000000'
        };
      } catch (e) {
        return { 
          id: eq.id, 
          segments: [], 
          lineWidth: eq.lineWidth, 
          dashStyle: eq.dashStyle,
          color: eq.color || '#000000'
        };
      }
    });
  }, [settings.equations, nXMin, nXMax, nYMin, nYMax, mathScope, settings.samplingStep]);

  // Combined Intercepts for all curves
  const intercepts = useMemo(() => {
    if (!settings.showXIntercepts && !settings.showYIntercepts) return [];
    
    let allIntercepts: (CustomLabel & { isAuto: boolean; axis: 'x' | 'y'; curveIdx: number })[] = [];
    allCurvePoints.forEach((curve, cIdx) => {
      curve.segments.forEach((segment, sIdx) => {
        if (segment.length < 2) return;
        
        const originalExpr = settings.equations.find(e => e.id === curve.id)?.expression.trim() || '';
        if (!originalExpr) return;
        
        // Handle both simple y=f(x) and implicit f(x,y)=g(x,y)
        const parts = originalExpr.split('=');
        const isSimpleFunction = parts.length === 2 && parts[0].trim() === 'y';
        
        try {
          if (isSimpleFunction && settings.showYIntercepts) {
            const expression = parts[1].trim();
            const compiled = math.parse(expression).compile();
            // Y-Intercept
            const y0 = compiled.evaluate({ ...mathScope, x: 0 });
            if (isFinite(Number(y0))) {
              // Check if x=0 is in this segment
              const xInRange = segment.some(p => Math.abs(p.x) < 0.02);
              if (xInRange && Math.abs(Number(y0)) > 0.001) {
                 allIntercepts.push({
                   id: `y-int-${cIdx}-${sIdx}`,
                   x: 0,
                   y: Number(Number(y0).toFixed(2)),
                   text: '', // No numbers for auto intercepts
                   symbol: 'dot',
                   isAuto: true,
                   axis: 'y',
                   curveIdx: cIdx
                 });
              }
            }
          }
        } catch {}

        // Generic X and Y Intercepts from segments
        // Geometrically detect if segment lies entirely on an axis.
        // Dynamic threshold catches x^2=0, x^3=0, x^2, x^3 etc.
        const axisThreshold = Math.max((nXMax - nXMin) / 200, (nYMax - nYMin) / 200, 0.05);
        const segLiesOnXAxis = segment.every(p => Math.abs(p.y) < axisThreshold);
        const segLiesOnYAxis = segment.every(p => Math.abs(p.x) < axisThreshold);

        if (!segLiesOnXAxis && !segLiesOnYAxis) {
        for (let i = 0; i < segment.length - 1; i++) {
          const p1 = segment[i];
          const p2 = segment[i+1];

          // X-Intercept (y crosses 0)
          if (settings.showXIntercepts && ((p1.y >= 0 && p2.y <= 0) || (p1.y <= 0 && p2.y >= 0))) {
            let xInt;
            const dy = p1.y - p2.y;
            if (Math.abs(dy) < 0.000001) {
              xInt = p1.x;
            } else {
              xInt = p1.x - (p1.y * (p2.x - p1.x)) / (p2.y - p1.y);
            }
            if (Math.abs(xInt) > 0.001) {
              allIntercepts.push({
                id: `x-int-${cIdx}-${sIdx}-${i}`,
                x: Number(xInt.toFixed(2)),
                y: 0,
                text: '',
                symbol: 'dot',
                isAuto: true,
                axis: 'x',
                curveIdx: cIdx
              });
            }
          }

          // Y-Intercept (x crosses 0)
          if (settings.showYIntercepts && ((p1.x >= 0 && p2.x <= 0) || (p1.x <= 0 && p2.x >= 0))) {
            let yInt;
            const dx = p1.x - p2.x;
            if (Math.abs(dx) < 0.000001) {
              yInt = p1.y;
            } else {
              yInt = p1.y - (p1.x * (p2.y - p1.y)) / (p2.x - p1.x);
            }
            if (Math.abs(yInt) > 0.001) {
              allIntercepts.push({
                id: `y-seg-int-${cIdx}-${sIdx}-${i}`,
                x: 0,
                y: Number(yInt.toFixed(2)),
                text: '',
                symbol: 'dot',
                isAuto: true,
                axis: 'y',
                curveIdx: cIdx
              });
            }
          }
        }
        } // end if (!segLiesOnXAxis && !segLiesOnYAxis)
      });
    });

    // Deduplicate and filter out start/end points
    return allIntercepts
      .filter((item, index, self) =>
        index === self.findIndex((t) => (
          Math.abs(t.x - item.x) < 0.05 && Math.abs(t.y - item.y) < 0.05
        ))
      )
      .filter(i => {
        // Hide if at or outside x-axis boundaries (approx)
        if (Math.abs(i.y) < 0.001) {
          if (!(i.x > nXMin + 0.01 && i.x < nXMax - 0.01)) return false;
        }
        // Hide if at or outside y-axis boundaries (approx)
        if (Math.abs(i.x) < 0.001) {
          if (!(i.y > nYMin + 0.01 && i.y < nYMax - 0.01)) return false;
        }
        return true;
      });
  }, [allCurvePoints, settings.showXIntercepts, settings.showYIntercepts, settings.equations, nXMin, nXMax, nYMin, nYMax]);

  const xTicks = useMemo(() => {
    if (settings.customXTicks.trim()) {
      return settings.customXTicks.split(',').map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
    }
    return Array.from({ length: Math.floor(nXMax - nXMin) + 1 }, (_, i) => Math.ceil(nXMin) + i)
      .filter(x => x !== 0 && x > (nXMin + 0.01) && x < (nXMax - 0.01));
  }, [settings.customXTicks, nXMin, nXMax]);

  const yTicks = useMemo(() => {
    if (settings.customYTicks.trim()) {
      return settings.customYTicks.split(',').map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
    }
    return Array.from({ length: Math.floor(nYMax - nYMin) + 1 }, (_, i) => Math.ceil(nYMin) + i)
      .filter(y => y !== 0 && y > (nYMin + 0.01) && y < (nYMax - 0.01));
  }, [settings.customYTicks, nYMin, nYMax]);

  // Gridline positions (in mathematical coordinates)
  const gridLines = useMemo(() => {
    if (!settings.showGrid) return { xs: [], ys: [] };
    const sx = Math.max(0.01, settings.gridSpacingX);
    const sy = Math.max(0.01, settings.gridSpacingY);
    const xs: number[] = [];
    const ys: number[] = [];
    // anchor at 0 so gridlines pass through the origin
    const xStart = Math.ceil(nXMin / sx) * sx;
    for (let x = xStart; x <= nXMax + 1e-9; x += sx) {
      // skip the axis line itself (x=0) since the axis draws it
      if (Math.abs(x) > 1e-9) xs.push(Number(x.toFixed(6)));
    }
    const yStart = Math.ceil(nYMin / sy) * sy;
    for (let y = yStart; y <= nYMax + 1e-9; y += sy) {
      if (Math.abs(y) > 1e-9) ys.push(Number(y.toFixed(6)));
    }
    return { xs, ys };
  }, [settings.showGrid, settings.gridSpacingX, settings.gridSpacingY, nXMin, nXMax, nYMin, nYMax]);

  // Map VBA String Constants to Integers for maximum compatibility
  const VBA_CONSTANTS: Record<string, number> = {
    'msoArrowheadNone': 1,
    'msoArrowheadTriangle': 2,
    'msoArrowheadOpen': 3,
    'msoArrowheadStealth': 4,
    'msoArrowheadDiamond': 5,
    'msoArrowheadOval': 6,
    'msoLineSolid': 1,
    'msoLineSquareDot': 2,
    'msoLineRoundDot': 3,
    'msoLineDash': 4,
    'msoLineDashDot': 5,
    'msoLineLongDash': 7,
  };

  // --- VBA Code Generation ---
  const generatedVBA = useMemo(() => {
    const arrowInt = VBA_CONSTANTS[settings.arrowStyle] || 1;
    // Calculate Origin in JS to ensure curve points are injected as exact numbers
    const jsOriginX = 100 + scaledOriginX;
    const jsOriginY = 100 + scaledOriginY;

    // VBA coordinate helpers: precompute scaled positions in JS and inject as numeric
    // literals into the generated VBA. This keeps the VBA simple (no log math inside
    // Word) and ensures that all scales, including custom and log, export correctly.
    const vbaX = (x: number): string => {
      const xs = tx(x);
      if (xs === null) return `${jsOriginX}`; // fallback
      return `(${jsOriginX} + (${xs} * unitSize))`;
    };
    const vbaY = (y: number): string => {
      const ys = ty(y);
      if (ys === null) return `${jsOriginY}`;
      return `(${jsOriginY} - (${ys} * unitSize))`;
    };

    return `
Sub DrawGraph()
    ' Generated by MathGraph Designer - Robust Compatibility Version
    Dim debugStep As String: debugStep = "Initialization"
    On Error GoTo ErrorHandler
    
    Dim doc As Document: Set doc = ActiveDocument
    Dim unitSize As Double: unitSize = ${settings.unitSize}
    Dim xMin As Double: xMin = ${nXMin}: Dim xMax As Double: xMax = ${nXMax}
    Dim yMin As Double: yMin = ${nYMin}: Dim yMax As Double: yMax = ${nYMax}
    
    Dim originX As Double, originY As Double
    originX = ${jsOriginX}: originY = ${jsOriginY}
    
    Dim shpArray() As Variant
    Dim shpCount As Long: shpCount = 0
    
    ' --- 0. Draw Grid ---
    ${settings.showGrid ? `
    debugStep = "Drawing Grid"
    Dim gridLine As Shape
    ${gridLines.xs.map(gx => `
    Set gridLine = doc.Shapes.AddLine(${vbaX(gx)}, ${vbaY(nYMin)}, ${vbaX(gx)}, ${vbaY(nYMax)})
    gridLine.Line.ForeColor.RGB = ${hexToVbaRgb(settings.gridColor)}
    gridLine.Line.Weight = ${settings.gridLineWidth}
    gridLine.Line.Transparency = ${(1 - settings.gridOpacity).toFixed(3)}
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = gridLine.Name`).join('')}
    ${gridLines.ys.map(gy => `
    Set gridLine = doc.Shapes.AddLine(${vbaX(nXMin)}, ${vbaY(gy)}, ${vbaX(nXMax)}, ${vbaY(gy)})
    gridLine.Line.ForeColor.RGB = ${hexToVbaRgb(settings.gridColor)}
    gridLine.Line.Weight = ${settings.gridLineWidth}
    gridLine.Line.Transparency = ${(1 - settings.gridOpacity).toFixed(3)}
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = gridLine.Name`).join('')}
    ` : ''}

    ' --- 1. Draw Axes ---
    debugStep = "Drawing Axes"
    Dim xAxis As Shape, yAxis As Shape
    
    If ${settings.showXAxis ? 'True' : 'False'} Then
        ' Using AddLine instead of AddConnector for maximum compatibility
        Set xAxis = doc.Shapes.AddLine(${vbaX(nXMin)}, ${vbaY(0)}, ${vbaX(nXMax)}, ${vbaY(0)})
        xAxis.Line.EndArrowheadStyle = ${arrowInt}
        xAxis.Line.Weight = 0.75: xAxis.Line.ForeColor.RGB = 0
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = xAxis.Name
    End If

    If ${settings.showYAxis ? 'True' : 'False'} Then
        Set yAxis = doc.Shapes.AddLine(${vbaX(0)}, ${vbaY(nYMin)}, ${vbaX(0)}, ${vbaY(nYMax)})
        yAxis.Line.EndArrowheadStyle = ${arrowInt}
        yAxis.Line.Weight = 0.75: yAxis.Line.ForeColor.RGB = 0
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = yAxis.Name
    End If

    ' --- 2. Draw Ticks & Labels ---
    If ${settings.showTicks ? 'True' : 'False'} Then
        debugStep = "Drawing Ticks and Labels"
        Dim val As Variant, tick As Shape, lbl As Shape
        ' X Ticks
        Dim xTickVals: xTickVals = Array(${xTicks.length > 0 ? xTicks.map(t => vbaX(t)).join(', ') : ''})
        If UBound(xTickVals) >= LBound(xTickVals) Then
            For Each val In xTickVals
                Set tick = doc.Shapes.AddLine(val, originY - 4, val, originY + 4)
                tick.Line.ForeColor.RGB = 0: tick.Line.Weight = ${settings.tickWidth}
                shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name
            Next val
        End If
        
        ' Y Ticks
        Dim yTickVals: yTickVals = Array(${yTicks.length > 0 ? yTicks.map(t => vbaY(t)).join(', ') : ''})
        If UBound(yTickVals) >= LBound(yTickVals) Then
            For Each val In yTickVals
                Set tick = doc.Shapes.AddLine(originX - 4, val, originX + 4, val)
                tick.Line.ForeColor.RGB = RGB(0, 0, 0): tick.Line.Weight = ${settings.tickWidth}
                shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name
            Next val
        End If
    End If

    ' --- 2.5 Auto Intercept Ticks ---
    If ${settings.showXIntercepts || settings.showYIntercepts ? 'True' : 'False'} Then
        debugStep = "Drawing Intercept Ticks"
        ${intercepts.map(i => {
           const eq = (i as any).curveIdx !== undefined ? settings.equations[(i as any).curveIdx] : undefined;
           const useComp = !!eq?.useCompression;
           const xsRaw = tx(i.x);
           const ysRaw = ty(i.y);
           const exRaw = xsRaw === null ? i.x : xsRaw;
           const eyRaw = ysRaw === null ? i.y : ysRaw;
           const ey = useComp ? compressYForCurve(eyRaw, exRaw) : eyRaw;
           const ex = useComp ? compressXForCurve(exRaw) : exRaw;
           const xExpr = `(${jsOriginX} + (${ex} * unitSize))`;
           const yExpr = `(${jsOriginY} - (${ey} * unitSize))`;
           if (i.y === 0) { // X-intercept: tick perpendicular to x-axis, centered on curve
             return `
        Set tick = doc.Shapes.AddLine(${xExpr}, ${yExpr} - 4, ${xExpr}, ${yExpr} + 4)
        tick.Line.ForeColor.RGB = 0: tick.Line.Weight = ${settings.tickWidth}
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name`;
           } else { // Y-intercept: tick perpendicular to y-axis, centered on curve
             return `
        Set tick = doc.Shapes.AddLine(${xExpr} - 4, ${yExpr}, ${xExpr} + 4, ${yExpr})
        tick.Line.ForeColor.RGB = 0: tick.Line.Weight = ${settings.tickWidth}
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name`;
           }
        }).join('')}
    End If

    ' --- 2.6 Manual Custom Labels ---
    ${settings.customLabels.map((l, lIdx) => {
      const sym = l.symbol || 'cross';
      return `
    ' Manual Label ${lIdx + 1}: ${l.text}
    Dim mx${lIdx}, my${lIdx}
    mx${lIdx} = ${vbaX(l.x)}
    my${lIdx} = ${vbaY(l.y)}
    ${sym === 'dot' ? `
    Set tick = doc.Shapes.AddShape(9, mx${lIdx} - 1.5, my${lIdx} - 1.5, 3, 3)
    tick.Fill.ForeColor.RGB = 0: tick.Line.Visible = 0
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name` : `
    Dim mla${lIdx}, mlb${lIdx}
    Set mla${lIdx} = doc.Shapes.AddLine(mx${lIdx} - 2, my${lIdx} - 2, mx${lIdx} + 2, my${lIdx} + 2)
    Set mlb${lIdx} = doc.Shapes.AddLine(mx${lIdx} - 2, my${lIdx} + 2, mx${lIdx} + 2, my${lIdx} - 2)
    mla${lIdx}.Line.ForeColor.RGB = 0: mlb${lIdx}.Line.ForeColor.RGB = 0
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = mla${lIdx}.Name
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = mlb${lIdx}.Name`}
    Set lbl = doc.Shapes.AddTextbox(1, mx${lIdx} + 4, my${lIdx} - 15, 50, 25)
    lbl.Fill.Visible = 0: lbl.Line.Visible = 0
    lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
    lbl.TextFrame.WordWrap = 0: lbl.TextFrame.TextRange.Text = "${l.text}"
    lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Size = ${settings.customLabelFontSize}
    lbl.TextFrame.TextRange.Font.Italic = True
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name`;
    }).join('')}

    ' --- 3. Draw Origin "O" ---
    If ${settings.showOrigin ? 'True' : 'False'} Then
        debugStep = "Drawing Origin Label"
        ' Positioned snugly at the bottom-left corner of the intersection
        Set lbl = doc.Shapes.AddTextbox(1, originX - 12, originY - 3, 30, 30)
        lbl.Fill.Visible = 0: lbl.Line.Visible = 0
        lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        lbl.TextFrame.WordWrap = 0
        lbl.TextFrame.TextRange.Text = "O"
        lbl.TextFrame.TextRange.Font.Name = "Times New Roman"
        lbl.TextFrame.TextRange.Font.Italic = True
        lbl.TextFrame.TextRange.Font.Size = ${settings.axisLabelFontSize}
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name
    End If

    ' --- 4. Draw Curves ---
    Dim fb As FreeformBuilder, curveComp As Shape
    ${allCurvePoints.map((curve, idx) => {
      const dashInt = VBA_CONSTANTS[curve.dashStyle] || 1;
      const curveCompress = settings.equations[idx]?.useCompression === true;
      return curve.segments.map((segment, sIdx) => {
        if (segment.length < 2) return '';
        const pointBuild = segment.map((p, pIdx) => {
          let xScaled = tx(p.x);
          let yScaled = ty(p.y);
          if (xScaled === null || yScaled === null) return ''; // skip invalid
          if (curveCompress) {
            const sxOrig = xScaled;
            yScaled = compressYForCurve(yScaled, sxOrig);
            xScaled = compressXForCurve(sxOrig);
          }
          const xPos = jsOriginX + (xScaled * settings.unitSize);
          const yPos = jsOriginY - (yScaled * settings.unitSize);
          if (pIdx === 0) return `Set fb = doc.Shapes.BuildFreeform(1, ${xPos}, ${yPos})`;
          return `fb.AddNodes 0, 1, ${xPos}, ${yPos}`;
        }).filter(Boolean).join('\n    ');
        
        return `
    ' --- Curve ${idx + 1} Segment ${sIdx + 1} ---
    debugStep = "Drawing Curve ${idx + 1} (Expression: ${settings.equations[idx].expression})"
    ${pointBuild}
    Set curveComp = fb.ConvertToShape
    curveComp.Line.ForeColor.RGB = ${hexToVbaRgb(curve.color)}
    curveComp.Line.Weight = ${curve.lineWidth}
    curveComp.Line.DashStyle = ${dashInt}
    shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = curveComp.Name`;
      }).join('\n');
    }).join('\n')}

    ' --- 5. Axis Labels ---
    debugStep = "Drawing Axis Extension Labels"
    
    If ${settings.showXLabel ? 'True' : 'False'} Then
        ' X Label: Restored to original position with additional -5 shift total
        Set lbl = doc.Shapes.AddTextbox(1, ${vbaX(nXMax)} - 7, ${vbaY(0)} - 6, 30, 30)
        lbl.Fill.Visible = 0: lbl.Line.Visible = 0
        lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        lbl.TextFrame.WordWrap = 0
        lbl.TextFrame.TextRange.Text = " " & "${settings.xAxisLabel}"
        lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Italic = True: lbl.TextFrame.TextRange.Font.Size = ${settings.axisLabelFontSize}
        lbl.TextFrame.TextRange.ParagraphFormat.Alignment = 1 ' Center
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name
    End If

    If ${settings.showYLabel ? 'True' : 'False'} Then
        ' Y Label: Restored to original position with additional -5 shift total
        Set lbl = doc.Shapes.AddTextbox(1, ${vbaX(0)} - 16, ${vbaY(nYMax)} - 9, 30, 30)
        lbl.Fill.Visible = 0: lbl.Line.Visible = 0
        lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        lbl.TextFrame.WordWrap = 0
        lbl.TextFrame.TextRange.Text = " " & "${settings.yAxisLabel}"
        lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Italic = True: lbl.TextFrame.TextRange.Font.Size = ${settings.axisLabelFontSize}
        lbl.TextFrame.TextRange.ParagraphFormat.Alignment = 0 ' Left alignment to prevent italic clipping
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name
    End If

    ' --- 6. Group Everything ---
    debugStep = "Final Grouping"
    ' Optional: Resume Next handles cases where grouping might fail due to page constraints
    On Error Resume Next
    If shpCount > 1 Then doc.Shapes.Range(shpArray).Group
    On Error GoTo 0

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred in Step: " & debugStep & vbCrLf & "Error Description: " & Err.Description, 16, "Graph Designer Error"
End Sub
    `.trim();
  }, [settings, nXMin, nXMax, nYMin, nYMax, allCurvePoints, xTicks, yTicks, intercepts, gridLines]);

  const copyToClipboard = () => {
    navigator.clipboard.writeText(generatedVBA);
    alert('VBA Code copied to clipboard!');
  };

  return (
    <div className="min-h-screen grid grid-cols-1 lg:grid-cols-[400px_1fr] h-screen overflow-hidden">
      {/* Settings Side Panel */}
      <aside className="bg-white border-r border-stone-200 overflow-y-auto p-6 space-y-8 flex flex-col justify-between shadow-xl z-20">
        <div className="space-y-6">
          <header className="flex items-center gap-3 border-b border-stone-100 pb-4">
            <div className="bg-stone-900 p-2 rounded-lg text-white">
              <FunctionSquare size={24} />
            </div>
            <div className="flex-1">
              <h1 className="text-xl font-bold tracking-tight text-stone-900 uppercase">MathGraph</h1>
              <p className="text-xs text-stone-500 font-mono italic">Word VBA Designer</p>
            </div>
            <div className="group relative">
              <Info size={16} className="text-stone-300 hover:text-stone-600 cursor-help" />
              <div className="absolute right-0 top-6 w-64 bg-stone-900 text-white text-[10px] p-3 rounded-xl opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-50 shadow-2xl space-y-2">
                <p><span className="text-amber-400 font-bold">Implicit Functions:</span> You can now type relations like <code className="bg-white/10 px-1 rounded">x^2 + y^2 = 9</code> or <code className="bg-white/10 px-1 rounded">y^2 = x</code>.</p>
                <p><span className="text-amber-400 font-bold">Logarithms:</span> Use <code className="bg-white/10 px-1 rounded">log(base, value)</code>. Example: <code className="bg-white/10 px-1 rounded">log(2, x)</code> for base-2.</p>
                <p><span className="text-amber-400 font-bold">Functions:</span> Use <code className="bg-white/10 px-1 rounded">y = sin(x)</code> for the fastest rendering.</p>
              </div>
            </div>
          </header>

          <section className="space-y-4">
            <div className="flex items-center justify-between border-b pb-1">
              <label className="text-xs font-bold uppercase tracking-widest text-stone-400 block">Equations</label>
              <button 
                onClick={() => setSettings({
                  ...settings,
                  equations: [...settings.equations, {
                    id: Date.now().toString(),
                    expression: '',
                    lineWidth: 0.75,
                    dashStyle: 'msoLineSolid',
                    color: '#000000',
                    useCompression: false,
                    intervals: [
                      {
                        id: 'int-' + Date.now(),
                        xMin: settings.xMin,
                        xMax: settings.xMax,
                        yMin: settings.yMin,
                        yMax: settings.yMax,
                        useCustomDomain: false,
                        useCustomRange: false
                      }
                    ]
                  }]
                })}
                className="text-stone-400 hover:text-stone-900 transition-colors"
              >
                <Plus size={16} />
              </button>
            </div>
            <div className="space-y-4 max-h-[400px] overflow-y-auto pr-1">
              {settings.equations.map((eq, idx) => (
                <div key={eq.id} className="p-3 bg-stone-50 rounded-xl space-y-3 relative group border border-stone-100">
                  <div className="flex items-center gap-2">
                    <div className="bg-stone-200 text-stone-600 text-[8px] font-bold px-1.5 py-0.5 rounded uppercase">Eq {idx + 1}</div>
                    <input 
                      value={eq.expression}
                      onChange={(e) => {
                        const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, expression: e.target.value } : l);
                        setSettings({ ...settings, equations: newEqs });
                      }}
                      className="flex-1 bg-white border border-stone-200 px-3 py-2 rounded-lg font-mono text-xs focus:ring-1 focus:ring-stone-900"
                    />
                    {settings.equations.length > 1 && (
                      <button 
                        onClick={() => setSettings({ ...settings, equations: settings.equations.filter(l => l.id !== eq.id) })}
                        className="text-stone-300 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={14} />
                      </button>
                    )}
                  </div>

                  {/* Multiple Intervals Support */}
                  <div className="space-y-2 border-t border-stone-100 pt-2">
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-[10px] text-stone-500 font-bold uppercase tracking-tighter">Domain/Range Intervals</span>
                      <button 
                        onClick={() => {
                          const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                            ...l, 
                            intervals: [...l.intervals, {
                              id: 'int-' + Date.now(),
                              xMin: settings.xMin,
                              xMax: settings.xMax,
                              yMin: settings.yMin,
                              yMax: settings.yMax,
                              useCustomDomain: false,
                              useCustomRange: false
                            }] 
                          } : l);
                          setSettings({ ...settings, equations: newEqs });
                        }}
                        className="text-stone-400 hover:text-stone-900"
                      >
                        <Plus size={12} />
                      </button>
                    </div>
                    
                    <div className="space-y-3">
                      {eq.intervals.map((interval, iIdx) => (
                        <div key={interval.id} className="bg-white/50 p-2 rounded-lg border border-stone-100 relative group/int">
                          {eq.intervals.length > 1 && (
                            <button 
                              onClick={() => {
                                const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, intervals: l.intervals.filter(i => i.id !== interval.id) } : l);
                                setSettings({ ...settings, equations: newEqs });
                              }}
                              className="absolute -right-1 -top-1 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover/int:opacity-100 transition-opacity"
                            >
                              <Trash2 size={8} />
                            </button>
                          )}
                          <div className="space-y-2">
                            <div className="space-y-1">
                              <div className="flex items-center justify-between">
                                <span className="text-[9px] text-stone-400 uppercase font-bold">X Range</span>
                                <button 
                                  onClick={() => {
                                    const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                      ...l, 
                                      intervals: l.intervals.map(i => i.id === interval.id ? { ...i, useCustomDomain: !i.useCustomDomain } : i)
                                    } : l);
                                    setSettings({ ...settings, equations: newEqs });
                                  }}
                                  className={cn("text-[8px] uppercase px-1 py-0.5 rounded", interval.useCustomDomain ? "bg-stone-900 text-white" : "bg-stone-200 text-stone-500")}
                                >
                                  {interval.useCustomDomain ? "Custom" : "Auto"}
                                </button>
                              </div>
                              {interval.useCustomDomain && (
                                <div className="grid grid-cols-2 gap-1 animate-in fade-in duration-200">
                                  <input 
                                    type="text" value={interval.xMin} 
                                    onChange={(e) => {
                                      const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                        ...l, 
                                        intervals: l.intervals.map(i => i.id === interval.id ? { ...i, xMin: e.target.value } : i)
                                      } : l);
                                      setSettings({ ...settings, equations: newEqs });
                                    }}
                                    className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                                    placeholder="min"
                                  />
                                  <input 
                                    type="text" value={interval.xMax} 
                                    onChange={(e) => {
                                      const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                        ...l, 
                                        intervals: l.intervals.map(i => i.id === interval.id ? { ...i, xMax: e.target.value } : i)
                                      } : l);
                                      setSettings({ ...settings, equations: newEqs });
                                    }}
                                    className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                                    placeholder="max"
                                  />
                                </div>
                              )}
                            </div>

                            <div className="space-y-1">
                              <div className="flex items-center justify-between">
                                <span className="text-[9px] text-stone-400 uppercase font-bold">Y Limit</span>
                                <button 
                                  onClick={() => {
                                    const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                      ...l, 
                                      intervals: l.intervals.map(i => i.id === interval.id ? { ...i, useCustomRange: !i.useCustomRange } : i)
                                    } : l);
                                    setSettings({ ...settings, equations: newEqs });
                                  }}
                                  className={cn("text-[8px] uppercase px-1 py-0.5 rounded", interval.useCustomRange ? "bg-stone-900 text-white" : "bg-stone-200 text-stone-500")}
                                >
                                  {interval.useCustomRange ? "Custom" : "Auto"}
                                </button>
                              </div>
                              {interval.useCustomRange && (
                                <div className="grid grid-cols-2 gap-1 animate-in fade-in duration-200">
                                  <input 
                                    type="text" value={interval.yMin} 
                                    onChange={(e) => {
                                      const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                        ...l, 
                                        intervals: l.intervals.map(i => i.id === interval.id ? { ...i, yMin: e.target.value } : i)
                                      } : l);
                                      setSettings({ ...settings, equations: newEqs });
                                    }}
                                    className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                                    placeholder="min"
                                  />
                                  <input 
                                    type="text" value={interval.yMax} 
                                    onChange={(e) => {
                                      const newEqs = settings.equations.map(l => l.id === eq.id ? { 
                                        ...l, 
                                        intervals: l.intervals.map(i => i.id === interval.id ? { ...i, yMax: e.target.value } : i)
                                      } : l);
                                      setSettings({ ...settings, equations: newEqs });
                                    }}
                                    className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                                    placeholder="max"
                                  />
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {settings.axisBreakEnabled && (
                    <div className="flex items-center justify-between border-t border-stone-100 pt-2">
                      <span className="text-[10px] text-stone-500 font-bold uppercase tracking-tighter">Use compression bands</span>
                      <button
                        onClick={() => {
                          const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, useCompression: !l.useCompression } : l);
                          setSettings({ ...settings, equations: newEqs });
                        }}
                        className={cn("text-[9px] uppercase px-2 py-0.5 rounded font-bold tracking-wide", eq.useCompression ? "bg-stone-900 text-white" : "bg-stone-200 text-stone-500")}
                      >
                        {eq.useCompression ? "On" : "Off"}
                      </button>
                    </div>
                  )}

                  <div className="grid grid-cols-3 gap-2 border-t border-stone-100 pt-2">
                    <div className="space-y-1">
                      <span className="text-[10px] text-stone-400 uppercase">Width</span>
                      <input
                        type="number" step="0.25" min="0.25"
                        value={eq.lineWidth}
                        onChange={(e) => {
                          const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, lineWidth: Number(e.target.value) } : l);
                          setSettings({ ...settings, equations: newEqs });
                        }}
                        className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                      />
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] text-stone-400 uppercase">Style</span>
                      <select 
                        value={eq.dashStyle}
                        onChange={(e) => {
                          const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, dashStyle: e.target.value } : l);
                          setSettings({ ...settings, equations: newEqs });
                        }}
                        className="w-full bg-white border border-stone-200 p-1 rounded text-[10px]"
                      >
                        {DASH_STYLES.map(s => <option key={s.value} value={s.value}>{s.label}</option>)}
                      </select>
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] text-stone-400 uppercase">Color</span>
                      <div className="flex gap-1 items-center h-[26px]">
                        <input 
                          type="color"
                          value={eq.color || '#000000'}
                          onChange={(e) => {
                            const newEqs = settings.equations.map(l => l.id === eq.id ? { ...l, color: e.target.value } : l);
                            setSettings({ ...settings, equations: newEqs });
                          }}
                          className="w-6 h-6 p-0 border-0 bg-transparent cursor-pointer block"
                        />
                        <span className="text-[8px] font-mono text-stone-400 uppercase whitespace-nowrap">{(eq.color || '#000000')}</span>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </section>

          <section className="grid grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="text-[10px] font-bold uppercase tracking-widest text-stone-400">X Range</label>
              <div className="flex items-center gap-2">
                <input 
                  type="number"
                  value={settings.xMin}
                  onChange={(e) => setSettings({ ...settings, xMin: e.target.value })}
                  className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-xs"
                />
                <span className="text-stone-300">to</span>
                <input 
                  type="number"
                  value={settings.xMax}
                  onChange={(e) => setSettings({ ...settings, xMax: e.target.value })}
                  className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-xs"
                />
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold uppercase tracking-widest text-stone-400">Y Range</label>
              <div className="flex items-center gap-2">
                <input 
                  type="number"
                  value={settings.yMin}
                  onChange={(e) => setSettings({ ...settings, yMin: e.target.value })}
                  className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-xs"
                />
                <span className="text-stone-300">to</span>
                <input 
                  type="number"
                  value={settings.yMax}
                  onChange={(e) => setSettings({ ...settings, yMax: e.target.value })}
                  className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-xs"
                />
              </div>
            </div>
          </section>

          <section className="space-y-3">
            <label className="text-[10px] font-bold uppercase tracking-widest text-stone-400 block border-b pb-1">Visual Settings</label>
            <div className="space-y-4 pt-1">
              <div className="flex items-center justify-between">
                <span className="text-sm text-stone-600">Unit Size (px)</span>
                <input 
                  type="range" min="10" max="100" 
                  value={settings.unitSize}
                  onChange={(e) => setSettings({ ...settings, unitSize: Number(e.target.value) })}
                  className="w-32 accent-stone-900"
                />
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm text-stone-600">Tick Width</span>
                <input 
                  type="number" step="0.05" min="0.1"
                  value={settings.tickWidth}
                  onChange={(e) => setSettings({ ...settings, tickWidth: Number(e.target.value) })}
                  className="w-16 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                />
              </div>
              <div className="flex items-center justify-between">
                <div className="flex flex-col">
                  <span className="text-sm text-stone-600">Axis Label Size</span>
                  <span className="text-[10px] text-stone-400 -mt-1 italic">x, y, O</span>
                </div>
                <input 
                  type="number" step="1" min="6" max="48"
                  value={settings.axisLabelFontSize}
                  onChange={(e) => setSettings({ ...settings, axisLabelFontSize: Number(e.target.value) })}
                  className="w-16 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                />
              </div>
              <div className="flex items-center justify-between">
                <div className="flex flex-col">
                  <span className="text-sm text-stone-600">Custom Label Size</span>
                  <span className="text-[10px] text-stone-400 -mt-1 italic">point labels</span>
                </div>
                <input 
                  type="number" step="1" min="6" max="48"
                  value={settings.customLabelFontSize}
                  onChange={(e) => setSettings({ ...settings, customLabelFontSize: Number(e.target.value) })}
                  className="w-16 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                />
              </div>

              {/* Axis Scaling */}
              <div className="space-y-2 pt-2 border-t">
                <span className="text-[10px] text-stone-400 uppercase font-bold block">Axis Scaling</span>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs text-stone-600 w-12">X Axis</span>
                  <select 
                    value={settings.xScaleMode}
                    onChange={(e) => setSettings({ ...settings, xScaleMode: e.target.value })}
                    className="flex-1 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                  >
                    {SCALE_OPTIONS.map(o => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ))}
                  </select>
                </div>
                {settings.xScaleMode === 'linear-custom' && (
                  <div className="flex items-center justify-between gap-2 pl-14">
                    <span className="text-[10px] text-stone-400 italic">multiplier</span>
                    <input
                      type="number" step="0.1"
                      value={settings.xScaleMultiplier}
                      onChange={(e) => setSettings({ ...settings, xScaleMultiplier: Number(e.target.value) })}
                      className="w-20 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                    />
                  </div>
                )}
                {settings.xScaleMode === 'custom' && (
                  <div className="flex items-center justify-between gap-2 pl-14">
                    <span className="text-[10px] text-stone-400 italic">f(x) =</span>
                    <input
                      type="text"
                      value={settings.xScaleCustom}
                      placeholder="e.g. sqrt(x), 1/x"
                      onChange={(e) => setSettings({ ...settings, xScaleCustom: e.target.value })}
                      className="flex-1 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                    />
                  </div>
                )}

                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs text-stone-600 w-12">Y Axis</span>
                  <select 
                    value={settings.yScaleMode}
                    onChange={(e) => setSettings({ ...settings, yScaleMode: e.target.value })}
                    className="flex-1 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                  >
                    {SCALE_OPTIONS.map(o => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ))}
                  </select>
                </div>
                {settings.yScaleMode === 'linear-custom' && (
                  <div className="flex items-center justify-between gap-2 pl-14">
                    <span className="text-[10px] text-stone-400 italic">multiplier</span>
                    <input
                      type="number" step="0.1"
                      value={settings.yScaleMultiplier}
                      onChange={(e) => setSettings({ ...settings, yScaleMultiplier: Number(e.target.value) })}
                      className="w-20 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                    />
                  </div>
                )}
                {settings.yScaleMode === 'custom' && (
                  <div className="flex items-center justify-between gap-2 pl-14">
                    <span className="text-[10px] text-stone-400 italic">f(y) =</span>
                    <input
                      type="text"
                      value={settings.yScaleCustom}
                      placeholder="e.g. sqrt(y), 1/y"
                      onChange={(e) => setSettings({ ...settings, yScaleCustom: e.target.value })}
                      className="flex-1 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                    />
                  </div>
                )}
                {(xIsLog || yIsLog) && (
                  <p className="text-[9px] text-amber-700 bg-amber-50 border border-amber-200 p-1.5 rounded italic leading-snug">
                    ⚠ On log scales, only positive values are plottable. The range must be &gt; 0.
                  </p>
                )}
              </div>

              {/* Axis Compression Bands */}
              <div className="space-y-2 pt-2 border-t">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] text-stone-400 uppercase font-bold block">Axis Compression Bands</span>
                  <button
                    onClick={() => {
                      const enabling = !settings.axisBreakEnabled;
                      const seedBand: AxisBreakBand = {
                        id: 'band-' + Date.now(),
                        axis: 'y',
                        min: '-2',
                        max: '2',
                        compression: 0.2,
                      };
                      setSettings({
                        ...settings,
                        axisBreakEnabled: enabling,
                        axisBreakBands: enabling && settings.axisBreakBands.length === 0 ? [seedBand] : settings.axisBreakBands,
                      });
                    }}
                    className={cn("w-10 h-5 rounded-full transition-colors relative", settings.axisBreakEnabled ? "bg-stone-900" : "bg-stone-200")}
                  >
                    <div className={cn("absolute top-1 w-3 h-3 bg-white rounded-full transition-all", settings.axisBreakEnabled ? "left-6" : "left-1")} />
                  </button>
                </div>
                {settings.axisBreakEnabled && (
                  <div className="space-y-2 animate-in fade-in slide-in-from-top-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] text-stone-500">{settings.axisBreakBands.length} band{settings.axisBreakBands.length === 1 ? '' : 's'}</span>
                      <button
                        onClick={() => setSettings({
                          ...settings,
                          axisBreakBands: [
                            ...settings.axisBreakBands,
                            { id: 'band-' + Date.now(), axis: 'x', min: '-1', max: '1', compression: 0.2 },
                          ],
                        })}
                        className="text-[10px] text-stone-500 hover:text-stone-900 font-bold flex items-center gap-0.5"
                      >
                        <Plus size={12} /> Add band
                      </button>
                    </div>
                    <div className="space-y-2">
                      {settings.axisBreakBands.map((band, bIdx) => (
                        <div key={band.id} className="bg-white border border-stone-200 rounded p-2 space-y-1.5 relative group">
                          <div className="flex items-center justify-between">
                            <span className="text-[9px] text-stone-400 uppercase font-bold tracking-wider">Band {bIdx + 1}</span>
                            <button
                              onClick={() => setSettings({
                                ...settings,
                                axisBreakBands: settings.axisBreakBands.filter(b => b.id !== band.id),
                              })}
                              className="text-stone-300 hover:text-red-500 transition-colors"
                              title="Delete band"
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[10px] text-stone-500 w-10">Axis</span>
                            <div className="flex bg-stone-50 border border-stone-200 rounded p-0.5 flex-1">
                              <button
                                onClick={() => setSettings({
                                  ...settings,
                                  axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, axis: 'x' } : b),
                                })}
                                className={cn("flex-1 text-[10px] py-0.5 rounded transition-all", band.axis === 'x' ? "bg-stone-900 text-white" : "text-stone-500")}
                              >X</button>
                              <button
                                onClick={() => setSettings({
                                  ...settings,
                                  axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, axis: 'y' } : b),
                                })}
                                className={cn("flex-1 text-[10px] py-0.5 rounded transition-all", band.axis === 'y' ? "bg-stone-900 text-white" : "text-stone-500")}
                              >Y</button>
                            </div>
                          </div>
                          <div className="flex items-center justify-between gap-1">
                            <span className="text-[10px] text-stone-500 w-10">Range</span>
                            <input
                              type="number" step="0.1" value={band.min}
                              onChange={(e) => setSettings({
                                ...settings,
                                axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, min: e.target.value } : b),
                              })}
                              className="w-full bg-stone-50 border border-stone-200 p-1 rounded text-[10px] font-mono"
                              placeholder="min"
                            />
                            <span className="text-stone-300 text-[9px]">to</span>
                            <input
                              type="number" step="0.1" value={band.max}
                              onChange={(e) => setSettings({
                                ...settings,
                                axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, max: e.target.value } : b),
                              })}
                              className="w-full bg-stone-50 border border-stone-200 p-1 rounded text-[10px] font-mono"
                              placeholder="max"
                            />
                          </div>
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[10px] text-stone-500 w-10">Scale</span>
                            <input
                              type="range" min="0" max="3" step="0.01"
                              value={band.compression}
                              onChange={(e) => setSettings({
                                ...settings,
                                axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, compression: Number(e.target.value) } : b),
                              })}
                              className="flex-1 accent-stone-900"
                            />
                            <input
                              type="number" min="0" step="0.05"
                              value={band.compression}
                              onChange={(e) => setSettings({
                                ...settings,
                                axisBreakBands: settings.axisBreakBands.map(b => b.id === band.id ? { ...b, compression: Number(e.target.value) } : b),
                              })}
                              className="w-14 bg-stone-50 border border-stone-200 p-1 rounded text-[10px] font-mono text-right"
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                    <p className="text-[9px] text-stone-400 italic leading-snug">
                      Range is always an x-interval. Axis X rescales the x-coordinate inside that interval (0 = collapsed to a vertical line, 1 = unchanged, &gt;1 = horizontal stretch). Axis Y multiplies the curve's y values for x in that interval (0 = flattens to x-axis, 1 = unchanged, &gt;1 = vertical stretch). Bands on the same axis should be disjoint, and curves only show the effect when "Use compression bands" is enabled on the equation.
                    </p>
                  </div>
                )}
              </div>

              {/* Sampling Step */}
              <div className="flex items-center justify-between pt-2 border-t">
                <div className="flex flex-col">
                  <span className="text-sm text-stone-600">Sampling Step</span>
                  <span className="text-[10px] text-stone-400 -mt-1 italic">smaller = smoother &amp; slower</span>
                </div>
                <input
                  type="number" step="0.001" min="0.0001"
                  value={settings.samplingStep}
                  onChange={(e) => setSettings({ ...settings, samplingStep: Number(e.target.value) })}
                  className="w-20 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                />
              </div>

              <div className="flex items-center justify-between">
                <span className="text-sm text-stone-600">Show Ticks</span>
                <button 
                  onClick={() => setSettings({ ...settings, showTicks: !settings.showTicks })}
                  className={cn(
                    "w-10 h-5 rounded-full transition-colors relative",
                    settings.showTicks ? "bg-stone-900" : "bg-stone-200"
                  )}
                >
                  <div className={cn(
                    "absolute top-1 w-3 h-3 bg-white rounded-full transition-all",
                    settings.showTicks ? "left-6" : "left-1"
                  )} />
                </button>
              </div>

              {settings.showTicks && (
                <div className="space-y-3 animate-in fade-in slide-in-from-top-2">
                  <div className="space-y-1">
                    <span className="text-[10px] text-stone-400 uppercase font-bold">Custom X Ticks (e.g. -2, 0, 2)</span>
                    <input 
                      type="text" 
                      placeholder="Leave empty for auto"
                      value={settings.customXTicks}
                      onChange={(e) => setSettings({ ...settings, customXTicks: e.target.value })}
                      className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-[10px] font-mono"
                    />
                  </div>
                  <div className="space-y-1">
                    <span className="text-[10px] text-stone-400 uppercase font-bold">Custom Y Ticks</span>
                    <input 
                      type="text" 
                      placeholder="Leave empty for auto"
                      value={settings.customYTicks}
                      onChange={(e) => setSettings({ ...settings, customYTicks: e.target.value })}
                      className="w-full bg-stone-50 border border-stone-200 p-2 rounded text-[10px] font-mono"
                    />
                  </div>
                </div>
              )}
              <div className="flex items-center justify-between">
                <span className="text-sm text-stone-600">Arrow Style</span>
                <div className="flex items-center gap-1 bg-stone-50 border border-stone-200 p-1 rounded-lg">
                  {ARROW_STYLES.map(s => (
                    <button 
                      key={s.value}
                      onClick={() => setSettings({ ...settings, arrowStyle: s.value })}
                      title={s.label}
                      className={cn(
                        "p-1 rounded transition-all",
                        settings.arrowStyle === s.value ? "bg-stone-900 text-white shadow-inner" : "text-stone-400 hover:bg-stone-200"
                      )}
                    >
                      <div className="w-4 h-4 flex items-center justify-center">
                        <svg viewBox="0 0 10 10" className="w-full h-full fill-current">
                          {s.value === 'msoArrowheadNone' && <rect x="0" y="4.5" width="10" height="1" />}
                          {s.value === 'msoArrowheadTriangle' && <polygon points="0 2, 8 5, 0 8" />}
                          {s.value === 'msoArrowheadOpen' && <polyline points="0 2, 8 5, 0 8" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />}
                          {s.value === 'msoArrowheadStealth' && <polygon points="0 2, 8 5, 0 8, 3 5" />}
                          {s.value === 'msoArrowheadDiamond' && <polygon points="3 2, 8 5, 3 8, 0 5" />}
                          {s.value === 'msoArrowheadOval' && <circle cx="5" cy="5" r="3.5" />}
                        </svg>
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-3 pt-2">
                <span className="text-[10px] text-stone-400 uppercase font-bold block border-b pb-1">Visibility Controls</span>
                <div className="grid grid-cols-2 gap-x-4 gap-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] text-stone-600">X-Axis</span>
                    <button onClick={() => setSettings({ ...settings, showXAxis: !settings.showXAxis })} className={cn("w-8 h-4 rounded-full transition-colors relative", settings.showXAxis ? "bg-stone-900" : "bg-stone-200")}>
                      <div className={cn("absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all", settings.showXAxis ? "left-4.5" : "left-0.5")} />
                    </button>
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] text-stone-600 italic">X Label</span>
                    <button onClick={() => setSettings({ ...settings, showXLabel: !settings.showXLabel })} className={cn("w-8 h-4 rounded-full transition-colors relative", settings.showXLabel ? "bg-stone-900" : "bg-stone-200")}>
                      <div className={cn("absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all", settings.showXLabel ? "left-4.5" : "left-0.5")} />
                    </button>
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] text-stone-600">Y-Axis</span>
                    <button onClick={() => setSettings({ ...settings, showYAxis: !settings.showYAxis })} className={cn("w-8 h-4 rounded-full transition-colors relative", settings.showYAxis ? "bg-stone-900" : "bg-stone-200")}>
                      <div className={cn("absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all", settings.showYAxis ? "left-4.5" : "left-0.5")} />
                    </button>
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] text-stone-600 italic">Y Label</span>
                    <button onClick={() => setSettings({ ...settings, showYLabel: !settings.showYLabel })} className={cn("w-8 h-4 rounded-full transition-colors relative", settings.showYLabel ? "bg-stone-900" : "bg-stone-200")}>
                      <div className={cn("absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all", settings.showYLabel ? "left-4.5" : "left-0.5")} />
                    </button>
                  </div>
                  <div className="flex items-center justify-between col-span-2 border-t pt-1 mt-1">
                    <span className="text-[10px] text-stone-600 italic">Origin "O" Label</span>
                    <button onClick={() => setSettings({ ...settings, showOrigin: !settings.showOrigin })} className={cn("w-8 h-4 rounded-full transition-colors relative", settings.showOrigin ? "bg-stone-900" : "bg-stone-200")}>
                      <div className={cn("absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all", settings.showOrigin ? "left-4.5" : "left-0.5")} />
                    </button>
                  </div>
                </div>
              </div>

              <div className="flex items-center justify-between">
                <div className="flex flex-col">
                  <span className="text-sm text-stone-600">Auto Intercepts</span>
                  <span className="text-[10px] text-stone-400 -mt-1 italic">Adding tick to intercept</span>
                </div>
                <div className="flex items-center gap-4">
                  <div className="flex flex-col items-center gap-1">
                    <span className="text-[8px] text-stone-400 font-bold uppercase tracking-tight">X-Axis</span>
                    <button 
                      onClick={() => setSettings({ ...settings, showXIntercepts: !settings.showXIntercepts })}
                      className={cn(
                        "w-10 h-5 rounded-full transition-colors relative",
                        settings.showXIntercepts ? "bg-stone-900" : "bg-stone-200"
                      )}
                    >
                      <div className={cn(
                        "absolute top-1 w-3 h-3 bg-white rounded-full transition-all",
                        settings.showXIntercepts ? "left-6" : "left-1"
                      )} />
                    </button>
                  </div>
                  <div className="flex flex-col items-center gap-1">
                    <span className="text-[8px] text-stone-400 font-bold uppercase tracking-tight">Y-Axis</span>
                    <button 
                      onClick={() => setSettings({ ...settings, showYIntercepts: !settings.showYIntercepts })}
                      className={cn(
                        "w-10 h-5 rounded-full transition-colors relative",
                        settings.showYIntercepts ? "bg-stone-900" : "bg-stone-200"
                      )}
                    >
                      <div className={cn(
                        "absolute top-1 w-3 h-3 bg-white rounded-full transition-all",
                        settings.showYIntercepts ? "left-6" : "left-1"
                      )} />
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </section>

          {/* Grid Settings */}
          <section className="space-y-3">
            <div className="flex items-center justify-between border-b pb-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-stone-400 block">Grid</label>
              <button 
                onClick={() => setSettings({ ...settings, showGrid: !settings.showGrid })}
                className={cn(
                  "w-10 h-5 rounded-full transition-colors relative",
                  settings.showGrid ? "bg-stone-900" : "bg-stone-200"
                )}
              >
                <div className={cn(
                  "absolute top-1 w-3 h-3 bg-white rounded-full transition-all",
                  settings.showGrid ? "left-6" : "left-1"
                )} />
              </button>
            </div>
            {settings.showGrid && (
              <div className="space-y-2 animate-in fade-in slide-in-from-top-2">
                <div className="grid grid-cols-2 gap-2">
                  <div className="flex flex-col">
                    <span className="text-[10px] text-stone-400 uppercase font-bold">X Spacing</span>
                    <input 
                      type="number" step="0.1" min="0.01"
                      value={settings.gridSpacingX}
                      onChange={(e) => setSettings({ ...settings, gridSpacingX: Number(e.target.value) })}
                      className="bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                    />
                  </div>
                  <div className="flex flex-col">
                    <span className="text-[10px] text-stone-400 uppercase font-bold">Y Spacing</span>
                    <input 
                      type="number" step="0.1" min="0.01"
                      value={settings.gridSpacingY}
                      onChange={(e) => setSettings({ ...settings, gridSpacingY: Number(e.target.value) })}
                      className="bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                    />
                  </div>
                </div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[10px] text-stone-400 uppercase font-bold">Color</span>
                  <div className="flex items-center gap-2 flex-1 justify-end">
                    <input 
                      type="color"
                      value={settings.gridColor}
                      onChange={(e) => setSettings({ ...settings, gridColor: e.target.value })}
                      className="w-8 h-6 border border-stone-200 rounded cursor-pointer"
                    />
                    <input 
                      type="text"
                      value={settings.gridColor}
                      onChange={(e) => setSettings({ ...settings, gridColor: e.target.value })}
                      className="w-20 bg-stone-50 border border-stone-200 p-1 rounded text-xs font-mono"
                    />
                  </div>
                </div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[10px] text-stone-400 uppercase font-bold">Opacity</span>
                  <div className="flex items-center gap-2 flex-1 justify-end">
                    <input 
                      type="range" min="0.05" max="1" step="0.05"
                      value={settings.gridOpacity}
                      onChange={(e) => setSettings({ ...settings, gridOpacity: Number(e.target.value) })}
                      className="w-24 accent-stone-900"
                    />
                    <span className="text-xs text-stone-500 font-mono w-8 text-right">{settings.gridOpacity.toFixed(2)}</span>
                  </div>
                </div>
                <div className="flex items-center justify-between">
                  <span className="text-[10px] text-stone-400 uppercase font-bold">Line Width</span>
                  <input 
                    type="number" step="0.05" min="0.1"
                    value={settings.gridLineWidth}
                    onChange={(e) => setSettings({ ...settings, gridLineWidth: Number(e.target.value) })}
                    className="w-16 bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                  />
                </div>
              </div>
            )}
          </section>

          <section className="space-y-3">
            <div className="flex items-center justify-between border-b pb-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-stone-400 block">Custom Labels</label>
              <button 
                onClick={() => setSettings({
                  ...settings,
                  customLabels: [...settings.customLabels, { id: Date.now().toString(), x: 1, y: 1, text: 'P', symbol: 'cross' }]
                })}
                className="text-stone-400 hover:text-stone-900 transition-colors"
                title="Add manual label"
              >
                <Plus size={16} />
              </button>
            </div>
            <div className="space-y-2 max-h-[150px] overflow-y-auto">
              {settings.customLabels.map(label => (
                <div key={label.id} className="flex items-center gap-1 bg-stone-50 p-2 rounded-lg group">
                  <input 
                    type="text" value={label.text} 
                    onChange={(e) => {
                      const newLabels = settings.customLabels.map(l => l.id === label.id ? { ...l, text: e.target.value } : l);
                      setSettings({ ...settings, customLabels: newLabels });
                    }}
                    className="bg-transparent border-none text-[10px] font-bold w-10 focus:ring-0"
                  />
                  <div className="flex bg-stone-200 rounded p-0.5">
                    <button 
                      onClick={() => {
                        const newLabels = settings.customLabels.map(l => l.id === label.id ? { ...l, symbol: 'dot' } : l);
                        setSettings({ ...settings, customLabels: newLabels });
                      }}
                      className={cn("w-4 h-4 rounded text-[8px] flex items-center justify-center transition-all", label.symbol === 'dot' ? "bg-white text-stone-900 shadow-sm" : "text-stone-500")}
                    >●</button>
                    <button 
                      onClick={() => {
                        const newLabels = settings.customLabels.map(l => l.id === label.id ? { ...l, symbol: 'cross' } : l);
                        setSettings({ ...settings, customLabels: newLabels });
                      }}
                      className={cn("w-4 h-4 rounded text-[8px] flex items-center justify-center transition-all", (label.symbol === 'cross' || !label.symbol) ? "bg-white text-stone-900 shadow-sm" : "text-stone-500")}
                    >×</button>
                  </div>
                  <input 
                    type="number" value={label.x} 
                    onChange={(e) => {
                      const newLabels = settings.customLabels.map(l => l.id === label.id ? { ...l, x: Number(e.target.value) } : l);
                      setSettings({ ...settings, customLabels: newLabels });
                    }}
                    className="bg-transparent border-none text-[10px] w-8 focus:ring-0"
                  />
                  <input 
                    type="number" value={label.y} 
                    onChange={(e) => {
                      const newLabels = settings.customLabels.map(l => l.id === label.id ? { ...l, y: Number(e.target.value) } : l);
                      setSettings({ ...settings, customLabels: newLabels });
                    }}
                    className="bg-transparent border-none text-[10px] w-8 focus:ring-0"
                  />
                  <button 
                    onClick={() => setSettings({ ...settings, customLabels: settings.customLabels.filter(l => l.id !== label.id) })}
                    className="opacity-0 group-hover:opacity-100 text-stone-300 hover:text-red-500 transition-all ml-auto"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              ))}
            </div>
          </section>

          <section className="bg-stone-50 border border-stone-100 rounded-xl p-4 space-y-2">
            <div className="flex items-center gap-2 text-stone-900">
              <Info size={14} />
              <h3 className="text-[10px] font-bold uppercase tracking-widest">How to Use</h3>
            </div>
            <ol className="text-[10px] text-stone-500 space-y-2 list-decimal list-inside leading-relaxed">
              <li>Configure your graph settings here</li>
              <li>Click <span className="text-stone-900 font-bold">Generate VBA Script</span></li>
              <li>In MS Word: <span className="italic">Turn on Developer Mode &gt; Visual Basic</span></li>
              <li><span className="italic">Insert &gt; Module</span> and paste code</li>
              <li>Press <span className="text-stone-900 font-bold">F5</span> or run <span className="font-mono">DrawGraph</span></li>
            </ol>
          </section>
        </div>

        <button 
          onClick={() => setVbaVisible(true)}
          className="w-full bg-stone-900 text-white py-4 rounded-xl flex items-center justify-center gap-2 hover:bg-stone-800 transition-all group mt-6"
        >
          <Code2 size={18} className="group-hover:scale-110 transition-transform" />
          <span className="font-bold uppercase tracking-wider text-xs">Generate VBA Script</span>
        </button>
      </aside>

      {/* Main View Area */}
      <main className="relative bg-stone-100 p-8 flex items-center justify-center overflow-auto">
        {/* Background Paper Feel */}
        <div 
          className="bg-white shadow-2xl relative overflow-hidden flex items-center justify-center"
          style={{ width: canvasWidth + 100, height: canvasHeight + 100 }}
        >
          {/* Subtle Grid Substructure */}
          <div className="absolute inset-0 opacity-[0.03] pointer-events-none" style={{ backgroundSize: `${settings.unitSize}px ${settings.unitSize}px`, backgroundImage: 'radial-gradient(circle, #000 1px, transparent 1px)' }} />

          <svg 
            width={canvasWidth + 80} 
            height={canvasHeight + 80} 
            viewBox={`-40 -40 ${canvasWidth + 80} ${canvasHeight + 80}`}
            className="overflow-visible"
          >
            {/* Definitions (Arrowheads) */}
            <defs>
              <marker id="arrowhead" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
                {settings.arrowStyle !== 'msoArrowheadNone' && (
                  <path 
                    d={
                      settings.arrowStyle === 'msoArrowheadTriangle' ? 'M 0 0 L 10 3.5 L 0 7 Z' : 
                      settings.arrowStyle === 'msoArrowheadStealth' ? 'M 0 0 L 10 3.5 L 0 7 L 3 3.5 Z' :
                      settings.arrowStyle === 'msoArrowheadDiamond' ? 'M 3 0 L 10 3.5 L 3 7 L 0 3.5 Z' :
                      settings.arrowStyle === 'msoArrowheadOval' ? 'M 1 3.5 A 4 4 0 1 0 9 3.5 A 4 4 0 1 0 1 3.5' :
                      settings.arrowStyle === 'msoArrowheadOpen' ? 'M 0 0 L 10 3.5 L 0 7 L 0 6 L 8 3.5 L 0 1 Z' : 'M 0 0 L 10 3.5 L 0 7 Z'
                    } 
                    fill="#000" 
                  />
                )}
              </marker>
            </defs>

            {/* Grid */}
            {settings.showGrid && (
              <g opacity={settings.gridOpacity} className="animate-in fade-in duration-300">
                {gridLines.xs.map((gx) => {
                  const p1 = toPoints(gx, Number(settings.yMin));
                  const p2 = toPoints(gx, Number(settings.yMax));
                  return (
                    <line key={`gx-${gx}`} x1={p1.x} y1={p1.y} x2={p2.x} y2={p2.y}
                          stroke={settings.gridColor} strokeWidth={settings.gridLineWidth} />
                  );
                })}
                {gridLines.ys.map((gy) => {
                  const p1 = toPoints(Number(settings.xMin), gy);
                  const p2 = toPoints(Number(settings.xMax), gy);
                  return (
                    <line key={`gy-${gy}`} x1={p1.x} y1={p1.y} x2={p2.x} y2={p2.y}
                          stroke={settings.gridColor} strokeWidth={settings.gridLineWidth} />
                  );
                })}
              </g>
            )}

            {/* Axes */}
            {settings.showXAxis && (
              <line 
                x1={toPoints(Number(settings.xMin), 0).x} y1={toPoints(Number(settings.xMin), 0).y}
                x2={toPoints(Number(settings.xMax), 0).x} y2={toPoints(Number(settings.xMax), 0).y}
                stroke="#000" strokeWidth="0.75" markerEnd="url(#arrowhead)"
                className="transition-all duration-300"
              />
            )}
            {settings.showYAxis && (
              <line 
                x1={toPoints(0, Number(settings.yMin)).x} y1={toPoints(0, Number(settings.yMin)).y}
                x2={toPoints(0, Number(settings.yMax)).x} y2={toPoints(0, Number(settings.yMax)).y}
                stroke="#000" strokeWidth="0.75" markerEnd="url(#arrowhead)"
                className="transition-all duration-300"
              />
            )}

            {/* Ticks & Number Labels */}
            {settings.showTicks && (
              <g className="animate-in fade-in duration-300">
                {/* X-Axis Ticks */}
                {xTicks.map((x) => {
                  if (x === 0) return null;
                  const { x: px, y: py } = toPoints(x, 0);
                  return (
                    <g key={`x-${x}`}>
                      <line x1={px} y1={py - 4} x2={px} y2={py + 4} stroke="#000" strokeWidth={settings.tickWidth} />
                    </g>
                  );
                })}
                
                {/* Y-Axis Ticks */}
                {yTicks.map((y) => {
                  if (y === 0) return null;
                  const { x: px, y: py } = toPoints(0, y);
                  return (
                    <g key={`y-${y}`}>
                      <line x1={px - 4} y1={py} x2={px + 4} y2={py} stroke="#000" strokeWidth={settings.tickWidth} />
                    </g>
                  );
                })}
              </g>
            )}

            {/* Curves */}
            {allCurvePoints.map((curve, idx) => (
              <g key={curve.id}>
                {curve.segments.map((segment, sIdx) => {
                  const curveCompress = settings.equations[idx]?.useCompression === true;
                  const pts = segment.map(p => {
                    let xs = tx(p.x);
                    let ys = ty(p.y);
                    const ex0 = xs === null ? p.x : xs;
                    const ey0 = ys === null ? p.y : ys;
                    const ey = curveCompress ? compressYForCurve(ey0, ex0) : ey0;
                    const ex = curveCompress ? compressXForCurve(ex0) : ex0;
                    return {
                      x: scaledOriginX + ex * settings.unitSize,
                      y: scaledOriginY - ey * settings.unitSize,
                    };
                  });
                  let d = '';
                  if (pts.length === 0) {
                    d = '';
                  } else if (pts.length <= 2) {
                    d = pts.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
                  } else {
                    // Catmull-Rom → cubic bezier for smooth curves
                    d = `M ${pts[0].x} ${pts[0].y}`;
                    for (let i = 0; i < pts.length - 1; i++) {
                      const p0 = pts[Math.max(i - 1, 0)];
                      const p1 = pts[i];
                      const p2 = pts[i + 1];
                      const p3 = pts[Math.min(i + 2, pts.length - 1)];
                      const cp1x = p1.x + (p2.x - p0.x) / 6;
                      const cp1y = p1.y + (p2.y - p0.y) / 6;
                      const cp2x = p2.x - (p3.x - p1.x) / 6;
                      const cp2y = p2.y - (p3.y - p1.y) / 6;
                      d += ` C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${p2.x} ${p2.y}`;
                    }
                  }
                  return (
                    <path
                      key={`${curve.id}-${sIdx}`}
                      d={d}
                      stroke={curve.color}
                      strokeWidth={curve.lineWidth * 1.5}
                      strokeLinejoin="round"
                      strokeLinecap="round"
                      strokeDasharray={DASH_STYLES.find(s => s.value === curve.dashStyle)?.dash}
                      fill="none"
                      className="animate-in fade-in duration-500"
                    />
                  );
                })}
              </g>
            ))}

            {/* Axis Labels */}
            {settings.showXLabel && (
              <text x={toPoints(settings.xMax, 0).x + 10} y={toPoints(settings.xMax, 0).y + 15} fontSize={settings.axisLabelFontSize} className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>{settings.xAxisLabel}</text>
            )}
            {settings.showYLabel && (
              <text x={toPoints(0, settings.yMax).x - 15} y={toPoints(0, settings.yMax).y - 12} fontSize={settings.axisLabelFontSize} className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>{settings.yAxisLabel}</text>
            )}

            {/* Origin */}
            {settings.showOrigin && (
              <text x={originX - 12} y={originY + 12} fontSize={settings.axisLabelFontSize} className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>O</text>
            )}

            {/* Custom Labels & Detected Intercepts */}
            {[...intercepts, ...settings.customLabels].map((label, idx) => {
              const curveIdx = (label as any).curveIdx as number | undefined;
              const eqForLabel = curveIdx !== undefined ? settings.equations[curveIdx] : undefined;
              const compress = !!eqForLabel?.useCompression;
              const xs = tx(label.x);
              const ys = ty(label.y);
              const exRaw = xs === null ? label.x : xs;
              const eyRaw = ys === null ? label.y : ys;
              const ey = compress ? compressYForCurve(eyRaw, exRaw) : eyRaw;
              const ex = compress ? compressXForCurve(exRaw) : exRaw;
              const x = scaledOriginX + ex * settings.unitSize;
              const y = scaledOriginY - ey * settings.unitSize;
              if (x < -20 || x > canvasWidth + 20 || y < -20 || y > canvasHeight + 20) return null;
              const sym = label.symbol || 'cross';
              const isAutoInt = (label as any).isAuto;
              const axis = (label as any).axis;

              return (
                <g key={label.id + idx}>
                  {/* Ticks for auto intercepts */}
                  {isAutoInt && axis === 'x' && (
                    <line x1={x} y1={y - 4} x2={x} y2={y + 4} stroke="#000" strokeWidth={settings.tickWidth} />
                  )}
                  {isAutoInt && axis === 'y' && (
                    <line x1={x - 4} y1={y} x2={x + 4} y2={y} stroke="#000" strokeWidth={settings.tickWidth} />
                  )}

                  {!isAutoInt && (sym === 'dot' ? (
                    <circle cx={x} cy={y} r="1.5" fill="#000" />
                  ) : (
                    <g stroke="#000" strokeWidth="0.75">
                      <line x1={x - 2} y1={y - 2} x2={x + 2} y2={y + 2} />
                      <line x1={x - 2} y1={y + 2} x2={x + 2} y2={y - 2} />
                    </g>
                  ))}
                  {label.text && (
                    <text 
                      x={x + 4} 
                      y={y - 4} 
                      fontSize={settings.customLabelFontSize} 
                      className="math-label" 
                      style={{ fontFamily: 'Times New Roman', fontStyle: 'italic' }}
                    >
                      {label.text}
                    </text>
                  )}
                </g>
              );
            })}
          </svg>
        </div>

        {/* Legend / Status Overlay */}
        <div className="absolute bottom-8 right-8 flex flex-col gap-2 max-w-[200px]">
          <div className="bg-white/90 backdrop-blur p-3 rounded-lg shadow-sm border border-stone-200 text-[10px] space-y-2">
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 rounded-full bg-stone-900"/>
              <span className="text-stone-500 font-bold uppercase tracking-tighter">Active Equations</span>
            </div>
            <div className="space-y-1 max-h-[100px] overflow-y-auto">
              {settings.equations.map(eq => (
                <p key={eq.id} className="font-mono text-stone-900 truncate" title={eq.expression}>{eq.expression}</p>
              ))}
            </div>
          </div>
        </div>
      </main>

      {/* VBA Backdrop/Modal */}
      <AnimatePresence>
        {vbaVisible && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center p-4 lg:p-12 bg-stone-900/60 backdrop-blur-sm"
            onClick={() => setVbaVisible(false)}
          >
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              onClick={(e) => e.stopPropagation()}
              className="bg-white w-full max-w-4xl max-h-full flex flex-col rounded-2xl shadow-2xl overflow-hidden"
            >
              <header className="p-6 border-b border-stone-100 flex items-center justify-between bg-stone-50">
                <div className="flex items-center gap-3">
                  <div className="bg-stone-900 p-2 rounded-lg text-white">
                    <Code2 size={24} />
                  </div>
                  <div>
                    <h2 className="text-xl font-bold text-stone-900">VBA Export</h2>
                    <p className="text-xs text-stone-500 italic">Copy this code into a Word Layout Module</p>
                  </div>
                </div>
                <button 
                  onClick={() => setVbaVisible(false)}
                  className="text-stone-400 hover:text-stone-900 transition-colors p-2"
                >
                  <RefreshCcw size={20} />
                </button>
              </header>
              <div className="flex-1 overflow-y-auto p-6 bg-[#0a0a0a]">
                <pre className="code-block m-0 border-none shadow-none text-xs leading-relaxed">
                  {generatedVBA}
                </pre>
              </div>
              <footer className="p-4 bg-stone-50 border-t border-stone-100 flex justify-end gap-3">
                <button 
                  onClick={() => setVbaVisible(false)}
                  className="px-6 py-2 text-xs font-bold uppercase tracking-widest text-stone-500 hover:text-stone-900 transition-colors"
                >
                  Close
                </button>
                <button 
                  onClick={copyToClipboard}
                  className="bg-stone-900 text-white px-8 py-2 rounded-lg text-xs font-bold uppercase tracking-widest hover:bg-stone-800 transition-all flex items-center gap-2"
                >
                  <Download size={16} />
                  Copy Script
                </button>
              </footer>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

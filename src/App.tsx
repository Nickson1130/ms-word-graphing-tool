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

  const canvasWidth = (nXMax - nXMin) * settings.unitSize;
  const canvasHeight = (nYMax - nYMin) * settings.unitSize;

  // Origin position in SVG pixels
  const originX = -nXMin * settings.unitSize;
  const originY = nYMax * settings.unitSize;

  // Coordinate conversion
  const toPoints = (x: number, y: number) => ({
    x: originX + x * settings.unitSize,
    y: originY - y * settings.unitSize,
  });

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
            const samples = Math.max(2, Math.ceil((iXMax - iXMin) / Math.max(0.0001, settings.samplingStep)) + 1);
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
          // Convert to LHS - RHS
          const lhs = parts[0] || '0';
          const rhs = parts[1] || '0';
          const compiled = math.parse(`(${lhs}) - (${rhs})`).compile();
          
          eq.intervals.forEach(interval => {
            const iXMin = interval.useCustomDomain ? (parseFloat(interval.xMin) || 0) : nXMin;
            const iXMax = interval.useCustomDomain ? (parseFloat(interval.xMax) || 0) : nXMax;
            const iYMin = interval.useCustomRange ? (parseFloat(interval.yMin) || 0) : nYMin;
            const iYMax = interval.useCustomRange ? (parseFloat(interval.yMax) || 0) : nYMax;

            // Use a high-resolution grid for smooth implicit plotting
            const gridX = 300;
            const gridY = 300;
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
  }, [settings.equations, nXMin, nXMax, nYMin, nYMax, settings.samplingStep, mathScope]);

  // Combined Intercepts for all curves
  const intercepts = useMemo(() => {
    if (!settings.showXIntercepts && !settings.showYIntercepts) return [];
    
    let allIntercepts: (CustomLabel & { isAuto: boolean; axis: 'x' | 'y' })[] = [];
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
              const xInRange = segment.some(p => Math.abs(p.x) < settings.samplingStep * 2);
              if (xInRange && Math.abs(Number(y0)) > 0.001) {
                 allIntercepts.push({ 
                   id: `y-int-${cIdx}-${sIdx}`, 
                   x: 0, 
                   y: Number(Number(y0).toFixed(2)), 
                   text: '', // No numbers for auto intercepts
                   symbol: 'dot',
                   isAuto: true,
                   axis: 'y'
                 });
              }
            }
          }
        } catch {}

        // Generic X and Y Intercepts from segments
        // Skip tick detection entirely if the equation itself is x=0 or y=0
        const isXEquals0 = originalExpr.replace(/\s/g, '') === 'x=0';
        const isYEquals0 = originalExpr.replace(/\s/g, '') === 'y=0';

        for (let i = 0; i < segment.length - 1; i++) {
          const p1 = segment[i];
          const p2 = segment[i+1];

          // X-Intercept (y crosses 0)
          if (!isYEquals0 && settings.showXIntercepts && ((p1.y >= 0 && p2.y <= 0) || (p1.y <= 0 && p2.y >= 0))) {
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
                axis: 'x'
              });
            }
          }

          // Y-Intercept (x crosses 0)
          if (!isXEquals0 && settings.showYIntercepts && ((p1.x >= 0 && p2.x <= 0) || (p1.x <= 0 && p2.x >= 0))) {
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
                axis: 'y'
              });
            }
          }
        }
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
          return i.x > nXMin + 0.01 && i.x < nXMax - 0.01;
        }
        // Hide if at or outside y-axis boundaries (approx)
        if (Math.abs(i.x) < 0.001) {
          return i.y > nYMin + 0.01 && i.y < nYMax - 0.01;
        }
        return true;
      });
  }, [allCurvePoints, settings.showXIntercepts, settings.showYIntercepts, settings.equations, settings.samplingStep, nXMin, nXMax, nYMin, nYMax]);

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
    const jsOriginX = 100 + (-nXMin * settings.unitSize);
    const jsOriginY = 100 + (nYMax * settings.unitSize);

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
    
    ' --- 1. Draw Axes ---
    debugStep = "Drawing Axes"
    Dim xAxis As Shape, yAxis As Shape
    
    If ${settings.showXAxis ? 'True' : 'False'} Then
        ' Using AddLine instead of AddConnector for maximum compatibility
        Set xAxis = doc.Shapes.AddLine(originX + (xMin * unitSize), originY, originX + (xMax * unitSize), originY)
        xAxis.Line.EndArrowheadStyle = ${arrowInt}
        xAxis.Line.Weight = 0.75: xAxis.Line.ForeColor.RGB = 0
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = xAxis.Name
    End If

    If ${settings.showYAxis ? 'True' : 'False'} Then
        Set yAxis = doc.Shapes.AddLine(originX, originY - (yMin * unitSize), originX, originY - (yMax * unitSize))
        yAxis.Line.EndArrowheadStyle = ${arrowInt}
        yAxis.Line.Weight = 0.75: yAxis.Line.ForeColor.RGB = 0
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = yAxis.Name
    End If

    ' --- 2. Draw Ticks & Labels ---
    If ${settings.showTicks ? 'True' : 'False'} Then
        debugStep = "Drawing Ticks and Labels"
        Dim val As Variant, tick As Shape, lbl As Shape
        ' X Ticks
        Dim xTickVals: xTickVals = Array(${xTicks.length > 0 ? xTicks.join(', ') : ''})
        If UBound(xTickVals) >= LBound(xTickVals) Then
            For Each val In xTickVals
                Set tick = doc.Shapes.AddLine(originX + (val * unitSize), originY - 4, originX + (val * unitSize), originY + 4)
                tick.Line.ForeColor.RGB = 0: tick.Line.Weight = ${settings.tickWidth}
                shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name
            Next val
        End If
        
        ' Y Ticks
        Dim yTickVals: yTickVals = Array(${yTicks.length > 0 ? yTicks.join(', ') : ''})
        If UBound(yTickVals) >= LBound(yTickVals) Then
            For Each val In yTickVals
                Set tick = doc.Shapes.AddLine(originX - 4, originY - (val * unitSize), originX + 4, originY - (val * unitSize))
                tick.Line.ForeColor.RGB = RGB(0, 0, 0): tick.Line.Weight = ${settings.tickWidth}
                shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name
            Next val
        End If
    End If

    ' --- 2.5 Auto Intercept Ticks ---
    If ${settings.showXIntercepts || settings.showYIntercepts ? 'True' : 'False'} Then
        debugStep = "Drawing Intercept Ticks"
        ${intercepts.map(i => {
           if (i.y === 0) { // X-intercept
             return `
        Set tick = doc.Shapes.AddLine(originX + (${i.x} * unitSize), originY - 4, originX + (${i.x} * unitSize), originY + 4)
        tick.Line.ForeColor.RGB = 0: tick.Line.Weight = ${settings.tickWidth}
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = tick.Name`;
           } else { // Y-intercept
             return `
        Set tick = doc.Shapes.AddLine(originX - 4, originY - (${i.y} * unitSize), originX + 4, originY - (${i.y} * unitSize))
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
    mx${lIdx} = originX + (${l.x} * unitSize)
    my${lIdx} = originY - (${l.y} * unitSize)
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
    lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Size = 10
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
        lbl.TextFrame.TextRange.Font.Size = 12
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name
    End If

    ' --- 4. Draw Curves ---
    Dim fb As FreeformBuilder, curveComp As Shape
    ${allCurvePoints.map((curve, idx) => {
      const dashInt = VBA_CONSTANTS[curve.dashStyle] || 1;
      return curve.segments.map((segment, sIdx) => {
        if (segment.length < 2) return '';
        const pointBuild = segment.map((p, pIdx) => {
          const xPos = jsOriginX + (p.x * settings.unitSize);
          const yPos = jsOriginY - (p.y * settings.unitSize);
          if (pIdx === 0) return `Set fb = doc.Shapes.BuildFreeform(1, ${xPos}, ${yPos})`;
          return `fb.AddNodes 0, 1, ${xPos}, ${yPos}`;
        }).join('\n    ');
        
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
        Set lbl = doc.Shapes.AddTextbox(1, originX + (xMax * unitSize) - 7, originY - 6, 30, 30)
        lbl.Fill.Visible = 0: lbl.Line.Visible = 0
        lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        lbl.TextFrame.WordWrap = 0
        lbl.TextFrame.TextRange.Text = " " & "${settings.xAxisLabel}"
        lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Italic = True: lbl.TextFrame.TextRange.Font.Size = 12
        lbl.TextFrame.TextRange.ParagraphFormat.Alignment = 1 ' Center
        shpCount = shpCount + 1: ReDim Preserve shpArray(1 To shpCount): shpArray(shpCount) = lbl.Name
    End If

    If ${settings.showYLabel ? 'True' : 'False'} Then
        ' Y Label: Restored to original position with additional -5 shift total
        Set lbl = doc.Shapes.AddTextbox(1, originX - 16, originY - (yMax * unitSize) - 9, 30, 30)
        lbl.Fill.Visible = 0: lbl.Line.Visible = 0
        lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0: lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        lbl.TextFrame.WordWrap = 0
        lbl.TextFrame.TextRange.Text = " " & "${settings.yAxisLabel}"
        lbl.TextFrame.TextRange.Font.Name = "Times New Roman": lbl.TextFrame.TextRange.Font.Italic = True: lbl.TextFrame.TextRange.Font.Size = 12
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
  }, [settings, nXMin, nXMax, nYMin, nYMax, allCurvePoints, xTicks, yTicks, intercepts]);

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
                <span className="text-sm text-stone-600">Sampling Step</span>
                <select 
                  value={settings.samplingStep}
                  onChange={(e) => setSettings({ ...settings, samplingStep: Number(e.target.value) })}
                  className="bg-stone-50 border border-stone-200 p-1 rounded text-xs"
                >
                  <option value={0.01}>0.01 (Precise)</option>
                  <option value={0.001}>0.001 (Micro)</option>
                </select>
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
            {allCurvePoints.map((curve) => (
              <g key={curve.id}>
                {curve.segments.map((segment, sIdx) => {
                  const pts = segment.map(p => toPoints(p.x, p.y));
                  let d = '';
                  if (pts.length === 0) {
                    d = '';
                  } else if (pts.length === 1) {
                    d = `M ${pts[0].x} ${pts[0].y}`;
                  } else if (pts.length === 2) {
                    d = `M ${pts[0].x} ${pts[0].y} L ${pts[1].x} ${pts[1].y}`;
                  } else {
                    // Catmull-Rom to cubic bezier for smooth curves
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
              <text x={toPoints(settings.xMax, 0).x + 10} y={toPoints(settings.xMax, 0).y + 15} fontSize="12" className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>{settings.xAxisLabel}</text>
            )}
            {settings.showYLabel && (
              <text x={toPoints(0, settings.yMax).x - 15} y={toPoints(0, settings.yMax).y - 12} fontSize="12" className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>{settings.yAxisLabel}</text>
            )}

            {/* Origin */}
            {settings.showOrigin && (
              <text x={originX - 12} y={originY + 12} fontSize="11" className="math-label italic" style={{ fontFamily: 'Times New Roman' }}>O</text>
            )}

            {/* Custom Labels & Detected Intercepts */}
            {[...intercepts, ...settings.customLabels].map((label, idx) => {
              const { x, y } = toPoints(label.x, label.y);
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
                      fontSize="10" 
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

import React, { useState, useMemo, useRef } from "react";
import { Plus, Minus, AlertTriangle, TrendingDown, Info, Anchor, Upload, Check, X } from "lucide-react";
import * as XLSX from "xlsx";

const fmt = (n) => Math.round(n).toLocaleString("zh-TW");
const fmt1 = (n) => (Math.round(n * 10) / 10).toLocaleString("zh-TW");

const STYLE = `
.pas-root {
  --bg-page: #F2F2F7;
  --bg-card: #FFFFFF;
  --bg-fill: #F2F2F7;
  --sep: rgba(60,60,67,0.22);
  --text-1: #1C1C1E;
  --text-2: #8E8E93;
  --text-3: #C7C7CC;
  --blue: #007AFF;
  --blue-tint: #E8F2FF;
  --green: #34C759;
  --green-tint: #E8F9EC;
  --orange: #FF9500;
  --orange-tint: #FFF2E0;
  --purple: #AF52DE;
  --purple-tint: #F5E9FB;
  --red: #FF3B30;
  --red-tint: #FFECEB;
  --teal: #30B0C7;
  --teal-tint: #E4F6F9;
  font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang TC", "Helvetica Neue", Arial, sans-serif;
  background: var(--bg-page);
  border-radius: 24px;
  padding: 28px;
}
@media (prefers-color-scheme: dark) {
  .pas-root {
    --bg-page: #000000;
    --bg-card: #1C1C1E;
    --bg-fill: #2C2C2E;
    --sep: rgba(84,84,88,0.55);
    --text-1: #F2F2F7;
    --text-2: #98989D;
    --text-3: #48484A;
    --blue: #0A84FF; --blue-tint: rgba(10,132,255,0.16);
    --green: #30D158; --green-tint: rgba(48,209,88,0.16);
    --orange: #FF9F0A; --orange-tint: rgba(255,159,10,0.16);
    --purple: #BF5AF2; --purple-tint: rgba(191,90,242,0.16);
    --red: #FF453A; --red-tint: rgba(255,69,58,0.16);
    --teal: #40C8E0; --teal-tint: rgba(64,200,224,0.16);
  }
}
.pas-root * { box-sizing: border-box; }
.pas-card {
  background: var(--bg-card);
  border-radius: 14px;
  overflow: hidden;
}
.pas-row {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 14px;
  border-bottom: 0.5px solid var(--sep);
}
.pas-row:last-child { border-bottom: none; }
.pas-input {
  border: none;
  background: var(--bg-fill);
  border-radius: 8px;
  padding: 6px 9px;
  font-size: 14px;
  color: var(--text-1);
  font-family: inherit;
  outline: none;
}
.pas-input:focus { box-shadow: 0 0 0 3px var(--blue-tint); }
.pas-iconbtn {
  width: 24px; height: 24px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  border: none; cursor: pointer; flex-shrink: 0;
}
.pas-add { background: var(--blue-tint); color: var(--blue); }
.pas-remove { background: var(--red-tint); color: var(--red); }
.pas-tile {
  width: 30px; height: 30px; border-radius: 8px;
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.pas-eyebrow {
  font-size: 12px; font-weight: 600; letter-spacing: 0.02em;
  color: var(--text-2); text-transform: uppercase;
  margin: 0 0 8px 4px;
}
`;

const VENDOR_TILE = [
  { key: "blue" }, { key: "orange" }, { key: "green" }, { key: "purple" }, { key: "red" }, { key: "teal" },
];

function vendorVar(idx, suffix) { return `var(--${VENDOR_TILE[idx % VENDOR_TILE.length].key}${suffix})`; }

export default function ProcurementAllocationSystem() {
  const idRef = useRef(100);
  const nextId = () => ++idRef.current;

  const [vendors, setVendors] = useState([
    { id: 1, name: "冠緯", capKg: "" },
    { id: 2, name: "合眾", capKg: "" },
  ]);

  const [specs, setSpecs] = useState([
    { id: 1, name: "30/40", usage: 3900 },
    { id: 2, name: "40/50", usage: 2812 },
    { id: 3, name: "50/60", usage: 7238.5 },
    { id: 4, name: "60/70", usage: 5770 },
    { id: 5, name: "70/90", usage: 5070 },
    { id: 6, name: "90/120", usage: 23485 },
    { id: 7, name: "120/150", usage: 3622.5 },
    { id: 8, name: "100/200", usage: 31075.5 },
    { id: 9, name: "200/300", usage: 2572 },
    { id: 10, name: "300/500", usage: 0 },
  ]);

  const [quotes, setQuotes] = useState(() => {
    const priceTable = {
      1: { 1: 380, 2: 380 }, 2: { 1: 370, 2: 370 }, 3: { 1: 360, 2: 350 },
      4: { 1: 340, 2: 330 }, 5: { 1: 330, 2: 310 }, 6: { 1: 310, 2: 285 },
      7: { 1: 290, 2: 265 }, 8: { 1: 275, 2: 260 }, 9: { 1: 255, 2: 240 },
      10: { 1: 240 },
    };
    const ratioTable = {
      1: { 2: 0 }, 2: { 2: 0 }, 3: { 2: 50 }, 4: { 2: 50 }, 5: { 2: 80 },
      6: { 2: 80 }, 7: { 2: 50 }, 8: { 2: 80 }, 9: { 2: 50 },
    };
    const q = {};
    for (const s of [1,2,3,4,5,6,7,8,9,10]) {
      q[s] = {};
      for (const v of [1,2]) {
        const price = priceTable[s]?.[v];
        if (price !== undefined) q[s][v] = { price: String(price), ratio: ratioTable[s]?.[v] ?? 100 };
      }
    }
    return q;
  });

  const addVendor = () => { const id = nextId(); setVendors((v) => [...v, { id, name: `廠商${v.length + 1}`, capKg: "" }]); };
  const removeVendor = (id) => {
    setVendors((v) => v.filter((x) => x.id !== id));
    setQuotes((q) => { const nq = { ...q }; for (const s of Object.keys(nq)) { const row = { ...nq[s] }; delete row[id]; nq[s] = row; } return nq; });
  };
  const updateVendor = (id, field, val) => setVendors((v) => v.map((x) => (x.id === id ? { ...x, [field]: val } : x)));

  const addSpec = () => { const id = nextId(); setSpecs((s) => [...s, { id, name: "新規格", usage: 0 }]); setQuotes((q) => ({ ...q, [id]: {} })); };
  const removeSpec = (id) => { setSpecs((s) => s.filter((x) => x.id !== id)); setQuotes((q) => { const nq = { ...q }; delete nq[id]; return nq; }); };
  const updateSpec = (id, field, val) => setSpecs((s) => s.map((x) => (x.id === id ? { ...x, [field]: val } : x)));

  const updateQuote = (specId, vendorId, field, val) => {
    setQuotes((q) => {
      const specRow = { ...(q[specId] || {}) };
      const cell = specRow[vendorId] || { price: "", ratio: 100 };
      specRow[vendorId] = { ...cell, [field]: val };

      if (field === "ratio") {
        const activeIds = Object.keys(specRow)
          .map(Number)
          .filter((vid) => specRow[vid].price !== "" && specRow[vid].price !== undefined && specRow[vid].price !== null);
        const others = activeIds.filter((vid) => vid !== vendorId);
        const newVal = Math.max(0, Math.min(100, Number(val) || 0));
        specRow[vendorId] = { ...specRow[vendorId], ratio: newVal };
        const remaining = 100 - newVal;
        if (others.length > 0) {
          const currentSum = others.reduce((sum, vid) => sum + (Number(specRow[vid].ratio) || 0), 0);
          let allocated = 0;
          others.forEach((vid, i) => {
            let share;
            if (i === others.length - 1) {
              share = remaining - allocated; // last vendor takes the remainder so the total is exactly 100
            } else if (currentSum > 0) {
              share = Math.round(remaining * ((Number(specRow[vid].ratio) || 0) / currentSum));
            } else {
              share = Math.round(remaining / others.length);
            }
            share = Math.max(0, Math.min(100, share));
            allocated += share;
            specRow[vid] = { ...specRow[vid], ratio: share };
          });
        }
      }
      return { ...q, [specId]: specRow };
    });
  };

  // ---- ERP file import (xlsx/csv, parsed client-side) ----
  const fileInputRef = useRef(null);
  const [importStage, setImportStage] = useState(null); // null | 'detected' | 'manual' | 'preview'
  const [rawImport, setRawImport] = useState(null); // {aoa, headerIdx, header, fileName}
  const [detection, setDetection] = useState(null); // {mode, confidence, reportType, balanceInfo}
  const [specColSel, setSpecColSel] = useState(-1);
  const [qtyColsSel, setQtyColsSel] = useState([]); // array of column indices to SUM together
  const [importPreview, setImportPreview] = useState(null); // [{name, usage}]
  const [importInfo, setImportInfo] = useState(null); // {fileName, matchedRows, skippedRows}
  const [importError, setImportError] = useState(null);

  const extractGrade = (text) => {
    if (!text) return null;
    const m = String(text).match(/(\d{2,3})\s*[\/\-]\s*(\d{2,3})/);
    return m ? `${m[1]}/${m[2]}` : null;
  };

  // Score + detect which columns are 規格 and 耗用量(KG)，優先用進耗存表的勾稽關係驗證
  // norm() strips whitespace so headers like "規　　格"（ERP排版常見的全形空格）也能比對到
  const norm = (s) => String(s).replace(/\s+/g, "");
  // strictNum: parseFloat 對 "502白仁" 這種文字會誤讀成 502（只看開頭數字），這裡要求整個字串就是數字才算數
  const strictNum = (v) => {
    const s = String(v).trim().replace(/,/g, "");
    if (s === "") return null;
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  };

  const detectColumns = (aoa, headerIdx, header) => {
    const dataRows = aoa.slice(headerIdx + 1).filter((r) => r && !r.every((c) => c === ""));
    const normHeader = header.map(norm);

    // 1) spec column
    let specCol = normHeader.findIndex((h) => h.includes("規格"));
    if (specCol === -1) specCol = normHeader.findIndex((h) => h.includes("品名"));
    if (specCol === -1) {
      let best = { idx: -1, ratio: 0 };
      header.forEach((h, idx) => {
        const ratio = dataRows.filter((r) => extractGrade(r[idx])).length / Math.max(dataRows.length, 1);
        if (ratio > best.ratio) best = { idx, ratio };
      });
      specCol = best.idx;
    }

    // 2) parse numeric columns (excluding spec col and obvious "包裝"件數欄，那些單位是件不是KG)
    const numericCols = header
      .map((h, idx) => {
        if (idx === specCol) return null;
        if (norm(h).includes("包裝")) return null;
        const vals = dataRows.map((r) => strictNum(r[idx]));
        const validCount = vals.filter((v) => v !== null).length;
        const numericRatio = validCount / Math.max(dataRows.length, 1);
        return numericRatio > 0.5 ? { idx, header: h, normHeader: norm(h), vals, numericRatio } : null;
      })
      .filter(Boolean);
    const numericColIdxs = numericCols.map((c) => c.idx);

    // 3) try to find 期初+進貨-耗用=期末 relationship across 4 numeric columns (進耗存表 signature)
    let balanceResult = null;
    const n = numericCols.length;
    if (n >= 4 && n <= 9) {
      outer: for (let a = 0; a < n; a++) {
        for (let b = 0; b < n; b++) {
          if (b === a) continue;
          for (let c = 0; c < n; c++) {
            if (c === a || c === b) continue;
            for (let d = 0; d < n; d++) {
              if (d === a || d === b || d === c) continue;
              let matches = 0, total = 0;
              for (let r = 0; r < dataRows.length; r++) {
                const begin = numericCols[a].vals[r], purchase = numericCols[b].vals[r];
                const usage = numericCols[c].vals[r], end = numericCols[d].vals[r];
                if (begin == null || purchase == null || usage == null || end == null) continue;
                total++;
                if (Math.abs((begin + purchase - usage) - end) <= Math.max(1, Math.abs(end) * 0.02)) matches++;
              }
              if (total >= 5 && matches / total > 0.85) {
                balanceResult = { usageCol: numericCols[c].idx, matchRate: matches / total, total };
                break outer;
              }
            }
          }
        }
      }
    }

    if (balanceResult) {
      return {
        specCol, qtyCols: [balanceResult.usageCol], numericColIdxs,
        confidence: "high", mode: "balance", reportType: "進耗存表",
        detail: `已透過「期初結存＋本期進貨－本期耗用＝期末結存」的勾稽關係驗證（${dataRows.length}列中${balanceResult.total}列可比對，符合率${Math.round(balanceResult.matchRate * 100)}%）`,
      };
    }

    // 4) fallback: keyword scoring — 找出所有「消耗/流出」性質的欄位
    // CONFIRMED_KEYWORDS：已用真實進耗存表驗證過欄位命名方式，命中就直接給高信心
    // GUESS_KEYWORDS：其他ERP可能的講法，尚未驗證過，命中只給中等信心
    // DEFAULT_CHECK_KEYWORDS：使用者確認過的採購基準（領料＋銷貨），預設勾選加總；
    // 「出庫」雖然辨識得出來，但使用者說了公司自己也有進口一部分、且採購基準只看領料+銷貨，
    // 所以出庫不預設勾選，只在找不到領料/銷貨時作為備援
    const CONFIRMED_KEYWORDS = ["領料", "銷貨", "出庫"];
    const DEFAULT_CHECK_KEYWORDS = ["領料", "銷貨"];
    const GUESS_KEYWORDS = ["耗用", "領用", "銷售", "耗損", "消耗"];
    const USAGE_KEYWORDS = [...CONFIRMED_KEYWORDS, ...GUESS_KEYWORDS];
    const NEGATIVE_KEYWORDS = ["進貨", "入庫", "期初", "期末", "結存", "採購"];
    const scored = numericCols.map((col) => {
      let score = col.numericRatio * 5;
      const isConfirmed = CONFIRMED_KEYWORDS.some((k) => col.normHeader.includes(k));
      const isDefaultCheck = DEFAULT_CHECK_KEYWORDS.some((k) => col.normHeader.includes(k));
      const isUsage = isConfirmed || GUESS_KEYWORDS.some((k) => col.normHeader.includes(k));
      const isNegative = NEGATIVE_KEYWORDS.some((k) => col.normHeader.includes(k));
      if (isConfirmed) score += 150;
      else if (isUsage) score += 100;
      else if (col.normHeader.includes("數量")) score += 30;
      if (isNegative) score -= 80;
      const unitCol = col.idx + 1;
      if (unitCol < header.length) {
        const sampleVals = dataRows.slice(0, 20).map((r) => String(r[unitCol] || ""));
        if (sampleVals.some((v) => /kg|公斤/i.test(v))) score += 20;
      }
      return { ...col, score, isUsage, isConfirmed, isDefaultCheck, isNegative };
    });

    const matchedUsageCols = scored.filter((c) => c.isUsage && !c.isNegative);
    const defaultCheckCols = matchedUsageCols.filter((c) => c.isDefaultCheck);
    let qtyCols, confidence, detail;
    if (matchedUsageCols.length > 0) {
      // 有領料/銷貨可用就只勾這兩個當採購基準；沒有的話（例如其他報表沒有這兩個欄位）才退回全部消耗性質欄位
      const useDefault = defaultCheckCols.length > 0;
      const chosen = useDefault ? defaultCheckCols : matchedUsageCols;
      qtyCols = chosen.map((c) => c.idx);
      const allConfirmed = chosen.every((c) => c.isConfirmed);
      confidence = allConfirmed ? "high" : "medium";
      detail = useDefault
        ? `依已驗證過的採購基準（領料＋銷貨）找到${qtyCols.length}個欄位（${qtyCols.map((i) => header[i]).join("、")}），預設加總；其餘消耗性質欄位（如出庫）可自行勾選加入`
        : `依欄位名稱找到${qtyCols.length}個消耗性質欄位（${qtyCols.map((i) => header[i]).join("、")}），部分為尚未驗證過的命名方式，建議確認樣本數值`;
    } else {
      const best = scored.reduce((a, b) => (b.score > a.score ? b : a), { score: -Infinity });
      qtyCols = best.idx !== undefined ? [best.idx] : [];
      confidence = "low";
      detail = "找不到明確的進耗存勾稽關係，也沒有明顯的「領料／銷貨／出庫」欄位名稱，此為系統最佳猜測，建議手動確認";
    }
    return { specCol, qtyCols, numericColIdxs, confidence, mode: "keyword", reportType: "單一數量表", detail };
  };

  const handleFileSelect = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImportError(null);
    try {
      const isCsv = /\.csv$/i.test(file.name);
      let wb;
      if (isCsv) {
        const text = await file.text();
        wb = XLSX.read(text, { type: "string" });
      } else {
        const buf = await file.arrayBuffer();
        wb = XLSX.read(buf, { type: "array" });
      }
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      let headerIdx = -1;
      for (let i = 0; i < Math.min(aoa.length, 15); i++) {
        const row = aoa[i].map((c) => norm(c));
        if (row.some((c) => c.includes("規格") || c.includes("品名"))) { headerIdx = i; break; }
      }
      if (headerIdx === -1) throw new Error("找不到包含「規格」或「品名」的表頭列，請確認檔案格式");

      const header = aoa[headerIdx].map((c) => String(c));
      const det = detectColumns(aoa, headerIdx, header);
      if (det.specCol === -1 || det.qtyCols.length === 0) throw new Error("無法自動判斷規格欄或數量欄，請確認檔案格式");

      setRawImport({ aoa, headerIdx, header, fileName: file.name });
      setDetection(det);
      setSpecColSel(det.specCol);
      setQtyColsSel(det.qtyCols);
      setImportStage("detected");
    } catch (err) {
      setImportError(err.message || "檔案解析失敗，請確認格式");
    }
  };

  const columnSample = (colIdx) => {
    if (!rawImport || colIdx < 0) return [];
    return rawImport.aoa.slice(rawImport.headerIdx + 1, rawImport.headerIdx + 4).map((r) => r[colIdx]).filter((v) => v !== "");
  };

  const toggleQtyCol = (idx) => {
    setQtyColsSel((prev) => (prev.includes(idx) ? prev.filter((i) => i !== idx) : [...prev, idx]));
  };

  const runAggregation = () => {
    if (!rawImport || specColSel < 0 || qtyColsSel.length === 0) { setImportError("請選擇規格欄，並至少勾選一個數量欄"); return; }
    const { aoa, headerIdx, fileName } = rawImport;
    const agg = {};
    let matched = 0, skipped = 0;
    for (let i = headerIdx + 1; i < aoa.length; i++) {
      const row = aoa[i];
      if (!row || row.every((c) => c === "")) continue;
      const specText = row[specColSel];
      let qty = 0, anyValid = false;
      for (const col of qtyColsSel) {
        const v = parseFloat(String(row[col]).replace(/,/g, ""));
        if (!isNaN(v)) { qty += v; anyValid = true; }
      }
      if (!anyValid) { skipped++; continue; }
      const grade = extractGrade(specText);
      if (!grade) { skipped++; continue; }
      agg[grade] = (agg[grade] || 0) + Math.abs(qty);
      matched++;
    }
    const rows = Object.entries(agg).map(([name, usage]) => ({ name, usage: Math.round(usage * 10) / 10 }));
    if (rows.length === 0) { setImportError("解析後沒有可用的規格數量，請換一組欄位試試"); return; }
    rows.sort((a, b) => b.usage - a.usage);
    setImportError(null);
    setImportPreview(rows);
    setImportInfo({ fileName, matchedRows: matched, skippedRows: skipped });
    setImportStage("preview");
  };

  const applyImport = () => {
    if (!importPreview) return;
    setSpecs((prev) => {
      const next = [...prev];
      const newIds = [];
      importPreview.forEach((row) => {
        const idx = next.findIndex((s) => s.name.trim() === row.name.trim());
        if (idx !== -1) next[idx] = { ...next[idx], usage: row.usage };
        else { const id = nextId(); next.push({ id, name: row.name, usage: row.usage }); newIds.push(id); }
      });
      if (newIds.length) setQuotes((q) => { const nq = { ...q }; newIds.forEach((id) => { nq[id] = {}; }); return nq; });
      return next;
    });
    setImportPreview(null);
    setImportInfo(null);
    setImportStage(null);
    setRawImport(null);
    setDetection(null);
  };

  const cancelImport = () => {
    setImportPreview(null);
    setImportInfo(null);
    setImportError(null);
    setImportStage(null);
    setRawImport(null);
    setDetection(null);
  };

  const vendorIdx = (vendorId) => vendors.findIndex((v) => v.id === vendorId);

  const result = useMemo(() => {
    const capRemaining = {};
    vendors.forEach((v) => { capRemaining[v.id] = v.capKg === "" || v.capKg === null ? Infinity : Number(v.capKg); });

    const specList = specs.map((s) => {
      const rows = vendors
        .map((v) => { const cell = quotes[s.id]?.[v.id]; if (!cell || cell.price === "" || cell.price === null) return null; return { vendorId: v.id, price: Number(cell.price), ratio: cell.ratio ?? 100 }; })
        .filter(Boolean).sort((a, b) => a.price - b.price);
      const savingPotential = rows.length >= 2 ? (rows[1].price - rows[0].price) * s.usage : 0;
      return { spec: s, rows, savingPotential };
    });
    specList.sort((a, b) => b.savingPotential - a.savingPotential);

    const allocations = {};
    const warnings = [];

    specList.forEach(({ spec, rows }) => {
      let remaining = spec.usage;
      const allocRows = [];
      if (rows.length === 0) { if (spec.usage > 0) warnings.push({ type: "no-quote", spec: spec.name }); allocations[spec.id] = []; return; }
      for (const r of rows) {
        if (remaining <= 0) break;
        const ratioCapKg = spec.usage * ((r.ratio ?? 100) / 100);
        const avail = Math.min(remaining, ratioCapKg, capRemaining[r.vendorId]);
        if (avail > 0.0001) { allocRows.push({ vendorId: r.vendorId, qty: avail, price: r.price, exceeded: false }); remaining -= avail; capRemaining[r.vendorId] -= avail; }
      }
      if (remaining > 0.0001) {
        for (const r of rows) {
          if (remaining <= 0.0001) break;
          const avail = Math.min(remaining, capRemaining[r.vendorId]);
          if (avail > 0.0001) {
            const existing = allocRows.find((a) => a.vendorId === r.vendorId);
            if (existing) { existing.qty += avail; existing.exceeded = true; } else allocRows.push({ vendorId: r.vendorId, qty: avail, price: r.price, exceeded: true });
            remaining -= avail; capRemaining[r.vendorId] -= avail;
          }
        }
      }
      if (remaining > 0.0001) warnings.push({ type: "shortfall", spec: spec.name, qty: remaining });
      allocations[spec.id] = allocRows;
    });

    const vendorTotals = {};
    vendors.forEach((v) => (vendorTotals[v.id] = { qty: 0, cost: 0 }));
    let totalCost = 0, totalQty = 0;
    Object.values(allocations).forEach((rows) => { rows.forEach((r) => { vendorTotals[r.vendorId].qty += r.qty; vendorTotals[r.vendorId].cost += r.qty * r.price; totalCost += r.qty * r.price; totalQty += r.qty; }); });

    let minCost = 0;
    specs.forEach((s) => {
      const rows = vendors.map((v) => { const cell = quotes[s.id]?.[v.id]; return cell && cell.price !== "" ? Number(cell.price) : null; }).filter((p) => p !== null);
      if (rows.length > 0) minCost += Math.min(...rows) * s.usage;
    });

    return { allocations, vendorTotals, totalCost, totalQty, minCost, warnings };
  }, [specs, vendors, quotes]);

  const extraCost = result.totalCost - result.minCost;
  const savingsPct = result.minCost > 0 ? (extraCost / result.minCost) * 100 : 0;

  return (
    <div className="pas-root">
      <style>{STYLE}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <div style={{ width: 44, height: 44, borderRadius: 12, background: "var(--blue)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Anchor size={22} color="#fff" aria-hidden="true" />
        </div>
        <div>
          <div style={{ fontSize: 22, fontWeight: 700, color: "var(--text-1)", letterSpacing: "-0.01em" }}>採購分配系統</div>
          <div style={{ fontSize: 13, color: "var(--text-2)" }}>規格用量 × 多廠商報價 × 風險上限</div>
        </div>
      </div>

      {/* Import from ERP export */}
      <div style={{ marginBottom: 22 }}>
        <div className="pas-eyebrow">資料匯入</div>
        <div className="pas-card" style={{ padding: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div className="pas-tile" style={{ background: "var(--teal-tint)" }}>
              <Upload size={15} color="var(--teal)" aria-hidden="true" />
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14, fontWeight: 600, color: "var(--text-1)" }}>從 Excel / CSV 匯入年度用量</div>
              <div style={{ fontSize: 12, color: "var(--text-2)" }}>支援鼎新等ERP匯出的領料量或銷售量報表，自動依規格加總</div>
            </div>
            <button onClick={() => fileInputRef.current?.click()} className="pas-input" style={{ background: "var(--blue)", color: "#fff", fontWeight: 600, cursor: "pointer", border: "none" }}>
              選擇檔案
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFileSelect} style={{ display: "none" }} />
          </div>
          {importError && (
            <div style={{ marginTop: 10, fontSize: 12.5, color: "var(--red)" }}>{importError}</div>
          )}
          <div style={{ fontSize: 11.5, color: "var(--text-2)", marginTop: 10 }}>
            若匯出檔為 PDF（無法在此直接解析），請直接把檔案傳給我，我會幫您整理成數字後貼進這裡。
          </div>
        </div>

        {importStage === "detected" && rawImport && detection && (
          <div className="pas-card" style={{ marginTop: 10, padding: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
              <span
                style={{
                  fontSize: 11, fontWeight: 600, borderRadius: 5, padding: "2px 8px",
                  background: detection.confidence === "high" ? "var(--green-tint)" : detection.confidence === "medium" ? "var(--orange-tint)" : "var(--red-tint)",
                  color: detection.confidence === "high" ? "var(--green)" : detection.confidence === "medium" ? "var(--orange)" : "var(--red)",
                }}
              >
                信心度：{detection.confidence === "high" ? "高" : detection.confidence === "medium" ? "中" : "低"}
              </span>
              <span style={{ fontSize: 13, fontWeight: 600, color: "var(--text-1)" }}>
                偵測為{detection.reportType}：{rawImport.fileName}
              </span>
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-2)", marginBottom: 12 }}>{detection.detail}</div>
            <div style={{ marginBottom: 4 }}>
              <div style={{ fontSize: 11, color: "var(--text-2)", marginBottom: 4 }}>規格欄</div>
              <select value={specColSel} onChange={(e) => setSpecColSel(Number(e.target.value))} className="pas-input" style={{ width: "100%", marginBottom: 12 }}>
                {rawImport.header.map((h, i) => <option key={i} value={i}>{`${h || "(無標題)"}`}</option>)}
              </select>
              <div style={{ fontSize: 11, color: "var(--text-2)", marginBottom: 4 }}>耗用量欄（KG）— 已勾選的會加總，可自行調整</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {(detection.numericColIdxs || []).map((i) => {
                  const h = rawImport.header[i];
                  const sample = columnSample(i);
                  return (
                    <label key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "var(--text-1)", cursor: "pointer" }}>
                      <input type="checkbox" checked={qtyColsSel.includes(i)} onChange={() => toggleQtyCol(i)} />
                      <span style={{ fontWeight: qtyColsSel.includes(i) ? 600 : 400 }}>{h || "(無標題)"}</span>
                      <span style={{ color: "var(--text-2)", fontSize: 11 }}>樣本：{sample.slice(0, 3).join("、") || "—"}</span>
                    </label>
                  );
                })}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
              <button onClick={runAggregation} style={{ display: "flex", alignItems: "center", gap: 5, background: "var(--blue)", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                <Check size={14} aria-hidden="true" /> 看起來正確，確認解析
              </button>
              <button onClick={cancelImport} style={{ display: "flex", alignItems: "center", gap: 5, background: "var(--bg-fill)", color: "var(--text-2)", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                <X size={14} aria-hidden="true" /> 取消
              </button>
            </div>
          </div>
        )}

        {importStage === "preview" && importPreview && (
          <div className="pas-card" style={{ marginTop: 10, padding: 14 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: "var(--text-1)" }}>
              解析結果：{importInfo?.fileName}
            </div>
            <div style={{ fontSize: 12, color: "var(--text-2)", marginBottom: 10 }}>
              成功比對 {importInfo?.matchedRows} 列，略過 {importInfo?.skippedRows} 列（無法辨識規格或數量）
            </div>
            <div style={{ maxHeight: 220, overflowY: "auto", borderRadius: 8, background: "var(--bg-fill)" }}>
              {importPreview.map((row, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "6px 10px", fontSize: 13, borderBottom: i === importPreview.length - 1 ? "none" : "0.5px solid var(--sep)" }}>
                  <span style={{ color: "var(--text-1)" }}>{row.name}</span>
                  <span style={{ color: "var(--text-2)" }}>{row.usage.toLocaleString("zh-TW")} KG</span>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button onClick={applyImport} style={{ display: "flex", alignItems: "center", gap: 5, background: "var(--blue)", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                <Check size={14} aria-hidden="true" /> 套用到規格表
              </button>
              <button onClick={cancelImport} style={{ display: "flex", alignItems: "center", gap: 5, background: "var(--bg-fill)", color: "var(--text-1)", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                <X size={14} aria-hidden="true" /> 取消
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Section 1: Vendors */}
      <div style={{ marginBottom: 22 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div className="pas-eyebrow">廠商</div>
          <button onClick={addVendor} className="pas-iconbtn pas-add" aria-label="新增廠商"><Plus size={14} strokeWidth={2.5} aria-hidden="true" /></button>
        </div>
        <div className="pas-card">
          {vendors.map((v, i) => {
            const idx = vendorIdx(v.id);
            return (
              <div key={v.id} className="pas-row">
                <div className="pas-tile" style={{ background: vendorVar(idx, "-tint") }}>
                  <span style={{ width: 10, height: 10, borderRadius: "50%", background: vendorVar(idx, "") }} />
                </div>
                <input value={v.name} onChange={(e) => updateVendor(v.id, "name", e.target.value)} className="pas-input" style={{ width: 90, fontWeight: 600 }} />
                <input placeholder="年產能上限 KG" value={v.capKg} onChange={(e) => updateVendor(v.id, "capKg", e.target.value)} className="pas-input" style={{ width: 130, marginLeft: "auto", textAlign: "right" }} />
                {vendors.length > 1 && (
                  <button onClick={() => removeVendor(v.id)} className="pas-iconbtn pas-remove" aria-label={`移除${v.name}`}><Minus size={14} strokeWidth={2.5} aria-hidden="true" /></button>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Section 2: Specs + quotes */}
      <div style={{ marginBottom: 22 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div className="pas-eyebrow">規格・報價・分配上限</div>
          <button onClick={addSpec} className="pas-iconbtn pas-add" aria-label="新增規格"><Plus size={14} strokeWidth={2.5} aria-hidden="true" /></button>
        </div>
        <div className="pas-card" style={{ overflowX: "auto" }}>
          <div style={{ minWidth: 480 + vendors.length * 180 }}>
            <div className="pas-row" style={{ background: "var(--bg-fill)", borderBottom: "0.5px solid var(--sep)" }}>
              <div style={{ width: 96, fontSize: 12, fontWeight: 600, color: "var(--text-2)" }}>規格</div>
              <div style={{ width: 90, fontSize: 12, fontWeight: 600, color: "var(--text-2)" }}>用量(KG)</div>
              {vendors.map((v) => {
                const idx = vendorIdx(v.id);
                return (
                  <div key={v.id} style={{ width: 164, fontSize: 12, fontWeight: 600, color: vendorVar(idx, "") }}>
                    {v.name}<span style={{ fontWeight: 400, color: "var(--text-2)" }}> 報價 / 上限%</span>
                  </div>
                );
              })}
            </div>
            {specs.map((s) => (
              <div key={s.id} className="pas-row">
                <input value={s.name} onChange={(e) => updateSpec(s.id, "name", e.target.value)} className="pas-input" style={{ width: 88, fontWeight: 600 }} />
                <input type="number" value={s.usage} onChange={(e) => updateSpec(s.id, "usage", Number(e.target.value))} className="pas-input" style={{ width: 82 }} />
                {vendors.map((v) => {
                  const cell = quotes[s.id]?.[v.id] || { price: "", ratio: 100 };
                  return (
                    <div key={v.id} style={{ width: 164, display: "flex", gap: 6 }}>
                      <input type="number" placeholder="—" value={cell.price} onChange={(e) => updateQuote(s.id, v.id, "price", e.target.value)} className="pas-input" style={{ width: 66, textAlign: "right" }} />
                      <input type="number" min={0} max={100} value={cell.ratio} onChange={(e) => updateQuote(s.id, v.id, "ratio", Number(e.target.value))} className="pas-input" style={{ width: 68, textAlign: "right", padding: "6px 6px" }} disabled={cell.price === "" || cell.price === undefined} />
                    </div>
                  );
                })}
                <button onClick={() => removeSpec(s.id)} className="pas-iconbtn pas-remove" style={{ marginLeft: "auto" }} aria-label={`移除${s.name}`}><Minus size={14} strokeWidth={2.5} aria-hidden="true" /></button>
              </div>
            ))}
          </div>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-2)", marginTop: 8, marginLeft: 4 }}>上限% — 該規格內各廠商分配比例會自動連動，調整一家會等比例反映到其他有報價的廠商，加總維持100%</div>
      </div>

      {/* Warnings */}
      {result.warnings.length > 0 && (
        <div className="pas-card" style={{ marginBottom: 22, padding: "12px 14px", background: "var(--orange-tint)" }}>
          {result.warnings.map((w, i) => (
            <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: i === result.warnings.length - 1 ? 0 : 8 }}>
              <AlertTriangle size={15} color="var(--orange)" style={{ flexShrink: 0, marginTop: 1 }} aria-hidden="true" />
              <span style={{ fontSize: 13, color: "var(--text-1)" }}>
                {w.type === "shortfall" && <>{w.spec}：預估用量有 {fmt1(w.qty)} KG 無法被目前廠商供應涵蓋，需尋找備援廠商或提高上限。</>}
                {w.type === "no-quote" && <>{w.spec}：目前沒有任何廠商報價，無法分配。</>}
              </span>
            </div>
          ))}
        </div>
      )}

      {/* Section 3: Results */}
      <div className="pas-eyebrow">分配結果</div>

      <div style={{ display: "grid", gridTemplateColumns: `repeat(${Math.min(vendors.length, 3)}, minmax(0,1fr))`, gap: 12, marginBottom: 16 }}>
        {vendors.map((v) => {
          const t = result.vendorTotals[v.id] || { qty: 0, cost: 0 };
          const pct = result.totalQty > 0 ? (t.qty / result.totalQty) * 100 : 0;
          const idx = vendorIdx(v.id);
          return (
            <div key={v.id} className="pas-card" style={{ padding: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 10 }}>
                <div className="pas-tile" style={{ width: 24, height: 24, background: vendorVar(idx, "-tint") }}>
                  <span style={{ width: 8, height: 8, borderRadius: "50%", background: vendorVar(idx, "") }} />
                </div>
                <span style={{ fontSize: 13, color: "var(--text-2)", fontWeight: 600 }}>{v.name}</span>
              </div>
              <div style={{ fontSize: 24, fontWeight: 700, color: "var(--text-1)", letterSpacing: "-0.02em" }}>{fmt1(t.qty)}<span style={{ fontSize: 13, fontWeight: 500, color: "var(--text-2)" }}> KG</span></div>
              <div style={{ fontSize: 12.5, color: "var(--text-2)", marginTop: 3 }}>NT$ {fmt(t.cost)} · {fmt1(pct)}%</div>
              <div style={{ height: 5, borderRadius: 3, background: "var(--bg-fill)", marginTop: 10, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: vendorVar(idx, ""), borderRadius: 3 }} />
              </div>
            </div>
          );
        })}
      </div>

      <div className="pas-card" style={{ padding: 16, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 14 }}>
          <div className="pas-tile" style={{ width: 24, height: 24, background: "var(--purple-tint)" }}>
            <TrendingDown size={13} color="var(--purple)" aria-hidden="true" />
          </div>
          <span style={{ fontSize: 14, fontWeight: 600, color: "var(--text-1)" }}>成本比較</span>
        </div>

        {[
          { label: "目前方案總成本", value: result.totalCost, ref: Math.max(result.totalCost, result.minCost), color: "var(--blue)" },
          { label: "理論最低成本（忽略上限與產能）", value: result.minCost, ref: Math.max(result.totalCost, result.minCost), color: "var(--text-3)" },
        ].map((row, i) => (
          <div key={i} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 5 }}>
              <span style={{ color: "var(--text-2)" }}>{row.label}</span>
              <span style={{ fontWeight: 600, color: "var(--text-1)" }}>NT$ {fmt(row.value)}</span>
            </div>
            <div style={{ height: 8, borderRadius: 4, background: "var(--bg-fill)", overflow: "hidden" }}>
              <div style={{ width: `${row.ref > 0 ? (row.value / row.ref) * 100 : 0}%`, height: "100%", background: row.color, borderRadius: 4 }} />
            </div>
          </div>
        ))}

        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, paddingTop: 12, borderTop: "0.5px solid var(--sep)" }}>
          <span style={{ color: "var(--text-2)" }}>風險分散代價</span>
          <span style={{ fontWeight: 600, color: extraCost > 0 ? "var(--orange)" : "var(--text-1)" }}>
            NT$ {fmt(extraCost)}{result.minCost > 0 && <span style={{ fontWeight: 400, color: "var(--text-2)" }}> （+{fmt1(savingsPct)}%）</span>}
          </span>
        </div>
      </div>

      <div className="pas-eyebrow">逐規格分配明細</div>
      <div className="pas-card">
        {specs.map((s) => {
          const rows = result.allocations[s.id] || [];
          if (s.usage === 0 && rows.length === 0) return null;
          return (
            <div key={s.id} className="pas-row" style={{ flexDirection: "column", alignItems: "stretch", gap: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}>
                <span style={{ fontWeight: 600, color: "var(--text-1)" }}>{s.name}</span>
                <span style={{ color: "var(--text-2)" }}>{fmt1(s.usage)} KG</span>
              </div>
              <div style={{ display: "flex", height: 7, borderRadius: 4, overflow: "hidden", background: "var(--bg-fill)", gap: 1 }}>
                {rows.map((r, i) => {
                  const idx = vendorIdx(r.vendorId);
                  const vname = vendors.find((v) => v.id === r.vendorId)?.name;
                  const widthPct = s.usage > 0 ? (r.qty / s.usage) * 100 : 0;
                  return <div key={i} title={`${vname}: ${fmt1(r.qty)} KG`} style={{ width: `${widthPct}%`, background: vendorVar(idx, ""), minWidth: r.qty > 0 ? 2 : 0 }} />;
                })}
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {rows.map((r, i) => {
                  const idx = vendorIdx(r.vendorId);
                  const vname = vendors.find((v) => v.id === r.vendorId)?.name;
                  return (
                    <span key={i} style={{ fontSize: 12, color: vendorVar(idx, ""), background: vendorVar(idx, "-tint"), borderRadius: 6, padding: "3px 9px", fontWeight: 600 }}>
                      {vname} {fmt1(r.qty)}KG @{r.price}{r.exceeded && " · 超上限"}
                    </span>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ display: "flex", gap: 7, marginTop: 20, fontSize: 12, color: "var(--text-2)" }}>
        <Info size={14} style={{ flexShrink: 0, marginTop: 1 }} aria-hidden="true" />
        <span>分配邏輯：報價相同或僅一家報價時優先維持既有關係／唯一供應；有價差時依上限%決定各廠商最高承接量，優先讓價差最大的規格取得便宜產能；上限用盡仍不足時，會在產能範圍內超額分配並標示「超上限」；所有廠商產能都不足時，會標示缺口提示。</span>
      </div>
    </div>
  );
}

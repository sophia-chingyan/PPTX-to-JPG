import { useState, useRef, useCallback } from "react";

const FONT_LINK = "https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=DM+Sans:wght@300;400;500;600&display=swap";
if (!document.querySelector(`link[href*="DM+Sans"]`)) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = FONT_LINK;
  document.head.appendChild(link);
}

const RESOLUTIONS = [
  { label: "SD",  w: 960,  h: 540,  hint: "960×540"   },
  { label: "FHD", w: 1920, h: 1080, hint: "1920×1080" },
  { label: "2K",  w: 2560, h: 1440, hint: "2560×1440" },
  { label: "4K",  w: 3840, h: 2160, hint: "3840×2160" },
];
const PREVIEW_W = 480, PREVIEW_H = 270;

// ── Utilities ─────────────────────────────────────────────────────────────────

function loadJSZip() {
  return new Promise((resolve, reject) => {
    if (window.JSZip) return resolve();
    if (document.querySelector(`script[src*="jszip"]`)) {
      const t = setInterval(() => { if (window.JSZip) { clearInterval(t); resolve(); } }, 40);
      return;
    }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
}

async function parsePptx(file) {
  await loadJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1]) - parseInt(b.match(/slide(\d+)/)[1]));
  const mediaUrls = {};
  for (const key of Object.keys(zip.files))
    if (key.startsWith("ppt/media/"))
      mediaUrls[key] = URL.createObjectURL(await zip.files[key].async("blob"));
  const slides = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const xml   = await zip.files[slideFiles[i]].async("text");
    const num   = parseInt(slideFiles[i].match(/slide(\d+)/)[1]);
    const rPath = `ppt/slides/_rels/slide${num}.xml.rels`;
    const imgMap = {};
    if (zip.files[rPath]) {
      const rd = new DOMParser().parseFromString(await zip.files[rPath].async("text"), "text/xml");
      for (const r of rd.getElementsByTagName("Relationship")) {
        const t = r.getAttribute("Target");
        if (t?.includes("media/")) {
          const fp = "ppt/" + t.replace("../", "");
          if (mediaUrls[fp]) imgMap[r.getAttribute("Id")] = mediaUrls[fp];
        }
      }
    }
    let bgColor = null;
    const sd = new DOMParser().parseFromString(xml, "text/xml");
    const bg = sd.getElementsByTagName("p:bg")[0];
    if (bg) { const sc = bg.getElementsByTagName("a:srgbClr")[0]; if (sc) bgColor = "#" + sc.getAttribute("val"); }
    slides.push({ index: i, xml, images: imgMap, bgColor });
  }
  return slides;
}

function renderToCanvas(slide, w, h) {
  return new Promise(resolve => {
    const canvas = document.createElement("canvas");
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = slide.bgColor || "#ffffff";
    ctx.fillRect(0, 0, w, h);
    const doc = new DOMParser().parseFromString(slide.xml, "text/xml");
    const SX = w / 12192000, SY = h / 6858000;
    const eX = v => (parseInt(v) || 0) * SX;
    const eY = v => (parseInt(v) || 0) * SY;
    const szPx = sz => ((parseInt(sz) || 1800) / 100) * (w / 960);
    const textItems = [], imageItems = [];
    const shapes = [...Array.from(doc.getElementsByTagName("p:sp")), ...Array.from(doc.getElementsByTagName("p:pic"))];
    for (const sh of shapes) {
      const off = sh.getElementsByTagName("a:off")[0];
      const ext = sh.getElementsByTagName("a:ext")[0];
      if (!off || !ext) continue;
      const x = eX(off.getAttribute("x")), y = eY(off.getAttribute("y"));
      const sw = eX(ext.getAttribute("cx")), sh2 = eY(ext.getAttribute("cy"));
      const blip = sh.getElementsByTagName("a:blip")[0];
      if (blip) { const rid = blip.getAttribute("r:embed"); if (rid && slide.images[rid]) imageItems.push({ src: slide.images[rid], x, y, w: sw, h: sh2 }); continue; }
      const spPr = sh.getElementsByTagName("p:spPr")[0] || sh.getElementsByTagName("a:spPr")[0];
      if (spPr) { const sf = spPr.getElementsByTagName("a:solidFill")[0]; if (sf) { const sc = sf.getElementsByTagName("a:srgbClr")[0]; if (sc) { ctx.fillStyle = "#" + sc.getAttribute("val"); ctx.fillRect(x, y, sw, sh2); } } }
      const tb = sh.getElementsByTagName("p:txBody")[0];
      if (!tb) continue;
      let yo = 0;
      for (const para of tb.getElementsByTagName("a:p")) {
        let txt = "", fs = szPx(1800), fb = false, fc = "#000";
        for (const run of para.getElementsByTagName("a:r")) {
          const rp = run.getElementsByTagName("a:rPr")[0];
          if (rp) { const sz = rp.getAttribute("sz"); if (sz) fs = szPx(sz); if (rp.getAttribute("b") === "1") fb = true; const sfl = rp.getElementsByTagName("a:solidFill")[0]; if (sfl) { const scc = sfl.getElementsByTagName("a:srgbClr")[0]; if (scc) fc = "#" + scc.getAttribute("val"); } }
          const te = run.getElementsByTagName("a:t")[0]; if (te) txt += te.textContent;
        }
        if (txt.trim()) textItems.push({ text: txt, x, y: y + yo, w: sw, fontSize: fs, fontBold: fb, fontColor: fc });
        yo += fs * 1.4;
      }
    }
    Promise.all(imageItems.map(it => new Promise(res => {
      const im = new Image(); im.crossOrigin = "anonymous";
      im.onload = () => { ctx.drawImage(im, it.x, it.y, it.w, it.h); res(); }; im.onerror = () => res(); im.src = it.src;
    }))).then(() => {
      for (const it of textItems) {
        const fz = Math.max(8, it.fontSize);
        ctx.font = `${it.fontBold ? "bold " : ""}${fz}px 'DM Sans',sans-serif`;
        ctx.fillStyle = it.fontColor; ctx.textBaseline = "top";
        const words = it.text.split(" "); let line = "", ly = it.y;
        for (const word of words) { const test = line + (line ? " " : "") + word; if (ctx.measureText(test).width > it.w && line) { ctx.fillText(line, it.x, ly); line = word; ly += fz * 1.3; } else line = test; }
        ctx.fillText(line, it.x, ly);
      }
      resolve(canvas.toDataURL("image/png"));
    });
  });
}

async function toJpg(pngUrl) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => { const c = document.createElement("canvas"); c.width = img.width; c.height = img.height; const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0); resolve(c.toDataURL("image/jpeg", 0.92)); };
    img.src = pngUrl;
  });
}

// ── Atoms ─────────────────────────────────────────────────────────────────────

function ResChip({ active, onClick, label, hint }) {
  return (
    <button onClick={onClick} style={{
      padding: "5px 11px", borderRadius: 3,
      border: `1px solid ${active ? "#00e5cc" : "#252525"}`,
      background: active ? "#00e5cc18" : "transparent",
      color: active ? "#00e5cc" : "#555",
      fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, fontWeight: 600,
      cursor: "pointer", letterSpacing: 1, transition: "all 0.15s",
    }}>
      {label} <span style={{ opacity: 0.45, fontSize: 8, marginLeft: 3 }}>{hint}</span>
    </button>
  );
}

function FmtPill({ value, onChange, disabled }) {
  return (
    <div style={{ display: "flex", borderRadius: 3, overflow: "hidden", border: `1px solid ${disabled ? "#1a1a1a" : "#252525"}`, flexShrink: 0 }}>
      {["png", "jpg"].map(f => (
        <button key={f} onClick={e => { e.stopPropagation(); if (!disabled) onChange(f); }} style={{
          padding: "3px 9px",
          background: value === f ? (f === "jpg" ? "#b84200" : "#004d66") : "transparent",
          color: value === f ? "#fff" : disabled ? "#2a2a2a" : "#3a3a3a",
          fontFamily: "'IBM Plex Mono',monospace", fontSize: 9, fontWeight: 700,
          letterSpacing: 1, textTransform: "uppercase", border: "none",
          cursor: disabled ? "default" : "pointer", transition: "all 0.12s",
        }}>{f}</button>
      ))}
    </div>
  );
}

// ── Step 1: Config card ───────────────────────────────────────────────────────

function ConfigCard({ slide, selected, format, onToggleSelect, onChangeFormat }) {
  const isJpg  = format === "jpg";
  const accent = selected ? (isJpg ? "#e65100" : "#00e5cc") : "#1e1e1e";
  return (
    <div style={{
      background: "#111", borderRadius: 8, overflow: "hidden",
      border: `1px solid ${accent}`, opacity: selected ? 1 : 0.42,
      transition: "border-color 0.18s, opacity 0.18s, transform 0.18s",
    }}
      onMouseEnter={e => e.currentTarget.style.transform = "scale(1.013)"}
      onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}
    >
      <div style={{ position: "relative", cursor: "pointer" }} onClick={onToggleSelect}>
        <img src={slide.preview} alt={`Slide ${slide.index + 1}`} style={{ width: "100%", display: "block" }} />
        {/* Checkbox */}
        <div style={{
          position: "absolute", top: 8, left: 8, width: 20, height: 20, borderRadius: 4,
          border: `2px solid ${selected ? (isJpg ? "#e65100" : "#00e5cc") : "#444"}`,
          background: selected ? (isJpg ? "#e65100" : "#00e5cc") : "rgba(0,0,0,0.65)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 11, color: "#000", fontWeight: 800, transition: "all 0.15s",
        }}>{selected ? "✓" : ""}</div>
        {/* Format badge */}
        <div style={{
          position: "absolute", top: 8, right: 8,
          background: isJpg ? "rgba(184,66,0,0.9)" : "rgba(0,77,102,0.9)",
          color: "#fff", fontFamily: "'IBM Plex Mono',monospace",
          fontSize: 9, fontWeight: 700, letterSpacing: 1, padding: "2px 7px", borderRadius: 3,
        }}>.{format.toUpperCase()}</div>
        {/* Slide number */}
        <div style={{
          position: "absolute", bottom: 8, right: 8,
          fontFamily: "'IBM Plex Mono',monospace", fontSize: 9,
          color: "rgba(255,255,255,0.4)", background: "rgba(0,0,0,0.55)", padding: "2px 7px", borderRadius: 3,
        }}>{String(slide.index + 1).padStart(2, "0")}</div>
      </div>
      <div style={{ padding: "7px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", borderTop: "1px solid #181818", gap: 6 }}>
        <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 10, color: "#3a3a3a", letterSpacing: 1 }}>
          SLIDE {String(slide.index + 1).padStart(2, "0")}
        </span>
        <FmtPill value={format} onChange={onChangeFormat} disabled={!selected} />
      </div>
    </div>
  );
}

// ── Step 2: Result card ───────────────────────────────────────────────────────

function ResultCard({ item, renderedRes }) {
  const isJpg = item.format === "jpg";
  const doDownload = useCallback(async () => {
    const a = document.createElement("a");
    const suffix = `_${renderedRes.w}x${renderedRes.h}`;
    if (isJpg) { a.href = await toJpg(item.dataUrl); a.download = `slide_${item.index + 1}${suffix}.jpg`; }
    else        { a.href = item.dataUrl; a.download = `slide_${item.index + 1}${suffix}.png`; }
    a.click();
  }, [item, renderedRes, isJpg]);

  return (
    <div style={{ background: "#111", borderRadius: 8, overflow: "hidden", border: "1px solid #1e1e1e", transition: "border-color 0.18s, transform 0.18s" }}
      onMouseEnter={e => { e.currentTarget.style.transform = "scale(1.013)"; e.currentTarget.style.borderColor = isJpg ? "#e6510044" : "#00e5cc33"; }}
      onMouseLeave={e => { e.currentTarget.style.transform = "scale(1)"; e.currentTarget.style.borderColor = "#1e1e1e"; }}
    >
      <div style={{ position: "relative" }}>
        <img src={item.dataUrl} alt={`Slide ${item.index + 1}`} style={{ width: "100%", display: "block" }} />
        <div style={{
          position: "absolute", top: 8, right: 8,
          background: isJpg ? "rgba(184,66,0,0.9)" : "rgba(0,77,102,0.9)",
          color: "#fff", fontFamily: "'IBM Plex Mono',monospace",
          fontSize: 9, fontWeight: 700, letterSpacing: 1, padding: "2px 7px", borderRadius: 3,
        }}>.{item.format.toUpperCase()}</div>
        <div style={{
          position: "absolute", bottom: 8, right: 8,
          fontFamily: "'IBM Plex Mono',monospace", fontSize: 9,
          color: "rgba(255,255,255,0.4)", background: "rgba(0,0,0,0.55)", padding: "2px 7px", borderRadius: 3,
        }}>{String(item.index + 1).padStart(2, "0")}</div>
      </div>
      <div style={{ padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", borderTop: "1px solid #181818" }}>
        <span style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 9, color: "#2e2e2e", letterSpacing: 1 }}>
          {renderedRes.w}×{renderedRes.h}
        </span>
        <button onClick={doDownload} style={{
          padding: "4px 12px", borderRadius: 3,
          border: `1px solid ${isJpg ? "#6b2500" : "#1e3a44"}`,
          background: "transparent", color: isJpg ? "#b84200" : "#007799",
          fontFamily: "'IBM Plex Mono',monospace", fontSize: 9,
          cursor: "pointer", letterSpacing: 1, transition: "all 0.15s",
        }}
          onMouseEnter={e => { e.currentTarget.style.color = isJpg ? "#e65100" : "#00e5cc"; e.currentTarget.style.borderColor = isJpg ? "#e65100" : "#00e5cc"; }}
          onMouseLeave={e => { e.currentTarget.style.color = isJpg ? "#b84200" : "#007799"; e.currentTarget.style.borderColor = isJpg ? "#6b2500" : "#1e3a44"; }}
        >↓ .{item.format.toUpperCase()}</button>
      </div>
    </div>
  );
}

// ── Main ──────────────────────────────────────────────────────────────────────

export default function App() {
  // phase: idle | previewing | config | converting | done | error
  const [phase,        setPhase]        = useState("idle");
  const [fileName,     setFileName]     = useState("");
  const [parsedSlides, setParsedSlides] = useState([]);
  const [previews,     setPreviews]     = useState([]);
  const [slideFormats, setSlideFormats] = useState({});
  const [selected,     setSelected]     = useState(new Set());
  const [resIdx,       setResIdx]       = useState(1);
  const [progress,     setProgress]     = useState(0);
  const [progressMsg,  setProgressMsg]  = useState("");
  const [results,      setResults]      = useState([]);
  const [dragOver,     setDragOver]     = useState(false);
  const inputRef = useRef(null);

  const renderedRes = RESOLUTIONS[resIdx];
  const isLoading   = phase === "previewing" || phase === "converting";
  const selArr      = [...selected].sort((a, b) => a - b);
  const jpgSel      = selArr.filter(i => slideFormats[i] === "jpg").length;
  const pngSel      = selArr.length - jpgSel;

  const processFile = useCallback(async file => {
    if (!file?.name.match(/\.pptx$/i)) { setPhase("error"); return; }
    setFileName(file.name); setPhase("previewing");
    setProgress(0); setProgressMsg("Parsing file…");
    setParsedSlides([]); setPreviews([]); setResults([]);
    setSelected(new Set()); setSlideFormats({});
    try {
      const parsed = await parsePptx(file);
      setParsedSlides(parsed);
      const pv = [];
      for (let i = 0; i < parsed.length; i++) {
        setProgress(((i + 1) / parsed.length) * 100);
        setProgressMsg(`Loading preview ${i + 1} / ${parsed.length}…`);
        pv.push({ index: i, preview: await renderToCanvas(parsed[i], PREVIEW_W, PREVIEW_H) });
      }
      setPreviews(pv);
      const fmts = {}; pv.forEach((_, i) => { fmts[i] = "png"; });
      setSlideFormats(fmts);
      setSelected(new Set(pv.map((_, i) => i)));
      setPhase("config");
    } catch (e) { console.error(e); setPhase("error"); }
  }, []);

  const handleConvert = useCallback(async () => {
    if (!selArr.length) return;
    setPhase("converting"); setProgress(0); setResults([]);
    const res = RESOLUTIONS[resIdx];
    const out = [];
    for (let i = 0; i < selArr.length; i++) {
      const idx = selArr[i];
      setProgress(((i + 1) / selArr.length) * 100);
      setProgressMsg(`Converting slide ${idx + 1} as ${(slideFormats[idx] || "png").toUpperCase()} at ${res.hint} (${i + 1}/${selArr.length})…`);
      const dataUrl = await renderToCanvas(parsedSlides[idx], res.w, res.h);
      out.push({ index: idx, dataUrl, format: slideFormats[idx] || "png" });
    }
    setResults(out); setPhase("done");
  }, [selArr, parsedSlides, slideFormats, resIdx]);

  const downloadAll = useCallback(async () => {
    const res = RESOLUTIONS[resIdx];
    for (let i = 0; i < results.length; i++) {
      const item = results[i]; const a = document.createElement("a");
      const suffix = `_${res.w}x${res.h}`;
      if (item.format === "jpg") { a.href = await toJpg(item.dataUrl); a.download = `slide_${item.index + 1}${suffix}.jpg`; }
      else { a.href = item.dataUrl; a.download = `slide_${item.index + 1}${suffix}.png`; }
      a.click(); await new Promise(r => setTimeout(r, 280));
    }
  }, [results, resIdx]);

  const toggleSelect = idx => setSelected(prev => { const n = new Set(prev); n.has(idx) ? n.delete(idx) : n.add(idx); return n; });
  const toggleAll    = () => setSelected(selected.size === previews.length ? new Set() : new Set(previews.map((_, i) => i)));
  const setAllFmt    = fmt => { const n = {}; previews.forEach((_, i) => { n[i] = fmt; }); setSlideFormats(n); };

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0a", color: "#d0d0d0", fontFamily: "'DM Sans',sans-serif" }}>
      <style>{`
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.2}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(11px)}to{opacity:1;transform:translateY(0)}}
        *{box-sizing:border-box}
        ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:#0a0a0a}
        ::-webkit-scrollbar-thumb{background:#1e1e1e;border-radius:3px}
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: "1px solid #141414", padding: "15px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 30, height: 30, background: "linear-gradient(135deg,#00e5cc,#0077ff)", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>⬡</div>
          <div>
            <div style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600, letterSpacing: 1, color: "#ddd" }}>PPTX → IMAGE</div>
            <div style={{ fontSize: 10, color: "#333", marginTop: 1 }}>configure per slide, then convert</div>
          </div>
        </div>
        <div style={{
          display: "flex", alignItems: "center", gap: 6, padding: "3px 11px",
          borderRadius: 3, border: "1px solid #1a1a1a", background: "#0e0e0e",
          fontFamily: "'IBM Plex Mono',monospace", fontSize: 9, letterSpacing: 2,
          color: { idle:"#2a2a2a", previewing:"#00e5cc", config:"#4ade80", converting:"#f59e0b", done:"#4ade80", error:"#f87171" }[phase],
        }}>
          {isLoading && <span style={{ width:5, height:5, borderRadius:"50%", background:"currentColor", animation:"pulse 1s infinite" }}/>}
          {{ idle:"IDLE", previewing:"LOADING", config:"CONFIGURE", converting:"CONVERTING", done:"DONE", error:"ERROR" }[phase]}
        </div>
      </div>

      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "22px 20px" }}>

        {/* Drop zone */}
        {(phase === "idle" || phase === "config" || phase === "done") && (
          <div
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); }}
            onClick={() => inputRef.current?.click()}
            style={{
              border: `1.5px dashed ${dragOver ? "#00e5cc" : "#1e1e1e"}`,
              borderRadius: 10, padding: phase === "idle" ? "50px 22px" : "14px 20px",
              textAlign: "center", cursor: "pointer",
              background: dragOver ? "rgba(0,229,204,0.03)" : "#0d0d0d",
              transition: "all 0.2s", marginBottom: 18,
            }}
          >
            <input ref={inputRef} type="file" accept=".pptx" style={{ display:"none" }} onChange={e => processFile(e.target.files[0])} />
            {phase === "idle" ? (
              <>
                <div style={{ fontSize:32, opacity:0.14, marginBottom:10 }}>↑</div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:12, color:"#444", marginBottom:4, letterSpacing:1 }}>DROP .PPTX FILE HERE</div>
                <div style={{ fontSize:11, color:"#282828" }}>or click to browse</div>
              </>
            ) : (
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <span style={{ fontSize:17 }}>📄</span>
                  <div style={{ textAlign:"left" }}>
                    <div style={{ fontSize:13, fontWeight:500, color:"#ccc" }}>{fileName}</div>
                    <div style={{ fontSize:10, color:"#3a3a3a", marginTop:1 }}>{previews.length} slides</div>
                  </div>
                </div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#282828", letterSpacing:1 }}>CLICK TO REPLACE</div>
              </div>
            )}
          </div>
        )}

        {/* Progress */}
        {isLoading && (
          <>
            <div style={{ background:"#151515", borderRadius:3, height:2, marginBottom:6, overflow:"hidden" }}>
              <div style={{ height:"100%", width:`${progress}%`, background:"linear-gradient(90deg,#00e5cc,#0077ff)", borderRadius:3, transition:"width 0.3s" }}/>
            </div>
            <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#333", letterSpacing:1, marginBottom:20 }}>{progressMsg}</div>
          </>
        )}

        {/* ─────────── STEP 1: CONFIG ─────────── */}
        {phase === "config" && (
          <>
            {/* Step label */}
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
              <div style={{ width:22, height:22, borderRadius:"50%", background:"#00e5cc", color:"#000", fontFamily:"'IBM Plex Mono',monospace", fontSize:11, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>1</div>
              <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:11, color:"#555", letterSpacing:1 }}>CHOOSE FORMAT PER SLIDE, THEN CLICK CONVERT</div>
            </div>

            {/* Toolbar */}
            <div style={{ background:"#0d0d0d", border:"1px solid #171717", borderRadius:8, padding:"13px 16px", marginBottom:14, display:"flex", flexWrap:"wrap", gap:14, alignItems:"flex-start" }}>
              {/* Resolution */}
              <div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:8, color:"#333", letterSpacing:2, marginBottom:6, textTransform:"uppercase" }}>Output Resolution</div>
                <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                  {RESOLUTIONS.map((r, i) => <ResChip key={r.label} active={resIdx === i} label={r.label} hint={r.hint} onClick={() => setResIdx(i)} />)}
                </div>
              </div>
              {/* Set-all */}
              <div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:8, color:"#333", letterSpacing:2, marginBottom:6, textTransform:"uppercase" }}>Set All To</div>
                <div style={{ display:"flex", gap:5 }}>
                  {[["png","#004d66","#0099cc","rgba(0,77,102,"],["jpg","#6b2500","#e65100","rgba(184,66,0,"]].map(([f,b,c,bg]) => (
                    <button key={f} onClick={() => setAllFmt(f)} style={{
                      padding:"5px 12px", borderRadius:3, border:`1px solid ${b}`,
                      background:`${bg}0.1)`, color:c,
                      fontFamily:"'IBM Plex Mono',monospace", fontSize:10, fontWeight:700,
                      cursor:"pointer", letterSpacing:1, transition:"all 0.15s",
                    }}
                      onMouseEnter={e=>e.currentTarget.style.background=`${bg}0.22)`}
                      onMouseLeave={e=>e.currentTarget.style.background=`${bg}0.1)`}
                    >ALL → {f.toUpperCase()}</button>
                  ))}
                </div>
              </div>
              <div style={{ flex:1 }}/>
              {/* Select + Convert */}
              <div style={{ display:"flex", alignItems:"flex-end", gap:7 }}>
                <button onClick={toggleAll} style={{ padding:"5px 12px", borderRadius:3, border:"1px solid #222", background:"transparent", color:"#555", fontFamily:"'IBM Plex Mono',monospace", fontSize:9, cursor:"pointer", letterSpacing:1, transition:"all 0.15s" }}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor="#3a3a3a";e.currentTarget.style.color="#999";}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor="#222";e.currentTarget.style.color="#555";}}
                >{selected.size === previews.length ? "DESELECT ALL" : "SELECT ALL"}</button>
                <button onClick={handleConvert} disabled={selected.size === 0} style={{
                  padding:"6px 20px", borderRadius:3, border:"none",
                  background: selected.size > 0 ? "linear-gradient(90deg,#00e5cc,#0077ff)" : "#181818",
                  color: selected.size > 0 ? "#000" : "#2a2a2a",
                  fontFamily:"'IBM Plex Mono',monospace", fontSize:11, fontWeight:700,
                  cursor: selected.size > 0 ? "pointer" : "not-allowed", letterSpacing:1,
                }}>
                  {selected.size > 0
                    ? `⚙ CONVERT ${selected.size} · ${pngSel > 0 ? `${pngSel} PNG` : ""}${pngSel > 0 && jpgSel > 0 ? " + " : ""}${jpgSel > 0 ? `${jpgSel} JPG` : ""}`
                    : "⚙ CONVERT"}
                </button>
              </div>
            </div>

            {/* Hint */}
            <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#242424", marginBottom:14, letterSpacing:1 }}>
              {selected.size}/{previews.length} SELECTED &nbsp;·&nbsp; CLICK THUMBNAIL TO INCLUDE / EXCLUDE &nbsp;·&nbsp; USE PNG · JPG TOGGLE PER CARD
            </div>

            {/* Config grid */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(220px, 1fr))", gap:10 }}>
              {previews.map(slide => (
                <div key={slide.index} style={{ animation:"fadeUp 0.26s ease-out both", animationDelay:`${Math.min(slide.index * 0.028, 0.42)}s` }}>
                  <ConfigCard
                    slide={slide}
                    selected={selected.has(slide.index)}
                    format={slideFormats[slide.index] || "png"}
                    onToggleSelect={() => toggleSelect(slide.index)}
                    onChangeFormat={fmt => setSlideFormats(p => ({ ...p, [slide.index]: fmt }))}
                  />
                </div>
              ))}
            </div>
          </>
        )}

        {/* ─────────── STEP 2: DONE ─────────── */}
        {phase === "done" && (
          <>
            {/* Step label */}
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
              <div style={{ width:22, height:22, borderRadius:"50%", background:"#4ade80", color:"#000", fontFamily:"'IBM Plex Mono',monospace", fontSize:11, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>2</div>
              <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:11, color:"#555", letterSpacing:1 }}>CONVERSION COMPLETE — DOWNLOAD INDIVIDUALLY OR ALL AT ONCE</div>
            </div>

            {/* Results toolbar */}
            <div style={{ background:"#0d0d0d", border:"1px solid #171717", borderRadius:8, padding:"13px 16px", marginBottom:14, display:"flex", flexWrap:"wrap", gap:12, alignItems:"center" }}>
              <div style={{ flex:1 }}>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:11, color:"#4ade80", letterSpacing:1 }}>
                  ✓ {results.length} SLIDE{results.length !== 1 ? "S" : ""} · {renderedRes.hint}
                </div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#333", marginTop:3, letterSpacing:1 }}>
                  {[results.filter(r=>r.format==="png").length > 0 && `${results.filter(r=>r.format==="png").length} PNG`, results.filter(r=>r.format==="jpg").length > 0 && `${results.filter(r=>r.format==="jpg").length} JPG`].filter(Boolean).join("  +  ")}
                </div>
              </div>
              <button onClick={() => setPhase("config")} style={{ padding:"6px 14px", borderRadius:3, border:"1px solid #222", background:"transparent", color:"#666", fontFamily:"'IBM Plex Mono',monospace", fontSize:9, cursor:"pointer", letterSpacing:1, transition:"all 0.15s" }}
                onMouseEnter={e=>{e.currentTarget.style.borderColor="#444";e.currentTarget.style.color="#aaa";}}
                onMouseLeave={e=>{e.currentTarget.style.borderColor="#222";e.currentTarget.style.color="#666";}}
              >← BACK TO CONFIG</button>
              <button onClick={downloadAll} style={{ padding:"6px 22px", borderRadius:3, border:"none", background:"linear-gradient(90deg,#00e5cc,#0077ff)", color:"#000", fontFamily:"'IBM Plex Mono',monospace", fontSize:11, fontWeight:700, cursor:"pointer", letterSpacing:1 }}>
                ↓ DOWNLOAD ALL
              </button>
            </div>

            {/* Results grid */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(240px, 1fr))", gap:11 }}>
              {results.map(item => (
                <div key={item.index} style={{ animation:"fadeUp 0.26s ease-out both", animationDelay:`${Math.min(item.index * 0.028, 0.42)}s` }}>
                  <ResultCard item={item} renderedRes={renderedRes} />
                </div>
              ))}
            </div>
          </>
        )}

        {/* Error */}
        {phase === "error" && (
          <div style={{ textAlign:"center", padding:40, color:"#f87171", fontFamily:"'IBM Plex Mono',monospace", fontSize:11, letterSpacing:1 }}>
            FAILED TO PROCESS FILE — please ensure it is a valid .pptx
          </div>
        )}
      </div>
    </div>
  );
}

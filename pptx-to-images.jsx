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

// ─── Utilities ────────────────────────────────────────────────────────────────

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (window.JSZip) return resolve();
    if (document.querySelector(`script[src="${src}"]`)) {
      const check = setInterval(() => { if (window.JSZip) { clearInterval(check); resolve(); } }, 50);
      return;
    }
    const s = document.createElement("script");
    s.src = src; s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
}

async function parsePptx(file) {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
  const ab = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(ab);

  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1]) - parseInt(b.match(/slide(\d+)/)[1]));

  const images = {};
  for (const key of Object.keys(zip.files)) {
    if (key.startsWith("ppt/media/"))
      images[key] = URL.createObjectURL(await zip.files[key].async("blob"));
  }

  const slides = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const slideXml = await zip.files[slideFiles[i]].async("text");
    const slideNum = parseInt(slideFiles[i].match(/slide(\d+)/)[1]);
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    let slideImages = {};
    if (zip.files[relsPath]) {
      const relsDoc = new DOMParser().parseFromString(await zip.files[relsPath].async("text"), "text/xml");
      for (const rel of relsDoc.getElementsByTagName("Relationship")) {
        const target = rel.getAttribute("Target");
        if (target?.includes("media/")) {
          const fp = "ppt/" + target.replace("../", "");
          if (images[fp]) slideImages[rel.getAttribute("Id")] = images[fp];
        }
      }
    }
    const sd = new DOMParser().parseFromString(slideXml, "text/xml");
    let bgColor = null;
    const bg = sd.getElementsByTagName("p:bg")[0];
    if (bg) { const sc = bg.getElementsByTagName("a:srgbClr")[0]; if (sc) bgColor = "#" + sc.getAttribute("val"); }
    slides.push({ index: i, xml: slideXml, images: slideImages, bgColor });
  }
  return slides;
}

function renderSlideToCanvas(slide, width, height) {
  return new Promise((resolve) => {
    const canvas = document.createElement("canvas");
    canvas.width = width; canvas.height = height;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = slide.bgColor || "#ffffff";
    ctx.fillRect(0, 0, width, height);

    const doc = new DOMParser().parseFromString(slide.xml, "text/xml");
    const SX = width / 12192000, SY = height / 6858000;
    const eX = (v) => (parseInt(v) || 0) * SX;
    const eY = (v) => (parseInt(v) || 0) * SY;
    const szPx = (sz) => ((parseInt(sz) || 1800) / 100) * (width / 960);

    const textItems = [], imageItems = [];
    const shapes = [...Array.from(doc.getElementsByTagName("p:sp")), ...Array.from(doc.getElementsByTagName("p:pic"))];

    for (const sh of shapes) {
      const off = sh.getElementsByTagName("a:off")[0];
      const ext = sh.getElementsByTagName("a:ext")[0];
      if (!off || !ext) continue;
      const x = eX(off.getAttribute("x")), y = eY(off.getAttribute("y"));
      const w = eX(ext.getAttribute("cx")), h = eY(ext.getAttribute("cy"));

      const blip = sh.getElementsByTagName("a:blip")[0];
      if (blip) {
        const rid = blip.getAttribute("r:embed");
        if (rid && slide.images[rid]) imageItems.push({ src: slide.images[rid], x, y, w, h });
        continue;
      }
      const spPr = sh.getElementsByTagName("p:spPr")[0] || sh.getElementsByTagName("a:spPr")[0];
      if (spPr) {
        const sf = spPr.getElementsByTagName("a:solidFill")[0];
        if (sf) { const sc = sf.getElementsByTagName("a:srgbClr")[0]; if (sc) { ctx.fillStyle = "#" + sc.getAttribute("val"); ctx.fillRect(x, y, w, h); } }
      }
      const tb = sh.getElementsByTagName("p:txBody")[0];
      if (!tb) continue;
      let yOff = 0;
      for (const para of tb.getElementsByTagName("a:p")) {
        let txt = "", fs = szPx(1800), fb = false, fc = "#000000";
        for (const run of para.getElementsByTagName("a:r")) {
          const rp = run.getElementsByTagName("a:rPr")[0];
          if (rp) {
            const sz = rp.getAttribute("sz"); if (sz) fs = szPx(sz);
            if (rp.getAttribute("b") === "1") fb = true;
            const sfl = rp.getElementsByTagName("a:solidFill")[0];
            if (sfl) { const scc = sfl.getElementsByTagName("a:srgbClr")[0]; if (scc) fc = "#" + scc.getAttribute("val"); }
          }
          const te = run.getElementsByTagName("a:t")[0];
          if (te) txt += te.textContent;
        }
        if (txt.trim()) textItems.push({ text: txt, x, y: y + yOff, w, fontSize: fs, fontBold: fb, fontColor: fc });
        yOff += fs * 1.4;
      }
    }

    Promise.all(imageItems.map((it) => new Promise((res) => {
      const im = new Image(); im.crossOrigin = "anonymous";
      im.onload = () => { ctx.drawImage(im, it.x, it.y, it.w, it.h); res(); };
      im.onerror = () => res();
      im.src = it.src;
    }))).then(() => {
      for (const it of textItems) {
        const fz = Math.max(10, it.fontSize);
        ctx.font = `${it.fontBold ? "bold " : ""}${fz}px 'DM Sans', sans-serif`;
        ctx.fillStyle = it.fontColor; ctx.textBaseline = "top";
        const words = it.text.split(" ");
        let line = "", lineY = it.y;
        for (const word of words) {
          const test = line + (line ? " " : "") + word;
          if (ctx.measureText(test).width > it.w && line) { ctx.fillText(line, it.x, lineY); line = word; lineY += fz * 1.3; }
          else line = test;
        }
        ctx.fillText(line, it.x, lineY);
      }
      resolve(canvas.toDataURL("image/png"));
    });
  });
}

async function toJpgDataUrl(pngDataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement("canvas");
      c.width = img.width; c.height = img.height;
      const ctx = c.getContext("2d");
      ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
      ctx.drawImage(img, 0, 0);
      resolve(c.toDataURL("image/jpeg", 0.92));
    };
    img.src = pngDataUrl;
  });
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function ResChip({ active, onClick, children, warn }) {
  const accent = warn ? "#f59e0b" : "#00e5cc";
  return (
    <button onClick={onClick} style={{
      padding: "5px 12px", borderRadius: 3,
      border: `1px solid ${active ? accent : "#2a2a2a"}`,
      background: active ? `${accent}18` : "transparent",
      color: active ? accent : "#555",
      fontFamily: "'IBM Plex Mono', monospace", fontSize: 10,
      fontWeight: 600, cursor: "pointer", letterSpacing: 1,
      textTransform: "uppercase", transition: "all 0.15s", lineHeight: 1.4,
    }}>
      {children}
    </button>
  );
}

// PNG / JPG pill toggle embedded in each card
function FmtToggle({ value, onChange }) {
  return (
    <div style={{ display: "flex", borderRadius: 3, overflow: "hidden", border: "1px solid #252525" }}>
      {["png", "jpg"].map((f) => (
        <button
          key={f}
          onClick={(e) => { e.stopPropagation(); onChange(f); }}
          style={{
            padding: "3px 10px",
            background: value === f ? (f === "jpg" ? "#b84200" : "#004d66") : "transparent",
            color: value === f ? "#fff" : "#3a3a3a",
            fontFamily: "'IBM Plex Mono', monospace", fontSize: 9,
            fontWeight: 700, letterSpacing: 1, textTransform: "uppercase",
            border: "none", cursor: "pointer", transition: "all 0.15s",
          }}
        >
          {f}
        </button>
      ))}
    </div>
  );
}

function SlideCard({ slide, selected, format, onToggleSelect, onChangeFormat, onDownload, renderedRes }) {
  const isJpg = format === "jpg";
  const selColor = isJpg ? "#e65100" : "#00e5cc";
  return (
    <div
      style={{
        background: "#111", borderRadius: 8, overflow: "hidden",
        border: `1px solid ${selected ? selColor + "44" : "#1e1e1e"}`,
        transition: "border-color 0.2s, transform 0.18s, box-shadow 0.18s",
      }}
      onMouseEnter={e => { e.currentTarget.style.transform = "scale(1.013)"; e.currentTarget.style.boxShadow = `0 5px 28px ${selColor}14`; }}
      onMouseLeave={e => { e.currentTarget.style.transform = "scale(1)"; e.currentTarget.style.boxShadow = "none"; }}
    >
      {/* Thumbnail — click to toggle selection */}
      <div style={{ position: "relative", cursor: "pointer" }} onClick={onToggleSelect}>
        <img
          src={slide.dataUrl}
          alt={`Slide ${slide.index + 1}`}
          style={{ width: "100%", display: "block", opacity: selected ? 1 : 0.28, transition: "opacity 0.2s" }}
        />

        {/* Checkbox (top-left) */}
        <div style={{
          position: "absolute", top: 8, left: 8,
          width: 20, height: 20, borderRadius: 4,
          border: `2px solid ${selected ? selColor : "#3a3a3a"}`,
          background: selected ? selColor : "rgba(0,0,0,0.6)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 11, color: "#000", fontWeight: 800, transition: "all 0.15s",
        }}>
          {selected ? "✓" : ""}
        </div>

        {/* Format badge (top-right) — reflects current per-slide choice */}
        <div style={{
          position: "absolute", top: 8, right: 8,
          background: isJpg ? "rgba(184,66,0,0.88)" : "rgba(0,77,102,0.88)",
          color: "#fff", fontFamily: "'IBM Plex Mono', monospace",
          fontSize: 9, fontWeight: 700, letterSpacing: 1,
          padding: "2px 7px", borderRadius: 3,
        }}>
          .{format.toUpperCase()}
        </div>

        {/* Slide number (bottom-right) */}
        <div style={{
          position: "absolute", bottom: 8, right: 8,
          fontFamily: "'IBM Plex Mono', monospace", fontSize: 9,
          color: "rgba(255,255,255,0.35)", background: "rgba(0,0,0,0.5)",
          padding: "2px 7px", borderRadius: 3,
        }}>
          {String(slide.index + 1).padStart(2, "0")}
        </div>
      </div>

      {/* Footer row */}
      <div style={{
        padding: "8px 10px", display: "flex", alignItems: "center",
        justifyContent: "space-between", borderTop: "1px solid #181818", gap: 6,
      }}>
        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 9, color: "#2e2e2e", letterSpacing: 1, flexShrink: 0 }}>
          {renderedRes.w}×{renderedRes.h}
        </span>

        {/* Per-slide format toggle — the key new feature */}
        <FmtToggle value={format} onChange={onChangeFormat} />

        {/* Individual download */}
        <button
          onClick={(e) => { e.stopPropagation(); onDownload(); }}
          style={{
            padding: "4px 10px", borderRadius: 3,
            border: `1px solid ${isJpg ? "#6b2500" : "#1e3a44"}`,
            background: "transparent",
            color: isJpg ? "#b84200" : "#007799",
            fontFamily: "'IBM Plex Mono', monospace", fontSize: 9,
            cursor: "pointer", letterSpacing: 1, transition: "all 0.15s", flexShrink: 0,
          }}
          onMouseEnter={e => { e.currentTarget.style.borderColor = selColor; e.currentTarget.style.color = selColor; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = isJpg ? "#6b2500" : "#1e3a44"; e.currentTarget.style.color = isJpg ? "#b84200" : "#007799"; }}
        >
          ↓ .{format.toUpperCase()}
        </button>
      </div>
    </div>
  );
}

// ─── Main ─────────────────────────────────────────────────────────────────────

export default function PptxToImages() {
  const [parsedSlides, setParsedSlides] = useState([]);
  const [slides, setSlides] = useState([]);
  const [status, setStatus] = useState("idle");
  const [fileName, setFileName] = useState("");
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [dragOver, setDragOver] = useState(false);

  // Per-slide format map: { [slideIndex]: "png" | "jpg" }
  const [slideFormats, setSlideFormats] = useState({});
  const [selectedSlides, setSelectedSlides] = useState(new Set());

  const [resIdx, setResIdx] = useState(1);
  const [renderedResIdx, setRenderedResIdx] = useState(1);
  const inputRef = useRef(null);

  const needsRerender = resIdx !== renderedResIdx && slides.length > 0;
  const currentRes = RESOLUTIONS[resIdx];
  const renderedRes = RESOLUTIONS[renderedResIdx];
  const isLoading = status === "loading" || status === "rerendering";

  // Export summary counts
  const selectedCount = selectedSlides.size;
  const jpgCount = slides.filter((s) => selectedSlides.has(s.index) && slideFormats[s.index] === "jpg").length;
  const pngCount = selectedCount - jpgCount;

  // ── Rendering helpers ──

  const renderAll = useCallback(async (parsed, targetResIdx) => {
    const res = RESOLUTIONS[targetResIdx];
    const rendered = [];
    for (let i = 0; i < parsed.length; i++) {
      setProgress(((i + 1) / parsed.length) * 100);
      setProgressLabel(`Rendering slide ${i + 1} / ${parsed.length} at ${res.hint}…`);
      rendered.push({ index: i, dataUrl: await renderSlideToCanvas(parsed[i], res.w, res.h) });
    }
    return rendered;
  }, []);

  const processFile = useCallback(async (file) => {
    if (!file?.name.match(/\.pptx$/i)) { setStatus("error"); return; }
    setFileName(file.name); setStatus("loading");
    setSlides([]); setParsedSlides([]); setProgress(0); setProgressLabel("Reading file…");
    setSelectedSlides(new Set()); setSlideFormats({});
    try {
      const parsed = await parsePptx(file);
      setParsedSlides(parsed);
      const rendered = await renderAll(parsed, resIdx);
      setSlides(rendered);
      setSelectedSlides(new Set(rendered.map((_, i) => i)));
      // Default every slide to PNG
      const fmts = {};
      rendered.forEach((_, i) => { fmts[i] = "png"; });
      setSlideFormats(fmts);
      setRenderedResIdx(resIdx);
      setStatus("done");
    } catch (e) { console.error(e); setStatus("error"); }
  }, [resIdx, renderAll]);

  const handleRerender = useCallback(async () => {
    if (!parsedSlides.length) return;
    setStatus("rerendering"); setSlides([]); setProgress(0);
    try {
      const rendered = await renderAll(parsedSlides, resIdx);
      setSlides(rendered);
      setRenderedResIdx(resIdx);
      setStatus("done");
    } catch (e) { console.error(e); setStatus("error"); }
  }, [parsedSlides, resIdx, renderAll]);

  // ── Per-slide actions ──

  const setSlideFormat = (idx, fmt) =>
    setSlideFormats((prev) => ({ ...prev, [idx]: fmt }));

  const setAllFormats = (fmt) => {
    const next = {};
    slides.forEach((s) => { next[s.index] = fmt; });
    setSlideFormats(next);
  };

  const toggleSelect = (idx) =>
    setSelectedSlides((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });

  const toggleAll = () => {
    if (selectedSlides.size === slides.length) setSelectedSlides(new Set());
    else setSelectedSlides(new Set(slides.map((_, i) => i)));
  };

  // ── Download ──

  const downloadSlide = useCallback(async (slide, fmt) => {
    const a = document.createElement("a");
    const suffix = `_${renderedRes.w}x${renderedRes.h}`;
    if (fmt === "jpg") {
      a.href = await toJpgDataUrl(slide.dataUrl);
      a.download = `slide_${slide.index + 1}${suffix}.jpg`;
    } else {
      a.href = slide.dataUrl;
      a.download = `slide_${slide.index + 1}${suffix}.png`;
    }
    a.click();
  }, [renderedRes]);

  const downloadSelected = useCallback(() => {
    slides
      .filter((s) => selectedSlides.has(s.index))
      .forEach((slide, idx) =>
        setTimeout(() => downloadSlide(slide, slideFormats[slide.index] || "png"), idx * 280)
      );
  }, [slides, selectedSlides, slideFormats, downloadSlide]);

  // ─────────────────────────────────────────────────────────────────────────────

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0a", color: "#d0d0d0", fontFamily: "'DM Sans', sans-serif" }}>
      <style>{`
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.2} }
        @keyframes fadeUp { from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)} }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar{width:5px} ::-webkit-scrollbar-track{background:#0a0a0a}
        ::-webkit-scrollbar-thumb{background:#1e1e1e;border-radius:3px}
      `}</style>

      {/* ── Header ── */}
      <div style={{ borderBottom: "1px solid #141414", padding: "16px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 30, height: 30, background: "linear-gradient(135deg,#00e5cc,#0077ff)", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>⬡</div>
          <div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, fontWeight: 600, letterSpacing: 1, color: "#ddd" }}>PPTX → IMAGE</div>
            <div style={{ fontSize: 10, color: "#333", marginTop: 1 }}>per-slide format export</div>
          </div>
        </div>
        <div style={{
          display: "flex", alignItems: "center", gap: 6, padding: "3px 11px",
          borderRadius: 3, border: "1px solid #1a1a1a", background: "#0e0e0e",
          fontFamily: "'IBM Plex Mono', monospace", fontSize: 9, letterSpacing: 2,
          color: { idle:"#2e2e2e", loading:"#00e5cc", rerendering:"#f59e0b", done:"#4ade80", error:"#f87171" }[status],
        }}>
          {isLoading && <span style={{ width:5, height:5, borderRadius:"50%", background:"currentColor", animation:"pulse 1s infinite" }}/>}
          {{ idle:"IDLE", loading:"PROCESSING", rerendering:"RE-RENDERING", done:"DONE", error:"ERROR" }[status]}
        </div>
      </div>

      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "24px 20px" }}>

        {/* ── Drop Zone ── */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); }}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `1.5px dashed ${dragOver ? "#00e5cc" : "#1e1e1e"}`,
            borderRadius: 10, padding: slides.length || isLoading ? "16px 22px" : "52px 22px",
            textAlign: "center", cursor: "pointer",
            background: dragOver ? "rgba(0,229,204,0.03)" : "#0d0d0d",
            transition: "all 0.2s", marginBottom: 18,
          }}
        >
          <input ref={inputRef} type="file" accept=".pptx" style={{ display:"none" }} onChange={(e) => processFile(e.target.files[0])} />
          {!slides.length && !isLoading ? (
            <>
              <div style={{ fontSize: 34, opacity: 0.15, marginBottom: 12 }}>↑</div>
              <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:12, color:"#444", marginBottom:5, letterSpacing:1 }}>DROP .PPTX FILE HERE</div>
              <div style={{ fontSize:11, color:"#2a2a2a" }}>or click to browse</div>
            </>
          ) : (
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                <span style={{ fontSize:18 }}>📄</span>
                <div style={{ textAlign:"left" }}>
                  <div style={{ fontSize:13, fontWeight:500, color:"#ccc" }}>{fileName}</div>
                  <div style={{ fontSize:10, color:"#3a3a3a", marginTop:2 }}>
                    {isLoading ? progressLabel : `${slides.length} slides · ${renderedRes.hint}`}
                  </div>
                </div>
              </div>
              {!isLoading && <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#2a2a2a", letterSpacing:1 }}>CLICK TO REPLACE</div>}
            </div>
          )}
        </div>

        {/* ── Progress Bar ── */}
        {isLoading && (
          <div style={{ background:"#151515", borderRadius:3, height:2, marginBottom:18, overflow:"hidden" }}>
            <div style={{ height:"100%", width:`${progress}%`, background:"linear-gradient(90deg,#00e5cc,#0077ff)", borderRadius:3, transition:"width 0.3s" }}/>
          </div>
        )}

        {/* ── Toolbar ── */}
        {(slides.length > 0 || isLoading) && (
          <div style={{
            background:"#0d0d0d", border:"1px solid #171717", borderRadius:8,
            padding:"14px 16px", marginBottom:16,
            display:"flex", flexWrap:"wrap", gap:16, alignItems:"flex-start",
          }}>

            {/* Resolution */}
            <div>
              <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:8, color:"#333", letterSpacing:2, marginBottom:7, textTransform:"uppercase" }}>
                Resolution{needsRerender && <span style={{ color:"#f59e0b", marginLeft:5 }}>· pending re-render</span>}
              </div>
              <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                {RESOLUTIONS.map((r, i) => (
                  <ResChip key={r.label} active={resIdx === i} warn={needsRerender && resIdx === i} onClick={() => setResIdx(i)}>
                    {r.label}&nbsp;<span style={{ opacity:0.5, fontSize:8 }}>{r.hint}</span>
                  </ResChip>
                ))}
              </div>
            </div>

            {/* Re-render */}
            {needsRerender && (
              <div style={{ display:"flex", alignItems:"flex-end" }}>
                <button onClick={handleRerender} style={{
                  padding:"5px 16px", borderRadius:3, border:"1px solid #f59e0b",
                  background:"rgba(245,158,11,0.08)", color:"#f59e0b",
                  fontFamily:"'IBM Plex Mono',monospace", fontSize:10, fontWeight:600,
                  cursor:"pointer", letterSpacing:1, transition:"all 0.15s",
                }}
                  onMouseEnter={e=>e.currentTarget.style.background="rgba(245,158,11,0.18)"}
                  onMouseLeave={e=>e.currentTarget.style.background="rgba(245,158,11,0.08)"}
                >
                  ↻ RE-RENDER AT {currentRes.hint}
                </button>
              </div>
            )}

            {/* Set-all shortcuts */}
            {slides.length > 0 && (
              <div>
                <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:8, color:"#333", letterSpacing:2, marginBottom:7, textTransform:"uppercase" }}>Set All Slides To</div>
                <div style={{ display:"flex", gap:5 }}>
                  <button onClick={() => setAllFormats("png")} style={{
                    padding:"5px 13px", borderRadius:3, border:"1px solid #004d66",
                    background:"rgba(0,77,102,0.12)", color:"#0099cc",
                    fontFamily:"'IBM Plex Mono',monospace", fontSize:10, fontWeight:700,
                    cursor:"pointer", letterSpacing:1, transition:"all 0.15s",
                  }}
                    onMouseEnter={e=>e.currentTarget.style.background="rgba(0,77,102,0.25)"}
                    onMouseLeave={e=>e.currentTarget.style.background="rgba(0,77,102,0.12)"}
                  >ALL → PNG</button>
                  <button onClick={() => setAllFormats("jpg")} style={{
                    padding:"5px 13px", borderRadius:3, border:"1px solid #6b2500",
                    background:"rgba(184,66,0,0.1)", color:"#e65100",
                    fontFamily:"'IBM Plex Mono',monospace", fontSize:10, fontWeight:700,
                    cursor:"pointer", letterSpacing:1, transition:"all 0.15s",
                  }}
                    onMouseEnter={e=>e.currentTarget.style.background="rgba(184,66,0,0.22)"}
                    onMouseLeave={e=>e.currentTarget.style.background="rgba(184,66,0,0.1)"}
                  >ALL → JPG</button>
                </div>
              </div>
            )}

            <div style={{ flex:1 }}/>

            {/* Select + Export */}
            {slides.length > 0 && (
              <div style={{ display:"flex", alignItems:"flex-end", gap:7 }}>
                <button onClick={toggleAll} style={{
                  padding:"5px 12px", borderRadius:3, border:"1px solid #222",
                  background:"transparent", color:"#555",
                  fontFamily:"'IBM Plex Mono',monospace", fontSize:9,
                  cursor:"pointer", letterSpacing:1, transition:"all 0.15s",
                }}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor="#3a3a3a";e.currentTarget.style.color="#999";}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor="#222";e.currentTarget.style.color="#555";}}
                >
                  {selectedSlides.size === slides.length ? "DESELECT ALL" : "SELECT ALL"}
                </button>
                <button
                  onClick={downloadSelected}
                  disabled={selectedCount === 0}
                  style={{
                    padding:"5px 18px", borderRadius:3, border:"none",
                    background: selectedCount > 0 ? "linear-gradient(90deg,#00e5cc,#0077ff)" : "#181818",
                    color: selectedCount > 0 ? "#000" : "#2a2a2a",
                    fontFamily:"'IBM Plex Mono',monospace", fontSize:10, fontWeight:700,
                    cursor: selectedCount > 0 ? "pointer" : "not-allowed", letterSpacing:1,
                  }}
                >
                  {selectedCount > 0
                    ? `↓ EXPORT ${selectedCount} · ${pngCount > 0 ? `${pngCount} PNG` : ""}${pngCount > 0 && jpgCount > 0 ? " + " : ""}${jpgCount > 0 ? `${jpgCount} JPG` : ""}`
                    : "↓ EXPORT"}
                </button>
              </div>
            )}
          </div>
        )}

        {/* ── Hint ── */}
        {slides.length > 0 && (
          <div style={{ fontFamily:"'IBM Plex Mono',monospace", fontSize:9, color:"#252525", marginBottom:14, letterSpacing:1 }}>
            {selectedCount}/{slides.length} SELECTED &nbsp;·&nbsp; CLICK THUMBNAIL TO TOGGLE &nbsp;·&nbsp; SWITCH PNG/JPG ON EACH CARD INDEPENDENTLY
          </div>
        )}

        {/* ── Slide Grid ── */}
        {slides.length > 0 && (
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(290px, 1fr))", gap:12 }}>
            {slides.map((slide) => (
              <div key={slide.index} style={{ animation:"fadeUp 0.3s ease-out both", animationDelay:`${Math.min(slide.index * 0.035, 0.5)}s` }}>
                <SlideCard
                  slide={slide}
                  selected={selectedSlides.has(slide.index)}
                  format={slideFormats[slide.index] || "png"}
                  onToggleSelect={() => toggleSelect(slide.index)}
                  onChangeFormat={(fmt) => setSlideFormat(slide.index, fmt)}
                  onDownload={() => downloadSlide(slide, slideFormats[slide.index] || "png")}
                  renderedRes={renderedRes}
                />
              </div>
            ))}
          </div>
        )}

        {/* ── Error ── */}
        {status === "error" && (
          <div style={{ textAlign:"center", padding:40, color:"#f87171", fontFamily:"'IBM Plex Mono',monospace", fontSize:11, letterSpacing:1 }}>
            FAILED TO PROCESS FILE — please ensure it is a valid .pptx
          </div>
        )}
      </div>
    </div>
  );
}

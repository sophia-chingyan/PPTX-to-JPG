import { useState, useRef, useCallback } from "react";

const FONT_LINK = "https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=DM+Sans:wght@300;400;500;600&display=swap";
if (!document.querySelector(`link[href*="DM+Sans"]`)) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = FONT_LINK;
  document.head.appendChild(link);
}

const RESOLUTIONS = [
  { label: "SD", w: 960, h: 540, hint: "960×540" },
  { label: "FHD", w: 1920, h: 1080, hint: "1920×1080" },
  { label: "2K", w: 2560, h: 1440, hint: "2560×1440" },
  { label: "4K", w: 3840, h: 2160, hint: "3840×2160" },
];

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (window.JSZip) return resolve();
    if (document.querySelector(`script[src="${src}"]`)) {
      const check = setInterval(() => { if (window.JSZip) { clearInterval(check); resolve(); } }, 50);
      return;
    }
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = reject;
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
    if (key.startsWith("ppt/media/")) {
      images[key] = URL.createObjectURL(await zip.files[key].async("blob"));
    }
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
    const SX = width / 12192000;
    const SY = height / 6858000;
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

function toJpgDataUrl(pngDataUrl) {
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

function Chip({ active, onClick, children, accent = "#00e5cc" }) {
  return (
    <button onClick={onClick} style={{
      padding: "5px 14px", borderRadius: 3, border: `1px solid ${active ? accent : "#2c2c2c"}`,
      background: active ? `${accent}18` : "transparent",
      color: active ? accent : "#666", fontFamily: "'IBM Plex Mono', monospace",
      fontSize: 11, fontWeight: 600, cursor: "pointer", letterSpacing: 1,
      textTransform: "uppercase", transition: "all 0.15s",
    }}>
      {children}
    </button>
  );
}

function SlideCard({ slide, selected, onToggle, onDownload, format, renderedRes }) {
  return (
    <div style={{
      background: "#111", borderRadius: 8, overflow: "hidden",
      border: `1px solid ${selected ? "#00e5cc33" : "#1e1e1e"}`,
      transition: "border-color 0.2s, transform 0.2s, box-shadow 0.2s",
    }}
      onMouseEnter={e => { e.currentTarget.style.transform = "scale(1.015)"; e.currentTarget.style.boxShadow = "0 6px 30px rgba(0,229,204,0.08)"; }}
      onMouseLeave={e => { e.currentTarget.style.transform = "scale(1)"; e.currentTarget.style.boxShadow = "none"; }}
    >
      <div style={{ position: "relative", cursor: "pointer" }} onClick={onToggle}>
        <img src={slide.dataUrl} alt={`Slide ${slide.index + 1}`} style={{
          width: "100%", display: "block",
          opacity: selected ? 1 : 0.35, transition: "opacity 0.2s",
        }} />
        {/* Checkbox */}
        <div style={{
          position: "absolute", top: 10, left: 10,
          width: 20, height: 20, borderRadius: 4,
          border: `2px solid ${selected ? "#00e5cc" : "#3a3a3a"}`,
          background: selected ? "#00e5cc" : "rgba(0,0,0,0.6)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 12, color: "#000", fontWeight: 700, transition: "all 0.15s",
        }}>
          {selected ? "✓" : ""}
        </div>
        {/* Slide number badge */}
        <div style={{
          position: "absolute", bottom: 8, right: 8,
          fontFamily: "'IBM Plex Mono', monospace", fontSize: 10,
          color: "rgba(255,255,255,0.5)", background: "rgba(0,0,0,0.55)",
          padding: "2px 7px", borderRadius: 3,
        }}>
          {String(slide.index + 1).padStart(2, "0")}
        </div>
      </div>
      <div style={{
        padding: "8px 12px", display: "flex", alignItems: "center", justifyContent: "space-between",
        borderTop: "1px solid #1a1a1a",
      }}>
        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#3a3a3a", letterSpacing: 1 }}>
          {renderedRes.w}×{renderedRes.h}
        </span>
        <button onClick={e => { e.stopPropagation(); onDownload(); }} style={{
          padding: "4px 12px", borderRadius: 3, border: "1px solid #2a2a2a",
          background: "transparent", color: "#888", fontFamily: "'IBM Plex Mono', monospace",
          fontSize: 10, cursor: "pointer", letterSpacing: 1, transition: "all 0.15s",
        }}
          onMouseEnter={e => { e.currentTarget.style.borderColor = "#00e5cc"; e.currentTarget.style.color = "#00e5cc"; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = "#2a2a2a"; e.currentTarget.style.color = "#888"; }}
        >
          ↓ .{format.toUpperCase()}
        </button>
      </div>
    </div>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function PptxToImages() {
  const [parsedSlides, setParsedSlides] = useState([]);
  const [slides, setSlides] = useState([]);         // rendered
  const [status, setStatus] = useState("idle");      // idle | loading | rerendering | done | error
  const [fileName, setFileName] = useState("");
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [format, setFormat] = useState("png");
  const [selectedSlides, setSelectedSlides] = useState(new Set());
  const [resIdx, setResIdx] = useState(1);           // default FHD
  const [renderedResIdx, setRenderedResIdx] = useState(1);
  const inputRef = useRef(null);

  const needsRerender = resIdx !== renderedResIdx && slides.length > 0;
  const currentRes = RESOLUTIONS[resIdx];
  const renderedRes = RESOLUTIONS[renderedResIdx];

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
    setSelectedSlides(new Set());
    try {
      setProgressLabel("Parsing PPTX…");
      const parsed = await parsePptx(file);
      setParsedSlides(parsed);
      const rendered = await renderAll(parsed, resIdx);
      setSlides(rendered);
      setSelectedSlides(new Set(rendered.map((_, i) => i)));
      setRenderedResIdx(resIdx);
      setStatus("done");
    } catch (e) {
      console.error(e); setStatus("error");
    }
  }, [resIdx, renderAll]);

  const handleRerender = useCallback(async () => {
    if (!parsedSlides.length) return;
    setStatus("rerendering"); setSlides([]); setProgress(0);
    try {
      const rendered = await renderAll(parsedSlides, resIdx);
      setSlides(rendered);
      setSelectedSlides(new Set(rendered.map((_, i) => i)));
      setRenderedResIdx(resIdx);
      setStatus("done");
    } catch (e) { console.error(e); setStatus("error"); }
  }, [parsedSlides, resIdx, renderAll]);

  const downloadSlide = useCallback(async (slide) => {
    const a = document.createElement("a");
    const suffix = `_${renderedRes.w}x${renderedRes.h}`;
    if (format === "jpg") {
      a.href = await toJpgDataUrl(slide.dataUrl);
      a.download = `slide_${slide.index + 1}${suffix}.jpg`;
    } else {
      a.href = slide.dataUrl;
      a.download = `slide_${slide.index + 1}${suffix}.png`;
    }
    a.click();
  }, [format, renderedRes]);

  const downloadSelected = useCallback(() => {
    slides.filter((_, i) => selectedSlides.has(i))
      .forEach((slide, idx) => setTimeout(() => downloadSlide(slide), idx * 280));
  }, [slides, selectedSlides, downloadSlide]);

  const toggleAll = () => {
    if (selectedSlides.size === slides.length) setSelectedSlides(new Set());
    else setSelectedSlides(new Set(slides.map((_, i) => i)));
  };

  const isLoading = status === "loading" || status === "rerendering";

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0a", color: "#d0d0d0", fontFamily: "'DM Sans', sans-serif", margin: 0, padding: 0 }}>
      <style>{`
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.25} }
        @keyframes fadeUp { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-track { background: #0a0a0a; }
        ::-webkit-scrollbar-thumb { background: #222; border-radius: 3px; }
      `}</style>

      {/* ── Header ── */}
      <div style={{ borderBottom: "1px solid #161616", padding: "18px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 32, height: 32, background: "linear-gradient(135deg, #00e5cc, #0077ff)", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⬡</div>
          <div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600, letterSpacing: 1, color: "#e0e0e0" }}>PPTX → IMAGE</div>
            <div style={{ fontSize: 11, color: "#3a3a3a", marginTop: 1 }}>slide export tool</div>
          </div>
        </div>
        {/* Status pill */}
        <div style={{
          display: "flex", alignItems: "center", gap: 6, padding: "4px 12px",
          borderRadius: 3, border: "1px solid #1e1e1e", background: "#111",
          fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, letterSpacing: 2,
          color: { idle: "#333", loading: "#00e5cc", rerendering: "#f59e0b", done: "#4ade80", error: "#f87171" }[status],
        }}>
          {isLoading && <span style={{ width: 5, height: 5, borderRadius: "50%", background: "currentColor", animation: "pulse 1s infinite" }} />}
          {{ idle: "IDLE", loading: "PROCESSING", rerendering: "RE-RENDERING", done: "DONE", error: "ERROR" }[status]}
        </div>
      </div>

      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "28px 24px" }}>

        {/* ── Drop Zone ── */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); }}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `1.5px dashed ${dragOver ? "#00e5cc" : "#222"}`,
            borderRadius: 10, padding: slides.length || isLoading ? "18px 24px" : "56px 24px",
            textAlign: "center", cursor: "pointer",
            background: dragOver ? "rgba(0,229,204,0.03)" : "#0e0e0e",
            transition: "all 0.25s", marginBottom: 20,
          }}
        >
          <input ref={inputRef} type="file" accept=".pptx" style={{ display: "none" }} onChange={(e) => processFile(e.target.files[0])} />
          {!slides.length && !isLoading ? (
            <>
              <div style={{ fontSize: 36, opacity: 0.2, marginBottom: 14 }}>↑</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#555", marginBottom: 6, letterSpacing: 1 }}>DROP .PPTX FILE HERE</div>
              <div style={{ fontSize: 12, color: "#333" }}>or click to browse</div>
            </>
          ) : (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 20 }}>📄</span>
                <div style={{ textAlign: "left" }}>
                  <div style={{ fontSize: 14, fontWeight: 500, color: "#ccc" }}>{fileName}</div>
                  <div style={{ fontSize: 11, color: "#444", marginTop: 2 }}>
                    {isLoading ? progressLabel : `${slides.length} slides · ${renderedRes.hint}`}
                  </div>
                </div>
              </div>
              {!isLoading && <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#333", letterSpacing: 1 }}>CLICK TO REPLACE</div>}
            </div>
          )}
        </div>

        {/* ── Progress Bar ── */}
        {isLoading && (
          <div style={{ background: "#161616", borderRadius: 4, height: 3, marginBottom: 20, overflow: "hidden" }}>
            <div style={{ height: "100%", width: `${progress}%`, background: "linear-gradient(90deg, #00e5cc, #0077ff)", borderRadius: 4, transition: "width 0.3s" }} />
          </div>
        )}

        {/* ── Toolbar ── */}
        {(slides.length > 0 || isLoading) && (
          <div style={{
            background: "#0e0e0e", border: "1px solid #191919", borderRadius: 8,
            padding: "14px 18px", marginBottom: 20,
            display: "flex", flexWrap: "wrap", gap: 16, alignItems: "flex-start",
          }}>

            {/* Format */}
            <div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 9, color: "#444", letterSpacing: 2, marginBottom: 7, textTransform: "uppercase" }}>Format</div>
              <div style={{ display: "flex", gap: 6 }}>
                {["png", "jpg"].map((f) => (
                  <Chip key={f} active={format === f} onClick={() => setFormat(f)}>
                    .{f}
                  </Chip>
                ))}
              </div>
            </div>

            {/* Resolution */}
            <div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 9, color: "#444", letterSpacing: 2, marginBottom: 7, textTransform: "uppercase" }}>
                Resolution {needsRerender && <span style={{ color: "#f59e0b", marginLeft: 4 }}>· pending re-render</span>}
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {RESOLUTIONS.map((r, i) => (
                  <Chip key={r.label} active={resIdx === i} accent={needsRerender && resIdx === i ? "#f59e0b" : "#00e5cc"} onClick={() => setResIdx(i)}>
                    {r.label} <span style={{ opacity: 0.6, fontSize: 9, marginLeft: 4 }}>{r.hint}</span>
                  </Chip>
                ))}
              </div>
            </div>

            {/* Re-render button */}
            {needsRerender && (
              <div style={{ display: "flex", alignItems: "flex-end" }}>
                <button
                  onClick={handleRerender}
                  style={{
                    padding: "6px 18px", borderRadius: 3, border: "1px solid #f59e0b",
                    background: "rgba(245,158,11,0.1)", color: "#f59e0b",
                    fontFamily: "'IBM Plex Mono', monospace", fontSize: 11,
                    fontWeight: 600, cursor: "pointer", letterSpacing: 1, transition: "all 0.15s",
                  }}
                  onMouseEnter={e => e.currentTarget.style.background = "rgba(245,158,11,0.2)"}
                  onMouseLeave={e => e.currentTarget.style.background = "rgba(245,158,11,0.1)"}
                >
                  ↻ RE-RENDER AT {currentRes.hint}
                </button>
              </div>
            )}

            {/* Spacer */}
            <div style={{ flex: 1 }} />

            {/* Selection & Download */}
            {slides.length > 0 && (
              <div style={{ display: "flex", alignItems: "flex-end", gap: 8 }}>
                <button onClick={toggleAll} style={{
                  padding: "6px 14px", borderRadius: 3, border: "1px solid #2a2a2a",
                  background: "transparent", color: "#666", fontFamily: "'IBM Plex Mono', monospace",
                  fontSize: 10, cursor: "pointer", letterSpacing: 1, transition: "all 0.15s",
                }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = "#444"; e.currentTarget.style.color = "#aaa"; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = "#2a2a2a"; e.currentTarget.style.color = "#666"; }}
                >
                  {selectedSlides.size === slides.length ? "DESELECT ALL" : "SELECT ALL"}
                </button>
                <button
                  onClick={downloadSelected}
                  disabled={selectedSlides.size === 0}
                  style={{
                    padding: "6px 20px", borderRadius: 3, border: "none",
                    background: selectedSlides.size > 0 ? "linear-gradient(90deg,#00e5cc,#0077ff)" : "#1a1a1a",
                    color: selectedSlides.size > 0 ? "#000" : "#333",
                    fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 600,
                    cursor: selectedSlides.size > 0 ? "pointer" : "not-allowed", letterSpacing: 1,
                    transition: "opacity 0.15s",
                  }}
                >
                  ↓ EXPORT {selectedSlides.size > 0 ? `${selectedSlides.size} SLIDE${selectedSlides.size > 1 ? "S" : ""}` : ""}
                </button>
              </div>
            )}
          </div>
        )}

        {/* ── Selection hint ── */}
        {slides.length > 0 && (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#2e2e2e", marginBottom: 16, letterSpacing: 1 }}>
            {selectedSlides.size} / {slides.length} SELECTED · CLICK SLIDE TO TOGGLE
          </div>
        )}

        {/* ── Slide Grid ── */}
        {slides.length > 0 && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
            {slides.map((slide) => (
              <div key={slide.index} style={{ animation: "fadeUp 0.35s ease-out both", animationDelay: `${Math.min(slide.index * 0.04, 0.6)}s` }}>
                <SlideCard
                  slide={slide}
                  selected={selectedSlides.has(slide.index)}
                  onToggle={() => setSelectedSlides((prev) => {
                    const next = new Set(prev);
                    next.has(slide.index) ? next.delete(slide.index) : next.add(slide.index);
                    return next;
                  })}
                  onDownload={() => downloadSlide(slide)}
                  format={format}
                  renderedRes={renderedRes}
                />
              </div>
            ))}
          </div>
        )}

        {/* ── Error ── */}
        {status === "error" && (
          <div style={{ textAlign: "center", padding: 40, color: "#f87171", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, letterSpacing: 1 }}>
            FAILED TO PROCESS FILE — please ensure it is a valid .pptx
          </div>
        )}
      </div>
    </div>
  );
}

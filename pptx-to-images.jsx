import { useState, useRef, useCallback } from "react";

const FONT_LINK = "https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Outfit:wght@300;400;500;600;700&display=swap";

// Inject font
if (!document.querySelector(`link[href*="Outfit"]`)) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = FONT_LINK;
  document.head.appendChild(link);
}

// Load JSZip from CDN
function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve();
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// Parse PPTX and render slides to canvas images
async function parsePptx(file) {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");

  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  // Get slide files sorted by number
  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1]);
      const nb = parseInt(b.match(/slide(\d+)/)[1]);
      return na - nb;
    });

  // Parse relationships for each slide to find images
  const images = {};
  for (const key of Object.keys(zip.files)) {
    if (key.startsWith("ppt/media/")) {
      const blob = await zip.files[key].async("blob");
      const url = URL.createObjectURL(blob);
      images[key] = url;
    }
  }

  const slides = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const slideXml = await zip.files[slideFiles[i]].async("text");
    const slideNum = parseInt(slideFiles[i].match(/slide(\d+)/)[1]);

    // Parse slide relationships for image refs
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    let slideImages = {};
    if (zip.files[relsPath]) {
      const relsXml = await zip.files[relsPath].async("text");
      const parser = new DOMParser();
      const relsDoc = parser.parseFromString(relsXml, "text/xml");
      const rels = relsDoc.getElementsByTagName("Relationship");
      for (const rel of rels) {
        const target = rel.getAttribute("Target");
        if (target && target.includes("media/")) {
          const rId = rel.getAttribute("Id");
          const fullPath = "ppt/" + target.replace("../", "");
          if (images[fullPath]) {
            slideImages[rId] = images[fullPath];
          }
        }
      }
    }

    slides.push({
      index: i,
      xml: slideXml,
      images: slideImages,
    });
  }

  return slides;
}

// Render a slide to a canvas and return a data URL
function renderSlideToCanvas(slide, width = 1920, height = 1080) {
  return new Promise((resolve) => {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");

    // White background
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    const parser = new DOMParser();
    const doc = parser.parseFromString(slide.xml, "text/xml");

    // EMU to pixels conversion (1 inch = 914400 EMU, 96 DPI)
    const emuToPx = (emu) => (parseInt(emu) || 0) * 96 / 914400;

    // Extract shapes from the slide
    const spTree = doc.getElementsByTagName("p:spTree")[0];
    if (!spTree) {
      resolve(canvas.toDataURL("image/png"));
      return;
    }

    const textItems = [];
    const imageItems = [];

    // Process all shape types
    const shapes = [
      ...Array.from(doc.getElementsByTagName("p:sp")),
      ...Array.from(doc.getElementsByTagName("p:pic")),
    ];

    for (const shape of shapes) {
      // Get position and size
      const off = shape.getElementsByTagName("a:off")[0];
      const ext = shape.getElementsByTagName("a:ext")[0];
      if (!off || !ext) continue;

      const x = emuToPx(off.getAttribute("x"));
      const y = emuToPx(off.getAttribute("y"));
      const w = emuToPx(ext.getAttribute("cx"));
      const h = emuToPx(ext.getAttribute("cy"));

      // Check for image (blipFill)
      const blip = shape.getElementsByTagName("a:blip")[0];
      if (blip) {
        const rEmbed = blip.getAttribute("r:embed");
        if (rEmbed && slide.images[rEmbed]) {
          imageItems.push({ src: slide.images[rEmbed], x, y, w, h });
        }
        continue;
      }

      // Check for solid fill on the shape
      const spPr = shape.getElementsByTagName("p:spPr")[0] || shape.getElementsByTagName("a:spPr")[0];
      if (spPr) {
        const solidFill = spPr.getElementsByTagName("a:solidFill")[0];
        if (solidFill) {
          const srgbClr = solidFill.getElementsByTagName("a:srgbClr")[0];
          if (srgbClr) {
            const color = "#" + srgbClr.getAttribute("val");
            ctx.fillStyle = color;
            ctx.fillRect(x * (width / 1920) * 2, y * (height / 1080) * 2, w * (width / 1920) * 2, h * (height / 1080) * 2);
          }
        }
      }

      // Extract text
      const txBody = shape.getElementsByTagName("p:txBody")[0];
      if (!txBody) continue;

      const paragraphs = txBody.getElementsByTagName("a:p");
      let yOffset = 0;

      for (const para of paragraphs) {
        const runs = para.getElementsByTagName("a:r");
        let paraText = "";
        let fontSize = 18;
        let fontBold = false;
        let fontColor = "#000000";

        for (const run of runs) {
          const rPr = run.getElementsByTagName("a:rPr")[0];
          if (rPr) {
            const sz = rPr.getAttribute("sz");
            if (sz) fontSize = parseInt(sz) / 100;
            const b = rPr.getAttribute("b");
            if (b === "1") fontBold = true;
            const solidFill = rPr.getElementsByTagName("a:solidFill")[0];
            if (solidFill) {
              const srgb = solidFill.getElementsByTagName("a:srgbClr")[0];
              if (srgb) fontColor = "#" + srgb.getAttribute("val");
            }
          }
          const t = run.getElementsByTagName("a:t")[0];
          if (t) paraText += t.textContent;
        }

        if (paraText.trim()) {
          textItems.push({
            text: paraText,
            x, y: y + yOffset,
            w, h,
            fontSize,
            fontBold,
            fontColor,
          });
        }
        yOffset += fontSize * 1.5;
      }
    }

    // Scale factor
    const scaleX = width / 1920 * 2;
    const scaleY = height / 1080 * 2;

    // Draw images first, then text on top
    const imagePromises = imageItems.map((item) => {
      return new Promise((res) => {
        const img = new Image();
        img.crossOrigin = "anonymous";
        img.onload = () => {
          ctx.drawImage(img, item.x * scaleX, item.y * scaleY, item.w * scaleX, item.h * scaleY);
          res();
        };
        img.onerror = () => res();
        img.src = item.src;
      });
    });

    Promise.all(imagePromises).then(() => {
      // Draw text
      for (const item of textItems) {
        const scaledSize = Math.max(12, item.fontSize * scaleX * 0.5);
        ctx.font = `${item.fontBold ? "bold " : ""}${scaledSize}px Arial, sans-serif`;
        ctx.fillStyle = item.fontColor;
        ctx.textBaseline = "top";

        // Word wrap
        const maxWidth = item.w * scaleX;
        const words = item.text.split(" ");
        let line = "";
        let lineY = item.y * scaleY;

        for (const word of words) {
          const testLine = line + (line ? " " : "") + word;
          const metrics = ctx.measureText(testLine);
          if (metrics.width > maxWidth && line) {
            ctx.fillText(line, item.x * scaleX, lineY);
            line = word;
            lineY += scaledSize * 1.3;
          } else {
            line = testLine;
          }
        }
        ctx.fillText(line, item.x * scaleX, lineY);
      }

      resolve(canvas.toDataURL("image/png"));
    });
  });
}

// Status badge component
function StatusBadge({ status }) {
  const colors = {
    idle: { bg: "#2a2a2a", text: "#666", label: "WAITING" },
    loading: { bg: "#1a1a2e", text: "#7c83ff", label: "PROCESSING" },
    done: { bg: "#0a2a1a", text: "#4ade80", label: "COMPLETE" },
    error: { bg: "#2a1a1a", text: "#ff6b6b", label: "ERROR" },
  };
  const c = colors[status] || colors.idle;
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "4px 12px",
        borderRadius: 4,
        background: c.bg,
        color: c.text,
        fontFamily: "'Space Mono', monospace",
        fontSize: 11,
        fontWeight: 700,
        letterSpacing: 2,
        textTransform: "uppercase",
      }}
    >
      {status === "loading" && (
        <span
          style={{
            width: 6,
            height: 6,
            borderRadius: "50%",
            background: c.text,
            animation: "pulse 1s infinite",
          }}
        />
      )}
      {c.label}
    </span>
  );
}

export default function PptxToImages() {
  const [slides, setSlides] = useState([]);
  const [status, setStatus] = useState("idle");
  const [fileName, setFileName] = useState("");
  const [progress, setProgress] = useState(0);
  const [dragOver, setDragOver] = useState(false);
  const [format, setFormat] = useState("png");
  const [selectedSlides, setSelectedSlides] = useState(new Set());
  const inputRef = useRef(null);

  const processFile = useCallback(
    async (file) => {
      if (!file || !file.name.match(/\.pptx$/i)) {
        setStatus("error");
        return;
      }

      setFileName(file.name);
      setStatus("loading");
      setSlides([]);
      setProgress(0);
      setSelectedSlides(new Set());

      try {
        const parsedSlides = await parsePptx(file);
        const rendered = [];

        for (let i = 0; i < parsedSlides.length; i++) {
          setProgress(((i + 1) / parsedSlides.length) * 100);
          const dataUrl = await renderSlideToCanvas(parsedSlides[i]);
          rendered.push({ index: i, dataUrl });
        }

        setSlides(rendered);
        setSelectedSlides(new Set(rendered.map((_, i) => i)));
        setStatus("done");
      } catch (e) {
        console.error(e);
        setStatus("error");
      }
    },
    []
  );

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      processFile(file);
    },
    [processFile]
  );

  const toggleSlide = (i) => {
    setSelectedSlides((prev) => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i);
      else next.add(i);
      return next;
    });
  };

  const downloadSlide = (slide) => {
    const a = document.createElement("a");
    const ext = format === "jpg" ? "jpeg" : format;
    // Convert if jpg
    if (format === "jpg") {
      const canvas = document.createElement("canvas");
      const img = new Image();
      img.onload = () => {
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0);
        a.href = canvas.toDataURL("image/jpeg", 0.92);
        a.download = `slide_${slide.index + 1}.jpg`;
        a.click();
      };
      img.src = slide.dataUrl;
    } else {
      a.href = slide.dataUrl;
      a.download = `slide_${slide.index + 1}.png`;
      a.click();
    }
  };

  const downloadAll = () => {
    slides
      .filter((_, i) => selectedSlides.has(i))
      .forEach((slide, idx) => {
        setTimeout(() => downloadSlide(slide), idx * 300);
      });
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#0d0d0d",
        color: "#e0e0e0",
        fontFamily: "'Outfit', sans-serif",
        padding: 0,
        margin: 0,
      }}
    >
      <style>{`
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.3; }
        }
        @keyframes slideUp {
          from { opacity: 0; transform: translateY(20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes barGrow {
          from { width: 0%; }
        }
        .slide-card:hover { transform: scale(1.02); box-shadow: 0 8px 40px rgba(124,131,255,0.15); }
        .slide-card { transition: transform 0.2s, box-shadow 0.2s; }
        .dl-btn:hover { background: #7c83ff !important; color: #000 !important; }
        .format-btn:hover { border-color: #7c83ff !important; }
      `}</style>

      {/* Header */}
      <div
        style={{
          borderBottom: "1px solid #1a1a1a",
          padding: "20px 32px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div
            style={{
              width: 36,
              height: 36,
              background: "linear-gradient(135deg, #7c83ff, #4a4fbf)",
              borderRadius: 8,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 18,
            }}
          >
            ⬡
          </div>
          <div>
            <div
              style={{
                fontFamily: "'Space Mono', monospace",
                fontSize: 15,
                fontWeight: 700,
                letterSpacing: 1,
              }}
            >
              PPTX → IMG
            </div>
            <div style={{ fontSize: 11, color: "#555", marginTop: 2 }}>
              PowerPoint slide converter
            </div>
          </div>
        </div>
        <StatusBadge status={status} />
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "32px 24px" }}>
        {/* Drop Zone */}
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? "#7c83ff" : "#2a2a2a"}`,
            borderRadius: 12,
            padding: slides.length ? "24px" : "64px 24px",
            textAlign: "center",
            cursor: "pointer",
            background: dragOver ? "rgba(124,131,255,0.05)" : "#111",
            transition: "all 0.3s",
            marginBottom: 24,
          }}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".pptx"
            style={{ display: "none" }}
            onChange={(e) => processFile(e.target.files[0])}
          />
          {!slides.length && status !== "loading" ? (
            <>
              <div
                style={{
                  fontSize: 48,
                  marginBottom: 16,
                  opacity: 0.3,
                }}
              >
                ↑
              </div>
              <div
                style={{
                  fontFamily: "'Space Mono', monospace",
                  fontSize: 14,
                  color: "#888",
                  marginBottom: 8,
                }}
              >
                DROP .PPTX FILE HERE
              </div>
              <div style={{ fontSize: 13, color: "#444" }}>
                or click to browse
              </div>
            </>
          ) : (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{ fontSize: 20 }}>📄</span>
                <div style={{ textAlign: "left" }}>
                  <div style={{ fontSize: 14, fontWeight: 500 }}>
                    {fileName}
                  </div>
                  <div style={{ fontSize: 12, color: "#555" }}>
                    {slides.length
                      ? `${slides.length} slides converted`
                      : "Processing..."}
                  </div>
                </div>
              </div>
              <div
                style={{
                  fontFamily: "'Space Mono', monospace",
                  fontSize: 11,
                  color: "#555",
                }}
              >
                Click to replace
              </div>
            </div>
          )}
        </div>

        {/* Progress Bar */}
        {status === "loading" && (
          <div
            style={{
              background: "#1a1a1a",
              borderRadius: 6,
              height: 4,
              marginBottom: 24,
              overflow: "hidden",
            }}
          >
            <div
              style={{
                height: "100%",
                width: `${progress}%`,
                background: "linear-gradient(90deg, #7c83ff, #a78bfa)",
                borderRadius: 6,
                transition: "width 0.3s",
                animation: "barGrow 0.5s ease-out",
              }}
            />
          </div>
        )}

        {/* Controls */}
        {slides.length > 0 && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: 24,
              flexWrap: "wrap",
              gap: 12,
            }}
          >
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <span
                style={{
                  fontFamily: "'Space Mono', monospace",
                  fontSize: 11,
                  color: "#555",
                  letterSpacing: 1,
                }}
              >
                FORMAT
              </span>
              {["png", "jpg"].map((f) => (
                <button
                  key={f}
                  className="format-btn"
                  onClick={() => setFormat(f)}
                  style={{
                    padding: "6px 16px",
                    borderRadius: 6,
                    border: `1px solid ${format === f ? "#7c83ff" : "#2a2a2a"}`,
                    background: format === f ? "rgba(124,131,255,0.1)" : "transparent",
                    color: format === f ? "#7c83ff" : "#666",
                    fontFamily: "'Space Mono', monospace",
                    fontSize: 12,
                    fontWeight: 700,
                    cursor: "pointer",
                    letterSpacing: 1,
                    textTransform: "uppercase",
                  }}
                >
                  .{f}
                </button>
              ))}
            </div>

            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <button
                onClick={() => {
                  if (selectedSlides.size === slides.length) {
                    setSelectedSlides(new Set());
                  } else {
                    setSelectedSlides(new Set(slides.map((_, i) => i)));
                  }
                }}
                style={{
                  padding: "8px 16px",
                  borderRadius: 6,
                  border: "1px solid #2a2a2a",
                  background: "transparent",
                  color: "#888",
                  fontSize: 12,
                  cursor: "pointer",
                  fontFamily: "'Outfit', sans-serif",
                }}
              >
                {selectedSlides.size === slides.length
                  ? "Deselect All"
                  : "Select All"}
              </button>
              <button
                className="dl-btn"
                onClick={downloadAll}
                disabled={selectedSlides.size === 0}
                style={{
                  padding: "8px 24px",
                  borderRadius: 6,
                  border: "none",
                  background:
                    selectedSlides.size > 0 ? "#5a5fff" : "#1a1a1a",
                  color: selectedSlides.size > 0 ? "#fff" : "#444",
                  fontFamily: "'Space Mono', monospace",
                  fontSize: 12,
                  fontWeight: 700,
                  cursor:
                    selectedSlides.size > 0 ? "pointer" : "not-allowed",
                  letterSpacing: 1,
                }}
              >
                ↓ DOWNLOAD {selectedSlides.size > 0 ? `(${selectedSlides.size})` : ""}
              </button>
            </div>
          </div>
        )}

        {/* Slide Grid */}
        {slides.length > 0 && (
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))",
              gap: 16,
            }}
          >
            {slides.map((slide) => (
              <div
                key={slide.index}
                className="slide-card"
                style={{
                  background: "#141414",
                  borderRadius: 10,
                  overflow: "hidden",
                  border: `1px solid ${selectedSlides.has(slide.index) ? "#7c83ff44" : "#1e1e1e"}`,
                  animation: "slideUp 0.4s ease-out",
                  animationDelay: `${slide.index * 0.05}s`,
                  animationFillMode: "both",
                }}
              >
                {/* Slide image */}
                <div
                  style={{
                    position: "relative",
                    cursor: "pointer",
                  }}
                  onClick={() => toggleSlide(slide.index)}
                >
                  <img
                    src={slide.dataUrl}
                    alt={`Slide ${slide.index + 1}`}
                    style={{
                      width: "100%",
                      display: "block",
                      opacity: selectedSlides.has(slide.index) ? 1 : 0.4,
                      transition: "opacity 0.2s",
                    }}
                  />
                  {/* Selection indicator */}
                  <div
                    style={{
                      position: "absolute",
                      top: 10,
                      left: 10,
                      width: 22,
                      height: 22,
                      borderRadius: 5,
                      border: `2px solid ${selectedSlides.has(slide.index) ? "#7c83ff" : "#555"}`,
                      background: selectedSlides.has(slide.index)
                        ? "#7c83ff"
                        : "rgba(0,0,0,0.5)",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 14,
                      color: "#fff",
                      fontWeight: 700,
                      transition: "all 0.2s",
                    }}
                  >
                    {selectedSlides.has(slide.index) ? "✓" : ""}
                  </div>
                </div>

                {/* Slide footer */}
                <div
                  style={{
                    padding: "10px 14px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                  }}
                >
                  <span
                    style={{
                      fontFamily: "'Space Mono', monospace",
                      fontSize: 12,
                      color: "#555",
                    }}
                  >
                    SLIDE {String(slide.index + 1).padStart(2, "0")}
                  </span>
                  <button
                    className="dl-btn"
                    onClick={(e) => {
                      e.stopPropagation();
                      downloadSlide(slide);
                    }}
                    style={{
                      padding: "4px 14px",
                      borderRadius: 4,
                      border: "1px solid #2a2a2a",
                      background: "transparent",
                      color: "#888",
                      fontFamily: "'Space Mono', monospace",
                      fontSize: 11,
                      cursor: "pointer",
                      letterSpacing: 1,
                    }}
                  >
                    ↓ .{format.toUpperCase()}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Error state */}
        {status === "error" && (
          <div
            style={{
              textAlign: "center",
              padding: 40,
              color: "#ff6b6b",
              fontFamily: "'Space Mono', monospace",
              fontSize: 13,
            }}
          >
            Failed to process file. Please ensure it's a valid .pptx file.
          </div>
        )}
      </div>
    </div>
  );
}

import React, { useState, useEffect } from "react";

// ── Load PptxGenJS from CDN ──────────────────────────────────────────────────
function usePptxGen() {
  const [ready, setReady] = useState(!!window.PptxGenJS);
  useEffect(() => {
    if (window.PptxGenJS) { setReady(true); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js";
    s.onload = () => setReady(true);
    document.head.appendChild(s);
  }, []);
  return ready;
}

// ── Constants ────────────────────────────────────────────────────────────────
const CAMPAIGN_COLORS = [
  { label: "Mint",       value: "#7DD4B8", text: "#1a4a3a" },
  { label: "Rose",       value: "#E8A0B4", text: "#4a1a2a" },
  { label: "Mauve",      value: "#C17BAD", text: "#ffffff" },
  { label: "Sand",       value: "#E8D5B0", text: "#4a3a1a" },
  { label: "Sky Blue",   value: "#A8D4E8", text: "#1a3a4a" },
  { label: "Terracotta", value: "#C89080", text: "#ffffff" },
  { label: "Lavender",   value: "#B0A8D4", text: "#1a1a4a" },
  { label: "Sage",       value: "#A8C4A0", text: "#1a3a1a" },
  { label: "Peach",      value: "#F0C0A0", text: "#4a2a1a" },
  { label: "Steel",      value: "#A0B4C8", text: "#ffffff" },
];

const CAMPAIGN_TYPES = [
  "All Types",
  "TX Wave",
  "Retail",
  "Bold Ruby",
  "ZYN",
  "IQOSPHERE",
  "Ongoing Campaign",
  "Offline",
  "Loyalty",
  "Other",
];

const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const DAY_NAMES   = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

const getDaysInMonth  = (y, m) => new Date(y, m + 1, 0).getDate();
const getFirstDayOfMonth = (y, m) => new Date(y, m, 1).getDay();

function buildCalendarDays(year, month) {
  const days = [];
  const first = getFirstDayOfMonth(year, month);
  for (let i = 0; i < first; i++) days.push(null);
  for (let d = 1; d <= getDaysInMonth(year, month); d++) days.push(d);
  while (days.length % 7 !== 0) days.push(null);
  return days;
}

let _id = 10;
const uid = () => ++_id;

const hexToRgb = hex => {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return { r, g, b };
};

// ── Seed data ────────────────────────────────────────────────────────────────
const today = new Date();
const SEED = [
  { id:1, name:"TX Wave 4",      subtitle:"Prize Refresh",       notes:"Confirm assets by Mon", type:"TX Wave",          color:"#7DD4B8", textColor:"#1a4a3a", date:{year:today.getFullYear(),month:today.getMonth(),day:3}  },
  { id:2, name:"Retail 2.0",     subtitle:"Grand Opening",       notes:"",                      type:"Retail",           color:"#C17BAD", textColor:"#ffffff", date:{year:today.getFullYear(),month:today.getMonth(),day:5}  },
  { id:3, name:"Bold Ruby",      subtitle:"LAS Launch",          notes:"Hajimetewari collab",   type:"Bold Ruby",        color:"#C89080", textColor:"#ffffff", date:{year:today.getFullYear(),month:today.getMonth(),day:10} },
  { id:4, name:"ZYN",            subtitle:"FSS Multicategory",   notes:"",                      type:"ZYN",              color:"#A8D4E8", textColor:"#1a3a4a", date:{year:today.getFullYear(),month:today.getMonth(),day:12} },
  { id:5, name:"IQOSPHERE",      subtitle:"BAU Earn & Burn",     notes:"Check token limits",    type:"IQOSPHERE",        color:"#A8C4A0", textColor:"#1a3a1a", date:{year:today.getFullYear(),month:today.getMonth(),day:15} },
  { id:6, name:"Loyalty 2027",   subtitle:"IQOSPHERE Retirement",notes:"",                      type:"Loyalty",          color:"#B0A8D4", textColor:"#1a1a4a", date:{year:today.getFullYear(),month:today.getMonth(),day:20} },
];

// ── Main component ───────────────────────────────────────────────────────────
export default function CampaignCalendar() {
  const pptxReady = usePptxGen();
  const [viewYear,  setViewYear]  = useState(today.getFullYear());
  const [viewMonth, setViewMonth] = useState(today.getMonth());
  const [campaigns, setCampaigns] = useState(SEED);
  const [filterType, setFilterType] = useState("All Types");

  // Modal state
  const [modal, setModal]     = useState(false);
  const [editId, setEditId]   = useState(null);
  const [form,  setForm]      = useState({ name:"", subtitle:"", notes:"", type:CAMPAIGN_TYPES[1], color:CAMPAIGN_COLORS[0].value, textColor:CAMPAIGN_COLORS[0].text, day:"" });

  // Drag state
  const [dragId,      setDragId]      = useState(null);
  const [dragOverDay, setDragOverDay] = useState(null);

  // Export state
  const [exporting, setExporting] = useState(false);

  const calDays     = buildCalendarDays(viewYear, viewMonth);
  const daysInMonth = getDaysInMonth(viewYear, viewMonth);

  const visibleCampaigns = campaigns.filter(c =>
    c.date.year === viewYear && c.date.month === viewMonth &&
    (filterType === "All Types" || c.type === filterType)
  );

  const forDay = day => visibleCampaigns.filter(c => c.date.day === day);

  // Navigation
  const prevMonth = () => viewMonth === 0 ? (setViewMonth(11), setViewYear(y=>y-1)) : setViewMonth(m=>m-1);
  const nextMonth = () => viewMonth === 11 ? (setViewMonth(0),  setViewYear(y=>y+1)) : setViewMonth(m=>m+1);

  // Modal helpers
  const openNew = (day="") => {
    setEditId(null);
    setForm({ name:"", subtitle:"", notes:"", type:CAMPAIGN_TYPES[1], color:CAMPAIGN_COLORS[0].value, textColor:CAMPAIGN_COLORS[0].text, day:day||"" });
    setModal(true);
  };
  const openEdit = (c, e) => {
    e.stopPropagation();
    setEditId(c.id);
    setForm({ name:c.name, subtitle:c.subtitle, notes:c.notes||"", type:c.type, color:c.color, textColor:c.textColor, day:c.date.day });
    setModal(true);
  };
  const pickColor = v => {
    const m = CAMPAIGN_COLORS.find(c=>c.value===v);
    setForm(f=>({...f, color:v, textColor: m ? m.text : "#ffffff"}));
  };
  const handleSave = () => {
    if (!form.name.trim() || !form.day) return;
    const day = parseInt(form.day);
    if (editId) {
      setCampaigns(cs => cs.map(c => c.id===editId ? {...c,...form, date:{year:viewYear,month:viewMonth,day}} : c));
    } else {
      setCampaigns(cs => [...cs, {id:uid(),...form, date:{year:viewYear,month:viewMonth,day}}]);
    }
    setModal(false);
  };
  const handleDelete = () => { setCampaigns(cs=>cs.filter(c=>c.id!==editId)); setModal(false); };

  // Drag & drop
  const onDragStart = (e, id) => { setDragId(id); e.dataTransfer.effectAllowed="move"; };
  const onDragOver  = (e, day) => { if(!day) return; e.preventDefault(); setDragOverDay(day); };
  const onDrop      = (e, day) => {
    e.preventDefault();
    if (dragId && day) setCampaigns(cs => cs.map(c => c.id===dragId ? {...c, date:{year:viewYear,month:viewMonth,day}} : c));
    setDragId(null); setDragOverDay(null);
  };
  const onDragEnd   = () => { setDragId(null); setDragOverDay(null); };

  // ── PPTX Export ─────────────────────────────────────────────────────────────
  const exportPptx = async () => {
    if (!pptxReady) return;
    setExporting(true);
    try {
      const pptx = new window.PptxGenJS();
      pptx.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"

      const W = 13.33, H = 7.5;
      const MARGIN = 0.3;
      const HEADER_H = 0.7;
      const DAY_LABEL_H = 0.28;
      const GRID_TOP = MARGIN + HEADER_H + DAY_LABEL_H;
      const GRID_H   = H - GRID_TOP - MARGIN;
      const COLS = 7;
      const first = getFirstDayOfMonth(viewYear, viewMonth);
      const totalDays = getDaysInMonth(viewYear, viewMonth);
      const totalCells = Math.ceil((first + totalDays) / 7) * 7;
      const ROWS = totalCells / 7;
      const CELL_W = (W - MARGIN * 2) / COLS;
      const CELL_H = GRID_H / ROWS;

      const exportList = filterType === "All Types"
        ? campaigns.filter(c => c.date.year===viewYear && c.date.month===viewMonth)
        : campaigns.filter(c => c.date.year===viewYear && c.date.month===viewMonth && c.type===filterType);

      const slide = pptx.addSlide();

      // Background
      slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:"F7F8FA"}, line:{color:"F7F8FA"} });

      // Title bar
      slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:HEADER_H, fill:{color:"1a1d2e"}, line:{color:"1a1d2e"} });
      slide.addText(`${MONTH_NAMES[viewMonth]} ${viewYear} — Campaign Calendar`, {
        x:MARGIN, y:0.1, w:8, h:0.5,
        fontSize:22, bold:true, color:"FFFFFF", fontFace:"Calibri", valign:"middle"
      });
      if (filterType !== "All Types") {
        slide.addText(`Filtered: ${filterType}`, {
          x:W-3.5, y:0.15, w:3.2, h:0.4,
          fontSize:11, color:"A8D4E8", fontFace:"Calibri", align:"right", valign:"middle"
        });
      }

      // Day headers
      DAY_NAMES.forEach((d, i) => {
        const x = MARGIN + i * CELL_W;
        slide.addShape(pptx.ShapeType.rect, { x, y:MARGIN+HEADER_H, w:CELL_W, h:DAY_LABEL_H, fill:{color:"2d3250"}, line:{color:"2d3250"} });
        slide.addText(d, { x, y:MARGIN+HEADER_H, w:CELL_W, h:DAY_LABEL_H, fontSize:9, bold:true, color:"FFFFFF", fontFace:"Calibri", align:"center", valign:"middle" });
      });

      // Grid cells
      for (let i = 0; i < totalCells; i++) {
        const col = i % COLS;
        const row = Math.floor(i / COLS);
        const day = i - first + 1;
        const isValid = day >= 1 && day <= totalDays;
        const cx = MARGIN + col * CELL_W;
        const cy = GRID_TOP + row * CELL_H;

        slide.addShape(pptx.ShapeType.rect, {
          x:cx, y:cy, w:CELL_W, h:CELL_H,
          fill:{ color: isValid ? "FFFFFF" : "F0F1F5" },
          line:{ color:"E0E3EC", pt:0.5 }
        });

        if (isValid) {
          // Day number
          const isToday = viewYear===today.getFullYear() && viewMonth===today.getMonth() && day===today.getDate();
          if (isToday) {
            slide.addShape(pptx.ShapeType.ellipse, { x:cx+0.06, y:cy+0.04, w:0.22, h:0.22, fill:{color:"3B5BDB"}, line:{color:"3B5BDB"} });
            slide.addText(String(day), { x:cx+0.06, y:cy+0.04, w:0.22, h:0.22, fontSize:8, bold:true, color:"FFFFFF", fontFace:"Calibri", align:"center", valign:"middle" });
          } else {
            slide.addText(String(day), { x:cx+0.06, y:cy+0.04, w:0.22, h:0.22, fontSize:8, bold:true, color:"4a5068", fontFace:"Calibri", align:"center", valign:"middle" });
          }

          // Campaign boxes
          const dayCampaigns = exportList.filter(c => c.date.day === day);
          dayCampaigns.forEach((c, ci) => {
            const boxH = 0.28;
            const boxY = cy + 0.28 + ci * (boxH + 0.04);
            if (boxY + boxH > cy + CELL_H - 0.04) return; // overflow guard
            const bg = c.color.replace("#","");
            const fg = c.textColor.replace("#","");
            slide.addShape(pptx.ShapeType.roundRect, {
              x: cx+0.06, y: boxY, w: CELL_W-0.12, h: boxH,
              fill:{color:bg}, line:{color:bg}, rectRadius:0.05
            });
            // Name line
            slide.addText(c.name, {
              x:cx+0.09, y:boxY+0.01, w:CELL_W-0.18, h:0.13,
              fontSize:7.5, bold:true, color:fg, fontFace:"Calibri", valign:"middle"
            });
            // Subtitle line
            if (c.subtitle) {
              slide.addText(c.subtitle, {
                x:cx+0.09, y:boxY+0.13, w:CELL_W-0.18, h:0.1,
                fontSize:6, color:fg, fontFace:"Calibri", valign:"top"
              });
            }
            // Notes line
            if (c.notes) {
              slide.addText(`📝 ${c.notes}`, {
                x:cx+0.09, y:boxY+0.2, w:CELL_W-0.18, h:0.09,
                fontSize:5.5, color:fg, fontFace:"Calibri", valign:"top", italic:true
              });
            }
          });

          const overflow = exportList.filter(c => c.date.day === day).length;
          const shown = Math.floor((CELL_H - 0.32) / 0.32);
          if (overflow > shown) {
            slide.addText(`+${overflow - shown} more`, {
              x:cx+0.06, y:cy+CELL_H-0.18, w:CELL_W-0.12, h:0.14,
              fontSize:6, color:"8a90a0", fontFace:"Calibri", align:"center"
            });
          }
        }
      }

      // Footer
      slide.addText(`Exported ${new Date().toLocaleDateString()} · Campaign Calendar`, {
        x:MARGIN, y:H-0.25, w:W-MARGIN*2, h:0.2,
        fontSize:7, color:"8a90a0", fontFace:"Calibri", align:"right"
      });

      const fname = `Campaign_Calendar_${MONTH_NAMES[viewMonth]}_${viewYear}${filterType!=="All Types"?`_${filterType.replace(/\s/g,"_")}`:""}`;
      await pptx.writeFile({ fileName: fname });
    } catch(err) {
      console.error("Export failed:", err);
    }
    setExporting(false);
  };

  // ── Render ───────────────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',sans-serif", minHeight:"100vh", background:"#F7F8FA", padding:24 }}>
      <div style={{ maxWidth:1200, margin:"0 auto" }}>

        {/* ── Top bar ── */}
        <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:12 }}>
          <div>
            <h1 style={{ margin:0, fontSize:22, fontWeight:700, color:"#1a1d2e", letterSpacing:"-0.3px" }}>Campaign Calendar</h1>
            <p style={{ margin:"4px 0 0", fontSize:13, color:"#8a90a0" }}>Drag campaigns between dates · Click a date to add</p>
          </div>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
            {/* Filter */}
            <select value={filterType} onChange={e=>setFilterType(e.target.value)} style={selectStyle}>
              {CAMPAIGN_TYPES.map(t=><option key={t}>{t}</option>)}
            </select>
            {/* Export */}
            <button onClick={exportPptx} disabled={!pptxReady||exporting} style={{...btnStyle, background:"#2d3250", color:"#fff", opacity:(!pptxReady||exporting)?0.6:1}}>
              {exporting ? "Exporting…" : "⬇ Export PPTX"}
            </button>
            {/* Add */}
            <button onClick={()=>openNew()} style={{...btnStyle, background:"#3B5BDB", color:"#fff"}}>
              + Add Campaign
            </button>
          </div>
        </div>

        {/* ── Month nav ── */}
        <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:14 }}>
          <button onClick={prevMonth} style={navBtn}>‹</button>
          <h2 style={{ margin:0, fontSize:18, fontWeight:700, color:"#1a1d2e", minWidth:200, textAlign:"center" }}>
            {MONTH_NAMES[viewMonth]} {viewYear}
          </h2>
          <button onClick={nextMonth} style={navBtn}>›</button>
          {filterType !== "All Types" && (
            <span style={{ fontSize:12, background:"#EEF2FF", color:"#3B5BDB", borderRadius:20, padding:"3px 10px", fontWeight:600 }}>
              {filterType} <span style={{cursor:"pointer",marginLeft:4}} onClick={()=>setFilterType("All Types")}>✕</span>
            </span>
          )}
        </div>

        {/* ── Calendar ── */}
        <div style={{ background:"#fff", borderRadius:14, boxShadow:"0 2px 12px rgba(0,0,0,0.07)", overflow:"hidden" }}>
          {/* Day headers */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", borderBottom:"1px solid #eef0f5" }}>
            {DAY_NAMES.map(d=>(
              <div key={d} style={{ padding:"10px 0", textAlign:"center", fontSize:12, fontWeight:700, color:"#8a90a0", letterSpacing:"0.5px", textTransform:"uppercase" }}>{d}</div>
            ))}
          </div>

          {/* Cells */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)" }}>
            {calDays.map((day, i) => {
              const isToday = day && viewYear===today.getFullYear() && viewMonth===today.getMonth() && day===today.getDate();
              const isOver  = dragOverDay===day && day!==null;
              const dayCamps = day ? forDay(day) : [];
              return (
                <div
                  key={i}
                  onClick={()=>day && openNew(day)}
                  onDragOver={e=>onDragOver(e,day)}
                  onDrop={e=>onDrop(e,day)}
                  style={{
                    minHeight:120,
                    borderRight:(i+1)%7===0?"none":"1px solid #eef0f5",
                    borderBottom:i<calDays.length-7?"1px solid #eef0f5":"none",
                    padding:"7px 5px 5px",
                    cursor:day?"pointer":"default",
                    background:isOver?"#EEF2FF":day?"#fff":"#FAFAFA",
                    transition:"background 0.15s",
                  }}
                >
                  {day && (
                    <>
                      <div style={{ fontSize:12, fontWeight:600, color:isToday?"#fff":"#4a5068",
                        background:isToday?"#3B5BDB":"transparent", borderRadius:"50%",
                        width:24, height:24, display:"flex", alignItems:"center", justifyContent:"center", marginBottom:4 }}>
                        {day}
                      </div>
                      <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
                        {dayCamps.map(c=>(
                          <div
                            key={c.id}
                            draggable
                            onDragStart={e=>{e.stopPropagation();onDragStart(e,c.id);}}
                            onDragEnd={onDragEnd}
                            onClick={e=>openEdit(c,e)}
                            style={{
                              background:c.color, color:c.textColor,
                              borderRadius:5, padding:"4px 6px", fontSize:11, fontWeight:700,
                              cursor:"grab", opacity:dragId===c.id?0.35:1,
                              boxShadow:"0 1px 3px rgba(0,0,0,0.12)", userSelect:"none", lineHeight:1.3,
                            }}
                          >
                            <div>{c.name}</div>
                            {c.subtitle && <div style={{ fontWeight:400, fontSize:10, opacity:0.85 }}>{c.subtitle}</div>}
                            {c.notes    && <div style={{ fontWeight:400, fontSize:9,  opacity:0.7,  marginTop:2, fontStyle:"italic" }}>📝 {c.notes}</div>}
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* ── Legend ── */}
        <div style={{ marginTop:14, display:"flex", flexWrap:"wrap", gap:8 }}>
          {CAMPAIGN_COLORS.map(c=>(
            <div key={c.value} style={{ display:"flex", alignItems:"center", gap:5, fontSize:11, color:"#6a7080" }}>
              <div style={{ width:12, height:12, borderRadius:3, background:c.value }} />
              {c.label}
            </div>
          ))}
        </div>
      </div>

      {/* ── Modal ── */}
      {modal && (
        <div style={{ position:"fixed", inset:0, background:"rgba(10,15,40,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:200 }}
          onClick={()=>setModal(false)}>
          <div style={{ background:"#fff", borderRadius:14, padding:28, width:440, maxWidth:"92vw", boxShadow:"0 20px 60px rgba(0,0,0,0.2)", maxHeight:"90vh", overflowY:"auto" }}
            onClick={e=>e.stopPropagation()}>
            <h3 style={{ margin:"0 0 18px", fontSize:17, fontWeight:700, color:"#1a1d2e" }}>
              {editId ? "Edit Campaign" : "New Campaign"}
            </h3>

            <Label>Campaign Name *</Label>
            <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="e.g. TX Wave 4" style={inputStyle} />

            <Label>Subtitle</Label>
            <input value={form.subtitle} onChange={e=>setForm(f=>({...f,subtitle:e.target.value}))} placeholder="e.g. Prize Refresh" style={inputStyle} />

            <Label>Notes <span style={{fontWeight:400,color:"#aaa"}}>(shown in small text on card)</span></Label>
            <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}
              placeholder="e.g. Confirm assets by Monday" rows={2}
              style={{...inputStyle, resize:"vertical", minHeight:52}} />

            <Label>Campaign Type</Label>
            <select value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))} style={{...inputStyle, cursor:"pointer"}}>
              {CAMPAIGN_TYPES.filter(t=>t!=="All Types").map(t=><option key={t}>{t}</option>)}
            </select>

            <Label>Day of Month *</Label>
            <input type="number" min={1} max={daysInMonth} value={form.day}
              onChange={e=>setForm(f=>({...f,day:e.target.value}))} placeholder={`1–${daysInMonth}`}
              style={{...inputStyle, width:80}} />

            <Label>Color</Label>
            <div style={{ display:"flex", flexWrap:"wrap", gap:8, marginBottom:16 }}>
              {CAMPAIGN_COLORS.map(c=>(
                <button key={c.value} onClick={()=>pickColor(c.value)} title={c.label} style={{
                  width:32, height:32, borderRadius:7, background:c.value, border:"none", cursor:"pointer",
                  outline:form.color===c.value?"3px solid #3B5BDB":"2px solid transparent", outlineOffset:2
                }} />
              ))}
            </div>

            {/* Preview */}
            <Label>Preview</Label>
            <div style={{ marginBottom:18 }}>
              <div style={{ display:"inline-block", background:form.color, color:form.textColor, borderRadius:6, padding:"6px 12px", fontSize:12, fontWeight:700 }}>
                {form.name||"Campaign Name"}
                {form.subtitle && <div style={{fontWeight:400,fontSize:11}}>{form.subtitle}</div>}
                {form.notes    && <div style={{fontWeight:400,fontSize:9,opacity:0.75,marginTop:2,fontStyle:"italic"}}>📝 {form.notes}</div>}
              </div>
            </div>

            <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
              {editId && <button onClick={handleDelete} style={{...btnStyle,background:"#fee2e2",color:"#c0392b"}}>Delete</button>}
              <button onClick={()=>setModal(false)} style={{...btnStyle,background:"#f0f1f5",color:"#4a5068"}}>Cancel</button>
              <button onClick={handleSave} style={{...btnStyle,background:"#3B5BDB",color:"#fff"}}>{editId?"Save Changes":"Add to Calendar"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Sub-components ────────────────────────────────────────────────────────────
const Label = ({children}) => <label style={{display:"block",fontSize:12,fontWeight:600,color:"#4a5068",marginBottom:5}}>{children}</label>;

// ── Shared styles ─────────────────────────────────────────────────────────────
const navBtn = { background:"#fff", border:"1px solid #e0e3ec", borderRadius:8, width:36, height:36, fontSize:18, cursor:"pointer", color:"#4a5068", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700 };
const btnStyle = { border:"none", borderRadius:7, padding:"9px 16px", fontSize:13, fontWeight:600, cursor:"pointer" };
const selectStyle = { border:"1px solid #e0e3ec", borderRadius:7, padding:"8px 12px", fontSize:13, color:"#1a1d2e", cursor:"pointer", background:"#fff", fontFamily:"inherit" };
const inputStyle = { display:"block", width:"100%", boxSizing:"border-box", border:"1px solid #e0e3ec", borderRadius:7, padding:"9px 12px", fontSize:14, color:"#1a1d2e", marginBottom:14, outline:"none", fontFamily:"inherit" };
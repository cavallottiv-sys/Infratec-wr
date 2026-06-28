import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import { db } from "./firebase.js";
import { ref, set, get, onValue } from "firebase/database";

pdfjsLib.GlobalWorkerOptions.workerSrc = `${import.meta.env.BASE_URL}pdf.worker.min.mjs`;

const UTENTI = [
  { id: "admin", nome: "Admin", pin: "1234", ruolo: "admin" },
  { id: "op1", nome: "Operatore 1", pin: "5678", ruolo: "operatore" },
  { id: "op2", nome: "Operatore 2", pin: "9012", ruolo: "operatore" },
];

const STATUS = {
  da_fare: { label: "Da fare", bg: "#E8EDF5", color: "#1C2B4A", dot: "#6B7FA3" },
  in_corso: { label: "In corso", bg: "#E0F4F7", color: "#004D5C", dot: "#00A8C8" },
  completato: { label: "Completato", bg: "#DFF2E4", color: "#1A5C2A", dot: "#28A745" },
};

const STATUS_ORDER = ["da_fare", "in_corso", "completato"];
const CATEGORIE = ["Opere boschive", "Opere accessorie"];
const DB_PATH = "infratec/data";

const LS = {
  display: "block",
  fontSize: 11,
  fontWeight: 700,
  color: "#6B7FA3",
  textTransform: "uppercase",
  letterSpacing: 0.8,
  marginBottom: 5,
};

const IS = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: 9,
  border: "1.5px solid #D4D8DE",
  fontSize: 13,
  fontFamily: "inherit",
  outline: "none",
  boxSizing: "border-box",
  background: "#fff",
};

const BS = {
  background: "none",
  border: "none",
  color: "#003087",
  fontWeight: 600,
  fontSize: 13,
  cursor: "pointer",
  padding: 0,
  display: "flex",
  alignItems: "center",
  gap: 4,
};

async function loadData() {
  try {
    const snap = await get(ref(db, DB_PATH));
    if (snap.exists()) {
      const val = snap.val();
      return {
        wrs: Array.isArray(val.wrs) ? val.wrs : [],
        prezzi: val.prezzi && typeof val.prezzi === "object" ? val.prezzi : {},
      };
    }
    return { wrs: [], prezzi: {} };
  } catch (e) {
    console.error("loadData error", e);
    return { wrs: [], prezzi: {} };
  }
}

async function saveData(d) {
  try {
    const clean = {
      ...d,
      wrs: (d.wrs || []).map((w) => ({
        ...w,
        photos: (w.photos || []).map((p) => ({
          id: p.id ?? "",
          dataUrl: p.dataUrl ?? "",
          ts: p.ts ?? 0,
          nome_operatore: p.nome_operatore ?? "",
        })),
      })),
    };
    await set(ref(db, DB_PATH), clean);
  } catch (e) {
    console.error("saveData error", e);
  }
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2);
}

function fileToBase64(f) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(f);
  });
}

function fmtDate(iso) {
  return iso ? new Date(iso).toLocaleDateString("it-IT") : "—";
}

function fmtDateTime(iso) {
  return iso
    ? new Date(iso).toLocaleString("it-IT", {
        day: "2-digit",
        month: "2-digit",
        year: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "—";
}

function nomeUtente(id) {
  return UTENTI.find((u) => u.id === id)?.nome || id || "—";
}

function fmtTempoObiettivo(val) {
  if (!val) return "—";
  const d = new Date(val);
  if (!isNaN(d.getTime())) {
    return d.toLocaleDateString("it-IT", { day: "2-digit", month: "2-digit", year: "2-digit" });
  }
  return val;
}

// ── Helpers parser WR ────────────────────────────────────────────────────────

function normalizeText(text) {
  if (!text) return "";
  return String(text)
    .replace(/\r\n/g, "\n").replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
}

function cleanValue(val, maxLen = 120) {
  if (!val) return "";
  const c = String(val).replace(/^[\s:;\-=|,]+/, "").replace(/[\s]+$/, "").replace(/\n.*$/s, "").trim();
  if (c.length < 2 || c.length > maxLen) return "";
  return c;
}

// Stop-word: tronca il valore estratto quando incontra un altro label noto
const FIELD_STOPS = /\b(Comune|Appuntamento|Telefono Reclamante|Centrale|Job Type|Sq\.|Tecnico|Data Dispaccio|NOME_ASSISTENTE|OPERATORE_SIM|RICHIESTA_PERMESSO|TELEFONO_ASSISTENTE|TIPO_INTERVENTO_SVOA|URGENZA|LATITUDINE|LONGITUDINE|UNIQUEID|COGNOME|CELL\.|TELEF\.)/i;

function extractByLabel(text, label) {
  const esc = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const m = text.match(new RegExp(esc + "\\s*[-=|]?\\s*([^\\n]{2,120})", "i"));
  if (!m) return null;
  const stopMatch = m[1].match(FIELD_STOPS);
  return stopMatch ? m[1].slice(0, stopMatch.index).trim() : m[1];
}

function pickFirstMatch(text, labelList) {
  for (const label of labelList) {
    const val = cleanValue(extractByLabel(text, label));
    if (val) return val;
  }
  return "";
}

const CATEGORIA_MAP = [
  { pattern: /ACCESSORI[EA]/i, label: "Opere accessorie" },
  { pattern: /BOSCHIV[AEO]/i,  label: "Opere boschive" },
];

function resolveCategoria(raw) {
  if (!raw) return "";
  for (const { pattern, label } of CATEGORIA_MAP) {
    if (pattern.test(raw)) return label;
  }
  return cleanValue(raw, 60);
}

const PROVINCIA_MAP = [
  { pattern: /\bpalermo\b/i,   value: "Palermo" },
  { pattern: /\bagrigento\b/i, value: "Agrigento" },
  { pattern: /\btrapani\b/i,   value: "Trapani" },
  { pattern: /\b(monreale|bagheria|misilmeri|termini\s*imerese|belmonte\s*mezzagno|partinico|carini|cinisi|capaci|isola\s*delle\s*femmine|torretta|giardinello|borgetto|terrasini|trappeto|balestrate|castellammare\s*del\s*golfo|camporeale|corleone|lercara|prizzi|marineo|baucina|ciminna|mezzojuso|ventimiglia|bolognetta|piana\s*degli\s*albanesi|san\s*giuseppe\s*jato|san\s*cipirrello|bisacquino|giuliana|chiusa\s*sclafani|contessa\s*entellina|sambuca\s*di\s*sicilia|sciacca|menfi)\b/i, value: "Palermo" },
  { pattern: /\b(favara|canicatt[iì]|licata|racalmuto|palma\s*di\s*montechiaro|naro|porto\s*empedocle|realmonte|siculiana|montallegro|ribera|caltabellotta|burgio|lucca\s*sicula|bivona|cammarata|san\s*giovanni\s*gemini|casteltermini)\b/i, value: "Agrigento" },
  { pattern: /\b(marsala|mazara\s*del\s*vallo|castelvetrano|alcamo|salemi|vita|paceco|buseto\s*palizzolo|calatafimi|campobello\s*di\s*mazara|gibellina|partanna|santa\s*ninfa|salaparuta|poggioreale|erice|valderice|custonaci)\b/i, value: "Trapani" },
];

function deduceProvincia(str) {
  if (!str) return "";
  for (const { pattern, value } of PROVINCIA_MAP) {
    if (pattern.test(str)) return value;
  }
  return "";
}

// ── Parser principale ─────────────────────────────────────────────────────────

function parseWRText(text) {
  const t = normalizeText(text);

  // ── Numero WR ──
  // Label esatti nel PDF: "WR:", "NUMBER", "WORKREQUESTID"
  // Prende solo la parte numerica pura
  const numero_wr_raw = pickFirstMatch(t, ["WORKREQUESTID", "NUMBER", "WR:"]);
  const numero_wr = (numero_wr_raw.match(/\d+/) || [""])[0];

  // ── Tipo intervento: "Accessorie" o "Boschive" ──
  // Nel PDF: "CODICE CL./O.L.: 0001 - TI26-OP.ACCESSORIE"
  const codiceOL = pickFirstMatch(t, ["CODICE CL./O.L."]);
  const tipo_intervento = /ACCESSOR/i.test(codiceOL)
    ? "Accessorie"
    : /BOSCHIV/i.test(codiceOL)
    ? "Boschive"
    : /ACCESSOR/i.test(t)
    ? "Accessorie"
    : /BOSCHIV/i.test(t)
    ? "Boschive"
    : "";

  // ── Categoria: derivata dal tipo ──
  const categoria = tipo_intervento === "Accessorie"
    ? "Opere accessorie"
    : "Opere boschive";

  // ── Indirizzo ──
  // "INDIRIZZO_INTERV" è su riga propria → valore pulito
  // "Indiriz.:" in testa è seguito da " Comune: ..." sulla stessa riga → tronco lì
  const indirizzoRaw = pickFirstMatch(t, ["INDIRIZZO_INTERV", "STREET", "Indiriz.:"]);
  const indirizzoClean = indirizzoRaw
    .replace(/\s*Comune:.*$/i, "")
    .replace(/\s{2,}/g, " ")
    .trim();

  // ── Comune: solo nome città, taglia codice centrale e resto ──
  const comuneRaw = pickFirstMatch(t, ["CITTA", "CITY", "Comune"]);
  // "CITTA - PALERMO" → "PALERMO", taglia tutto dopo spazio+codice
  const comune = comuneRaw.replace(/\s+[A-Z0-9_]{4,}.*$/, "").trim();

  // ── Indirizzo = Comune + via ──
  const indirizzo = comune && indirizzoClean
    ? comune.charAt(0) + comune.slice(1).toLowerCase() + " " + indirizzoClean
    : indirizzoClean || comune;

  // ── Provincia: dedotta dal comune ──
  const provincia = deduceProvincia(comune) || deduceProvincia(indirizzo);

  // ── Descrizione ──
  const oggettoMatch = t.match(/_Oggetto:\s*([^_\n]+?)(?:\s*_\s*_|\s*_Referente:|\s*_Segnalante:|$)/i);
  const descrizione = oggettoMatch
    ? oggettoMatch[1].replace(/\s+/g, " ").trim()
    : "";

  // ── Referente: solo nome — "NOME_ASSISTENTE" è su riga propria ──
  const referenteRaw = pickFirstMatch(t, ["NOME_ASSISTENTE"]);
  // Taglia tutto dopo il primo separatore noto (es. spazio + label successivo)
  const referente = referenteRaw
    .replace(/\s*(OPERATORE|RICHIESTA|TELEFONO|TIPO_INT|URGENZA).*/i, "")
    .trim();

  // ── Telefono: solo cifre ──
  const telefonoRaw = pickFirstMatch(t, ["TELEFONO_ASSISTENTE"]);
  const telefonoClean = telefonoRaw.replace(/[^\d+]/g, "");
  const telefono = /\d{6,}/.test(telefonoClean) ? telefonoClean : "";

  // ── Urgenza ──
  const urgenzaRaw = pickFirstMatch(t, ["URGENZA", "NETWORK_URGENTE"]);

  // ── Latitudine: estrae solo il pattern DMS (dd-dd-dd.ddd) o decimale ──
  const latitudineRaw = pickFirstMatch(t, ["LATITUDINE_INTERVENTO"]);
  const latMatch = (latitudineRaw || "").match(/(-?\d{1,3}-\d{2}-[\d.]+|-?\d{1,3}[.,]\d{3,})/);
  const latitudine_intervento = latMatch ? latMatch[1] : "";

  // ── Longitudine: stessa logica ──
  const longitudineRaw = pickFirstMatch(t, ["LONGITUDINE_INTERVENTO"]);
  const lonMatch = (longitudineRaw || "").match(/(-?\d{1,3}-\d{2}-[\d.]+|-?\d{1,3}[.,]\d{3,})/);
  const longitudine_intervento = lonMatch ? lonMatch[1] : "";

  // ── Tempo obiettivo: estrae solo la parte ISO o data ──
  const tempoRaw = pickFirstMatch(t, ["DATA_PREV_FINE_LAV", "FINE_LAVORI_PREV", "DUEDATE"]);
  const tempoMatch = (tempoRaw || "").match(/(\d{4}-\d{2}-\d{2}T[\d:+.]+|\d{2}\s+\w{3}\s+\d{4}[\s\d:]*)/);
  const tempo_obiettivo = tempoMatch ? tempoMatch[1].trim() : "";

  return {
    numero_wr, tipo_intervento, categoria,
    indirizzo, comune, provincia, descrizione,
    referente, telefono,
    urgenza: urgenzaRaw === "1" ? "Sì" : urgenzaRaw === "0" ? "No" : "",
    latitudine_intervento, longitudine_intervento, tempo_obiettivo,
  };
}

async function extractWRFromEmail(text) {
  const dati = parseWRText(text);
  return {
    numero_wr:              dati.numero_wr || "",
    tipo_intervento:        dati.tipo_intervento || "Pozzetto",
    categoria:              dati.categoria || "Opere boschive",
    indirizzo:              dati.indirizzo || "",
    comune:                 dati.comune || "",
    provincia:              dati.provincia || "",
    note:                   dati.descrizione || "",
    referente:              dati.referente || "",
    telefono:               dati.telefono || "",
    priorita:               /^s/i.test(dati.urgenza || "") ? "urgente" : "normale",
    latitudine_intervento:  dati.latitudine_intervento || "",
    longitudine_intervento: dati.longitudine_intervento || "",
    tempo_obiettivo:        dati.tempo_obiettivo || "",
  };
}

async function fetchGmailWRs() {
  return [];
}

function dmsToDecimal(dms) {
  const m = String(dms || "").trim().match(/^(-?\d+)-(\d+)-([\d.]+)$/);
  if (!m) return "";
  const deg = Number(m[1]);
  const min = Number(m[2]);
  const sec = Number(m[3]);
  const sign = deg < 0 ? -1 : 1;
  const absDeg = Math.abs(deg);
  return (sign * (absDeg + min / 60 + sec / 3600)).toFixed(6);
}

function openNavigation(wr, operatorPosition) {
  const lat = dmsToDecimal(wr.latitudine_intervento || wr.lat);
  const lng = dmsToDecimal(wr.longitudine_intervento || wr.lng);

  let destination = "";
  if (lat && lng) {
    destination = `${lat},${lng}`;
  } else {
    // Fallback: indirizzo testuale (comune + via)
    const addr = [wr.indirizzo, wr.comune].filter(Boolean).join(", ");
    if (!addr) return;
    destination = encodeURIComponent(addr);
  }

  let url = `https://www.google.com/maps/dir/?api=1&destination=${destination}&travelmode=driving`;
  if (operatorPosition?.lat && operatorPosition?.lng) {
    url += `&origin=${operatorPosition.lat},${operatorPosition.lng}`;
  }
  window.open(url, "_blank", "noopener,noreferrer");
}

function callReferente(telefono) {
  const cleanPhone = String(telefono || "").replace(/[^\d+]/g, "");
  if (!cleanPhone) return;
  window.location.href = `tel:${cleanPhone}`;
}

async function extractTextFromPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const uint8Array = new Uint8Array(arrayBuffer);
  const loadingTask = pdfjsLib.getDocument({ data: uint8Array });
  const pdf = await loadingTask.promise;
  let fullText = "";

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ");
    fullText += `\n${pageText}`;
  }

  return fullText.trim();
}

function exportExcel(wrs, prezzi) {
  const rows = wrs.map((w) => ({
    "N° WR": w.numero_wr || "",
    Data: fmtDate(w.created_at),
    Categoria: w.categoria || "",
    Tipo: w.tipo_intervento || "",
    Indirizzo: w.indirizzo || "",
    "Note WR": w.note || "",
    "Lavoro eseguito": w.lavoro_eseguito || "",
    Stato: STATUS[w.status]?.label || "",
    Operatore: nomeUtente(w.ultimo_operatore),
    "Ultimo agg.": fmtDateTime(w.updated_at),
    "Prezzo (€)": prezzi[w.id] != null && prezzi[w.id] !== "" ? Number(prezzi[w.id]).toFixed(2) : "",
    Foto: (w.photos || []).length,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [12, 10, 18, 14, 30, 30, 30, 12, 14, 16, 12, 5].map((w) => ({ wch: w }));

  let r = rows.length + 3;
  CATEGORIE.forEach((cat) => {
    const tot = wrs
      .filter((w) => w.categoria === cat)
      .reduce((s, w) => {
        const p = prezzi[w.id];
        return p != null && p !== "" ? s + Number(p) : s;
      }, 0);

    XLSX.utils.sheet_add_aoa(
      ws,
      [[`Totale ${cat}`, "", "", "", "", "", "", "", "", "", tot.toFixed(2), ""]],
      { origin: r++ }
    );
  });

  const grand = wrs.reduce((s, w) => {
    const p = prezzi[w.id];
    return p != null && p !== "" ? s + Number(p) : s;
  }, 0);

  XLSX.utils.sheet_add_aoa(
    ws,
    [["TOTALE GENERALE", "", "", "", "", "", "", "", "", "", grand.toFixed(2), ""]],
    { origin: r }
  );

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "WR Infratec");
  XLSX.writeFile(wb, `Infratec_WR_${new Date().toLocaleDateString("it-IT").replace(/\//g, "-")}.xlsx`);
}

function Badge({ status }) {
  const s = STATUS[status] || STATUS.da_fare;
  return (
    <span
      style={{
        padding: "4px 8px",
        borderRadius: 999,
        background: s.bg,
        color: s.color,
        fontSize: 11,
        fontWeight: 700,
      }}
    >
      {s.label}
    </span>
  );
}

function CatBadge({ cat }) {
  const b = cat === "Opere boschive";
  return (
    <span style={{ fontSize: 11, fontWeight: 700, color: b ? "#2E7D32" : "#6B7FA3" }}>
      {b ? "🌿 Boschive" : "🔧 Accessorie"}
    </span>
  );
}

function Lightbox({ photo, onClose }) {
  if (!photo) return null;
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,.75)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 9999,
        padding: 20,
      }}
    >
      <button
        onClick={onClose}
        style={{
          position: "absolute",
          top: 20,
          right: 20,
          background: "#fff",
          border: "none",
          borderRadius: "50%",
          width: 36,
          height: 36,
          cursor: "pointer",
        }}
      >
        ×
      </button>
      <img src={photo.dataUrl} alt="Foto WR" style={{ maxWidth: "95vw", maxHeight: "90vh", borderRadius: 12 }} />
    </div>
  );
}

function PhotoSection({ photos = [], onAdd, onDelete, onPreview }) {
  const inputRef = useRef(null);

  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 8,
        }}
      >
        <div style={{ fontWeight: 700, fontSize: 13 }}>Foto ({photos.length})</div>
        <button
          onClick={() => inputRef.current?.click()}
          style={{
            padding: "6px 14px",
            borderRadius: 8,
            background: "#003087",
            color: "#fff",
            fontWeight: 700,
            fontSize: 12,
            border: "none",
            cursor: "pointer",
          }}
        >
          📷 Aggiungi
        </button>
        <input
          ref={inputRef}
          type="file"
          accept="image/*"
          capture="environment"
          multiple
          style={{ display: "none" }}
          onChange={async (e) => {
            for (const f of Array.from(e.target.files || [])) {
              await onAdd(f);
            }
            e.target.value = "";
          }}
        />
      </div>

      {photos.length === 0 ? (
        <div
          onClick={() => inputRef.current?.click()}
          style={{
            border: "2px dashed #D4D8DE",
            borderRadius: 12,
            padding: 20,
            textAlign: "center",
            color: "#9BA5B4",
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          📷 Tocca per scattare o allegare
        </div>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
          {photos.map((p) => (
            <div
              key={p.id}
              style={{
                position: "relative",
                aspectRatio: "1",
                borderRadius: 10,
                overflow: "hidden",
                background: "#F5F7FA",
              }}
            >
              <img
                src={p.dataUrl}
                alt="Foto WR"
                onClick={() => onPreview(p)}
                style={{
                  width: "100%",
                  height: "100%",
                  objectFit: "cover",
                  cursor: "pointer",
                  display: "block",
                }}
              />
              <button
                onClick={() => onDelete(p.id)}
                style={{
                  position: "absolute",
                  top: 4,
                  right: 4,
                  background: "rgba(0,0,0,.55)",
                  border: "none",
                  borderRadius: "50%",
                  width: 22,
                  height: 22,
                  color: "#fff",
                  fontSize: 13,
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
              >
                ×
              </button>
            </div>
          ))}

          <div
            onClick={() => inputRef.current?.click()}
            style={{
              borderRadius: 10,
              border: "2px dashed #D4D8DE",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              aspectRatio: "1",
              cursor: "pointer",
              color: "#9BA5B4",
              fontSize: 26,
            }}
          >
            +
          </div>
        </div>
      )}
    </div>
  );
}

function Row({ label, value, highlight }) {
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "space-between",
        gap: 12,
        padding: "10px 0",
        borderBottom: "1px solid #EEF2F7",
      }}
    >
      <div style={{ fontSize: 12, color: "#6B7FA3", fontWeight: 700 }}>{label}</div>
      <div style={{ fontSize: 13, fontWeight: highlight ? 700 : 500, textAlign: "right" }}>{value || "—"}</div>
    </div>
  );
}

function Field({ label, value, onChange, placeholder, type = "text" }) {
  return (
    <div>
      <label style={LS}>{label}</label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        style={IS}
      />
    </div>
  );
}

function TBtn({ label, onClick, active, accent }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: "8px 12px",
        borderRadius: 10,
        border: "none",
        cursor: "pointer",
        background: active ? (accent ? "#F05A22" : "#003087") : "rgba(255,255,255,.12)",
        color: "#fff",
        fontWeight: 700,
        fontSize: 12,
      }}
    >
      {label}
    </button>
  );
}

function LoginScreen({ onLogin }) {
  const [nome, setNome] = useState("");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");

  function handleSubmit() {
    const u = UTENTI.find((x) => x.nome === nome && x.pin === pin);
    if (u) onLogin(u);
    else {
      setErr("Nome o PIN errato");
      setPin("");
    }
  }

  return (
    <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: "#F3F6FB", padding: 20 }}>
      <div
        style={{
          width: "100%",
          maxWidth: 380,
          background: "#fff",
          borderRadius: 16,
          padding: 20,
          boxShadow: "0 8px 30px rgba(0,0,0,.08)",
        }}
      >
        <div style={{ fontSize: 24, fontWeight: 800, color: "#003087" }}>INFRATEC</div>
        <div style={{ fontSize: 13, color: "#6B7FA3", marginBottom: 16 }}>WR Manager · TIM/Fibercop</div>

        <label style={LS}>Chi sei?</label>
        <select
          value={nome}
          onChange={(e) => {
            setNome(e.target.value);
            setErr("");
          }}
          style={IS}
        >
          <option value="">— Seleziona —</option>
          {UTENTI.map((u) => (
            <option key={u.id} value={u.nome}>
              {u.nome}
            </option>
          ))}
        </select>

        <div style={{ height: 12 }} />

        <label style={LS}>PIN</label>
        <input
          value={pin}
          type="password"
          onChange={(e) => {
            setPin(e.target.value);
            setErr("");
          }}
          placeholder="••••"
          style={{ ...IS, borderColor: err ? "#C0392B" : "#D4D8DE" }}
          onKeyDown={(e) => e.key === "Enter" && handleSubmit()}
        />

        {err && <div style={{ marginTop: 10, color: "#C0392B", fontSize: 12 }}>⚠ {err}</div>}

        <button
          onClick={handleSubmit}
          style={{
            marginTop: 16,
            width: "100%",
            padding: 12,
            borderRadius: 10,
            background: "#003087",
            color: "#fff",
            fontWeight: 700,
            border: "none",
            cursor: "pointer",
          }}
        >
          Accedi
        </button>
      </div>
    </div>
  );
}


// ── Componente distanza stradale ─────────────────────────────────────────────
function DistanzaBox({ wr }) {
  const [stato, setStato] = useState("idle"); // idle | loading | ok | error
  const [km, setKm] = useState(null);
  const [durata, setDurata] = useState(null);

  const calcola = () => {
    if (!navigator.geolocation) {
      setStato("error");
      return;
    }
    setStato("loading");
    navigator.geolocation.getCurrentPosition(
      async (pos) => {
        const origin = `${pos.coords.latitude},${pos.coords.longitude}`;
        const lat = dmsToDecimal(wr.latitudine_intervento || wr.lat);
        const lng = dmsToDecimal(wr.longitudine_intervento || wr.lng);
        const dest = lat && lng
          ? `${lat},${lng}`
          : encodeURIComponent([wr.indirizzo, wr.comune].filter(Boolean).join(", "));

        if (!dest) { setStato("error"); return; }

        // Usa Google Maps Distance Matrix via URL embed (no API key)
        // Apriamo Maps con indicazioni — calcoliamo con Haversine come stima
        const R = 6371;
        const dLat = (Number(lat) - pos.coords.latitude) * Math.PI / 180;
        const dLon = (Number(lng) - pos.coords.longitude) * Math.PI / 180;
        const a =
          Math.sin(dLat / 2) ** 2 +
          Math.cos(pos.coords.latitude * Math.PI / 180) *
          Math.cos(Number(lat) * Math.PI / 180) *
          Math.sin(dLon / 2) ** 2;
        const distLinea = R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
        // Stima stradale ~1.3x distanza in linea d'aria
        const distStrada = distLinea * 1.3;
        const minStimati = Math.round(distStrada / 40 * 60); // ~40 km/h media urbana

        setKm(distStrada.toFixed(1));
        setDurata(minStimati < 60
          ? `~${minStimati} min`
          : `~${Math.floor(minStimati/60)}h ${minStimati%60}min`);
        setStato("ok");
      },
      () => setStato("error"),
      { timeout: 8000 }
    );
  };

  const mapsUrl = () => {
    const lat = dmsToDecimal(wr.latitudine_intervento || wr.lat);
    const lng = dmsToDecimal(wr.longitudine_intervento || wr.lng);
    const dest = lat && lng
      ? `${lat},${lng}`
      : encodeURIComponent([wr.indirizzo, wr.comune].filter(Boolean).join(", "));
    return `https://www.google.com/maps/dir/?api=1&destination=${dest}&travelmode=driving`;
  };

  return (
    <div style={{
      background: "linear-gradient(135deg, #003087 0%, #0055CC 100%)",
      borderRadius: 14,
      padding: "14px 16px",
      color: "#fff",
    }}>
      <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: 0.8, textTransform: "uppercase", opacity: 0.75, marginBottom: 10 }}>
        📍 Distanza dal cantiere
      </div>

      {stato === "idle" && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button
            onClick={calcola}
            style={{
              flex: 1,
              padding: "11px 0",
              borderRadius: 10,
              border: "none",
              background: "#fff",
              color: "#003087",
              fontWeight: 800,
              fontSize: 14,
              cursor: "pointer",
            }}
          >
            📐 Calcola distanza
          </button>
          <a
            href={mapsUrl()}
            target="_blank"
            rel="noopener noreferrer"
            style={{
              flex: 1,
              padding: "11px 0",
              borderRadius: 10,
              border: "2px solid rgba(255,255,255,0.4)",
              background: "transparent",
              color: "#fff",
              fontWeight: 700,
              fontSize: 14,
              cursor: "pointer",
              textDecoration: "none",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 6,
            }}
          >
            🧭 Apri Maps
          </a>
        </div>
      )}

      {stato === "loading" && (
        <div style={{ textAlign: "center", fontSize: 13, opacity: 0.85, padding: "8px 0" }}>
          Rilevamento posizione…
        </div>
      )}

      {stato === "ok" && (
        <div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginBottom: 6 }}>
            <span style={{ fontSize: 32, fontWeight: 900, letterSpacing: -1 }}>{km}</span>
            <span style={{ fontSize: 16, fontWeight: 700, opacity: 0.85 }}>km</span>
            <span style={{ fontSize: 14, opacity: 0.75, marginLeft: 4 }}>{durata}</span>
          </div>
          <div style={{ fontSize: 11, opacity: 0.6, marginBottom: 10 }}>stima stradale (×1.3 linea d'aria)</div>
          <div style={{ display: "flex", gap: 8 }}>
            <a
              href={mapsUrl()}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                flex: 1,
                padding: "9px 0",
                borderRadius: 9,
                border: "2px solid rgba(255,255,255,0.4)",
                background: "transparent",
                color: "#fff",
                fontWeight: 700,
                fontSize: 13,
                cursor: "pointer",
                textDecoration: "none",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              🧭 Apri Maps
            </a>
            <button
              onClick={() => { setStato("idle"); setKm(null); setDurata(null); }}
              style={{ padding: "9px 14px", borderRadius: 9, border: "2px solid rgba(255,255,255,0.3)", background: "transparent", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" }}
            >
              ↺
            </button>
          </div>
        </div>
      )}

      {stato === "error" && (
        <div>
          <div style={{ fontSize: 13, opacity: 0.85, marginBottom: 10 }}>
            ⚠️ Posizione GPS non disponibile su questo dispositivo.
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <a
              href={mapsUrl()}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                flex: 1,
                padding: "9px 0",
                borderRadius: 9,
                border: "none",
                background: "#fff",
                color: "#003087",
                fontWeight: 800,
                fontSize: 13,
                cursor: "pointer",
                textDecoration: "none",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              🧭 Apri Maps comunque
            </a>
            <button
              onClick={() => setStato("idle")}
              style={{ padding: "9px 14px", borderRadius: 9, border: "2px solid rgba(255,255,255,0.3)", background: "transparent", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" }}
            >
              ↺
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default function WRManager() {
  const [utente, setUtente] = useState(null);
  const [data, setData] = useState({ wrs: [], prezzi: {} });
  const [view, setView] = useState("lista");
  const [selectedId, setSelectedId] = useState(null);
  const [filter, setFilter] = useState("tutti");
  const [catFilter, setCatFilter] = useState("tutte");
  const [lightbox, setLightbox] = useState(null);
  const [toast, setToast] = useState("");
  const [editLavoro, setEditLavoro] = useState("");
  const [priceDraft, setPriceDraft] = useState("");
  const [dbStatus, setDbStatus] = useState("connecting");

  const [emailText, setEmailText] = useState("");
  const [extracting, setExtracting] = useState(false);
  const [extracted, setExtracted] = useState(null);
  const [extractErr, setExtractErr] = useState("");
  const [gmailLoading, setGmailLoading] = useState(false);
  const [gmailEmails, setGmailEmails] = useState([]);
  const [gmailErr, setGmailErr] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [pdfName, setPdfName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [provFilter, setProvFilter] = useState("tutte");
  const [editingId, setEditingId] = useState(null);

  const emptyForm = {
    numero_wr: "",
    tipo_intervento: "Pozzetto",
    categoria: "Opere boschive",
    indirizzo: "",
    comune: "",
    provincia: "",
    note: "",
    priorita: "normale",
    referente: "",
    telefono: "",
    latitudine_intervento: "",
    longitudine_intervento: "",
    tempo_obiettivo: "",
  };

  const [form, setForm] = useState(emptyForm);
  const [saving, setSaving] = useState(false);

  const selected = data.wrs.find((w) => w.id === selectedId) || null;
  const isAdmin = utente?.ruolo === "admin";

  useEffect(() => {
    loadData()
      .then((d) => {
        setData(d);
        setDbStatus("ok");
      })
      .catch(() => setDbStatus("error"));

    const unsubscribe = onValue(
      ref(db, DB_PATH),
      (snap) => {
        if (snap.exists()) {
          const val = snap.val();
          setData({
            wrs: Array.isArray(val.wrs) ? val.wrs : [],
            prezzi: val.prezzi && typeof val.prezzi === "object" ? val.prezzi : {},
          });
          setDbStatus("ok");
        } else {
          setData({ wrs: [], prezzi: {} });
        }
      },
      () => setDbStatus("error")
    );

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    setEditLavoro(selected?.lavoro_eseguito || "");
    setPriceDraft(selected ? String(data.prezzi[selected.id] ?? "") : "");
  }, [selected, data.prezzi]);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2800);
  };

  const persist = useCallback(async (next) => {
    setData(next);
    await saveData(next);
  }, []);

  const persistAndSync = useCallback(
    async (next) => {
      await persist(next);
    },
    [persist]
  );

  function logOp() {
    return { updated_at: new Date().toISOString(), ultimo_operatore: utente.id };
  }

  async function addPhoto(wrId, file) {
    const dataUrl = await fileToBase64(file);
    const photo = { id: uid(), dataUrl, ts: Date.now(), nome_operatore: utente.nome };
    const next = {
      ...data,
      wrs: data.wrs.map((w) =>
        w.id === wrId ? { ...w, photos: [...(w.photos || []), photo], ...logOp() } : w
      ),
    };
    await persistAndSync(next);
    showToast("Foto aggiunta ✓");
  }

  async function deletePhoto(wrId, photoId) {
    const next = {
      ...data,
      wrs: data.wrs.map((w) =>
        w.id === wrId ? { ...w, photos: (w.photos || []).filter((p) => p.id !== photoId), ...logOp() } : w
      ),
    };
    await persistAndSync(next);
    showToast("Foto eliminata");
  }

  async function fetchGmail() {
    setGmailLoading(true);
    setGmailErr("");
    setGmailEmails([]);
    try {
      const emails = await fetchGmailWRs();
      setGmailEmails(emails);
      if (!emails.length) setGmailErr("Nessuna mail WR trovata.");
    } catch (e) {
      setGmailErr(`Errore Gmail: ${e.message}`);
    } finally {
      setGmailLoading(false);
    }
  }

  async function importFromGmailEmail(email) {
    setExtracting(true);
    setExtractErr("");
    setExtracted(null);
    try {
      const d = await extractWRFromEmail(`Da: ${email.from}\nOggetto: ${email.subject}\nData: ${email.date}\n\n${email.snippet}`);
      setExtracted(d);
    } catch (e) {
      setExtractErr(`Errore: ${e.message}`);
    } finally {
      setExtracting(false);
    }
  }

  async function handleExtract() {
    if (!emailText.trim()) return;
    setExtracting(true);
    setExtractErr("");
    setExtracted(null);
    try {
      setExtracted(await extractWRFromEmail(emailText));
    } catch {
      setExtractErr("Errore estrazione.");
    } finally {
      setExtracting(false);
    }
  }

  async function handlePdfFile(file) {
    if (!file) return;
    if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
      setExtractErr("Seleziona un file PDF valido.");
      return;
    }

    setPdfLoading(true);
    setExtractErr("");
    setExtracted(null);
    setPdfName(file.name);
    setGmailEmails([]);

    try {
      const pdfText = await extractTextFromPdf(file);
      if (!pdfText.trim()) throw new Error("Il PDF non contiene testo leggibile.");
      setExtracted(await extractWRFromEmail(pdfText));
    } catch (e) {
      setExtractErr(`Errore lettura PDF: ${e.message}`);
    } finally {
      setPdfLoading(false);
    }
  }

  async function handleAdd() {
    if (!form.numero_wr.trim() || !form.indirizzo.trim()) return;
    setSaving(true);
    try {
      const wr = {
        id: uid(),
        numero_wr: form.numero_wr || "",
        tipo_intervento: form.tipo_intervento || "Pozzetto",
        categoria: form.categoria || "Opere accessorie",
        indirizzo: form.indirizzo || "",
        comune: form.comune || "",
        provincia: form.provincia || deduceProvincia(form.comune || "") || deduceProvincia(form.indirizzo || ""),
        note: form.note || "",
        priorita: form.priorita || "normale",
        referente: form.referente || "",
        telefono: form.telefono || "",
        latitudine_intervento: form.latitudine_intervento || "",
        longitudine_intervento: form.longitudine_intervento || "",
        tempo_obiettivo: form.tempo_obiettivo || "",
        status: "da_fare",
        photos: [],
        lavoro_eseguito: "",
        created_at: new Date().toISOString(),
        ...logOp(),
      };
      const next = { ...data, wrs: [wr, ...data.wrs] };
      await persist(next);
      setForm(emptyForm);
      setView("lista");
      showToast("WR registrata ✓");
    } finally {
      setSaving(false);
    }
  }

  function handleEditWR(wr) {
    setEditingId(wr.id);
    setForm({
      numero_wr: wr.numero_wr || "",
      tipo_intervento: wr.tipo_intervento || "Pozzetto",
      categoria: wr.categoria || "Opere boschive",
      indirizzo: wr.indirizzo || "",
      comune: wr.comune || "",
      provincia: wr.provincia || "",
      note: wr.note || "",
      priorita: wr.priorita || "normale",
      referente: wr.referente || "",
      telefono: wr.telefono || "",
      latitudine_intervento: wr.latitudine_intervento || "",
      longitudine_intervento: wr.longitudine_intervento || "",
      tempo_obiettivo: wr.tempo_obiettivo || "",
    });
    setView("modifica");
  }

  async function handleSaveEdit() {
    if (!form.numero_wr.trim() || !form.indirizzo.trim()) return;
    setSaving(true);
    try {
      const next = {
        ...data,
        wrs: data.wrs.map((w) =>
          w.id === editingId
            ? {
                ...w,
                numero_wr: form.numero_wr || "",
                tipo_intervento: form.tipo_intervento || "",
                categoria: form.categoria || "",
                indirizzo: form.indirizzo || "",
                comune: form.comune || "",
                provincia: form.provincia || deduceProvincia(form.comune || "") || deduceProvincia(form.indirizzo || ""),
                note: form.note || "",
                priorita: form.priorita || "normale",
                referente: form.referente || "",
                telefono: form.telefono || "",
                latitudine_intervento: form.latitudine_intervento || "",
                longitudine_intervento: form.longitudine_intervento || "",
                tempo_obiettivo: form.tempo_obiettivo || "",
                ...logOp(),
              }
            : w
        ),
      };
      await persist(next);
      setEditingId(null);
      setForm(emptyForm);
      setSelectedId(editingId);
      setView("dettaglio");
      showToast("WR modificata ✓");
    } finally {
      setSaving(false);
    }
  }

  async function confirmImport() {
    if (!extracted) return;

    const wr = {
      id: uid(),
      ...extracted,
      status: "da_fare",
      photos: [],
      lavoro_eseguito: "",
      created_at: new Date().toISOString(),
      ...logOp(),
    };

    await persist({ ...data, wrs: [wr, ...data.wrs] });
    setView("lista");
    setEmailText("");
    setExtracted(null);
    setGmailEmails([]);
    setPdfName("");
    showToast("WR importata ✓");
  }

  async function updateStatus(id, status) {
    const next = {
      ...data,
      wrs: data.wrs.map((w) => (w.id === id ? { ...w, status, ...logOp() } : w)),
    };
    await persistAndSync(next);
    showToast("Stato aggiornato ✓");
  }

  function cleanInlineText(v) {
    return String(v || "").replace(/\s+/g, " ").trim();
  }

  function shortText(v, max = 90) {
    const t = cleanInlineText(v);
    return t.length > max ? `${t.slice(0, max)}...` : t;
  }

  async function saveLavoro(id) {
    const next = {
      ...data,
      wrs: data.wrs.map((w) => (w.id === id ? { ...w, lavoro_eseguito: editLavoro, ...logOp() } : w)),
    };
    await persistAndSync(next);
    showToast("Descrizione salvata ✓");
  }

  async function savePrezzo(id) {
    const next = {
      ...data,
      prezzi: { ...data.prezzi, [id]: priceDraft === "" ? null : priceDraft },
    };
    await persist(next);
    showToast("Prezzo salvato ✓");
  }

  async function deleteWR(id) {
    const prezzi = { ...data.prezzi };
    delete prezzi[id];
    await persist({ wrs: data.wrs.filter((w) => w.id !== id), prezzi });
    setSelectedId(null);
    setView("lista");
    showToast("WR eliminata");
  }

  const wrs = data.wrs;
  const filtered = wrs
    .filter((w) => filter === "tutti" || w.status === filter)
    .filter((w) => catFilter === "tutte" || w.categoria === catFilter)
    .filter((w) => {
      if (provFilter === "tutte") return true;
      const prov = w.provincia || deduceProvincia(w.comune || "") || deduceProvincia(w.indirizzo || "");
      return prov === provFilter;
    });

  const counts = STATUS_ORDER.reduce((acc, s) => ({ ...acc, [s]: wrs.filter((w) => w.status === s).length }), {});

  const totaleVisibile = filtered.reduce((s, w) => {
    const p = data.prezzi[w.id];
    return p != null && p !== "" ? s + Number(p) : s;
  }, 0);

  if (!utente) {
    return <LoginScreen onLogin={(u) => { setUtente(u); setView("lista"); }} />;
  }

  return (
    <div style={{ minHeight: "100vh", background: "#F3F6FB", fontFamily: "Arial, sans-serif" }}>
      <div
        style={{
          background: "#003087",
          color: "#fff",
          padding: "14px 16px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
        }}
      >
        <div>
          <div style={{ fontSize: 22, fontWeight: 800 }}>INFRATEC</div>
          <div style={{ fontSize: 12, opacity: 0.9 }}>WR Manager</div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div style={{ fontSize: 12 }}>{utente.nome}</div>
          <div style={{ fontSize: 12 }}>{dbStatus === "ok" ? "🟢" : dbStatus === "error" ? "🔴" : "🟡"}</div>

          {isAdmin && (
            <>
              <TBtn
                label="+ WR"
                onClick={() => {
                  setView("aggiungi");
                  setSelectedId(null);
                  setForm(emptyForm);
                }}
                active={view === "aggiungi"}
              />
              <TBtn
                label="↓ Import"
                onClick={() => {
                  setView("importa");
                  setSelectedId(null);
                  setExtracted(null);
                  setGmailEmails([]);
                }}
                active={view === "importa"}
                accent
              />
              <TBtn label="Excel" onClick={() => exportExcel(wrs, data.prezzi)} />
            </>
          )}

          <button
            onClick={() => {
              setUtente(null);
              setSelectedId(null);
              setView("lista");
            }}
            style={{
              padding: "5px 10px",
              borderRadius: 8,
              fontSize: 11,
              fontWeight: 700,
              cursor: "pointer",
              border: "none",
              background: "rgba(255,255,255,.12)",
              color: "#fff",
            }}
          >
            Esci
          </button>
        </div>
      </div>

      {toast && (
        <div
          style={{
            position: "fixed",
            bottom: 20,
            left: "50%",
            transform: "translateX(-50%)",
            background: "#111",
            color: "#fff",
            padding: "10px 14px",
            borderRadius: 10,
            zIndex: 1000,
          }}
        >
          {toast}
        </div>
      )}

      <Lightbox photo={lightbox} onClose={() => setLightbox(null)} />

      <div style={{ maxWidth: 900, margin: "0 auto", padding: 16 }}>
        {view === "lista" && !selected && (
          <>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
              {STATUS_ORDER.map((s) => (
                <button
                  key={s}
                  onClick={() => setFilter(filter === s ? "tutti" : s)}
                  style={{
                    background: filter === s ? STATUS[s].bg : "#fff",
                    border: `1.5px solid ${filter === s ? STATUS[s].dot : "#E0E4EA"}`,
                    borderRadius: 10,
                    padding: "8px 10px",
                    cursor: "pointer",
                  }}
                >
                  <div style={{ fontWeight: 800 }}>{counts[s]}</div>
                  <div style={{ fontSize: 12 }}>{STATUS[s].label}</div>
                </button>
              ))}
            </div>

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
              {["tutte", "Opere boschive", "Opere accessorie"].map((c) => (
                <button
                  key={c}
                  onClick={() => setCatFilter(c)}
                  style={{
                    padding: "4px 11px", borderRadius: 16, border: "1.5px solid", fontSize: 11,
                    fontWeight: 600, cursor: "pointer",
                    borderColor: catFilter === c ? "#003087" : "#D4D8DE",
                    background: catFilter === c ? "#003087" : "#fff",
                    color: catFilter === c ? "#fff" : "#6B7FA3",
                  }}
                >
                  {c === "tutte" ? "Tutte" : c === "Opere boschive" ? "Boschive" : "Accessorie"}
                </button>
              ))}

              {["tutti", ...STATUS_ORDER].map((s) => (
                <button
                  key={s}
                  onClick={() => setFilter(s)}
                  style={{
                    padding: "4px 11px", borderRadius: 16, border: "1.5px solid", fontSize: 11,
                    fontWeight: 600, cursor: "pointer",
                    borderColor: filter === s ? "#555" : "#D4D8DE",
                    background: filter === s ? "#555" : "#fff",
                    color: filter === s ? "#fff" : "#6B7FA3",
                  }}
                >
                  {s === "tutti" ? "Tutti stati" : STATUS[s].label}
                </button>
              ))}

              <select
                value={provFilter}
                onChange={(e) => setProvFilter(e.target.value)}
                style={{
                  padding: "4px 10px", borderRadius: 16, border: "1.5px solid #D4D8DE",
                  fontSize: 11, fontWeight: 600, cursor: "pointer",
                  background: provFilter !== "tutte" ? "#E8EDF5" : "#fff",
                  color: provFilter !== "tutte" ? "#003087" : "#6B7FA3",
                  outline: "none",
                }}
              >
                <option value="tutte">Tutte province</option>
                <option value="Palermo">Palermo</option>
                <option value="Agrigento">Agrigento</option>
                <option value="Trapani">Trapani</option>
              </select>
            </div>

            {isAdmin && (
              <div
                style={{
                  background: "#fff",
                  borderRadius: 12,
                  padding: 12,
                  marginBottom: 12,
                  border: "1px solid #E8ECF2",
                }}
              >
                <b>{filtered.length}</b> WR selezionate · € <b>{totaleVisibile.toFixed(2)}</b>
              </div>
            )}

            {filtered.length === 0 ? (
              <div style={{ background: "#fff", borderRadius: 12, padding: 16, color: "#6B7FA3" }}>
                {isAdmin ? "Nessuna WR. Usa + WR o Import." : "Nessun lavoro assegnato."}
              </div>
            ) : (
              <div style={{ display: "grid", gap: 10 }}>
                {filtered.map((wr) => (
                  <div
                    key={wr.id}
                    onClick={() => {
                      setSelectedId(wr.id);
                      setView("dettaglio");
                    }}
                    style={{
                      background: "#fff",
                      borderRadius: 14,
                      padding: 14,
                      cursor: "pointer",
                      boxShadow: "0 2px 8px rgba(0,0,0,.06)",
                      border: "1.5px solid #E8ECF2",
                      display: "flex",
                      alignItems: "flex-start",
                      gap: 12,
                      justifyContent: "space-between",
                    }}
                  >
                    <div style={{ minWidth: 0, flex: 1 }}>
                      <div
                        style={{
                          display: "flex",
                          gap: 8,
                          alignItems: "center",
                          flexWrap: "wrap",
                          marginBottom: 4,
                        }}
                      >
                        <div style={{ fontWeight: 800, fontSize: 15, color: "#18243A" }}>{wr.numero_wr || "—"}</div>

                        {wr.priorita === "urgente" && (
                          <span style={{ color: "#C0392B", fontWeight: 800, fontSize: 11 }}>URGENTE</span>
                        )}

                        {(wr.photos || []).length > 0 && (
                          <span style={{ fontSize: 11, color: "#6B7FA3", fontWeight: 700 }}>📷 {wr.photos.length}</span>
                        )}

                        {isAdmin && data.prezzi[wr.id] != null && data.prezzi[wr.id] !== "" && (
                          <span style={{ fontSize: 11, color: "#2E7D32", fontWeight: 800 }}>
                            € {Number(data.prezzi[wr.id]).toFixed(2)}
                          </span>
                        )}
                      </div>

                      <div style={{ fontSize: 13, color: "#2B3545", marginTop: 2 }}>{shortText(wr.indirizzo, 60)}</div>
                      <div style={{ fontSize: 12, color: "#5B6575", marginTop: 4 }}>Tipo: {shortText(wr.tipo_intervento, 40)}</div>

                      <div
                        style={{
                          fontSize: 12,
                          color: "#6B7FA3",
                          display: "flex",
                          gap: 8,
                          flexWrap: "wrap",
                          marginTop: 8,
                          alignItems: "center",
                        }}
                      >
                        <CatBadge cat={wr.categoria} />
                        <span>•</span>
                        <span>{fmtDateTime(wr.updated_at)}</span>
                        {wr.ultimo_operatore && (
                          <>
                            <span>•</span>
                            <span>👷 {nomeUtente(wr.ultimo_operatore)}</span>
                          </>
                        )}
                      </div>
                    </div>

                    <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 8, flexShrink: 0 }}>
                      <Badge status={wr.status} />
                      <span style={{ fontSize: 18, color: "#9BA5B4", lineHeight: 1 }}>›</span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {view === "dettaglio" && selected && (
          <div style={{ background: "#fff", borderRadius: 14, padding: 16, border: "1px solid #E8ECF2" }}>
            <button
              onClick={() => {
                setSelectedId(null);
                setView("lista");
              }}
              style={BS}
            >
              ← Lista
            </button>

            <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#18243A" }}>{selected.numero_wr || "WR"}</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginTop: 6 }}>
                  <Badge status={selected.status} />
                  <CatBadge cat={selected.categoria} />
                  {selected.priorita === "urgente" && (
                    <span style={{ color: "#C0392B", fontWeight: 800, fontSize: 12 }}>URGENTE</span>
                  )}
                </div>
              </div>

              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button
                  onClick={() => {
                    if (navigator.geolocation) {
                      navigator.geolocation.getCurrentPosition(
                        (pos) => openNavigation(selected, { lat: pos.coords.latitude, lng: pos.coords.longitude }),
                        () => openNavigation(selected)
                      );
                    } else {
                      openNavigation(selected);
                    }
                  }}
                  style={{
                    padding: "8px 12px",
                    borderRadius: 10,
                    border: "none",
                    cursor: "pointer",
                    background: "#003087",
                    color: "#fff",
                    fontWeight: 700,
                  }}
                >
                  🧭 Naviga
                </button>

                <button
                  onClick={() => callReferente(selected.telefono || selected.tel_referente)}
                  disabled={!(selected.telefono || selected.tel_referente)}
                  style={{
                    padding: "8px 12px",
                    borderRadius: 10,
                    border: "none",
                    cursor: selected.telefono || selected.tel_referente ? "pointer" : "not-allowed",
                    background: selected.telefono || selected.tel_referente ? "#F05A22" : "#E8ECF2",
                    color: selected.telefono || selected.tel_referente ? "#fff" : "#9BA5B4",
                    fontWeight: 700,
                  }}
                >
                  📞 Chiama
                </button>
              </div>
            </div>

            <div style={{ height: 14 }} />

            {/* ── Distanza stradale ── */}
            <DistanzaBox wr={selected} />

            <div style={{ height: 14 }} />

            <div style={{ display: "grid", gap: 2 }}>
              <Row label="Tipo intervento" value={selected.tipo_intervento} />
              <Row label="Indirizzo" value={selected.indirizzo} highlight />
              {selected.comune && <Row label="Comune" value={selected.comune} />}
              {selected.provincia && <Row label="Provincia" value={selected.provincia} />}
              <Row label="Referente" value={selected.referente} />
              <Row label="Telefono" value={selected.telefono || selected.tel_referente} />
              <Row label="Priorità" value={selected.priorita} />
              {selected.tempo_obiettivo && <Row label="Tempo obiettivo" value={fmtTempoObiettivo(selected.tempo_obiettivo)} />}
              <Row label="Latitudine" value={selected.latitudine_intervento || selected.lat} />
              <Row label="Longitudine" value={selected.longitudine_intervento || selected.lng} />
              <Row label="Creata il" value={fmtDateTime(selected.created_at)} />
              <Row label="Ultimo aggiornamento" value={fmtDateTime(selected.updated_at)} />
              <Row label="Ultimo operatore" value={nomeUtente(selected.ultimo_operatore)} />
            </div>

            <div style={{ height: 14 }} />

            <div>
              <div style={LS}>Note WR</div>
              <div
                style={{
                  background: "#FAFBFD",
                  border: "1px solid #E8ECF2",
                  borderRadius: 10,
                  padding: 12,
                  fontSize: 13,
                  color: "#2B3545",
                  whiteSpace: "pre-wrap",
                }}
              >
                {selected.note || "—"}
              </div>
            </div>

            <div style={{ height: 14 }} />

            <div>
              <div style={LS}>Stato</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {STATUS_ORDER.map((s) => (
                  <button
                    key={s}
                    onClick={() => updateStatus(selected.id, s)}
                    style={{
                      padding: "8px 12px",
                      borderRadius: 10,
                      border: "none",
                      cursor: "pointer",
                      background: selected.status === s ? STATUS[s].dot : STATUS[s].bg,
                      color: selected.status === s ? "#fff" : STATUS[s].color,
                      fontWeight: 700,
                      fontSize: 12,
                    }}
                  >
                    {STATUS[s].label}
                  </button>
                ))}
              </div>
            </div>

            <div style={{ height: 14 }} />

            <div>
              <div style={LS}>Lavoro eseguito</div>
              <textarea
                value={editLavoro}
                onChange={(e) => setEditLavoro(e.target.value)}
                rows={5}
                placeholder="Descrivi il lavoro eseguito…"
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  borderRadius: 10,
                  border: "1.5px solid #D4D8DE",
                  fontSize: 13,
                  fontFamily: "inherit",
                  resize: "vertical",
                  boxSizing: "border-box",
                  outline: "none",
                }}
              />
              <button
                onClick={() => saveLavoro(selected.id)}
                style={{
                  marginTop: 8,
                  padding: "10px 14px",
                  borderRadius: 10,
                  background: "#003087",
                  color: "#fff",
                  fontWeight: 700,
                  border: "none",
                  cursor: "pointer",
                }}
              >
                Salva descrizione
              </button>
            </div>

            {isAdmin && (
              <>
                <div style={{ height: 14 }} />
                <div>
                  <div style={LS}>Prezzo (€)</div>
                  <input
                    type="number"
                    step="0.01"
                    value={priceDraft}
                    onChange={(e) => setPriceDraft(e.target.value)}
                    style={IS}
                    placeholder="0.00"
                  />
                  <button
                    onClick={() => savePrezzo(selected.id)}
                    style={{
                      marginTop: 8,
                      padding: "10px 14px",
                      borderRadius: 10,
                      background: "#28A745",
                      color: "#fff",
                      fontWeight: 700,
                      border: "none",
                      cursor: "pointer",
                    }}
                  >
                    Salva prezzo
                  </button>
                </div>
              </>
            )}

            <div style={{ height: 14 }} />

            <PhotoSection
              photos={selected.photos || []}
              onAdd={(file) => addPhoto(selected.id, file)}
              onDelete={(photoId) => deletePhoto(selected.id, photoId)}
              onPreview={(photo) => setLightbox(photo)}
            />

            {isAdmin && (
              <>
                <div style={{ height: 18 }} />
                <div style={{ display: "flex", gap: 10 }}>
                  <button
                    onClick={() => handleEditWR(selected)}
                    style={{
                      flex: 1,
                      padding: 12,
                      borderRadius: 10,
                      background: "#003087",
                      color: "#fff",
                      fontWeight: 700,
                      border: "none",
                      cursor: "pointer",
                    }}
                  >
                    ✏️ Modifica WR
                  </button>
                  <button
                    onClick={() => deleteWR(selected.id)}
                    style={{
                      flex: 1,
                      padding: 12,
                      borderRadius: 10,
                      background: "#C0392B",
                      color: "#fff",
                      fontWeight: 700,
                      border: "none",
                      cursor: "pointer",
                    }}
                  >
                    Elimina WR
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {view === "importa" && isAdmin && (
          <div style={{ background: "#fff", borderRadius: 14, padding: 16, border: "1px solid #E8ECF2" }}>
            <button onClick={() => setView("lista")} style={BS}>
              ← Lista
            </button>

            <div style={{ fontSize: 20, fontWeight: 800, color: "#003087", marginTop: 10 }}>Importa WR da Gmail</div>
            <div style={{ fontSize: 13, color: "#6B7FA3", marginBottom: 12 }}>Cerca automaticamente nelle mail TIM/Fibercop.</div>

            <button
              onClick={fetchGmail}
              style={{
                padding: "10px 14px",
                borderRadius: 10,
                background: "#003087",
                color: "#fff",
                fontWeight: 700,
                border: "none",
                cursor: "pointer",
              }}
            >
              {gmailLoading ? "Ricerca…" : "🔍 Cerca mail WR in Gmail"}
            </button>

            {gmailErr && <div style={{ marginTop: 10, color: "#C0392B", fontSize: 12 }}>{gmailErr}</div>}

            {gmailEmails.length > 0 && !extracted && (
              <div style={{ marginTop: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#6B7FA3", marginBottom: 8 }}>
                  Mail trovate ({gmailEmails.length})
                </div>

                <div style={{ display: "grid", gap: 10 }}>
                  {gmailEmails.map((e, i) => (
                    <div key={i} style={{ border: "1px solid #E8ECF2", borderRadius: 10, padding: 12 }}>
                      <div style={{ fontWeight: 700 }}>{e.subject}</div>
                      <div style={{ fontSize: 12, color: "#6B7FA3" }}>{e.from} · {e.date}</div>
                      <div style={{ fontSize: 12, margin: "8px 0" }}>{e.snippet?.slice(0, 150)}…</div>
                      <button
                        onClick={() => importFromGmailEmail(e)}
                        disabled={extracting}
                        style={{
                          padding: "7px 16px",
                          borderRadius: 8,
                          background: "#003087",
                          color: "#fff",
                          fontWeight: 600,
                          fontSize: 12,
                          border: "none",
                          cursor: "pointer",
                        }}
                      >
                        {extracting ? "Estrazione…" : "Estrai WR"}
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div style={{ borderTop: "1px solid #E8ECF2", paddingTop: 14, marginTop: 14, marginBottom: 14 }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "#9BA5B4",
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  marginBottom: 8,
                }}
              >
                Oppure importa un PDF
              </div>

              <input
                id="pdf-import-input"
                type="file"
                accept="application/pdf"
                style={{ display: "none" }}
                onChange={async (e) => {
                  const file = e.target.files?.[0];
                  if (file) await handlePdfFile(file);
                  e.target.value = "";
                }}
              />

              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={async (e) => {
                  e.preventDefault();
                  setDragOver(false);
                  const file = e.dataTransfer.files?.[0];
                  if (file) await handlePdfFile(file);
                }}
                onClick={() => document.getElementById("pdf-import-input")?.click()}
                style={{
                  border: `2px dashed ${dragOver ? "#003087" : "#D4D8DE"}`,
                  background: dragOver ? "#F0F7FF" : "#FAFBFC",
                  borderRadius: 12,
                  padding: "20px 16px",
                  textAlign: "center",
                  cursor: "pointer",
                }}
              >
                <div style={{ fontSize: 28, marginBottom: 6 }}>📄</div>
                <div style={{ fontSize: 14, fontWeight: 700, color: "#003087", marginBottom: 4 }}>Trascina qui il PDF</div>
                <div style={{ fontSize: 12, color: "#6B7FA3", marginBottom: 10 }}>oppure clicca per selezionarlo</div>

                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    document.getElementById("pdf-import-input")?.click();
                  }}
                  style={{
                    padding: "8px 14px",
                    borderRadius: 8,
                    background: "#F05A22",
                    color: "#fff",
                    fontWeight: 700,
                    fontSize: 12,
                    border: "none",
                    cursor: "pointer",
                  }}
                >
                  Importa PDF
                </button>

                {pdfName && <div style={{ marginTop: 10, fontSize: 12, color: "#4A5568" }}>File: {pdfName}</div>}
                {pdfLoading && <div style={{ marginTop: 10, fontSize: 12, color: "#003087", fontWeight: 600 }}>Lettura PDF in corso…</div>}
              </div>
            </div>

            <div style={{ borderTop: "1px solid #E8ECF2", paddingTop: 14 }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "#9BA5B4",
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  marginBottom: 8,
                }}
              >
                Oppure incolla il testo della mail
              </div>

              <textarea
                value={emailText}
                onChange={(e) => setEmailText(e.target.value)}
                rows={5}
                placeholder="Incolla il testo della mail WR…"
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  borderRadius: 10,
                  border: "1.5px solid #D4D8DE",
                  fontSize: 13,
                  fontFamily: "inherit",
                  resize: "vertical",
                  boxSizing: "border-box",
                  outline: "none",
                }}
              />

              <button
                onClick={handleExtract}
                disabled={extracting || !emailText.trim()}
                style={{
                  marginTop: 8,
                  width: "100%",
                  padding: 12,
                  borderRadius: 10,
                  background: extracting ? "#E8ECF2" : "#F05A22",
                  color: extracting ? "#9BA5B4" : "#fff",
                  fontWeight: 700,
                  fontSize: 14,
                  border: "none",
                  cursor: extracting ? "not-allowed" : "pointer",
                }}
              >
                {extracting ? "Estrazione…" : "Estrai dati WR"}
              </button>
            </div>

            {extractErr && <div style={{ marginTop: 10, color: "#C0392B", fontSize: 12 }}>{extractErr}</div>}

            {extracted && (
              <div
                style={{
                  marginTop: 14,
                  border: "1px solid #E8ECF2",
                  borderRadius: 12,
                  padding: 12,
                  background: "#FAFBFD",
                }}
              >
                <div style={{ fontWeight: 800, marginBottom: 8 }}>Dati estratti – controlla</div>

                {[
                  ["Numero WR", "numero_wr"],
                  ["Tipo", "tipo_intervento"],
                  ["Comune", "comune"],
                  ["Indirizzo", "indirizzo"],
                  ["Referente", "referente"],
                  ["Telefono", "telefono"],
                  ["Latitudine", "latitudine_intervento"],
                  ["Longitudine", "longitudine_intervento"],
                ].map(([l, k]) => (
                  <div key={k} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                    <div style={{ width: 110, fontSize: 12, fontWeight: 700, color: "#6B7FA3" }}>{l}</div>
                    <input
                      value={extracted[k] ?? ""}
                      onChange={(e) => setExtracted({ ...extracted, [k]: e.target.value })}
                      style={{
                        flex: 1,
                        padding: "5px 10px",
                        borderRadius: 7,
                        border: "1.5px solid #C6D8EE",
                        fontSize: 13,
                        fontFamily: "inherit",
                        outline: "none",
                        background: "#fff",
                      }}
                    />
                  </div>
                ))}

                {/* Priorità – tendina */}
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                  <div style={{ width: 110, fontSize: 12, fontWeight: 700, color: "#6B7FA3" }}>Priorità</div>
                  <select
                    value={extracted.priorita ?? "normale"}
                    onChange={(e) => setExtracted({ ...extracted, priorita: e.target.value })}
                    style={{ flex: 1, padding: "5px 10px", borderRadius: 7, border: "1.5px solid #C6D8EE", fontSize: 13, fontFamily: "inherit", outline: "none", background: "#fff" }}
                  >
                    <option value="normale">Normale</option>
                    <option value="urgente">Urgente</option>
                  </select>
                </div>

                {/* Note / Descrizione */}
                <div style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: 8 }}>
                  <div style={{ width: 110, fontSize: 12, fontWeight: 700, color: "#6B7FA3", paddingTop: 6 }}>Note</div>
                  <textarea
                    value={extracted.note ?? ""}
                    onChange={(e) => setExtracted({ ...extracted, note: e.target.value })}
                    rows={4}
                    style={{ flex: 1, padding: "5px 10px", borderRadius: 7, border: "1.5px solid #C6D8EE", fontSize: 13, fontFamily: "inherit", outline: "none", background: "#fff", resize: "vertical" }}
                  />
                </div>

                <button
                  onClick={confirmImport}
                  style={{
                    marginTop: 8,
                    width: "100%",
                    padding: 12,
                    borderRadius: 10,
                    background: "#28A745",
                    color: "#fff",
                    fontWeight: 800,
                    border: "none",
                    cursor: "pointer",
                  }}
                >
                  ✓ Conferma e aggiungi WR
                </button>
              </div>
            )}
          </div>
        )}

        {(view === "aggiungi" || view === "modifica") && isAdmin && (
          <div style={{ background: "#fff", borderRadius: 14, padding: 16, border: "1px solid #E8ECF2" }}>
            <button
              onClick={() => {
                if (view === "modifica") { setView("dettaglio"); setEditingId(null); setForm(emptyForm); }
                else setView("lista");
              }}
              style={BS}
            >
              ← {view === "modifica" ? "Dettaglio" : "Lista"}
            </button>

            <div style={{ fontSize: 18, fontWeight: 800, color: "#003087", margin: "10px 0 14px" }}>
              {view === "modifica" ? "✏️ Modifica WR" : "Nuova WR"}
            </div>

            <Field label="Numero WR" value={form.numero_wr} onChange={(v) => setForm({ ...form, numero_wr: v })} placeholder="es. WR-2024-001234" />
            <div style={{ height: 12 }} />

            <label style={LS}>Categoria</label>
            <select value={form.categoria} onChange={(e) => setForm({ ...form, categoria: e.target.value })} style={IS}>
              {CATEGORIE.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <div style={{ height: 12 }} />

            <label style={LS}>Tipo intervento</label>
            <select value={form.tipo_intervento} onChange={(e) => setForm({ ...form, tipo_intervento: e.target.value })} style={IS}>
              <option>Pozzetto</option>
              <option>Taglio rami</option>
              <option>Misto</option>
              <option>Altro</option>
            </select>
            <div style={{ height: 12 }} />

            <Field label="Indirizzo" value={form.indirizzo} onChange={(v) => setForm({ ...form, indirizzo: v })} placeholder="Via, città…" />
            <div style={{ height: 12 }} />

            <Field label="Comune" value={form.comune} onChange={(v) => setForm({ ...form, comune: v })} placeholder="es. Palermo" />
            <div style={{ height: 12 }} />

            <label style={LS}>Provincia</label>
            <select value={form.provincia} onChange={(e) => setForm({ ...form, provincia: e.target.value })} style={IS}>
              <option value="">— Dedotta automaticamente —</option>
              <option value="Palermo">Palermo</option>
              <option value="Agrigento">Agrigento</option>
              <option value="Trapani">Trapani</option>
              <option value="Altro">Altro</option>
            </select>
            <div style={{ height: 12 }} />

            <Field label="Referente" value={form.referente} onChange={(v) => setForm({ ...form, referente: v })} placeholder="Nome referente" />
            <div style={{ height: 12 }} />

            <Field label="Telefono" value={form.telefono} onChange={(v) => setForm({ ...form, telefono: v })} placeholder="es. 3331234567" />
            <div style={{ height: 12 }} />

            <Field label="Note" value={form.note} onChange={(v) => setForm({ ...form, note: v })} placeholder="Descrizione lavoro…" />
            <div style={{ height: 12 }} />

            <Field label="Tempo obiettivo" value={form.tempo_obiettivo} onChange={(v) => setForm({ ...form, tempo_obiettivo: v })} placeholder="es. 2025-12-31" />
            <div style={{ height: 12 }} />

            <label style={LS}>Priorità</label>
            <select value={form.priorita} onChange={(e) => setForm({ ...form, priorita: e.target.value })} style={IS}>
              <option value="normale">Normale</option>
              <option value="urgente">Urgente</option>
            </select>

            <button
              onClick={view === "modifica" ? handleSaveEdit : handleAdd}
              disabled={saving}
              style={{
                marginTop: 16, width: "100%", padding: 12, borderRadius: 10,
                background: saving ? "#E8ECF2" : view === "modifica" ? "#28A745" : "#003087",
                color: saving ? "#9BA5B4" : "#fff",
                fontWeight: 700, border: "none", cursor: saving ? "not-allowed" : "pointer",
              }}
            >
              {saving ? "Salvataggio…" : view === "modifica" ? "✓ Salva modifiche" : "Aggiungi WR"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

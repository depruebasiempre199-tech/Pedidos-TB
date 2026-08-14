import { useState, useMemo, useEffect, Fragment } from "react";
import * as XLSX from "xlsx";
import {
  ChevronDown,
  ChevronUp,
  Settings2,
  Plus,
  Trash2,
  Upload,
  Download,
  Printer,
  RefreshCw,
  Boxes,
  PackageCheck,
  TriangleAlert,
  ArrowLeft,
  Home,
  ClipboardList,
  History,
  CheckCircle2,
  Circle,
  Sun,
  Moon,
  Monitor,
} from "lucide-react";
import { fetchHistorial, saveSnapshot, updateSnapshot, deleteSnapshot, fetchEstado, saveEstado } from "./firebase-service";

const WEEKDAYS = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const DIAS_PEDIDO = [...WEEKDAYS];
const UNIDADES_BASE = ["Libras", "Unidades", "Onzas", "Galones", "Litros"];

// unidad = presentación de compra (lo que se le pide al proveedor)
// unidadBase = unidad común para comparar consumo/existencia/tránsito del mismo producto
// factor = cuántas unidadBase equivalen a 1 unidad de compra
const RAW = [
  { id: 1, producto: "Café Ristreto / Buen Café", proveedor: "Ristreto", tipo: "externo", dia: "Domingo", unidad: "Cs 40 / Lbs", unidadBase: "Libras", consumoSemanal: 168, diasInv: 12, existencia: 44, transito: 40, factor: 40, pedir: 6, pedirTocado: false },
  { id: 2, producto: "Carne Molida 40LB", proveedor: "Pexport", tipo: "externo", dia: "Domingo", unidad: "Cs 40 / Lbs", unidadBase: "Libras", consumoSemanal: 296.52, diasInv: 8, existencia: 322.8, transito: 80, factor: 40, pedir: 0, pedirTocado: false },
  { id: 3, producto: "Tomate", proveedor: "Grupo Incoagro", tipo: "externo", dia: "Domingo", unidad: "Libras", unidadBase: "Libras", consumoSemanal: 52.3, diasInv: 11, existencia: 45, transito: 30, factor: 1, pedir: 8, pedirTocado: false },
  { id: 4, producto: "Lechuga", proveedor: "Recursos Selectivos", tipo: "externo", dia: "Domingo", unidad: "BG 2.5 LB", unidadBase: "Libras", consumoSemanal: 45.23, diasInv: 10, existencia: 38.7, transito: 17.5, factor: 2.5, pedir: 4, pedirTocado: false },
  { id: 5, producto: "Café Ristreto / Buen Café", proveedor: "Ristreto", tipo: "externo", dia: "Martes", unidad: "Paquete 20 bls 3oz", unidadBase: "Libras", consumoSemanal: 168, diasInv: 16, existencia: 45, transito: 75, factor: 3.75, pedir: 71, pedirTocado: false },
  { id: 6, producto: "Carne Molida 40LB", proveedor: "Pexport", tipo: "externo", dia: "Martes", unidad: "Cs 40 / Lbs", unidadBase: "Libras", consumoSemanal: 296.52, diasInv: 10, existencia: 353, transito: 160, factor: 40, pedir: 0, pedirTocado: false },
  { id: 7, producto: "Steak", proveedor: "CD Congelado", tipo: "externo", dia: "Martes", unidad: "CS 40 LBS", unidadBase: "Libras", consumoSemanal: 91, diasInv: 8, existencia: 56.8, transito: 0, factor: 40, pedir: 2, pedirTocado: false },
  { id: 8, producto: "Tortilla HP 10.25", proveedor: "Bodega central", tipo: "interno", dia: "Domingo", unidad: "Cs / 144 Uni", unidadBase: "Unidades", consumoSemanal: 2053, diasInv: 6, existencia: 1995, transito: 0, factor: 144, pedir: 0, pedirTocado: false },
  { id: 9, producto: "Servilleta interfoleada", proveedor: "Bodega central", tipo: "interno", dia: "Domingo", unidad: "CS 6000 UN", unidadBase: "Unidades", consumoSemanal: 5000, diasInv: 6, existencia: 18000, transito: 6000, factor: 6000, pedir: 0, pedirTocado: false },
  { id: 10, producto: "Ketchup Heinz", proveedor: "Bodega central", tipo: "interno", dia: "Domingo", unidad: "Caja 1500 U", unidadBase: "Unidades", consumoSemanal: 1155, diasInv: 6, existencia: 1800, transito: 1400, factor: 1500, pedir: 0, pedirTocado: false },
  { id: 11, producto: "Tortilla HP 10.25", proveedor: "Bodega central", tipo: "interno", dia: "Jueves", unidad: "Cs / 144 Uni", unidadBase: "Unidades", consumoSemanal: 2053, diasInv: 6, existencia: 1995, transito: 10, factor: 144, pedir: 0, pedirTocado: false },
  { id: 12, producto: "Salsa queso cheddar", proveedor: "Bodega central", tipo: "interno", dia: "Jueves", unidad: "CS 10BG 3.75LB", unidadBase: "Libras", consumoSemanal: 37, diasInv: 6, existencia: 81, transito: 3, factor: 37.5, pedir: 0, pedirTocado: false },
];

const PROVEEDORES_INICIALES = [
  { nombre: "Ristreto", tipo: "externo" },
  { nombre: "Pexport", tipo: "externo" },
  { nombre: "Grupo Incoagro", tipo: "externo" },
  { nombre: "Recursos Selectivos", tipo: "externo" },
  { nombre: "CD Congelado", tipo: "externo" },
  { nombre: "Bodega central", tipo: "interno" },
];

const DEFAULT_SCHEDULE = { Ristreto: "Jueves", Pexport: "Lunes", "Grupo Incoagro": "Jueves", "Recursos Selectivos": "Viernes", "CD Congelado": "Lunes", "Bodega central": "Viernes" };

function defaultDiaProveedores() {
  const map = {};
  DIAS_PEDIDO.forEach((d) => (map[d] = new Set()));
  RAW.forEach((r) => map[r.dia]?.add(r.proveedor));
  return map;
}

function normalize(s) {
  return (s || "").toString().trim().toLowerCase();
}

function parseNumeroExcel(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined) return NaN;
  const limpio = v.toString().trim().replace(/,/g, "");
  return Number(limpio);
}

function itemKey(producto, proveedor, dia) {
  return `${normalize(producto)}|${normalize(proveedor)}|${dia}`;
}

function nextDateForWeekday(baseDate, weekdayName) {
  const base = new Date(baseDate + "T00:00:00");
  const targetIdx = WEEKDAYS.indexOf(weekdayName);
  if (targetIdx < 0 || isNaN(base.getTime())) return null;
  const baseIdx = base.getDay();
  let diff = (targetIdx - baseIdx + 7) % 7;
  if (diff === 0) diff = 7;
  const result = new Date(base);
  result.setDate(base.getDate() + diff);
  return result;
}

function formatDate(d) {
  if (!d) return "—";
  return d.toLocaleDateString("es-GT", { weekday: "short", day: "2-digit", month: "short" });
}

function computeRow(r, diasInvConfig) {
  const key = `${r.proveedor}|${r.dia}`;
  const diasInv = diasInvConfig[key] ?? r.diasInv;
  const diario = r.consumoSemanal / 7;
  const proyeccion = diario * diasInv - r.existencia - r.transito;
  const pedirSugerido = Math.max(0, Math.ceil(proyeccion / r.factor));
  const cobertura = r.existencia / (diario || 1);
  return { ...r, diasInv, diario, proyeccion, pedirSugerido, cobertura };
}

function PreviewSeleccion({ preview, onToggle, onMarcarTodos, onAplicar, aplicarLabel, colActualLabel }) {
  if (preview.length === 0) return null;
  return (
    <div style={{ marginTop: 14 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
        <button onClick={() => onMarcarTodos(true)} style={miniButtonStyle}>Marcar todos</button>
        <button onClick={() => onMarcarTodos(false)} style={miniButtonStyle}>Desmarcar todos</button>
        <button onClick={onAplicar} style={{ ...buttonStyle, background: "var(--ok)", color: "#fff", borderColor: "var(--ok)" }}>
          {aplicarLabel}
        </button>
        <span style={{ fontSize: 11.5, color: "var(--text-secondary)" }}>
          {preview.filter((p) => p.aplicar).length} de {preview.length} seleccionados
        </span>
      </div>
      <div style={{ maxHeight: 360, overflowY: "auto", border: "1px solid var(--border)", borderRadius: 8 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
          <thead>
            <tr style={{ background: "var(--track)", textAlign: "left" }}>
              <th style={{ padding: "6px 8px", width: 30 }}></th>
              <th style={{ padding: "6px 8px" }}>Producto</th>
              <th style={{ padding: "6px 8px" }}>Proveedor</th>
              <th style={{ padding: "6px 8px" }}>Día</th>
              <th style={{ padding: "6px 8px", textAlign: "right" }}>{colActualLabel}</th>
              <th style={{ padding: "6px 8px", textAlign: "right" }}>Valor nuevo</th>
            </tr>
          </thead>
          <tbody>
            {preview.map((p) => (
              <tr key={p.id} style={{ borderTop: "1px solid var(--border)", background: p.aplicar ? "var(--track)" : "transparent" }}>
                <td style={{ padding: "6px 8px" }}>
                  <input type="checkbox" checked={p.aplicar} onChange={() => onToggle(p.id)} />
                </td>
                <td style={{ padding: "6px 8px" }}>{p.producto}</td>
                <td style={{ padding: "6px 8px", color: "var(--text-secondary)" }}>{p.proveedor}</td>
                <td style={{ padding: "6px 8px", color: "var(--text-secondary)" }}>{p.dia}</td>
                <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: "var(--mono)" }}>{p.actual}</td>
                <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: "var(--mono)", fontWeight: 600 }}>{p.nuevo}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ProductCard({ r, despachoLabel, updateText, update, updateUnidadBase, eliminarFila }) {
  const tipoColor = r.tipo === "externo" ? "var(--accent-ext)" : "var(--accent-int)";
  return (
    <div style={{ background: "var(--card)", border: "1px solid var(--border)", borderLeft: `3px solid ${tipoColor}`, borderRadius: 10, padding: 12, boxShadow: "var(--shadow)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8, marginBottom: 8 }}>
        <input
          type="text"
          value={r.producto}
          onChange={(e) => updateText(r.id, "producto", e.target.value)}
          style={{ ...textInputStyle, fontSize: 14, fontWeight: 600, border: "none", padding: "2px 0", background: "transparent" }}
        />
        <button onClick={() => eliminarFila(r.id)} style={dangerLinkStyle}><Trash2 size={13} /></button>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{r.proveedor}</span>
        <span style={{ fontSize: 10.5, fontWeight: 500, padding: "2px 8px", borderRadius: 20, background: r.tipo === "externo" ? "#F0E6DC" : "#F5EBC8", color: tipoColor }}>
          {r.tipo}
        </span>
      </div>

      <Bar cobertura={r.cobertura} objetivo={r.diasInv} />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10 }}>
        <div>
          <div style={miniLabelStyle}>Existencia (auto)</div>
          <div style={{ fontFamily: "var(--mono)", fontSize: 14 }}>{r.existencia} <span style={{ fontSize: 10, color: "var(--text-secondary)" }}>{r.unidadBase}</span></div>
        </div>
        <div>
          <div style={miniLabelStyle}>Tránsito</div>
          <ConvertibleInput value={r.transito} unidadBase={r.unidadBase} unidadCompra={r.unidad} factor={r.factor} onChange={(v) => update(r.id, "transito", v)} />
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10 }}>
        <div>
          <div style={miniLabelStyle}>Proyección</div>
          <div style={{ fontFamily: "var(--mono)", fontSize: 14, color: r.proyeccion < 0 ? "var(--text-secondary)" : "var(--text-primary)" }}>
            {r.proyeccion.toFixed(1)}
          </div>
        </div>
        <div>
          <div style={miniLabelStyle}>Sugerido a pedir</div>
          <div style={{ fontFamily: "var(--mono)", fontSize: 14, color: r.pedirSugerido > 0 ? "var(--danger)" : "var(--ok)" }}>
            {r.pedirSugerido} <span style={{ fontSize: 10, color: "var(--text-secondary)" }}>{r.unidad}</span>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 10 }}>
        <div>
          <div style={{ ...miniLabelStyle, display: "flex", alignItems: "center", gap: 4 }}>
            Pedir {r.pedirTocado ? <CheckCircle2 size={12} color="var(--ok)" /> : <Circle size={12} color="var(--border)" />}
          </div>
          <input type="number" value={r.pedir} onChange={(e) => update(r.id, "pedir", e.target.value)} style={{ ...inputStyle, width: "100%", fontWeight: 600, color: r.pedir > 0 ? "var(--danger)" : "var(--ok)" }} />
          {r.pedirSugerido !== r.pedir && (
            <button onClick={() => update(r.id, "pedir", r.pedirSugerido)} style={{ ...miniButtonStyle, fontSize: 10, marginTop: 3 }}>sug. {r.pedirSugerido}</button>
          )}
        </div>
        <div>
          <div style={miniLabelStyle}>Entrega</div>
          <div style={{ fontSize: 13, marginTop: 6 }}>{despachoLabel}</div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 10, paddingTop: 8, borderTop: "1px solid var(--border)" }}>
        <input type="text" value={r.unidad} onChange={(e) => updateText(r.id, "unidad", e.target.value)} style={{ ...textInputStyle, width: "60%", fontSize: 11, border: "none", padding: 0, color: "var(--text-secondary)" }} />
        <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>{r.diario.toFixed(1)}/día &middot; obj. {r.diasInv}d</span>
      </div>
    </div>
  );
}

const miniLabelStyle = { fontSize: 10, color: "var(--text-secondary)", textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 3 };

const TABS = [
  { key: "resumen", label: "Resumen", icon: Home },
  { key: "pedidos", label: "Pedidos", icon: ClipboardList },
  { key: "historial", label: "Historial", icon: History },
  { key: "config", label: "Configuración", icon: Settings2 },
];

function Accordion({ title, isOpen, onToggle, children }) {
  return (
    <div style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 10, boxShadow: "var(--shadow)", overflow: "hidden" }}>
      <button
        onClick={onToggle}
        style={{
          width: "100%",
          textAlign: "left",
          padding: "11px 14px",
          background: isOpen ? "var(--header-bg)" : "var(--card)",
          border: "none",
          cursor: "pointer",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          fontSize: 12.5,
          fontWeight: 600,
          color: isOpen ? "var(--header-text)" : "var(--text-primary)",
        }}
      >
        {title}
        {isOpen ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
      </button>
      {isOpen && <div style={{ padding: 14 }}>{children}</div>}
    </div>
  );
}

function KpiCard({ icon, label, value, tone }) {
  const color = tone === "danger" ? "var(--danger)" : "var(--text-primary)";
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 10,
        background: "var(--card)",
        border: "1px solid var(--border)",
        borderRadius: 10,
        padding: "8px 14px",
        boxShadow: "var(--shadow)",
      }}
    >
      <span style={{ color: "var(--text-secondary)" }}>{icon}</span>
      <div>
        <div style={{ fontFamily: "var(--mono)", fontSize: 17, fontWeight: 600, color, lineHeight: 1.1 }}>{value}</div>
        <div style={{ fontSize: 10, color: "var(--text-secondary)", textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</div>
      </div>
    </div>
  );
}

function Bar({ cobertura, objetivo }) {
  const pct = Math.min(150, (cobertura / objetivo) * 100);
  const color = pct < 60 ? "var(--danger)" : pct < 100 ? "var(--warn)" : "var(--ok)";
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ width: 64, height: 6, borderRadius: 3, background: "var(--track)", overflow: "hidden" }}>
        <div style={{ width: `${Math.min(100, pct)}%`, height: "100%", background: color, borderRadius: 3 }} />
      </div>
      <span style={{ fontSize: 11, color: "var(--text-secondary)", fontFamily: "var(--mono)", minWidth: 34 }}>
        {cobertura.toFixed(1)}d
      </span>
    </div>
  );
}

// Input principal en unidad base, más un campo chico para convertir desde la
// unidad de compra (ej: "tengo 2 cajas" -> escribe 2, toca usar, y carga 80 lb).
function ConvertibleInput({ value, unidadBase, unidadCompra, factor, onChange }) {
  const [convValue, setConvValue] = useState("");
  const aplicarConversion = () => {
    const n = Number(convValue);
    if (!isNaN(n)) onChange(n * factor);
    setConvValue("");
  };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
        <input type="number" value={value} onChange={(e) => onChange(Number(e.target.value) || 0)} style={inputStyle} />
        <span style={{ fontSize: 10, color: "var(--text-secondary)" }}>{unidadBase}</span>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
        <input
          type="number"
          placeholder={`en ${unidadCompra}`}
          value={convValue}
          onChange={(e) => setConvValue(e.target.value)}
          style={{ ...inputStyle, width: 54, fontSize: 10 }}
        />
        <button onClick={aplicarConversion} style={miniButtonStyle} title={`convertir a ${unidadBase} (x${factor})`}>
          <RefreshCw size={11} />
        </button>
      </div>
    </div>
  );
}

function chequearConsistencia(rows) {
  const porProducto = {};
  rows.forEach((r) => {
    const k = normalize(r.producto);
    if (!porProducto[k]) porProducto[k] = { nombre: r.producto, unidades: new Set() };
    porProducto[k].unidades.add(r.unidadBase);
  });
  return Object.values(porProducto).filter((p) => p.unidades.size > 1);
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

const LIGHT_VARS = {
  "--bg": "#F5F3EF",
  "--card": "#FFFFFF",
  "--border": "#E6E1D6",
  "--text-primary": "#1B1B1B",
  "--text-secondary": "#8A8275",
  "--danger": "#C1401C",
  "--warn": "#A87615",
  "--ok": "#147D5A",
  "--track": "#EFE9DA",
  "--accent-ext": "#5B4636",
  "--accent-int": "#8A6B00",
  "--header-bg": "#1B1B1B",
  "--header-text": "#F4B400",
  "--brand-yellow": "#F4B400",
  "--hover-bg": "#EDEEE7",
  "--row-hover": "#FAFAF6",
  "--danger-soft-bg": "#FBEEE7",
  "--danger-soft-border": "#F0D3C4",
  "--mono": "ui-monospace, SFMono-Regular, Menlo, monospace",
  "--shadow": "0 1px 2px rgba(27,27,27,0.05), 0 6px 16px rgba(27,27,27,0.06)",
};

const DARK_VARS = {
  "--bg": "#141414",
  "--card": "#1F1E1B",
  "--border": "#37352C",
  "--text-primary": "#F2EFE6",
  "--text-secondary": "#A39C8C",
  "--danger": "#E2694A",
  "--warn": "#D9A63D",
  "--ok": "#3FBE95",
  "--track": "#2A2822",
  "--accent-ext": "#C9A98A",
  "--accent-int": "#E4C34A",
  "--header-bg": "#0E0E0D",
  "--header-text": "#F4B400",
  "--brand-yellow": "#F4B400",
  "--hover-bg": "#2A2822",
  "--row-hover": "#242219",
  "--danger-soft-bg": "#3A241C",
  "--danger-soft-border": "#5A3527",
  "--mono": "ui-monospace, SFMono-Regular, Menlo, monospace",
  "--shadow": "0 1px 2px rgba(0,0,0,0.4), 0 6px 16px rgba(0,0,0,0.5)",
};

const GLOBAL_CSS = `

  .pp-kpi-row { }
  .pp-filters { }
  .pp-table-wrap { display: block; }
  .pp-cards { display: none; }
  .pp-btn:hover { background: var(--hover-bg); }
  .pp-btn-dark:hover { opacity: 0.88; }
  .pp-row:hover { background: var(--row-hover); }
  .pp-tabs { display: flex; gap: 6px; overflow-x: auto; }
  .pp-tab { display: flex; align-items: center; gap: 6px; padding: 9px 14px; border-radius: 9px 9px 0 0; border: none; background: transparent; font-size: 12.5px; font-weight: 500; color: var(--text-secondary); cursor: pointer; white-space: nowrap; }
  .pp-tab.active { background: var(--header-bg); color: var(--header-text); box-shadow: var(--shadow); }
  @media (max-width: 760px) {
    .pp-shell { padding: 14px !important; }
    .pp-kpi-row { flex-wrap: nowrap !important; overflow-x: auto; padding-bottom: 4px; width: 100%; }
    .pp-kpi-row > div { flex: 0 0 auto; }
    .pp-filters { flex-direction: column !important; align-items: stretch !important; }
    .pp-filters > label { width: 100%; }
    .pp-filters select { width: 100%; }
    .pp-filters > button { width: 100%; justify-content: center; }
    .pp-table-wrap { display: none !important; }
    .pp-cards { display: flex !important; flex-direction: column; gap: 10px; }
    .pp-config-inline { flex-direction: column !important; align-items: stretch !important; }
    .pp-tab { padding: 8px 10px; font-size: 11px; }
    .pp-tab span.pp-tab-label { display: none; }
  }
`;

function StyleBlock() {
  return <style>{GLOBAL_CSS}</style>;
}

function LogoIcon({ size = 40 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx="60" cy="60" r="56" fill="#1B1B1B" />
      <path
        d="M 60.0,26.0 L 65.74,46.14 L 84.04,35.96 L 73.86,54.26 L 94.0,60.0 L 73.86,65.74 L 84.04,84.04 L 65.74,73.86 L 60.0,94.0 L 54.26,73.86 L 35.96,84.04 L 46.14,65.74 L 26.0,60.0 L 46.14,54.26 L 35.96,35.96 L 54.26,46.14 Z"
        fill="#F4B400"
      />
    </svg>
  );
}

function Logo({ subtitle = "ABASTECIMIENTO" }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      <LogoIcon size={40} />
      <div>
        <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: -0.4, color: "var(--text-primary)" }}>BURST KITCHEN</div>
        <div style={{ fontSize: 9.5, fontWeight: 600, letterSpacing: 2, color: "var(--text-secondary)" }}>{subtitle}</div>
      </div>
    </div>
  );
}

const THEME_OPTIONS = [
  { key: "light", label: "Claro", icon: Sun },
  { key: "system", label: "Sistema", icon: Monitor },
  { key: "dark", label: "Oscuro", icon: Moon },
];

function ThemeToggle({ themeMode, setThemeMode }) {
  return (
    <div style={{ display: "flex", gap: 2, background: "var(--track)", borderRadius: 9, padding: 2, border: "1px solid var(--border)" }}>
      {THEME_OPTIONS.map(({ key, label, icon: Icon }) => (
        <button
          key={key}
          onClick={() => setThemeMode(key)}
          title={label}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 5,
            padding: "5px 9px",
            borderRadius: 7,
            border: "none",
            cursor: "pointer",
            fontSize: 11,
            fontWeight: 500,
            background: themeMode === key ? "var(--header-bg)" : "transparent",
            color: themeMode === key ? "var(--header-text)" : "var(--text-secondary)",
          }}
        >
          <Icon size={13} />
        </button>
      ))}
    </div>
  );
}

let nextId = 100;

export default function PedidosPrototype() {
  const [rows, setRows] = useState(RAW);
  const [extraKeys, setExtraKeys] = useState([]);
  const [nuevaCombo, setNuevaCombo] = useState({ proveedor: "", dia: DIAS_PEDIDO[0] });
  const [dia, setDia] = useState("Todos");
  const [tipo, setTipo] = useState("Todos");
  const [proveedor, setProveedor] = useState("Todos");
  const [activeTab, setActiveTab] = useState("pedidos");
  const [themeMode, setThemeMode] = useState("system");
  const [systemDark, setSystemDark] = useState(false);

  useEffect(() => {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    setSystemDark(mq.matches);
    const handler = (e) => setSystemDark(e.matches);
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);

  const isDark = themeMode === "dark" || (themeMode === "system" && systemDark);
  const [sectionOpen, setSectionOpen] = useState({ uploads: true, transito: false, proveedores: false, diasInv: false, calendario: false, consistencia: false });
  const [diasInvConfig, setDiasInvConfig] = useState({});
  const [schedule, setSchedule] = useState(DEFAULT_SCHEDULE);
  const [fechaPedido, setFechaPedido] = useState(new Date().toISOString().slice(0, 10));
  const [uploadStatus, setUploadStatus] = useState("");
  const [existenciaDia, setExistenciaDia] = useState("Jueves");
  const [existenciaUltima, setExistenciaUltima] = useState(() => Object.fromEntries(DIAS_PEDIDO.map((d) => [d, null])));
  const [existenciaPreview, setExistenciaPreview] = useState([]);
  const [existenciaArchivo, setExistenciaArchivo] = useState("");
  const [consumoPreview, setConsumoPreview] = useState([]);
  const [consumoArchivo, setConsumoArchivo] = useState("");
  const [transitoPreview, setTransitoPreview] = useState([]);
  const [transitoArchivo, setTransitoArchivo] = useState("");
  const [consumoUltima, setConsumoUltima] = useState(null);
  const [diaProveedores, setDiaProveedores] = useState(defaultDiaProveedores);
  const [proveedoresList, setProveedoresList] = useState(PROVEEDORES_INICIALES);
  const [nuevoProveedor, setNuevoProveedor] = useState({ nombre: "", tipo: "externo" });
  const [historial, setHistorial] = useState([]);
  const [historialStatus, setHistorialStatus] = useState("");
  const [dbStatus, setDbStatus] = useState("conectando…");
  const [vistaResumen, setVistaResumen] = useState(false);
  const [resumenDia, setResumenDia] = useState("Todos");
  const [resumenMostrarTodos, setResumenMostrarTodos] = useState(false);
  const [columnasCompactas, setColumnasCompactas] = useState(true);
  const [estadoListo, setEstadoListo] = useState(false);
  const [guardadoStatus, setGuardadoStatus] = useState("");

  useEffect(() => {
    fetchHistorial()
      .then((data) => {
        setHistorial(data);
        setDbStatus("conectado a Firebase (Realtime Database)");
      })
      .catch((err) => {
        console.error(err);
        setDbStatus("error de conexión a Firebase — revisá firebase-config.js");
      });
  }, []);

  // Cargar el estado guardado (productos, proveedores, configuración) al abrir la app.
  useEffect(() => {
    fetchEstado()
      .then((estado) => {
        if (estado) {
          if (estado.rows) setRows(estado.rows);
          if (estado.extraKeys) setExtraKeys(estado.extraKeys);
          if (estado.diasInvConfig) setDiasInvConfig(estado.diasInvConfig);
          if (estado.schedule) setSchedule(estado.schedule);
          if (estado.proveedoresList) setProveedoresList(estado.proveedoresList);
          if (estado.diaProveedores) {
            const reconstruido = {};
            DIAS_PEDIDO.forEach((d) => {
              reconstruido[d] = new Set(estado.diaProveedores[d] || []);
            });
            setDiaProveedores(reconstruido);
          }
          if (estado.fechaPedido) setFechaPedido(estado.fechaPedido);
          const maxId = Math.max(0, ...(estado.rows || []).map((r) => r.id));
          if (maxId >= nextId) nextId = maxId + 1;
        }
        setEstadoListo(true);
      })
      .catch((err) => {
        console.error(err);
        setEstadoListo(true);
      });
  }, []);

  // Autoguardado: cada vez que cambian productos, proveedores o configuración,
  // se guarda en Firebase (con una pequeña espera para no guardar en cada tecla).
  useEffect(() => {
    if (!estadoListo) return;
    setGuardadoStatus("guardando…");
    const timeout = setTimeout(() => {
      const estado = {
        rows,
        extraKeys,
        diasInvConfig,
        schedule,
        proveedoresList,
        diaProveedores: Object.fromEntries(Object.entries(diaProveedores).map(([k, v]) => [k, [...v]])),
        fechaPedido,
      };
      saveEstado(estado)
        .then(() => setGuardadoStatus("guardado ✓"))
        .catch((err) => {
          console.error(err);
          setGuardadoStatus("no se pudo guardar — revisá tu conexión");
        });
    }, 1200);
    return () => clearTimeout(timeout);
  }, [estadoListo, rows, extraKeys, diasInvConfig, schedule, proveedoresList, diaProveedores, fechaPedido]);

  const toggleSection = (key) => setSectionOpen((prev) => ({ ...prev, [key]: !prev[key] }));

  const todosProveedores = useMemo(() => proveedoresList.map((p) => p.nombre), [proveedoresList]);
  const proveedoresFiltro = useMemo(() => ["Todos", ...todosProveedores], [todosProveedores]);
  const diaOptions = ["Todos", ...DIAS_PEDIDO];
  const tipoOptions = ["Todos", "externo", "interno"];
  const proveedorDiaKeys = useMemo(
    () => [...new Set([...rows.map((r) => `${r.proveedor}|${r.dia}`), ...extraKeys])],
    [rows, extraKeys]
  );

  const update = (id, field, value) => {
    setRows((prev) =>
      prev.map((r) =>
        r.id === id
          ? { ...r, [field]: Number(value) || 0, ...(field === "pedir" ? { pedirTocado: true } : {}) }
          : r
      )
    );
  };

  const updateText = (id, field, value) => {
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  };

  const agregarCombo = () => {
    if (!nuevaCombo.proveedor) return;
    const key = `${nuevaCombo.proveedor}|${nuevaCombo.dia}`;
    setExtraKeys((prev) => (prev.includes(key) ? prev : [...prev, key]));
  };

  const eliminarCombo = (key) => {
    setExtraKeys((prev) => prev.filter((k) => k !== key));
    setDiasInvConfig((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    setSchedule((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  const updateDiasInv = (key, value) => {
    setDiasInvConfig((prev) => ({ ...prev, [key]: Number(value) || 0 }));
  };

  const updateSchedule = (prov, weekday) => {
    setSchedule((prev) => ({ ...prev, [prov]: weekday }));
  };

  const toggleProveedorDia = (prov, diaKey) => {
    setDiaProveedores((prev) => {
      const next = { ...prev, [diaKey]: new Set(prev[diaKey]) };
      if (next[diaKey].has(prov)) next[diaKey].delete(prov);
      else next[diaKey].add(prov);
      return next;
    });
  };

  const agregarProveedor = () => {
    const nombre = nuevoProveedor.nombre.trim();
    if (!nombre || todosProveedores.includes(nombre)) return;
    setProveedoresList((prev) => [...prev, { nombre, tipo: nuevoProveedor.tipo }]);
    setSchedule((prev) => ({ ...prev, [nombre]: "Lunes" }));
    setNuevoProveedor({ nombre: "", tipo: "externo" });
  };

  const moveProveedor = (nombre, direction) => {
    setProveedoresList((prev) => {
      const idx = prev.findIndex((p) => p.nombre === nombre);
      const newIdx = idx + direction;
      if (idx < 0 || newIdx < 0 || newIdx >= prev.length) return prev;
      const next = [...prev];
      [next[idx], next[newIdx]] = [next[newIdx], next[idx]];
      return next;
    });
  };

  const updateProveedorTipo = (nombre, tipo) => {
    setProveedoresList((prev) => prev.map((p) => (p.nombre === nombre ? { ...p, tipo } : p)));
    setRows((prev) => prev.map((r) => (r.proveedor === nombre ? { ...r, tipo } : r)));
  };

  const eliminarProveedor = (nombre) => {
    setProveedoresList((prev) => prev.filter((p) => p.nombre !== nombre));
    setRows((prev) => prev.filter((r) => r.proveedor !== nombre));
    setSchedule((prev) => {
      const next = {};
      Object.entries(prev).forEach(([k, v]) => {
        if (k !== nombre && !k.startsWith(`${nombre}|`)) next[k] = v;
      });
      return next;
    });
    setDiaProveedores((prev) => {
      const next = {};
      DIAS_PEDIDO.forEach((d) => {
        next[d] = new Set(prev[d]);
        next[d].delete(nombre);
      });
      return next;
    });
    if (proveedor === nombre) setProveedor("Todos");
  };

  const agregarProducto = (prov, diaSel) => {
    const tipoProv = proveedoresList.find((p) => p.nombre === prov)?.tipo || "externo";
    const nuevo = {
      id: nextId++,
      producto: "Nuevo producto",
      proveedor: prov,
      tipo: tipoProv,
      dia: diaSel,
      unidad: "unidad",
      unidadBase: "Unidades",
      consumoSemanal: 0,
      diasInv: 7,
      existencia: 0,
      transito: 0,
      factor: 1,
      pedir: 0,
      pedirTocado: false,
    };
    setRows((prev) => [...prev, nuevo]);
    setDiaProveedores((prev) => {
      const next = { ...prev, [diaSel]: new Set(prev[diaSel]) };
      next[diaSel].add(prov);
      return next;
    });
  };

  const eliminarFila = (id) => {
    setRows((prev) => prev.filter((r) => r.id !== id));
  };

  const handleConsumoUpload = async (file) => {
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const updates = {};
      data.forEach((row) => {
        const nombre = normalize(row[0]);
        const valor = parseNumeroExcel(row[29]);
        if (nombre && !isNaN(valor)) updates[nombre] = valor;
      });
      const preview = rows
        .filter((r) => updates[normalize(r.producto)] !== undefined)
        .map((r) => ({
          id: r.id,
          producto: r.producto,
          proveedor: r.proveedor,
          dia: r.dia,
          actual: r.consumoSemanal,
          nuevo: updates[normalize(r.producto)],
          aplicar: true,
        }));
      setConsumoPreview(preview);
      setConsumoArchivo(file.name);
      setUploadStatus(
        preview.length > 0
          ? `Se encontraron ${preview.length} coincidencias de consumo. Revisá la lista y desmarcá lo que no quieras aplicar.`
          : "No se encontraron productos coincidentes en el archivo."
      );
    } catch (err) {
      setUploadStatus("No se pudo leer el archivo de consumo semanal.");
    }
  };

  const toggleConsumoPreview = (id) => {
    setConsumoPreview((prev) => prev.map((p) => (p.id === id ? { ...p, aplicar: !p.aplicar } : p)));
  };

  const marcarTodosConsumo = (valor) => {
    setConsumoPreview((prev) => prev.map((p) => ({ ...p, aplicar: valor })));
  };

  const aplicarConsumoSeleccionados = () => {
    const seleccionados = consumoPreview.filter((p) => p.aplicar);
    if (seleccionados.length === 0) return;
    const map = {};
    seleccionados.forEach((p) => {
      map[p.id] = p.nuevo;
    });
    setRows((prev) => prev.map((r) => (map[r.id] !== undefined ? { ...r, consumoSemanal: map[r.id] } : r)));
    setUploadStatus(`Consumo semanal actualizado en ${seleccionados.length} producto(s).`);
    setConsumoUltima({ fecha: new Date(), archivo: consumoArchivo, count: seleccionados.length });
    setConsumoPreview((prev) => prev.filter((p) => !p.aplicar));
  };

  const handleExistenciaUpload = async (file) => {
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const updates = {};
      data.forEach((row) => {
        const nombre = normalize(row[0]);
        const valor = parseNumeroExcel(row[16]);
        if (nombre && !isNaN(valor)) updates[nombre] = valor;
      });
      const preview = rows
        .filter((r) => r.dia === existenciaDia && updates[normalize(r.producto)] !== undefined)
        .map((r) => ({
          id: r.id,
          producto: r.producto,
          proveedor: r.proveedor,
          dia: r.dia,
          actual: r.existencia,
          nuevo: updates[normalize(r.producto)],
          aplicar: true,
        }));
      setExistenciaPreview(preview);
      setExistenciaArchivo(file.name);
      setUploadStatus(
        preview.length > 0
          ? `Se encontraron ${preview.length} coincidencias de existencia para ${existenciaDia}. Revisá la lista y desmarcá lo que no quieras aplicar.`
          : `No se encontraron productos coincidentes de ${existenciaDia} en el archivo.`
      );
    } catch (err) {
      setUploadStatus("No se pudo leer el archivo de existencia.");
    }
  };

  const toggleExistenciaPreview = (id) => {
    setExistenciaPreview((prev) => prev.map((p) => (p.id === id ? { ...p, aplicar: !p.aplicar } : p)));
  };

  const marcarTodosExistencia = (valor) => {
    setExistenciaPreview((prev) => prev.map((p) => ({ ...p, aplicar: valor })));
  };

  const aplicarExistenciaSeleccionados = () => {
    const seleccionados = existenciaPreview.filter((p) => p.aplicar);
    if (seleccionados.length === 0) return;
    const map = {};
    seleccionados.forEach((p) => {
      map[p.id] = p.nuevo;
    });
    setRows((prev) => prev.map((r) => (map[r.id] !== undefined ? { ...r, existencia: map[r.id] } : r)));
    setUploadStatus(`Existencia de ${existenciaDia} actualizada en ${seleccionados.length} producto(s).`);
    setExistenciaUltima((prev) => ({ ...prev, [existenciaDia]: { fecha: new Date(), archivo: existenciaArchivo, count: seleccionados.length } }));
    setExistenciaPreview((prev) => prev.filter((p) => !p.aplicar));
  };

  const handleTransitoUpload = async (file) => {
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const updates = {};
      data.forEach((row) => {
        const nombre = normalize(row[0]);
        const valor = parseNumeroExcel(row[16]);
        if (nombre && !isNaN(valor)) updates[nombre] = valor;
      });
      const preview = rows
        .filter((r) => updates[normalize(r.producto)] !== undefined)
        .map((r) => ({
          id: r.id,
          producto: r.producto,
          proveedor: r.proveedor,
          dia: r.dia,
          actual: r.transito,
          nuevo: updates[normalize(r.producto)],
          aplicar: false,
        }));
      setTransitoPreview(preview);
      setTransitoArchivo(file.name);
      setUploadStatus(
        preview.length > 0
          ? `Se encontraron ${preview.length} coincidencias. Revisá la lista abajo y elegí cuáles aplicar como tránsito.`
          : "No se encontraron productos coincidentes en el archivo."
      );
    } catch (err) {
      setUploadStatus("No se pudo leer el archivo de tránsito.");
    }
  };

  const toggleTransitoPreview = (id) => {
    setTransitoPreview((prev) => prev.map((p) => (p.id === id ? { ...p, aplicar: !p.aplicar } : p)));
  };

  const marcarTodosTransito = (valor) => {
    setTransitoPreview((prev) => prev.map((p) => ({ ...p, aplicar: valor })));
  };

  const aplicarTransitoSeleccionados = () => {
    const seleccionados = transitoPreview.filter((p) => p.aplicar);
    if (seleccionados.length === 0) return;
    const map = {};
    seleccionados.forEach((p) => {
      map[p.id] = p.nuevo;
    });
    setRows((prev) => prev.map((r) => (map[r.id] !== undefined ? { ...r, transito: map[r.id] } : r)));
    setUploadStatus(`Tránsito actualizado en ${seleccionados.length} producto(s).`);
    setTransitoPreview((prev) => prev.filter((p) => !p.aplicar));
  };

  const computed = rows.map((r) => computeRow(r, diasInvConfig));
  const inconsistencias = useMemo(() => chequearConsistencia(rows), [rows]);

  const updateUnidadBase = (id, value) => {
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, unidadBase: value } : r)));
  };

  const guardarSemana = () => {
    const items = computed.map((r) => ({
      producto: r.producto,
      proveedor: r.proveedor,
      dia: r.dia,
      unidad: r.unidad,
      pedir: r.pedir,
      // Por defecto se sugieren para concatenar los que sí tenían cantidad a pedir;
      // el usuario elige manualmente cuáles realmente aplican.
      selected: r.pedir > 0,
    }));
    const snapshot = { id: Date.now(), fecha: fechaPedido, items };
    setHistorialStatus("Guardando semana…");
    saveSnapshot(snapshot)
      .then((data) => {
        setHistorial(data);
        setHistorialStatus(`Semana del ${fechaPedido} guardada en el historial (${items.filter((i) => i.pedir > 0).length} productos con pedido).`);
      })
      .catch((err) => {
        console.error(err);
        setHistorialStatus("No se pudo guardar en Firebase — revisá la conexión.");
      });
  };

  const toggleItemSeleccion = (snapshotId, idx) => {
    setHistorial((prev) =>
      prev.map((snap) => {
        if (snap.id !== snapshotId) return snap;
        const items = snap.items.map((it, i) => (i === idx ? { ...it, selected: !it.selected } : it));
        updateSnapshot(snapshotId, items).catch((err) => console.error(err));
        return { ...snap, items };
      })
    );
  };

  const eliminarSemana = (snapshotId, fecha) => {
    if (!window.confirm(`¿Eliminar el pedido guardado de la semana del ${fecha}? Esta acción no se puede deshacer.`)) return;
    setHistorialStatus("Eliminando…");
    deleteSnapshot(snapshotId)
      .then((data) => {
        setHistorial(data);
        setHistorialStatus(`Semana del ${fecha} eliminada del historial.`);
      })
      .catch((err) => {
        console.error(err);
        setHistorialStatus("No se pudo eliminar — revisá la conexión.");
      });
  };

  const aplicarComoTransito = (snapshotId) => {
    const snap = historial.find((h) => h.id === snapshotId);
    if (!snap) return;
    const map = {};
    snap.items.forEach((it) => {
      if (it.selected && it.pedir > 0) {
        const k = itemKey(it.producto, it.proveedor, it.dia);
        map[k] = (map[k] || 0) + it.pedir;
      }
    });
    if (Object.keys(map).length === 0) {
      setHistorialStatus("No hay productos seleccionados para concatenar en esa semana.");
      return;
    }
    let count = 0;
    setRows((prev) =>
      prev.map((r) => {
        const k = itemKey(r.producto, r.proveedor, r.dia);
        if (map[k] !== undefined) {
          count++;
          return { ...r, transito: r.transito + map[k] };
        }
        return r;
      })
    );
    setHistorialStatus(`Se sumó al tránsito de ${count} productos seleccionados de la semana del ${snap.fecha}.`);
  };

  const filtered = computed.filter((r) => {
    if (dia !== "Todos" && !diaProveedores[dia]?.has(r.proveedor)) return false;
    if (tipo !== "Todos" && r.tipo !== tipo) return false;
    if (proveedor !== "Todos" && r.proveedor !== proveedor) return false;
    return true;
  });
  const provOrderIndex = useMemo(() => {
    const map = {};
    proveedoresList.forEach((p, idx) => { map[p.nombre] = idx; });
    return map;
  }, [proveedoresList]);
  const filteredGrouped = useMemo(
    () => [...filtered].sort((a, b) => (provOrderIndex[a.proveedor] ?? 999) - (provOrderIndex[b.proveedor] ?? 999)),
    [filtered, provOrderIndex]
  );
  const colCount = columnasCompactas ? 7 : 14;
  const totalPedir = filtered.reduce((s, r) => s + (r.pedir > 0 ? 1 : 0), 0);

  const uploadChecklist = [
    { label: "Consumo semanal", done: !!consumoUltima },
    ...DIAS_PEDIDO.map((d) => ({ label: `Existencia ${d}`, done: !!existenciaUltima[d] })),
  ];
  const uploadsListos = uploadChecklist.filter((u) => u.done).length;

  const pedirCompletados = filtered.filter((r) => r.pedirTocado).length;
  const pedirProgreso = filtered.length ? Math.round((pedirCompletados / filtered.length) * 100) : 0;

  const resumenTodos = computed
    .filter((r) => resumenDia === "Todos" || diaProveedores[resumenDia]?.has(r.proveedor))
    .map((r) => ({
      producto: r.producto,
      proveedor: r.proveedor,
      dia: r.dia,
      cantidad: r.pedir,
      unidad: r.unidad,
      fecha: formatDate(nextDateForWeekday(fechaPedido, schedule[`${r.proveedor}|${r.dia}`] || "Lunes")),
    }));
  const resumenItems = resumenTodos.filter((i) => i.cantidad > 0);

  const descargarCSV = () => {
    const lista = resumenMostrarTodos ? resumenTodos : resumenItems;
    const header = "Producto,Proveedor,Día de pedido,Fecha de entrega,Cantidad,Unidad\n";
    const body = lista
      .map((i) => `"${i.producto}","${i.proveedor}","${i.dia}","${i.fecha}",${i.cantidad},"${i.unidad}"`)
      .join("\n");
    const blob = new Blob([header + body], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `pedido_${resumenDia}_${fechaPedido}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const descargarExcel = () => {
    const lista = resumenMostrarTodos ? resumenTodos : resumenItems;
    const data = lista.map((i) => ({
      Producto: i.producto,
      Proveedor: i.proveedor,
      "Día de pedido": i.dia,
      "Fecha de entrega": i.fecha,
      Cantidad: i.cantidad,
      Unidad: i.unidad,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 28 }, { wch: 20 }, { wch: 14 }, { wch: 16 }, { wch: 10 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pedido");
    XLSX.writeFile(wb, `pedido_${resumenDia}_${fechaPedido}.xlsx`);
  };

  const vars = isDark ? DARK_VARS : LIGHT_VARS;

  if (vistaResumen) {
    const listaActual = resumenMostrarTodos ? resumenTodos : resumenItems;
    const porProveedor = {};
    listaActual.forEach((i) => {
      if (!porProveedor[i.proveedor]) porProveedor[i.proveedor] = [];
      porProveedor[i.proveedor].push(i);
    });
    return (
      <div className="pp-shell" style={{ ...vars, background: "var(--bg)", padding: 20, fontFamily: "Inter, system-ui, sans-serif", color: "var(--text-primary)", borderRadius: 12 }}>
        <StyleBlock />
        <div style={{ marginBottom: 14 }}>
          <Logo subtitle="RESUMEN DE PEDIDO" />
        </div>
        <div className="pp-config-inline" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 13, color: "var(--text-secondary)" }}>
              Fecha de pedido: {fechaPedido} &middot; {resumenItems.length} productos a pedir
              {resumenMostrarTodos && ` · ${resumenTodos.length} en total`}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={() => setVistaResumen(false)} style={buttonStyle}><ArrowLeft size={13} /> Volver a la tabla</button>
            <button onClick={descargarCSV} style={buttonStyle}><Download size={13} /> Descargar CSV</button>
            <button onClick={descargarExcel} style={buttonStyle}><Download size={13} /> Descargar Excel</button>
            <button onClick={() => window.print()} style={{ ...buttonStyle, background: "var(--header-bg)", color: "var(--header-text)", borderColor: "var(--header-bg)" }}><Printer size={13} /> Imprimir</button>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 18, marginBottom: 18, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 500, color: "var(--text-secondary)" }}>Mostrando el pedido de:</span>
            <select value={resumenDia} onChange={(e) => setResumenDia(e.target.value)} style={selectStyle}>
              <option value="Todos">Todos los días</option>
              {DIAS_PEDIDO.map((d) => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>
          </div>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "var(--text-secondary)", cursor: "pointer" }}>
            <input type="checkbox" checked={resumenMostrarTodos} onChange={(e) => setResumenMostrarTodos(e.target.checked)} />
            Mostrar todos los productos del día (incluye los que no hace falta pedir)
          </label>
        </div>

        {listaActual.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--text-secondary)" }}>
            {resumenMostrarTodos
              ? `No hay productos asignados a ${resumenDia === "Todos" ? "ningún día" : resumenDia} todavía.`
              : `Ningún producto tiene cantidad a pedir para ${resumenDia === "Todos" ? "ningún día" : resumenDia} todavía.`}
          </div>
        ) : (
          Object.entries(porProveedor).map(([prov, items]) => (
            <div key={prov} style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 10, marginBottom: 14, overflow: "hidden", boxShadow: "var(--shadow)" }}>
              <div style={{ padding: "10px 14px", background: "var(--header-bg)", color: "var(--header-text)", fontSize: 13, fontWeight: 500 }}>
                {prov} &middot; entrega {items[0].fecha}
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ textAlign: "left" }}>
                    {["Producto", "Cantidad", "Unidad", ...(resumenMostrarTodos ? ["Estado"] : [])].map((h) => (
                      <th key={h} style={{ padding: "8px 14px", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.3, color: "var(--text-secondary)", fontWeight: 500 }}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {items.map((i, idx) => (
                    <tr key={idx} style={{ borderTop: "1px solid var(--border)", opacity: i.cantidad > 0 ? 1 : 0.55 }}>
                      <td style={{ padding: "8px 14px", fontWeight: 500 }}>{i.producto}</td>
                      <td style={{ padding: "8px 14px", fontFamily: "var(--mono)", fontWeight: i.cantidad > 0 ? 600 : 400, color: i.cantidad > 0 ? "var(--danger)" : "var(--text-secondary)" }}>
                        {i.cantidad}
                      </td>
                      <td style={{ padding: "8px 14px", color: "var(--text-secondary)" }}>{i.unidad}</td>
                      {resumenMostrarTodos && (
                        <td style={{ padding: "8px 14px", fontSize: 11, color: i.cantidad > 0 ? "var(--danger)" : "var(--ok)" }}>
                          {i.cantidad > 0 ? "a pedir" : "cubierto"}
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))
        )}
      </div>
    );
  }

  return (
    <div className="pp-shell" style={{ ...vars, background: "var(--bg)", padding: 22, fontFamily: "Inter, system-ui, sans-serif", color: "var(--text-primary)", borderRadius: 14 }}>
      <StyleBlock />
      <div style={{ marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <Logo />
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {estadoListo && (
            <span
              style={{
                fontSize: 11,
                color: guardadoStatus === "no se pudo guardar — revisá tu conexión" ? "var(--danger, #d33)" : "var(--text-secondary)",
              }}
            >
              {guardadoStatus === "guardando…" ? "guardando…" : guardadoStatus === "guardado ✓" ? "✓ guardado" : guardadoStatus}
            </span>
          )}
          <ThemeToggle themeMode={themeMode} setThemeMode={setThemeMode} />
        </div>
      </div>

      <div className="pp-tabs" style={{ borderBottom: "1px solid var(--border)", marginBottom: 18 }}>
        {TABS.map(({ key, label, icon: Icon }) => (
          <button key={key} onClick={() => setActiveTab(key)} className={`pp-tab ${activeTab === key ? "active" : ""}`}>
            <Icon size={14} />
            <span className="pp-tab-label">{label}</span>
          </button>
        ))}
      </div>

      {activeTab === "resumen" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <div className="pp-kpi-row" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <KpiCard icon={<Boxes size={15} />} label="Productos en vista" value={filtered.length} />
            <KpiCard icon={<TriangleAlert size={15} />} label="Requieren pedido" value={totalPedir} tone="danger" />
            <KpiCard icon={<PackageCheck size={15} />} label="Proveedores" value={todosProveedores.length} />
            <KpiCard
              icon={uploadsListos === uploadChecklist.length ? <CheckCircle2 size={15} /> : <Circle size={15} />}
              label="Archivos cargados"
              value={`${uploadsListos}/${uploadChecklist.length}`}
              tone={uploadsListos === uploadChecklist.length ? undefined : "danger"}
            />
          </div>

          <div style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 12, boxShadow: "var(--shadow)", overflow: "hidden" }}>
            <div style={{ padding: "12px 16px", background: "var(--header-bg)", color: "var(--header-text)", fontSize: 13, fontWeight: 500 }}>
              Productos más urgentes
            </div>
            {resumenItems.length === 0 ? (
              <div style={{ padding: 16, fontSize: 13, color: "var(--text-secondary)" }}>Ningún producto requiere pedido en este momento.</div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <tbody>
                  {[...resumenItems]
                    .sort((a, b) => b.cantidad - a.cantidad)
                    .slice(0, 8)
                    .map((i, idx) => (
                      <tr key={idx} style={{ borderTop: idx === 0 ? "none" : "1px solid var(--border)" }}>
                        <td style={{ padding: "9px 16px", fontWeight: 500 }}>{i.producto}</td>
                        <td style={{ padding: "9px 16px", color: "var(--text-secondary)" }}>{i.proveedor}</td>
                        <td style={{ padding: "9px 16px", color: "var(--text-secondary)" }}>{i.fecha}</td>
                        <td style={{ padding: "9px 16px", fontFamily: "var(--mono)", fontWeight: 600, color: "var(--danger)", textAlign: "right" }}>
                          {i.cantidad} {i.unidad}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            )}
            {resumenItems.length > 8 && (
              <div style={{ padding: "8px 16px", fontSize: 11, color: "var(--text-secondary)" }}>
                y {resumenItems.length - 8} más — vé a la pestaña Pedidos para verlos todos.
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === "config" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <Accordion title="1. Cargar excels de consumo y existencia" isOpen={sectionOpen.uploads} onToggle={() => toggleSection("uploads")}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <span style={{ fontSize: 12, fontWeight: 600 }}>Archivos cargados</span>
              <span style={{ fontFamily: "var(--mono)", fontSize: 12, fontWeight: 600, color: uploadsListos === uploadChecklist.length ? "var(--ok)" : "var(--text-secondary)" }}>
                {uploadsListos}/{uploadChecklist.length}
              </span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 }}>
              {uploadChecklist.map((u) => (
                <div key={u.label} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: u.done ? "var(--text-primary)" : "var(--text-secondary)" }}>
                  {u.done ? <CheckCircle2 size={15} color="var(--ok)" /> : <Circle size={15} color="var(--border)" />}
                  {u.label}
                </div>
              ))}
            </div>
            <div className="pp-config-inline" style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
              <div>
                <div style={{ fontSize: 12, marginBottom: 4 }}>Consumo semanal (elegís qué productos aplicar)</div>
                <input type="file" accept=".xlsx,.xls" onChange={(e) => handleConsumoUpload(e.target.files[0])} style={fileInputStyle} />
              </div>
              <div>
                <div style={{ fontSize: 12, marginBottom: 4 }}>Existencia para el día (elegís qué productos aplicar)</div>
                <div style={{ display: "flex", gap: 8 }}>
                  <select value={existenciaDia} onChange={(e) => setExistenciaDia(e.target.value)} style={selectStyle}>
                    {DIAS_PEDIDO.map((d) => (
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                  <input type="file" accept=".xlsx,.xls" onChange={(e) => handleExistenciaUpload(e.target.files[0])} style={fileInputStyle} />
                </div>
              </div>
            </div>
            {uploadStatus && <div style={{ fontSize: 12, color: "var(--ok)", marginTop: 8 }}>{uploadStatus}</div>}
            {consumoPreview.length > 0 && (
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, marginTop: 14 }}>Vista previa — Consumo semanal ({consumoArchivo})</div>
                <PreviewSeleccion
                  preview={consumoPreview}
                  onToggle={toggleConsumoPreview}
                  onMarcarTodos={marcarTodosConsumo}
                  onAplicar={aplicarConsumoSeleccionados}
                  aplicarLabel="Aplicar seleccionados a Consumo semanal"
                  colActualLabel="Consumo actual"
                />
              </div>
            )}
            {existenciaPreview.length > 0 && (
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, marginTop: 14 }}>
                  Vista previa — Existencia {existenciaDia} ({existenciaArchivo})
                </div>
                <PreviewSeleccion
                  preview={existenciaPreview}
                  onToggle={toggleExistenciaPreview}
                  onMarcarTodos={marcarTodosExistencia}
                  onAplicar={aplicarExistenciaSeleccionados}
                  aplicarLabel="Aplicar seleccionados a Existencia"
                  colActualLabel="Existencia actual"
                />
              </div>
            )}
            <div style={{ display: "flex", gap: 14, marginTop: 8, flexWrap: "wrap" }}>
              {DIAS_PEDIDO.map((d) => {
                const u = existenciaUltima[d];
                return (
                  <span key={d} style={{ fontSize: 11, color: "var(--text-secondary)", display: "flex", alignItems: "center", gap: 4 }}>
                    <span style={{ width: 6, height: 6, borderRadius: "50%", background: u ? "var(--ok)" : "var(--track)", display: "inline-block", border: u ? "none" : "1px solid var(--border)" }} />
                    {d}: {u ? `${u.count} productos · ${u.archivo}` : "sin datos aún"}
                  </span>
                );
              })}
            </div>
          </Accordion>

          <Accordion title="2. Cargar tránsito desde Excel (Inventario Final, selección manual)" isOpen={sectionOpen.transito} onToggle={() => toggleSection("transito")}>
            <div style={{ fontSize: 12, color: "var(--text-secondary)", marginBottom: 10 }}>
              Subí el reporte (columna A = nombre del producto, columna Q = Inventario Final). La app va a mostrar
              una lista de coincidencias para que elijas producto por producto cuáles aplicar como Tránsito — no se
              aplica nada automáticamente.
            </div>
            <input type="file" accept=".xlsx,.xls" onChange={(e) => handleTransitoUpload(e.target.files[0])} style={fileInputStyle} />
            {uploadStatus && transitoArchivo && <div style={{ fontSize: 12, color: "var(--ok)", marginTop: 8 }}>{uploadStatus}</div>}
            <PreviewSeleccion
              preview={transitoPreview}
              onToggle={toggleTransitoPreview}
              onMarcarTodos={marcarTodosTransito}
              onAplicar={aplicarTransitoSeleccionados}
              aplicarLabel="Aplicar seleccionados a Tránsito"
              colActualLabel="Tránsito actual"
            />
          </Accordion>

          <Accordion title="3. Proveedores — agregar, eliminar y asignar días" isOpen={sectionOpen.proveedores} onToggle={() => toggleSection("proveedores")}>
            <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap", alignItems: "center" }}>
              <input
                type="text"
                placeholder="Nombre del proveedor nuevo"
                value={nuevoProveedor.nombre}
                onChange={(e) => setNuevoProveedor((p) => ({ ...p, nombre: e.target.value }))}
                style={{ ...textInputStyle, width: 200 }}
              />
              <select value={nuevoProveedor.tipo} onChange={(e) => setNuevoProveedor((p) => ({ ...p, tipo: e.target.value }))} style={selectStyle}>
                <option value="externo">externo</option>
                <option value="interno">interno</option>
              </select>
              <button onClick={agregarProveedor} style={buttonStyle}><Plus size={13} /> Agregar proveedor</button>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>
                    <th style={{ padding: "4px 6px", color: "var(--text-secondary)", fontWeight: 500 }}>Orden</th>
                    <th style={{ textAlign: "left", padding: "4px 10px", color: "var(--text-secondary)", fontWeight: 500 }}>Proveedor</th>
                    <th style={{ padding: "4px 10px", color: "var(--text-secondary)", fontWeight: 500 }}>Tipo</th>
                    {DIAS_PEDIDO.map((d) => (
                      <th key={d} style={{ padding: "4px 10px", color: "var(--text-secondary)", fontWeight: 500 }}>{d}</th>
                    ))}
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {proveedoresList.map(({ nombre: prov, tipo: tp }, idx) => (
                    <tr key={prov} style={{ borderTop: "1px solid var(--border)" }}>
                      <td style={{ padding: "4px 6px" }}>
                        <div style={{ display: "flex", gap: 2 }}>
                          <button
                            onClick={() => moveProveedor(prov, -1)}
                            disabled={idx === 0}
                            title="Subir"
                            style={{ ...miniButtonStyle, opacity: idx === 0 ? 0.35 : 1, cursor: idx === 0 ? "default" : "pointer" }}
                          >
                            <ChevronUp size={12} />
                          </button>
                          <button
                            onClick={() => moveProveedor(prov, 1)}
                            disabled={idx === proveedoresList.length - 1}
                            title="Bajar"
                            style={{ ...miniButtonStyle, opacity: idx === proveedoresList.length - 1 ? 0.35 : 1, cursor: idx === proveedoresList.length - 1 ? "default" : "pointer" }}
                          >
                            <ChevronDown size={12} />
                          </button>
                        </div>
                      </td>
                      <td style={{ padding: "4px 10px" }}>{prov}</td>
                      <td style={{ padding: "4px 10px", textAlign: "center" }}>
                        <select value={tp} onChange={(e) => updateProveedorTipo(prov, e.target.value)} style={{ fontSize: 11.5, border: "1px solid var(--border)", borderRadius: 6, padding: "2px 4px", color: "var(--text-primary)", background: "var(--card)" }}>
                          <option value="externo">externo</option>
                          <option value="interno">interno</option>
                        </select>
                      </td>
                      {DIAS_PEDIDO.map((d) => (
                        <td key={d} style={{ padding: "4px 10px", textAlign: "center" }}>
                          <input type="checkbox" checked={diaProveedores[d].has(prov)} onChange={() => toggleProveedorDia(prov, d)} />
                        </td>
                      ))}
                      <td style={{ padding: "4px 10px" }}>
                        <button onClick={() => eliminarProveedor(prov)} style={dangerLinkStyle}><Trash2 size={12} /> eliminar</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: "var(--text-secondary)", marginTop: 6 }}>
              Eliminar un proveedor quita también sus productos de la tabla. Los checkboxes definen en qué día de pedido aparece.
            </div>
          </Accordion>

          <Accordion title="4. Días de inventario objetivo por proveedor y día de pedido" isOpen={sectionOpen.diasInv} onToggle={() => toggleSection("diasInv")}>
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
              <select value={nuevaCombo.proveedor} onChange={(e) => setNuevaCombo((c) => ({ ...c, proveedor: e.target.value }))} style={selectStyle}>
                <option value="">Elegir proveedor…</option>
                {todosProveedores.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
              <select value={nuevaCombo.dia} onChange={(e) => setNuevaCombo((c) => ({ ...c, dia: e.target.value }))} style={selectStyle}>
                {DIAS_PEDIDO.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
              <button onClick={agregarCombo} style={buttonStyle}><Plus size={13} /> Agregar combinación</button>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
              {proveedorDiaKeys.map((key) => {
                const [prov, d] = key.split("|");
                const rowMatch = rows.find((r) => `${r.proveedor}|${r.dia}` === key);
                const original = rowMatch?.diasInv ?? 0;
                return (
                  <div key={key} style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid var(--border)", borderRadius: 8, padding: "6px 10px" }}>
                    <span style={{ fontSize: 12 }}>{prov} · {d}</span>
                    <input
                      type="number"
                      defaultValue={original}
                      onChange={(e) => updateDiasInv(key, e.target.value)}
                      style={{ width: 48, fontFamily: "var(--mono)", fontSize: 12, border: "1px solid var(--border)", borderRadius: 6, padding: "2px 4px" }}
                    />
                    <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>días</span>
                    <button onClick={() => eliminarCombo(key)} title="Quitar combinación" style={{ ...miniButtonStyle, padding: "2px 5px" }}><Trash2 size={11} /></button>
                  </div>
                );
              })}
            </div>
          </Accordion>

          <Accordion title="5. Calendario de despacho por proveedor y día de pedido" isOpen={sectionOpen.calendario} onToggle={() => toggleSection("calendario")}>
            <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10 }}>
              <span style={{ fontSize: 12 }}>Fecha del pedido</span>
              <input type="date" value={fechaPedido} onChange={(e) => setFechaPedido(e.target.value)} style={selectStyle} />
            </div>
            <div style={{ fontSize: 11, color: "var(--text-secondary)", marginBottom: 8 }}>
              Cada combinación de proveedor + día de pedido tiene su propio día de entrega (por ejemplo, Ristreto puede entregar Jueves lo pedido el Domingo, y Martes lo pedido el Martes).
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
              <select value={nuevaCombo.proveedor} onChange={(e) => setNuevaCombo((c) => ({ ...c, proveedor: e.target.value }))} style={selectStyle}>
                <option value="">Elegir proveedor…</option>
                {todosProveedores.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
              <select value={nuevaCombo.dia} onChange={(e) => setNuevaCombo((c) => ({ ...c, dia: e.target.value }))} style={selectStyle}>
                {DIAS_PEDIDO.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
              <button onClick={agregarCombo} style={buttonStyle}><Plus size={13} /> Agregar combinación</button>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
              {proveedorDiaKeys.map((key) => {
                const [prov, d] = key.split("|");
                return (
                  <div key={key} style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid var(--border)", borderRadius: 8, padding: "6px 10px" }}>
                    <span style={{ fontSize: 12 }}>{prov} · {d}</span>
                    <select value={schedule[key] || "Lunes"} onChange={(e) => updateSchedule(key, e.target.value)} style={{ fontSize: 12, border: "1px solid var(--border)", borderRadius: 6, padding: "2px 4px" }}>
                      {WEEKDAYS.map((w) => (
                        <option key={w} value={w}>{w}</option>
                      ))}
                    </select>
                    <button onClick={() => eliminarCombo(key)} title="Quitar combinación" style={{ ...miniButtonStyle, padding: "2px 5px" }}><Trash2 size={11} /></button>
                  </div>
                );
              })}
            </div>
          </Accordion>

          <Accordion title="6. Consistencia de unidades" isOpen={sectionOpen.consistencia} onToggle={() => toggleSection("consistencia")}>
            {inconsistencias.length === 0 ? (
              <div style={{ fontSize: 12, color: "var(--ok)" }}>Todos los productos usan la misma unidad base en todas sus filas.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {inconsistencias.map((p) => (
                  <div key={p.nombre} style={{ fontSize: 12, color: "var(--danger)", background: "var(--danger-soft-bg)", border: "1px solid var(--danger-soft-border)", borderRadius: 8, padding: "6px 10px" }}>
                    "{p.nombre}" tiene filas en distintas unidades base ({[...p.unidades].join(" y ")}) — revisá que no se estén sumando libras con cajas o bolsas con unidades.
                  </div>
                ))}
              </div>
            )}
            <div style={{ fontSize: 11, color: "var(--text-secondary)", marginTop: 6 }}>
              Cada producto tiene una unidad base (columna "u. base" en la tabla) que se usa para comparar consumo, existencia y tránsito de forma justa entre proveedores. La cantidad a pedir siempre se muestra convertida a la presentación de compra (caja, bolsa, paquete).
            </div>
          </Accordion>
        </div>
      )}

      {activeTab === "historial" && (
        <div style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 12, boxShadow: "var(--shadow)", padding: 16 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={guardarSemana} style={buttonStyle}><PackageCheck size={13} /> Guardar pedido de esta semana en el historial</button>
            <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>{dbStatus}</span>
          </div>
          {historialStatus && <div style={{ fontSize: 12, color: "var(--ok)", marginBottom: 8 }}>{historialStatus}</div>}
          {historial.length === 0 && <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>Todavía no hay semanas guardadas.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {historial.map((snap) => {
              const pedidos = snap.items.filter((i) => i.pedir > 0);
              const seleccionados = pedidos.filter((i) => i.selected).length;
              return (
                <div key={snap.id} style={{ border: "1px solid var(--border)", borderRadius: 8, padding: "8px 12px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                    <span style={{ fontSize: 12, fontWeight: 500 }}>
                      Semana del {snap.fecha} &middot; {pedidos.length} productos pedidos &middot; {seleccionados} seleccionados
                    </span>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button onClick={() => aplicarComoTransito(snap.id)} style={buttonStyle}>Concatenar seleccionados como tránsito</button>
                      <button onClick={() => eliminarSemana(snap.id, snap.fecha)} style={dangerLinkStyle} title="Eliminar esta semana del historial">
                        <Trash2 size={13} /> Eliminar
                      </button>
                    </div>
                  </div>
                  {pedidos.length > 0 && (
                    <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 4 }}>
                      {snap.items.map((it, idx) =>
                        it.pedir > 0 ? (
                          <label key={idx} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
                            <input type="checkbox" checked={!!it.selected} onChange={() => toggleItemSeleccion(snap.id, idx)} />
                            <span>{it.producto} <span style={{ color: "var(--text-secondary)" }}>({it.proveedor}, {it.dia})</span> — {it.pedir} {it.unidad}</span>
                          </label>
                        ) : null
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          <div style={{ fontSize: 11, color: "var(--text-secondary)", marginTop: 10 }}>
            Marcá solo los productos que realmente siguen en tránsito esa semana — no todos los pedidos anteriores aplican. "Concatenar seleccionados" suma esa cantidad al tránsito de los productos que coincidan por nombre, proveedor y día.
          </div>
        </div>
      )}

      {activeTab === "pedidos" && (
        <>
      <div style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 10, boxShadow: "var(--shadow)", padding: "10px 14px", marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <span style={{ fontSize: 12, fontWeight: 600 }}>Pedido completado</span>
          <span style={{ fontFamily: "var(--mono)", fontSize: 12, fontWeight: 600, color: pedirProgreso === 100 ? "var(--ok)" : "var(--text-secondary)" }}>
            {pedirCompletados}/{filtered.length} &middot; {pedirProgreso}%
          </span>
        </div>
        <div style={{ width: "100%", height: 8, borderRadius: 4, background: "var(--track)", overflow: "hidden" }}>
          <div style={{ width: `${pedirProgreso}%`, height: "100%", background: pedirProgreso === 100 ? "var(--ok)" : "var(--warn)", borderRadius: 4, transition: "width 0.2s" }} />
        </div>
        <div style={{ fontSize: 10.5, color: "var(--text-secondary)", marginTop: 4 }}>
          Cuenta como completado cada producto donde ya tocaste el campo "Pedir" en la vista actual.
        </div>
      </div>

      <div className="pp-filters" style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
        <Select label="Día de pedido" value={dia} onChange={setDia} options={diaOptions} />
        <Select label="Tipo" value={tipo} onChange={setTipo} options={tipoOptions} />
        <Select label="Proveedor" value={proveedor} onChange={setProveedor} options={proveedoresFiltro} />
        {proveedor !== "Todos" && (
          <button onClick={() => agregarProducto(proveedor, dia === "Todos" ? "Domingo" : dia)} style={buttonStyle}>
            <Plus size={13} /> Agregar producto a {proveedor}
          </button>
        )}
        <button
          onClick={() => {
            setResumenDia(dia);
            setVistaResumen(true);
          }}
          style={{ ...buttonStyle, marginLeft: "auto", background: "var(--header-bg)", color: "var(--header-text)", borderColor: "var(--header-bg)" }}
        >
          <Printer size={13} /> Ver resumen para descargar / imprimir
        </button>
        <button onClick={() => setColumnasCompactas((v) => !v)} style={buttonStyle}>
          {columnasCompactas ? "Ver todas las columnas" : "Vista compacta"}
        </button>
      </div>

      <div className="pp-table-wrap" style={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 12, overflow: "hidden", boxShadow: "var(--shadow)" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "var(--header-bg)", textAlign: "left" }}>
                {(columnasCompactas
                  ? ["Producto", "Proveedor", "Cobertura", "Existencia", "Tránsito", "Pedir", ""]
                  : ["Producto", "Proveedor", "Tipo", "Unidad compra", "U. base", "Consumo/día", "Días obj.", "Cobertura", "Existencia", "Tránsito", "Proyección", "Pedir", "Fecha despacho", ""]
                ).map((h) => (
                  <th key={h} style={{ padding: "11px 12px", fontWeight: 500, color: "var(--header-text)", whiteSpace: "nowrap", fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.5 }}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredGrouped.map((r, idx) => {
                const despachoDate = nextDateForWeekday(fechaPedido, schedule[`${r.proveedor}|${r.dia}`] || "Lunes");
                const esNuevoGrupo = idx === 0 || filteredGrouped[idx - 1].proveedor !== r.proveedor;
                return (
                  <Fragment key={r.id}>
                    {esNuevoGrupo && (
                      <tr>
                        <td colSpan={colCount} style={{ padding: "8px 12px", background: "var(--track)", fontSize: 11.5, fontWeight: 600, color: "var(--text-secondary)", textTransform: "uppercase", letterSpacing: 0.4 }}>
                          {r.proveedor}
                        </td>
                      </tr>
                    )}
                    <tr style={{ borderTop: "1px solid var(--border)", borderLeft: `3px solid ${r.tipo === "externo" ? "var(--accent-ext)" : "var(--accent-int)"}` }}>
                    <td style={{ padding: "6px 12px", minWidth: 180 }}>
                      <input type="text" value={r.producto} onChange={(e) => updateText(r.id, "producto", e.target.value)} style={textInputStyle} />
                    </td>
                    <td style={{ padding: "10px 12px", color: "var(--text-secondary)" }}>{r.proveedor}</td>
                    {!columnasCompactas && (
                      <td style={{ padding: "10px 12px" }}>
                        <span
                          style={{
                            fontSize: 10.5,
                            fontWeight: 500,
                            padding: "2px 8px",
                            borderRadius: 20,
                            background: r.tipo === "externo" ? "#F0E6DC" : "#F5EBC8",
                            color: r.tipo === "externo" ? "var(--accent-ext)" : "var(--accent-int)",
                          }}
                        >
                          {r.tipo}
                        </span>
                      </td>
                    )}
                    {!columnasCompactas && (
                      <td style={{ padding: "6px 12px", minWidth: 120 }}>
                        <input type="text" value={r.unidad} onChange={(e) => updateText(r.id, "unidad", e.target.value)} style={textInputStyle} />
                      </td>
                    )}
                    {!columnasCompactas && (
                      <td style={{ padding: "6px 12px" }}>
                        <select value={r.unidadBase} onChange={(e) => updateUnidadBase(r.id, e.target.value)} style={{ ...selectStyle, padding: "4px 6px", fontSize: 12 }}>
                          {UNIDADES_BASE.map((u) => (
                            <option key={u} value={u}>{u}</option>
                          ))}
                        </select>
                      </td>
                    )}
                    {!columnasCompactas && <td style={{ padding: "10px 12px", fontFamily: "var(--mono)" }}>{r.diario.toFixed(2)}</td>}
                    {!columnasCompactas && <td style={{ padding: "10px 12px", fontFamily: "var(--mono)" }}>{r.diasInv}</td>}
                    <td style={{ padding: "10px 12px" }}>
                      <Bar cobertura={r.cobertura} objetivo={r.diasInv} />
                    </td>
                    <td style={{ padding: "10px 12px" }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                        <span style={{ fontFamily: "var(--mono)", fontSize: 13 }}>{r.existencia}</span>
                        <span style={{ fontSize: 9, color: "var(--accent-int)", textTransform: "uppercase", letterSpacing: 0.3 }}>
                          auto · {r.unidadBase}
                        </span>
                      </div>
                    </td>
                    <td style={{ padding: "6px 12px" }}>
                      <ConvertibleInput
                        value={r.transito}
                        unidadBase={r.unidadBase}
                        unidadCompra={r.unidad}
                        factor={r.factor}
                        onChange={(v) => update(r.id, "transito", v)}
                      />
                    </td>
                    {!columnasCompactas && (
                      <td style={{ padding: "10px 12px", fontFamily: "var(--mono)", color: r.proyeccion < 0 ? "var(--text-secondary)" : "var(--text-primary)" }}>
                        {r.proyeccion.toFixed(1)}
                      </td>
                    )}
                    <td style={{ padding: "6px 12px" }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                          <input
                            type="number"
                            value={r.pedir}
                            onChange={(e) => update(r.id, "pedir", e.target.value)}
                            style={{ ...inputStyle, fontWeight: 500, color: r.pedir > 0 ? "var(--danger)" : "var(--ok)" }}
                          />
                          {r.pedirTocado ? <CheckCircle2 size={14} color="var(--ok)" /> : <Circle size={14} color="var(--border)" />}
                        </div>
                        {r.pedirSugerido !== r.pedir && (
                          <button onClick={() => update(r.id, "pedir", r.pedirSugerido)} style={{ ...miniButtonStyle, fontSize: 10 }}>
                            usar sugerido: {r.pedirSugerido}
                          </button>
                        )}
                      </div>
                    </td>
                    {!columnasCompactas && (
                      <td style={{ padding: "10px 12px", color: "var(--text-secondary)", whiteSpace: "nowrap" }}>{formatDate(despachoDate)}</td>
                    )}
                    <td style={{ padding: "10px 12px" }}>
                      <button onClick={() => eliminarFila(r.id)} style={dangerLinkStyle}><Trash2 size={12} /></button>
                    </td>
                  </tr>
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="pp-cards">
        {filteredGrouped.map((r, idx) => {
          const esNuevoGrupo = idx === 0 || filteredGrouped[idx - 1].proveedor !== r.proveedor;
          return (
            <Fragment key={r.id}>
              {esNuevoGrupo && (
                <div style={{ fontSize: 11.5, fontWeight: 600, color: "var(--text-secondary)", textTransform: "uppercase", letterSpacing: 0.4, padding: "6px 4px 2px" }}>
                  {r.proveedor}
                </div>
              )}
              <ProductCard
                r={r}
                despachoLabel={formatDate(nextDateForWeekday(fechaPedido, schedule[`${r.proveedor}|${r.dia}`] || "Lunes"))}
                updateText={updateText}
                update={update}
                updateUnidadBase={updateUnidadBase}
                eliminarFila={eliminarFila}
              />
            </Fragment>
          );
        })}
      </div>

      <div style={{ marginTop: 10, fontSize: 12, color: "var(--text-secondary)" }}>
        Existencia se llena sola desde el excel que cargues arriba (elegí el día correcto antes de subirlo) — ya no se edita a mano. Tránsito y la cantidad a pedir sí son manuales. Agregá o eliminá proveedores y productos desde el panel de configuración, y usá "Guardar pedido de esta semana" cada vez que cierres un pedido.
      </div>
        </>
      )}
    </div>
  );
}

function Select({ label, value, onChange, options }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 11, color: "#726e63" }}>
      {label}
      <select value={value} onChange={(e) => onChange(e.target.value)} style={selectStyle}>
        {options.map((o) => (
          <option key={o} value={o}>{o}</option>
        ))}
      </select>
    </label>
  );
}

const sectionLabel = {
  fontSize: 11,
  fontWeight: 600,
  color: "var(--text-secondary)",
  textTransform: "uppercase",
  letterSpacing: 0.6,
  marginBottom: 10,
  paddingLeft: 9,
  borderLeft: "3px solid var(--accent-int)",
};

const inputStyle = {
  width: 70,
  padding: "4px 6px",
  borderRadius: 6,
  border: "1px solid var(--border)",
  fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
  fontSize: 12,
  background: "var(--card)",
  color: "var(--text-primary)",
};

const textInputStyle = {
  width: "100%",
  padding: "4px 6px",
  borderRadius: 6,
  border: "1px solid var(--border)",
  fontSize: 12,
  background: "var(--card)",
  color: "var(--text-primary)",
};

const selectStyle = {
  padding: "6px 10px",
  borderRadius: 8,
  border: "1px solid var(--border)",
  background: "var(--card)",
  fontSize: 13,
  color: "var(--text-primary)",
};

const fileInputStyle = {
  fontSize: 12,
  maxWidth: 220,
  color: "var(--text-primary)",
};

const buttonStyle = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  padding: "7px 13px",
  borderRadius: 8,
  border: "1px solid var(--border)",
  background: "var(--card)",
  fontSize: 12,
  fontWeight: 500,
  cursor: "pointer",
  color: "var(--text-primary)",
};

const dangerLinkStyle = {
  display: "inline-flex",
  alignItems: "center",
  gap: 4,
  background: "transparent",
  border: "none",
  color: "var(--danger)",
  fontSize: 11,
  cursor: "pointer",
  padding: 0,
};

const miniButtonStyle = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  padding: "3px 6px",
  borderRadius: 5,
  border: "1px solid var(--border)",
  background: "var(--track)",
  cursor: "pointer",
  color: "var(--text-primary)",
  lineHeight: 1,
};

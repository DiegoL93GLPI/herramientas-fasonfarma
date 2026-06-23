import { useState } from "react";
import * as XLSX from "xlsx";

const IPS_EMPLEADOR = 0.165;

const TIPOS_HHEE = [
  { id: "diurna", label: "Diurna (06:00–20:00)", mult: 1.5, color: "#fb923c" },
  { id: "nocturna", label: "Nocturna (20:00–06:00)", mult: 2.0, color: "#f472b6" },
  { id: "feriado", label: "Domingo / Feriado", mult: 2.0, color: "#ef4444" },
];

const EXTRAS_PREDEFINIDOS = [
  { id: "vacaciones", label: "Vacaciones", hint: "30 dias/año = salario / 12", auto: true },
  { id: "preaviso", label: "Preaviso", hint: "1 mes por año trabajado / 12", auto: true },
  { id: "indemnizacion", label: "Provision indemnizacion", hint: "Provision mensual por riesgo de desvinculacion", auto: true },
  { id: "epp", label: "EPP / Uniforme", hint: "Guantes, barbijo, cofia, delantal — costo anual / 12", auto: false },
  { id: "capacitacion", label: "Capacitacion BPM", hint: "Cursos, auditorias — costo anual / 12", auto: false },
  { id: "transporte", label: "Transporte", hint: "Pasaje o movilidad mensual", auto: false },
  { id: "bonofam", label: "Bonificacion familiar", hint: "Monto fijo mensual si aplica", auto: false },
];

function fmt(n) {
  if (!n && n !== 0) return "—";
  return `Gs. ${Math.round(n).toLocaleString("es-PY")}`;
}

function NumInput({ label, value, onChange, placeholder, unit, hint, min = "0", step = "any" }) {
  return (
    <div className="field">
      {label && <label className="flabel">{label}</label>}
      {hint && <p className="fhint">{hint}</p>}
      <div className="finput-wrap">
        <input type="number" className="finput" value={value}
          onChange={e => onChange(e.target.value)} placeholder={placeholder} min={min} step={step} />
        {unit && <span className="funit">{unit}</span>}
      </div>
    </div>
  );
}

let nextOpId = 1;
const newOp = () => ({ id: nextOpId++, nombre: "", salario: "" });

// ---- Exportar a Excel ----
function exportarExcel(historial) {
  const wb = XLSX.utils.book_new();

  historial.forEach((r, idx) => {
    const fecha = new Date(r.timestamp).toLocaleString("es-PY");
    const rows = [];

    // Encabezado
    rows.push(["CALCULO DE TARIFA HORA MOD — FasonFarma"]);
    rows.push([`Fecha: ${fecha}`]);
    rows.push([]);

    // Inputs — jornada
    rows.push(["JORNADA"]);
    rows.push(["Horas por dia", r.inputs.horasDia, "h"]);
    rows.push(["Dias del mes", r.inputs.diasMes, "dias"]);
    rows.push(["Horas ordinarias por operario", r.resultado.horasOrdinarias, "h"]);
    rows.push([]);

    // Inputs — operarios
    rows.push(["OPERARIOS"]);
    rows.push(["Nombre", "Salario neto", "IPS (16.5%)", "Aguinaldo (1/12)", "Total base"]);
    r.resultado.detalle.forEach(d => {
      rows.push([
        d.nombre,
        Math.round(d.salario),
        Math.round(d.ips),
        Math.round(d.aguinaldo),
        Math.round(d.salario + d.ips + d.aguinaldo),
      ]);
    });
    rows.push([]);

    // Horas extra
    if (r.resultado.hheeDetalle.length > 0) {
      rows.push(["HORAS EXTRAS"]);
      rows.push(["Tipo", "Horas grupo", "Multiplicador", "Costo"]);
      r.resultado.hheeDetalle.forEach(h => {
        rows.push([h.label, h.horas, `x${h.mult}`, Math.round(h.costo)]);
      });
      rows.push([]);
    }

    // Otros costos
    if (r.resultado.extrasActivos.length > 0) {
      rows.push(["OTROS COSTOS LABORALES"]);
      rows.push(["Concepto", "Monto mensual"]);
      r.resultado.extrasActivos.forEach(ex => {
        rows.push([ex.label, Math.round(ex.monto)]);
      });
      rows.push([]);
    }

    // Desglose total
    rows.push(["DESGLOSE COSTO MENSUAL"]);
    rows.push(["Concepto", "Monto (Gs.)", "% del total"]);
    const total = r.resultado.costoMensualTotal;
    [
      ["Salarios netos", r.resultado.totalSalarios],
      ["IPS empleador (16.5%)", r.resultado.totalIPS],
      ["Aguinaldo (1/12)", r.resultado.totalAguinaldo],
      ...r.resultado.hheeDetalle.map(h => [`HHEE ${h.label}`, h.costo]),
      ...r.resultado.extrasActivos.map(ex => [ex.label, ex.monto]),
    ].forEach(([label, val]) => {
      rows.push([label, Math.round(val), `${((val / total) * 100).toFixed(1)}%`]);
    });
    rows.push([]);

    // Resultado final
    rows.push(["RESULTADO"]);
    rows.push(["Costo mensual total del personal", Math.round(r.resultado.costoMensualTotal), "Gs."]);
    rows.push(["Total horas trabajadas en el mes (grupo)", r.resultado.horasTotales, "h"]);
    rows.push(["TARIFA HORA MOD", Math.round(r.resultado.tarifaHoraMOD), "Gs./h"]);

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Anchos de columna
    ws["!cols"] = [{ wch: 38 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];

    const sheetName = `Calculo ${idx + 1}`;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  // Hoja resumen si hay más de uno
  if (historial.length > 1) {
    const resumen = [["RESUMEN DE CALCULOS"]];
    resumen.push([]);
    resumen.push(["#", "Fecha", "Operarios", "Hs. ord./op.", "Total hs. grupo", "Costo total (Gs.)", "Tarifa MOD (Gs./h)"]);
    historial.forEach((r, idx) => {
      resumen.push([
        idx + 1,
        new Date(r.timestamp).toLocaleString("es-PY"),
        r.resultado.detalle.length,
        r.resultado.horasOrdinarias,
        r.resultado.horasTotales,
        Math.round(r.resultado.costoMensualTotal),
        Math.round(r.resultado.tarifaHoraMOD),
      ]);
    });
    const wsRes = XLSX.utils.aoa_to_sheet(resumen);
    wsRes["!cols"] = [{ wch: 4 }, { wch: 22 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  }

  XLSX.writeFile(wb, `Tarifa_MOD_FasonFarma_${new Date().toISOString().slice(0,10)}.xlsx`);
}

export default function TarifaMOD() {
  const [horasDia, setHorasDia] = useState("8");
  const [diasMes, setDiasMes] = useState("30");
  const [operarios, setOperarios] = useState([newOp()]);
  const [bloqHHEE, setBloqHHEE] = useState([]);
  const [extras, setExtras] = useState(
    EXTRAS_PREDEFINIDOS.map(e => ({ ...e, activo: false, montoManual: "" }))
  );
  const [resultado, setResultado] = useState(null);
  const [copiado, setCopiado] = useState(false);
  const [historial, setHistorial] = useState([]); // acumula calculos

  const addOp = () => setOperarios(o => [...o, newOp()]);
  const removeOp = (id) => setOperarios(o => o.filter(x => x.id !== id));
  const updateOp = (id, field, val) =>
    setOperarios(o => o.map(x => x.id === id ? { ...x, [field]: val } : x));

  const addBloq = () => setBloqHHEE(b => [...b, { id: Date.now(), tipo: "diurna", horas: "" }]);
  const removeBloq = (id) => setBloqHHEE(b => b.filter(x => x.id !== id));
  const updateBloq = (id, field, val) =>
    setBloqHHEE(b => b.map(x => x.id === id ? { ...x, [field]: val } : x));

  const toggleExtra = (id) =>
    setExtras(e => e.map(x => x.id === id ? { ...x, activo: !x.activo } : x));
  const updateExtraMonto = (id, val) =>
    setExtras(e => e.map(x => x.id === id ? { ...x, montoManual: val } : x));

  const opsValidas = operarios.filter(o => o.nombre.trim() && parseFloat(o.salario) > 0);
  const salarioPromedio = opsValidas.length > 0
    ? opsValidas.reduce((s, o) => s + parseFloat(o.salario), 0) / opsValidas.length : 0;
  const horaOrdPromedio = salarioPromedio > 0 && parseFloat(horasDia) > 0
    ? salarioPromedio / 30 / parseFloat(horasDia) : 0;

  const calcular = () => {
    const hDia = parseFloat(horasDia);
    const dMes = parseFloat(diasMes);
    if (!hDia || hDia <= 0 || !dMes || dMes <= 0) { alert("Ingresá horas por día y días del mes."); return; }
    const ops = operarios.filter(o => o.nombre.trim() && parseFloat(o.salario) > 0);
    if (ops.length === 0) { alert("Ingresá al menos un operario con salario."); return; }

    const horasOrdinarias = hDia * dMes;
    let totalSalarios = 0, totalIPS = 0, totalAguinaldo = 0;

    const detalle = ops.map(o => {
      const salario = parseFloat(o.salario);
      const valorHoraOrdinaria = salario / 30 / hDia;
      const ips = salario * IPS_EMPLEADOR;
      const aguinaldo = salario / 12;
      totalSalarios += salario;
      totalIPS += ips;
      totalAguinaldo += aguinaldo;
      return { nombre: o.nombre, salario, valorHoraOrdinaria, ips, aguinaldo };
    });

    let totalCostoHHEE = 0, totalHorasHHEE = 0;
    const hheeDetalle = [];
    const valorHoraPromedio = detalle.reduce((s, d) => s + d.valorHoraOrdinaria, 0) / detalle.length;

    bloqHHEE.forEach(b => {
      const horas = parseFloat(b.horas) || 0;
      if (horas <= 0) return;
      const tipo = TIPOS_HHEE.find(t => t.id === b.tipo);
      const costoBloque = horas * valorHoraPromedio * tipo.mult * ops.length;
      totalCostoHHEE += costoBloque;
      totalHorasHHEE += horas;
      hheeDetalle.push({ label: tipo.label, horas, mult: tipo.mult, costo: costoBloque, color: tipo.color });
    });

    const extrasActivos = [];
    let totalExtras = 0;
    extras.filter(e => e.activo).forEach(e => {
      let monto = e.auto ? totalSalarios / 12 : (parseFloat(e.montoManual) || 0) * ops.length;
      if (monto > 0) {
        extrasActivos.push({ label: e.label, monto, color: "#38bdf8" });
        totalExtras += monto;
      }
    });

    const costoMensualTotal = totalSalarios + totalIPS + totalAguinaldo + totalCostoHHEE + totalExtras;
    const horasTotales = horasOrdinarias * ops.length + totalHorasHHEE;
    const tarifaHoraMOD = costoMensualTotal / horasTotales;

    const res = {
      detalle, horasOrdinarias, horasTotales,
      totalSalarios, totalIPS, totalAguinaldo, totalCostoHHEE,
      hheeDetalle, extrasActivos, totalExtras,
      costoMensualTotal, tarifaHoraMOD, valorHoraPromedio, hDia,
    };

    setResultado(res);

    // Guardar en historial
    setHistorial(h => [...h, {
      timestamp: Date.now(),
      inputs: { horasDia: hDia, diasMes: dMes },
      resultado: res,
    }]);
  };

  const pct = (val) => resultado ? ((val / resultado.costoMensualTotal) * 100).toFixed(1) : "0";

  const copiar = () => {
    if (!resultado) return;
    navigator.clipboard?.writeText(Math.round(resultado.tarifaHoraMOD).toString());
    setCopiado(true);
    setTimeout(() => setCopiado(false), 2000);
  };

  return (
    <>
      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0c0f18; }
        .app { min-height: 100vh; background: #0c0f18; color: #dde1ea; font-family: 'Inter','Segoe UI',system-ui,sans-serif; padding: 2rem 1rem 3rem; }
        .header { max-width: 680px; margin: 0 auto 2rem; text-align: center; }
        .eyebrow { font-size: .68rem; letter-spacing: .18em; text-transform: uppercase; color: #4a7fa5; margin-bottom: .5rem; }
        .title { font-size: 1.75rem; font-weight: 800; letter-spacing: -.03em; color: #dde1ea; }
        .title em { font-style: normal; color: #5ba3f5; }
        .subtitle { font-size: .82rem; color: #4b5568; margin-top: .4rem; }
        .main { max-width: 680px; margin: 0 auto; display: flex; flex-direction: column; gap: 1.25rem; }
        .card { background: #141926; border: 1px solid #1e2535; border-radius: 14px; padding: 1.5rem; }
        .card-result { border-color: #2a3f5f; }
        .sec-header { display: flex; align-items: center; gap: .6rem; margin-bottom: 1.25rem; }
        .sec-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
        .sec-title { font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .12em; color: #6b7280; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: .75rem; }
        .field { margin-bottom: .9rem; }
        .flabel { display: block; font-size: .72rem; color: #5b6475; margin-bottom: .3rem; font-weight: 500; }
        .fhint { font-size: .67rem; color: #374151; margin-bottom: .3rem; line-height: 1.4; }
        .finput-wrap { position: relative; display: flex; align-items: center; }
        .finput { width: 100%; padding: .55rem 2.5rem .55rem .75rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .88rem; outline: none; transition: border-color .15s; }
        .finput:focus { border-color: #5ba3f5; }
        .finput::placeholder { color: #2d3748; }
        .funit { position: absolute; right: .65rem; font-size: .68rem; color: #374151; pointer-events: none; }
        .info-box { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .65rem .9rem; font-size: .72rem; color: #4b5568; line-height: 1.7; margin-bottom: .9rem; }
        .info-box strong { color: #6b7280; }
        .info-box .hi { color: #5ba3f5; font-weight: 700; }
        .op-header { display: grid; grid-template-columns: 1fr 160px 28px; gap: .5rem; margin-bottom: .4rem; }
        .col-lbl { font-size: .64rem; color: #374151; text-transform: uppercase; letter-spacing: .08em; }
        .col-lbl.right { text-align: right; }
        .op-row { display: grid; grid-template-columns: 1fr 160px 28px; gap: .5rem; align-items: center; margin-bottom: .55rem; }
        .op-text { padding: .52rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; width: 100%; }
        .op-text:focus { border-color: #5ba3f5; }
        .op-text::placeholder { color: #2d3748; }
        .op-num { padding: .52rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; width: 100%; text-align: right; }
        .op-num:focus { border-color: #5ba3f5; }
        .op-num::placeholder { color: #2d3748; }
        .btn-remove { width: 28px; height: 28px; border-radius: 6px; border: 1px solid #1e2535; background: #0c0f18; color: #4b5568; cursor: pointer; font-size: 1rem; display: flex; align-items: center; justify-content: center; transition: all .15s; flex-shrink: 0; }
        .btn-remove:hover { border-color: #ef4444; color: #ef4444; }
        .hhee-row { display: grid; grid-template-columns: 1fr 100px 28px; gap: .5rem; align-items: center; margin-bottom: .55rem; }
        .fselect { width: 100%; padding: .52rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .82rem; outline: none; cursor: pointer; }
        .fselect:focus { border-color: #5ba3f5; }
        .btn-add { width: 100%; padding: .42rem; border: 1px dashed #1e2535; border-radius: 7px; background: transparent; color: #4b5568; font-size: .78rem; cursor: pointer; transition: all .15s; margin-top: .2rem; }
        .btn-add:hover { border-color: #5ba3f5; color: #5ba3f5; }
        .divider { height: 1px; background: #1e2535; margin: 1rem 0; }
        .extras-grid { display: flex; flex-direction: column; gap: .6rem; }
        .extra-item { background: #0c0f18; border: 1px solid #1e2535; border-radius: 9px; padding: .65rem .85rem; transition: border-color .15s; }
        .extra-item.active { border-color: #2a3f5f; background: #0d1525; }
        .extra-top { display: flex; align-items: center; gap: .65rem; cursor: pointer; }
        .extra-check { width: 16px; height: 16px; border-radius: 4px; border: 1px solid #374151; background: #141926; display: flex; align-items: center; justify-content: center; flex-shrink: 0; transition: all .15s; }
        .extra-check.on { background: #5ba3f5; border-color: #5ba3f5; }
        .extra-check-mark { font-size: .65rem; color: #0c0f18; font-weight: 900; }
        .extra-name { font-size: .82rem; color: #9ca3af; font-weight: 500; flex: 1; }
        .extra-item.active .extra-name { color: #dde1ea; }
        .extra-hint { font-size: .67rem; color: #374151; margin-top: .1rem; }
        .extra-tag { font-size: .65rem; padding: .15rem .45rem; border-radius: 4px; background: #0b2218; color: #34d399; border: 1px solid #064e2e; }
        .extra-monto { margin-top: .65rem; padding-top: .65rem; border-top: 1px solid #1e2535; }
        .btn-primary { width: 100%; padding: .75rem; background: #5ba3f5; color: #0c0f18; border: none; border-radius: 8px; font-size: .88rem; font-weight: 800; cursor: pointer; transition: background .15s; letter-spacing: .02em; }
        .btn-primary:hover { background: #4a8de0; }
        .btn-ghost { width: 100%; padding: .55rem; background: transparent; color: #4b5568; border: 1px solid #1e2535; border-radius: 8px; font-size: .78rem; cursor: pointer; margin-top: .6rem; transition: all .15s; }
        .btn-ghost:hover { color: #dde1ea; border-color: #374151; }
        .btn-excel { width: 100%; padding: .65rem; background: #0b2e14; color: #34d399; border: 1px solid #0d5c2a; border-radius: 8px; font-size: .82rem; font-weight: 700; cursor: pointer; margin-top: .6rem; transition: all .15s; display: flex; align-items: center; justify-content: center; gap: .5rem; }
        .btn-excel:hover { background: #0d3d1b; border-color: #34d399; }
        .btn-excel:disabled { opacity: .35; cursor: not-allowed; }
        .historial-badge { display: inline-flex; align-items: center; gap: .4rem; padding: .25rem .65rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 20px; font-size: .7rem; color: #6b7280; margin-top: .5rem; }
        .historial-badge span { color: #5ba3f5; font-weight: 700; }
        .result-hero { text-align: center; padding: 1.5rem 0 1.25rem; border-bottom: 1px solid #1e2535; margin-bottom: 1.25rem; }
        .result-eyebrow { font-size: .65rem; letter-spacing: .15em; text-transform: uppercase; color: #4a7fa5; margin-bottom: .4rem; }
        .result-value { font-size: 2.8rem; font-weight: 900; color: #5ba3f5; letter-spacing: -.04em; line-height: 1; }
        .result-label { font-size: .75rem; color: #4b5568; margin-top: .35rem; }
        .copy-btn { display: inline-block; margin-top: .75rem; padding: .3rem .9rem; border-radius: 20px; font-size: .72rem; cursor: pointer; transition: all .2s; border: none; }
        .copy-btn.idle { background: #172038; border: 1px solid #2a3f5f; color: #5ba3f5; }
        .copy-btn.done { background: #0b2218; border: 1px solid #064e2e; color: #34d399; }
        .stats-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: .75rem; margin-bottom: 1.25rem; }
        .stat-box { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .65rem .75rem; text-align: center; }
        .stat-val { font-size: .92rem; font-weight: 700; color: #dde1ea; }
        .stat-lbl { font-size: .62rem; color: #374151; text-transform: uppercase; letter-spacing: .06em; margin-top: .2rem; }
        .bar-section { display: flex; flex-direction: column; gap: .85rem; margin-bottom: 1rem; }
        .bar-top { display: flex; justify-content: space-between; margin-bottom: .3rem; }
        .bar-lbl { font-size: .75rem; color: #6b7280; }
        .bar-val { font-size: .75rem; color: #dde1ea; font-weight: 600; }
        .bar-track { height: 5px; background: #1e2535; border-radius: 3px; overflow: hidden; margin-bottom: .15rem; }
        .bar-fill { height: 100%; border-radius: 3px; transition: width .5s ease; }
        .bar-pct { font-size: .65rem; color: #2d3748; }
        .formula-box { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .75rem 1rem; font-size: .72rem; color: #4b5568; margin-bottom: 1.1rem; line-height: 1.8; }
        .formula-box strong { color: #6b7280; }
        .total-row { display: flex; justify-content: space-between; align-items: center; padding-top: .9rem; border-top: 1px solid #1e2535; margin-bottom: 1.25rem; }
        .total-lbl { font-size: .8rem; color: #5b6475; }
        .total-val { font-size: 1rem; font-weight: 800; color: #dde1ea; }
        .op-table { width: 100%; border-collapse: collapse; font-size: .75rem; margin-bottom: .5rem; }
        .op-table th { text-align: left; color: #374151; font-size: .62rem; text-transform: uppercase; letter-spacing: .07em; padding: .4rem .5rem; border-bottom: 1px solid #1e2535; }
        .op-table th.right { text-align: right; }
        .op-table td { padding: .45rem .5rem; color: #9ca3af; border-bottom: 1px solid #0f1420; }
        .op-table td.right { text-align: right; }
        .op-table td.name { color: #dde1ea; font-weight: 600; }
        .op-table tr:last-child td { border-bottom: none; }
        .section-mini-label { font-size: .68rem; color: #374151; text-transform: uppercase; letter-spacing: .07em; margin-bottom: .6rem; }
        @media (max-width:520px) {
          .two-col { grid-template-columns: 1fr; }
          .stats-grid { grid-template-columns: 1fr 1fr; }
          .op-header, .op-row { grid-template-columns: 1fr 120px 28px; }
          .hhee-row { grid-template-columns: 1fr 90px 28px; }
        }
      `}</style>

      <div className="app">
        <div className="header">
          <p className="eyebrow">FasonFarma — MOD</p>
          <h1 className="title">Tarifa <em>hora</em> de mano de obra</h1>
          <p className="subtitle">Salario + IPS + Aguinaldo + Horas extra + Otros costos</p>
        </div>

        <div className="main">

          {/* Jornada */}
          <div className="card">
            <div className="sec-header">
              <div className="sec-dot" style={{ background: "#5ba3f5" }} />
              <span className="sec-title">Jornada laboral</span>
            </div>
            <div className="two-col">
              <NumInput label="Horas por dia" value={horasDia} onChange={setHorasDia} placeholder="8" unit="h" step="0.5" min="1" />
              <NumInput label="Dias del mes" value={diasMes} onChange={setDiasMes} placeholder="30" unit="dias" step="1" min="1" hint="Base legal: 30 dias" />
            </div>
            <div className="info-box">
              <strong>Formula legal (Codigo Laboral PY):</strong><br />
              Valor hora ordinaria = Salario / 30 / horas diarias<br />
              <span className="hi">Hora ord. promedio: {horaOrdPromedio > 0 ? fmt(horaOrdPromedio) : "—"}</span>
            </div>
          </div>

          {/* Operarios */}
          <div className="card">
            <div className="sec-header">
              <div className="sec-dot" style={{ background: "#a78bfa" }} />
              <span className="sec-title">Operarios</span>
            </div>
            <div className="op-header">
              <span className="col-lbl">Nombre / Rol</span>
              <span className="col-lbl right">Salario neto mensual</span>
              <span />
            </div>
            {operarios.map(o => (
              <div key={o.id} className="op-row">
                <input type="text" className="op-text" placeholder="Ej: Operario linea 1"
                  value={o.nombre} onChange={e => updateOp(o.id, "nombre", e.target.value)} />
                <input type="number" className="op-num" placeholder="2.500.000"
                  min="0" step="any" value={o.salario} onChange={e => updateOp(o.id, "salario", e.target.value)} />
                <button className="btn-remove" onClick={() => removeOp(o.id)}>×</button>
              </div>
            ))}
            <button className="btn-add" onClick={addOp}>+ Agregar operario</button>
            <div className="divider" />
            <div className="info-box">
              <strong>IPS empleador:</strong> 16.5% del salario (Art. 9, Ley 98/92) — costo tuyo<br />
              <strong>IPS empleado:</strong> 9% — se descuenta del salario del trabajador, no es tu costo<br />
              <strong>Aguinaldo:</strong> salario / 12 (provision mensual del 13er salario anual)
            </div>
          </div>

          {/* Horas extra */}
          <div className="card">
            <div className="sec-header">
              <div className="sec-dot" style={{ background: "#fb923c" }} />
              <span className="sec-title">Horas extras — opcional</span>
            </div>
            <div className="info-box">
              <strong>Diurna</strong> (06:00–20:00): x1.5 &nbsp;|&nbsp;
              <strong>Nocturna</strong> (20:00–06:00): x2.0 &nbsp;|&nbsp;
              <strong>Dom/Feriado</strong>: x2.0<br />
              Limite legal: max 3 hs extra/dia, max 3 veces/semana.
            </div>
            {bloqHHEE.length > 0 && (
              <div className="op-header" style={{ gridTemplateColumns: "1fr 100px 28px" }}>
                <span className="col-lbl">Tipo</span>
                <span className="col-lbl right">Horas grupo</span>
                <span />
              </div>
            )}
            {bloqHHEE.map(b => (
              <div key={b.id} className="hhee-row">
                <select className="fselect" value={b.tipo} onChange={e => updateBloq(b.id, "tipo", e.target.value)}>
                  {TIPOS_HHEE.map(t => <option key={t.id} value={t.id}>{t.label} (x{t.mult})</option>)}
                </select>
                <input type="number" className="op-num" placeholder="0" min="0" step="0.5"
                  value={b.horas} onChange={e => updateBloq(b.id, "horas", e.target.value)} />
                <button className="btn-remove" onClick={() => removeBloq(b.id)}>×</button>
              </div>
            ))}
            <button className="btn-add" onClick={addBloq}>+ Agregar bloque de horas extra</button>
          </div>

          {/* Otros costos */}
          <div className="card">
            <div className="sec-header">
              <div className="sec-dot" style={{ background: "#38bdf8" }} />
              <span className="sec-title">Otros costos laborales — opcional</span>
            </div>
            <div className="extras-grid">
              {extras.map(e => (
                <div key={e.id} className={`extra-item ${e.activo ? "active" : ""}`}>
                  <div className="extra-top" onClick={() => toggleExtra(e.id)}>
                    <div className={`extra-check ${e.activo ? "on" : ""}`}>
                      {e.activo && <span className="extra-check-mark">✓</span>}
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: ".5rem" }}>
                        <span className="extra-name">{e.label}</span>
                        {e.auto && <span className="extra-tag">auto</span>}
                      </div>
                      <p className="extra-hint">{e.hint}</p>
                    </div>
                  </div>
                  {e.activo && !e.auto && (
                    <div className="extra-monto">
                      <NumInput label="Monto por operario / mes" value={e.montoManual}
                        onChange={v => updateExtraMonto(e.id, v)} placeholder="Ej: 150000" unit="Gs." />
                    </div>
                  )}
                  {e.activo && e.auto && (
                    <div className="extra-monto">
                      <p className="fhint" style={{ color: "#4b5568" }}>
                        Se calculara como <strong style={{ color: "#5ba3f5" }}>salario total / 12</strong> al momento de calcular.
                      </p>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          <button className="btn-primary" onClick={calcular}>Calcular tarifa hora MOD</button>

          {/* Resultado */}
          {resultado && (
            <div className="card card-result">
              <div className="sec-header">
                <div className="sec-dot" style={{ background: "#5ba3f5" }} />
                <span className="sec-title">Resultado</span>
              </div>

              <div className="result-hero">
                <p className="result-eyebrow">Tarifa hora MOD</p>
                <p className="result-value">{fmt(resultado.tarifaHoraMOD)}</p>
                <p className="result-label">por hora trabajada</p>
                <button className={`copy-btn ${copiado ? "done" : "idle"}`} onClick={copiar}>
                  {copiado ? "Copiado" : "Copiar → usar en estimador de ordenes"}
                </button>
              </div>

              <div className="stats-grid">
                <div className="stat-box">
                  <p className="stat-val">{resultado.detalle.length}</p>
                  <p className="stat-lbl">Operarios</p>
                </div>
                <div className="stat-box">
                  <p className="stat-val">{resultado.horasOrdinarias.toLocaleString("es-PY")}</p>
                  <p className="stat-lbl">Hs. ord. / op.</p>
                </div>
                <div className="stat-box">
                  <p className="stat-val">{resultado.horasTotales.toLocaleString("es-PY")}</p>
                  <p className="stat-lbl">Total hs. grupo</p>
                </div>
              </div>

              <div className="formula-box">
                <strong>Formula aplicada:</strong><br />
                Valor hora ord. = Salario / 30 / {resultado.hDia}h = {fmt(resultado.valorHoraPromedio)} (promedio)<br />
                Tarifa MOD = {fmt(resultado.costoMensualTotal)} / {resultado.horasTotales.toLocaleString("es-PY")}h = <strong style={{ color: "#5ba3f5" }}>{fmt(resultado.tarifaHoraMOD)}</strong>
              </div>

              <div className="bar-section">
                {[
                  { label: "Salarios netos", val: resultado.totalSalarios, color: "#5ba3f5" },
                  { label: "IPS empleador (16.5%)", val: resultado.totalIPS, color: "#a78bfa" },
                  { label: "Aguinaldo (1/12)", val: resultado.totalAguinaldo, color: "#34d399" },
                  ...resultado.hheeDetalle.map(h => ({ label: `HHEE ${h.label} — ${h.horas}h x${h.mult}`, val: h.costo, color: h.color })),
                  ...resultado.extrasActivos.map(ex => ({ label: ex.label, val: ex.monto, color: "#38bdf8" })),
                ].map(b => (
                  <div key={b.label}>
                    <div className="bar-top">
                      <span className="bar-lbl">{b.label}</span>
                      <span className="bar-val">{fmt(b.val)}</span>
                    </div>
                    <div className="bar-track">
                      <div className="bar-fill" style={{ width: `${pct(b.val)}%`, background: b.color }} />
                    </div>
                    <span className="bar-pct">{pct(b.val)}% del costo total</span>
                  </div>
                ))}
              </div>

              <div className="total-row">
                <span className="total-lbl">Costo mensual total del personal</span>
                <span className="total-val">{fmt(resultado.costoMensualTotal)}</span>
              </div>

              {resultado.detalle.length > 1 && (
                <>
                  <p className="section-mini-label">Desglose por operario</p>
                  <table className="op-table">
                    <thead>
                      <tr>
                        <th>Nombre</th>
                        <th className="right">Salario</th>
                        <th className="right">IPS</th>
                        <th className="right">Aguinaldo</th>
                        <th className="right">Total base</th>
                      </tr>
                    </thead>
                    <tbody>
                      {resultado.detalle.map((d, i) => (
                        <tr key={i}>
                          <td className="name">{d.nombre}</td>
                          <td className="right">{fmt(d.salario)}</td>
                          <td className="right">{fmt(d.ips)}</td>
                          <td className="right">{fmt(d.aguinaldo)}</td>
                          <td className="right" style={{ color: "#dde1ea", fontWeight: 700 }}>
                            {fmt(d.salario + d.ips + d.aguinaldo)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              )}

              <div className="divider" />

              {/* Exportar */}
              <div style={{ textAlign: "center" }}>
                <div className="historial-badge">
                  Calculos acumulados en esta sesion: <span>{historial.length}</span>
                </div>
              </div>

              <button className="btn-excel" onClick={() => exportarExcel(historial)}>
                Exportar a Excel — {historial.length} {historial.length === 1 ? "calculo" : "calculos"}
                {historial.length > 1 ? " + hoja Resumen" : ""}
              </button>

              <button className="btn-ghost" onClick={() => setResultado(null)}>Nuevo calculo</button>
            </div>
          )}

        </div>
      </div>
    </>
  );
}

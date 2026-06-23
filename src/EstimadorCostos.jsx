import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const TIPOS_PROCESO = ["Elaboración", "Etiquetado", "Sellado", "Acondicionamiento"];

const CIF_RUBROS_DEFAULT = [
  { id: 1, nombre: "Alquiler del local", monto: "" },
  { id: 2, nombre: "Energía eléctrica", monto: "" },
  { id: 3, nombre: "Agua", monto: "" },
  { id: 4, nombre: "Depreciación maquinaria", monto: "" },
  { id: 5, nombre: "Mantenimiento equipos", monto: "" },
  { id: 6, nombre: "Seguro", monto: "" },
];

let nextCifId = 7;
let nextProcId = 1;

const newProceso = () => ({
  id: nextProcId++,
  tipo: TIPOS_PROCESO[0],
  personas: "",
  horasMOD: "",
  tarifaMOD: "",
  pctMerma: "",
});

function fmt(n) {
  if (!n && n !== 0) return "—";
  return `Gs. ${Math.round(n).toLocaleString("es-PY")}`;
}

function NumInput({ label, value, onChange, placeholder, unit, hint, min = "0", step = "any" }) {
  return (
    <div className="pinput-field">
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

// ---- CIF por horas de produccion ----
function calcularProcesos(procesos, cantidadInicial, totalCIF, horasMensualesPlanta) {
  let unidadesEntrada = cantidadInicial;
  const resultados = [];
  let costoMODTotal = 0;
  let horasTotalesOrden = 0;

  const datosProc = procesos.map(p => {
    const personas = parseFloat(p.personas) || 1;
    const horas = parseFloat(p.horasMOD) || 0;
    const tarifa = parseFloat(p.tarifaMOD) || 0;
    const merma = parseFloat(p.pctMerma) || 0;
    const costoMOD = personas * horas * tarifa;
    horasTotalesOrden += horas;
    costoMODTotal += costoMOD;
    return { personas, horas, tarifa, merma, costoMOD };
  });

  const cifPorHora = totalCIF / horasMensualesPlanta;
  const cifTotalOrden = horasTotalesOrden * cifPorHora;

  let costoTotalAcumulado = 0;

  datosProc.forEach((dp, idx) => {
    const proporcionHoras = horasTotalesOrden > 0 ? dp.horas / horasTotalesOrden : 0;
    const cifAsignado = cifTotalOrden * proporcionHoras;
    const costoProc = dp.costoMOD + cifAsignado;
    const unidadesSalida = Math.floor(unidadesEntrada * (1 - dp.merma / 100));
    const costoUnitario = unidadesSalida > 0 ? costoProc / unidadesSalida : 0;
    costoTotalAcumulado += costoProc;

    resultados.push({
      idx: idx + 1,
      tipo: procesos[idx].tipo,
      personas: dp.personas,
      horas: dp.horas,
      tarifa: dp.tarifa,
      merma: dp.merma,
      unidadesEntrada: Math.floor(unidadesEntrada),
      unidadesSalida,
      costoMOD: dp.costoMOD,
      cifAsignado,
      costoProc,
      costoUnitario,
    });

    unidadesEntrada = unidadesSalida;
  });

  const unidadesFinales = resultados.length > 0 ? resultados[resultados.length - 1].unidadesSalida : cantidadInicial;
  const costoUnitarioFinal = unidadesFinales > 0 ? costoTotalAcumulado / unidadesFinales : 0;

  return {
    resultados, costoTotalAcumulado, unidadesFinales, costoUnitarioFinal,
    costoMODTotal, cifTotalOrden, horasTotalesOrden, cifPorHora,
  };
}

// ---- Export Excel ----
function exportarExcel(historial, cifRubros, totalCIF, horasDia, diasMes) {
  const horasMes = horasDia * diasMes;
  const cifPorHora = totalCIF / horasMes;
  const wb = XLSX.utils.book_new();

  const cifRows = [
    ["COSTOS INDIRECTOS DE FABRICACION (CIF) — MENSUAL"], [],
    ["Rubro", "Monto mensual (Gs.)"],
    ...cifRubros.filter(r => parseFloat(r.monto) > 0).map(r => [r.nombre, Math.round(parseFloat(r.monto))]),
    [],
    ["TOTAL CIF MENSUAL", Math.round(totalCIF)],
    [],
    ["DISTRIBUCION POR HORAS"],
    ["Horas por dia", horasDia],
    ["Dias habiles por mes", diasMes],
    ["Horas mensuales de planta", horasMes],
    ["CIF por hora de planta", Math.round(cifPorHora), "Gs./h"],
  ];
  const wsCIF = XLSX.utils.aoa_to_sheet(cifRows);
  wsCIF["!cols"] = [{ wch: 35 }, { wch: 22 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsCIF, "CIF Mensual");

  historial.forEach((h, idx) => {
    const c = h.calculo;
    const rows = [
      [`ESTIMACION — ${h.inputs.producto}`],
      [`Fecha: ${new Date(h.timestamp).toLocaleString("es-PY")}`], [],
      ["Cantidad inicial", h.inputs.cantidadInicial, "u."], [],
      ["PASO 1 — MOD TOTAL DE LA ORDEN"],
      ["#", "Proceso", "Personas", "Horas", "Tarifa (Gs./h)", "Calculo", "Costo MOD"],
    ];
    c.resultados.forEach(r => {
      rows.push([r.idx, r.tipo, r.personas, r.horas, Math.round(r.tarifa),
        `${r.personas} x ${r.horas}h x Gs. ${Math.round(r.tarifa).toLocaleString("es-PY")}`, Math.round(r.costoMOD)]);
    });
    rows.push(["", "", "", "", "", "TOTAL MOD", Math.round(c.costoMODTotal)]);
    rows.push([]);
    rows.push(["PASO 2 — CIF ASIGNADO A LA ORDEN"]);
    rows.push(["Horas totales de la orden", c.horasTotalesOrden, "h"]);
    rows.push(["CIF por hora", Math.round(c.cifPorHora), "Gs./h"]);
    rows.push(["CIF orden", `${c.horasTotalesOrden}h x Gs. ${Math.round(c.cifPorHora).toLocaleString("es-PY")}`, Math.round(c.cifTotalOrden)]);
    rows.push([]);
    rows.push(["PASO 3 — COSTO TOTAL DE LA ORDEN"]);
    rows.push(["MOD", Math.round(c.costoMODTotal)]);
    rows.push(["CIF", Math.round(c.cifTotalOrden)]);
    rows.push(["TOTAL ORDEN", Math.round(c.costoTotalAcumulado)]);
    rows.push([]);
    rows.push(["PASO 4 — COSTO UNITARIO"]);
    rows.push(["Unidades iniciales", h.inputs.cantidadInicial]);
    rows.push(["Unidades finales (con merma)", c.unidadesFinales]);
    rows.push(["COSTO UNITARIO FINAL", Math.round(c.costoUnitarioFinal), "Gs./u."]);
    rows.push([]);
    rows.push(["DESGLOSE POR PROCESO"]);
    rows.push(["#", "Proceso", "% Merma", "Entrada", "Salida", "MOD", "CIF", "Costo proc.", "Costo unit."]);
    c.resultados.forEach(r => {
      rows.push([r.idx, r.tipo, `${r.merma}%`, r.unidadesEntrada, r.unidadesSalida,
        Math.round(r.costoMOD), Math.round(r.cifAsignado), Math.round(r.costoProc), Math.round(r.costoUnitario)]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 4 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 28 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws, `Orden ${idx + 1}`);
  });

  if (historial.length > 1) {
    const resumen = [
      ["RESUMEN DE ORDENES"], [],
      ["#", "Producto", "Procesos", "Horas ord.", "Unid. iniciales", "Unid. finales", "MOD", "CIF", "Costo total", "Costo unitario"],
      ...historial.map((h, i) => {
        const c = h.calculo;
        return [i + 1, h.inputs.producto, c.resultados.length, c.horasTotalesOrden,
          h.inputs.cantidadInicial, c.unidadesFinales,
          Math.round(c.costoMODTotal), Math.round(c.cifTotalOrden),
          Math.round(c.costoTotalAcumulado), Math.round(c.costoUnitarioFinal)];
      }),
    ];
    const wsRes = XLSX.utils.aoa_to_sheet(resumen);
    wsRes["!cols"] = [{ wch: 4 }, { wch: 24 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsRes, "Resumen");
  }

  XLSX.writeFile(wb, `Costos_Produccion_FasonFarma_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ---- Export PDF ----
function exportarPDF(historial, cifRubros, totalCIF, horasDia, diasMes) {
  const fecha = new Date().toLocaleDateString("es-PY", { day: "2-digit", month: "2-digit", year: "numeric" });
  const horasMes = horasDia * diasMes;
  const cifFiltered = cifRubros.filter(r => parseFloat(r.monto) > 0);

  const cifRows = cifFiltered.map(r =>
    `<tr><td>${r.nombre}</td><td class="right">${fmt(parseFloat(r.monto))}</td></tr>`
  ).join("");

  const ordenesHTML = historial.map((h, idx) => {
    const c = h.calculo;

    const modRows = c.resultados.map(r =>
      `<tr><td class="proc-num">${r.idx}</td><td>${r.tipo}</td>
       <td class="right">${r.personas}</td><td class="right">${r.horas}h</td>
       <td class="right">${fmt(r.tarifa)}/h</td>
       <td class="right">${fmt(r.costoMOD)}</td></tr>`
    ).join("");

    const procesosRows = c.resultados.map(r =>
      `<tr><td class="proc-num">${r.idx}</td><td>${r.tipo}</td>
       <td class="right">${r.merma}%</td>
       <td class="right">${r.unidadesEntrada.toLocaleString("es-PY")}</td>
       <td class="right">${r.unidadesSalida.toLocaleString("es-PY")}</td>
       <td class="right">${fmt(r.costoMOD)}</td>
       <td class="right">${fmt(r.cifAsignado)}</td>
       <td class="right bold">${fmt(r.costoUnitario)}</td></tr>`
    ).join("");

    return `
    <div class="orden-block${idx > 0 ? " page-break" : ""}">
      <div class="orden-header">
        <span class="orden-num">${h.inputs.producto}</span>
        <span class="orden-fecha">${new Date(h.timestamp).toLocaleDateString("es-PY")}</span>
      </div>
      <div class="orden-meta">Cantidad inicial: <strong>${h.inputs.cantidadInicial.toLocaleString("es-PY")} u.</strong></div>

      <div class="paso-label">Paso 1 — MOD total de la orden</div>
      <table class="proc-table">
        <thead><tr><th>#</th><th>Proceso</th><th class="right">Pers.</th><th class="right">Horas</th><th class="right">Tarifa</th><th class="right">Costo MOD</th></tr></thead>
        <tbody>${modRows}</tbody>
        <tfoot><tr><td colspan="5" class="right total-td">Total MOD</td><td class="right total-td">${fmt(c.costoMODTotal)}</td></tr></tfoot>
      </table>

      <div class="paso-label">Paso 2 — CIF asignado a la orden</div>
      <table class="data-table">
        <tr><td class="lbl">Horas totales de la orden</td><td class="val">${c.horasTotalesOrden} h</td></tr>
        <tr><td class="lbl">CIF por hora de planta</td><td class="val">${fmt(c.cifPorHora)}/h</td></tr>
        <tr><td class="lbl">CIF orden (${c.horasTotalesOrden}h × ${fmt(c.cifPorHora)})</td><td class="val bold">${fmt(c.cifTotalOrden)}</td></tr>
      </table>

      <div class="paso-label">Paso 3 — Costo total de la orden</div>
      <table class="data-table">
        <tr><td class="lbl">MOD</td><td class="val">${fmt(c.costoMODTotal)}</td></tr>
        <tr><td class="lbl">CIF</td><td class="val">${fmt(c.cifTotalOrden)}</td></tr>
        <tr><td class="lbl bold">Total orden</td><td class="val bold">${fmt(c.costoTotalAcumulado)}</td></tr>
      </table>

      <div class="paso-label">Paso 4 — Costo unitario</div>
      <table class="data-table" style="margin-bottom:12px">
        <tr><td class="lbl">Unidades iniciales</td><td class="val">${h.inputs.cantidadInicial.toLocaleString("es-PY")} u.</td></tr>
        <tr><td class="lbl">Unidades finales (con merma)</td><td class="val">${c.unidadesFinales.toLocaleString("es-PY")} u.</td></tr>
      </table>

      <div class="resultado-box">
        <div class="resultado-cols">
          <div class="resultado-item">
            <div class="resultado-label">Costo unitario final</div>
            <div class="resultado-value">${fmt(c.costoUnitarioFinal)}</div>
            <div class="resultado-sub">${fmt(c.costoTotalAcumulado)} / ${c.unidadesFinales.toLocaleString("es-PY")} u.</div>
          </div>
          <div class="resultado-item right">
            <div class="resultado-label">MOD por unidad</div>
            <div class="resultado-val-sm">${fmt(c.costoMODTotal / c.unidadesFinales)}</div>
            <div class="resultado-label" style="margin-top:8px">CIF por unidad</div>
            <div class="resultado-val-sm">${fmt(c.cifTotalOrden / c.unidadesFinales)}</div>
          </div>
        </div>
      </div>

      <div class="paso-label" style="margin-top:14px">Desglose por proceso</div>
      <table class="proc-table">
        <thead><tr><th>#</th><th>Proceso</th><th class="right">Merma</th><th class="right">Entrada</th><th class="right">Salida</th><th class="right">MOD</th><th class="right">CIF</th><th class="right">Costo unit.</th></tr></thead>
        <tbody>${procesosRows}</tbody>
      </table>
    </div>`;
  }).join("");

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Costos FasonFarma ${fecha}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Helvetica Neue',Arial,sans-serif;font-size:9.5px;color:#1a1a2e;background:#fff;padding:28px;max-width:780px;margin:0 auto}
  .doc-title{font-size:17px;font-weight:700;color:#0f1117;letter-spacing:-.03em;margin-bottom:3px}
  .doc-sub{font-size:9.5px;color:#6b7280;margin-bottom:3px}
  .doc-date{font-size:8.5px;color:#9ca3af;margin-bottom:24px}
  .section-label{font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:#9ca3af;margin-bottom:7px;padding-bottom:3px;border-bottom:1px solid #e5e7eb}
  .cif-block{margin-bottom:24px}
  .cif-table{width:100%;border-collapse:collapse;margin-bottom:7px}
  .cif-table td{padding:4px 0;font-size:9px;border-bottom:1px solid #f3f4f6}
  .cif-table td.right{text-align:right;color:#374151}
  .cif-total{display:flex;justify-content:space-between;padding:6px 0;border-top:1.5px solid #1a1a2e;font-weight:700;font-size:9.5px}
  .cif-meta{font-size:8.5px;color:#6b7280;margin-top:4px}
  .orden-block{margin-bottom:32px}
  .orden-block.page-break{page-break-before:always;padding-top:24px}
  .orden-header{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px}
  .orden-num{font-size:13px;font-weight:700;color:#0f1117}
  .orden-fecha{font-size:8.5px;color:#9ca3af}
  .orden-meta{font-size:9px;color:#6b7280;margin-bottom:10px}
  .orden-meta strong{color:#1a1a2e}
  .paso-label{font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#4a7fa5;margin-bottom:5px;margin-top:10px;padding:3px 0;border-bottom:1px solid #dbeafe}
  .proc-table{width:100%;border-collapse:collapse;margin-bottom:8px}
  .proc-table th{font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#9ca3af;padding:4px 3px;border-bottom:1.5px solid #e5e7eb;text-align:left}
  .proc-table th.right{text-align:right}
  .proc-table td{padding:5px 3px;font-size:9px;border-bottom:1px solid #f3f4f6;color:#374151}
  .proc-table td.proc-num{color:#9ca3af;font-weight:700}
  .proc-table td.right{text-align:right}
  .proc-table td.bold{font-weight:700;color:#1a1a2e}
  .proc-table td.total-td{font-weight:700;color:#1a1a2e;border-top:1.5px solid #d1d5db;border-bottom:none}
  .proc-table tfoot td{padding:5px 3px}
  .data-table{width:100%;border-collapse:collapse;margin-bottom:8px}
  .data-table td{padding:4px 0;font-size:9px;border-bottom:1px solid #f3f4f6}
  .data-table td.lbl{color:#6b7280;width:55%}
  .data-table td.val{color:#1a1a2e;font-weight:500}
  .data-table td.bold{font-weight:700}
  .resultado-box{background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px 16px}
  .resultado-cols{display:flex;justify-content:space-between;align-items:center}
  .resultado-item{flex:1}
  .resultado-item.right{text-align:right}
  .resultado-label{font-size:7.5px;text-transform:uppercase;letter-spacing:.1em;color:#6b7280;margin-bottom:3px}
  .resultado-value{font-size:22px;font-weight:800;color:#1a1a2e;letter-spacing:-.03em}
  .resultado-val-sm{font-size:12px;font-weight:700;color:#1a1a2e}
  .resultado-sub{font-size:8px;color:#9ca3af;margin-top:2px}
  .print-note{margin-top:20px;padding:8px;background:#f9fafb;border:1px solid #e5e7eb;border-radius:4px;font-size:8.5px;color:#6b7280;text-align:center}
  @media print{.print-note{display:none}.orden-block.page-break{page-break-before:always}}
</style>
</head>
<body>
  <div class="doc-title">Estimacion de Costos de Produccion</div>
  <div class="doc-sub">FasonFarma — Distribucion CIF por horas de planta</div>
  <div class="doc-date">Generado: ${fecha}</div>
  <div class="cif-block">
    <div class="section-label">CIF Mensual</div>
    <table class="cif-table">${cifRows}</table>
    <div class="cif-total"><span>Total CIF mensual</span><span>${fmt(totalCIF)}</span></div>
    <div class="cif-meta">Jornada: ${horasDia}h/dia x ${diasMes} dias = ${horasMes}h/mes &nbsp;|&nbsp; CIF por hora: ${fmt(totalCIF / horasMes)}/h</div>
  </div>
  <div class="section-label">Ordenes — ${historial.length} ${historial.length === 1 ? "orden" : "ordenes"}</div>
  ${ordenesHTML}
  <div class="print-note">Para guardar como PDF: Archivo → Imprimir → Guardar como PDF</div>
</body>
</html>`;

  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Costos_FasonFarma_${new Date().toISOString().slice(0, 10)}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function EstimadorCostos() {
  const [step, setStep] = useState(1);

  // CIF
  const [rubros, setRubros] = useState(CIF_RUBROS_DEFAULT);
  const [horasDia, setHorasDia] = useState("9.5");
  const [diasMes, setDiasMes] = useState("22");
  const [cifGuardado, setCifGuardado] = useState(false);

  // Orden
  const [producto, setProducto] = useState("");
  const [cantidadInicial, setCantidadInicial] = useState("");
  const [procesos, setProcesos] = useState([newProceso()]);

  const [calculo, setCalculo] = useState(null);
  const [historial, setHistorial] = useState([]);

  const totalCIF = rubros.reduce((s, r) => s + (parseFloat(r.monto) || 0), 0);
  const horasMes = (parseFloat(horasDia) || 0) * (parseFloat(diasMes) || 0);
  const cifPorHoraPreview = horasMes > 0 && totalCIF > 0 ? totalCIF / horasMes : 0;

  const updateRubro = (id, field, val) => setRubros(r => r.map(x => x.id === id ? { ...x, [field]: val } : x));
  const addRubro = () => setRubros(r => [...r, { id: nextCifId++, nombre: "", monto: "" }]);
  const removeRubro = (id) => setRubros(r => r.filter(x => x.id !== id));

  const guardarCIF = () => {
    if (totalCIF <= 0) { alert("Ingresá al menos un rubro con monto."); return; }
    if (horasMes <= 0) { alert("Ingresá horas por día y días hábiles válidos."); return; }
    setCifGuardado(true);
    setStep(2);
  };

  const addProceso = () => setProcesos(p => [...p, newProceso()]);
  const removeProceso = (id) => setProcesos(p => p.filter(x => x.id !== id));
  const updateProceso = (id, field, val) => setProcesos(p => p.map(x => x.id === id ? { ...x, [field]: val } : x));

  const calcular = useCallback(() => {
    if (!producto.trim()) { alert("Ingresá el nombre del producto."); return; }
    const cant = parseFloat(cantidadInicial);
    if (!cant || cant <= 0) { alert("Ingresá la cantidad inicial."); return; }
    const procs = procesos.filter(p => parseFloat(p.horasMOD) > 0 && parseFloat(p.tarifaMOD) > 0);
    if (procs.length === 0) { alert("Ingresá al menos un proceso con horas y tarifa MOD."); return; }

    const result = calcularProcesos(procs, cant, totalCIF, horasMes);
    setCalculo(result);
    setHistorial(h => [...h, {
      timestamp: Date.now(),
      inputs: { producto, cantidadInicial: cant, procesos: procs },
      calculo: result,
    }]);
    setStep(3);
  }, [producto, cantidadInicial, procesos, totalCIF, horasMes]);

  const resetOrden = () => {
    setProducto("");
    setCantidadInicial("");
    setProcesos([newProceso()]);
    setCalculo(null);
    setStep(2);
  };

  return (
    <>
      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0c0f18; }
        .app { min-height: 100vh; background: #0c0f18; color: #dde1ea; font-family: 'Inter','Segoe UI',system-ui,sans-serif; padding: 2rem 1rem 3rem; }
        .header { max-width: 780px; margin: 0 auto 2rem; text-align: center; }
        .eyebrow { font-size: .68rem; letter-spacing: .18em; text-transform: uppercase; color: #4a7fa5; margin-bottom: .5rem; }
        .title { font-size: 1.75rem; font-weight: 800; letter-spacing: -.03em; color: #dde1ea; }
        .title em { font-style: normal; color: #5ba3f5; }
        .subtitle { font-size: .82rem; color: #4b5568; margin-top: .4rem; }
        .steps-nav { display: flex; justify-content: center; max-width: 480px; margin: 0 auto 2rem; }
        .step-item { flex: 1; text-align: center; position: relative; }
        .step-item:not(:last-child)::after { content: ''; position: absolute; top: 14px; right: -50%; width: 100%; height: 1px; background: #1e2535; z-index: 0; }
        .step-circle { width: 28px; height: 28px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: .72rem; font-weight: 700; position: relative; z-index: 1; }
        .step-done .step-circle { background: #1a3a5c; color: #5ba3f5; border: 1px solid #2a5a8c; }
        .step-active .step-circle { background: #5ba3f5; color: #0c0f18; border: 1px solid #5ba3f5; }
        .step-pending .step-circle { background: #141926; color: #374151; border: 1px solid #1e2535; }
        .step-label { display: block; font-size: .62rem; margin-top: .3rem; letter-spacing: .05em; text-transform: uppercase; }
        .step-done .step-label { color: #5ba3f5; }
        .step-active .step-label { color: #dde1ea; }
        .step-pending .step-label { color: #374151; }
        .main { max-width: 780px; margin: 0 auto; display: flex; flex-direction: column; gap: 1.25rem; }
        .card { background: #141926; border: 1px solid #1e2535; border-radius: 14px; padding: 1.5rem; }
        .card-active { border-color: #2a3f5f; }
        .sec-header { display: flex; align-items: center; gap: .6rem; margin-bottom: 1.25rem; }
        .sec-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
        .sec-title { font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .12em; color: #6b7280; }
        .sec-badge { margin-left: auto; font-size: .62rem; font-weight: 700; padding: .2rem .55rem; border-radius: 4px; }
        .badge-ok { background: #0b2218; color: #34d399; border: 1px solid #064e2e; }
        .badge-warn { background: #1c1a0b; color: #fbbf24; border: 1px solid #4e3d06; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: .75rem; }
        .field { margin-bottom: .9rem; }
        .pinput-field { margin-bottom: 0; }
        .flabel { display: block; font-size: .72rem; color: #5b6475; margin-bottom: .3rem; font-weight: 500; }
        .fhint { font-size: .67rem; color: #374151; margin-bottom: .3rem; line-height: 1.4; }
        .finput-wrap { position: relative; display: flex; align-items: center; }
        .finput { width: 100%; padding: .52rem 2.2rem .52rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; transition: border-color .15s; }
        .finput:focus { border-color: #5ba3f5; }
        .finput::placeholder { color: #2d3748; }
        .funit { position: absolute; right: .6rem; font-size: .65rem; color: #374151; pointer-events: none; }
        .fselect { width: 100%; padding: .52rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; cursor: pointer; }
        .fselect:focus { border-color: #5ba3f5; }
        .info-box { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .65rem .9rem; font-size: .72rem; color: #4b5568; line-height: 1.8; margin-bottom: .9rem; }
        .info-box strong { color: #5ba3f5; }
        .rubro-header { display: grid; grid-template-columns: 1fr 140px 28px; gap: .5rem; margin-bottom: .4rem; }
        .col-lbl { font-size: .63rem; color: #374151; text-transform: uppercase; letter-spacing: .08em; }
        .col-lbl.right { text-align: right; }
        .rubro-row { display: grid; grid-template-columns: 1fr 140px 28px; gap: .5rem; align-items: center; margin-bottom: .55rem; }
        .rubro-text { padding: .5rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; width: 100%; }
        .rubro-text:focus { border-color: #5ba3f5; }
        .rubro-text::placeholder { color: #2d3748; }
        .rubro-num { padding: .5rem .7rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 7px; color: #dde1ea; font-size: .85rem; outline: none; width: 100%; text-align: right; }
        .rubro-num:focus { border-color: #5ba3f5; }
        .rubro-num::placeholder { color: #2d3748; }
        .btn-remove { width: 28px; height: 28px; border-radius: 6px; border: 1px solid #1e2535; background: #0c0f18; color: #4b5568; cursor: pointer; font-size: 1rem; display: flex; align-items: center; justify-content: center; transition: all .15s; flex-shrink: 0; }
        .btn-remove:hover { border-color: #ef4444; color: #ef4444; }
        .cif-total-row { display: flex; justify-content: space-between; align-items: center; padding: .6rem .85rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; margin-top: .75rem; }
        .cif-total-lbl { font-size: .75rem; color: #5b6475; }
        .cif-total-val { font-size: .95rem; font-weight: 700; color: #5ba3f5; }
        .proceso-card { background: #0f1520; border: 1px solid #1e2535; border-radius: 10px; padding: 1rem 1rem .75rem; margin-bottom: .75rem; position: relative; }
        .proceso-card-header { display: flex; align-items: center; gap: .5rem; margin-bottom: .85rem; }
        .proceso-num { width: 20px; height: 20px; border-radius: 50%; background: #1a2744; border: 1px solid #2a3f5f; display: flex; align-items: center; justify-content: center; font-size: .65rem; font-weight: 700; color: #5ba3f5; flex-shrink: 0; }
        .proceso-tipo-label { font-size: .75rem; font-weight: 600; color: #9ca3af; flex: 1; }
        .proceso-grid { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr 1fr; gap: .6rem; align-items: end; }
        .proceso-remove { position: absolute; top: .75rem; right: .75rem; }
        .divider { height: 1px; background: #1e2535; margin: 1rem 0; }
        .hint { font-size: .7rem; color: #374151; line-height: 1.5; margin-bottom: .75rem; }
        .btn-add { width: 100%; padding: .42rem; border: 1px dashed #1e2535; border-radius: 7px; background: transparent; color: #4b5568; font-size: .78rem; cursor: pointer; transition: all .15s; margin-bottom: .75rem; }
        .btn-add:hover { border-color: #5ba3f5; color: #5ba3f5; }
        .btn-primary { width: 100%; padding: .72rem; background: #5ba3f5; color: #0c0f18; border: none; border-radius: 8px; font-size: .85rem; font-weight: 800; cursor: pointer; transition: background .15s; margin-top: .5rem; }
        .btn-primary:hover { background: #4a8de0; }
        .btn-primary:disabled { opacity: .35; cursor: not-allowed; }
        .btn-ghost { width: 100%; padding: .55rem; background: transparent; color: #4b5568; border: 1px solid #1e2535; border-radius: 8px; font-size: .78rem; cursor: pointer; margin-top: .6rem; transition: all .15s; }
        .btn-ghost:hover { color: #dde1ea; border-color: #374151; }
        .btn-excel { width: 100%; padding: .65rem; background: #0b2e14; color: #34d399; border: 1px solid #0d5c2a; border-radius: 8px; font-size: .82rem; font-weight: 700; cursor: pointer; margin-top: .6rem; transition: all .15s; }
        .btn-excel:hover { background: #0d3d1b; border-color: #34d399; }
        .btn-pdf { width: 100%; padding: .65rem; background: #1a1226; color: #c084fc; border: 1px solid #4c1d95; border-radius: 8px; font-size: .82rem; font-weight: 700; cursor: pointer; margin-top: .6rem; transition: all .15s; }
        .btn-pdf:hover { background: #221533; border-color: #c084fc; }
        .locked-box { background: #0a1a0f; border: 1px solid #0d2e1a; border-radius: 8px; padding: .7rem 1rem; display: flex; justify-content: space-between; align-items: center; }
        .locked-val { font-size: .8rem; color: #34d399; font-weight: 600; }
        .locked-edit { font-size: .72rem; color: #5ba3f5; cursor: pointer; background: none; border: none; padding: 0; text-decoration: underline; }
        .result-hero { text-align: center; padding: 1.5rem 0 1.25rem; border-bottom: 1px solid #1e2535; margin-bottom: 1.25rem; }
        .result-eyebrow { font-size: .65rem; letter-spacing: .15em; text-transform: uppercase; color: #4a7fa5; margin-bottom: .4rem; }
        .result-value { font-size: 2.8rem; font-weight: 900; color: #5ba3f5; letter-spacing: -.04em; line-height: 1; }
        .result-label { font-size: .75rem; color: #4b5568; margin-top: .35rem; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: .75rem; margin-bottom: 1.25rem; }
        .stat-box { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .65rem .75rem; text-align: center; }
        .stat-val { font-size: .88rem; font-weight: 700; color: #dde1ea; }
        .stat-lbl { font-size: .6rem; color: #374151; text-transform: uppercase; letter-spacing: .06em; margin-top: .2rem; }

        /* Resumen paso a paso */
        .resumen-pasos { display: flex; flex-direction: column; gap: .75rem; margin-bottom: 1.25rem; }
        .paso-block { background: #0c0f18; border: 1px solid #1e2535; border-radius: 8px; padding: .75rem 1rem; }
        .paso-header { font-size: .68rem; font-weight: 700; text-transform: uppercase; letter-spacing: .1em; color: #5ba3f5; margin-bottom: .6rem; }
        .paso-row { display: flex; justify-content: space-between; font-size: .78rem; padding: .25rem 0; border-bottom: 1px solid #141926; }
        .paso-row:last-child { border-bottom: none; }
        .paso-lbl { color: #6b7280; }
        .paso-val { color: #dde1ea; font-weight: 600; font-family: monospace; }
        .paso-val.blue { color: #5ba3f5; font-size: .85rem; }
        .paso-formula { font-size: .72rem; color: #374151; font-family: monospace; margin-top: .4rem; padding: .4rem .6rem; background: #141926; border-radius: 5px; }

        .proc-result-label { font-size: .68rem; color: #374151; text-transform: uppercase; letter-spacing: .08em; margin-bottom: .6rem; }
        .proc-result-table { width: 100%; border-collapse: collapse; margin-bottom: 1rem; font-size: .78rem; }
        .proc-result-table th { font-size: .63rem; color: #374151; text-transform: uppercase; letter-spacing: .06em; padding: .4rem .5rem; border-bottom: 1px solid #1e2535; text-align: left; }
        .proc-result-table th.right { text-align: right; }
        .proc-result-table td { padding: .5rem .5rem; color: #9ca3af; border-bottom: 1px solid #0f1420; }
        .proc-result-table td.right { text-align: right; }
        .proc-result-table td.highlight { color: #dde1ea; font-weight: 700; }
        .proc-result-table td.tipo { color: #dde1ea; }
        .proc-result-table tr:last-child td { border-bottom: none; }
        .total-row { display: flex; justify-content: space-between; align-items: center; padding-top: .9rem; border-top: 1px solid #1e2535; margin-bottom: 1.25rem; }
        .total-lbl { font-size: .8rem; color: #5b6475; }
        .total-val { font-size: 1rem; font-weight: 800; color: #dde1ea; }
        .historial-badge { display: inline-flex; align-items: center; gap: .4rem; padding: .25rem .65rem; background: #0c0f18; border: 1px solid #1e2535; border-radius: 20px; font-size: .7rem; color: #6b7280; margin-top: .5rem; }
        .historial-badge span { color: #5ba3f5; font-weight: 700; }
        .export-center { text-align: center; }
        @media (max-width: 600px) {
          .proceso-grid { grid-template-columns: 1fr 1fr; }
          .stats-grid { grid-template-columns: 1fr 1fr; }
          .rubro-row, .rubro-header { grid-template-columns: 1fr 110px 28px; }
          .two-col { grid-template-columns: 1fr; }
        }
      `}</style>

      <div className="app">
        <div className="header">
          <p className="eyebrow">FasonFarma — Costos</p>
          <h1 className="title">Costo <em>unitario</em> por orden</h1>
          <p className="subtitle">MOD + CIF por horas de planta · Multiples procesos en cascada</p>
        </div>

        <div className="steps-nav">
          {[{ n: 1, label: "CIF" }, { n: 2, label: "Orden" }, { n: 3, label: "Resultado" }].map(s => {
            const cls = step > s.n ? "step-done" : step === s.n ? "step-active" : "step-pending";
            return (
              <div key={s.n} className={`step-item ${cls}`}>
                <div className="step-circle">{step > s.n ? "✓" : s.n}</div>
                <span className="step-label">{s.label}</span>
              </div>
            );
          })}
        </div>

        <div className="main">

          {/* CIF */}
          <div className={`card ${step === 1 ? "card-active" : ""}`}>
            <div className="sec-header">
              <div className="sec-dot" style={{ background: cifGuardado ? "#34d399" : "#fbbf24" }} />
              <span className="sec-title">Costos Indirectos de Fabricacion — mensual</span>
              <span className={`sec-badge ${cifGuardado ? "badge-ok" : "badge-warn"}`}>{cifGuardado ? "Guardado" : "Pendiente"}</span>
            </div>
            {cifGuardado ? (
              <div className="locked-box">
                <span className="locked-val">
                  CIF: {fmt(totalCIF)} / mes &nbsp;|&nbsp; {horasMes}h/mes &nbsp;|&nbsp; {fmt(cifPorHoraPreview)}/h
                </span>
                <button className="locked-edit" onClick={() => { setCifGuardado(false); setStep(1); }}>Editar</button>
              </div>
            ) : (
              <>
                <div className="two-col" style={{ marginBottom: ".9rem" }}>
                  <div className="field" style={{ marginBottom: 0 }}>
                    <label className="flabel">Horas por dia</label>
                    <div className="finput-wrap">
                      <input type="number" className="finput" placeholder="9.5" min="1" step="0.5"
                        value={horasDia} onChange={e => setHorasDia(e.target.value)} />
                      <span className="funit">h</span>
                    </div>
                  </div>
                  <div className="field" style={{ marginBottom: 0 }}>
                    <label className="flabel">Dias habiles por mes</label>
                    <p className="fhint">Solo lunes a viernes</p>
                    <div className="finput-wrap">
                      <input type="number" className="finput" placeholder="22" min="1" step="1"
                        value={diasMes} onChange={e => setDiasMes(e.target.value)} />
                      <span className="funit">dias</span>
                    </div>
                  </div>
                </div>
                {horasMes > 0 && (
                  <div className="info-box">
                    Horas mensuales de planta: <strong>{horasMes}h</strong>
                    {cifPorHoraPreview > 0 && <> &nbsp;|&nbsp; CIF por hora: <strong>{fmt(cifPorHoraPreview)}/h</strong></>}
                  </div>
                )}
                <div className="divider" />
                <p className="hint">Ingresá cada gasto fijo mensual de tu planta.</p>
                <div className="rubro-header">
                  <span className="col-lbl">Rubro</span>
                  <span className="col-lbl right">Monto mensual</span>
                  <span />
                </div>
                {rubros.map(r => (
                  <div key={r.id} className="rubro-row">
                    <input type="text" className="rubro-text" placeholder="Ej: Alquiler" value={r.nombre} onChange={e => updateRubro(r.id, "nombre", e.target.value)} />
                    <input type="number" className="rubro-num" placeholder="0" min="0" step="any" value={r.monto} onChange={e => updateRubro(r.id, "monto", e.target.value)} />
                    <button className="btn-remove" onClick={() => removeRubro(r.id)}>×</button>
                  </div>
                ))}
                <button className="btn-add" onClick={addRubro}>+ Agregar rubro</button>
                <div className="cif-total-row">
                  <span className="cif-total-lbl">Total CIF mensual</span>
                  <span className="cif-total-val">{fmt(totalCIF)}</span>
                </div>
                <button className="btn-primary" onClick={guardarCIF}>Guardar y continuar</button>
              </>
            )}
          </div>

          {/* Orden */}
          {step >= 2 && (
            <div className={`card ${step === 2 ? "card-active" : ""}`}>
              <div className="sec-header">
                <div className="sec-dot" style={{ background: "#5ba3f5" }} />
                <span className="sec-title">Datos de la orden</span>
              </div>
              {step === 3 ? (
                <div className="locked-box">
                  <span className="locked-val">{historial[historial.length - 1]?.inputs.producto} — {historial[historial.length - 1]?.inputs.procesos.length} procesos</span>
                  <button className="locked-edit" onClick={resetOrden}>Nueva orden</button>
                </div>
              ) : (
                <>
                  <div className="field">
                    <label className="flabel">Producto</label>
                    <input type="text" className="finput" style={{ paddingRight: ".75rem" }}
                      placeholder="Ej: Amoxicilina 500mg" value={producto}
                      onChange={e => setProducto(e.target.value)} />
                  </div>
                  <div className="field">
                    <label className="flabel">Cantidad inicial de la orden</label>
                    <div className="finput-wrap">
                      <input type="number" className="finput" placeholder="Ej: 3200" min="1" step="any"
                        value={cantidadInicial} onChange={e => setCantidadInicial(e.target.value)} />
                      <span className="funit">u.</span>
                    </div>
                  </div>
                  <div className="divider" />
                  <div className="sec-header" style={{ marginBottom: ".75rem" }}>
                    <div className="sec-dot" style={{ background: "#a78bfa" }} />
                    <span className="sec-title">Procesos</span>
                  </div>
                  <p className="hint">La merma se aplica en cascada. La salida de cada proceso es la entrada del siguiente.</p>
                  {procesos.map((p, idx) => (
                    <div key={p.id} className="proceso-card">
                      <div className="proceso-card-header">
                        <div className="proceso-num">{idx + 1}</div>
                        <span className="proceso-tipo-label">{p.tipo}</span>
                        {procesos.length > 1 && (
                          <button className="btn-remove proceso-remove" onClick={() => removeProceso(p.id)}>×</button>
                        )}
                      </div>
                      <div className="proceso-grid">
                        <div className="pinput-field">
                          <label className="flabel">Tipo de proceso</label>
                          <select className="fselect" value={p.tipo} onChange={e => updateProceso(p.id, "tipo", e.target.value)}>
                            {TIPOS_PROCESO.map(t => <option key={t}>{t}</option>)}
                          </select>
                        </div>
                        <NumInput label="Personas" value={p.personas} onChange={v => updateProceso(p.id, "personas", v)} placeholder="1" unit="p." min="1" step="1" />
                        <NumInput label="Horas MOD" value={p.horasMOD} onChange={v => updateProceso(p.id, "horasMOD", v)} placeholder="8" unit="h" step="0.5" />
                        <NumInput label="Tarifa MOD" value={p.tarifaMOD} onChange={v => updateProceso(p.id, "tarifaMOD", v)} placeholder="Gs./h" unit="Gs." />
                        <NumInput label="% Merma" value={p.pctMerma} onChange={v => updateProceso(p.id, "pctMerma", v)} placeholder="0" unit="%" step="0.5" />
                      </div>
                    </div>
                  ))}
                  <button className="btn-add" onClick={addProceso}>+ Agregar proceso</button>
                  <button className="btn-primary" disabled={!cifGuardado} onClick={calcular}>Calcular costo unitario</button>
                </>
              )}
            </div>
          )}

          {/* Resultado */}
          {step === 3 && calculo && (
            <div className="card card-active">
              <div className="sec-header">
                <div className="sec-dot" style={{ background: "#5ba3f5" }} />
                <span className="sec-title">Resultado — {historial[historial.length - 1]?.inputs.producto}</span>
              </div>

              <div className="result-hero">
                <p className="result-eyebrow">Costo unitario final</p>
                <p className="result-value">{fmt(calculo.costoUnitarioFinal)}</p>
                <p className="result-label">por unidad neta · {calculo.resultados.length} {calculo.resultados.length === 1 ? "proceso" : "procesos"}</p>
              </div>

              <div className="stats-grid">
                <div className="stat-box">
                  <p className="stat-val">{historial[historial.length - 1]?.inputs.cantidadInicial.toLocaleString("es-PY")}</p>
                  <p className="stat-lbl">Unid. iniciales</p>
                </div>
                <div className="stat-box">
                  <p className="stat-val">{calculo.unidadesFinales.toLocaleString("es-PY")}</p>
                  <p className="stat-lbl">Unid. finales</p>
                </div>
                <div className="stat-box">
                  <p className="stat-val">{calculo.horasTotalesOrden}h</p>
                  <p className="stat-lbl">Horas orden</p>
                </div>
                <div className="stat-box">
                  <p className="stat-val">{fmt(calculo.cifPorHora)}</p>
                  <p className="stat-lbl">CIF / hora</p>
                </div>
              </div>

              {/* Resumen paso a paso */}
              <p className="proc-result-label" style={{ marginBottom: ".75rem" }}>Resumen paso a paso</p>
              <div className="resumen-pasos">

                <div className="paso-block">
                  <p className="paso-header">Paso 1 — MOD total de la orden</p>
                  {calculo.resultados.map(r => (
                    <div className="paso-row" key={r.idx}>
                      <span className="paso-lbl">Proceso {r.idx} — {r.tipo}</span>
                      <span className="paso-val">{r.personas} × {r.horas}h × {fmt(r.tarifa)}/h = {fmt(r.costoMOD)}</span>
                    </div>
                  ))}
                  <div className="paso-row" style={{ borderTop: "1px solid #1e2535", marginTop: ".25rem", paddingTop: ".35rem" }}>
                    <span className="paso-lbl" style={{ fontWeight: 700, color: "#9ca3af" }}>Total MOD</span>
                    <span className="paso-val blue">{fmt(calculo.costoMODTotal)}</span>
                  </div>
                </div>

                <div className="paso-block">
                  <p className="paso-header">Paso 2 — CIF asignado a la orden</p>
                  <div className="paso-row">
                    <span className="paso-lbl">Horas totales de la orden</span>
                    <span className="paso-val">{calculo.horasTotalesOrden} h</span>
                  </div>
                  <div className="paso-row">
                    <span className="paso-lbl">CIF por hora de planta</span>
                    <span className="paso-val">{fmt(calculo.cifPorHora)}/h</span>
                  </div>
                  <div className="paso-formula">{fmt(totalCIF)} / {horasMes}h = {fmt(calculo.cifPorHora)}/h</div>
                  <div className="paso-row" style={{ marginTop: ".4rem" }}>
                    <span className="paso-lbl" style={{ fontWeight: 700, color: "#9ca3af" }}>CIF orden ({calculo.horasTotalesOrden}h × {fmt(calculo.cifPorHora)})</span>
                    <span className="paso-val blue">{fmt(calculo.cifTotalOrden)}</span>
                  </div>
                </div>

                <div className="paso-block">
                  <p className="paso-header">Paso 3 — Costo total de la orden</p>
                  <div className="paso-row">
                    <span className="paso-lbl">MOD</span>
                    <span className="paso-val">{fmt(calculo.costoMODTotal)}</span>
                  </div>
                  <div className="paso-row">
                    <span className="paso-lbl">CIF</span>
                    <span className="paso-val">{fmt(calculo.cifTotalOrden)}</span>
                  </div>
                  <div className="paso-row" style={{ borderTop: "1px solid #1e2535", marginTop: ".25rem", paddingTop: ".35rem" }}>
                    <span className="paso-lbl" style={{ fontWeight: 700, color: "#9ca3af" }}>Total orden</span>
                    <span className="paso-val blue">{fmt(calculo.costoTotalAcumulado)}</span>
                  </div>
                </div>

                <div className="paso-block">
                  <p className="paso-header">Paso 4 — Costo unitario</p>
                  <div className="paso-row">
                    <span className="paso-lbl">Unidades finales (con merma)</span>
                    <span className="paso-val">{calculo.unidadesFinales.toLocaleString("es-PY")} u.</span>
                  </div>
                  <div className="paso-formula">{fmt(calculo.costoTotalAcumulado)} / {calculo.unidadesFinales.toLocaleString("es-PY")} u. = {fmt(calculo.costoUnitarioFinal)}</div>
                  <div className="paso-row" style={{ marginTop: ".4rem" }}>
                    <span className="paso-lbl" style={{ fontWeight: 700, color: "#9ca3af" }}>Costo unitario final</span>
                    <span className="paso-val blue" style={{ fontSize: "1rem" }}>{fmt(calculo.costoUnitarioFinal)}</span>
                  </div>
                </div>

              </div>

              {/* Desglose por proceso */}
              <p className="proc-result-label">Desglose por proceso</p>
              <table className="proc-result-table">
                <thead>
                  <tr>
                    <th>#</th><th>Proceso</th>
                    <th className="right">Entrada</th><th className="right">Salida</th>
                    <th className="right">MOD</th><th className="right">CIF</th>
                    <th className="right">Costo unit.</th>
                  </tr>
                </thead>
                <tbody>
                  {calculo.resultados.map(r => (
                    <tr key={r.idx}>
                      <td style={{ color: "#5ba3f5", fontWeight: 700 }}>{r.idx}</td>
                      <td className="tipo">{r.tipo}</td>
                      <td className="right">{r.unidadesEntrada.toLocaleString("es-PY")}</td>
                      <td className="right">{r.unidadesSalida.toLocaleString("es-PY")}</td>
                      <td className="right">{fmt(r.costoMOD)}</td>
                      <td className="right">{fmt(r.cifAsignado)}</td>
                      <td className="right highlight">{fmt(r.costoUnitario)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>

              <div className="total-row">
                <span className="total-lbl">Costo total acumulado de la orden</span>
                <span className="total-val">{fmt(calculo.costoTotalAcumulado)}</span>
              </div>

              <div className="export-center">
                <div className="historial-badge">
                  Ordenes en esta sesion: <span>{historial.length}</span>
                </div>
              </div>

              <button className="btn-excel" onClick={() => exportarExcel(historial, rubros, totalCIF, parseFloat(horasDia), parseFloat(diasMes))}>
                Exportar a Excel — {historial.length} {historial.length === 1 ? "orden" : "ordenes"}{historial.length > 1 ? " + Resumen" : ""}
              </button>
              <button className="btn-pdf" onClick={() => exportarPDF(historial, rubros, totalCIF, parseFloat(horasDia), parseFloat(diasMes))}>
                Descargar HTML → guardar como PDF
              </button>
              <button className="btn-ghost" onClick={resetOrden}>Calcular otra orden</button>
            </div>
          )}

        </div>
      </div>
    </>
  );
}

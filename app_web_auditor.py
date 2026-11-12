"""
=============================================================================
SERVIDOR WEB + BOT DE TELEGRAM INTEGRADO - AUDITORÍA CONTABLE CEA AUTO ORIENTE
=============================================================================
Incluye:
- Motor de Consolidación Multi-Curso por Cédula (231 clientes con 2+ cursos simultáneos).
- Separación inteligente entre Riesgo Cero Real ($0 en todos sus cursos) y Combos Multi-Curso
  (2da matrícula en $0 pero con abono en su 1ra matrícula).
- Carga dual: Archivo de Matrículas (`matriculas.xlsx` / `convenios.xlsx`) + Archivo opcional
  de Recibos de Caja (`ingresos.xlsx`) para cazar pagos huérfanos sin matrícula.
- Parametrización en vivo de Tarifas Base por Categoría (A1, A2, B1, C1, C2).
- Fila desplegable (Expander) al hacer clic en cualquier alumno para ver todas sus
  matrículas cruzadas, recibos de caja y saldo global consolidado por Cédula.
"""

import datetime
import io
import json
import logging
import os
import re
import threading
import time
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from dotenv import load_dotenv

import auditor_cea


def _parse_multipart_binary(content_type: str, raw_body: bytes) -> Dict[str, Tuple[Optional[str], bytes]]:
    """
    Parser binario exacto para multipart/form-data.
    Garantiza que ningún byte (\r\n) dentro de archivos .xlsx/.zip sea alterado por políticas MIME de texto.
    """
    m = re.search(r'boundary=(?:"([^"]+)"|([^;\s]+))', content_type, flags=re.IGNORECASE)
    if not m:
        return {}
    boundary_str = m.group(1) or m.group(2)
    delim = ("--" + boundary_str).encode("utf-8")

    parts_out: Dict[str, Tuple[Optional[str], bytes]] = {}
    chunks = raw_body.split(delim)
    for chunk in chunks[1:]:
        if chunk.startswith(b"--"):
            break
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        elif chunk.startswith(b"\n"):
            chunk = chunk[1:]

        if chunk.endswith(b"\r\n"):
            chunk = chunk[:-2]
        elif chunk.endswith(b"\n"):
            chunk = chunk[:-1]

        sep_idx = chunk.find(b"\r\n\r\n")
        sep_len = 4
        if sep_idx == -1:
            sep_idx = chunk.find(b"\n\n")
            sep_len = 2
        if sep_idx == -1:
            continue

        header_text = chunk[:sep_idx].decode("utf-8", errors="ignore")
        payload_bytes = chunk[sep_idx + sep_len :]

        name_m = re.search(r'\bname="([^"]+)"', header_text, flags=re.IGNORECASE)
        if not name_m:
            continue
        field_name = name_m.group(1)
        fn_m = re.search(r'\bfilename="([^"]*)"', header_text, flags=re.IGNORECASE)
        filename = fn_m.group(1).strip() if fn_m and fn_m.group(1).strip() else None
        parts_out[field_name] = (filename, payload_bytes)
    return parts_out

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

PORT = int(os.getenv("PORT", "10000"))
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "").strip()
DEFAULT_EXCEL_PATH = BASE_DIR / "matriculas11112026.xlsx"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("auditor_web")

STATE: Dict[str, Any] = {
    "filename": "matriculas11112026.xlsx",
    "ingresos_filename": "",
    "periodicidad": "ANUAL",
    "updated_at": "",
    "raw_matriculas_bytes": b"",
    "raw_ingresos_bytes": None,
    "tarifas_base": dict(auditor_cea.TARIFAS_BASE_DEFAULT),
    "hallazgos": None,
    "excel_bytes": b"",
    "pdf_bytes": b"",
}


def procesar_y_actualizar_estado(
    origen: Any,
    filename: str = "matriculas11112026.xlsx",
    periodicidad: str = "ANUAL",
    origen_ingresos: Any = None,
    ingresos_filename: str = "",
    tarifas_base: Optional[Dict[str, float]] = None,
    config_override: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    if isinstance(origen, (str, Path)):
        mat_bytes = Path(origen).read_bytes()
    elif isinstance(origen, bytes):
        mat_bytes = origen
    else:
        mat_bytes = origen.read()

    ing_bytes = None
    if origen_ingresos:
        if isinstance(origen_ingresos, (str, Path)):
            ing_bytes = Path(origen_ingresos).read_bytes()
        elif isinstance(origen_ingresos, bytes):
            ing_bytes = origen_ingresos
        else:
            ing_bytes = origen_ingresos.read()

    tarifas_usar = tarifas_base or STATE.get("tarifas_base") or dict(auditor_cea.TARIFAS_BASE_DEFAULT)

    registros, metadata = auditor_cea.cargar_y_limpiar_datos(
        mat_bytes,
        origen_ingresos=ing_bytes,
        tarifas_base=tarifas_usar,
    )
    hallazgos = auditor_cea.auditar_riesgos(registros, metadata=metadata, config=config_override)
    excel_bytes = auditor_cea.generar_reporte_excel_bytes(hallazgos, tipo_informe=periodicidad)
    pdf_bytes = auditor_cea.generar_reporte_pdf_bytes(hallazgos, tipo_informe=periodicidad)

    STATE["filename"] = filename
    STATE["ingresos_filename"] = ingresos_filename
    STATE["periodicidad"] = periodicidad
    STATE["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    STATE["raw_matriculas_bytes"] = mat_bytes
    STATE["raw_ingresos_bytes"] = ing_bytes
    STATE["tarifas_base"] = tarifas_usar
    STATE["hallazgos"] = hallazgos
    STATE["excel_bytes"] = excel_bytes
    STATE["pdf_bytes"] = pdf_bytes
    logger.info(
        "✅ Estado actualizado: '%s' (%d matrículas, %d clientes únicos, %d multi-curso)",
        filename,
        len(registros),
        hallazgos["kpis"]["total_clientes_unicos"],
        hallazgos["kpis"]["clientes_multicurso_count"],
    )
    return hallazgos


if DEFAULT_EXCEL_PATH.exists():
    try:
        procesar_y_actualizar_estado(DEFAULT_EXCEL_PATH, filename="matriculas11112026.xlsx", periodicidad="ANUAL")
    except Exception as exc:
        logger.warning("No se pudo precargar matriculas11112026.xlsx: %s", exc)


HTML_DASHBOARD = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Sistema de Auditoría Contable y Control Multi-Curso | CEA Auto Oriente</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <style>
    :root {
      --bg: #0e1626;
      --card: #131d31;
      --card-border: #1e2d4a;
      --text: #f8fafc;
      --muted: #94a3b8;
      --primary: #1f77b4;
      --primary-neon: #3b82f6;
      --orange: #ff7f0e;
      --danger: #d62728;
      --success: #10b981;
      --warning: #f59e0b;
      --purple: #8b5cf6;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding-bottom: 60px;
    }
    header {
      background: linear-gradient(135deg, #0b1324 0%, #172a4f 100%);
      border-bottom: 1px solid var(--card-border);
      padding: 18px 28px;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      position: sticky;
      top: 0;
      z-index: 50;
      box-shadow: 0 4px 20px rgba(0,0,0,0.45);
    }
    .brand { display: flex; align-items: center; gap: 14px; }
    .brand-badge {
      width: 46px; height: 46px; border-radius: 12px;
      background: linear-gradient(135deg, #3b82f6, #1d4ed8);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 19px;
      box-shadow: 0 4px 12px rgba(59,130,246,0.45);
    }
    .brand h1 { font-size: 19px; font-weight: 700; letter-spacing: -0.02em; }
    .brand p { font-size: 12.5px; color: #cbd5e1; }
    .header-actions { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .btn {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 9px 15px; border-radius: 8px;
      font-size: 13px; font-weight: 600; cursor: pointer;
      border: 1px solid transparent; transition: all 0.15s ease;
      text-decoration: none; color: #fff;
    }
    .btn-primary { background: var(--primary-neon); }
    .btn-primary:hover { background: #2563eb; }
    .btn-success { background: #059669; }
    .btn-success:hover { background: #047857; }
    .btn-danger { background: #dc2626; }
    .btn-danger:hover { background: #b91c1c; }
    .btn-outline { background: rgba(255,255,255,0.06); border-color: rgba(255,255,255,0.18); }
    .btn-outline:hover { background: rgba(255,255,255,0.14); }

    .container { max-width: 1480px; margin: 22px auto; padding: 0 22px; }

    .control-bar {
      display: grid;
      grid-template-columns: 1.4fr 1fr;
      gap: 16px;
      margin-bottom: 18px;
    }
    @media (max-width: 1024px) { .control-bar { grid-template-columns: 1fr; } }

    .panel {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 16px 20px;
    }
    .panel-title {
      font-size: 13px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 0.04em; color: var(--muted);
      margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center;
    }
    .upload-grid {
      display: grid;
      grid-template-columns: 1fr 1fr auto;
      gap: 10px;
      align-items: end;
    }
    @media (max-width: 768px) { .upload-grid { grid-template-columns: 1fr; } }
    .field-label { font-size: 11.5px; color: var(--muted); margin-bottom: 4px; display: block; font-weight: 600; }
    .upload-row { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    input[type="file"], select, input[type="text"], input[type="email"], input[type="number"] {
      background: #09101d; border: 1px solid #283a5e; color: var(--text);
      padding: 8px 11px; border-radius: 8px; font-size: 12.5px; width: 100%;
    }

    /* Drawer de Tarifas Parametrizables */
    .tarifas-drawer {
      background: #10192b;
      border: 1px solid #263859;
      border-radius: 10px;
      padding: 14px 18px;
      margin-bottom: 18px;
      display: none;
    }
    .tarifas-drawer.open { display: block; }
    .tarifas-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 12px;
      margin-top: 10px;
      align-items: end;
    }

    /* Grid de KPIs */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(205px, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }
    .kpi-card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 16px 18px;
      position: relative;
      overflow: hidden;
    }
    .kpi-card::before {
      content: ""; position: absolute; top: 0; left: 0; width: 4px; height: 100%;
      background: var(--primary-neon);
    }
    .kpi-card.success::before { background: var(--success); }
    .kpi-card.danger::before { background: var(--danger); }
    .kpi-card.warning::before { background: var(--warning); }
    .kpi-card.purple::before { background: var(--purple); }
    .kpi-card.orange::before { background: var(--orange); }
    .kpi-label { font-size: 11.5px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
    .kpi-value { font-size: 23px; font-weight: 800; margin: 5px 0 3px; letter-spacing: -0.02em; }
    .kpi-sub { font-size: 12px; color: var(--muted); }

    /* Gráficas */
    .charts-grid {
      display: grid;
      grid-template-columns: 1.55fr 1fr;
      gap: 16px;
      margin-bottom: 20px;
    }
    @media (max-width: 1024px) { .charts-grid { grid-template-columns: 1fr; } }
    .chart-box { height: 300px; position: relative; }

    /* Pills / Pestañas */
    .tabs-header {
      display: flex; flex-wrap: wrap; gap: 8px;
      border-bottom: 1px solid var(--card-border);
      margin-bottom: 14px; padding-bottom: 12px;
    }
    .tab-btn {
      background: #1a263e; color: #cbd5e1;
      border: 1px solid #2b3d63;
      padding: 8px 14px; border-radius: 8px;
      font-size: 12.5px; font-weight: 600; cursor: pointer;
      display: inline-flex; align-items: center; gap: 7px;
      transition: all 0.15s;
    }
    .tab-btn:hover { background: #263859; color: #fff; }
    .tab-btn.active {
      background: var(--primary-neon);
      border-color: var(--primary-neon);
      color: #fff;
      box-shadow: 0 2px 10px rgba(59,130,246,0.35);
    }
    .badge {
      padding: 2px 7px; border-radius: 999px;
      font-size: 11px; font-weight: 700;
      background: rgba(255,255,255,0.16);
    }
    .badge-danger { background: rgba(239,68,68,0.22); color: #fca5a5; border: 1px solid rgba(239,68,68,0.45); }
    .badge-success { background: rgba(16,185,129,0.22); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.45); }
    .badge-warning { background: rgba(245,158,11,0.22); color: #fde68a; border: 1px solid rgba(245,158,11,0.45); }
    .badge-purple { background: rgba(139,92,246,0.22); color: #ddd6fe; border: 1px solid rgba(139,92,246,0.45); }
    .badge-info { background: rgba(59,130,246,0.22); color: #93c5fd; border: 1px solid rgba(59,130,246,0.45); }

    .table-toolbar {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 12px; gap: 12px; flex-wrap: wrap;
    }
    .table-wrap {
      overflow-x: auto; max-height: 580px; overflow-y: auto;
      border: 1px solid var(--card-border); border-radius: 10px;
    }
    table { width: 100%; border-collapse: collapse; font-size: 12.5px; text-align: left; }
    thead th {
      position: sticky; top: 0; background: #18243c; color: #e2e8f0;
      font-weight: 700; padding: 11px 12px; border-bottom: 1px solid #2b3d63;
      white-space: nowrap; z-index: 10;
    }
    tbody td { padding: 10px 12px; border-bottom: 1px solid #1a263e; white-space: nowrap; }
    tbody tr.data-row { cursor: pointer; transition: background 0.12s; }
    tbody tr.data-row:hover { background: rgba(59, 130, 246, 0.11); }
    tbody tr.expander-row { background: #0b1220; display: none; }
    tbody tr.expander-row.open { display: table-row; }
    .expander-box {
      padding: 14px 18px;
      border-left: 3px solid var(--primary-neon);
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      white-space: normal;
    }
    @media (max-width: 900px) { .expander-box { grid-template-columns: 1fr; } }
    .mini-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 6px; }
    .mini-table th, .mini-table td { padding: 6px 8px; border: 1px solid #1e2d4a; }
    .mini-table th { background: #131d31; color: #93c5fd; }

    .money-pos { color: #34d399; font-weight: 600; }
    .money-neg { color: #f87171; font-weight: 700; }
    .status-toast { margin-top: 10px; padding: 9px 13px; border-radius: 8px; font-size: 12.5px; display: none; }
    .status-toast.ok { display: block; background: rgba(16,185,129,0.15); border: 1px solid #059669; color: #6ee7b7; }
    .status-toast.err { display: block; background: rgba(239,68,68,0.15); border: 1px solid #dc2626; color: #fca5a5; }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="brand-badge">CEA</div>
      <div>
        <h1>Sistema de Auditoría Contable y Control de Cartera — Sede CEA</h1>
        <p id="headerMeta">CEA Auto Oriente • Cargando motor de consolidación por Cédula...</p>
      </div>
    </div>
    <div class="header-actions">
      <button type="button" class="btn btn-outline" onclick="toggleTarifas()">⚙️ Tarifas por Categoría</button>
      <a href="/api/download/excel" class="btn btn-success">📊 Descargar Excel Multi-Pestaña (.xlsx)</a>
      <a href="/api/download/pdf" class="btn btn-danger">📄 Descargar PDF Ejecutivo</a>
    </div>
  </header>

  <div class="container">
    <!-- Panel desplegable de Parametrización de Tarifas -->
    <div id="tarifasDrawer" class="tarifas-drawer">
      <div class="panel-title">
        <span>⚙️ Parametrización de Tarifas Estándar por Categoría (COP)</span>
        <span style="font-size:11.5px; color:var(--muted);">Se aplica cuando una matrícula no tiene tarifa pactada explícita</span>
      </div>
      <div class="tarifas-grid">
        <div><label class="field-label">Moto A1</label><input type="number" id="tar_A1" value="800000" step="10000" /></div>
        <div><label class="field-label">Moto A2</label><input type="number" id="tar_A2" value="1000000" step="10000" /></div>
        <div><label class="field-label">Carro Particular B1</label><input type="number" id="tar_B1" value="1250000" step="10000" /></div>
        <div><label class="field-label">Servicio Público C1</label><input type="number" id="tar_C1" value="1450000" step="10000" /></div>
        <div><label class="field-label">Camión / C2</label><input type="number" id="tar_C2" value="1600000" step="10000" /></div>
        <div><button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="guardarTarifas()">💾 Aplicar Tarifas</button></div>
      </div>
    </div>

    <!-- Barra de Ingestión Dual y Despacho -->
    <div class="control-bar">
      <div class="panel">
        <div class="panel-title">
          <span>1. Ingestión de Matrículas (`convenios.xlsx`) + Recibos de Caja (`ingresos.xlsx`)</span>
          <span id="currentFileBadge" class="badge badge-info">matriculas11112026.xlsx</span>
        </div>
        <form id="uploadForm">
          <div class="upload-grid">
            <div>
              <label class="field-label">📋 Archivo de Matrículas / Convenios (.xlsx / .csv) *</label>
              <input type="file" id="fileInput" name="file" accept=".xlsx,.xls,.csv" required />
            </div>
            <div>
              <label class="field-label">🧾 Archivo de Ingresos / Recibos de Caja (Opcional)</label>
              <input type="file" id="ingresosInput" name="ingresos_file" accept=".xlsx,.xls,.csv" />
            </div>
            <div style="display:flex; gap:8px;">
              <select id="periodoInput" name="periodicidad" style="width:auto;">
                <option value="ANUAL">Anual</option>
                <option value="MENSUAL">Mensual</option>
                <option value="SEMANAL">Semanal</option>
              </select>
              <button type="submit" class="btn btn-primary">⚡ Auditar y Cruzar</button>
            </div>
          </div>
        </form>
        <div id="uploadToast" class="status-toast"></div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>4. Despacho Automático de Alertas (Email / Telegram)</span>
        </div>
        <div class="upload-row" style="margin-bottom: 8px;">
          <input type="email" id="emailDestInput" placeholder="Correo destino (ej. gerencia@autooriente.com)" style="flex:1;" />
          <button type="button" onclick="despacharEmail()" class="btn btn-outline">📧 Enviar Correo</button>
        </div>
        <div class="upload-row">
          <input type="text" id="tgChatInput" placeholder="Chat ID de Telegram (opcional si está en .env)" style="flex:1;" />
          <button type="button" onclick="despacharTelegram()" class="btn btn-primary">✈️ Enviar a Telegram</button>
        </div>
        <div id="dispatchToast" class="status-toast"></div>
      </div>
    </div>

    <!-- Tarjetas de KPIs Ejecutivos -->
    <div class="kpi-grid" id="kpiGrid"></div>

    <!-- Fila Superior de Gráficos -->
    <div class="charts-grid">
      <div class="panel">
        <div class="panel-title">
          <span>Gráfico 1: Matrículas vs. Cartera Pendiente por Mes (Enero a Noviembre 2026)</span>
        </div>
        <div class="chart-box">
          <canvas id="chartMensual"></canvas>
        </div>
      </div>
      <div class="panel">
        <div class="panel-title">
          <span>Gráfico 2: Distribución de la Cartera por Antigüedad de Mora</span>
        </div>
        <div class="chart-box">
          <canvas id="chartCartera"></canvas>
        </div>
      </div>
    </div>

    <!-- Fila de Filtros Rápidos (Pills Interactivos) + Tabla Maestra -->
    <div class="panel">
      <div class="tabs-header" id="tabsHeader">
        <button class="tab-btn active" onclick="cambiarTab('jt_cartera', this)">
          💛 Cartera Javier Torres <span class="badge badge-warning" id="cnt_jt">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('abono_cero_real', this)">
          🚨 Riesgo $0 REAL (Suspensión) <span class="badge badge-danger" id="cnt_cero_real">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('abono_cero_multicurso', this)">
          🔄 Falsos $0 (Combos Multi-Curso) <span class="badge badge-info" id="cnt_cero_mc">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('clientes_multicurso', this)">
          👥 Clientes Multi-Curso (Consolidado Cédula) <span class="badge badge-purple" id="cnt_mc_all">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('cartera_critica_90', this)">
          ⏳ Mora Crítica (&gt;90 Días) <span class="badge badge-danger" id="cnt_mora90">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('mora_30_90', this)">
          ⚠️ Mora 31 a 90 Días <span class="badge badge-warning" id="cnt_mora30">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('errores_digitacion', this)">
          🔍 Bandera 3: Errores / Saldos Negativos <span class="badge badge-danger" id="cnt_err">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('mensual', this)">
          📈 Cumplimiento Mensual de Metas
        </button>
        <button class="tab-btn" onclick="cambiarTab('convenios', this)">
          📊 Consolidado por Convenios
        </button>
      </div>

      <div class="table-toolbar">
        <input type="text" id="searchInput" placeholder="🔎 Buscar en vivo por cédula, nombre, matrícula, categoría o convenio..." oninput="renderizarTablaActiva()" style="max-width: 460px;" />
        <span id="tableCountInfo" style="font-size: 12.5px; color: var(--muted);">💡 Haz clic sobre cualquier alumno para desplegar sus cursos hermanos y recibos de caja</span>
      </div>

      <div class="table-wrap">
        <table id="auditTable">
          <thead id="auditThead"></thead>
          <tbody id="auditTbody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <script>
    let AUDIT_DATA = null;
    let ACTIVE_TAB = 'jt_cartera';
    let chartMensualInstance = null;
    let chartCarteraInstance = null;

    const fmtCOP = (val) => '$' + Math.round(Number(val || 0)).toLocaleString('es-CO');

    function toggleTarifas() {
      document.getElementById('tarifasDrawer').classList.toggle('open');
    }

    async function guardarTarifas() {
      const tarifas = {
        A1: Number(document.getElementById('tar_A1').value || 800000),
        A2: Number(document.getElementById('tar_A2').value || 1000000),
        B1: Number(document.getElementById('tar_B1').value || 1250000),
        C1: Number(document.getElementById('tar_C1').value || 1450000),
        C2: Number(document.getElementById('tar_C2').value || 1600000),
      };
      const res = await fetch('/api/tarifas', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tarifas })
      });
      const data = await res.json();
      if (data.ok) {
        AUDIT_DATA = data;
        actualizarVista();
        toggleTarifas();
      }
    }

    async function cargarDatos() {
      const res = await fetch('/api/data');
      AUDIT_DATA = await res.json();
      actualizarVista();
    }

    function actualizarVista() {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const h = AUDIT_DATA.hallazgos;
      const k = h.kpis;
      const m = h.metadata || {};

      const badgeTxt = AUDIT_DATA.ingresos_filename
        ? `${AUDIT_DATA.filename} + ${AUDIT_DATA.ingresos_filename}`
        : (AUDIT_DATA.filename || 'matriculas.xlsx');
      document.getElementById('currentFileBadge').textContent = badgeTxt;

      document.getElementById('headerMeta').textContent =
        `Sede: ${m.sede || 'CEA'} • Periodo: ${m.fecha_inicial || '01/01/2026'} al ${m.fecha_final || '11/11/2026'} • ${k.total_matriculas} Matrículas (${k.total_clientes_unicos} Clientes Únicos | ${k.clientes_multicurso_count} Multi-Curso)`;

      // Contadores de pestañas
      document.getElementById('cnt_jt').textContent = h.jt_cartera.length;
      document.getElementById('cnt_cero_real').textContent = h.abono_cero_real.length;
      document.getElementById('cnt_cero_mc').textContent = h.abono_cero_multicurso.length;
      document.getElementById('cnt_mc_all').textContent = h.clientes_multicurso.length;
      document.getElementById('cnt_mora90').textContent = h.cartera_critica_90.length;
      document.getElementById('cnt_mora30').textContent = (h.mora_media_60.length + h.mora_temprana_30.length);
      document.getElementById('cnt_err').textContent = h.errores_digitacion.length;

      // Renderizar KPIs
      const kpisHtml = `
        <div class="kpi-card">
          <div class="kpi-label">Total Facturado</div>
          <div class="kpi-value">${fmtCOP(k.total_facturado)}</div>
          <div class="kpi-sub"><b>${k.total_matriculas} matrículas</b> (${k.total_clientes_unicos} cédulas únicas)</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">Recaudo Efectivo en Caja</div>
          <div class="kpi-value" style="color:#34d399;">${fmtCOP(k.total_abonos)}</div>
          <div class="kpi-sub">Tasa Recaudo: <b>${k.tasa_recaudo_pct}%</b> (${k.alumnos_paz_y_salvo} al día)</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">Cartera Pendiente Total</div>
          <div class="kpi-value" style="color:#f87171;">${fmtCOP(k.total_pendiente)}</div>
          <div class="kpi-sub"><b>${k.alumnos_con_saldo} matrículas</b> con saldo (${k.tasa_morosidad_pct}%)</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">🚨 Riesgo $0 REAL (Suspender)</div>
          <div class="kpi-value" style="color:#fca5a5;">${k.abono_cero_real_count} Mat. (${k.abono_cero_real_clientes} Clientes)</div>
          <div class="kpi-sub">$0 en todos sus cursos: <b>${fmtCOP(k.abono_cero_real_monto)}</b></div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🔄 Clientes Multi-Curso (2+ Cursos)</div>
          <div class="kpi-value" style="color:#c4b5fd;">${k.clientes_multicurso_count} Clientes</div>
          <div class="kpi-sub"><b>${k.abono_cero_multicurso_count} falsos $0</b> (abonaron en su 1ra matrícula)</div>
        </div>
        <div class="kpi-card orange">
          <div class="kpi-label">💛 Cartera Javier Torres</div>
          <div class="kpi-value" style="color:#fdba74;">${fmtCOP(k.jt_pendiente)}</div>
          <div class="kpi-sub"><b>${k.jt_con_saldo_count} mat.</b> de ${k.jt_total_matriculas} (${k.jt_abono_cero_real_count} en $0 real)</div>
        </div>
      `;
      document.getElementById('kpiGrid').innerHTML = kpisHtml;

      renderizarGraficas(h);
      renderizarTablaActiva();
    }

    function renderizarGraficas(h) {
      const mesesValidos = h.mensual.filter(x => x.mes !== 'SIN_FECHA');
      const labels = mesesValidos.map(x => x.nombre_mes);
      const dataMat = mesesValidos.map(x => x.total_matriculas);
      const dataJT = mesesValidos.map(x => x.matriculas_jt);
      const dataCarteraMillones = mesesValidos.map(x => Math.round(x.pendiente / 1000000 * 10) / 10);

      if (chartMensualInstance) chartMensualInstance.destroy();
      const ctx1 = document.getElementById('chartMensual').getContext('2d');
      chartMensualInstance = new Chart(ctx1, {
        type: 'bar',
        data: {
          labels,
          datasets: [
            {
              label: 'Matrículas Totales Mes (Meta: 100)',
              data: dataMat,
              backgroundColor: 'rgba(59, 130, 246, 0.82)',
              borderColor: '#3b82f6',
              borderWidth: 1,
              borderRadius: 5,
              yAxisID: 'y',
            },
            {
              label: 'Matrículas Javier Torres (Meta: 30)',
              data: dataJT,
              backgroundColor: 'rgba(139, 92, 246, 0.85)',
              borderColor: '#8b5cf6',
              borderWidth: 1,
              borderRadius: 5,
              yAxisID: 'y',
            },
            {
              label: 'Cartera Pendiente Mes ($ Millones COP)',
              data: dataCarteraMillones,
              type: 'line',
              borderColor: '#ff7f0e',
              backgroundColor: 'rgba(255, 127, 14, 0.2)',
              borderWidth: 2.5,
              pointRadius: 4,
              tension: 0.3,
              yAxisID: 'y1',
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 11.5 } } } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.04)' } },
            y: { position: 'left', ticks: { color: '#93c5fd' }, grid: { color: 'rgba(255,255,255,0.06)' }, title: { display: true, text: 'N° Matrículas', color: '#93c5fd' } },
            y1: { position: 'right', ticks: { color: '#fdba74', callback: v => '$' + v + 'M' }, grid: { drawOnChartArea: false }, title: { display: true, text: 'Cartera ($ Millones)', color: '#fdba74' } }
          }
        }
      });

      const k = h.kpis;
      if (chartCarteraInstance) chartCarteraInstance.destroy();
      const ctx2 = document.getElementById('chartCartera').getContext('2d');
      chartCarteraInstance = new Chart(ctx2, {
        type: 'doughnut',
        data: {
          labels: [
            `Crítica > 90 días (${k.mora_90_count})`,
            `Mora Media 61-90 días (${k.mora_60_count})`,
            `Mora Temprana 31-60 días (${k.mora_30_count})`,
            `Reciente <= 30 días (${k.mora_aldia_count})`
          ],
          datasets: [{
            data: [k.mora_90_monto, k.mora_60_monto, k.mora_30_monto, k.mora_aldia_monto],
            backgroundColor: ['#ef4444', '#ff7f0e', '#eab308', '#3b82f6'],
            borderWidth: 2,
            borderColor: '#131d31'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '62%',
          plugins: {
            legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11 } } }
          }
        }
      });
    }

    function cambiarTab(tabKey, btnEl) {
      ACTIVE_TAB = tabKey;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      if (btnEl) btnEl.classList.add('active');
      renderizarTablaActiva();
    }

    function toggleExpander(rowId) {
      const el = document.getElementById(rowId);
      if (el) el.classList.toggle('open');
    }

    function renderizarTablaActiva() {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const h = AUDIT_DATA.hallazgos;
      const q = (document.getElementById('searchInput').value || '').trim().toLowerCase();

      const thead = document.getElementById('auditThead');
      const tbody = document.getElementById('auditTbody');

      // Pestaña de Cumplimiento Mensual
      if (ACTIVE_TAB === 'mensual') {
        thead.innerHTML = `<tr>
          <th>Mes</th><th>Matrículas CEA</th><th>Meta CEA</th><th>Estado CEA</th>
          <th>Matrículas Javier Torres</th><th>Meta JT</th><th>Estado JT</th>
          <th>Total Facturado</th><th>Total Recaudado</th><th>Cartera Pendiente</th><th>Riesgo $0 Real</th>
        </tr>`;
        const rows = h.mensual.filter(m => !q || m.nombre_mes.toLowerCase().includes(q));
        document.getElementById('tableCountInfo').textContent = `Mostrando ${rows.length} periodos mensuales`;
        tbody.innerHTML = rows.map(m => `<tr>
          <td><b>${m.nombre_mes}</b></td>
          <td><b>${m.total_matriculas}</b></td>
          <td>${m.meta_empresa}</td>
          <td><span class="badge ${m.estado_meta_empresa === 'CUMPLIDA' ? 'badge-success' : 'badge-warning'}">${m.cumplimiento_empresa_pct}% (${m.estado_meta_empresa})</span></td>
          <td><b>${m.matriculas_jt}</b></td>
          <td>${m.meta_jt}</td>
          <td><span class="badge ${m.estado_meta_jt === 'CUMPLIDA' ? 'badge-success' : 'badge-warning'}">${m.cumplimiento_jt_pct}% (${m.estado_meta_jt})</span></td>
          <td>${fmtCOP(m.facturado)}</td>
          <td class="money-pos">${fmtCOP(m.recaudado)}</td>
          <td class="${m.pendiente > 0 ? 'money-neg' : ''}">${fmtCOP(m.pendiente)}</td>
          <td><span class="badge badge-danger">${m.abono_cero_real_count}</span></td>
        </tr>`).join('');
        return;
      }

      // Pestaña de Consolidado por Convenios
      if (ACTIVE_TAB === 'convenios') {
        thead.innerHTML = `<tr>
          <th>Convenio / Canal</th><th>Matrículas</th><th>Clientes Únicos</th><th>Total Facturado</th>
          <th>Recaudado en Caja</th><th>% Recaudo</th><th>Cartera Pendiente</th><th>% Morosidad</th>
          <th>Deudores Activos</th><th>Riesgo $0 Real</th>
        </tr>`;
        const rows = h.convenios.filter(c => !q || c.convenio.toLowerCase().includes(q));
        document.getElementById('tableCountInfo').textContent = `Mostrando ${rows.length} convenios activos`;
        tbody.innerHTML = rows.map(c => `<tr>
          <td><b>${c.convenio}</b></td>
          <td><b>${c.matriculas}</b></td>
          <td>${c.clientes_unicos}</td>
          <td>${fmtCOP(c.facturado)}</td>
          <td class="money-pos">${fmtCOP(c.recaudado)}</td>
          <td><span class="badge badge-success">${c.tasa_recaudo}%</span></td>
          <td class="${c.pendiente > 0 ? 'money-neg' : ''}">${fmtCOP(c.pendiente)}</td>
          <td><span class="badge ${c.tasa_morosidad > 10 ? 'badge-danger' : 'badge-warning'}">${c.tasa_morosidad}%</span></td>
          <td>${c.con_saldo}</td>
          <td><span class="badge badge-danger">${c.abono_cero_real}</span></td>
        </tr>`).join('');
        return;
      }

      // Pestaña de Clientes Multi-Curso Consolidados por Cédula
      if (ACTIVE_TAB === 'clientes_multicurso') {
        const rows = h.clientes_multicurso.filter(c => {
          if (!q) return true;
          return `${c.NUMERODOCUMENTO} ${c.ALUMNO} ${c.CONVENIO_NORM} ${c.CATEGORIAS} ${c.MATRICULAS}`.toLowerCase().includes(q);
        });
        const sumaSaldo = rows.reduce((a, b) => a + Number(b.SALDO_REAL_CLIENTE || 0), 0);
        document.getElementById('tableCountInfo').textContent =
          `Mostrando ${rows.length} Clientes Multi-Curso (2+ matrículas) | Saldo Real Consolidado: ${fmtCOP(sumaSaldo)} (Haz clic en una fila para ver el desglose)`;

        thead.innerHTML = `<tr>
          <th>Cédula / Doc</th><th>Alumno (Cliente)</th><th>Convenio</th><th>N° Cursos</th>
          <th>Combo Categorías</th><th>Matrículas Asociadas</th><th>Total Contratado (Cliente)</th>
          <th>Total Abonado (Caja Global)</th><th>Saldo Real Cliente</th><th>Antigüedad Máx.</th>
        </tr>`;

        tbody.innerHTML = rows.map((c, idx) => {
          const expId = `exp_mc_${idx}`;
          const cursosHtml = (c.CURSOS_DETALLE || []).map(cd => `
            <tr>
              <td><b>${cd.NUMEROMATRICULA}</b></td>
              <td><span class="badge">${cd.CATEGORIA}</span></td>
              <td>${cd.FECHAINSCRIPCION}</td>
              <td>${fmtCOP(cd.VALORCURSO)}</td>
              <td class="money-pos">${fmtCOP(cd.ABONOS_ORIGINAL)}</td>
              <td class="${cd.VALORPENDIENTE_ORIGINAL > 0 ? 'money-neg' : ''}">${fmtCOP(cd.VALORPENDIENTE_ORIGINAL)}</td>
              <td>${cd.ESTADO}</td>
            </tr>
          `).join('');

          const recibosHtml = (c.RECIBOS_DETALLE || []).map(rc => `
            <tr>
              <td><b>${rc.RECIBO}</b></td>
              <td>${rc.FECHA}</td>
              <td>${rc.NUMEROMATRICULA || '<span class="badge badge-warning">HUÉRFANO (Por Cédula)</span>'}</td>
              <td class="money-pos">${fmtCOP(rc.VALOR)}</td>
              <td>${rc.TIPOPAGO}</td>
            </tr>
          `).join('') || '<tr><td colspan="5" style="color:var(--muted);">Sin recibos registrados</td></tr>';

          return `
            <tr class="data-row" onclick="toggleExpander('${expId}')">
              <td><code>${c.NUMERODOCUMENTO}</code></td>
              <td><b>🔍 ${c.ALUMNO}</b></td>
              <td>${c.CONVENIO_NORM}</td>
              <td><span class="badge badge-purple">${c.CANTIDAD_CURSOS} cursos</span></td>
              <td><b>${c.CATEGORIAS}</b></td>
              <td>${c.MATRICULAS}</td>
              <td>${fmtCOP(c.TOTAL_CONTRATADO)}</td>
              <td class="money-pos">${fmtCOP(c.TOTAL_ABONADO)}</td>
              <td class="${c.SALDO_REAL_CLIENTE > 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(c.SALDO_REAL_CLIENTE)}</td>
              <td><span class="badge ${c.DIAS_ANTIGUEDAD_MAX > 90 ? 'badge-danger' : 'badge-warning'}">${c.DIAS_ANTIGUEDAD_MAX} días</span></td>
            </tr>
            <tr id="${expId}" class="expander-row">
              <td colspan="10">
                <div class="expander-box">
                  <div>
                    <b style="color:#93c5fd;">📚 Matrículas Simultáneas de ${c.ALUMNO} (Cédula: ${c.NUMERODOCUMENTO})</b>
                    <table class="mini-table">
                      <thead><tr><th>Matrícula</th><th>Cat.</th><th>Fecha Insc.</th><th>Valor Curso</th><th>Abono Reg.</th><th>Saldo Reg.</th><th>Estado</th></tr></thead>
                      <tbody>${cursosHtml}</tbody>
                    </table>
                  </div>
                  <div>
                    <b style="color:#6ee7b7;">🧾 Historial de Recibos de Caja y Abonos Aplicados a la Cédula</b>
                    <table class="mini-table">
                      <thead><tr><th>Recibo</th><th>Fecha</th><th>Matrícula</th><th>Valor Pagado</th><th>Medio / Concepto</th></tr></thead>
                      <tbody>${recibosHtml}</tbody>
                    </table>
                  </div>
                </div>
              </td>
            </tr>
          `;
        }).join('');
        return;
      }

      // Listados por Matrícula con Contexto Multi-Curso y Expander
      let sourceList = [];
      let showMotivo = false;
      if (ACTIVE_TAB === 'jt_cartera') sourceList = h.jt_cartera;
      else if (ACTIVE_TAB === 'abono_cero_real') sourceList = h.abono_cero_real;
      else if (ACTIVE_TAB === 'abono_cero_multicurso') sourceList = h.abono_cero_multicurso;
      else if (ACTIVE_TAB === 'cartera_critica_90') sourceList = h.cartera_critica_90;
      else if (ACTIVE_TAB === 'mora_30_90') sourceList = [...h.mora_media_60, ...h.mora_temprana_30];
      else if (ACTIVE_TAB === 'errores_digitacion') {
        sourceList = h.errores_digitacion;
        showMotivo = true;
      }

      const filtered = sourceList.filter(r => {
        if (!q) return true;
        const txt = `${r.NUMEROMATRICULA} ${r.NUMERODOCUMENTO} ${r.ALUMNO} ${r.CONVENIO_NORM} ${r.CATEGORIA} ${r.ETIQUETA_MULTICURSO || ''} ${r.MOTIVO_ERROR || ''}`.toLowerCase();
        return txt.includes(q);
      });

      const totalSaldo = filtered.reduce((acc, r) => acc + Number(r.VALORPENDIENTE || 0), 0);
      document.getElementById('tableCountInfo').textContent =
        `Mostrando ${filtered.length} registros | Suma Saldo Pendiente: ${fmtCOP(totalSaldo)} • 💡 Clic en un alumno para ver cruce de cursos y recibos`;

      thead.innerHTML = `<tr>
        <th>Matrícula</th><th>Cédula / Doc</th><th>Alumno</th><th>Convenio</th>
        <th>Fecha Insc.</th><th>Cat.</th><th>Valor Curso</th><th>Abono Mat.</th>
        <th>Saldo Mat.</th><th>Contexto Cliente (Multi-Curso / Caja Global)</th>
        <th>Antigüedad</th><th>Estado</th>
        ${showMotivo ? '<th>Hallazgo / Error Detectado</th>' : ''}
      </tr>`;

      tbody.innerHTML = filtered.map((r, idx) => {
        const expId = `exp_row_${idx}`;
        let badgeCtx = `<span class="badge">${r.ETIQUETA_MULTICURSO || 'Curso Único'}</span>`;
        if (r.TIPO_RIESGO_CERO === 'RIESGO_CERO_REAL') {
          badgeCtx = `<span class="badge badge-danger">🚨 $0 Global (Sin abono en ningún curso)</span>`;
        } else if (r.TIPO_RIESGO_CERO === 'MULTICURSO_CON_ABONO') {
          badgeCtx = `<span class="badge badge-info">🔄 Combo ${r.CATEGORIAS_CLIENTE} (Abonó ${fmtCOP(r.TOTAL_ABONADO_CLIENTE)} en otra mat.)</span>`;
        } else if (r.ES_MULTICURSO) {
          badgeCtx = `<span class="badge badge-purple">👥 Multi-Curso (${r.CATEGORIAS_CLIENTE}) • Global: ${fmtCOP(r.SALDO_REAL_CLIENTE)}</span>`;
        }

        const cursosHtml = (r.CURSOS_HERMANOS || []).map(cd => `
          <tr>
            <td><b>${cd.NUMEROMATRICULA}</b></td>
            <td><span class="badge">${cd.CATEGORIA}</span></td>
            <td>${cd.FECHAINSCRIPCION}</td>
            <td>${fmtCOP(cd.VALORCURSO)}</td>
            <td class="money-pos">${fmtCOP(cd.ABONOS_ORIGINAL)}</td>
            <td class="${cd.VALORPENDIENTE_ORIGINAL > 0 ? 'money-neg' : ''}">${fmtCOP(cd.VALORPENDIENTE_ORIGINAL)}</td>
            <td>${cd.ESTADO}</td>
          </tr>
        `).join('');

        const recibosHtml = (r.RECIBOS_CLIENTE || []).map(rc => `
          <tr>
            <td><b>${rc.RECIBO}</b></td>
            <td>${rc.FECHA}</td>
            <td>${rc.NUMEROMATRICULA || '<span class="badge badge-warning">HUÉRFANO (Cédula)</span>'}</td>
            <td class="money-pos">${fmtCOP(rc.VALOR)}</td>
            <td>${rc.TIPOPAGO}</td>
          </tr>
        `).join('') || '<tr><td colspan="5" style="color:#fca5a5;">🚨 Sin recibos ni abonos registrados en caja ($0 Global)</td></tr>';

        return `
          <tr class="data-row" onclick="toggleExpander('${expId}')">
            <td><b>${r.NUMEROMATRICULA}</b></td>
            <td><code>${r.NUMERODOCUMENTO}</code></td>
            <td><b>🔍 ${r.ALUMNO}</b></td>
            <td>${r.CONVENIO_NORM}</td>
            <td>${r.FECHAINSCRIPCION}</td>
            <td><span class="badge">${r.CATEGORIA}</span></td>
            <td>${fmtCOP(r.VALORCURSO)}</td>
            <td class="${r.ABONOS === 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(r.ABONOS)}</td>
            <td class="money-neg">${fmtCOP(r.VALORPENDIENTE)}</td>
            <td>${badgeCtx}</td>
            <td><span class="badge ${r.DIAS_ANTIGUEDAD > 90 ? 'badge-danger' : (r.DIAS_ANTIGUEDAD > 30 ? 'badge-warning' : '')}">${r.DIAS_ANTIGUEDAD} días</span></td>
            <td>${r.ESTADO}</td>
            ${showMotivo ? `<td style="color:#fca5a5;font-weight:600;">${r.MOTIVO_ERROR}</td>` : ''}
          </tr>
          <tr id="${expId}" class="expander-row">
            <td colspan="${showMotivo ? 13 : 12}">
              <div class="expander-box">
                <div>
                  <b style="color:#93c5fd;">📊 Consolidado por Cédula (${r.NUMERODOCUMENTO}) — ${r.ALUMNO}</b>
                  <p style="font-size:12px; color:#cbd5e1; margin:4px 0;">
                    Total Contratado Cliente: <b>${fmtCOP(r.TOTAL_CONTRATADO_CLIENTE)}</b> |
                    Total Pagado Cliente: <b style="color:#34d399;">${fmtCOP(r.TOTAL_ABONADO_CLIENTE)}</b> |
                    Saldo Real Cliente: <b style="color:#f87171;">${fmtCOP(r.SALDO_REAL_CLIENTE)}</b>
                  </p>
                  <table class="mini-table">
                    <thead><tr><th>Matrícula</th><th>Cat.</th><th>Fecha Insc.</th><th>Valor Curso</th><th>Abono Reg.</th><th>Saldo Reg.</th><th>Estado</th></tr></thead>
                    <tbody>${cursosHtml}</tbody>
                  </table>
                </div>
                <div>
                  <b style="color:#6ee7b7;">🧾 Recibos de Caja Asociados a la Cédula (${r.NUMERODOCUMENTO})</b>
                  <table class="mini-table">
                    <thead><tr><th>N° Recibo</th><th>Fecha</th><th>Matrícula</th><th>Valor</th><th>Medio de Pago</th></tr></thead>
                    <tbody>${recibosHtml}</tbody>
                  </table>
                </div>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    document.getElementById('uploadForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fileEl = document.getElementById('fileInput');
      const ingEl = document.getElementById('ingresosInput');
      if (!fileEl.files.length) return;
      const fd = new FormData();
      fd.append('file', fileEl.files[0]);
      if (ingEl.files.length) {
        fd.append('ingresos_file', ingEl.files[0]);
      }
      fd.append('periodicidad', document.getElementById('periodoInput').value);

      const toast = document.getElementById('uploadToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Ejecutando motor contable y cruce por Cédula...';

      try {
        const res = await fetch('/api/upload', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.ok) {
          AUDIT_DATA = data;
          actualizarVista();
          toast.className = 'status-toast ok';
          const k = data.hallazgos.kpis;
          toast.textContent = `✅ Auditado: ${k.total_matriculas} matrículas (${k.total_clientes_unicos} clientes únicos, ${k.clientes_multicurso_count} multi-curso).`;
        } else {
          toast.className = 'status-toast err';
          toast.textContent = `❌ Error: ${data.error}`;
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = `❌ Error de red: ${err}`;
      }
    });

    async function despacharEmail() {
      const email = document.getElementById('emailDestInput').value.trim();
      const toast = document.getElementById('dispatchToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Enviando informe por correo SMTP...';
      try {
        const res = await fetch('/api/send/email', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '⚠️ ') + data.message;
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    async function despacharTelegram() {
      const chat_id = document.getElementById('tgChatInput').value.trim();
      const toast = document.getElementById('dispatchToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Enviando resumen + Excel + PDF a Telegram...';
      try {
        const res = await fetch('/api/send/telegram', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ chat_id })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '⚠️ ') + data.message;
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    cargarDatos();
  </script>
</body>
</html>
"""


class AuditorWebHandler(BaseHTTPRequestHandler):
    def _send_json(self, data: Dict[str, Any], status: int = 200) -> None:
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        if self.path in ("/", "/index.html"):
            raw = HTML_DASHBOARD.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if self.path == "/health":
            self._send_json({"ok": True, "service": "Auditor CEA Auto Oriente Multi-Curso"})
            return

        if self.path == "/api/data":
            self._send_json({
                "ok": True,
                "filename": STATE["filename"],
                "ingresos_filename": STATE["ingresos_filename"],
                "periodicidad": STATE["periodicidad"],
                "updated_at": STATE["updated_at"],
                "tarifas_base": STATE["tarifas_base"],
                "hallazgos": STATE["hallazgos"],
            })
            return

        if self.path.startswith("/api/download/excel"):
            excel_bytes = STATE.get("excel_bytes") or b""
            fecha_str = datetime.date.today().strftime("%Y%m%d")
            fname = f"Auditoria_CEA_{STATE['periodicidad']}_{fecha_str}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
            return

        if self.path.startswith("/api/download/pdf"):
            pdf_bytes = STATE.get("pdf_bytes") or b""
            fecha_str = datetime.date.today().strftime("%Y%m%d")
            fname = f"Resumen_Ejecutivo_CEA_{STATE['periodicidad']}_{fecha_str}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        self.send_response(404)
        self.end_headers()

    def do_POST(self) -> None:
        if self.path == "/api/upload":
            try:
                content_type = self.headers.get("Content-Type", "")
                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length) if length > 0 else b""
                parts = _parse_multipart_binary(content_type, raw_body)

                file_bytes = None
                filename = "matriculas.xlsx"
                ingresos_bytes = None
                ingresos_filename = ""
                periodicidad = "ANUAL"

                if "file" in parts:
                    fn, payload = parts["file"]
                    file_bytes = payload
                    if fn:
                        filename = fn

                if "ingresos_file" in parts:
                    fn_ing, payload_ing = parts["ingresos_file"]
                    # Si el usuario seleccionó por error el mismo archivo de matrículas en ambos campos,
                    # ignorar el segundo para no duplicar ni alterar el cruce.
                    if payload_ing and fn_ing and payload_ing != file_bytes:
                        ingresos_bytes = payload_ing
                        ingresos_filename = fn_ing

                if "periodicidad" in parts:
                    _, val_p = parts["periodicidad"]
                    if val_p:
                        periodicidad = val_p.decode("utf-8", errors="ignore").strip() or "ANUAL"

                if not file_bytes:
                    self._send_json({"ok": False, "error": "No se recibió el archivo de matrículas."}, 400)
                    return

                hallazgos = procesar_y_actualizar_estado(
                    file_bytes,
                    filename=filename,
                    periodicidad=periodicidad,
                    origen_ingresos=ingresos_bytes,
                    ingresos_filename=ingresos_filename,
                )
                self._send_json({
                    "ok": True,
                    "filename": filename,
                    "ingresos_filename": ingresos_filename,
                    "periodicidad": periodicidad,
                    "updated_at": STATE["updated_at"],
                    "hallazgos": hallazgos,
                })
            except Exception as exc:
                logger.exception("Error procesando archivos subidos")
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path == "/api/tarifas":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}
                nuevas_tarifas = body.get("tarifas", {})
                if STATE["raw_matriculas_bytes"]:
                    hallazgos = procesar_y_actualizar_estado(
                        STATE["raw_matriculas_bytes"],
                        filename=STATE["filename"],
                        periodicidad=STATE["periodicidad"],
                        origen_ingresos=STATE["raw_ingresos_bytes"],
                        ingresos_filename=STATE["ingresos_filename"],
                        tarifas_base=nuevas_tarifas,
                    )
                    self._send_json({
                        "ok": True,
                        "filename": STATE["filename"],
                        "ingresos_filename": STATE["ingresos_filename"],
                        "periodicidad": STATE["periodicidad"],
                        "hallazgos": hallazgos,
                    })
                else:
                    self._send_json({"ok": False, "error": "No hay archivo cargado."}, 400)
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path in ("/api/send/email", "/api/send/telegram"):
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}

            if not STATE["hallazgos"]:
                self._send_json({"ok": False, "message": "Primero carga un archivo de matrículas."}, 400)
                return

            if self.path == "/api/send/email":
                email_extra = [payload.get("email", "").strip()] if payload.get("email") else None
                ok, msg_out = auditor_cea.enviar_correo_auditoria(
                    STATE["hallazgos"],
                    STATE["excel_bytes"],
                    STATE["pdf_bytes"],
                    periodicidad=STATE["periodicidad"],
                    destinatarios_extra=email_extra,
                )
                self._send_json({"ok": ok, "message": msg_out})
                return

            if self.path == "/api/send/telegram":
                chat_id = payload.get("chat_id", "").strip() or None
                ok, msg_out = auditor_cea.enviar_telegram_auditoria_sync(
                    STATE["hallazgos"],
                    STATE["excel_bytes"],
                    STATE["pdf_bytes"],
                    periodicidad=STATE["periodicidad"],
                    chat_id=chat_id,
                )
                self._send_json({"ok": ok, "message": msg_out})
                return

        self.send_response(404)
        self.end_headers()

    def log_message(self, format: str, *args: Any) -> None:
        return


def _loop_keep_alive() -> None:
    while True:
        time.sleep(600)
        if RENDER_EXTERNAL_URL:
            try:
                urllib.request.urlopen(f"{RENDER_EXTERNAL_URL.rstrip('/')}/health", timeout=15).read()
            except Exception:
                pass


def main() -> None:
    server = ThreadingHTTPServer(("0.0.0.0", PORT), AuditorWebHandler)
    threading.Thread(target=_loop_keep_alive, daemon=True).start()
    logger.info("🌐 Dashboard Web Multi-Curso de Auditoría CEA activo en http://0.0.0.0:%s", PORT)
    server.serve_forever()


if __name__ == "__main__":
    main()

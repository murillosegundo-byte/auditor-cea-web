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
import hashlib
import hmac
import io
import json
import logging
import os
import re
import secrets
import threading
import time
import urllib.parse
import urllib.request
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from dotenv import load_dotenv

import auditor_cea
import clases_core


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

# =============================================================================
# SEGURIDAD Y CONTROL DE ACCESO (LOGIN GERENCIAL)
# =============================================================================
AUDITOR_USER = os.getenv("AUDITOR_USER", "admin").strip()
AUDITOR_PASSWORD = os.getenv("AUDITOR_PASSWORD", "AutoOriente2026*").strip()
SESSION_SECRET = os.getenv("SESSION_SECRET", secrets.token_hex(32))
COOKIE_NAME = "cea_auditor_session"


def _generar_token_sesion(usuario: str) -> str:
    exp = int(time.time()) + 86400 * 7  # 7 días de sesión activa
    payload = f"{usuario}:{exp}"
    firma = hmac.new(SESSION_SECRET.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    return f"{payload}:{firma}"


def _verificar_token_sesion(token: str) -> bool:
    if not token or token.count(":") != 2:
        return False
    usuario, exp_str, firma = token.split(":", 2)
    try:
        if int(exp_str) < int(time.time()):
            return False
    except ValueError:
        return False
    payload = f"{usuario}:{exp_str}"
    esperada = hmac.new(SESSION_SECRET.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    return hmac.compare_digest(firma, esperada) and usuario == AUDITOR_USER


HTML_LOGIN = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Acceso Restringido | Auditoría Contable CEA Auto Oriente</title>
  <style>
    :root {
      --bg: #0e1626;
      --card: #131d31;
      --border: #1e2d4a;
      --primary: #3b82f6;
      --text: #f8fafc;
      --muted: #94a3b8;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: radial-gradient(circle at top, #172a4f 0%, #0e1626 65%);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .login-card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 16px;
      width: 100%;
      max-width: 420px;
      padding: 34px 30px;
      box-shadow: 0 20px 50px rgba(0,0,0,0.6);
    }
    .logo-row {
      display: flex;
      align-items: center;
      gap: 14px;
      margin-bottom: 22px;
    }
    .badge-cea {
      width: 48px; height: 48px; border-radius: 12px;
      background: linear-gradient(135deg, #3b82f6, #1d4ed8);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 19px;
      box-shadow: 0 4px 14px rgba(59,130,246,0.45);
    }
    h1 { font-size: 18px; font-weight: 700; line-height: 1.3; }
    p.sub { font-size: 12.5px; color: var(--muted); margin-top: 3px; }
    .field { margin-bottom: 16px; }
    label { display: block; font-size: 12px; font-weight: 600; color: #cbd5e1; margin-bottom: 6px; }
    input {
      width: 100%;
      padding: 11px 13px;
      border-radius: 8px;
      border: 1px solid #283a5e;
      background: #09101d;
      color: #fff;
      font-size: 14px;
    }
    input:focus { outline: none; border-color: var(--primary); }
    button {
      width: 100%;
      padding: 12px;
      background: var(--primary);
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      margin-top: 6px;
      transition: background 0.15s;
    }
    button:hover { background: #2563eb; }
    .err-msg {
      background: rgba(239,68,68,0.15);
      border: 1px solid #dc2626;
      color: #fca5a5;
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 12.5px;
      margin-bottom: 16px;
      display: none;
    }
    .footer-note {
      margin-top: 20px;
      text-align: center;
      font-size: 11.5px;
      color: var(--muted);
      border-top: 1px solid var(--border);
      padding-top: 14px;
    }
  </style>
</head>
<body>
  <div class="login-card">
    <div class="logo-row">
      <div class="badge-cea">CEA</div>
      <div>
        <h1>Auditoría Contable y Cartera</h1>
        <p class="sub">CEA Auto Oriente • Acceso Gerencial</p>
      </div>
    </div>
    <div id="errBox" class="err-msg">Usuario o contraseña incorrectos.</div>
    <form id="loginForm">
      <div class="field">
        <label>Usuario Autorizado</label>
        <input type="text" id="userInput" name="username" placeholder="Ej: admin" required autocomplete="username" />
      </div>
      <div class="field">
        <label>Contraseña de Seguridad</label>
        <input type="password" id="passInput" name="password" placeholder="••••••••••••" required autocomplete="current-password" />
      </div>
      <button type="submit">🔒 Ingresar al Sistema</button>
    </form>
    <div class="footer-note">
      🛡️ Información financiera confidencial protegida por sesión cifrada HMAC-SHA256.
    </div>
  </div>
  <script>
    document.getElementById('loginForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      const errBox = document.getElementById('errBox');
      errBox.style.display = 'none';
      const username = document.getElementById('userInput').value.trim();
      const password = document.getElementById('passInput').value;
      try {
        const res = await fetch('/api/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, password })
        });
        const data = await res.json();
        if (data.ok) {
          window.location.href = '/';
        } else {
          errBox.textContent = data.error || 'Credenciales inválidas.';
          errBox.style.display = 'block';
        }
      } catch (err) {
        errBox.textContent = 'Error de conexión con el servidor.';
        errBox.style.display = 'block';
      }
    });
  </script>
</body>
</html>
"""

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

    /* Barra de Navegación Maestra de 5 Módulos */
    .mod-nav {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
      background: #0f172a;
      border: 1px solid #2b3d63;
      padding: 10px;
      border-radius: 14px;
    }
    .mod-btn {
      flex: 1 1 200px;
      background: #162238;
      color: #cbd5e1;
      border: 1px solid #2b3d63;
      padding: 11px 14px;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      transition: all 0.15s;
    }
    .mod-btn:hover { background: #223556; color: #fff; }
    .mod-btn.active {
      background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
      border-color: #60a5fa;
      color: #fff;
      box-shadow: 0 4px 16px rgba(37, 99, 235, 0.38);
    }
    .mod-section { display: none; }
    .mod-section.active { display: block; }
    .form-grid-4 {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
      gap: 10px;
      align-items: end;
    }
    .progress-bar-wrap {
      width: 110px;
      background: #0f172a;
      border: 1px solid #2b3d63;
      border-radius: 999px;
      height: 10px;
      overflow: hidden;
      display: inline-block;
      vertical-align: middle;
      margin-right: 6px;
    }
    .progress-bar-fill {
      height: 100%;
      background: linear-gradient(90deg, #3b82f6, #10b981);
    }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="brand-badge">CEA</div>
      <div>
        <h1>Plataforma Central CEA Auto Oriente — Cartera + Bot de Clases + PESV</h1>
        <p id="headerMeta">CEA Auto Oriente • Cargando motor de consolidación por Cédula y sincronización con Google Sheets...</p>
      </div>
    </div>
    <div class="header-actions">
      <button type="button" class="btn btn-outline" id="btnSheetsStatus" onclick="toggleSheetsDrawer()">🔑 Google Sheets</button>
      <button type="button" class="btn btn-outline" onclick="toggleTarifas()">⚙️ Tarifas por Categoría</button>
      <a href="/api/download/excel" class="btn btn-success">📊 Excel Cartera (.xlsx)</a>
      <a href="/api/download/pdf" class="btn btn-danger">📄 PDF Ejecutivo Cartera</a>
      <a href="/logout" class="btn btn-outline" title="Cerrar sesión segura">🔒 Salir</a>
    </div>
  </header>

  <div class="container">
    <!-- Panel desplegable para conectar Google Sheets en vivo -->
    <div id="sheetsDrawer" class="tarifas-drawer">
      <div class="panel-title">
        <span>🔑 Conexión en Tiempo Real con Google Sheets del Bot (`1xouy9DMkp4oZffZVx9E1y2IwY5qxRi5u2yr3p5wpj6U`)</span>
        <span id="sheetsStatusLabel" class="badge badge-info">Verificando conexión...</span>
      </div>
      <p style="font-size:12.5px; color:#cbd5e1; margin-bottom:10px;">
        El Dashboard y el Bot de Telegram comparten las mismas hojas de Google Sheets (<b>Alumnos</b>, <b>Instructores</b>, <b>Registro_Clases</b>, <b>Inasistencias</b>, <b>Mantenimientos</b>, <b>Vehiculos_PESV</b> y <b>Cierres_Jornada</b>). Si estás en Render y aún no has cargado <code>credentials.json</code>, súbelo o pégalo aquí con 1 clic:
      </p>
      <div class="form-grid-4">
        <div style="grid-column: span 2;">
          <label class="field-label">📂 Seleccionar archivo credentials.json de Google Service Account</label>
          <input type="file" id="credFileInput" accept=".json" onchange="leerArchivoCredenciales(this)" />
        </div>
        <div style="grid-column: span 2;">
          <label class="field-label">📋 O pegar contenido JSON de credentials.json</label>
          <input type="text" id="credJsonInput" placeholder='{"type": "service_account", "project_id": "..."}' />
        </div>
        <div>
          <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="conectarGoogleSheets()">🔌 Conectar Sheets Ahora</button>
        </div>
      </div>
      <div id="sheetsToast" class="status-toast"></div>
    </div>

    <!-- Barra de Navegación de los 5 Módulos Integrados -->
    <div class="mod-nav">
      <button type="button" class="mod-btn active" onclick="cambiarModulo('mod_cartera', this)">
        <span>💰 1. Auditoría de Cartera y Convenios</span>
        <span class="badge" id="nav_badge_cartera">Multi-Curso</span>
      </button>
      <button type="button" class="mod-btn" onclick="cambiarModulo('mod_clases', this)">
        <span>🎓 2. Seguimiento Clases y Alumnos</span>
        <span class="badge badge-info" id="nav_badge_clases">F-59 a F-62</span>
      </button>
      <button type="button" class="mod-btn" onclick="cambiarModulo('mod_nomina', this)">
        <span>👨‍🏫 3. Nómina y Horarios Instructores</span>
        <span class="badge badge-purple" id="nav_badge_nomina">Liquidación</span>
      </button>
      <button type="button" class="mod-btn" onclick="cambiarModulo('mod_inasistencias', this)">
        <span>🚫 4. Control de Inasistencias</span>
        <span class="badge badge-warning" id="nav_badge_inas">153+</span>
      </button>
      <button type="button" class="mod-btn" onclick="cambiarModulo('mod_pesv', this)">
        <span>🔧 5. Mantenimientos y Flota PESV</span>
        <span class="badge badge-success" id="nav_badge_pesv">SISI-F2-08</span>
      </button>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 1: AUDITORÍA CONTABLE Y CONTROL DE CARTERA                       -->
    <!-- ======================================================================= -->
    <div id="mod_cartera" class="mod-section active">
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
    </div> <!-- Fin mod_cartera -->

    <!-- ======================================================================= -->
    <!-- MÓDULO 2: SEGUIMIENTO DE CLASES PRÁCTICAS Y ALUMNOS (F-59 a F-62)       -->
    <!-- ======================================================================= -->
    <div id="mod_clases" class="mod-section">
      <div class="kpi-grid" id="kpiClasesGrid"></div>

      <div class="control-bar">
        <div class="panel">
          <div class="panel-title">
            <span>➕ Registrar o Actualizar Alumno en Google Sheets (Sincronizado con el Bot)</span>
            <span class="badge badge-info">Hoja: Alumnos</span>
          </div>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Cédula / Documento *</label>
              <input type="text" id="al_cedula" placeholder="Ej: 1094950123" />
            </div>
            <div>
              <label class="field-label">Nombre Completo del Alumno *</label>
              <input type="text" id="al_nombre" placeholder="Ej: JUAN PEREZ GOMEZ" />
            </div>
            <div>
              <label class="field-label">Categoría *</label>
              <select id="al_cat">
                <option value="A2">🏍️ A2 Moto (15 clases - F-60)</option>
                <option value="B1">🚗 B1 Carro Particular (20 clases - F-61)</option>
                <option value="C1">🚕 C1 Servicio Público (30 clases - F-62)</option>
                <option value="RECAT_C1">🔄 Recategorización C1 (10 clases - F-62)</option>
                <option value="A1">🏍️ A1 Moto (8 clases - F-59)</option>
              </select>
            </div>
            <div>
              <label class="field-label">Teléfono / Celular</label>
              <input type="text" id="al_tel" placeholder="Ej: 3101234567" />
            </div>
            <div>
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="registrarAlumnoWeb()">💾 Guardar Alumno</button>
            </div>
          </div>
          <div id="alumnoToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>📝 Registrar Clase Práctica Manualmente</span>
            <span class="badge badge-success">Hoja: Registro_Clases</span>
          </div>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Cédula Alumno *</label>
              <input type="text" id="cl_cedula" placeholder="Cédula" />
            </div>
            <div>
              <label class="field-label">Nombre Alumno *</label>
              <input type="text" id="cl_alumno" placeholder="Nombre Alumno" />
            </div>
            <div>
              <label class="field-label">Categoría</label>
              <select id="cl_cat">
                <option value="A2">A2</option>
                <option value="B1">B1</option>
                <option value="C1">C1</option>
                <option value="RECAT_C1">RECAT_C1</option>
                <option value="A1">A1</option>
              </select>
            </div>
            <div>
              <label class="field-label">Instructor *</label>
              <input type="text" id="cl_instructor" placeholder="Nombre Instructor" />
            </div>
            <div>
              <label class="field-label">Clases Dictadas (Ej: 1, 2)</label>
              <input type="text" id="cl_nums" value="1, 2" />
            </div>
            <div>
              <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="registrarClaseWeb()">✅ Registrar Clase</button>
            </div>
          </div>
          <div id="claseToast" class="status-toast"></div>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>🎓 Listado de Alumnos en Clases Prácticas + Cruce en Vivo con Cartera + Descarga PDF Oficial (F-59 a F-62)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="filtroCatClases" onchange="renderizarModuloClases()" style="width:auto;">
              <option value="TODAS">Todas las Categorías</option>
              <option value="A1">A1 (F-59)</option>
              <option value="A2">A2 (F-60)</option>
              <option value="B1">B1 (F-61)</option>
              <option value="C1">C1 (F-62)</option>
              <option value="RECAT_C1">RECAT_C1 (F-62)</option>
            </select>
            <select id="filtroAlertaCartera" onchange="renderizarModuloClases()" style="width:auto;">
              <option value="TODOS">Todos los Alumnos</option>
              <option value="CON_SALDO">⚠️ Alumnos en Clase con Saldo Pendiente</option>
              <option value="CERO_REAL">🚨 Alumnos en Clase con Abono $0 (Bloquear)</option>
              <option value="FINALIZADOS">✅ Curso Práctico Finalizado</option>
            </select>
            <button type="button" class="btn btn-outline" onclick="cargarBotOverview()">🔄 Sincronizar con Bot</button>
          </div>
        </div>

        <div class="table-toolbar">
          <input type="text" id="searchClasesInput" placeholder="🔎 Buscar alumno por nombre, cédula o categoría..." oninput="renderizarModuloClases()" style="max-width: 440px;" />
          <span id="clasesCountInfo" style="font-size:12.5px; color:var(--muted);">Haz clic en 📄 F-59/F-62 para descargar la ficha oficial con fotos y firmas</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Cédula</th>
                <th>Alumno</th>
                <th>Categoría</th>
                <th>Formato Oficial</th>
                <th>Avance Clases Prácticas</th>
                <th>Estado Curso</th>
                <th>Cruce Contable Cartera</th>
                <th>Saldo Real Cédula</th>
                <th>Acción / PDF Oficial</th>
              </tr>
            </thead>
            <tbody id="clasesTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 3: LIQUIDACIÓN / NÓMINA Y HORARIOS DE INSTRUCTORES               -->
    <!-- ======================================================================= -->
    <div id="mod_nomina" class="mod-section">
      <div class="kpi-grid" id="kpiNominaGrid"></div>

      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>👨‍🏫 Filtros de Liquidación de Instructores y Validación de Horario de Envío (Min 18-28 / Min 40-50)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="nom_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="mes">📅 Este Mes</option>
              <option value="dia">📅 Hoy</option>
              <option value="ayer">📅 Ayer</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="semana_ant">📅 Semana Anterior</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="anio">📅 Este Año</option>
              <option value="todo">📚 Todo el Histórico</option>
            </select>
            <select id="nom_instructor" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODOS">👥 Todos los Instructores</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfNomina()">📄 Descargar PDF Oficial de Nómina / Liquidación</button>
          </div>
        </div>

        <div class="table-wrap" style="max-height:290px; margin-bottom:16px;">
          <table>
            <thead>
              <tr>
                <th>Instructor</th>
                <th>Sesiones Registradas</th>
                <th>Total Clases (Horas)</th>
                <th>✅ A Tiempo (En Horario)</th>
                <th>⚠️ Fuera de Horario</th>
                <th>Desglose por Categoría</th>
                <th>Acción PDF Individual</th>
              </tr>
            </thead>
            <tbody id="nominaResumenTbody"></tbody>
          </table>
        </div>

        <div class="panel-title" style="margin-top:10px;">
          <span>📋 Detalle de Clases Dictadas y Validación de Minuto de Envío</span>
          <input type="text" id="searchNominaInput" placeholder="🔎 Filtrar por instructor, alumno, categoría o validación..." oninput="renderizarModuloNomina()" style="max-width: 360px;" />
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Hora Envío</th>
                <th>Instructor</th>
                <th>Cédula Alumno</th>
                <th>Alumno</th>
                <th>Cat.</th>
                <th>Clases Marcadas</th>
                <th>N° Horas</th>
                <th>Validación Horario</th>
                <th>Detalle Minuto</th>
              </tr>
            </thead>
            <tbody id="nominaDetalleTbody"></tbody>
          </table>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>📸 Cierres de Jornada Diarios (Foto de Agenda / Programación Enviada al Bot)</span>
          <span class="badge badge-purple" id="cntCierresBadge">0 cierres</span>
        </div>
        <div class="table-wrap" style="max-height:260px;">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Instructor</th>
                <th>Estado Cierre</th>
                <th>Evidencia / Enlace Drive</th>
              </tr>
            </thead>
            <tbody id="cierresTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 4: CONTROL DE INASISTENCIAS Y RECAUDO DE MULTAS                  -->
    <!-- ======================================================================= -->
    <div id="mod_inasistencias" class="mod-section">
      <div class="kpi-grid" id="kpiInasGrid"></div>

      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>➕ Registrar Nueva Inasistencia de Alumno (Sincronizado con Google Sheets y Bot)</span>
          <span class="badge badge-warning">Hoja: Inasistencias</span>
        </div>
        <div class="form-grid-4">
          <div>
            <label class="field-label">Fecha Programada *</label>
            <input type="text" id="in_fecha" placeholder="DD/MM/YYYY" />
          </div>
          <div>
            <label class="field-label">Nombre Completo del Alumno *</label>
            <input type="text" id="in_alumno" placeholder="Nombre y Apellidos" />
          </div>
          <div>
            <label class="field-label">Categoría *</label>
            <select id="in_cat">
              <option value="B1">🚗 B1 Carro</option>
              <option value="C1">🚕 C1 Público</option>
              <option value="A2">🏍️ A2 Moto</option>
              <option value="A1">🏍️ A1 Moto</option>
            </select>
          </div>
          <div>
            <label class="field-label">Hora de Clase *</label>
            <input type="text" id="in_hora" placeholder="Ej: 6:00 a 8:00 AM" />
          </div>
          <div>
            <label class="field-label">Instructor Asignado *</label>
            <input type="text" id="in_instructor" placeholder="Ej: ALDEMAR / HECTOR" />
          </div>
          <div>
            <label class="field-label">Estado de Pago *</label>
            <select id="in_estado">
              <option value="Pendiente Pago">⏳ Pendiente Pago</option>
              <option value="Cancelo Efectivo">💵 Cancelo Efectivo</option>
              <option value="Cancelo Trasferencia">🏦 Cancelo Trasferencia</option>
            </select>
          </div>
          <div>
            <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="registrarInasistenciaWeb()">💾 Registrar Inasistencia</button>
          </div>
        </div>
        <div id="inasToast" class="status-toast"></div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>🚫 Historial Oficial de Inasistencias y Control de Pagos</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="inas_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="todo">📚 Todo el Histórico</option>
              <option value="mes">📅 Este Mes</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="dia">📅 Hoy</option>
              <option value="ayer">📅 Ayer</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="anio">📅 Este Año</option>
            </select>
            <select id="inas_estado_filtro" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODOS">Todos los Estados</option>
              <option value="PENDIENTE">⏳ Solo Pendiente Pago</option>
              <option value="PAGADO">✅ Solo Pagadas (Efectivo / Transf.)</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfInasistenciasGeneral()">📄 Descargar PDF Consolidado Inasistencias</button>
          </div>
        </div>

        <div class="table-toolbar">
          <input type="text" id="searchInasInput" placeholder="🔎 Buscar inasistencia por alumno, instructor, fecha o estado..." oninput="renderizarModuloInasistencias()" style="max-width: 440px;" />
          <span id="inasCountInfo" style="font-size:12.5px; color:var(--muted);">Haz clic en 💵 Efectivo / 🏦 Transf. para actualizar el pago en vivo</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Fecha</th>
                <th>Alumno</th>
                <th>Cat.</th>
                <th>Horario</th>
                <th>Instructor</th>
                <th>Estado de Pago</th>
                <th>Actualizar Pago (1 Clic)</th>
                <th>PDF Individual</th>
              </tr>
            </thead>
            <tbody id="inasTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 5: MANTENIMIENTOS Y FLOTA VEHICULAR PESV (FORMATO SISI-F2-08)    -->
    <!-- ======================================================================= -->
    <div id="mod_pesv" class="mod-section">
      <div class="kpi-grid" id="kpiPesvGrid"></div>

      <div class="control-bar">
        <div class="panel">
          <div class="panel-title">
            <span>🔧 Registrar Nuevo Mantenimiento Vehicular (Formato Oficial SISI-F2-08)</span>
            <span class="badge badge-success">Hoja: Mantenimientos</span>
          </div>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Placa del Vehículo *</label>
              <select id="mt_placa"></select>
            </div>
            <div>
              <label class="field-label">Tipo Mantenimiento *</label>
              <select id="mt_tipo">
                <option value="PREVENTIVO">🟢 PREVENTIVO</option>
                <option value="CORRECTIVO">🟠 CORRECTIVO</option>
              </select>
            </div>
            <div>
              <label class="field-label">Kilometraje Actual *</label>
              <input type="text" id="mt_km" placeholder="Ej: 161500" />
            </div>
            <div>
              <label class="field-label">Taller / Proveedor *</label>
              <input type="text" id="mt_taller" placeholder="Ej: TALLER / MONTALLANTAS" />
            </div>
            <div>
              <label class="field-label">Mecánico / Técnico *</label>
              <input type="text" id="mt_mecanico" placeholder="Ej: HECTOR / JAIRO" />
            </div>
            <div style="grid-column: span 2;">
              <label class="field-label">Actividad Realizada / Repuestos *</label>
              <input type="text" id="mt_actividad" placeholder="Ej: CAMBIO DE ACEITE Y FILTROS / REVISION FRENOS" />
            </div>
            <div>
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="registrarMantenimientoWeb()">💾 Guardar Mantenimiento</button>
            </div>
          </div>
          <div id="mtToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>🚘 Gestión de Flota Activa PESV</span>
            <span class="badge badge-info">Hoja: Vehiculos_PESV</span>
          </div>
          <div class="form-grid-4" style="margin-bottom:10px;">
            <div>
              <label class="field-label">Placa *</label>
              <input type="text" id="vh_placa" placeholder="Ej: ABC123" />
            </div>
            <div>
              <label class="field-label">Tipo (CARRO / MOTO)</label>
              <select id="vh_tipo">
                <option value="CARRO">🚗 CARRO</option>
                <option value="MOTO">🏍️ MOTO</option>
              </select>
            </div>
            <div>
              <label class="field-label">Cat. (B1/C1/A2)</label>
              <input type="text" id="vh_cat" value="B1" />
            </div>
            <div>
              <label class="field-label">Marca / Línea</label>
              <input type="text" id="vh_marca" placeholder="RENAULT / AKT" />
            </div>
          </div>
          <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="agregarVehiculoWeb()">➕ Agregar / Reactivar Vehículo en Flota</button>
          <div id="vhToast" class="status-toast"></div>
          <div id="flotaChipsWrap" style="margin-top:10px; display:flex; flex-wrap:wrap; gap:6px;"></div>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>📋 Bitácora Oficial de Mantenimientos Preventivos y Correctivos (SISI-F2-08 Firmado)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="pesv_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="todo">📚 Todo el Histórico (2026)</option>
              <option value="mes">📅 Este Mes</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="anio">📅 Este Año</option>
              <option value="mes_01">Enero 2026</option>
              <option value="mes_02">Febrero 2026</option>
              <option value="mes_03">Marzo 2026</option>
              <option value="mes_04">Abril 2026</option>
              <option value="mes_05">Mayo 2026</option>
              <option value="mes_06">Junio 2026</option>
              <option value="mes_07">Julio 2026</option>
              <option value="mes_08">Agosto 2026</option>
              <option value="mes_09">Septiembre 2026</option>
              <option value="mes_10">Octubre 2026</option>
              <option value="mes_11">Noviembre 2026</option>
              <option value="mes_12">Diciembre 2026</option>
            </select>
            <select id="pesv_placa_filtro" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODAS">🚘 Todas las Placas</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfPesv()">📄 Descargar Formato Oficial SISI-F2-08 (Con Firmas)</button>
          </div>
        </div>

        <div class="table-toolbar">
          <input type="text" id="searchPesvInput" placeholder="🔎 Filtrar por placa, taller, mecánico, actividad o tipo..." oninput="renderizarModuloPesv()" style="max-width: 440px;" />
          <span id="pesvCountInfo" style="font-size:12.5px; color:var(--muted);">Firmado digitalmente por SEGUNDO MURILLO (Rep. Legal) y MARLY ALFONSO (Líder PESV)</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>ID</th>
                <th>Fecha</th>
                <th>Placa</th>
                <th>Tipo Vehículo</th>
                <th>Tipo Mantenimiento</th>
                <th>Kilometraje</th>
                <th>Actividad Realizada / Intervención</th>
                <th>Taller / Proveedor</th>
                <th>Mecánico</th>
                <th>PDF Placa</th>
              </tr>
            </thead>
            <tbody id="pesvTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div>

  <script>
    let AUDIT_DATA = null;
    let BOT_DATA = null;
    let ACTIVE_TAB = 'jt_cartera';
    let ACTIVE_MODULE = 'mod_cartera';
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

    // =========================================================================
    // NAVEGACIÓN ENTRE LOS 5 MÓDULOS Y SINCRONIZACIÓN CON EL BOT / SHEETS
    // =========================================================================
    function cambiarModulo(modId, btnEl) {
      ACTIVE_MODULE = modId;
      document.querySelectorAll('.mod-section').forEach(s => s.classList.remove('active'));
      document.querySelectorAll('.mod-btn').forEach(b => b.classList.remove('active'));
      const target = document.getElementById(modId);
      if (target) target.classList.add('active');
      if (btnEl) btnEl.classList.add('active');
      if (modId !== 'mod_cartera' && !BOT_DATA) {
        cargarBotOverview();
      }
    }

    function toggleSheetsDrawer() {
      document.getElementById('sheetsDrawer').classList.toggle('open');
    }

    function leerArchivoCredenciales(inputEl) {
      if (!inputEl.files || !inputEl.files.length) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        document.getElementById('credJsonInput').value = ev.target.result || '';
      };
      reader.readAsText(inputEl.files[0]);
    }

    async function conectarGoogleSheets() {
      const rawJson = document.getElementById('credJsonInput').value.trim();
      const toast = document.getElementById('sheetsToast');
      if (!rawJson) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Selecciona el archivo credentials.json o pega su contenido JSON.';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Conectando con Google Sheets del Bot...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'conectar_sheets', credentials_json: rawJson })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
        if (data.ok) {
          await cargarBotOverview();
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    async function cargarBotOverview() {
      const nomPeriodo = (document.getElementById('nom_periodo') || {}).value || 'mes';
      const nomInst = (document.getElementById('nom_instructor') || {}).value || 'TODOS';
      const inasPeriodo = (document.getElementById('inas_periodo') || {}).value || 'todo';
      const inasEstado = (document.getElementById('inas_estado_filtro') || {}).value || 'TODOS';
      const pesvPeriodo = (document.getElementById('pesv_periodo') || {}).value || 'todo';
      const pesvPlaca = (document.getElementById('pesv_placa_filtro') || {}).value || 'TODAS';

      const qs = new URLSearchParams({
        periodo_nomina: nomPeriodo,
        instructor_nomina: nomInst,
        periodo_inasistencias: inasPeriodo,
        estado_inasistencias: inasEstado,
        periodo_mantenimientos: pesvPeriodo,
        placa_mantenimientos: pesvPlaca,
      });

      try {
        const res = await fetch('/api/bot/overview?' + qs.toString());
        const data = await res.json();
        if (!data.ok) return;
        BOT_DATA = data;

        // Estado de Google Sheets en Header
        const btnSheets = document.getElementById('btnSheetsStatus');
        const lblSheets = document.getElementById('sheetsStatusLabel');
        if (data.sheets_conectado) {
          btnSheets.innerHTML = '🟢 Sheets Conectado';
          lblSheets.className = 'badge badge-success';
          lblSheets.textContent = '🟢 Sincronizado en Vivo con Google Sheets';
        } else {
          btnSheets.innerHTML = '🟡 Conectar Google Sheets';
          lblSheets.className = 'badge badge-warning';
          lblSheets.textContent = '🟡 Modo Base Local / Semilla Histórica (Clic para conectar Google Sheets en vivo)';
        }

        // Actualizar badges de la barra de módulos
        document.getElementById('nav_badge_clases').textContent = `${data.clases.kpis.total_alumnos} Alumnos`;
        document.getElementById('nav_badge_nomina').textContent = `${data.nomina.total_clases} Clases`;
        document.getElementById('nav_badge_inas').textContent = `${data.inasistencias.kpis.total} (${data.inasistencias.kpis.pendientes} Pend.)`;
        document.getElementById('nav_badge_pesv').textContent = `${data.mantenimientos.kpis.total} Mant.`;

        renderizarModuloClases();
        renderizarModuloNomina();
        renderizarModuloInasistencias();
        renderizarModuloPesv();
      } catch (err) {
        console.error('Error cargando panorama del bot:', err);
      }
    }

    function renderizarModuloClases() {
      if (!BOT_DATA || !BOT_DATA.clases) return;
      const kc = BOT_DATA.clases.kpis;
      document.getElementById('kpiClasesGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Total Alumnos en Seguimiento</div>
          <div class="kpi-value">${kc.total_alumnos}</div>
          <div class="kpi-sub"><b>${kc.en_curso} en curso</b> • ${kc.finalizados} finalizados</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">Sesiones Prácticas Registradas</div>
          <div class="kpi-value" style="color:#34d399;">${kc.total_sesiones_registradas}</div>
          <div class="kpi-sub">Formatos F-59, F-60, F-61 y F-62</div>
        </div>
        <div class="kpi-card warning">
          <div class="kpi-label">⚠️ En Clase con Saldo en Cartera</div>
          <div class="kpi-value" style="color:#fde68a;">${kc.alumnos_con_mora_en_clase} Alumnos</div>
          <div class="kpi-sub">Tienen clases prácticas y saldo pendiente</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">🚨 En Clase con Abono $0 Global</div>
          <div class="kpi-value" style="color:#fca5a5;">${kc.alumnos_abono_cero_en_clase} Alumnos</div>
          <div class="kpi-sub">Bloquear programación práctica inmediata</div>
        </div>
      `;

      const q = (document.getElementById('searchClasesInput').value || '').trim().toLowerCase();
      const catF = document.getElementById('filtroCatClases').value;
      const alertaF = document.getElementById('filtroAlertaCartera').value;

      const rows = (BOT_DATA.clases.alumnos || []).filter(a => {
        if (catF !== 'TODAS' && a.categoria !== catF) return false;
        if (alertaF === 'CON_SALDO' && (a.saldo_cartera || 0) <= 0) return false;
        if (alertaF === 'CERO_REAL' && a.alerta_cartera !== 'RIESGO_CERO_REAL') return false;
        if (alertaF === 'FINALIZADOS' && a.estado_curso !== 'FINALIZADO') return false;
        if (!q) return true;
        return `${a.cedula} ${a.nombre} ${a.categoria} ${a.formato_codigo}`.toLowerCase().includes(q);
      });

      document.getElementById('clasesCountInfo').textContent = `Mostrando ${rows.length} alumnos en seguimiento práctico`;

      document.getElementById('clasesTbody').innerHTML = rows.map(a => {
        let badgeCartera = '<span class="badge badge-success">✅ Al Día / Sin Deuda</span>';
        if (a.alerta_cartera === 'RIESGO_CERO_REAL') {
          badgeCartera = `<span class="badge badge-danger">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'MORA_CRITICA_90') {
          badgeCartera = `<span class="badge badge-danger">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'CON_SALDO') {
          badgeCartera = `<span class="badge badge-warning">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'NO_CRUZADO') {
          badgeCartera = `<span class="badge">Sin cruce en Excel actual</span>`;
        }

        const qParam = encodeURIComponent(a.cedula || a.nombre);
        const catParam = encodeURIComponent(a.categoria || 'A2');

        return `<tr>
          <td><code>${a.cedula || 'S/D'}</code></td>
          <td><b>${a.nombre}</b></td>
          <td><span class="badge badge-info">${a.categoria}</span></td>
          <td><b>${a.formato_codigo}</b></td>
          <td>
            <div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:${a.porcentaje_avance}%;"></div></div>
            <b>${a.clases_completadas}/${a.total_requerido}</b> (${a.porcentaje_avance}%)
          </td>
          <td><span class="badge ${a.estado_curso === 'FINALIZADO' ? 'badge-success' : 'badge-purple'}">${a.estado_curso}</span></td>
          <td>${badgeCartera}</td>
          <td class="${(a.saldo_cartera || 0) > 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(a.saldo_cartera || 0)}</td>
          <td>
            <a href="/api/bot/pdf/alumno?q=${qParam}&cat=${catParam}" class="btn btn-primary" style="padding:5px 10px; font-size:11.5px;">📄 PDF ${a.formato_codigo}</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="9" style="text-align:center; color:var(--muted); padding:20px;">No hay alumnos que coincidan con el filtro seleccionado.</td></tr>';
    }

    function renderizarModuloNomina() {
      if (!BOT_DATA || !BOT_DATA.nomina) return;
      const nm = BOT_DATA.nomina;

      // Poblar selector de instructores conservando selección
      const selInst = document.getElementById('nom_instructor');
      const currentVal = selInst.value || 'TODOS';
      const listaInst = nm.instructores_disponibles || [];
      selInst.innerHTML = '<option value="TODOS">👥 Todos los Instructores</option>' +
        listaInst.map(i => `<option value="${i}" ${i === currentVal ? 'selected' : ''}>👨‍🏫 ${i}</option>`).join('');

      document.getElementById('kpiNominaGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Periodo Consultado</div>
          <div class="kpi-value" style="font-size:18px;">${nm.titulo_periodo}</div>
          <div class="kpi-sub">Filtro: <b>${nm.instructor_filtro}</b></div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">Total Clases Prácticas (Horas)</div>
          <div class="kpi-value" style="color:#c4b5fd;">${nm.total_clases} Clases</div>
          <div class="kpi-sub"><b>${nm.total_sesiones} sesiones</b> registradas</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">✅ Envíos A Tiempo (Min 18-28 / 40-50)</div>
          <div class="kpi-value" style="color:#34d399;">${nm.a_tiempo}</div>
          <div class="kpi-sub">Cumplen ventana oficial del CEA</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">⚠️ Envíos Fuera de Horario</div>
          <div class="kpi-value" style="color:#fca5a5;">${nm.fuera_horario}</div>
          <div class="kpi-sub">Antes o después del minuto permitido</div>
        </div>
      `;

      const perCode = document.getElementById('nom_periodo').value || 'mes';
      document.getElementById('nominaResumenTbody').innerHTML = (nm.resumen_instructores || []).map(r => {
        const catsTxt = Object.entries(r.categorias || {}).map(([k, v]) => `${k}: <b>${v}</b>`).join(' • ');
        const instEnc = encodeURIComponent(r.instructor);
        return `<tr>
          <td><b>👨‍🏫 ${r.instructor}</b></td>
          <td>${r.sesiones} sesiones</td>
          <td><span class="badge badge-purple" style="font-size:12.5px;">${r.clases_dictadas} clases</span></td>
          <td><span class="badge badge-success">${r.a_tiempo} a tiempo</span></td>
          <td><span class="badge ${r.fuera_horario > 0 ? 'badge-danger' : ''}">${r.fuera_horario} fuera de hora</span></td>
          <td>${catsTxt || 'S/D'}</td>
          <td>
            <a href="/api/bot/pdf/nomina?instructor=${instEnc}&periodo=${perCode}" class="btn btn-outline" style="padding:5px 10px; font-size:11.5px;">📄 PDF Liquidación</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="7" style="text-align:center; color:var(--muted);">Sin registros en este periodo</td></tr>';

      const q = (document.getElementById('searchNominaInput').value || '').trim().toLowerCase();
      const detRows = (nm.detalle_clases || []).filter(d => {
        if (!q) return true;
        return `${d.fecha} ${d.instructor} ${d.alumno} ${d.cedula} ${d.categoria} ${d.estado_validacion}`.toLowerCase().includes(q);
      });

      document.getElementById('nominaDetalleTbody').innerHTML = detRows.map(d => {
        const okHora = d.estado_validacion.includes('A TIEMPO');
        return `<tr>
          <td>${d.fecha}</td>
          <td><code>${d.hora}</code></td>
          <td><b>${d.instructor}</b></td>
          <td><code>${d.cedula}</code></td>
          <td>${d.alumno}</td>
          <td><span class="badge">${d.categoria}</span></td>
          <td><b>${d.clases}</b></td>
          <td>${d.num_clases}</td>
          <td><span class="badge ${okHora ? 'badge-success' : 'badge-warning'}">${d.estado_validacion}</span></td>
          <td style="color:var(--muted); font-size:11.5px;">${d.detalle_validacion}</td>
        </tr>`;
      }).join('') || '<tr><td colspan="10" style="text-align:center; color:var(--muted);">No hay sesiones registradas para este filtro.</td></tr>';

      const cierres = BOT_DATA.cierres_jornada || [];
      document.getElementById('cntCierresBadge').textContent = `${cierres.length} cierres`;
      document.getElementById('cierresTbody').innerHTML = cierres.map(c => `<tr>
        <td>${c.fecha}</td>
        <td><code>${c.hora}</code></td>
        <td><b>${c.instructor}</b></td>
        <td><span class="badge badge-success">${c.estado || 'CERRADA'}</span></td>
        <td>${c.drive_link ? `<a href="${c.drive_link}" target="_blank" style="color:#60a5fa;">🔗 Ver Foto Agenda</a>` : 'Registrada en Bot'}</td>
      </tr>`).join('') || '<tr><td colspan="5" style="text-align:center; color:var(--muted);">Sin cierres de jornada recientes.</td></tr>';
    }

    function renderizarModuloInasistencias() {
      if (!BOT_DATA || !BOT_DATA.inasistencias) return;
      const ki = BOT_DATA.inasistencias.kpis;
      document.getElementById('kpiInasGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Total Inasistencias (${BOT_DATA.inasistencias.titulo_periodo})</div>
          <div class="kpi-value">${ki.total}</div>
          <div class="kpi-sub">Registradas en Bot y Google Sheets</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">⏳ Pendientes de Pago</div>
          <div class="kpi-value" style="color:#fca5a5;">${ki.pendientes}</div>
          <div class="kpi-sub">Alumnos por cancelar multa de inasistencia</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">💵 Cancelaron en Efectivo</div>
          <div class="kpi-value" style="color:#34d399;">${ki.efectivo}</div>
          <div class="kpi-sub">Pagadas en caja física</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🏦 Cancelaron por Transferencia</div>
          <div class="kpi-value" style="color:#c4b5fd;">${ki.transferencia}</div>
          <div class="kpi-sub">Pagadas por banco / Nequi / Daviplata</div>
        </div>
      `;

      const q = (document.getElementById('searchInasInput').value || '').trim().toLowerCase();
      const rows = (BOT_DATA.inasistencias.registros || []).filter(r => {
        if (!q) return true;
        return `${r.fecha} ${r.alumno} ${r.categoria} ${r.instructor} ${r.estado_pago}`.toLowerCase().includes(q);
      });

      document.getElementById('inasCountInfo').textContent = `Mostrando ${rows.length} inasistencias`;

      document.getElementById('inasTbody').innerHTML = rows.map(r => {
        const estUp = (r.estado_pago || '').toUpperCase();
        let badgeEst = `<span class="badge badge-danger">⏳ ${r.estado_pago}</span>`;
        if (estUp.includes('EFECTIVO')) badgeEst = `<span class="badge badge-success">💵 ${r.estado_pago}</span>`;
        else if (estUp.includes('TRASFERENCIA') || estUp.includes('TRANSFERENCIA')) badgeEst = `<span class="badge badge-info">🏦 ${r.estado_pago}</span>`;

        const alEnc = encodeURIComponent(r.alumno);
        return `<tr>
          <td><code>#${r.fila}</code></td>
          <td><b>${r.fecha}</b></td>
          <td><b>${r.alumno}</b></td>
          <td><span class="badge">${r.categoria}</span></td>
          <td>${r.hora}</td>
          <td>${r.instructor}</td>
          <td>${badgeEst}</td>
          <td>
            <div style="display:flex; gap:4px;">
              <button type="button" class="btn btn-success" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Cancelo Efectivo')">💵 Efectivo</button>
              <button type="button" class="btn btn-primary" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Cancelo Trasferencia')">🏦 Transf.</button>
              <button type="button" class="btn btn-outline" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Pendiente Pago')">⏳ Pend.</button>
            </div>
          </td>
          <td>
            <a href="/api/bot/pdf/inasistencias?alumno=${alEnc}" class="btn btn-outline" style="padding:4px 9px; font-size:11.5px;">📄 PDF Alumno</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="9" style="text-align:center; color:var(--muted);">No hay inasistencias para este filtro.</td></tr>';
    }

    function renderizarModuloPesv() {
      if (!BOT_DATA || !BOT_DATA.mantenimientos) return;
      const km = BOT_DATA.mantenimientos.kpis;
      const flota = BOT_DATA.mantenimientos.flota_activa || [];

      document.getElementById('kpiPesvGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Total Intervenciones (${BOT_DATA.mantenimientos.titulo_periodo})</div>
          <div class="kpi-value">${km.total}</div>
          <div class="kpi-sub">Formato Oficial <b>SISI-F2-08</b></div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">🟢 Mantenimientos Preventivos</div>
          <div class="kpi-value" style="color:#34d399;">${km.preventivos}</div>
          <div class="kpi-sub">Cambio aceite, filtros, alineación, revisión</div>
        </div>
        <div class="kpi-card orange">
          <div class="kpi-label">🟠 Mantenimientos Correctivos</div>
          <div class="kpi-value" style="color:#fdba74;">${km.correctivos}</div>
          <div class="kpi-sub">Reparación de frenos, embrague, llantas, etc.</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🚘 Vehículos Activos en Flota PESV</div>
          <div class="kpi-value" style="color:#c4b5fd;">${flota.length} Vehículos</div>
          <div class="kpi-sub">${km.vehiculos_intervenidos} intervenidos en el periodo</div>
        </div>
      `;

      // Poblar selectores de placas
      const selMtPlaca = document.getElementById('mt_placa');
      selMtPlaca.innerHTML = flota.map(v => `<option value="${v.placa}">${v.icono} ${v.placa} (${v.categoria} - ${v.marca})</option>`).join('');

      const selFiltroPlaca = document.getElementById('pesv_placa_filtro');
      const curPlaca = selFiltroPlaca.value || 'TODAS';
      selFiltroPlaca.innerHTML = '<option value="TODAS">🚘 Todas las Placas</option>' +
        flota.map(v => `<option value="${v.placa}" ${v.placa === curPlaca ? 'selected' : ''}>${v.icono} ${v.placa}</option>`).join('');

      // Chips de flota activa con opción de retiro
      document.getElementById('flotaChipsWrap').innerHTML = flota.map(v => `
        <span class="badge badge-info" style="padding:5px 10px; font-size:12px; display:inline-flex; align-items:center; gap:6px;">
          ${v.icono} <b>${v.placa}</b> (${v.categoria})
          <button type="button" onclick="retirarVehiculoWeb('${v.placa}')" title="Retirar vehículo de flota activa" style="background:none; border:none; color:#fca5a5; cursor:pointer; font-weight:800;">×</button>
        </span>
      `).join('');

      const q = (document.getElementById('searchPesvInput').value || '').trim().toLowerCase();
      const rows = (BOT_DATA.mantenimientos.registros || []).filter(r => {
        if (!q) return true;
        return `${r.id} ${r.fecha} ${r.placa} ${r.tipo_mantenimiento} ${r.actividad} ${r.taller} ${r.mecanico}`.toLowerCase().includes(q);
      });

      document.getElementById('pesvCountInfo').textContent = `Mostrando ${rows.length} mantenimientos registrados`;
      const perCode = document.getElementById('pesv_periodo').value || 'todo';

      document.getElementById('pesvTbody').innerHTML = rows.map(r => {
        const isPrev = (r.tipo_mantenimiento || '').toUpperCase().includes('PREVENTIVO');
        return `<tr>
          <td><code>${r.id}</code></td>
          <td><b>${r.fecha}</b></td>
          <td><span class="badge badge-info"><b>${r.placa}</b></span></td>
          <td>${r.tipo_vehiculo}</td>
          <td><span class="badge ${isPrev ? 'badge-success' : 'badge-warning'}">${r.tipo_mantenimiento}</span></td>
          <td><code>${r.kilometraje} km</code></td>
          <td style="white-space:normal; max-width:340px;"><b>${r.actividad}</b></td>
          <td>${r.taller}</td>
          <td>${r.mecanico}</td>
          <td>
            <a href="/api/bot/pdf/mantenimientos?periodo=${perCode}&placa=${encodeURIComponent(r.placa)}" class="btn btn-outline" style="padding:4px 9px; font-size:11.5px;">📄 SISI-F2-08</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="10" style="text-align:center; color:var(--muted);">No hay mantenimientos para este filtro.</td></tr>';
    }

    async function registrarAlumnoWeb() {
      const cedula = document.getElementById('al_cedula').value.trim();
      const nombre = document.getElementById('al_nombre').value.trim();
      const categoria = document.getElementById('al_cat').value;
      const telefono = document.getElementById('al_tel').value.trim();
      const toast = document.getElementById('alumnoToast');
      if (!cedula || !nombre) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa Cédula y Nombre del alumno.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'registrar_alumno', cedula, nombre, categoria, telefono })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function registrarClaseWeb() {
      const cedula = document.getElementById('cl_cedula').value.trim();
      const alumno = document.getElementById('cl_alumno').value.trim();
      const categoria = document.getElementById('cl_cat').value;
      const instructor = document.getElementById('cl_instructor').value.trim();
      const clases_str = document.getElementById('cl_nums').value.trim();
      const toast = document.getElementById('claseToast');
      if (!cedula || !alumno || !instructor) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Completa Cédula, Nombre e Instructor.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'registrar_clase', cedula, alumno, categoria, instructor, clases_str })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function registrarInasistenciaWeb() {
      const fecha_clase = document.getElementById('in_fecha').value.trim();
      const alumno = document.getElementById('in_alumno').value.trim();
      const categoria = document.getElementById('in_cat').value;
      const hora_programada = document.getElementById('in_hora').value.trim();
      const instructor = document.getElementById('in_instructor').value.trim();
      const estado_pago = document.getElementById('in_estado').value;
      const toast = document.getElementById('inasToast');
      if (!alumno || !instructor) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa el nombre del alumno y el instructor.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'registrar_inasistencia',
          fecha_clase, alumno, categoria, hora_programada, instructor, estado_pago
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function cambiarEstadoInasistencia(filaSheet, nuevoEstado) {
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'estado_inasistencia', fila: filaSheet, estado_pago: nuevoEstado })
      });
      const data = await res.json();
      if (data.ok) cargarBotOverview();
    }

    async function registrarMantenimientoWeb() {
      const placa = document.getElementById('mt_placa').value;
      const tipo_mantenimiento = document.getElementById('mt_tipo').value;
      const kilometraje = document.getElementById('mt_km').value.trim();
      const taller = document.getElementById('mt_taller').value.trim();
      const mecanico = document.getElementById('mt_mecanico').value.trim();
      const actividad = document.getElementById('mt_actividad').value.trim();
      const toast = document.getElementById('mtToast');
      if (!placa || !actividad) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Selecciona la placa y describe la actividad realizada.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'registrar_mantenimiento',
          placa, tipo_mantenimiento, kilometraje, taller, mecanico, actividad
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function agregarVehiculoWeb() {
      const placa = document.getElementById('vh_placa').value.trim().toUpperCase();
      const tipo = document.getElementById('vh_tipo').value;
      const categoria = document.getElementById('vh_cat').value.trim().toUpperCase();
      const marca = document.getElementById('vh_marca').value.trim().toUpperCase();
      const toast = document.getElementById('vhToast');
      if (!placa) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa la placa del vehículo.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'vehiculo_agregar', placa, tipo, categoria, marca })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function retirarVehiculoWeb(placa) {
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'vehiculo_retirar', placa })
      });
      const data = await res.json();
      if (data.ok) cargarBotOverview();
    }

    function descargarPdfNomina() {
      const per = document.getElementById('nom_periodo').value || 'mes';
      const inst = document.getElementById('nom_instructor').value || 'TODOS';
      window.location.href = `/api/bot/pdf/nomina?instructor=${encodeURIComponent(inst)}&periodo=${encodeURIComponent(per)}`;
    }

    function descargarPdfInasistenciasGeneral() {
      const per = document.getElementById('inas_periodo').value || 'todo';
      const est = document.getElementById('inas_estado_filtro').value || 'TODOS';
      window.location.href = `/api/bot/pdf/inasistencias?periodo=${encodeURIComponent(per)}&estado=${encodeURIComponent(est)}`;
    }

    function descargarPdfPesv() {
      const per = document.getElementById('pesv_periodo').value || 'todo';
      const placa = document.getElementById('pesv_placa_filtro').value || 'TODAS';
      window.location.href = `/api/bot/pdf/mantenimientos?periodo=${encodeURIComponent(per)}&placa=${encodeURIComponent(placa)}`;
    }

    // Inicializar fecha por defecto de inasistencia hoy
    (function initDefaults() {
      const hoy = new Date();
      const dd = String(hoy.getDate()).padStart(2, '0');
      const mm = String(hoy.getMonth() + 1).padStart(2, '0');
      const yyyy = hoy.getFullYear();
      const inFecha = document.getElementById('in_fecha');
      if (inFecha) inFecha.value = `${dd}/${mm}/${yyyy}`;
    })();

    cargarDatos();
    cargarBotOverview();
  </script>
</body>
</html>
"""


class AuditorWebHandler(BaseHTTPRequestHandler):
    def _send_json(self, data: Dict[str, Any], status: int = 200, extra_headers: Optional[Dict[str, str]] = None) -> None:
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(raw)

    def _is_authenticated(self) -> bool:
        cookie_hdr = self.headers.get("Cookie", "")
        if not cookie_hdr:
            return False
        c = SimpleCookie()
        try:
            c.load(cookie_hdr)
        except Exception:
            return False
        if COOKIE_NAME not in c:
            return False
        return _verificar_token_sesion(c[COOKIE_NAME].value)

    def do_GET(self) -> None:
        if self.path == "/health":
            self._send_json({"ok": True, "service": "Auditor CEA Auto Oriente Multi-Curso"})
            return

        if self.path == "/logout":
            self.send_response(302)
            self.send_header("Location", "/login")
            self.send_header("Set-Cookie", f"{COOKIE_NAME}=; Path=/; HttpOnly; Max-Age=0; SameSite=Lax")
            self.end_headers()
            return

        if self.path == "/login":
            raw = HTML_LOGIN.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if not self._is_authenticated():
            if self.path.startswith("/api/"):
                self._send_json({"ok": False, "error": "Sesión expirada o no autorizada."}, 401)
            else:
                raw = HTML_LOGIN.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
            return

        if self.path in ("/", "/index.html"):
            raw = HTML_DASHBOARD.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
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

        if self.path.startswith("/api/bot/overview"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            p_nomina = (qs.get("periodo_nomina", ["mes"])[0] or "mes").strip()
            inst_nomina = (qs.get("instructor_nomina", ["TODOS"])[0] or "TODOS").strip()
            p_inas = (qs.get("periodo_inasistencias", ["todo"])[0] or "todo").strip()
            est_inas = (qs.get("estado_inasistencias", ["TODOS"])[0] or "TODOS").strip()
            p_mant = (qs.get("periodo_mantenimientos", ["todo"])[0] or "todo").strip()
            placa_mant = (qs.get("placa_mantenimientos", ["TODAS"])[0] or "TODAS").strip()

            panorama = clases_core.sheets_service.obtener_panorama_web(
                periodo_nomina=p_nomina,
                instructor_nomina=inst_nomina,
                periodo_inasistencias=p_inas,
                estado_inasistencias=est_inas,
                periodo_mantenimientos=p_mant,
                placa_mantenimientos=placa_mant,
                hallazgos_cartera=STATE.get("hallazgos"),
            )
            self._send_json(panorama)
            return

        if self.path.startswith("/api/bot/pdf/alumno"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            q_val = (qs.get("q", [""])[0] or "").strip()
            cat_val = (qs.get("cat", ["A2"])[0] or "A2").strip().upper()
            res_al = clases_core.sheets_service.obtener_historial_alumno(q_val)
            alumno = res_al.get("alumno") or {"cedula": q_val, "nombre": q_val or "ALUMNO CEA", "categoria": cat_val, "telefono": ""}
            if cat_val and cat_val in clases_core.FORMATOS_OFICIALES:
                alumno["categoria"] = cat_val
            registros = [
                r for r in (res_al.get("registros") or [])
                if not cat_val or str(r.get("categoria", "")).strip().upper() == cat_val
            ] or (res_al.get("registros") or [])
            pdf_buf = clases_core.generar_pdf_seguimiento(alumno, registros)
            pdf_bytes = pdf_buf.getvalue()
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(alumno.get("nombre", "Alumno")))
            fname = f"Seguimiento_Clases_{alumno.get('categoria', 'CEA')}_{safe_name}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/nomina"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            inst = (qs.get("instructor", ["TODOS"])[0] or "TODOS").strip()
            periodo = (qs.get("periodo", ["mes"])[0] or "mes").strip()
            reporte = clases_core.sheets_service.obtener_clases_por_instructor_periodo(inst, periodo)
            pdf_buf = clases_core.generar_pdf_liquidacion_instructor(reporte)
            pdf_bytes = pdf_buf.getvalue()
            safe_inst = re.sub(r"[^A-Za-z0-9_-]+", "_", inst)
            fname = f"Nomina_Liquidacion_{safe_inst}_{periodo}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/inasistencias"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            alumno_q = (qs.get("alumno", [""])[0] or "").strip()
            if alumno_q:
                datos_al = clases_core.sheets_service.buscar_inasistencias_alumno(alumno_q)
                pdf_buf = clases_core.generar_pdf_inasistencias_individual(datos_al)
                safe_al = re.sub(r"[^A-Za-z0-9_-]+", "_", alumno_q)
                fname = f"Inasistencias_Alumno_{safe_al}.pdf"
            else:
                periodo = (qs.get("periodo", ["todo"])[0] or "todo").strip()
                estado = (qs.get("estado", ["TODOS"])[0] or "TODOS").strip()
                datos_per = clases_core.sheets_service.obtener_inasistencias_periodo(periodo, estado)
                pdf_buf = clases_core.generar_pdf_inasistencias_periodo(datos_per)
                fname = f"Reporte_Inasistencias_{periodo}_{estado}.pdf"
            pdf_bytes = pdf_buf.getvalue()
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/mantenimientos"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            periodo = (qs.get("periodo", ["todo"])[0] or "todo").strip()
            placa = (qs.get("placa", ["TODAS"])[0] or "TODAS").strip()
            datos_mt = clases_core.sheets_service.obtener_mantenimientos_periodo(periodo, placa)
            pdf_buf = clases_core.generar_pdf_mantenimientos_pesv(datos_mt)
            pdf_bytes = pdf_buf.getvalue()
            fname = f"PESV_SISI_F2_08_{placa}_{periodo}.pdf"
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
        if self.path == "/api/login":
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}
            user_in = str(payload.get("username", "")).strip()
            pass_in = str(payload.get("password", ""))
            if hmac.compare_digest(user_in, AUDITOR_USER) and hmac.compare_digest(pass_in, AUDITOR_PASSWORD):
                token = _generar_token_sesion(AUDITOR_USER)
                cookie_val = f"{COOKIE_NAME}={token}; Path=/; HttpOnly; Max-Age=604800; SameSite=Lax"
                self._send_json({"ok": True}, 200, extra_headers={"Set-Cookie": cookie_val})
            else:
                time.sleep(0.6)
                self._send_json({"ok": False, "error": "Usuario o contraseña incorrectos."}, 401)
            return

        if not self._is_authenticated():
            self._send_json({"ok": False, "error": "No autorizado. Inicie sesión primero."}, 401)
            return

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

        if self.path == "/api/bot/action":
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}

            action = str(payload.get("action", "")).strip()

            if action == "conectar_sheets":
                raw_json = str(payload.get("credentials_json", "")).strip()
                res = clases_core.sheets_service.configurar_credenciales_en_vivo(raw_json)
                self._send_json(res)
                return

            if action == "registrar_alumno":
                cedula = str(payload.get("cedula", "")).strip()
                nombre = str(payload.get("nombre", "")).strip().upper()
                categoria = str(payload.get("categoria", "A2")).strip().upper()
                telefono = str(payload.get("telefono", "")).strip()
                ok = clases_core.sheets_service.asegurar_alumno(cedula, nombre, categoria, telefono)
                self._send_json({
                    "ok": bool(ok),
                    "message": f"Alumno {nombre} ({categoria} - C.C. {cedula}) guardado y sincronizado.",
                })
                return

            if action == "registrar_clase":
                cedula = str(payload.get("cedula", "")).strip()
                alumno = str(payload.get("alumno", "")).strip().upper()
                categoria = str(payload.get("categoria", "A2")).strip().upper()
                instructor = str(payload.get("instructor", "")).strip().upper()
                clases_str = str(payload.get("clases_str", "1, 2")).strip()
                nums = [int(x) for x in re.findall(r"\d+", clases_str)] or [1]
                clases_core.sheets_service.asegurar_alumno(cedula, alumno, categoria, "")
                ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
                ts_str = ahora_col.strftime("%Y-%m-%d %H:%M:%S")
                f_str = ahora_col.strftime("%Y-%m-%d")
                h_str = ahora_col.strftime("%H:%M:%S")
                clases_fmt = clases_core.formatear_clases(nums)
                tema_auto = clases_core.obtener_temas_para_clases(categoria, nums)
                fila_idx = clases_core.sheets_service.registrar_clase(
                    timestamp=ts_str,
                    fecha_reporte=f_str,
                    hora_reporte=h_str,
                    instructor=instructor,
                    id_alumno=cedula,
                    nombre_alumno=alumno,
                    categoria=categoria,
                    clases_reportadas=clases_fmt,
                    tema_registrado=tema_auto,
                    url_foto_drive="Registrado en Web",
                    validacion_horario="EN TIEMPO (WEB)",
                    observacion_instructor=f"Registrado por {AUDITOR_USER} desde Dashboard Web",
                )
                self._send_json({
                    "ok": bool(fila_idx),
                    "message": f"Clase(s) {clases_fmt} registrada(s) para {alumno} ({categoria}) con {instructor}.",
                })
                return

            if action == "registrar_inasistencia":
                res_in = clases_core.sheets_service.registrar_inasistencia(
                    documento_alumno=str(payload.get("cedula", "")).strip(),
                    nombre_alumno=str(payload.get("alumno", "")).strip().upper(),
                    categoria=str(payload.get("categoria", "B1")).strip().upper(),
                    clases_faltadas="No especificada",
                    nombre_instructor=str(payload.get("instructor", "")).strip().upper(),
                    no_horas=2,
                    valor_hora=25000,
                    estado=str(payload.get("estado_pago", "Pendiente Pago")).strip(),
                    fecha=str(payload.get("fecha_clase", "")).strip(),
                    hora=str(payload.get("hora_programada", "")).strip(),
                )
                self._send_json({
                    "ok": bool(res_in),
                    "message": f"Inasistencia #{res_in.get('Consecutivo', '')} registrada y sincronizada con el Bot.",
                })
                return

            if action == "estado_inasistencia":
                fila = str(payload.get("fila", "")).strip()
                nuevo_estado = str(payload.get("estado_pago", "Cancelo Efectivo")).strip()
                res_up = clases_core.sheets_service.actualizar_estado_inasistencia(fila, nuevo_estado)
                self._send_json({
                    "ok": res_up is not None,
                    "message": f"Inasistencia #{fila} actualizada a {nuevo_estado}.",
                })
                return

            if action == "registrar_mantenimiento":
                res_mt = clases_core.sheets_service.registrar_mantenimiento(
                    placa=str(payload.get("placa", "")).strip().upper(),
                    tipo_mantenimiento=str(payload.get("tipo_mantenimiento", "PREVENTIVO")).strip(),
                    establecimiento=str(payload.get("taller", "TALLER AUTORIZADO")).strip().upper(),
                    nit_cc="0",
                    actividad_realizada=str(payload.get("actividad", "")).strip().upper(),
                    nombre_tecnico=str(payload.get("mecanico", "")).strip().upper(),
                    rol="MECANICO",
                    observaciones=f"Registrado desde Web ({AUDITOR_USER})",
                    kilometraje_actual=str(payload.get("kilometraje", "0")).strip(),
                )
                self._send_json({
                    "ok": bool(res_mt),
                    "message": f"Mantenimiento registrado en SISI-F2-08 para placa {res_mt.get('Placa', '')}.",
                })
                return

            if action == "vehiculo_agregar":
                res_vh = clases_core.sheets_service.agregar_vehiculo_pesv(
                    placa=str(payload.get("placa", "")).strip().upper(),
                    tipo=str(payload.get("tipo", "CARRO")).strip().upper(),
                    categoria=str(payload.get("categoria", "B1")).strip().upper(),
                    marca=str(payload.get("marca", "")).strip().upper(),
                )
                self._send_json({
                    "ok": bool(res_vh.get("ok")),
                    "message": f"Vehículo {res_vh.get('placa', '')} activo en flota PESV.",
                })
                return

            if action == "vehiculo_retirar":
                res_ret = clases_core.sheets_service.retirar_vehiculo_pesv(
                    placa=str(payload.get("placa", "")).strip().upper(),
                )
                self._send_json({
                    "ok": bool(res_ret.get("ok")),
                    "message": f"Vehículo {res_ret.get('placa', '')} retirado de flota activa.",
                })
                return

            self._send_json({"ok": False, "message": f"Acción desconocida: {action}"}, 400)
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

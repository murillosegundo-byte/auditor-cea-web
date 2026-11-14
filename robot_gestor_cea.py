"""
=============================================================================
ROBOT GESTOR DE ARCHIVOS Y VIGILANTE EN SEGUNDO PLANO — CEA AUTO ORIENTE
=============================================================================
Arquitectura modular en segundo plano (Daemon) que conecta 3 módulos:
1. Vigilante de Carpetas en tiempo real (`watchdog` o monitoreo ligero sin bloqueo).
2. Gestor y Clasificador Inteligente (`pathlib` + `shutil`):
   - Lee extensiones (`.pdf`, `.xlsx`, `.xls`, `.csv`, `.jpg`, `.png`).
   - Si detecta un archivo de Matrículas (`matriculas*.xlsx` o `convenios*.xlsx`),
     ejecuta automáticamente la Auditoría Multi-Curso (`auditor_cea.py`) y genera
     el Excel Multi-Pestaña + Dictamen PDF Ejecutivo.
   - Renombra con fecha/hora y mueve a `Archivo_Procesado/<Categoria>/<YYYY-MM>/`.
3. Conectores Multicanal de Mensajería (`requests` / HTTP API):
   - Despacha automáticamente el documento procesado a Telegram (Bot API) y
     WhatsApp (Cloud API / Evolution API / WAHA).

Uso en consola:
    python robot_gestor_cea.py

Uso silencioso 24/7 en Windows (sin ventana negra):
    Doble clic en `iniciar_robot_silencioso.vbs` o con `pythonw robot_gestor_cea.py`
=============================================================================
"""

import base64
import datetime
import json
import logging
import os
import re
import shutil
import time
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Carpetas configurables (usa C:/Automatizacion si es accesible, o subcarpetas del proyecto)
DEFAULT_ROOT = Path(os.getenv("ROBOT_ROOT_DIR", "C:/Automatizacion"))
try:
    DEFAULT_ROOT.mkdir(parents=True, exist_ok=True)
    CARPETA_ENTRADA = DEFAULT_ROOT / "Entrada"
    CARPETA_ORGANIZADA = DEFAULT_ROOT / "Archivo_Procesado"
except Exception:
    CARPETA_ENTRADA = BASE_DIR / "Entrada_Automatizacion"
    CARPETA_ORGANIZADA = BASE_DIR / "Archivo_Procesado"

CARPETA_ENTRADA.mkdir(parents=True, exist_ok=True)
CARPETA_ORGANIZADA.mkdir(parents=True, exist_ok=True)

# También sincronizar con la bitácora del Dashboard Web
ROBOT_CONFIG_PATH = BASE_DIR / "robot_config.json"
ROBOT_LOG_PATH = BASE_DIR / "robot_bitacora.json"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - [ROBOT-CEA] - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(BASE_DIR / "robot_gestor_cea.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("robot_gestor_cea")


def cargar_credenciales_robot() -> Dict[str, str]:
    cfg = {
        "tg_token": os.getenv("TELEGRAM_BOT_TOKEN", "").strip(),
        "tg_chat": (os.getenv("TELEGRAM_CHAT_IDS", "") or os.getenv("ADMIN_CHAT_ID", "")).split(",")[0].strip(),
        "wa_url": os.getenv("WHATSAPP_API_URL", "").strip(),
        "wa_token": os.getenv("WHATSAPP_API_TOKEN", "").strip(),
        "wa_phone": os.getenv("WHATSAPP_DEFAULT_PHONE", "573134908073").strip(),
    }
    if ROBOT_CONFIG_PATH.exists():
        try:
            saved = json.loads(ROBOT_CONFIG_PATH.read_text(encoding="utf-8"))
            for k in cfg:
                if saved.get(k):
                    cfg[k] = str(saved[k]).strip()
        except Exception:
            pass
    return cfg


def registrar_en_bitacora(
    archivo_original: str,
    tipo_clasificado: str,
    ruta_respaldo: str,
    canal: str,
    estado: str,
) -> None:
    ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
    item = {
        "fecha": ahora_col.strftime("%Y-%m-%d %H:%M:%S"),
        "archivo_original": archivo_original,
        "tipo_clasificado": tipo_clasificado,
        "ruta_respaldo": ruta_respaldo,
        "canal": canal,
        "estado": estado,
    }
    bitacora = []
    if ROBOT_LOG_PATH.exists():
        try:
            bitacora = json.loads(ROBOT_LOG_PATH.read_text(encoding="utf-8"))
            if not isinstance(bitacora, list):
                bitacora = []
        except Exception:
            bitacora = []
    bitacora.insert(0, item)
    try:
        ROBOT_LOG_PATH.write_text(json.dumps(bitacora[:100], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.warning("No se pudo actualizar robot_bitacora.json: %s", exc)


# =============================================================================
# 1. CONECTORES DE MENSAJERÍA (TELEGRAM + WHATSAPP API)
# =============================================================================
def enviar_a_telegram(ruta_archivo: Path, mensaje: str) -> Tuple[bool, str]:
    """Envía el archivo procesado al chat de Telegram usando Bot API."""
    cfg = cargar_credenciales_robot()
    token = cfg["tg_token"]
    chat_id = cfg["tg_chat"]
    if not token or not chat_id:
        return False, "Sin TELEGRAM_BOT_TOKEN o TELEGRAM_CHAT_IDS configurado."

    url = f"https://api.telegram.org/bot{token}/sendDocument"
    boundary = "----RobotCEABoundary" + str(int(time.time() * 1000))
    file_bytes = ruta_archivo.read_bytes()

    parts = [
        f"--{boundary}\r\n".encode("utf-8"),
        f'Content-Disposition: form-data; name="chat_id"\r\n\r\n{chat_id}\r\n'.encode("utf-8"),
        f"--{boundary}\r\n".encode("utf-8"),
        f'Content-Disposition: form-data; name="caption"\r\n\r\n{mensaje[:1000]}\r\n'.encode("utf-8"),
        f"--{boundary}\r\n".encode("utf-8"),
        (
            f'Content-Disposition: form-data; name="document"; filename="{ruta_archivo.name}"\r\n'
            f"Content-Type: application/octet-stream\r\n\r\n"
        ).encode("utf-8"),
        file_bytes,
        f"\r\n--{boundary}--\r\n".encode("utf-8"),
    ]
    body = b"".join(parts)
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))
            if data.get("ok"):
                logger.info("✈️ Enviado a Telegram (%s): %s", chat_id, ruta_archivo.name)
                return True, f"Telegram OK ({chat_id})"
            return False, f"Error Telegram: {data.get('description')}"
    except Exception as exc:
        logger.error("Error enviando a Telegram: %s", exc)
        return False, str(exc)


def enviar_a_whatsapp(ruta_archivo: Path, telefono: str, mensaje: str) -> Tuple[bool, str]:
    """
    Envía documentos vía petición HTTP a WhatsApp Cloud API / Evolution API / WAHA
    si `WHATSAPP_API_URL` está configurado.
    """
    cfg = cargar_credenciales_robot()
    wa_url = cfg["wa_url"]
    wa_token = cfg["wa_token"]
    clean_phone = re.sub(r"\D+", "", telefono or cfg["wa_phone"] or "573134908073")
    if not wa_url:
        return False, "WHATSAPP_API_URL no configurada (modo solo Telegram + respaldo)"

    try:
        payload = json.dumps({
            "number": clean_phone,
            "phone": clean_phone,
            "caption": mensaje,
            "fileName": ruta_archivo.name,
            "mediatype": "document",
            "media": base64.b64encode(ruta_archivo.read_bytes()).decode("ascii"),
        }).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        if wa_token:
            headers["Authorization"] = f"Bearer {wa_token}"
            headers["apikey"] = wa_token
        req = urllib.request.Request(wa_url, data=payload, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=25) as resp:
            if 200 <= resp.status < 300:
                logger.info("📲 Enviado a WhatsApp (+%s): %s", clean_phone, ruta_archivo.name)
                return True, f"WhatsApp OK (+{clean_phone})"
    except Exception as exc:
        logger.warning("No se pudo despachar por WhatsApp API: %s", exc)
    return False, "Fallo WhatsApp API"


# =============================================================================
# 2. GESTOR, AUDITOR AUTOMÁTICO Y CLASIFICADOR DE ARCHIVOS
# =============================================================================
def clasificar_y_procesar(ruta_origen: Path) -> None:
    """Clasifica el archivo por extensión, lo renombra, respalda y lo despacha."""
    if not ruta_origen.exists() or not ruta_origen.is_file():
        return

    # Ignorar archivos temporales de Excel o descargas incompletas
    if ruta_origen.name.startswith("~$") or ruta_origen.suffix.lower() in (".tmp", ".crdownload", ".part"):
        return

    extension = ruta_origen.suffix.lower()
    ahora = datetime.datetime.now()
    mes_actual = ahora.strftime("%Y-%m")
    fecha_str = ahora.strftime("%Y%m%d_%H%M%S")

    if extension == ".pdf":
        subcarpeta = "Reportes_PDF_CEA"
    elif extension in (".xlsx", ".xls", ".csv"):
        subcarpeta = "Auditorias_Excel_Cartera"
    elif extension in (".jpg", ".jpeg", ".png", ".webp"):
        subcarpeta = "Evidencias_Clases_y_Recibos"
    else:
        subcarpeta = "Otros_Documentos"

    directorio_destino = CARPETA_ORGANIZADA / subcarpeta / mes_actual
    directorio_destino.mkdir(parents=True, exist_ok=True)

    nombre_limpio = re.sub(r"[^A-Za-z0-9_.-]+", "_", ruta_origen.name)
    nuevo_nombre = f"{fecha_str}_{nombre_limpio}"
    ruta_final = directorio_destino / nuevo_nombre

    shutil.move(str(ruta_origen), str(ruta_final))
    logger.info("📂 Archivo clasificado en: %s", ruta_final)

    # Si es un archivo de matrículas/convenios en Excel, ejecutar también auditoría automática
    resumen_extra = ""
    if extension in (".xlsx", ".xls", ".csv") and any(
        w in nombre_limpio.lower() for w in ("matricula", "convenio", "cartera")
    ):
        try:
            import auditor_cea

            registros, meta = auditor_cea.cargar_y_limpiar_datos(ruta_final)
            hallazgos = auditor_cea.auditar_riesgos(registros, metadata=meta)
            kpis = hallazgos["kpis"]
            pdf_bytes = auditor_cea.generar_reporte_pdf_bytes(hallazgos, tipo_informe="AUTOMATICO")
            pdf_auto_path = directorio_destino / f"{fecha_str}_Dictamen_Ejecutivo_CEA.pdf"
            pdf_auto_path.write_bytes(pdf_bytes)
            resumen_extra = (
                f"\n📊 *Auditoría Automática Ejecutada:*\n"
                f"• Matrículas: {kpis['total_matriculas']} ({kpis['total_clientes_unicos']} clientes)\n"
                f"• Cartera Pendiente: ${kpis['total_saldo']:,.0f}\n"
                f"• Riesgo $0 Real: {kpis['riesgo_cero_count']} matrículas"
            )
            enviar_a_telegram(pdf_auto_path, f"📑 Dictamen PDF generado automáticamente de {ruta_origen.name}{resumen_extra}")
        except Exception as exc:
            logger.warning("Archivo Excel clasificado sin auditoría automática: %s", exc)

    notificacion = (
        f"🤖 *Robot Gestor CEA Auto Oriente*\n"
        f"📄 *Nuevo documento procesado:* `{nuevo_nombre}`\n"
        f"📁 *Clasificación:* `{subcarpeta}/{mes_actual}`"
        f"{resumen_extra}"
    )

    ok_tg, _ = enviar_a_telegram(ruta_final, mensaje=notificacion)
    cfg = cargar_credenciales_robot()
    ok_wa, _ = enviar_a_whatsapp(ruta_final, telefono=cfg["wa_phone"], mensaje=notificacion)

    canales = []
    if ok_tg:
        canales.append("TELEGRAM")
    if ok_wa:
        canales.append("WHATSAPP")
    canal_str = " + ".join(canales) if canales else "ARCHIVADO LOCAL"

    registrar_en_bitacora(
        archivo_original=ruta_origen.name,
        tipo_clasificado=f"{subcarpeta}/{mes_actual}",
        ruta_respaldo=str(ruta_final),
        canal=canal_str,
        estado="✅ Procesado por Vigilante",
    )


# =============================================================================
# 3. VIGILANTE DE CARPETAS EN TIEMPO REAL (WATCHDOG O POLLING LIGERO)
# =============================================================================
def iniciar_servicio() -> None:
    logger.info("👀 Vigilando carpeta de entrada: %s", CARPETA_ENTRADA)
    logger.info("📂 Carpeta de respaldo clasificado: %s", CARPETA_ORGANIZADA)

    # Procesar primero cualquier archivo que ya esté esperando en la carpeta de entrada
    for item in list(CARPETA_ENTRADA.iterdir()):
        if item.is_file():
            time.sleep(0.5)
            clasificar_y_procesar(item)

    try:
        from watchdog.events import FileSystemEventHandler
        from watchdog.observers import Observer

        class VigilanteCarpeta(FileSystemEventHandler):
            def on_created(self, event: Any) -> None:
                if not event.is_directory:
                    time.sleep(1.5)  # Esperar a que termine de copiarse en disco
                    clasificar_y_procesar(Path(event.src_path))

        observador = Observer()
        observador.schedule(VigilanteCarpeta(), str(CARPETA_ENTRADA), recursive=False)
        observador.start()
        logger.info("🟢 Motor Watchdog activo en tiempo real.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            observador.stop()
        observador.join()

    except ImportError:
        logger.info("🟡 Librería `watchdog` no instalada; usando monitor ligero cada 3 segundos.")
        try:
            while True:
                for item in list(CARPETA_ENTRADA.iterdir()):
                    if item.is_file():
                        time.sleep(1.2)
                        clasificar_y_procesar(item)
                time.sleep(3)
        except KeyboardInterrupt:
            logger.info("Servicio detenido por el usuario.")


if __name__ == "__main__":
    iniciar_servicio()

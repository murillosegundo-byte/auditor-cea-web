"""
=============================================================================
BOT DE AGENDA Y RECORDATORIOS GOOGLE CALENDAR + DASHBOARD WEB — CEA AUTO ORIENTE
=============================================================================
Conecta Google Calendar, el Bot de Telegram de Agenda y el Dashboard Web.
"""

import base64
import datetime
import json
import os
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent
TOKEN_CALENDAR_PATH = BASE_DIR / "token_calendar.json"
TOKEN_FALLBACK_PATH = BASE_DIR / "token.json"
CREDS_CALENDAR_PATH = BASE_DIR / "credentials_calendar_oauth.json"
AGENDA_LOCAL_PATH = BASE_DIR / "agenda_eventos_cea.json"

TELEGRAM_AGENDA_TOKEN = (
    os.getenv("TELEGRAM_AGENDA_TOKEN", "").strip()
    or os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
)
SCOPES = ["https://www.googleapis.com/auth/calendar"]
TIMEZONE_STR = "America/Bogota"
LOCAL_TZ = datetime.timezone(datetime.timedelta(hours=-5))

USUARIO_CHAT_ID: Optional[str] = os.getenv("AGENDA_CHAT_ID", "") or os.getenv("ADMIN_CHAT_ID", "")
EVENTOS_NOTIFICADOS = set()

TITULO, FECHA, HORA, DURACION = range(4)


def _cargar_eventos_locales() -> List[Dict[str, Any]]:
    if AGENDA_LOCAL_PATH.exists():
        try:
            data = json.loads(AGENDA_LOCAL_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    # Semilla inicial de ejemplo para el Dashboard si aún no hay eventos locales
    hoy = datetime.datetime.now(LOCAL_TZ)
    semilla = [
        {
            "id": "ev-cea-01",
            "titulo": "Auditoría de Cartera y Revisión Facturación DIAN (FV-2)",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "08:30",
            "duracion_min": 60,
            "tipo": "DIARIO",
            "responsable": "Gerencia CEA Auto Oriente",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
        {
            "id": "ev-cea-02",
            "titulo": "Campaña de Remarketing WhatsApp IA (Clientes Pendientes)",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "10:30",
            "duracion_min": 45,
            "tipo": "DIARIO",
            "responsable": "Comercial / Admisiones",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
        {
            "id": "ev-cea-03",
            "titulo": "Verificación Cierre de Jornada Instructores y Clases Prácticas",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "18:00",
            "duracion_min": 30,
            "tipo": "DIARIO",
            "responsable": "Coordinación Académica",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
        {
            "id": "ev-cea-04",
            "titulo": "Revisión Exámenes Médicos CRC La 27 y RUNT",
            "fecha": (hoy + datetime.timedelta(days=1)).strftime("%Y-%m-%d"),
            "hora": "09:00",
            "duracion_min": 60,
            "tipo": "PUNTUAL",
            "responsable": "Sede Bolarquí",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
    ]
    try:
        AGENDA_LOCAL_PATH.write_text(json.dumps(semilla, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return semilla


def _guardar_eventos_locales(eventos: List[Dict[str, Any]]) -> None:
    try:
        AGENDA_LOCAL_PATH.write_text(json.dumps(eventos, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def get_calendar_service():
    """Autentica con Google Calendar usando token_calendar.json, token.json o GOOGLE_CALENDAR_TOKEN_JSON."""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    env_token = os.getenv("GOOGLE_CALENDAR_TOKEN_JSON", "").strip()
    if env_token and env_token.startswith("{") and not TOKEN_CALENDAR_PATH.exists():
        try:
            TOKEN_CALENDAR_PATH.write_text(env_token, encoding="utf-8")
        except Exception:
            pass

    creds = None
    tk_path = TOKEN_CALENDAR_PATH if TOKEN_CALENDAR_PATH.exists() else TOKEN_FALLBACK_PATH
    if tk_path.exists():
        creds = Credentials.from_authorized_user_file(str(tk_path), SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        tk_path.write_text(creds.to_json(), encoding="utf-8")
    if not creds or not creds.valid:
        raise RuntimeError("Token de Google Calendar no disponible o requiere re-autorización.")
    return build("calendar", "v3", credentials=creds)


def obtener_panorama_agenda_web() -> Dict[str, Any]:
    """
    Obtiene los eventos de Hoy y de los Próximos 7 Días combinando Google Calendar en vivo
    (si hay conexión) con el almacén persistente del Dashboard CEA.
    """
    ahora = datetime.datetime.now(LOCAL_TZ)
    hoy_str = ahora.strftime("%Y-%m-%d")
    limite_7d = (ahora + datetime.timedelta(days=7)).strftime("%Y-%m-%d")

    eventos_locales = _cargar_eventos_locales()
    calendar_conectado = False
    eventos_gcal: List[Dict[str, Any]] = []

    try:
        service = get_calendar_service()
        inicio_dia = ahora.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
        fin_7d = (ahora + datetime.timedelta(days=7)).replace(hour=23, minute=59, second=59).isoformat()
        res = (
            service.events()
            .list(
                calendarId="primary",
                timeMin=inicio_dia,
                timeMax=fin_7d,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        calendar_conectado = True
        for ev in res.get("items", []):
            start_raw = ev.get("start", {}).get("dateTime") or ev.get("start", {}).get("date") or ""
            if "T" in start_raw:
                dt = datetime.datetime.fromisoformat(start_raw)
                f_str = dt.strftime("%Y-%m-%d")
                h_str = dt.strftime("%H:%M")
            else:
                f_str = start_raw[:10]
                h_str = "Todo el día"
            eventos_gcal.append({
                "id": ev.get("id", ""),
                "titulo": ev.get("summary", "Sin título"),
                "fecha": f_str,
                "hora": h_str,
                "duracion_min": 60,
                "tipo": "DIARIO" if ev.get("recurringEventId") else "PUNTUAL",
                "responsable": "Google Calendar",
                "estado": "SINCRONIZADO GCAL",
                "origen": "Google Calendar API",
            })
    except Exception:
        calendar_conectado = False

    # Combinar sin duplicar por (titulo, fecha, hora)
    vistos = set()
    todos: List[Dict[str, Any]] = []
    for ev in eventos_gcal + eventos_locales:
        ev_copy = dict(ev)
        # Si es recurrente diario, proyectar en hoy si su fecha era anterior
        if ev_copy.get("tipo") == "DIARIO" and ev_copy.get("fecha", "") < hoy_str:
            ev_copy["fecha"] = hoy_str
        clave = (ev_copy.get("titulo", "").strip().upper(), ev_copy.get("fecha", ""), ev_copy.get("hora", ""))
        if clave in vistos:
            continue
        vistos.add(clave)
        todos.append(ev_copy)

    todos.sort(key=lambda x: (x.get("fecha", ""), x.get("hora", "")))
    eventos_hoy = [e for e in todos if e.get("fecha") == hoy_str or e.get("tipo") == "DIARIO"]
    eventos_semana = [e for e in todos if hoy_str <= e.get("fecha", "") <= limite_7d or e.get("tipo") == "DIARIO"]

    return {
        "ok": True,
        "calendar_conectado": calendar_conectado,
        "bot_token_configurado": bool(TELEGRAM_AGENDA_TOKEN),
        "fecha_hoy": ahora.strftime("%d/%m/%Y"),
        "kpis": {
            "total_hoy": len(eventos_hoy),
            "total_semana": len(eventos_semana),
            "recurrentes_diarios": sum(1 for e in todos if e.get("tipo") == "DIARIO"),
            "puntuales": sum(1 for e in todos if e.get("tipo") != "DIARIO"),
        },
        "eventos_hoy": eventos_hoy,
        "eventos_semana": eventos_semana,
    }


def crear_evento_agenda(
    titulo: str,
    fecha: str,
    hora: str,
    duracion_min: int = 60,
    es_diario: bool = False,
    responsable: str = "Gerencia CEA",
    notificar_telegram: bool = True,
    chat_id_override: str = "",
) -> Tuple[bool, str]:
    ahora = datetime.datetime.now(LOCAL_TZ)
    fecha_usar = fecha.strip() if fecha else ahora.strftime("%Y-%m-%d")
    if "/" in fecha_usar:
        try:
            fecha_usar = datetime.datetime.strptime(fecha_usar, "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            pass
    hora_usar = hora.strip() if hora else "09:00"

    nuevo_ev = {
        "id": f"ev-{int(ahora.timestamp())}",
        "titulo": titulo.strip(),
        "fecha": fecha_usar,
        "hora": hora_usar,
        "duracion_min": int(duracion_min or 60),
        "tipo": "DIARIO" if es_diario else "PUNTUAL",
        "responsable": responsable.strip() or "Gerencia CEA",
        "estado": "PROGRAMADO",
        "origen": "Dashboard Web + Bot Agenda",
    }

    # Guardar localmente siempre
    evs = _cargar_eventos_locales()
    evs.insert(0, nuevo_ev)
    _guardar_eventos_locales(evs)

    # Intentar sincronizar con Google Calendar si hay conexión
    gcal_msg = "Guardado en Agenda CEA"
    try:
        service = get_calendar_service()
        h, m = map(int, hora_usar.split(":")[:2])
        inicio_dt = datetime.datetime.strptime(f"{fecha_usar} {h:02d}:{m:02d}", "%Y-%m-%d %H:%M").replace(tzinfo=LOCAL_TZ)
        fin_dt = inicio_dt + datetime.timedelta(minutes=int(duracion_min or 60))
        cuerpo: Dict[str, Any] = {
            "summary": titulo.strip(),
            "description": f"Responsable: {responsable}",
            "start": {"dateTime": inicio_dt.isoformat(), "timeZone": TIMEZONE_STR},
            "end": {"dateTime": fin_dt.isoformat(), "timeZone": TIMEZONE_STR},
        }
        if es_diario:
            cuerpo["recurrence"] = ["RRULE:FREQ=DAILY"]
        service.events().insert(calendarId="primary", body=cuerpo).execute()
        nuevo_ev["estado"] = "SINCRONIZADO GCAL"
        _guardar_eventos_locales(evs)
        gcal_msg = "Sincronizado en Google Calendar"
    except Exception:
        pass

    if notificar_telegram:
        enviar_recordatorio_telegram_agenda(
            titulo=titulo,
            fecha=fecha_usar,
            hora=hora_usar,
            tipo="DIARIO RECURRENTE" if es_diario else "EVENTO PUNTUAL",
            responsable=responsable,
            chat_id=chat_id_override,
        )

    return True, f"Evento '{titulo}' ({'Diario' if es_diario else fecha_usar} a las {hora_usar}) creado — {gcal_msg}."


def eliminar_evento_agenda(ev_id: str) -> Tuple[bool, str]:
    evs = _cargar_eventos_locales()
    nuevos = [e for e in evs if str(e.get("id")) != str(ev_id)]
    _guardar_eventos_locales(nuevos)
    return True, "Evento eliminado de la agenda."


def enviar_recordatorio_telegram_agenda(
    titulo: str,
    fecha: str,
    hora: str,
    tipo: str = "RECORDATORIO",
    responsable: str = "CEA Auto Oriente",
    chat_id: str = "",
) -> Tuple[bool, str]:
    cid = (
        chat_id
        or USUARIO_CHAT_ID
        or (os.getenv("TELEGRAM_CHAT_IDS", "") or os.getenv("ADMIN_CHAT_ID", "")).split(",")[0].strip()
    )
    if not cid:
        return False, "Ingresa tu Chat ID de Telegram para recibir el recordatorio."

    texto = (
        f"⏰ *AGENDA CEA AUTO ORIENTE — {tipo}*\n\n"
        f"📌 *Actividad:* {titulo}\n"
        f"📅 *Fecha:* {fecha}\n"
        f"🕒 *Hora:* {hora}\n"
        f"👤 *Responsable:* {responsable}"
    )

    for tk in (TELEGRAM_AGENDA_TOKEN, os.getenv("TELEGRAM_BOT_TOKEN", "").strip()):
        if not tk:
            continue
        try:
            url = f"https://api.telegram.org/bot{tk}/sendMessage"
            payload = json.dumps({"chat_id": cid, "text": texto, "parse_mode": "Markdown"}).encode("utf-8")
            req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8", errors="ignore"))
                if data.get("ok"):
                    return True, f"Notificación de agenda enviada a Telegram (Chat {cid})."
        except Exception:
            continue
    return False, "No se pudo conectar con Telegram API (verifica conexión o el Chat ID)."

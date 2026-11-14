"""
=============================================================================
CRM WHATSAPP IA, EMBUDO ESTADÍSTICO DE ABANDONO Y MOTOR DE REMARKETING — CEA
=============================================================================
Permite:
1. Llevar el registro completo de clientes que consultan por WhatsApp Business
   (Anuncios de Facebook/Instagram `fb.me`, Agente de IA de Meta `Responde la IA`,
   `Transferencia de IA`).
2. Medir estadísticamente cuántas personas preguntan y se quedan sin concluir
   tras la respuesta de la IA (% de Abandono IA vs. % de Conversión Real).
3. Cruzar automáticamente los prospectos de WhatsApp contra el archivo activo
   de Matrículas SISEC / Contabilidad DIAN para detectar quiénes ya se matricularon.
4. Ejecutar campañas de Remarketing desde plantillas persuasivas (individual o por lote)
   vía WhatsApp API o enlaces directos `wa.me`.
"""

import datetime
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent
LEADS_DB_PATH = BASE_DIR / "leads_whatsapp_cea.json"
PLANTILLAS_DB_PATH = BASE_DIR / "plantillas_remarketing_cea.json"

TARIFA_ESTIMADA_CAT = {
    "A2": 854000,
    "B1": 1204000,
    "C1": 1450000,
    "COMBO A2+B1": 1800000,
    "COMBO A2+C1": 1950000,
    "RECATEGORIZACIÓN C1": 750000,
    "POR DEFINIR": 1100000,
}

PLANTILLAS_DEFAULT: Dict[str, Dict[str, str]] = {
    "IA_RECUPERACION_50": {
        "nombre": "🔥 Recuperación IA — Inicia hoy con el 50% (General)",
        "texto": (
            "¡Hola {nombre}! 👋 Notamos que nos escribiste a *CEA Auto Oriente* (Sede Bolarquí, Cra 23 #55-19) "
            "para consultar por tu licencia de conducción *{categoria}* y quedamos atentos a tu elección.\n\n"
            "✅ Recuerda que nuestro paquete incluye *TODO*: Clases teóricas y prácticas, exámenes médicos autorizados "
            "y derechos de tránsito.\n"
            "💳 *Facilidad de pago:* Puedes iniciar hoy mismo con el *50%* y pagar el saldo al terminar la teoría.\n\n"
            "¿Quieres que te reserve el cupo con el beneficio de esta semana o te agendo de una vez el examen médico?"
        ),
    },
    "MOTO_A2_RAPIDO": {
        "nombre": "🏍️ Cierre Curso Moto A2 (Duración 11 días - Todo Incluido)",
        "texto": (
            "¡Hola {nombre}! 🏍️ Retomando tu consulta en *CEA Auto Oriente*: el curso de *Moto A2* tiene una duración "
            "de solo *11 días* (100% presencial en Bolarquí) e incluye exámenes médicos, derechos RUNT y todas tus clases.\n\n"
            "Puedes separar tu cupo e iniciar con el *50%*. ¿Te envío los horarios disponibles para empezar esta semana?"
        ),
    },
    "CARRO_B1_C1": {
        "nombre": "🚗 Cierre Carro Particular B1 / Público C1 (20 días)",
        "texto": (
            "¡Hola {nombre}! 🚗 En *CEA Auto Oriente* tenemos cupos abiertos para *Carro B1 y Servicio Público C1* "
            "(mínimo 20 días, práctica en vía paso a paso a tu propio ritmo + exámenes médicos y derechos incluidos).\n\n"
            "Inicias con el *50%* y el resto al terminar teoría. ¿Te gustaría que te agende tu primera clase?"
        ),
    },
    "COMBO_MOTO_CARRO": {
        "nombre": "⚡ Promoción Combo Moto A2 + Carro B1/C1",
        "texto": (
            "¡Hola {nombre}! 🔥 Si realizas el *Combo Moto + Carro* en *CEA Auto Oriente* ahorras en exámenes médicos "
            "y derechos de tránsito, y sales con ambas licencias listas.\n\n"
            "Recuerda que puedes iniciar con el *50%* en Cra 23 #55-19 Bolarquí. ¿Te comparto la tarifa especial del combo?"
        ),
    },
}


def _norm_txt(s: str) -> str:
    s = str(s or "").strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^A-Z0-9 ]+", " ", s).strip()


def _limpiar_telefono(tel: str) -> str:
    digits = re.sub(r"\D+", "", str(tel or ""))
    if len(digits) == 10 and digits.startswith("3"):
        return "57" + digits
    return digits


def cargar_plantillas() -> Dict[str, Dict[str, str]]:
    plantillas = dict(PLANTILLAS_DEFAULT)
    if PLANTILLAS_DB_PATH.exists():
        try:
            saved = json.loads(PLANTILLAS_DB_PATH.read_text(encoding="utf-8"))
            if isinstance(saved, dict):
                plantillas.update(saved)
        except Exception:
            pass
    return plantillas


def guardar_plantilla(codigo: str, nombre: str, texto: str) -> Dict[str, Dict[str, str]]:
    plantillas = cargar_plantillas()
    cod = re.sub(r"[^A-Z0-9_]+", "_", codigo.strip().upper()) or "PERSONALIZADA"
    plantillas[cod] = {"nombre": nombre.strip() or cod, "texto": texto.strip()}
    try:
        PLANTILLAS_DB_PATH.write_text(json.dumps(plantillas, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return plantillas


def _cargar_leads_raw() -> List[Dict[str, Any]]:
    if LEADS_DB_PATH.exists():
        try:
            data = json.loads(LEADS_DB_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list) and data:
                return data
        except Exception:
            pass

    ahora = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
    hoy_str = ahora.strftime("%Y-%m-%d")
    semilla = [
        {
            "id": "wa-1001",
            "telefono": "573112472122",
            "nombre": "YRTD (Prospecto Anuncio FB)",
            "cedula": "",
            "categoria_interes": "A2",
            "origen": "Anuncio Facebook/IG (fb.me)",
            "estado_ia": "RESPONDE_IA_PENDIENTE",
            "ultimo_mensaje_cliente": "¡Hola! Quiero más información",
            "ultima_respuesta_ia": "YRTD, para que compares mejor, el curso de moto A2 dura 11 días y el de carro B1 o C1 mínimo 20 días, puedes iniciar con el 50%...",
            "fecha_consulta": f"{hoy_str} 09:43",
            "fecha_ultimo_contacto": f"{hoy_str} 10:43",
            "intentos_remarketing": 0,
            "notas": "La IA respondió 2 veces (09:43 y 10:43 a.m.). Quedó sin responder si desea Moto, Carro o Combo.",
        },
        {
            "id": "wa-1002",
            "telefono": "573213678326",
            "nombre": "Mary",
            "cedula": "",
            "categoria_interes": "B1",
            "origen": "WhatsApp Directo / Anuncio",
            "estado_ia": "TRANSFERENCIA_ASESOR",
            "ultimo_mensaje_cliente": "hola buen día le escribe Mary",
            "ultima_respuesta_ia": "Saludo inicial recibido — Pendiente respuesta de cierre comercial",
            "fecha_consulta": f"{hoy_str} 10:49",
            "fecha_ultimo_contacto": f"{hoy_str} 10:49",
            "intentos_remarketing": 0,
            "notas": "Escribió a las 10:49 a.m., lista para atención humana o plantilla de cierre.",
        },
        {
            "id": "wa-1003",
            "telefono": "573017837069",
            "nombre": "Rocío",
            "cedula": "",
            "categoria_interes": "COMBO A2+B1",
            "origen": "Anuncio Meta IA",
            "estado_ia": "RESPONDE_IA_PENDIENTE",
            "ultimo_mensaje_cliente": "Precio del curso de conducción",
            "ultima_respuesta_ia": "Rocío, retomando tu consulta sobre los cursos, te dejo un dato útil...",
            "fecha_consulta": f"{hoy_str} 09:15",
            "fecha_ultimo_contacto": f"{hoy_str} 10:29",
            "intentos_remarketing": 1,
            "notas": "Agente IA envió mensaje de seguimiento automático a las 10:29 a.m., aún sin concretar matrícula.",
        },
        {
            "id": "wa-1004",
            "telefono": "573143018118",
            "nombre": "Wilson Jhoyner (@wilson_jhoyner31)",
            "cedula": "",
            "categoria_interes": "A2",
            "origen": "Instagram / Facebook Ads",
            "estado_ia": "RESPONDE_IA_PENDIENTE",
            "ultimo_mensaje_cliente": "Información curso de moto",
            "ultima_respuesta_ia": "Gracias por preferirnos. Esperamos poder servirte nuevamente.",
            "fecha_consulta": f"{hoy_str} 10:44",
            "fecha_ultimo_contacto": f"{hoy_str} 10:44",
            "intentos_remarketing": 0,
            "notas": "Consultó desde Instagram (@wilson_jhoyner31), pendiente cierre.",
        },
    ]
    _guardar_leads_raw(semilla)
    return semilla


def _guardar_leads_raw(leads: List[Dict[str, Any]]) -> None:
    try:
        LEADS_DB_PATH.write_text(json.dumps(leads, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def obtener_panorama_whatsapp_crm(hallazgos_cartera: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    leads = _cargar_leads_raw()
    plantillas = cargar_plantillas()

    # Construir índice de matrículas/contabilidad para cruzar automáticamente quién ya se matriculó
    matriculas_por_doc: Dict[str, Dict[str, Any]] = {}
    matriculas_por_nombre: List[Tuple[set, Dict[str, Any]]] = []
    if hallazgos_cartera:
        fuente = hallazgos_cartera.get("todas_matriculas") or hallazgos_cartera.get("registros") or []
        for r in fuente:
            doc = re.sub(r"\D+", "", str(r.get("NUMERODOCUMENTO", "")))
            nom = _norm_txt(r.get("ALUMNO") or f"{r.get('NOMBRES', '')} {r.get('APELLIDOS', '')}")
            if doc:
                matriculas_por_doc[doc] = r
            tokens = {t for t in nom.split() if len(t) >= 3}
            if len(tokens) >= 2:
                matriculas_por_nombre.append((tokens, r))

    leads_enriquecidos: List[Dict[str, Any]] = []
    cnt_pendientes_ia = 0
    cnt_transferidos = 0
    cnt_remarketing = 0
    cnt_matriculados = 0
    cnt_descartados = 0
    ingreso_potencial_pendiente = 0.0
    ingreso_cerrado_real = 0.0
    por_categoria: Dict[str, Dict[str, int]] = {}

    for item in leads:
        ld = dict(item)
        tel_clean = _limpiar_telefono(ld.get("telefono", ""))
        ld["telefono"] = tel_clean
        cat = str(ld.get("categoria_interes", "A2")).strip().upper() or "POR DEFINIR"
        ld["categoria_interes"] = cat
        valor_est = float(TARIFA_ESTIMADA_CAT.get(cat, 1050000))
        ld["valor_estimado"] = valor_est

        # Cruzar contra Matrículas / Contabilidad si tiene cédula o nombre completo (>= 2 palabras)
        ced_clean = re.sub(r"\D+", "", str(ld.get("cedula", "")))
        match_mat = matriculas_por_doc.get(ced_clean) if ced_clean else None
        if not match_mat:
            nom_tokens = {t for t in _norm_txt(ld.get("nombre", "")).split() if len(t) >= 3 and t not in ("PROSPECTO", "ANUNCIO", "CLIENTE")}
            if len(nom_tokens) >= 2:
                for m_tokens, m_reg in matriculas_por_nombre:
                    if len(nom_tokens & m_tokens) >= 2:
                        match_mat = m_reg
                        break

        if match_mat:
            ld["estado_ia"] = "MATRICULADO"
            ld["matricula_cruzada"] = match_mat.get("NUMEROMATRICULA", "")
            ld["valor_pagado_real"] = float(match_mat.get("ABONOS", 0) or match_mat.get("VALORCURSO", 0))
        else:
            ld["matricula_cruzada"] = ""
            ld["valor_pagado_real"] = 0.0

        est = ld.get("estado_ia", "RESPONDE_IA_PENDIENTE")
        if cat not in por_categoria:
            por_categoria[cat] = {"total": 0, "pendientes_ia": 0, "matriculados": 0}
        por_categoria[cat]["total"] += 1

        if est == "MATRICULADO":
            cnt_matriculados += 1
            por_categoria[cat]["matriculados"] += 1
            ingreso_cerrado_real += ld["valor_pagado_real"] or valor_est
        elif est == "REMARKETING_ENVIADO":
            cnt_remarketing += 1
            ingreso_potencial_pendiente += valor_est
        elif est == "TRANSFERENCIA_ASESOR":
            cnt_transferidos += 1
            por_categoria[cat]["pendientes_ia"] += 1
            ingreso_potencial_pendiente += valor_est
        elif est == "DESCARTADO":
            cnt_descartados += 1
        else:
            cnt_pendientes_ia += 1
            por_categoria[cat]["pendientes_ia"] += 1
            ingreso_potencial_pendiente += valor_est

        leads_enriquecidos.append(ld)

    total = len(leads_enriquecidos)
    sin_concluir = cnt_pendientes_ia + cnt_transferidos + cnt_remarketing
    tasa_abandono_ia = round((sin_concluir / total) * 100, 1) if total > 0 else 0.0
    tasa_conversion = round((cnt_matriculados / total) * 100, 1) if total > 0 else 0.0

    return {
        "ok": True,
        "kpis": {
            "total_consultas": total,
            "pendientes_ia": cnt_pendientes_ia,
            "transferidos_asesor": cnt_transferidos,
            "en_remarketing": cnt_remarketing,
            "sin_concluir_total": sin_concluir,
            "matriculados": cnt_matriculados,
            "descartados": cnt_descartados,
            "tasa_abandono_ia_pct": tasa_abandono_ia,
            "tasa_conversion_pct": tasa_conversion,
            "ingreso_potencial_pendiente": round(ingreso_potencial_pendiente, 0),
            "ingreso_cerrado_real": round(ingreso_cerrado_real, 0),
        },
        "por_categoria": [
            {"categoria": k, **v} for k, v in sorted(por_categoria.items(), key=lambda x: -x[1]["total"])
        ],
        "plantillas": plantillas,
        "leads": leads_enriquecidos,
    }


def registrar_o_actualizar_lead(payload: Dict[str, Any]) -> Tuple[bool, str]:
    leads = _cargar_leads_raw()
    ahora_str = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).strftime("%Y-%m-%d %H:%M")

    tel = _limpiar_telefono(payload.get("telefono", ""))
    nombre = str(payload.get("nombre", "")).strip() or f"Prospecto {tel[-4:]}"
    cedula = str(payload.get("cedula", "")).strip()
    categoria = str(payload.get("categoria_interes", "A2")).strip().upper()
    estado_ia = str(payload.get("estado_ia", "RESPONDE_IA_PENDIENTE")).strip().upper()
    origen = str(payload.get("origen", "Anuncio Facebook/IG (Meta IA)")).strip()
    ultimo_msg = str(payload.get("ultimo_mensaje_cliente", "¡Hola! Quiero más información")).strip()
    notas = str(payload.get("notas", "")).strip()

    if not tel and not nombre:
        return False, "Ingresa al menos el número de WhatsApp o nombre del prospecto."

    # Buscar si ya existe por ID o teléfono
    lead_id = str(payload.get("id", "")).strip()
    for ld in leads:
        if (lead_id and ld.get("id") == lead_id) or (tel and _limpiar_telefono(ld.get("telefono", "")) == tel):
            ld["nombre"] = nombre or ld.get("nombre", "")
            if cedula:
                ld["cedula"] = cedula
            if categoria:
                ld["categoria_interes"] = categoria
            if estado_ia:
                ld["estado_ia"] = estado_ia
            if ultimo_msg:
                ld["ultimo_mensaje_cliente"] = ultimo_msg
            if notas:
                ld["notas"] = notas
            ld["fecha_ultimo_contacto"] = ahora_str
            _guardar_leads_raw(leads)
            return True, f"Prospecto {ld['nombre']} (+{tel}) actualizado a {estado_ia}."

    nuevo = {
        "id": f"wa-{int(datetime.datetime.now().timestamp() * 1000)}",
        "telefono": tel,
        "nombre": nombre,
        "cedula": cedula,
        "categoria_interes": categoria,
        "origen": origen,
        "estado_ia": estado_ia,
        "ultimo_mensaje_cliente": ultimo_msg,
        "ultima_respuesta_ia": "Atendido inicialmente por Agente IA de Meta",
        "fecha_consulta": ahora_str,
        "fecha_ultimo_contacto": ahora_str,
        "intentos_remarketing": 0,
        "notas": notas or "Capturado en embudo WhatsApp IA",
    }
    leads.insert(0, nuevo)
    _guardar_leads_raw(leads)
    return True, f"Prospecto {nombre} (+{tel}) registrado en el embudo de WhatsApp IA."


def cambiar_estado_lead(lead_id: str, nuevo_estado: str, incrementar_remarketing: bool = False) -> Tuple[bool, str]:
    leads = _cargar_leads_raw()
    ahora_str = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).strftime("%Y-%m-%d %H:%M")
    for ld in leads:
        if str(ld.get("id")) == str(lead_id) or _limpiar_telefono(ld.get("telefono", "")) == _limpiar_telefono(lead_id):
            ld["estado_ia"] = nuevo_estado
            ld["fecha_ultimo_contacto"] = ahora_str
            if incrementar_remarketing:
                ld["intentos_remarketing"] = int(ld.get("intentos_remarketing", 0)) + 1
            _guardar_leads_raw(leads)
            return True, f"Estado actualizado a {nuevo_estado}."
    return False, "No se encontró el prospecto."


def importar_lote_whatsapp_texto(texto_bruto: str, categoria_default: str = "A2") -> Tuple[int, str]:
    """
    Analiza texto pegado desde WhatsApp Web o listas de teléfonos/nombres
    (ej: '+57 311 2472122 - Juan Perez - Moto A2' o bloques copiados de la barra lateral de WhatsApp Web)
    y registra automáticamente todos los prospectos detectados.
    """
    if not texto_bruto or not texto_bruto.strip():
        return 0, "Pega al menos una línea con teléfonos o chats de WhatsApp."

    leads = _cargar_leads_raw()
    telefonos_existentes = {_limpiar_telefono(x.get("telefono", "")): x for x in leads if x.get("telefono")}
    ahora_str = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).strftime("%Y-%m-%d %H:%M")

    agregados = 0
    lineas = [ln.strip() for ln in texto_bruto.splitlines() if ln.strip()]
    for ln in lineas:
        # Buscar número celular colombiano (+57 3xx xxx xxxx o 3xxxxxxxxx)
        m_tel = re.search(r"(?:\+?57[\s-]*)?(3\d{2}[\s.-]*\d{3}[\s.-]*\d{4})", ln)
        if not m_tel:
            continue
        tel = _limpiar_telefono(m_tel.group(0))
        resto = ln.replace(m_tel.group(0), " ").strip(" -•|:\t")

        # Detectar categoría si se menciona en la línea
        ln_up = ln.upper()
        cat = categoria_default
        if "COMBO" in ln_up:
            cat = "COMBO A2+B1"
        elif "C1" in ln_up or "PUBLICO" in ln_up or "PÚBLICO" in ln_up:
            cat = "C1"
        elif "B1" in ln_up or "CARRO" in ln_up:
            cat = "B1"
        elif "A2" in ln_up or "MOTO" in ln_up:
            cat = "A2"

        est = "RESPONDE_IA_PENDIENTE"
        if "TRANSFERENCIA" in ln_up or "ASESOR" in ln_up:
            est = "TRANSFERENCIA_ASESOR"

        nombre_limpio = re.sub(r"(RESPUESTA DE LA IA|RESPONDE LA IA|TRANSFERENCIA DE IA|\d{1,2}:\d{2}\s*[ap]\.?\s*m\.?)", "", resto, flags=re.IGNORECASE).strip(" -•|,:")
        if not nombre_limpio or len(nombre_limpio) < 2:
            nombre_limpio = f"Prospecto WhatsApp {tel[-4:]}"
        elif len(nombre_limpio) > 42:
            nombre_limpio = nombre_limpio[:42].strip()

        if tel in telefonos_existentes:
            telefonos_existentes[tel]["fecha_ultimo_contacto"] = ahora_str
            continue

        nuevo = {
            "id": f"wa-{int(datetime.datetime.now().timestamp() * 1000)}-{agregados}",
            "telefono": tel,
            "nombre": nombre_limpio,
            "cedula": "",
            "categoria_interes": cat,
            "origen": "Importado de WhatsApp Web",
            "estado_ia": est,
            "ultimo_mensaje_cliente": resto[:120] if resto else "Consulta desde anuncio de WhatsApp",
            "ultima_respuesta_ia": "Respondió Agente IA — Pendiente cierre",
            "fecha_consulta": ahora_str,
            "fecha_ultimo_contacto": ahora_str,
            "intentos_remarketing": 0,
            "notas": "Importado en lote para campaña de remarketing",
        }
        leads.insert(0, nuevo)
        telefonos_existentes[tel] = nuevo
        agregados += 1

    _guardar_leads_raw(leads)
    return agregados, f"Se importaron {agregados} nuevos prospectos al embudo de WhatsApp IA."

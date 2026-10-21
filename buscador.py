"""Buscador de proveedores de TS.

Revisa los correos nuevos de la bandeja de entrada (Microsoft Graph), detecta
órdenes de compra o solicitudes de cotización, le pide a Claude que busque
proveedores en internet y envía el resultado por correo al dueño.

Variables de entorno (en GitHub van como secretos):
  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET  registro de la app en Microsoft Entra
  BUZON                 correo que se revisa (el de tu papá)
  ENVIAR_A              a quién le llega el resultado (por defecto, BUZON)
  ANTHROPIC_API_KEY     clave de la API de Claude
Opcionales:
  ESTADO_ARCHIVO        dónde se guardan los correos ya revisados (estado.json)
  HORAS_INICIALES       en la primera ejecución, cuántas horas hacia atrás revisar (2)
  SIMULAR=1             no envía correos: imprime el resultado
"""

import base64
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import anthropic
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
MODELO = "claude-opus-5-5"
PREFIJO_ASUNTO = "[Proveedores]"
PALABRAS_CLAVE = [
    "cotizacion", "cotizar", "cotice", "orden de compra", "oc ", "oc-", "oc:", "o.c",
    "rfq", "requerimos", "requerimiento", "solicitamos", "solicitud", "necesitamos",
    "purchase order", "quotation", "quote",
]
MAX_PDF_BYTES = 10 * 1024 * 1024
MAX_RESTARTS = 5

INSTRUCCIONES = (Path(__file__).parent / "instrucciones_agente.md").read_text(encoding="utf-8")


def sin_tildes(texto):
    return texto.lower().translate(str.maketrans("áéíóúñ", "aeioun"))


# ---------- Microsoft Graph ----------

def token_graph():
    r = requests.post(
        f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}/oauth2/v2.0/token",
        data={
            "client_id": os.environ["MS_CLIENT_ID"],
            "client_secret": os.environ["MS_CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def correos_nuevos(token, buzon, desde):
    url = f"{GRAPH}/users/{buzon}/mailFolders/inbox/messages"
    params = {
        "$filter": f"receivedDateTime ge {desde.strftime('%Y-%m-%dT%H:%M:%SZ')}",
        "$select": "id,internetMessageId,subject,from,receivedDateTime,body,hasAttachments",
        "$orderby": "receivedDateTime asc",
        "$top": "50",
    }
    headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.body-content-type="text"'}
    correos = []
    while url:
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        datos = r.json()
        correos.extend(datos.get("value", []))
        url, params = datos.get("@odata.nextLink"), None
    return correos


def pdfs_adjuntos(token, buzon, mensaje_id):
    r = requests.get(
        f"{GRAPH}/users/{buzon}/messages/{mensaje_id}/attachments",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    r.raise_for_status()
    pdfs = []
    for a in r.json().get("value", []):
        es_pdf = a.get("contentType") == "application/pdf" or a.get("name", "").lower().endswith(".pdf")
        if a.get("@odata.type") == "#microsoft.graph.fileAttachment" and es_pdf and a.get("size", 0) <= MAX_PDF_BYTES:
            pdfs.append({"nombre": a["name"], "base64": a["contentBytes"]})
    return pdfs


def enviar_correo(token, buzon, para, asunto, html):
    r = requests.post(
        f"{GRAPH}/users/{buzon}/sendMail",
        headers={"Authorization": f"Bearer {token}"},
        json={
            "message": {
                "subject": asunto,
                "body": {"contentType": "HTML", "content": html},
                "toRecipients": [{"emailAddress": {"address": d.strip()}} for d in para.split(",") if d.strip()],
            },
            "saveToSentItems": True,
        },
        timeout=30,
    )
    r.raise_for_status()


# ---------- Filtro y estado ----------

def remitente(correo):
    return (correo.get("from") or {}).get("emailAddress", {}).get("address", "").lower()


def pasa_filtro(correo, buzon):
    asunto = correo.get("subject") or ""
    if asunto.startswith(PREFIJO_ASUNTO) or remitente(correo) == buzon.lower():
        return False
    texto = sin_tildes(f"{asunto}\n{(correo.get('body') or {}).get('content', '')[:5000]}")
    return any(p in texto for p in PALABRAS_CLAVE)


def leer_estado(ruta):
    try:
        return json.loads(Path(ruta).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {"ultima_revision": None, "procesados": []}


def guardar_estado(ruta, estado):
    estado["procesados"] = estado["procesados"][-500:]
    Path(ruta).write_text(json.dumps(estado, indent=2, ensure_ascii=False), encoding="utf-8")


def calcular_desde(estado, ahora, horas_iniciales):
    if estado.get("ultima_revision"):
        # Media hora de traslape por si GitHub Actions se retrasa; los duplicados
        # se descartan con la lista de procesados.
        desde = datetime.fromisoformat(estado["ultima_revision"]) - timedelta(minutes=30)
        return max(desde, ahora - timedelta(hours=24))
    return ahora - timedelta(hours=horas_iniciales)


# ---------- Claude ----------

def contenido_para_claude(correo, pdfs):
    texto = (
        "CORREO RECIBIDO\n"
        f"De: {remitente(correo)}\n"
        f"Fecha: {correo.get('receivedDateTime')}\n"
        f"Asunto: {correo.get('subject') or ''}\n\n"
        f"{(correo.get('body') or {}).get('content', '')[:30000]}"
    )
    bloques = [
        {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": p["base64"]},
            "title": p["nombre"],
        }
        for p in pdfs
    ]
    bloques.append({"type": "text", "text": texto})
    return bloques


def preguntar_a_claude(cliente, correo, pdfs):
    mensajes = [{"role": "user", "content": contenido_para_claude(correo, pdfs)}]
    for _ in range(MAX_RESTARTS + 1):
        with cliente.beta.messages.stream(
            model=MODELO,
            max_tokens=32000,
            system=INSTRUCCIONES,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            output_config={"effort": "high"},
            tools=[{
                "type": "web_search_20260209",
                "name": "web_search",
                "max_uses": 20,
                "user_location": {"type": "approximate", "country": "CO"},
            }],
            messages=mensajes,
        ) as stream:
            respuesta = stream.get_final_message()
        if respuesta.stop_reason == "pause_turn":
            mensajes.append({"role": "assistant", "content": respuesta.content})
            continue
        if respuesta.stop_reason == "refusal":
            raise RuntimeError("Claude rechazó la solicitud")
        return "".join(b.text for b in respuesta.content if b.type == "text")
    raise RuntimeError("La búsqueda no terminó después de varios intentos")


def interpretar(texto):
    """Devuelve (es_solicitud, asunto, html) a partir del formato pedido en las instrucciones."""
    m = re.search(r"ES_SOLICITUD:\s*(SI|SÍ|NO)", texto, re.IGNORECASE)
    if not m or m.group(1).upper() == "NO":
        return False, "", ""
    asunto = re.search(r"ASUNTO:\s*(.+)", texto)
    partes = texto.split("---HTML---", 1)
    html = partes[1].strip() if len(partes) == 2 else ""
    html = re.sub(r"^```(?:html)?\s*|\s*```$", "", html)
    return True, (asunto.group(1).strip() if asunto else "Solicitud nueva"), html


# ---------- Principal ----------

def main():
    buzon = os.environ["BUZON"]
    para = os.environ.get("ENVIAR_A") or buzon
    simular = os.environ.get("SIMULAR") == "1"
    ruta_estado = os.environ.get("ESTADO_ARCHIVO", "estado.json")
    ahora = datetime.now(timezone.utc)

    estado = leer_estado(ruta_estado)
    desde = calcular_desde(estado, ahora, float(os.environ.get("HORAS_INICIALES", "2")))
    token = token_graph()
    cliente = anthropic.Anthropic()

    correos = correos_nuevos(token, buzon, desde)
    candidatos = [
        c for c in correos
        if c["internetMessageId"] not in estado["procesados"] and pasa_filtro(c, buzon)
    ]
    print(f"{len(correos)} correos desde {desde:%Y-%m-%d %H:%M} UTC, {len(candidatos)} parecen solicitudes")

    fallos = []
    for correo in candidatos:
        asunto_original = correo.get("subject") or "(sin asunto)"
        try:
            pdfs = pdfs_adjuntos(token, buzon, correo["id"]) if correo.get("hasAttachments") else []
            es_solicitud, asunto, html = interpretar(preguntar_a_claude(cliente, correo, pdfs))
            if es_solicitud and html:
                asunto = f"{PREFIJO_ASUNTO} {asunto}"
                if simular:
                    print(f"\n=== {asunto} (para {para}) ===\n{html}\n")
                else:
                    enviar_correo(token, buzon, para, asunto, html)
                print(f"Enviado: {asunto}  <-  {asunto_original}")
            else:
                print(f"No es solicitud: {asunto_original}")
            estado["procesados"].append(correo["internetMessageId"])
        except Exception as e:  # un correo que falla no detiene los demás; se reintenta en la próxima revisión
            fallos.append(correo["receivedDateTime"])
            print(f"ERROR con '{asunto_original}': {e}", file=sys.stderr)

    # Si algo falló, la próxima revisión arranca desde ese correo para reintentarlo.
    reanudar = min([datetime.fromisoformat(f.replace("Z", "+00:00")) + timedelta(minutes=30) for f in fallos] + [ahora])
    estado["ultima_revision"] = reanudar.isoformat()
    if not simular:  # una simulación no cuenta como revisión: la siguiente ejecución real procesa los mismos correos
        guardar_estado(ruta_estado, estado)
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())

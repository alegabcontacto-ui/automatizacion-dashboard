"""
generar_cotizacion.py
Genera el PDF de cotización IMSS y envía la respuesta automática por Outlook COM.

MODO_PRUEBA = True  →  todos los correos van a CORREO_PRUEBA (no al remitente real)
MODO_PRUEBA = False →  responde al remitente original del correo recibido
"""

import json
import logging
import random
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    HRFlowable,
    Image,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    Paragraph,
)

log = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
#  CONFIGURACIÓN  —  ajustar antes de pasar a producción
# ──────────────────────────────────────────────────────────────────────────────
EMPRESA = dict(
    nombre    = "COMERCIALIZADORA Y OPERADORA DE SERVICIOS ADEMEX S. DE R.L. DE C.V.",
    rfc       = "COS190816K19",
    tel       = "5565537507 / 5548556411",
    domicilio = "CALLE GIOTTO 35 BIS INT 1 COL MIXCOAC, ALC BENITO JUAREZ, CIUDAD DE MÉXICO CP 03910",
    correo    = "ventas.ademex@serviciosademex.com",
    firmante  = "JUAN PÉREZ XXXX",
    logo      = r"C:\Users\AlejandroRodríguez-C\OneDrive - GADMAR SA DE CV\Imágenes\ademex.png",
)

CORREO_PRUEBA = "alejandro.rodriguez@pharmagbc.com"
MODO_PRUEBA   = True   # ← cambiar a False para enviar al remitente real

DIRECTORIO_PDF = Path(__file__).parent / "cotizaciones_generadas"

NOTAS = (
    "LOS PRECIOS OFERTADOS, PERMANECERÁN FIJOS DURANTE LA VIGENCIA DEL PEDIDO. "
    "EN EL CASO QUE EL INSTITUTO ME OTORGUE LA DEMANDA SOLICITADA, ME OBLIGO EN "
    "NOMBRE DE MI REPRESENTADA A SUSCRIBIR EL PEDIDO QUE SE DERIVE EN LOS TÉRMINOS, "
    "CONDICIONES Y PORCENTAJES ESTABLECIDOS EN ESTA ADJUDICACIÓN."
)

CUERPO_CORREO_HTML = """\
<html>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;margin:0;padding:0;">

<p><strong style="color:#215E99;">Buen d&iacute;a.</strong></p>

<p>Reciba un cordial saludo.</p>

<p>En referencia a la comunicaci&oacute;n precedente, anexo la&nbsp;<strong style="color:#215E99;">cotizaci&oacute;n oficial</strong>&nbsp;emitida por&nbsp;<strong style="color:#215E99;">ADEMEX&nbsp;S. DE R.L. DE C.V.</strong>, correspondiente a los insumos solicitados.</p>

<p>Con relaci&oacute;n a la calidad y cumplimiento normativo de los productos ofertados, se informa que:</p>

<ul style="margin-top:4px;margin-bottom:4px;">
  <li>Cada clave cuenta con&nbsp;<strong style="color:#215E99;">documentos de control y verificaci&oacute;n</strong>, incluyendo an&aacute;lisis emitidos por el fabricante.</li>
  <li>Todos los insumos poseen&nbsp;<strong style="color:#215E99;">Registros Sanitarios COFEPRIS vigentes</strong>, que garantizan su legalidad y correcta fabricaci&oacute;n.</li>
  <li>En caso de productos termo sensibles, se realiza manejo bajo&nbsp;<strong>cadena de fr&iacute;o</strong>, con evidencia documentada de monitoreo de temperatura.</li>
  <li>Se dispone de&nbsp;<strong style="color:#215E99;">informaci&oacute;n t&eacute;cnica y trazabilidad completa</strong>&nbsp;para consulta de la unidad cuando as&iacute; se requiera.</li>
</ul>

<p>Agradecemos la oportunidad de participar en su proceso de adquisici&oacute;n&nbsp;y reiteramos el compromiso de proveer productos&nbsp;<strong style="color:#215E99;">seguros, confiables y conformes al marco regulatorio vigente</strong>, contribuyendo al adecuado abastecimiento de la unidad.</p>

<p>Quedo a su disposici&oacute;n para atender cualquier aclaraci&oacute;n, requerimiento adicional o documentaci&oacute;n complementaria que considere necesaria.</p>

<br>
<p style="font-size:8.5pt;color:#555555;">
<strong style="color:#215E99;">AVISO DE CONFIDENCIALIDAD:</strong>
Este correo electr&oacute;nico, incluyendo en su caso, los archivos adjuntos al mismo pueden contener informaci&oacute;n de car&aacute;cter confidencial y/o privilegiada, y se env&iacute;an a la atenci&oacute;n &uacute;nica y exclusivamente de la persona y/o entidad a quien va dirigido. La copia, revisi&oacute;n, uso, revelaci&oacute;n y/o distribuci&oacute;n de dicha informaci&oacute;n confidencial sin la autorizaci&oacute;n por escrito de
<strong style="color:#215E99;">COMERCIALIZADORA Y OPERADORA DE SERVICIOS ADEMEX&nbsp;S. DE R.L. DE C.V.</strong>
(o cualquiera de sus afiliadas o subsidiarias) est&aacute; prohibida. Si usted no es el destinatario a quien se dirige el presente correo, favor de contactar al remitente respondiendo al presente correo y eliminar el correo original incluyendo sus archivos, as&iacute; como cualquiera copia de este. Mediante la recepci&oacute;n del presente correo usted reconoce y acepta que en caso de incumplimiento de su parte y/o de sus representantes a los t&eacute;rminos antes mencionados,
<strong style="color:#215E99;">COMERCIALIZADORA Y OPERADORA DE SERVICIOS&nbsp;ADEMEX&nbsp;S. DE R.L. DE C.V.</strong>
tendr&aacute; derecho a los da&ntilde;os y perjuicios que esto le cause.
</p>

</body>
</html>"""

MESES = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE",
]


# ──────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def _monto_letras(monto: float) -> str:
    entero   = int(monto)
    centavos = round((monto - entero) * 100)
    try:
        from num2words import num2words
        palabras = num2words(entero, lang="es").upper()
    except Exception:
        palabras = str(entero)
    return f"({palabras} PESOS {centavos:02d}/100 M.N.)"


def _precio_aleatorio(clave: str) -> float:
    seed = sum(ord(c) for c in str(clave)) % 99991
    return round(random.Random(seed).uniform(50.0, 4500.0), 2)


def _extraer_evento(asunto: str) -> str:
    m = re.search(r"(REQ\.?\s*MED\.?\s*[\d\-]+)", asunto or "", re.IGNORECASE)
    if m:
        return m.group(1).upper()
    m2 = re.search(r"(\d{5,})", asunto or "")
    return m2.group(1) if m2 else (asunto or "SIN EVENTO")[:60].upper()


def _p(text, style) -> Paragraph:
    return Paragraph(str(text), style)


# ──────────────────────────────────────────────────────────────────────────────
#  GENERACIÓN DEL PDF
# ──────────────────────────────────────────────────────────────────────────────
def generar_pdf(df_partidas: pd.DataFrame, asunto: str, fecha_correo=None) -> Path:
    """
    Genera el PDF de cotización replicando el formato IMSS de ADEMEX.
    df_partidas: DataFrame con columnas GPO/GEN/ESP/DIF/VAR/CANTIDAD/CLAVE
                 (acepta mayúsculas o minúsculas).
    Retorna la ruta del PDF generado.
    """
    DIRECTORIO_PDF.mkdir(exist_ok=True)

    df = df_partidas.copy()
    df.columns = [c.lower() for c in df.columns]

    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = DIRECTORIO_PDF / f"Cotizacion_{ts}.pdf"

    fecha_dt  = fecha_correo if isinstance(fecha_correo, datetime) else datetime.now()
    fecha_str = f"{fecha_dt.day} DE {MESES[fecha_dt.month-1]} DEL {fecha_dt.year}"

    # ── Medidas de página ──────────────────────────────────────────────────────
    W, H = landscape(letter)
    mg   = 0.65 * cm
    uw   = W - 2 * mg          # ancho útil total ≈ 26.6 cm en horizontal

    # ── Estilos ────────────────────────────────────────────────────────────────
    def ps(name, font="Helvetica", size=7, leading=None, align=None):
        kw = dict(fontName=font, fontSize=size, leading=leading or size + 1.5)
        if align is not None:
            kw["alignment"] = align
        return ParagraphStyle(name, **kw)

    S6   = ps("S6",   size=6,   leading=7.5)
    S65  = ps("S65",  size=6.5, leading=8)
    S65B = ps("S65B", font="Helvetica-Bold", size=6.5, leading=8)
    S6C  = ps("S6C",  size=6,   leading=7.5, align=TA_CENTER)
    S6CB = ps("S6CB", font="Helvetica-Bold", size=6, leading=7.5, align=TA_CENTER)
    S7R  = ps("S7R",  size=7,   align=TA_RIGHT)
    S7RB = ps("S7RB", font="Helvetica-Bold", size=7, align=TA_RIGHT)
    S8C  = ps("S8C",  size=8,   align=TA_CENTER)
    S8CB = ps("S8CB", font="Helvetica-Bold", size=8, align=TA_CENTER)
    SN   = ps("SN",   size=6.5, leading=9)

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(letter),
        leftMargin=mg, rightMargin=mg,
        topMargin=mg,  bottomMargin=mg,
    )
    story = []

    # ── 1. ENCABEZADO ──────────────────────────────────────────────────────────
    logo_path = Path(EMPRESA["logo"])
    logo_w    = 3.6 * cm
    info_w    = uw - logo_w     # ← ancho real disponible para el bloque de texto

    if logo_path.exists():
        logo_cell = Image(str(logo_path), width=logo_w, height=2.0*cm, kind="proportional")
    else:
        logo_cell = _p("<b>ADEMEX</b>", S65B)

    evento = _extraer_evento(asunto)

    # 5 filas de info; cada fila tiene celda izquierda (60 %) y derecha (40 %)
    lw = info_w * 0.60   # ≈ 9.96 cm
    rw = info_w * 0.40   # ≈ 6.64 cm — suficiente para RFC + TEL separados

    info_rows = [
        [
            _p(f"<b>NOMBRE:</b> INSTITUTO MEXICANO DEL SEGURO SOCIAL", S65),
            _p(f"<b>EVENTO:</b>  {evento}", S65),
        ],
        [
            _p(f"<b>FECHA:</b> {fecha_str}", S65),
            _p("<b>FAB. (  ).  DIST. ( X ).  No. DE PREI IMSS:</b>  ---", S65),
        ],
        [
            _p(f"<b>NOMBRE DEL PARTICIPANTE:</b>  {EMPRESA['nombre']}", S65),
            _p(f"<b>RFC:</b>  {EMPRESA['rfc']}", S65),
        ],
        [
            _p(f"<b>DOMICILIO:</b>  {EMPRESA['domicilio']}", S65),
            _p(f"<b>TEL.:</b>  {EMPRESA['tel']}", S65),
        ],
        [
            _p(f"<b>CORREO:</b>  {EMPRESA['correo']}", S65),
            _p("", S65),
        ],
    ]

    info_tbl = Table(info_rows, colWidths=[lw, rw])
    info_tbl.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
    ]))

    header_outer = Table([[logo_cell, info_tbl]], colWidths=[logo_w, info_w])
    header_outer.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    story.append(header_outer)
    story.append(Spacer(1, 0.2 * cm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.black))
    story.append(Spacer(1, 0.2 * cm))

    # ── 2. TABLA DE PARTIDAS ───────────────────────────────────────────────────
    # Anchos calculados para que la CLAVE (18 chars) y las columnas cortas
    # ("---", cantidades) quepan en una sola línea con padding mínimo.
    # Total = uw ≈ 20.3 cm
    cw = [
        1.30*cm,   # Partida
        4.00*cm,   # CLAVE
        4.00*cm,   # Descripción
        0.85*cm,   # Pres
        0.85*cm,   # Cant
        0.85*cm,   # Tipo
        1.85*cm,   # Registro Sanitario
        1.45*cm,   # Marca
        1.60*cm,   # País de Origen
        2.65*cm,   # Nombre del titular del Registro Sanitario
        1.10*cm,   # RFC
        1.25*cm,   # Cantidad
        2.40*cm,   # P.U
        2.65*cm,   # Importe
    ]
    # sum(cw) ≈ 26.65 cm  ≈  uw landscape

    pad_h = 2    # padding horizontal en puntos (izq y der de cada celda)
    pad_v = 2    # padding vertical

    encabezados = [
        _p("Part\nida",                              S6CB),
        _p("CLAVE",                                  S6CB),
        _p("Descripción",                            S6CB),
        _p("Pres",                                   S6CB),
        _p("Cant",                                   S6CB),
        _p("Tipo",                                   S6CB),
        _p("Registro\nSanitario",                    S6CB),
        _p("Marca",                                  S6CB),
        _p("País de\nOrigen",                        S6CB),
        _p("Nombre del\ntitular del\nReg. Sanitario",S6CB),
        _p("RFC",                                    S6CB),
        _p("Cantidad",                               S6CB),
        _p("P.U",                                    S6CB),
        _p("Importe",                                S6CB),
    ]

    filas    = [encabezados]
    subtotal = 0.0

    for idx, (_, row) in enumerate(df.iterrows(), start=1):
        clave = str(row.get("clave") or "")
        if not clave or clave.lower() == "nan":
            gpo = str(row.get("gpo", "000"))
            gen = str(row.get("gen", "000"))
            esp = str(row.get("esp", "0000"))
            dif = str(row.get("dif", "00"))
            var = str(row.get("var", "00"))
            clave = f"{gpo}.{gen}.{esp}.{dif}.{var}"

        try:
            cant = int(float(row.get("cantidad", 1))) or 1
        except (ValueError, TypeError):
            cant = 1

        pu      = _precio_aleatorio(clave)
        importe = round(cant * pu, 2)
        subtotal += importe

        partida_num = str(row.get("partida", idx))

        filas.append([
            _p(partida_num,          S6C),
            _p(clave,                S6C),
            _p("MEDICAMENTO IMSS",   S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("---",                S6C),
            _p("null",               S6C),
            _p(str(cant),            S6C),
            _p(f"${pu:,.2f}",        S6C),
            _p(f"${importe:,.2f}",   S6C),
        ])

    iva   = 0.0
    total = round(subtotal + iva, 2)

    tabla = Table(filas, colWidths=cw, repeatRows=1)
    tabla.setStyle(TableStyle([
        # Encabezado
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#d0d0d0")),
        ("ALIGN",         (0, 0), (-1, 0), "CENTER"),
        ("VALIGN",        (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, 0), 3),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
        ("LEFTPADDING",   (0, 0), (-1, 0), pad_h),
        ("RIGHTPADDING",  (0, 0), (-1, 0), pad_h),
        # Filas de datos
        ("ALIGN",         (0, 1), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 1), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 1), (-1, -1), pad_v),
        ("BOTTOMPADDING", (0, 1), (-1, -1), pad_v),
        ("LEFTPADDING",   (0, 1), (-1, -1), pad_h),
        ("RIGHTPADDING",  (0, 1), (-1, -1), pad_h),
        # CLAVE y Descripción: alineadas a la izquierda
        ("ALIGN",         (1, 1), (2, -1), "LEFT"),
        # Grid
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.black),
        # Filas alternas
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
    ]))
    story.append(tabla)
    story.append(Spacer(1, 0.35 * cm))

    # ── 3. TOTALES ─────────────────────────────────────────────────────────────
    tot_rows = [
        [_p("SubTotal", S7R),  _p(f"${subtotal:,.2f}", S7RB)],
        [_p("I.V.A.",   S7R),  _p(f"${iva:,.2f}",      S7RB)],
        [_p("Total",    S7RB), _p(f"${total:,.2f}",     S7RB)],
    ]
    tot_tbl = Table(tot_rows, colWidths=[4.0*cm, 2.8*cm], hAlign="RIGHT")
    tot_tbl.setStyle(TableStyle([
        ("LINEABOVE",     (0, 0), (-1, 0),  0.5, colors.black),
        ("LINEABOVE",     (0, 2), (-1, 2),  0.5, colors.black),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
    ]))
    story.append(tot_tbl)
    story.append(Spacer(1, 0.15 * cm))

    story.append(_p(_monto_letras(total), S65))
    story.append(Spacer(1, 0.4 * cm))

    # ── 4. NOTAS ───────────────────────────────────────────────────────────────
    story.append(_p(f"NOTAS: {NOTAS}", SN))
    story.append(Spacer(1, 0.8 * cm))

    # ── 5. FIRMA ───────────────────────────────────────────────────────────────
    firma_inner = Table(
        [
            [_p("ATENTAMENTE", S8C)],
            [Spacer(1, 0.9 * cm)],
            [HRFlowable(width=5 * cm, thickness=0.5, color=colors.black)],
            [_p(EMPRESA["firmante"], S8CB)],
        ],
        colWidths=[9 * cm],
        hAlign="CENTER",
    )
    firma_inner.setStyle(TableStyle([
        ("ALIGN",  (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(firma_inner)

    doc.build(story)
    log.info(f"PDF cotización generado: {pdf_path.name}")
    return pdf_path


# ──────────────────────────────────────────────────────────────────────────────
#  ENVÍO POR OUTLOOK COM
# ──────────────────────────────────────────────────────────────────────────────
def enviar_cotizacion_outlook(message_com, pdf_path: Path, asunto_original: str) -> bool:
    """
    Responde al correo original adjuntando el PDF de cotización.
    En MODO_PRUEBA redirige a CORREO_PRUEBA en lugar del remitente real.
    """
    try:
        reply = message_com.Reply()

        if MODO_PRUEBA:
            destinatario = CORREO_PRUEBA
        else:
            raw = message_com.SenderEmailAddress
            if raw.upper().startswith("/O="):
                try:
                    PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                    destinatario = message_com.Sender.PropertyAccessor.GetProperty(PR_SMTP)
                except Exception:
                    destinatario = message_com.Sender.GetExchangeUser().PrimarySmtpAddress
            else:
                destinatario = raw
        reply.To      = destinatario
        reply.Subject  = f"RE: {asunto_original} — Cotización ADEMEX"
        reply.HTMLBody = CUERPO_CORREO_HTML

        reply.Attachments.Add(str(pdf_path.resolve()))
        reply.Send()

        modo_txt = f"[PRUEBA → {destinatario}]" if MODO_PRUEBA else f"[→ {destinatario}]"
        log.info(f"Cotización enviada {modo_txt}  |  Archivo: {pdf_path.name}")
        return True

    except Exception as e:
        log.error(f"Error al enviar cotización por Outlook: {e}", exc_info=True)
        return False


# ──────────────────────────────────────────────────────────────────────────────
#  FUNCIÓN PRINCIPAL  —  llamada desde main.py
# ──────────────────────────────────────────────────────────────────────────────
def _guardar_sidecar(pdf_path: Path, df_partidas: pd.DataFrame, message_com,
                     asunto: str, fecha_correo, enviado: bool, error: str = None):
    def _col(df, nombre):
        for c in df.columns:
            if c.upper() == nombre:
                return str(df[c].iloc[0]) if len(df) else "—"
        return "—"

    try:
        PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        remitente_real = message_com.Sender.PropertyAccessor.GetProperty(PR_SMTP)
    except Exception:
        try:
            raw = message_com.SenderEmailAddress
            if raw.upper().startswith("/O="):
                try:
                    remitente_real = message_com.Sender.GetExchangeUser().PrimarySmtpAddress
                except Exception:
                    remitente_real = _col(df_partidas, "REMITENTE")
            else:
                remitente_real = raw
        except Exception:
            remitente_real = _col(df_partidas, "REMITENTE")

    destinatario = CORREO_PRUEBA if MODO_PRUEBA else remitente_real

    meta = {
        "expediente":          _col(df_partidas, "EXPEDIENTE"),
        "remitente":           remitente_real,
        "asunto":              asunto,
        "num_partidas":        len(df_partidas),
        "fecha_correo":        fecha_correo.isoformat() if isinstance(fecha_correo, datetime) else str(fecha_correo or ""),
        "fecha_generacion":    datetime.now().isoformat(),
        "estado_envio":        "exitoso" if enviado else "fallido",
        "destinatario_enviado": destinatario,
        "modo_prueba":         MODO_PRUEBA,
        "error":               error,
    }
    try:
        with open(pdf_path.with_suffix(".json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning(f"No se pudo guardar sidecar JSON: {e}")


def generar_y_enviar_cotizacion(message_com, df_partidas: pd.DataFrame,
                                asunto: str, fecha_correo=None) -> bool:
    """
    Punto de entrada: genera el PDF y envía la respuesta.
    Retorna True si ambas operaciones fueron exitosas.
    """
    pdf_path = None
    try:
        pdf_path = generar_pdf(df_partidas, asunto, fecha_correo)
        enviado  = enviar_cotizacion_outlook(message_com, pdf_path, asunto)
        _guardar_sidecar(pdf_path, df_partidas, message_com, asunto, fecha_correo, enviado)
        return enviado
    except Exception as e:
        log.error(f"Error en generar_y_enviar_cotizacion: {e}", exc_info=True)
        if pdf_path and pdf_path.exists():
            _guardar_sidecar(pdf_path, df_partidas, message_com, asunto,
                             fecha_correo, False, str(e))
        return False

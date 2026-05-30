"""
test_cotizacion.py
Prueba independiente del sistema de cotización.
Genera un PDF con claves simuladas y lo envía a CORREO_PRUEBA por Outlook.
No requiere esperar un correo entrante — ejecutar directamente.
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import win32com.client

sys.path.insert(0, str(Path(__file__).parent))
from generar_cotizacion import (
    generar_pdf, _guardar_sidecar, CORREO_PRUEBA, CUERPO_CORREO_HTML
)

# ── Datos simulados (como llegarían de main.py tras procesar un correo) ───────
EXPEDIENTE_PRUEBA = "I-2026-TEST-0001"
REMITENTE_PRUEBA  = CORREO_PRUEBA

DF_PRUEBA = pd.DataFrame([
    {"EXPEDIENTE": EXPEDIENTE_PRUEBA, "PARTIDA":1, "GPO":"010","GEN":"000","ESP":"2616","DIF":"00","VAR":"00","CANTIDAD":18, "CLAVE":"010.000.2616.00.00","REMITENTE":REMITENTE_PRUEBA},
    {"EXPEDIENTE": EXPEDIENTE_PRUEBA, "PARTIDA":2, "GPO":"010","GEN":"000","ESP":"3345","DIF":"00","VAR":"00","CANTIDAD":30, "CLAVE":"010.000.3345.00.00","REMITENTE":REMITENTE_PRUEBA},
    {"EXPEDIENTE": EXPEDIENTE_PRUEBA, "PARTIDA":3, "GPO":"020","GEN":"001","ESP":"1122","DIF":"00","VAR":"00","CANTIDAD":12, "CLAVE":"020.001.1122.00.00","REMITENTE":REMITENTE_PRUEBA},
    {"EXPEDIENTE": EXPEDIENTE_PRUEBA, "PARTIDA":4, "GPO":"030","GEN":"002","ESP":"4455","DIF":"01","VAR":"00","CANTIDAD":5,  "CLAVE":"030.002.4455.01.00","REMITENTE":REMITENTE_PRUEBA},
])

ASUNTO_PRUEBA = "SOLICITUD DE COTIZACION REQ.MED. 26-05-26-TEST"


class _FakeMessage:
    """Simula el objeto COM de Outlook para el sidecar."""
    SenderEmailAddress = REMITENTE_PRUEBA


def main():
    print("=" * 55)
    print("  PRUEBA DE COTIZACIÓN AUTOMÁTICA ADEMEX")
    print("=" * 55)

    # 1. Generar PDF
    print("\n[1/3] Generando PDF de cotización...")
    fecha_prueba = datetime.now()
    pdf_path = generar_pdf(DF_PRUEBA, ASUNTO_PRUEBA, fecha_prueba)
    print(f"      Archivo: {pdf_path.name}")
    print(f"      Ruta:    {pdf_path.parent}")

    # 2. Enviar por Outlook
    print(f"\n[2/3] Enviando a {CORREO_PRUEBA} por Outlook...")
    enviado = False
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail         = outlook.CreateItem(0)
        mail.To      = CORREO_PRUEBA
        mail.Subject = f"[PRUEBA] Cotización ADEMEX — {fecha_prueba.strftime('%d/%m/%Y %H:%M')}"
        mail.HTMLBody = CUERPO_CORREO_HTML
        mail.Attachments.Add(str(pdf_path.resolve()))
        mail.Send()
        enviado = True
        print(f"      OK - Correo enviado correctamente")
    except Exception as e:
        print(f"      ERROR al enviar: {e}")

    # 3. Guardar sidecar JSON (para que aparezca completo en el dashboard)
    print(f"\n[3/3] Guardando metadatos en dashboard...")
    _guardar_sidecar(pdf_path, DF_PRUEBA, _FakeMessage(), ASUNTO_PRUEBA, fecha_prueba, enviado)
    print(f"      OK - Metadatos guardados")

    print("\n" + "=" * 55)
    print("  PRUEBA COMPLETADA")
    print(f"  Revisa la bandeja de {CORREO_PRUEBA}")
    print("=" * 55)


if __name__ == "__main__":
    main()

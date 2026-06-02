import re
import argparse
import unicodedata
import logging
import json
import time
import hashlib
import tempfile
import win32com.client
import pandas as pd

from io import StringIO
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pdfplumber
    PDFPLUMBER_DISPONIBLE = True
except ImportError:
    PDFPLUMBER_DISPONIBLE = False

try:
    import mysql.connector
    MYSQL_DISPONIBLE = True
except ImportError:
    MYSQL_DISPONIBLE = False

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from generar_cotizacion import generar_y_enviar_cotizacion
    COTIZACION_DISPONIBLE = True
except ImportError:
    COTIZACION_DISPONIBLE = False

from os import getenv


# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────

BUZONES_AUTORIZADOS = [
    "ventas@serviciosademex.com",
    "ventas@gadmar.com.mx",
    "contacto.ademex@serviciosademex.com",
    "contacto@gadmar.com.mx",
]

COLUMNAS_FINALES = [
    "EXPEDIENTE", "PARTIDA",
    "GPO", "GEN", "ESP", "DIF", "VAR",
    "CANTIDAD", "CLAVE",
    "ASUNTO", "REMITENTE", "FECHA_CORREO",
]

INTERVALO_MONITOREO = int(getenv("INTERVALO_MONITOREO", 120))
# Ventana (segundos) para considerar dos correos como la misma cotización.
# Se mide contra la FECHA DE RECEPCIÓN de los correos, no el reloj: dos correos
# del mismo remitente con las mismas claves+cantidades recibidos dentro de esta
# ventana se tratan como duplicados (mismo requerimiento llegado a varios buzones,
# o eco reciente del hilo). Pedidos genuinos del mismo contenido separados por más
# de esta ventana se cotizan por separado.
# 600s = 10 min: el remitente IMSS suele enviar el mismo correo a varios buzones
# (ventas@ y contacto@) con hasta ~10 min de diferencia.
VENTANA_DEDUP_SEGUNDOS = int(getenv("VENTANA_DEDUP_SEGUNDOS", 600))
ESTADO_FILE = Path("estado_procesamiento.json")
LOG_FILE    = "cotizaciones.log"

DB_CONFIG = {
    "host":     getenv("MYSQL_HOST",     "localhost"),
    "port":     int(getenv("MYSQL_PORT", 3306)),
    "user":     getenv("MYSQL_USER",     "root"),
    "password": getenv("MYSQL_PASSWORD", ""),
    "database": getenv("MYSQL_DATABASE", "cotizaciones_imss"),
}


# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
#  ESTADO
# ─────────────────────────────────────────────

def cargar_estado() -> dict:
    if ESTADO_FILE.exists():
        try:
            with open(ESTADO_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            log.warning(f"No se pudo leer estado previo: {e}")
    return {"ids_procesados": [], "ultima_fecha": None, "firmas_enviadas": {}}


def guardar_estado(estado: dict):
    try:
        with open(ESTADO_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning(f"No se pudo guardar estado: {e}")


# ─────────────────────────────────────────────
#  EXPEDIENTE Y PARTIDA
# ─────────────────────────────────────────────

# Contador en memoria para fallback sin MySQL (se reinicia al reiniciar el proceso)
_fallback_consecutivo: dict = {}


def _formato_expediente(fecha_hoy, consecutivo: int) -> str:
    dd  = str(fecha_hoy.day).zfill(2)
    mm  = str(fecha_hoy.month).zfill(2)
    yy  = str(fecha_hoy.year)[-2:]
    año = str(fecha_hoy.year)
    return f"I-{año}-{dd}{mm}{yy}{consecutivo}"


def generar_expediente(conn, fecha: datetime) -> str:
    fecha_hoy = fecha.date() if hasattr(fecha, "date") else datetime.now().date()

    if conn is None or not conn.is_connected():
        # Fallback: contador en memoria, único dentro del proceso y con formato correcto
        _fallback_consecutivo[fecha_hoy] = _fallback_consecutivo.get(fecha_hoy, 0) + 1
        return _formato_expediente(fecha_hoy, _fallback_consecutivo[fecha_hoy])

    cursor = conn.cursor()
    try:
        conn.start_transaction()
        cursor.execute(
            "SELECT consecutivo FROM expediente_consecutivo WHERE fecha = %s FOR UPDATE",
            (fecha_hoy,)
        )
        row = cursor.fetchone()
        if row:
            consecutivo = row[0]
            cursor.execute(
                "UPDATE expediente_consecutivo SET consecutivo = consecutivo + 1 WHERE fecha = %s",
                (fecha_hoy,)
            )
        else:
            consecutivo = 1
            cursor.execute(
                "INSERT INTO expediente_consecutivo (fecha, consecutivo) VALUES (%s, %s)",
                (fecha_hoy, 2)
            )
        conn.commit()
    except Exception as e:
        log.warning(f"Error generando expediente: {e}")
        try:
            conn.rollback()
        except Exception:
            pass
        consecutivo = 1
    finally:
        cursor.close()

    return _formato_expediente(fecha_hoy, consecutivo)


def revertir_expediente(conn, fecha: datetime) -> None:
    """Decrementa el consecutivo cuando un expediente fue asignado pero el correo falló."""
    if conn is None:
        return
    fecha_hoy = fecha.date() if hasattr(fecha, "date") else datetime.now().date()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "UPDATE expediente_consecutivo SET consecutivo = consecutivo - 1 "
            "WHERE fecha = %s AND consecutivo > 1",
            (fecha_hoy,)
        )
        conn.commit()
    except Exception as e:
        log.warning(f"revertir_expediente falló: {e}")
    finally:
        cursor.close()


def generar_clave(gpo, gen, esp, dif, var) -> str:
    return f"{gpo}.{gen}.{esp}.{dif}.{var}"


def firma_cotizacion(remitente_real: str, df: pd.DataFrame) -> str:
    """Hash estable de (remitente + conjunto de claves+cantidades) para deduplicar
    cotizaciones idénticas que llegan a varios buzones o como eco reciente del hilo."""
    items = sorted(
        f"{r['GPO']}.{r['GEN']}.{r['ESP']}.{r['DIF']}.{r['VAR']}={r['CANTIDAD']}"
        for _, r in df.iterrows()
    )
    base = (remitente_real or "") + "|" + "|".join(items)
    return hashlib.md5(base.encode()).hexdigest()


def _firma_vigente(ts_iso: str, ahora: datetime) -> bool:
    """True si la firma (por fecha de recepción) sigue dentro de la ventana de dedup."""
    try:
        return (ahora - datetime.fromisoformat(ts_iso)).total_seconds() < VENTANA_DEDUP_SEGUNDOS
    except (ValueError, TypeError):
        return False


# ─────────────────────────────────────────────
#  MYSQL
# ─────────────────────────────────────────────

def conectar_mysql():
    if not MYSQL_DISPONIBLE:
        log.warning("mysql-connector-python no está instalado. Se omitirá MySQL.")
        return None
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        log.info("Conexión a MySQL exitosa.")
        return conn
    except Exception as e:
        log.warning(f"No se pudo conectar a MySQL: {e}. Se guardará solo CSV.")
        return None


def asegurar_tablas_mysql(conn):
    """Crea las tablas necesarias si aún no existen y agrega columnas nuevas."""
    if conn is None:
        return
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS correos_procesados (
                id                INT AUTO_INCREMENT PRIMARY KEY,
                outlook_entry_id  VARCHAR(512) UNIQUE,
                asunto            VARCHAR(500),
                remitente         VARCHAR(255),
                remitente_imss    VARCHAR(255),
                fecha_correo      DATETIME,
                estado            VARCHAR(50)  DEFAULT 'procesando',
                filas_detectadas  INT          DEFAULT 0,
                filas_insertadas  INT          DEFAULT 0,
                expediente        VARCHAR(100),
                intentos          INT          DEFAULT 1,
                error_mensaje     TEXT,
                actualizado_en    DATETIME     DEFAULT CURRENT_TIMESTAMP
                                               ON UPDATE CURRENT_TIMESTAMP,
                creado_en         DATETIME     DEFAULT CURRENT_TIMESTAMP
            ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
        """)
        try:
            cursor.execute(
                "ALTER TABLE correos_procesados ADD COLUMN remitente_imss VARCHAR(255)"
            )
            log.info("Columna remitente_imss agregada a correos_procesados.")
        except Exception:
            pass  # ya existe
        # Asegurar que `estado` acepte cualquier etiqueta (incl. 'duplicado').
        # Si la tabla se creó con un ENUM antiguo, esto lo convierte a VARCHAR.
        try:
            cursor.execute(
                "ALTER TABLE correos_procesados "
                "MODIFY COLUMN estado VARCHAR(50) DEFAULT 'procesando'"
            )
        except Exception:
            pass
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS cotizaciones (
                id           INT AUTO_INCREMENT PRIMARY KEY,
                expediente   VARCHAR(100),
                partida      INT,
                gpo          VARCHAR(3),
                gen          VARCHAR(3),
                esp          VARCHAR(4),
                dif          VARCHAR(2),
                var          VARCHAR(2),
                cantidad     INT,
                clave        VARCHAR(20),
                asunto       VARCHAR(500),
                remitente    VARCHAR(255),
                fecha_correo DATETIME,
                entry_id     VARCHAR(32) UNIQUE,
                creado_en    DATETIME DEFAULT CURRENT_TIMESTAMP
            ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS expediente_consecutivo (
                fecha        DATE PRIMARY KEY,
                consecutivo  INT NOT NULL DEFAULT 1
            ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
        """)
        conn.commit()
        log.info("Tablas MySQL verificadas.")
    except Exception as e:
        log.warning(f"asegurar_tablas_mysql: {e}")
    finally:
        cursor.close()


def asegurar_mysql(conn):
    if conn is None:
        return conectar_mysql()
    try:
        conn.ping(reconnect=True, attempts=3, delay=2)
        return conn
    except Exception:
        log.warning("MySQL ping falló, reconectando desde cero...")
        return conectar_mysql()


def limpiar_mysql_por_fechas(conn, fecha_desde: datetime, fecha_hasta: datetime):
    if conn is None:
        return
    cursor = conn.cursor()
    try:
        cursor.execute(
            "DELETE FROM cotizaciones WHERE fecha_correo BETWEEN %s AND %s",
            (fecha_desde.strftime("%Y-%m-%d 00:00:00"),
             fecha_hasta.strftime("%Y-%m-%d 23:59:59")),
        )
        log.info(f"MySQL cotizaciones eliminadas: {cursor.rowcount} registro(s)")
        cursor.execute(
            "DELETE FROM correos_procesados WHERE fecha_correo BETWEEN %s AND %s",
            (fecha_desde.strftime("%Y-%m-%d 00:00:00"),
             fecha_hasta.strftime("%Y-%m-%d 23:59:59")),
        )
        log.info(f"MySQL correos_procesados eliminados: {cursor.rowcount} registro(s)")
        dia = fecha_desde.date()
        while dia <= fecha_hasta.date():
            cursor.execute("DELETE FROM expediente_consecutivo WHERE fecha = %s", (dia,))
            dia += timedelta(days=1)
        conn.commit()
    except Exception as e:
        log.error(f"Error limpiando MySQL por fechas: {e}")
        try:
            conn.rollback()
        except Exception:
            pass
    finally:
        cursor.close()


def insertar_filas(conn, df: pd.DataFrame) -> int:
    SQL = """
        INSERT INTO cotizaciones
            (expediente, partida, gpo, gen, esp, dif, var,
             cantidad, clave, asunto, remitente, fecha_correo, entry_id)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE id = id
    """
    cursor     = conn.cursor()
    insertadas = 0

    for _, row in df.iterrows():
        hash_fila = hashlib.md5(
            f"{row['EXPEDIENTE']}{row['REMITENTE']}{row['ASUNTO']}{row['GPO']}{row['GEN']}"
            f"{row['ESP']}{row['DIF']}{row['VAR']}{row['CANTIDAD']}".encode()
        ).hexdigest()

        try:
            fecha_mysql = pd.to_datetime(row["FECHA_CORREO"]).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            fecha_mysql = None

        cantidad = int(row["CANTIDAD"]) if str(row["CANTIDAD"]).isdigit() else None

        try:
            cursor.execute(SQL, (
                row["EXPEDIENTE"], int(row["PARTIDA"]),
                row["GPO"], row["GEN"], row["ESP"],
                row["DIF"], row["VAR"], cantidad,
                row["CLAVE"], str(row["ASUNTO"])[:500],
                row["REMITENTE"], fecha_mysql, hash_fila,
            ))
            if cursor.rowcount > 0:
                insertadas += 1
            else:
                log.debug(f"Fila duplicada (ya existía en BD): hash={hash_fila[:8]} clave={row['CLAVE']}")
        except Exception as e:
            log.warning(f"Error en fila {hash_fila[:8]}: {e}")

    conn.commit()
    cursor.close()
    return insertadas


# ─────────────────────────────────────────────
#  AUDITORÍA DE CORREOS
# ─────────────────────────────────────────────

def _fecha_a_mysql(fecha) -> str | None:
    if fecha is None:
        return None
    try:
        dt = fecha.replace(tzinfo=None) if getattr(fecha, "tzinfo", None) else fecha
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def registrar_correo_inicio(conn, outlook_entry_id: str, asunto: str,
                             remitente: str, fecha_correo) -> None:
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        cursor.execute(
            """INSERT INTO correos_procesados
                   (outlook_entry_id, asunto, remitente, fecha_correo, estado, intentos)
               VALUES (%s, %s, %s, %s, 'procesando', 1)
               ON DUPLICATE KEY UPDATE
                   intentos = intentos + 1,
                   estado   = 'procesando',
                   actualizado_en = CURRENT_TIMESTAMP
            """,
            (outlook_entry_id, str(asunto)[:500], str(remitente)[:255],
             _fecha_a_mysql(fecha_correo)),
        )
        conn.commit()
        cursor.close()
    except Exception as e:
        log.warning(f"registrar_correo_inicio falló: {e}")


def actualizar_correo_estado(conn, outlook_entry_id: str, estado: str,
                              filas_detectadas: int = 0, filas_insertadas: int = 0,
                              expediente: str = None, error: str = None,
                              remitente_imss: str = None) -> None:
    if conn is None or not outlook_entry_id:
        return
    try:
        cursor = conn.cursor()
        cursor.execute(
            """UPDATE correos_procesados
               SET estado           = %s,
                   filas_detectadas = %s,
                   filas_insertadas = %s,
                   expediente       = %s,
                   error_mensaje    = %s,
                   remitente_imss   = COALESCE(%s, remitente_imss),
                   actualizado_en   = CURRENT_TIMESTAMP
               WHERE outlook_entry_id = %s
            """,
            (estado, filas_detectadas, filas_insertadas, expediente, error,
             remitente_imss or None, outlook_entry_id),
        )
        conn.commit()
        cursor.close()
    except Exception as e:
        log.warning(f"actualizar_correo_estado falló ({estado}): {e}")


# ─────────────────────────────────────────────
#  UTILIDADES DE TEXTO
# ─────────────────────────────────────────────

def limpiar_texto(texto: str) -> str:
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_valor(valor) -> str:
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.lower() in ["nan", "none"]:
        return ""
    return valor


# ─────────────────────────────────────────────
#  VALIDACIÓN DE CORREOS
# ─────────────────────────────────────────────

_PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


def _smtp_desde_sender(message) -> str:
    try:
        return message.Sender.PropertyAccessor.GetProperty(_PR_SMTP)
    except Exception:
        pass
    try:
        raw = str(message.SenderEmailAddress)
        if not raw.upper().startswith("/O="):
            return raw
        try:
            return message.Sender.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            pass
    except Exception:
        pass
    return ""


def obtener_remitente_real(message) -> str:
    addr = _smtp_desde_sender(message)
    return addr if addr else ""


def extraer_remitente_original(message) -> str:
    try:
        body = str(message.Body)
    except Exception:
        return obtener_remitente_real(message)

    for match in re.finditer(
        r'(?:De|From):\s*[^<\n]*<([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})>',
        body, re.IGNORECASE,
    ):
        if "imss.gob.mx" in match.group(1).lower():
            return match.group(1).lower()

    for match in re.finditer(
        r'(?:De|From):\s*([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        body, re.IGNORECASE,
    ):
        if "imss.gob.mx" in match.group(1).lower():
            return match.group(1).lower()

    match = re.search(r'([a-zA-Z0-9._%+\-]+@imss\.gob\.mx)', body, re.IGNORECASE)
    if match:
        return match.group(1).lower()

    return obtener_remitente_real(message)


def viene_de_buzon_autorizado(remitente_limpio: str) -> bool:
    return any(b in remitente_limpio for b in BUZONES_AUTORIZADOS)


def contiene_variantes_cotizacion(texto: str) -> bool:
    palabras = [
        "cot", "coti", "cotiz", "cotizar", "cotizacion",
        "compra emergente", "solicitud",
    ]
    return any(p in texto for p in palabras)


def es_correo_valido(asunto: str, cuerpo: str, remitente: str) -> bool:
    remitente_limpio = limpiar_texto(remitente)
    texto_completo   = limpiar_texto(asunto + " " + cuerpo)

    if "se le ha asignado una tarea" in texto_completo:
        return False
    if not viene_de_buzon_autorizado(remitente_limpio):
        return False
    if "@imss.gob.mx" not in texto_completo:
        return False
    if not contiene_variantes_cotizacion(texto_completo):
        return False
    return True


# ─────────────────────────────────────────────
#  NORMALIZACIÓN DE COLUMNAS
# ─────────────────────────────────────────────

EQUIVALENCIAS_COLUMNAS = {
    "gpo": "GPO", "grupo": "GPO",
    "gen": "GEN", "generico": "GEN",
    "esp": "ESP", "especifico": "ESP",
    "dif": "DIF", "diferenciador": "DIF",
    "var": "VAR", "variante": "VAR",
    "cantidad": "CANTIDAD", "piezas": "CANTIDAD",
    "pieza": "CANTIDAD", "necesidad": "CANTIDAD",
    "solicitado": "CANTIDAD", "cantidad_solicitada": "CANTIDAD",
    "cantidad_autorizada": "CANTIDAD", "cantidad_requerida": "CANTIDAD",
    "requerida": "CANTIDAD", "cant_requerida": "CANTIDAD",
    "solicit": "CANTIDAD", "solicitud": "CANTIDAD",
    "cant_solicitada": "CANTIDAD", "cant_solicit": "CANTIDAD",
    "total": "CANTIDAD",
    # "unidad" removido: en tablas IMSS UNIDAD es código de hospital, no cantidad
    "ingreso": "CANTIDAD",   # columna INGRESO usada como cantidad en formato FOLIO/UNIDAD
    "medicamento_faltante": "CANTIDAD",
    "faltante": "CANTIDAD",
}


def normalizar_columna(col: str) -> str:
    col = limpiar_texto(col).replace(" ", "_")
    return EQUIVALENCIAS_COLUMNAS.get(col, col.upper())


# ─────────────────────────────────────────────
#  DETECCIÓN Y LIMPIEZA DE TABLAS
# ─────────────────────────────────────────────

def detectar_fila_encabezado(df: pd.DataFrame):
    for idx, row in df.iterrows():
        texto = " ".join(limpiar_texto(str(v)) for v in row.values)
        tiene_gpo  = any(p in texto for p in ["gpo", "grupo"])
        tiene_esp  = any(p in texto for p in ["esp", "especifico"])
        tiene_cant = any(p in texto for p in [
            "cantidad", "piezas", "necesidad", "solicitado",
            "solicit", "solicitud",
            "total", "unidad", "faltante", "medicamento_faltante",
            "ingreso",
        ])
        if tiene_gpo and tiene_esp and tiene_cant:
            return idx
    return None


def _asignar_dif_var_posicional(df: pd.DataFrame) -> pd.DataFrame:
    sin_nombre = [c for c in df.columns
                  if str(c).strip() in ("", "NAN", "NONE")
                  or str(c).upper().startswith("UNNAMED")]
    falta_dif = "DIF" not in df.columns
    falta_var = "VAR" not in df.columns
    rename = {}
    if falta_dif and len(sin_nombre) >= 1:
        rename[sin_nombre[0]] = "DIF"
    if falta_var and len(sin_nombre) >= 2:
        rename[sin_nombre[1]] = "VAR"
    elif falta_var and len(sin_nombre) == 1 and not falta_dif:
        rename[sin_nombre[0]] = "VAR"
    return df.rename(columns=rename) if rename else df


def filtrar_formato_clave(df: pd.DataFrame) -> pd.DataFrame:
    m_gpo = df["GPO"].apply(lambda x: bool(re.fullmatch(r"\d{3}", str(x).strip())))
    m_gen = df["GEN"].apply(lambda x: bool(re.fullmatch(r"\d{3}", str(x).strip())))
    m_esp = df["ESP"].apply(lambda x: bool(re.fullmatch(r"\d{4}", str(x).strip())))
    m_dif = df["DIF"].apply(lambda x: bool(re.fullmatch(r"\d{2}", str(x).strip())))
    m_var = df["VAR"].apply(lambda x: bool(re.fullmatch(r"\d{2}", str(x).strip())))
    mascara = m_gpo & m_gen & m_esp & m_dif & m_var
    descartadas = (~mascara).sum()
    if descartadas:
        cols_fallo = [c for c, m in [("GPO",m_gpo),("GEN",m_gen),("ESP",m_esp),("DIF",m_dif),("VAR",m_var)] if not m.all()]
        vals_muestra = df.loc[~mascara, ["GPO","GEN","ESP","DIF","VAR"]].head(3).to_dict("records")
        log.warning(f"filtrar_formato_clave: {descartadas} fila(s) descartada(s). Columnas con error: {cols_fallo}. Muestra: {vals_muestra}")
    return df[mascara].reset_index(drop=True)


def limpiar_tabla_con_encabezado(df_original: pd.DataFrame):
    df = df_original.copy().dropna(how="all").dropna(axis=1, how="all")
    if df.empty:
        return None

    cols_norm = [normalizar_columna(str(c)) for c in df.columns]
    if "GPO" in cols_norm and "ESP" in cols_norm:
        df.columns = cols_norm
        # Validar que la columna GPO realmente contenga claves de 3 dígitos.
        # En algunos formatos IMSS (FOLIO/UNIDAD) los encabezados vienen
        # corridos respecto a los datos: la columna etiquetada "GPO" contiene
        # el código de unidad concatenado y NO el grupo real. En ese caso se
        # ignora el mapeo por nombre y se delega al modo sin encabezado.
        try:
            serie_gpo = df.loc[:, df.columns == "GPO"].iloc[:, 0].map(limpiar_valor)
            proporcion_valida = serie_gpo.apply(
                lambda x: bool(re.fullmatch(r"\d{3}", str(x).strip()))
            ).mean()
        except Exception:
            proporcion_valida = 0.0
        if proporcion_valida < 0.5:
            log.debug(
                "Encabezados corridos detectados (GPO no es clave de 3 dígitos); "
                "se delega al modo sin encabezado."
            )
            return None
    else:
        fila_enc = detectar_fila_encabezado(df)
        if fila_enc is None:
            return None
        encabezados = df.loc[fila_enc].tolist()
        df = df.loc[fila_enc + 1:].copy()
        df.columns = [normalizar_columna(str(c)) for c in encabezados]

    df = df.dropna(how="all")
    try:
        df = df.map(limpiar_valor)
    except AttributeError:
        df = df.applymap(limpiar_valor)

    # Coalesce duplicate CANTIDAD columns (INGRESO y SOLICITADO ambos mapean a CANTIDAD)
    if (df.columns == "CANTIDAD").sum() > 1:
        cant_merged = df.loc[:, df.columns == "CANTIDAD"].apply(
            lambda row: next(
                (str(v).strip() for v in row if str(v).strip() not in ("", "nan", "None")),
                ""
            ),
            axis=1,
        )
        df = df.loc[:, df.columns != "CANTIDAD"].copy()
        df["CANTIDAD"] = cant_merged.values

    df = df.loc[:, ~df.columns.duplicated()]
    df = _asignar_dif_var_posicional(df)

    for col in ["GPO", "GEN", "ESP", "DIF", "VAR", "CANTIDAD"]:
        if col not in df.columns:
            df[col] = ""

    df = df[["GPO", "GEN", "ESP", "DIF", "VAR", "CANTIDAD"]]

    # Reparar antes del filtro de vacíos: columna GEN con 4 dígitos y ESP ausente
    # significa que la tabla omitió la columna ESP y puso el específico en GEN.
    _gen4_pre = df["GEN"].astype(str).str.strip().apply(lambda x: bool(re.fullmatch(r"\d{4}", x)))
    _esp_vacia = df["ESP"].astype(str).str.strip().eq("")
    _fix_pre = _gen4_pre & _esp_vacia
    if _fix_pre.any():
        df.loc[_fix_pre, "ESP"] = df.loc[_fix_pre, "GEN"]
        df.loc[_fix_pre, "GEN"] = "0"
        log.info(f"GEN→ESP reparado (columna GEN contenía específico de 4 dígitos): {_fix_pre.sum()} fila(s).")

    df = df[df["GPO"].astype(str).str.strip().ne("") & df["ESP"].astype(str).str.strip().ne("")]
    df = df[df["GPO"].apply(lambda x: limpiar_texto(str(x)) != "gpo")]

    def _norm_int(v):
        s = str(v).strip()
        try:
            return str(int(float(s))) if s else s
        except (ValueError, OverflowError):
            return s

    df["GPO"] = df["GPO"].apply(_norm_int).str.zfill(3)
    df["GEN"] = df["GEN"].apply(_norm_int).str.zfill(3)
    df["ESP"] = df["ESP"].apply(_norm_int).str.zfill(4)
    df["DIF"] = df["DIF"].apply(_norm_int).str.zfill(2)
    df["VAR"] = df["VAR"].apply(_norm_int).str.zfill(2)

    df = filtrar_formato_clave(df)

    def _extraer_cantidad(v: str) -> str:
        # Ignorar secuencias de 9+ dígitos: probablemente claves concatenadas, no cantidades
        validas = [s for s in re.findall(r"\d+", str(v)) if len(s) <= 8]
        return max(validas, key=len, default="")

    df["CANTIDAD"] = df["CANTIDAD"].astype(str).apply(_extraer_cantidad)
    df = df[df["CANTIDAD"].str.strip().ne("")]
    return df if not df.empty else None


def es_clave_3(v):  return bool(re.fullmatch(r"\d{3}", limpiar_valor(v)))
def es_clave_4(v):  return bool(re.fullmatch(r"\d{4}", limpiar_valor(v)))
def es_clave_12(v): return bool(re.fullmatch(r"\d{1,2}", limpiar_valor(v)))

_RE_FECHA = re.compile(r"\b\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}\b")
_RE_AÑO  = re.compile(r"^\d{4}$")

def extraer_numero(v: str) -> str:
    s = limpiar_valor(v).strip()
    # Ignorar fechas completas (dd/mm/aaaa, dd-mm-aa, etc.)
    if _RE_FECHA.search(s):
        return ""
    # Ignorar años sueltos (ej. "2026") que aparecen en columnas FECHA SOLICITUD
    if _RE_AÑO.match(s) and 1900 <= int(s) <= 2100:
        return ""
    m = re.search(r"\d+", s)
    if not m:
        return ""
    # Ignorar secuencias de 9+ dígitos: probablemente clave concatenada, no cantidad
    return m.group(0) if len(m.group(0)) <= 8 else ""


def limpiar_tabla_sin_encabezado(df_original: pd.DataFrame):
    df = df_original.copy().dropna(how="all").dropna(axis=1, how="all")
    if df.empty:
        return None
    try:
        df = df.map(limpiar_valor)
    except AttributeError:
        df = df.applymap(limpiar_valor)

    filas = []
    for _, row in df.iterrows():
        valores = [v for v in row.values if v != ""]
        if len(valores) < 6:
            continue
        for i in range(len(valores) - 4):
            gpo, gen, esp, dif, var = (
                valores[i], valores[i+1], valores[i+2],
                valores[i+3], valores[i+4],
            )
            gpo_z = str(gpo).strip().zfill(3)
            gen_z = str(gen).strip().zfill(3)
            esp_z = str(esp).strip().zfill(4)
            dif_z = str(dif).strip().zfill(2)
            var_z = str(var).strip().zfill(2)
            if es_clave_3(gpo_z) and es_clave_3(gen_z) and es_clave_4(esp_z) \
               and es_clave_12(dif_z) and es_clave_12(var_z):
                claves_raw = {gpo, gen, esp, dif, var}
                claves_z   = {gpo_z, gen_z, esp_z, dif_z, var_z}
                cantidad = next(
                    (extraer_numero(v) for v in reversed(valores)
                     if extraer_numero(v) and v not in claves_raw and v not in claves_z), ""
                )
                if cantidad:
                    filas.append({
                        "GPO": gpo_z, "GEN": gen_z, "ESP": esp_z,
                        "DIF": dif_z, "VAR": var_z, "CANTIDAD": cantidad,
                    })
                break

    if not filas:
        return None
    df_limpio = pd.DataFrame(filas)
    return df_limpio if not df_limpio.empty else None


def limpiar_tabla_imss(tabla: pd.DataFrame):
    df = limpiar_tabla_con_encabezado(tabla)
    if df is not None and not df.empty:
        return df
    df = limpiar_tabla_sin_encabezado(tabla)
    if df is not None and not df.empty:
        return df
    return None


# ── DETECCIÓN DE HILOS ───────────────────────────────────────────────────────

_SEPARADORES_HILO = [
    r'<div[^>]+id=["\']divRplyFwdMsg["\']',                       # OWA (Outlook Web)
    r'<hr[^>]*style=["\'][^"\']*border-top[^"\']*["\']',          # Outlook desktop HR
    r'<div[^>]*style=["\'][^"\']*border:none;border-top[^"\']*["\']',  # Outlook 365 div
    r'<div[^>]*style=["\'][^"\']*border-top:solid[^"\']*["\']',   # Outlook 365 div alt
    r'<div[^>]+class=["\'][^"\']*gmail_quote[^"\']*["\']',        # Gmail
    r'<blockquote[^>]+type=["\']cite["\']',                       # Lotus/genérico
    r'[-]{5,}\s*(?:Mensaje original|Original Message|Forwarded)',
    # Frase de cierre del cliente IMSS — marca el fin del contenido nuevo del día
    r'sin(?:\s|&nbsp;)+m(?:á|&aacute;|&#225;|&#xE1;)s(?:\s|&nbsp;)+por(?:\s|&nbsp;)+el(?:\s|&nbsp;)+momento',
]


def _html_solo_respuesta_nueva(html: str) -> str:
    """Devuelve solo la parte más reciente de un correo en hilo.

    Busca separadores de respuesta/reenvío y corta antes del primero.
    Si no encuentra separadores, devuelve el HTML completo.
    """
    primera_pos = len(html)
    for patron in _SEPARADORES_HILO:
        m = re.search(patron, html, re.IGNORECASE)
        if m and m.start() < primera_pos:
            primera_pos = m.start()
    porcion = html[:primera_pos].strip()
    return porcion if porcion else html


# ── EXTRACCIÓN DE TABLAS HTML ────────────────────────────────────────────────

_PREFIJOS_REENVIO = ("rv:", "fwd:", "fw:", "reenvio:", "reenvío:", "tr:", "re: rv:", "re: fwd:", "re: re:")


def _es_reenvio(message) -> bool:
    try:
        asunto = limpiar_texto(str(message.Subject or ""))
        return asunto.startswith(_PREFIJOS_REENVIO)
    except Exception:
        return False


def _html_a_tablas_imss(html: str) -> list:
    """Extrae y valida todas las tablas IMSS de un fragmento HTML."""
    try:
        tablas_raw = pd.read_html(StringIO(html), header=None)
    except Exception as e:
        log.debug(f"pd.read_html falló: {e}")
        return []
    tablas_limpias = []
    for i, tabla in enumerate(tablas_raw):
        limpia = limpiar_tabla_imss(tabla)
        if limpia is not None and not limpia.empty:
            log.debug(f"Tabla HTML #{i+1}: {len(limpia)} filas válidas.")
            tablas_limpias.append(limpia)
    return tablas_limpias


def _html_primera_seccion_reenviada(html: str) -> str:
    """
    Devuelve solo el primer bloque reenviado del HTML, es decir, el fragmento
    entre el primer separador de hilo y el segundo (o el final si no hay segundo).
    Evita procesar toda la cadena histórica de un hilo largo.
    """
    posiciones = []
    for patron in _SEPARADORES_HILO:
        for m in re.finditer(patron, html, re.IGNORECASE):
            posiciones.append(m.start())
    posiciones.sort()

    if not posiciones:
        return html
    inicio = posiciones[0]
    fin    = posiciones[1] if len(posiciones) >= 2 else len(html)
    return html[inicio:fin].strip()


def extraer_tablas_html(message) -> list:
    """
    Lee las tablas HTML del correo.
    Prioriza la parte más reciente del hilo. Para reenvíos (RV:/FWD:) sin tabla
    en la parte reciente, intenta solo el primer bloque reenviado (entre el primer
    y segundo separador), evitando procesar toda la cadena histórica.
    """
    try:
        html_completo = message.HTMLBody
    except Exception as e:
        log.debug(f"No se pudo leer HTMLBody: {e}")
        return []

    html_reciente = _html_solo_respuesta_nueva(html_completo)
    hilo_detectado = html_reciente != html_completo

    if hilo_detectado:
        log.debug("Hilo detectado: procesando solo la parte reciente del correo.")

    html_a_parsear = html_reciente if hilo_detectado else html_completo
    tablas_limpias = _html_a_tablas_imss(html_a_parsear)

    # Fallback para reenvíos: buscar la tabla en el primer bloque reenviado
    # (entre separador 1 y separador 2), ignorando el historial del hilo.
    if not tablas_limpias and hilo_detectado and _es_reenvio(message):
        html_reenvio = _html_primera_seccion_reenviada(html_completo)
        log.info("Reenvío sin tabla en parte reciente — intentando primer bloque reenviado.")
        tablas_limpias = _html_a_tablas_imss(html_reenvio)

    if tablas_limpias:
        log.info(f"Tablas HTML válidas encontradas: {len(tablas_limpias)}")
    return tablas_limpias


def extraer_tablas_adjuntos(message) -> list:
    """Lee adjuntos Excel y PDF y devuelve TODAS las tablas IMSS válidas."""
    EXTENSIONES = {".xlsx", ".xls", ".pdf"}
    try:
        count = message.Attachments.Count
    except Exception:
        return []

    tablas_limpias = []

    for i in range(1, count + 1):
        try:
            att    = message.Attachments.Item(i)
            nombre = str(att.FileName)
            ext    = Path(nombre).suffix.lower()

            if ext not in EXTENSIONES:
                log.debug(f"Adjunto ignorado: {nombre}")
                continue

            tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
            tmp.close()
            tmp_path = Path(tmp.name)

            try:
                att.SaveAsFile(str(tmp_path))

                if ext in {".xlsx", ".xls"}:
                    hojas = pd.read_excel(tmp_path, sheet_name=None, header=None)
                    for nombre_hoja, df_hoja in hojas.items():
                        limpia = limpiar_tabla_imss(df_hoja)
                        if limpia is not None and not limpia.empty:
                            log.info(f"FUENTE: adjunto Excel \"{nombre}\" hoja \"{nombre_hoja}\"")
                            tablas_limpias.append(limpia)

                elif ext == ".pdf":
                    if not PDFPLUMBER_DISPONIBLE:
                        log.warning(f"PDF \"{nombre}\" ignorado: instala pdfplumber.")
                        continue
                    with pdfplumber.open(str(tmp_path)) as pdf:
                        for num_pag, pagina in enumerate(pdf.pages, start=1):
                            for tabla_raw in pagina.extract_tables():
                                if not tabla_raw:
                                    continue
                                df_pdf = pd.DataFrame(tabla_raw)
                                limpia = limpiar_tabla_imss(df_pdf)
                                if limpia is not None and not limpia.empty:
                                    log.info(f"FUENTE: PDF \"{nombre}\" página {num_pag}")
                                    tablas_limpias.append(limpia)

            except Exception as e:
                log.warning(f"Error procesando adjunto \"{nombre}\": {e}")
            finally:
                try:
                    tmp_path.unlink()
                except Exception:
                    pass

        except Exception as e:
            log.warning(f"Error accediendo al adjunto #{i}: {e}")

    return tablas_limpias


# ─────────────────────────────────────────────
#  CONEXIÓN A OUTLOOK
# ─────────────────────────────────────────────

def conectar_outlook():
    for intento in range(1, 4):
        try:
            outlook   = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox     = namespace.GetDefaultFolder(6)
            log.info("Conexión a Outlook exitosa.")
            return inbox
        except Exception as e:
            log.warning(f"Intento {intento}/3 falló: {e}")
    raise ConnectionError("No se pudo conectar a Outlook después de 3 intentos.")


# ─────────────────────────────────────────────
#  CICLO DE PROCESAMIENTO
# ─────────────────────────────────────────────

def procesar_correos(inbox, conn, ids_procesados: set, ultima_fecha,
                     fecha_desde: datetime = None, fecha_hasta: datetime = None,
                     reprocesar: bool = False, firmas_enviadas: dict = None):
    conn = asegurar_mysql(conn)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    if firmas_enviadas is None:
        firmas_enviadas = {}

    correos_validos    = 0
    correos_con_tabla  = 0
    total_insertadas   = 0
    correos_duplicados = 0
    nueva_ultima_fecha = ultima_fecha
    resultados_csv     = []

    for message in messages:
        entry_id       = None
        remitente_real = ""
        try:
            entry_id     = str(message.EntryID)
            fecha_correo = message.ReceivedTime
            fecha_correo_cmp = fecha_correo.replace(tzinfo=None) if getattr(fecha_correo, "tzinfo", None) else fecha_correo

            if reprocesar:
                if fecha_hasta and fecha_correo_cmp > fecha_hasta:
                    continue
                if fecha_desde and fecha_correo_cmp < fecha_desde:
                    break
            else:
                if entry_id in ids_procesados:
                    continue
                if ultima_fecha and fecha_correo_cmp <= ultima_fecha:
                    break

            asunto          = str(message.Subject)
            cuerpo          = str(message.Body)
            remitente_buzon = obtener_remitente_real(message)

            if not es_correo_valido(asunto, cuerpo, remitente_buzon):
                continue

            correos_validos += 1
            registrar_correo_inicio(conn, entry_id, asunto, remitente_buzon, fecha_correo)
            remitente_real = extraer_remitente_original(message)

            if nueva_ultima_fecha is None or fecha_correo_cmp > nueva_ultima_fecha:
                nueva_ultima_fecha = fecha_correo_cmp

            log.info("=" * 60)
            log.info(f"CORREO VÁLIDO #{correos_validos}")
            log.info(f"ASUNTO   : {asunto}")
            log.info(f"REMITENTE: {remitente_real} (vía {remitente_buzon})")
            log.info(f"FECHA    : {fecha_correo}")

            # Buscar en cuerpo HTML primero, luego en adjuntos
            tablas_limpias = extraer_tablas_html(message)
            if not tablas_limpias:
                tablas_limpias = extraer_tablas_adjuntos(message)

            if not tablas_limpias:
                log.info("Sin tabla válida (cuerpo ni adjuntos).")
                actualizar_correo_estado(conn, entry_id, 'sin_tabla', remitente_imss=remitente_real)
                ids_procesados.add(entry_id)
                continue

            # Concatenar TODAS las tablas del correo en un solo DataFrame
            filas_brutas = sum(len(t) for t in tablas_limpias)
            df_correo = pd.concat(tablas_limpias, ignore_index=True)
            df_correo = filtrar_formato_clave(df_correo)

            if df_correo.empty:
                log.info("Sin filas válidas tras filtro de formato.")
                actualizar_correo_estado(conn, entry_id, 'sin_tabla', filas_brutas, remitente_imss=remitente_real)
                ids_procesados.add(entry_id)
                continue

            # Deduplicación: ¿ya se cotizó este mismo requerimiento (remitente +
            # claves + cantidades) dentro de la ventana? Cubre el mismo correo
            # llegado a varios buzones y ecos recientes del hilo. Se compara contra
            # la fecha de recepción del correo original, no el reloj.
            firma = firma_cotizacion(remitente_real, df_correo)
            ts_previo = firmas_enviadas.get(firma)
            if ts_previo:
                try:
                    dt_previo = datetime.fromisoformat(ts_previo)
                    delta = abs((fecha_correo_cmp - dt_previo).total_seconds())
                    if delta < VENTANA_DEDUP_SEGUNDOS:
                        correos_duplicados += 1
                        log.info(
                            f"DUPLICADO omitido — mismo remitente/claves/cantidades a "
                            f"{int(delta)}s del original (ventana {VENTANA_DEDUP_SEGUNDOS}s). "
                            f"No se inserta en MySQL ni se envía PDF."
                        )
                        actualizar_correo_estado(conn, entry_id, 'duplicado', filas_brutas,
                                                 remitente_imss=remitente_real)
                        ids_procesados.add(entry_id)
                        continue
                except ValueError:
                    pass

            expediente = generar_expediente(conn, fecha_correo)
            expediente_usado = False
            try:
                # Partida consecutiva sobre el total de filas de todas las tablas
                df_correo["PARTIDA"] = range(1, len(df_correo) + 1)
                df_correo["CLAVE"]   = df_correo.apply(
                    lambda r: generar_clave(r["GPO"], r["GEN"], r["ESP"], r["DIF"], r["VAR"]),
                    axis=1,
                )
                df_correo["EXPEDIENTE"]   = expediente
                df_correo["ASUNTO"]       = asunto
                df_correo["REMITENTE"]    = remitente_real
                df_correo["FECHA_CORREO"] = str(fecha_correo)
                df_correo = df_correo[COLUMNAS_FINALES]

                correos_con_tabla += 1
                resultados_csv.append(df_correo)

                log.info(f"Expediente : {expediente}")
                log.info(f"Filas total: {len(df_correo)}")
                log.info(f"\n{df_correo[['EXPEDIENTE','PARTIDA','GPO','GEN','ESP','DIF','VAR','CANTIDAD','CLAVE']].to_string(index=False)}")

                filas_insertadas_correo = 0
                if conn is not None:
                    try:
                        if not conn.is_connected():
                            conn = conectar_mysql()
                        if conn:
                            filas_insertadas_correo = insertar_filas(conn, df_correo)
                            total_insertadas += filas_insertadas_correo
                            log.info(f"Insertadas en MySQL: {filas_insertadas_correo}")
                    except Exception as e:
                        log.error(f"Error insertando en MySQL: {e}", exc_info=True)

                expediente_usado = True
                # Registrar la firma con la fecha de recepción del correo: los
                # duplicados posteriores dentro de la ventana se omitirán por completo.
                firmas_enviadas[firma] = fecha_correo_cmp.isoformat()
                actualizar_correo_estado(
                    conn, entry_id, 'completado',
                    len(df_correo), filas_insertadas_correo, expediente,
                    remitente_imss=remitente_real,
                )
                ids_procesados.add(entry_id)

                if COTIZACION_DISPONIBLE:
                    try:
                        generar_y_enviar_cotizacion(message, df_correo, asunto, fecha_correo)
                    except Exception as e_cot:
                        log.error(f"Error generando cotización: {e_cot}", exc_info=True)

            except Exception as e_inner:
                if not expediente_usado:
                    revertir_expediente(conn, fecha_correo)
                    log.warning(f"Expediente {expediente} revertido por error post-asignación: {e_inner}")
                raise

        except Exception as e:
            log.error(f"Error procesando correo: {e}", exc_info=True)
            actualizar_correo_estado(conn, entry_id, 'error', error=str(e), remitente_imss=remitente_real)

    if resultados_csv:
        df_final   = pd.concat(resultados_csv, ignore_index=True)
        nombre_csv = f"cotizaciones_imss_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df_final.to_csv(nombre_csv, index=False, encoding="utf-8-sig")
        log.info(f"CSV de respaldo: {nombre_csv} ({len(df_final)} filas)")

    log.info(
        f"Ciclo terminado — válidos: {correos_validos} | "
        f"con tabla: {correos_con_tabla} | duplicados: {correos_duplicados} | "
        f"MySQL: {total_insertadas}"
    )
    return ids_procesados, nueva_ultima_fecha, conn


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Monitor de Cotizaciones IMSS")
    parser.add_argument("--desde", metavar="YYYY-MM-DD", help="Reprocesar desde esta fecha")
    parser.add_argument("--hasta", metavar="YYYY-MM-DD", help="Reprocesar hasta esta fecha")
    args = parser.parse_args()

    fecha_desde = datetime.strptime(args.desde, "%Y-%m-%d") if args.desde else None
    fecha_hasta = (
        datetime.strptime(args.hasta, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
        if args.hasta else None
    )
    reprocesar = fecha_desde is not None or fecha_hasta is not None

    log.info("╔══════════════════════════════════════════════╗")
    log.info("║      Monitor de Cotizaciones IMSS            ║")
    if reprocesar:
        log.info(f"║  MODO REPROCESAR: {args.desde} – {args.hasta}  ║")
    else:
        log.info(f"║      Intervalo: {INTERVALO_MONITOREO}s                          ║")
    log.info("╚══════════════════════════════════════════════╝")

    inbox = conectar_outlook()
    conn  = conectar_mysql()
    asegurar_tablas_mysql(conn)

    if conn is None:
        log.warning("⚠  MySQL no disponible. Se guardará solo CSV.")

    if reprocesar:
        if conn and fecha_desde:
            limpiar_mysql_por_fechas(conn, fecha_desde, fecha_hasta or fecha_desde)
        try:
            procesar_correos(
                inbox, conn, set(), None,
                fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, reprocesar=True,
                firmas_enviadas={},
            )
        finally:
            if conn and MYSQL_DISPONIBLE:
                try:
                    if conn.is_connected():
                        conn.close()
                        log.info("Conexión MySQL cerrada.")
                except Exception:
                    pass
        return

    estado          = cargar_estado()
    ids_procesados  = set(estado.get("ids_procesados", []))
    firmas_enviadas = dict(estado.get("firmas_enviadas", {}))
    ultima_fecha    = None

    if estado.get("ultima_fecha"):
        try:
            ultima_fecha = datetime.fromisoformat(estado["ultima_fecha"])
            log.info(f"Retomando desde: {ultima_fecha}")
        except ValueError:
            log.warning("Fecha en estado inválida, procesando desde el inicio.")

    try:
        while True:
            log.info(f"\n⏱  [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Revisando correos nuevos...")
            ids_procesados, ultima_fecha, conn = procesar_correos(
                inbox, conn, ids_procesados, ultima_fecha,
                firmas_enviadas=firmas_enviadas,
            )
            # Podar firmas cuya fecha de recepción ya salió de la ventana de dedup
            ahora = datetime.now()
            firmas_enviadas = {
                h: ts for h, ts in firmas_enviadas.items()
                if _firma_vigente(ts, ahora)
            }
            guardar_estado({
                "ids_procesados": list(ids_procesados),
                "ultima_fecha": ultima_fecha.isoformat() if ultima_fecha else None,
                "firmas_enviadas": firmas_enviadas,
            })
            log.info(f"Próxima revisión en {INTERVALO_MONITOREO}s. (Ctrl+C para detener)\n")
            time.sleep(INTERVALO_MONITOREO)

    except KeyboardInterrupt:
        log.info("Monitor detenido manualmente.")
    finally:
        if conn and MYSQL_DISPONIBLE:
            try:
                if conn.is_connected():
                    conn.close()
                    log.info("Conexión MySQL cerrada.")
            except Exception:
                pass


if __name__ == "__main__":
    main()
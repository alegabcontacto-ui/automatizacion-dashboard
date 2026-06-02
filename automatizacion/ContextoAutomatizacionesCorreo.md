# Contexto — Automatización de Cotizaciones IMSS (ADEMEX / GADMAR)

## ¿Qué hace el sistema?

Monitor continuo que revisa correos entrantes en Outlook, detecta requerimientos de cotización del IMSS, extrae la tabla de claves, genera un expediente, inserta los datos en MySQL y responde automáticamente con un PDF de cotización.

---

## Archivos principales

| Archivo | Rol |
|---------|-----|
| `main.py` | Monitor principal: Outlook → parseo → MySQL → CSV |
| `generar_cotizacion.py` | Genera PDF (ReportLab) y envía respuesta por Outlook COM |
| `dashboard/dashboard.py` | Dashboard Streamlit — lee **solo de MySQL**, no de CSVs |
| `estado_procesamiento.json` | Persiste `ids_procesados` + `ultima_fecha` entre reinicios |
| `cotizaciones.log` | Log completo de operación |
| `cotizaciones_generadas/` | PDFs generados + sidecars JSON por cotización |

> Los archivos `cotizaciones_imss_YYYYMMDD_HHMMSS.csv` son **solo respaldo local** generados por `main.py`. El dashboard NO los usa — se pueden borrar sin impacto.

---

## Flujo de procesamiento

```
Outlook inbox
    │
    ▼
¿Correo válido? (buzon autorizado + @imss.gob.mx + variante "cotiz")
    │ sí
    ▼
Extraer tabla HTML → fallback adjunto Excel/PDF
    │
    ▼
Limpiar y normalizar columnas (GPO, GEN, ESP, DIF, VAR, CANTIDAD)
    │
    ▼
filtrar_formato_clave() → descartar filas con formato inválido
    │
    ▼
generar_expediente() → I-YYYY-DDMMYY{n}
    │
    ▼
insertar_filas() → MySQL tabla `cotizaciones`
    │
    ▼
generar_pdf() + enviar por Outlook COM
```

---

## Configuración clave

```python
# generar_cotizacion.py
MODO_PRUEBA = True        # True → PDFs van a CORREO_PRUEBA, no al remitente real
CORREO_PRUEBA = "alejandro.rodriguez@pharmagbc.com"

# main.py
INTERVALO_MONITOREO = 120  # segundos entre ciclos (configurable por env var)
BUZONES_AUTORIZADOS = [
    "ventas@serviciosademex.com",
    "ventas@gadmar.com.mx",
    "contacto.ademex@serviciosademex.com",
    "contacto@gadmar.com.mx",
]
```

Variables de entorno (`.env` o sistema):
```
MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DATABASE
INTERVALO_MONITOREO
```

---

## MySQL — Tablas

### `cotizaciones`
| Campo | Tipo | Notas |
|-------|------|-------|
| expediente | VARCHAR | I-YYYY-DDMMYY{n} |
| partida | INT | Consecutivo dentro del expediente |
| gpo/gen/esp/dif/var | VARCHAR | Componentes de la clave IMSS |
| cantidad | INT | Piezas solicitadas |
| clave | VARCHAR | gpo.gen.esp.dif.var |
| entry_id | VARCHAR UNIQUE | MD5(expediente+remitente+asunto+clave+cantidad) |

### `correos_procesados`
Auditoría de cada correo procesado: estado, intentos, expediente asignado, remitente real IMSS.

### `expediente_consecutivo`
Un registro por fecha con el consecutivo actual del día.

> ⚠️ La tabla `cotizaciones` **no se auto-crea** — debe existir antes de correr el monitor.

---

## Modo reprocesar

```bash
python main.py --desde 2026-06-01 --hasta 2026-06-01
```

- Borra `cotizaciones`, `correos_procesados` y `expediente_consecutivo` del rango de fechas
- Reprocesa todos los correos del rango desde Outlook
- Regenera PDFs y los envía al correo de prueba
- Útil cuando se corrige un bug de parseo y se necesita recargar datos

---

## Formatos de tabla IMSS reconocidos

El parser maneja dos modos:

**Con encabezado:** detecta columnas por nombre usando `EQUIVALENCIAS_COLUMNAS`:
- GPO / GRUPO
- GEN / GENERICO
- ESP / ESPECIFICO
- DIF / DIFERENCIADOR
- VAR / VARIANTE
- CANTIDAD / PIEZAS / NECESIDAD / SOLICITADO / **SOLICIT** / SOLICITUD / TOTAL / FALTANTE / INGRESO / REQUERIDA

**Sin encabezado:** búsqueda posicional de patrón `3-3-4-2-2 dígitos` en cada fila.

**Reparaciones automáticas:**
- Columna GEN con 4 dígitos + ESP ausente → se mueve GEN→ESP, GEN="000"
- Valores DIF/VAR en formato float (`0.0`) → se convierten a entero antes del zfill
- CANTIDAD con 9+ dígitos (clave concatenada) → se descarta, se toma el valor más corto

---

## Deduplicación de cotizaciones (persistente, ventana 10 min)

Evita procesar dos veces el mismo requerimiento. Cubre dos escenarios:
1. **Mismo correo a varios buzones** — el mismo remitente IMSS (ej. juanperez@imss.gob.mx) envía las mismas claves+cantidades a `ventas@` y `contacto@` con hasta ~10 min de diferencia. Solo se cotiza una vez. *(Antes el flujo lo trataba como 2 cotizaciones distintas y enviaba el PDF dos veces al correo de prueba.)*
2. **Eco reciente del hilo** — alguien de ventas responde el hilo re-conteniendo una cotización vieja. Si llegó dentro de la ventana, se omite.

**Cómo funciona:**
```python
firma = md5(remitente_real + "|" + sorted("gpo.gen.esp.dif.var=cantidad"))
```
- La firma se guarda en `estado_procesamiento.json` → `firmas_enviadas` con la **fecha de recepción** del correo original (persiste entre ciclos).
- Antes de insertar/enviar, se compara la firma del correo nuevo contra las guardadas. Si coincide y la diferencia de fechas de recepción es **< `VENTANA_DEDUP_SEGUNDOS` (600s = 10 min)** → es duplicado.
- **Un duplicado se omite por completo:** no genera expediente, no inserta en MySQL, no envía PDF. Se marca como `'duplicado'` en `correos_procesados`.

**Por qué la ventana se mide contra la fecha de recepción y no el reloj:**
Durante un reprocesar de día completo, todos los correos se procesan en segundos. Si se usara el reloj, dos pedidos genuinos de las mismas claves separados por horas se verían como duplicados. Usando la fecha de recepción del correo, solo los que realmente llegaron juntos (< 5 min) se deduplican.

**Configuración:** variable de entorno `VENTANA_DEDUP_SEGUNDOS` (default 600 = 10 min).

**Limitación conocida:** un eco que llega **más de 10 minutos** después del original NO se detecta (el escenario 2 menciona ecos de "horas o días"). Para cubrir esos casos hay que aumentar `VENTANA_DEDUP_SEGUNDOS` o cambiar a deduplicación permanente.

> Nota: como el inbox se ordena por fecha descendente (más nuevo primero), ante duplicados se conserva el correo más reciente y se omite el anterior. El contenido (claves+cantidades) es idéntico, así que la cotización resultante es la misma.

**Importante — la dedup NO está atada a ningún remitente específico:**
La firma usa `remitente_real` extraído dinámicamente de cada correo (línea "De: ...@imss.gob.mx" del cuerpo). Aplica igual para cualquier funcionario IMSS — no hay direcciones hardcodeadas. La regla es genérica: mismo remitente (sea quien sea) + mismas claves + mismas cantidades + dentro de 10 min → duplicado.

**Ciclo de vida de la firma:**
- Se registra en `firmas_enviadas` con la fecha de recepción del correo original.
- Al final de cada ciclo se poda: firmas con más de 10 min de antigüedad se eliminan del JSON.
- Esto mantiene el archivo pequeño y evita falsos positivos en pedidos futuros del mismo contenido.

**Resultado verificado en producción (reprocesar 2026-06-01):**
5 duplicados correctamente detectados y omitidos:
- COTIZACION 18022610 → a 0s del original
- SOLICITUD FARM 01 JUNIO → a 76s del original
- UMF 57 → a 53s del original
- HGZ 14 → a 6s del original
- UMF 55 → a 200s del original

---

## Expediente

Formato: `I-{AÑO}-{DD}{MM}{YY}{consecutivo}`  
Ejemplo: `I-2026-010626{n}` para el día 01/06/2026.

- Consecutivo almacenado en MySQL tabla `expediente_consecutivo`
- Si MySQL no disponible: contador en memoria (único dentro del proceso, se reinicia al reiniciar)
- Si ocurre excepción después de asignar el expediente: se revierte el consecutivo automáticamente (`revertir_expediente`)

---

## Bugs conocidos pendientes

| # | Descripción | Impacto |
|---|-------------|---------|
| 1 | `estado_procesamiento.json` crece indefinidamente (sin limpieza de IDs antiguos) | Lentitud a largo plazo |
| 2 | Sin reconexión automática a Outlook COM si Outlook se cierra/reinicia | Falla silenciosa hasta reiniciar |
| 3 | Tabla `cotizaciones` no se auto-crea en MySQL | Error en instalación nueva |
| 4 | CSVs de respaldo se acumulan en directorio raíz sin rotación ni límite | Acumulación de archivos |
| 5 | El inbox se itera completo cada ciclo (sin filtro `Restrict` por fecha) | Lento con inbox grande |
| 6 | Firmante en PDF es placeholder `"JUAN PÉREZ XXXX"` | PDF enviado con datos incorrectos |

---

## Bugs corregidos (sesión 2026-06-01)

### Bug 1 — GEN con 4 dígitos (UMF 24 AMECA, vidal.rodriguez@imss.gob.mx)
**Síntoma:** Correo procesado como "Sin tabla válida". Log: `filtrar_formato_clave: 9 fila(s) descartadas. Columnas con error: ['GEN']`.  
**Causa:** La tabla del correo omite la columna ESP y pone el específico (4 dígitos) en la columna GEN. El filtro esperaba GEN de 3 dígitos exactos.  
**Fix:** Antes del filtro de vacíos, detectar si GEN tiene 4 dígitos y ESP está vacía → mover GEN→ESP, GEN="0".  
**Código:** `main.py` líneas 607-615 (función `limpiar_tabla_con_encabezado`).

### Bug 2 — DIF/VAR en formato float "0.0"
**Síntoma:** Mismo correo UMF 24 AMECA seguía fallando. Log: `Columnas con error: ['DIF']. Muestra: [{'DIF': '0.0', ...}]`.  
**Causa:** Pandas lee celdas con cero como `float 0.0`. `str(0.0)` = `"0.0"`, que no pasa `fullmatch(r"\d{2}", ...)`.  
**Fix:** Función `_norm_int()` que convierte `float→int` antes de `zfill` para todos los campos numéricos (GPO, GEN, ESP, DIF, VAR).  
**Código:** `main.py` líneas 619-625.

### Bug 3 — CANTIDAD tomaba clave concatenada (UMF 55, ramon.martinezbe@imss.gob.mx)
**Síntoma:** Clave `060.456.0631.00.01` cargó CANTIDAD=`06045606310001` (la clave sin puntos). MySQL rechazó la fila con `Out of range value for column 'cantidad'`.  
**Causa:** `max(re.findall(r"\d+", x), key=len)` elegía la secuencia de dígitos **más larga**, que era la clave concatenada cuando aparecía en la celda.  
**Fix:** Filtrar secuencias de 9+ dígitos antes de elegir. Cantidades reales no superan 8 dígitos.  
**Código:** `main.py` función `_extraer_cantidad()` y `extraer_numero()`.

### Bug 4 — Hash de unicidad MySQL no incluía expediente (HGZ 14, victor.morenop@imss.gob.mx)
**Síntoma:** De 25 claves del correo HGZ 14 solo se insertaban 14 en MySQL. Las otras 11 se bloqueaban como duplicadas.  
**Causa:** El hash MD5 era `REMITENTE+ASUNTO+GPO+GEN+ESP+DIF+VAR+CANTIDAD` sin incluir el expediente. El mismo correo enviado en días anteriores ya tenía esas combinaciones en MySQL con el mismo hash.  
**Fix:** Agregar `EXPEDIENTE` al hash. El expediente incluye la fecha (`I-2026-010626n`) por lo que hashes de días distintos son siempre diferentes.  
**Código:** `main.py` línea 265-268.

### Bug 5 — `firmas_cotizadas` suprimía cotizaciones con mismas claves y distinta cantidad
**Síntoma:** Si el mismo proveedor mandaba el mismo set de claves con cantidades distintas (compra emergente), solo se generaba y enviaba la primera cotización.  
**Causa:** `firma = (remitente_real, frozenset(df["CLAVE"].tolist()))` — no incluía cantidades, por lo que el frozenset era idéntico.  
**Fix:** Incluir cantidad en la firma: `frozenset((row["CLAVE"], str(row["CANTIDAD"])) for _, row in df.iterrows())`.  
**Código:** `main.py` líneas 1071-1072.

### Bug 6 — Fallback expediente sin MySQL incorrecto
**Síntoma:** Sin conexión MySQL, el programa generaba `f"I-{year}-{ts}1"` donde `ts` es un timestamp en segundos. Dos correos en el mismo segundo recibían el mismo expediente. El formato tampoco coincidía con el estándar.  
**Fix:** Contador en memoria `_fallback_consecutivo` (dict por fecha), con el mismo formato `I-YYYY-DDMMYY{n}`.  
**Código:** `main.py` líneas 115-133.

### Bug 7 — `limpiar_mysql_por_fechas` no limpiaba `correos_procesados`
**Síntoma:** Al reprocesar varias veces el mismo día, el campo `intentos` en `correos_procesados` seguía incrementando (llegó a valores altos sin sentido).  
**Causa:** La función solo borraba `cotizaciones` y `expediente_consecutivo`, dejando intactos los registros de auditoría.  
**Fix:** Agregar DELETE de `correos_procesados` por el rango de fechas.  
**Código:** `main.py` líneas 266-271.

### Bug 8 — Expediente quemado en excepción post-asignación
**Síntoma:** Si ocurría cualquier error después de llamar `generar_expediente()` (que ya incrementó el contador en MySQL), el número quedaba "quemado" — el contador avanzó pero no hay registro con ese expediente.  
**Fix:** Envolver todo el bloque post-expediente en `try/except`. Si `expediente_usado` sigue en `False` cuando se captura la excepción, llamar `revertir_expediente()` que hace `consecutivo - 1`.  
**Código:** `main.py` función `revertir_expediente()` + bloque try/except en `procesar_correos`.

### Bug 9 — Columna "SOLICIT" no reconocida como CANTIDAD (UMF 55, ramon.martinezbe@imss.gob.mx)
**Síntoma:** Las cantidades cargadas en MySQL no coincidían con las del correo original. El correo usaba columna "SOLICIT" pero el programa tomaba los valores de otra columna.  
**Causa:** `EQUIVALENCIAS_COLUMNAS` tenía "solicitado" pero no "solicit" ni "solicitud". `normalizar_columna("SOLICIT")` devolvía `"SOLICIT"` sin mapear.  
**Fix:** Agregar `"solicit": "CANTIDAD"` y `"solicitud": "CANTIDAD"` a `EQUIVALENCIAS_COLUMNAS` y a las keywords de `detectar_fila_encabezado`.  
**Código:** `main.py` líneas 514-515 y 541-542.

---

## Historial de reprocesares del día 2026-06-01

Se realizaron múltiples reprocesares durante la sesión de depuración, cada uno corrigiendo un bug diferente:

| Hora | Motivo | Resultado |
|------|--------|-----------|
| 10:28 | Primer reprocesar post-reinicio | Bug DIF "0.0" encontrado |
| 10:30 | Fix DIF float aplicado | UMF 24 AMECA procesado ✅ |
| 10:41 | Fix hash + fix CANTIDAD | HGZ 14: 25/25 claves ✅ |
| 10:54 | Fix expediente + firmas + fallback + limpiar | Todos los bugs corregidos ✅ |
| 11:52 | Fix SOLICIT → CANTIDAD | UMF 55 cantidades correctas ✅ |

---

## Correos del 2026-06-01 y su estado final

| Correo | Remitente | Filas | Estado |
|--------|-----------|-------|--------|
| RV: cotización de compra 29/05/2026 UMF 57 | — | 31 | ✅ procesado |
| RV: CE HGR 46 01.06.2026 | — | — | ❌ sin tabla válida |
| RV: Solicitud de cotización UMAE HTOP | — | 1 | ✅ procesado |
| RV: SOLICITUD DE COTIZACIONES HGZ 14 | victor.morenop@imss.gob.mx | 25 | ✅ procesado (×2 cuentas) |
| RV: COTIZACION DE MEDICAMENTO UMF 24 AMECA | vidal.rodriguez@imss.gob.mx | 9 | ✅ procesado (fix GEN→ESP) |
| RV: Solicitud de cotización - Lidocaína | hector.abrica@imss.gob.mx | — | ❌ sin tabla válida |
| RV: COTIZACION UMF 55 - 28 05 2025 | ramon.martinezbe@imss.gob.mx | 9 | ✅ procesado (fix SOLICIT) |
| RV: Solicitud de cotización - Lacosamida | hector.abrica@imss.gob.mx | — | ❌ sin tabla válida |
| RV: Solicitud de cotización. | eliezer.avinar@imss.gob.mx | — | ❌ sin tabla válida |

---

## Notas operativas

- El monitor se corre con `python main.py` en una ventana normal de Windows
- El dashboard Streamlit corre como proceso separado: `streamlit run dashboard.py`
- Para detener el monitor: cerrar la ventana o Ctrl+C
- Para reprocesar un día: `python main.py --desde YYYY-MM-DD --hasta YYYY-MM-DD`
- Los PDFs se guardan en `cotizaciones_generadas/` con nombre `Cotizacion_YYYYMMDD_HHMMSS.pdf`
- Cada PDF tiene un sidecar `.json` con metadatos del envío
- Los CSVs `cotizaciones_imss_*.csv` son solo respaldo — el dashboard NO los usa, se pueden borrar sin impacto

---

## Cómo revisar el log

El log principal es `cotizaciones.log` (~8MB+). Para ver las últimas líneas:
```powershell
Get-Content cotizaciones.log -Tail 30
```

Para buscar un correo específico por remitente o asunto:
```powershell
Select-String -Path cotizaciones.log -Pattern "vidal.rodriguez|UMF 24" | Select-Object -ExpandProperty Line
```

**Indicadores clave en el log:**
- `CORREO VÁLIDO #N` → correo que pasó los filtros de validación
- `Tablas HTML válidas encontradas: N` → tabla detectada en el cuerpo HTML
- `Reenvío sin tabla en parte reciente — intentando primer bloque reenviado` → correo reenviado, buscando en bloque original
- `GEN→ESP reparado` → se aplicó la corrección de columna GEN con 4 dígitos
- `filtrar_formato_clave: N fila(s) descartadas. Columnas con error: [...]` → filas con formato incorrecto, el log ahora muestra qué columna falla y muestra 3 valores de ejemplo
- `Sin tabla válida (cuerpo ni adjuntos)` → correo sin tabla procesable
- `Insertadas en MySQL: N` → filas guardadas exitosamente
- `Expediente X revertido` → el consecutivo se decrementó porque hubo error post-asignación
- `Próxima revisión en 120s` → ciclo terminado, monitor esperando

---

## Cómo enviar un PDF manualmente por Outlook

Si necesitas reenviar un PDF de cotización generado a un correo específico:
```powershell
$pdf = "ruta\completa\al\Cotizacion_YYYYMMDD_HHMMSS.pdf"
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = "destinatario@ejemplo.com"
$mail.Subject = "RE: ASUNTO DEL CORREO — Cotización ADEMEX"
$mail.HTMLBody = "<p>Buen día, adjunto la cotización.</p>"
$mail.Attachments.Add($pdf) | Out-Null
$mail.Send()
```

---

## Revisión completa de robustez (2026-06-01)

Se realizó una revisión exhaustiva del código. Bugs encontrados en total: 14.

### Bugs corregidos en la sesión (9)
Ver sección "Bugs corregidos" más arriba.

### Bugs pendientes priorizados (5)

**🔴 Críticos:**
1. **Tabla `cotizaciones` no se auto-crea** — `asegurar_tablas_mysql()` solo crea `correos_procesados` y `expediente_consecutivo`. En instalación nueva el monitor falla silenciosamente fila por fila.
2. **Sin reconexión a Outlook COM** — `inbox` se obtiene una sola vez al arrancar. Si Outlook se cierra/reinicia, todas las iteraciones siguientes fallan con error COM capturado silenciosamente. El programa sigue vivo pero no procesa nada.

**🟠 Altos:**
3. **`estado_procesamiento.json` crece sin límite** — Ya ~339KB con miles de entry_ids. Se serializa completo cada 2 minutos. En meses será lento.
4. **Bug de timezone** — Línea 949: `fecha_correo <= ultima_fecha` puede lanzar `TypeError` si uno es timezone-aware y el otro naive.
5. **Rendimiento: itera todo el inbox cada ciclo** — Sin `Restrict` por fecha, con miles de correos cada revisión escanea todo.

**🟢 Bajos:**
- CSVs sin rotación ni límite
- Firmante `"JUAN PÉREZ XXXX"` en PDFs
- `detectar_fila_encabezado` incluye "unidad" pero `EQUIVALENCIAS_COLUMNAS` no lo mapea (inconsistencia)
- `_asignar_dif_var_posicional` frágil con columnas extra sin nombre

---

## Mejora de diagnóstico aplicada

`filtrar_formato_clave()` ahora loguea en WARNING exactamente qué columna falla y una muestra de los valores problemáticos:
```
[WARNING] filtrar_formato_clave: 9 fila(s) descartadas.
          Columnas con error: ['DIF'].
          Muestra: [{'GPO': '010', 'GEN': '000', 'ESP': '3047', 'DIF': '0.0', 'VAR': '00'}, ...]
```
Esto facilita diagnosticar futuros formatos de tabla no reconocidos sin necesidad de depurar en código.

---

## Historial de arranques del monitor (2026-06-01)

| Hora | PID | Motivo |
|------|-----|--------|
| 10:06 | 36860 | Arranque inicial del día |
| 10:21 | — | Reinicio tras correcciones (bugs 1-4) |
| 10:30 | — | Reprocesar UMF 24 AMECA |
| 10:41 | — | Reprocesar fix hash HGZ 14 |
| 10:54 | — | Reprocesar fix expediente/firmas/fallback/limpiar |
| 11:52 | — | Reprocesar fix SOLICIT |
| 13:19 | 12144 | Arranque tras contexto guardado en MD |
| 16:49 | — | Reprocesar con deduplicación activa (test) — 5 duplicados detectados |
| 16:52 | — | Reprocesar confirmación fix columna estado VARCHAR |
| ~17:00 | 20216 | Arranque con dedup 5 min |
| ~17:10 | 11964 | Reinicio con dedup ampliada a 10 min |

---

## Repositorio Git

- **Remote:** `https://github.com/alegabcontacto-ui/automatizacion-dashboard.git`
- **Rama principal:** `main`

### Historial de commits relevantes

| Hash | Mensaje |
|------|---------|
| `030d6ee` | feat: deduplicacion persistente de cotizaciones y mejoras de robustez |
| `add6827` | docs: actualizar contexto operativo de automatizacion |
| `19457ce` | feat: mejorar robustez en expedientes, parsing de tablas y contexto |
| `4e47f2d` | docs: agregar README, requirements y template de usuarios |
| `01e691f` | feat: commit inicial - automatizacion y dashboard IMSS |

### Qué incluye el commit `030d6ee`
Todo lo trabajado en la sesión 2026-06-01:
- Deduplicación persistente (ventana 10 min, basada en fecha de recepción)
- 9 bugs corregidos (GEN→ESP, DIF float, CANTIDAD clave, hash, firmas, fallback expediente, correos_procesados, expediente quemado, SOLICIT)
- Columna `estado` en `correos_procesados` ampliada a VARCHAR(50)
- Contexto completo en `ContextoAutomatizacionesCorreo.md`

### Comandos útiles de git
```bash
# Ver estado actual
git status

# Ver últimos commits
git log --oneline -5

# Subir cambios
git add main.py ContextoAutomatizacionesCorreo.md
git commit -m "descripcion"
git push origin main
```

---

## Auditoría post-fixes — 2026-06-02

Commit aplicado: `3c6bc68` (tres fixes críticos implementados por Claude Code).  
Fixes importantes implementados por el usuario de forma independiente.

### ✅ OK — verificado en código

**`generar_expediente` — `SELECT ... FOR UPDATE`**
`conn.start_transaction()` antes del SELECT, `FOR UPDATE` en la query, `conn.rollback()` en el except, `cursor.close()` en `finally`. Cubre row lock en fila existente y gap lock de InnoDB en primera inserción del día. Correcto para multi-proceso.

**Comparación fechas naive vs aware**
`fecha_correo_cmp` (tzinfo=None) usado consistentemente en las tres comparaciones: línea 1056 (`break` por `ultima_fecha`), línea 1070 (comparación `nueva_ultima_fecha`), línea 1071 (asignación `nueva_ultima_fecha`). El JSON siempre guarda y lee naive. Ciclo consistente.

**Auto-creación tablas `cotizaciones` y `expediente_consecutivo`**
Ambas en `asegurar_tablas_mysql`. Schema de `cotizaciones` coincide exactamente con `insertar_filas`: `entry_id VARCHAR(32) UNIQUE` activa el `ON DUPLICATE KEY UPDATE`, `cantidad INT` acepta los `None` que genera línea 376.

---

### ❌ Sigue mal — pendiente de fix

| # | Problema | Ubicación |
|---|---|---|
| 3 | `ids_procesados` crece sin límite en JSON, sin purga por fecha | línea 1276 |
| 6 | Cursores sin `finally` en `registrar_correo_inicio`, `actualizar_correo_estado`, `insertar_filas` | líneas 416, 442, 362 |
| 8 | Sin `Restrict` de Outlook — itera todo el inbox cada ciclo | líneas 1027-1028 |
| 9 | Sin límite de tamaño de adjuntos (`att.Size` no verificado) | línea 958 |
| 10 | Sin reconexión COM — `COMError` de Outlook crashea el proceso sin recuperación | línea 1027 |

---

### ⚠️ Riesgos nuevos introducidos por los fixes

**`conn.rollback()` en `generar_expediente` puede revertir un INSERT pendiente de `registrar_correo_inicio`**

Si `registrar_correo_inicio` ejecuta su `cursor.execute()` con éxito pero su `conn.commit()` falla (ventana de red microscópica), la transacción queda abierta sin commit. Cuando `generar_expediente` llama `conn.start_transaction()`, InnoDB lanza `InternalError: already in transaction`. El except llama `conn.rollback()`, que revierte el INSERT staged de `correos_procesados`. El correo no queda en auditoría. Impacto bajo con MySQL local.

**`ultima_fecha` en JSON existente puede ser offset-aware**

Si el archivo `estado_procesamiento.json` fue escrito por el código anterior al fix (que guardaba `fecha_correo` directamente con timezone), el campo `ultima_fecha` podría contener `"2026-06-01T10:30:00+06:00"`. `datetime.fromisoformat()` preserva el offset, devolviendo un datetime aware. La comparación `fecha_correo_cmp <= ultima_fecha` (naive <= aware) seguiría lanzando `TypeError`.

**Acción requerida:** verificar `estado_procesamiento.json`. Si `ultima_fecha` contiene `+` o `Z`, borrar ese campo o el archivo antes de arrancar.

---

### Pruebas recomendadas

| Ítem | Prueba |
|---|---|
| `FOR UPDATE` atómico | Ejecutar dos instancias simultáneas y verificar en MySQL que no hay expedientes duplicados en `cotizaciones`. |
| Timezone en JSON | Abrir `estado_procesamiento.json` y revisar si `ultima_fecha` contiene `+` o `Z`. Si sí, limpiar el campo. |
| Tablas auto-creadas | `DROP TABLE IF EXISTS cotizaciones; DROP TABLE IF EXISTS expediente_consecutivo;` → reiniciar monitor → `SHOW TABLES;` |
| Leak de cursores | `SHOW PROCESSLIST` antes y después de 10 ciclos; el número de conexiones `Sleep` no debe crecer. |
| Crash COM | Cerrar Outlook a mano mientras el monitor corre y esperar al siguiente ciclo. Verificar si el proceso muere o recupera. |
| Crecimiento JSON | Revisar tamaño de `estado_procesamiento.json` después de una semana. Si supera 1 MB, el fix de purga sigue pendiente. |

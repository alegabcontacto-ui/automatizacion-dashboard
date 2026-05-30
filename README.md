# Automatización y Dashboard — Cotizaciones IMSS

Sistema completo para procesar solicitudes de cotización IMSS recibidas por correo, generar cotizaciones en PDF y visualizar el historial en un dashboard web.

## Arquitectura

```
automatizacion/   → Servicio de fondo: monitorea Outlook, extrae datos, genera y envía PDFs
dashboard/        → Dashboard web: visualización, análisis y administración
```

Los dos componentes comparten una base de datos MySQL (`cotizaciones_imss`) y una carpeta de PDFs generados.

---

## Requisitos previos

- Windows 10/11 con **Microsoft Outlook** instalado y configurado
- **Python 3.10+**
- **MySQL 8.0+** corriendo en localhost
- Cuenta de GitHub para clonar el repositorio

---

## Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/alegabcontacto-ui/automatizacion-dashboard.git
cd automatizacion-dashboard
```

### 2. Crear la base de datos MySQL

Conectarse a MySQL y ejecutar:

```sql
CREATE DATABASE cotizaciones_imss CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE cotizaciones_imss;

CREATE TABLE cotizaciones (
    id INT AUTO_INCREMENT PRIMARY KEY,
    expediente VARCHAR(30),
    partida INT,
    gpo VARCHAR(10),
    gen VARCHAR(10),
    esp VARCHAR(10),
    dif VARCHAR(10),
    var VARCHAR(10),
    clave VARCHAR(30),
    cantidad INT,
    asunto TEXT,
    remitente VARCHAR(200),
    fecha_correo DATETIME,
    fecha_insercion DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uk_expediente_partida (expediente, partida)
);

CREATE TABLE correos_procesados (
    id INT AUTO_INCREMENT PRIMARY KEY,
    entry_id VARCHAR(200) UNIQUE,
    asunto TEXT,
    remitente VARCHAR(200),
    fecha_correo DATETIME,
    fecha_proceso DATETIME DEFAULT CURRENT_TIMESTAMP,
    estado VARCHAR(50),
    filas_detectadas INT DEFAULT 0,
    filas_insertadas INT DEFAULT 0,
    expediente VARCHAR(30),
    intentos INT DEFAULT 1,
    notas TEXT
);

CREATE TABLE expediente_consecutivo (
    fecha DATE PRIMARY KEY,
    consecutivo INT DEFAULT 0
);
```

### 3. Configurar la automatización

```bash
cd automatizacion
pip install -r requirements.txt
cp .env.example .env
```

Editar `.env` con tus datos reales.

### 4. Configurar el dashboard

```bash
cd ../dashboard
pip install -r requirements_dashboard.txt
cp .env.example .env
cp usuarios.yaml.example usuarios.yaml
```

Editar `.env` con tus datos reales.

Crear el primer usuario administrador:

```bash
python agregar_usuario.py
```

---

## Uso

### Iniciar la automatización

```bash
cd automatizacion
python main.py
```

El servicio monitorea Outlook cada 120 segundos (configurable con `INTERVALO_MONITOREO` en `.env`).

Para reprocesar un rango de fechas:

```bash
python main.py --desde 2026-05-01 --hasta 2026-05-31
```

### Iniciar el dashboard

```bash
cd dashboard
streamlit run dashboard.py
```

O usar el script de Windows:

```
iniciar_dashboard.bat
```

El dashboard queda disponible en `http://localhost:8501`.

---

## Variables de entorno

### `automatizacion/.env`

| Variable | Descripción | Ejemplo |
|---|---|---|
| `MYSQL_HOST` | Host de MySQL | `localhost` |
| `MYSQL_PORT` | Puerto de MySQL | `3306` |
| `MYSQL_USER` | Usuario MySQL | `root` |
| `MYSQL_PASSWORD` | Contraseña MySQL | `tu_password` |
| `MYSQL_DATABASE` | Nombre de la base de datos | `cotizaciones_imss` |
| `INTERVALO_MONITOREO` | Segundos entre revisiones de correo | `120` |

### `dashboard/.env`

| Variable | Descripción | Ejemplo |
|---|---|---|
| `MYSQL_HOST` | Host de MySQL | `localhost` |
| `MYSQL_PORT` | Puerto de MySQL | `3306` |
| `MYSQL_USER` | Usuario MySQL | `root` |
| `MYSQL_PASSWORD` | Contraseña MySQL | `tu_password` |
| `MYSQL_DATABASE` | Nombre de la base de datos | `cotizaciones_imss` |
| `PDF_DIR` | Ruta a la carpeta de PDFs generados | `C:\ruta\a\automatizacion\cotizaciones_generadas` |
| `LOGO_PATH` | Ruta al logo de la empresa (.png) | `C:\ruta\a\ademex.png` |

---

## Estructura de archivos

```
cotizaciones-ademex/
├── automatizacion/
│   ├── main.py                  # Motor principal de procesamiento
│   ├── generar_cotizacion.py    # Generación de PDF y envío por Outlook
│   ├── test_cotizacion.py       # Pruebas unitarias
│   ├── requirements.txt
│   └── .env.example
├── dashboard/
│   ├── dashboard.py             # Aplicación Streamlit
│   ├── agregar_usuario.py       # Gestión de usuarios por CLI
│   ├── iniciar_dashboard.bat    # Acceso directo Windows
│   ├── requirements_dashboard.txt
│   ├── .env.example
│   └── usuarios.yaml.example
└── .gitignore
```

---

## Modo prueba

En `generar_cotizacion.py`, la variable `MODO_PRUEBA` controla el destino de los correos:

- `MODO_PRUEBA = True` → todas las cotizaciones se envían a `CORREO_PRUEBA` (para desarrollo)
- `MODO_PRUEBA = False` → se responde al remitente real del IMSS

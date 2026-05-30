# =============================================================================
#  Monitor de Cotizaciones IMSS — v4.0
#  GADMAR SA DE CV
#  Mejoras: rutas relativas via .env, KPIs con deltas, filtro global de fechas,
#  insights operativos en home, código modular, manejo de errores robusto.
# =============================================================================

import base64
import json as _json_mod
import os
from datetime import datetime, timedelta
from pathlib import Path

import bcrypt
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
import yaml
from dotenv import load_dotenv
from streamlit_autorefresh import st_autorefresh
from yaml.loader import SafeLoader

# ── Cargar variables de entorno (.env junto al script) ────────────────────────
load_dotenv(Path(__file__).parent / ".env")

# ── Rutas relativas al directorio del script ──────────────────────────────────
BASE_DIR      = Path(__file__).parent
USUARIOS_PATH = BASE_DIR / "usuarios.yaml"
REMEMBER_PATH = BASE_DIR / ".remember.json"
LOGO_PATH     = Path(os.getenv("LOGO_PATH", str(BASE_DIR / "ademex.png")))
PDF_DIR       = Path(os.getenv("PDF_DIR",   str(BASE_DIR / "cotizaciones_generadas")))

# ── Paleta ────────────────────────────────────────────────────────────────────
NAVY   = "#0f172a"
BLUE   = "#2563eb"
SKY    = "#0ea5e9"
GREEN  = "#059669"
AMBER  = "#d97706"
RED    = "#dc2626"
MUTED  = "#6b7280"

st.set_page_config(
    page_title="Monitor · Cotizaciones IMSS",
    page_icon="📋",
    layout="wide",
)


# =============================================================================
#  AUTENTICACIÓN
# =============================================================================

def _verificar(usuario: str, clave: str) -> bool:
    try:
        with open(USUARIOS_PATH, encoding="utf-8") as f:
            cfg = yaml.load(f, Loader=SafeLoader)
        datos = cfg["credentials"]["usernames"].get(usuario)
        return datos is not None and bcrypt.checkpw(clave.encode(), datos["password"].encode())
    except Exception:
        return False

def _nombre(usuario: str) -> str:
    try:
        with open(USUARIOS_PATH, encoding="utf-8") as f:
            cfg = yaml.load(f, Loader=SafeLoader)
        return cfg["credentials"]["usernames"][usuario].get("name", usuario)
    except Exception:
        return usuario

def _load_remember():
    try:
        if REMEMBER_PATH.exists():
            data = _json_mod.loads(REMEMBER_PATH.read_text(encoding="utf-8"))
            u = data.get("u", "")
            p = base64.b64decode(data.get("p", "").encode()).decode() if data.get("p") else ""
            return u, p
    except Exception:
        pass
    return "", ""

def _save_remember(usuario: str, clave: str):
    REMEMBER_PATH.write_text(
        _json_mod.dumps({"u": usuario, "p": base64.b64encode(clave.encode()).decode()}),
        encoding="utf-8",
    )

def _clear_remember():
    if REMEMBER_PATH.exists():
        REMEMBER_PATH.unlink()


# =============================================================================
#  ADMIN — USUARIOS
# =============================================================================

def _cargar_usuarios() -> dict:
    with open(USUARIOS_PATH, encoding="utf-8") as f:
        return yaml.load(f, Loader=SafeLoader)

def _guardar_usuarios(cfg: dict):
    with open(USUARIOS_PATH, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False)

def _es_admin(usuario: str) -> bool:
    try:
        cfg = _cargar_usuarios()
        return cfg["credentials"]["usernames"].get(usuario, {}).get("role") == "admin"
    except Exception:
        return False

def _agregar_usuario(usuario: str, nombre: str, email: str, clave: str, rol: str = "user"):
    cfg = _cargar_usuarios()
    hashed = bcrypt.hashpw(clave.encode(), bcrypt.gensalt()).decode()
    cfg["credentials"]["usernames"][usuario] = {
        "name": nombre, "email": email or usuario, "password": hashed, "role": rol,
    }
    _guardar_usuarios(cfg)

def _eliminar_usuario(usuario: str):
    cfg = _cargar_usuarios()
    cfg["credentials"]["usernames"].pop(usuario, None)
    _guardar_usuarios(cfg)

def _cambiar_password(usuario: str, nueva_clave: str):
    cfg = _cargar_usuarios()
    if usuario in cfg["credentials"]["usernames"]:
        cfg["credentials"]["usernames"][usuario]["password"] = (
            bcrypt.hashpw(nueva_clave.encode(), bcrypt.gensalt()).decode()
        )
        _guardar_usuarios(cfg)

def _cambiar_rol(usuario: str, nuevo_rol: str):
    cfg = _cargar_usuarios()
    if usuario in cfg["credentials"]["usernames"]:
        cfg["credentials"]["usernames"][usuario]["role"] = nuevo_rol
        _guardar_usuarios(cfg)

def _contar_admins() -> int:
    try:
        cfg = _cargar_usuarios()
        return sum(1 for u in cfg["credentials"]["usernames"].values() if u.get("role") == "admin")
    except Exception:
        return 0


# =============================================================================
#  PANTALLA DE LOGIN
# =============================================================================

if not st.session_state.get("auth_ok"):
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    * { font-family: 'Inter', sans-serif !important; }
    [data-testid="stAppViewContainer"] { background: #060d1f !important; }
    [data-testid="stSidebar"], [data-testid="collapsedControl"],
    [data-testid="stHeader"], #MainMenu, footer, header { display: none !important; }
    .block-container { max-width: 380px !important; padding: 7vh 1.5rem 2rem !important; margin: 0 auto !important; }
    [data-testid="stTextInput"] label,
    [data-testid="stTextInput"] label p {
        color: rgba(255,255,255,0.4) !important; font-size: 0.7rem !important;
        font-weight: 600 !important; letter-spacing: 0.12em !important; text-transform: uppercase !important;
    }
    [data-testid="stTextInput"] > div > div > input {
        background: rgba(255,255,255,0.04) !important; border: 1px solid rgba(255,255,255,0.09) !important;
        border-radius: 8px !important; color: #f1f5f9 !important; font-size: 0.88rem !important;
        height: 2.85rem !important; padding: 0 0.9rem !important; caret-color: #2563eb !important;
    }
    [data-testid="stTextInput"] > div > div > input:focus {
        background: rgba(255,255,255,0.07) !important; border-color: #2563eb !important;
        box-shadow: 0 0 0 3px rgba(37,99,235,0.18) !important; outline: none !important;
    }
    [data-testid="stTextInput"] > div > div > input::placeholder { color: rgba(255,255,255,0.14) !important; }
    [data-testid="stFormSubmitButton"] > button {
        background: #2563eb !important; color: white !important; border: none !important;
        border-radius: 8px !important; font-weight: 700 !important; font-size: 0.78rem !important;
        letter-spacing: 0.14em !important; text-transform: uppercase !important;
        height: 2.85rem !important; margin-top: 0.4rem !important;
        box-shadow: 0 4px 24px rgba(37,99,235,0.4) !important;
    }
    [data-testid="stFormSubmitButton"] > button:hover {
        background: #1d4ed8 !important; box-shadow: 0 6px 32px rgba(37,99,235,0.55) !important;
        transform: translateY(-1px) !important;
    }
    [data-testid="stForm"] {
        border: 1px solid rgba(255,255,255,0.07) !important; border-radius: 14px !important;
        padding: 1.6rem 1.4rem !important; background: rgba(255,255,255,0.025) !important;
    }
    [data-testid="stForm"] [data-testid="stCheckbox"] label p {
        color: rgba(255,255,255,0.55) !important; font-size: 0.75rem !important;
    }
    [data-testid="stAlert"] { border-radius: 8px !important; }
    [data-testid="stAlert"] p { font-size: 0.82rem !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align:center;margin:1.8rem 0 2.2rem">
        <div style="font-size:1.15rem;font-weight:700;color:#f8fafc;letter-spacing:-.02em;margin-bottom:.45rem">
            Monitor de Cotizaciones IMSS
        </div>
        <div style="font-size:.68rem;color:rgba(255,255,255,0.28);letter-spacing:.14em;text-transform:uppercase;font-weight:500">
            GADMAR SA DE CV &nbsp;&middot;&nbsp; Acceso restringido
        </div>
    </div>
    """, unsafe_allow_html=True)

    _rem_u, _rem_p = _load_remember()
    with st.form("login_form"):
        usuario    = st.text_input("Usuario",    value=_rem_u, placeholder="correo@empresa.com")
        contrasena = st.text_input("Contraseña", value=_rem_p, type="password", placeholder="••••••••")
        recordar   = st.checkbox("Recordar mis datos", value=bool(_rem_u))
        ok         = st.form_submit_button("Ingresar", use_container_width=True)

    if ok:
        if _verificar(usuario, contrasena):
            if recordar:
                _save_remember(usuario, contrasena)
            else:
                _clear_remember()
            st.session_state.update({"auth_ok": True, "auth_user": usuario, "auth_name": _nombre(usuario)})
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")

    st.markdown("""
    <p style="text-align:center;color:rgba(255,255,255,0.13);font-size:.62rem;margin-top:2rem;letter-spacing:.06em">
        &copy; 2026 GADMAR SA DE CV
    </p>
    """, unsafe_allow_html=True)
    st.stop()


# =============================================================================
#  CSS GLOBAL DEL DASHBOARD
# =============================================================================

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
* {{ font-family: 'Inter', sans-serif !important; }}
/* Ocultar ícono roto del expander — cubre múltiples versiones de Streamlit */
[data-testid="stExpanderToggleIcon"],
[data-testid="stExpander"] summary [data-testid="stIcon"],
[data-testid="stExpander"] summary [data-testid="stIconMaterial"] {{
    font-size: 0 !important; line-height: 0 !important;
    width: 0 !important; overflow: hidden !important;
}}
/* Chevron CSS como reemplazo visual — sin dependencia de fuentes externas */
[data-testid="stExpander"] > summary::before {{
    content: '' !important; display: inline-block !important; flex-shrink: 0 !important;
    width: 6px !important; height: 6px !important;
    border-right: 2px solid #64748b !important; border-bottom: 2px solid #64748b !important;
    transform: rotate(-45deg) !important; transition: transform .15s !important;
    margin-right: .6rem !important; margin-bottom: 1px !important;
}}
[data-testid="stExpander"][open] > summary::before {{
    transform: rotate(45deg) !important; margin-bottom: -2px !important;
}}

[data-testid="stAppViewContainer"] {{ background: #f4f6f9 !important; }}
[data-testid="stHeader"] {{ display: none; }}
#MainMenu, footer {{ display: none; }}
.block-container {{ padding: 1.8rem 2.2rem !important; max-width: 1440px; }}

/* ── Sidebar ── */
[data-testid="stSidebar"] {{ background: {NAVY} !important; border-right: none !important; }}
[data-testid="stSidebar"] > div:first-child {{
    display: flex !important; flex-direction: column !important;
    height: 100% !important; padding-bottom: 1rem !important;
}}
.sidebar-spacer {{ flex: 1 !important; min-height: 1rem !important; }}
[data-testid="stSidebar"] * {{ color: #94a3b8; }}
[data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"] {{ display: none !important; }}

/* ── Nav radio ── */
[data-testid="stSidebar"] [data-testid="stRadio"] > div > div {{ gap: 2px !important; }}
[data-testid="stSidebar"] [data-testid="stRadio"] label {{
    background: transparent !important; border-radius: 8px !important;
    padding: .48rem .9rem !important; cursor: pointer !important;
    color: #94a3b8 !important; font-size: .83rem !important; font-weight: 500 !important;
    transition: background .15s, color .15s !important;
}}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {{
    background: rgba(255,255,255,0.06) !important; color: #e2e8f0 !important;
}}
[data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child {{ display: none !important; }}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {{
    background: rgba(37,99,235,0.22) !important; color: #93c5fd !important; font-weight: 700 !important;
}}

/* ── KPI Cards ── */
.kpi-card {{
    background: #ffffff; border-radius: 14px; padding: 1.2rem 1.4rem;
    border-top: 3px solid transparent;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.05);
    transition: transform .18s ease, box-shadow .18s ease;
}}
.kpi-card:hover {{ transform: translateY(-3px); box-shadow: 0 2px 6px rgba(0,0,0,0.06), 0 16px 40px rgba(0,0,0,0.09); }}
.kpi-icon  {{ font-size:1.15rem; margin-bottom:.4rem; opacity:.85; }}
.kpi-label {{ font-size:.67rem; font-weight:700; color:#9ca3af; text-transform:uppercase; letter-spacing:.1em; margin-bottom:.35rem; }}
.kpi-value {{ font-size:1.85rem; font-weight:800; color:{NAVY}; line-height:1.1; }}
.kpi-delta-up   {{ font-size:.73rem; color:#059669; font-weight:600; margin-top:.25rem; }}
.kpi-delta-down {{ font-size:.73rem; color:#dc2626; font-weight:600; margin-top:.25rem; }}
.kpi-delta-flat {{ font-size:.73rem; color:#9ca3af;  font-weight:600; margin-top:.25rem; }}

/* ── Botones ── */
.stButton > button {{
    border-radius: 8px !important; font-weight: 600 !important;
    font-size: 0.8rem !important; border: none !important; transition: all .18s ease !important;
}}
.stButton > button:hover {{ transform: translateY(-1px) !important; box-shadow: 0 4px 14px rgba(37,99,235,.25) !important; }}
.stDownloadButton > button {{
    border-radius: 8px !important; font-weight: 600 !important; font-size: 0.8rem !important;
    background: {GREEN} !important; color: white !important; border: none !important;
}}
.stDownloadButton > button:hover {{ transform: translateY(-1px) !important; box-shadow: 0 4px 14px rgba(5,150,105,.3) !important; }}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
    background: #ffffff; border-radius: 10px; padding: .25rem; gap: .2rem;
    border: 1px solid rgba(0,0,0,0.07); box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}}
.stTabs [data-baseweb="tab"] {{
    border-radius: 7px !important; font-weight: 500 !important; font-size: 0.82rem !important;
    padding: .45rem 1.1rem !important; color: #9ca3af !important; transition: all .15s ease !important;
}}
.stTabs [data-baseweb="tab"]:hover {{ background: #f9fafb !important; color: {NAVY} !important; }}
.stTabs [aria-selected="true"] {{
    background: {NAVY} !important; color: white !important; font-weight: 700 !important;
    box-shadow: 0 2px 8px rgba(15,23,42,.25) !important;
}}

/* ── DataFrames ── */
[data-testid="stDataFrame"] > div {{
    border-radius: 12px !important; border: 1px solid rgba(0,0,0,0.07) !important;
    box-shadow: 0 1px 3px rgba(0,0,0,.04) !important; overflow: hidden;
}}
[data-testid="stDataFrame"] .dvn-scroller .ag-row-even {{ background-color: #f9fafb !important; }}
[data-testid="stDataFrame"] .dvn-scroller .ag-row-odd  {{ background-color: #ffffff !important; }}
[data-testid="stDataFrame"] .dvn-scroller .ag-row:hover {{ background-color: #eff6ff !important; }}

/* ── Insight cards ── */
.insight-card {{
    background: #ffffff; border-radius: 14px; border: 1px solid rgba(0,0,0,0.07);
    padding: 1.4rem 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.05);
    margin-bottom: 1rem; height: 100%;
}}
.insight-num   {{ font-size:2rem; font-weight:800; color:{BLUE}; line-height:1; }}
.insight-label {{ font-size:.7rem; font-weight:700; color:#9ca3af; text-transform:uppercase; letter-spacing:.1em; margin-top:.3rem; }}
.insight-text  {{ font-size:.84rem; color:#374151; margin-top:.8rem; line-height:1.65; border-top:1px solid rgba(0,0,0,.07); padding-top:.8rem; }}
.tag {{ display:inline-block; padding:.18rem .55rem; border-radius:5px; font-size:.72rem; font-weight:600; }}
.tag-up   {{ background:#dcfce7; color:#166534; }}
.tag-down {{ background:#fee2e2; color:#991b1b; }}
.tag-flat {{ background:#f1f5f9; color:#64748b; }}

/* ── Alerta banner ── */
.alert-warn {{
    background:#fffbeb; border:1px solid #fbbf24; border-left:4px solid #f59e0b;
    border-radius:10px; padding:.9rem 1.2rem; margin-bottom:1rem;
    display:flex; align-items:flex-start; gap:.8rem;
}}
.alert-err {{
    background:#fef2f2; border:1px solid #fca5a5; border-left:4px solid #ef4444;
    border-radius:10px; padding:.9rem 1.2rem; margin-bottom:1rem;
    display:flex; align-items:flex-start; gap:.8rem;
}}

/* ── Skeleton ── */
@keyframes shimmer {{ 0%{{background-position:-900px 0}} 100%{{background-position:900px 0}} }}
.skel {{
    background: linear-gradient(90deg,#e9ecef 25%,#dee2e6 50%,#e9ecef 75%);
    background-size: 900px 100%; animation: shimmer 1.5s infinite linear; border-radius: 10px;
}}

/* ── Pulse ── */
@keyframes pulse {{ 0%,100%{{opacity:1}} 50%{{opacity:.3}} }}
.pulse {{ animation: pulse 2s infinite; }}

hr  {{ border-color:rgba(0,0,0,.08) !important; margin:1rem 0 !important; }}
.sec {{
    font-size:.67rem; font-weight:700; color:#9ca3af;
    text-transform:uppercase; letter-spacing:.12em; margin-bottom:.75rem;
}}

/* ── Filtro global ── */
.filter-bar {{
    background:#ffffff; border-radius:12px; border:1px solid rgba(0,0,0,0.07);
    padding:.9rem 1.2rem; margin-bottom:1rem;
    box-shadow:0 1px 3px rgba(0,0,0,.04);
}}
</style>
""", unsafe_allow_html=True)

# Auto-refresh cada 2 minutos
st_autorefresh(interval=120_000, key="auto_refresh")

# Botón flotante para abrir/cerrar sidebar
components.html("""
<script>
(function() {
    function injectBtn() {
        var doc = window.parent.document;
        if (doc.getElementById('sidebar-fab')) return;
        var btn = doc.createElement('button');
        btn.id = 'sidebar-fab';
        btn.innerHTML = '&#9776;';
        btn.title = 'Abrir / Cerrar panel';
        btn.style.cssText =
            'position:fixed;top:14px;left:14px;z-index:9999999;background:#0f172a;color:white;' +
            'border:none;border-radius:8px;width:40px;height:40px;font-size:1.25rem;cursor:pointer;' +
            'box-shadow:0 2px 12px rgba(0,0,0,0.3);transition:all .18s ease;' +
            'display:flex;align-items:center;justify-content:center;';
        btn.onmouseenter = function(){ btn.style.background='#1e293b'; btn.style.transform='scale(1.06)'; };
        btn.onmouseleave = function(){ btn.style.background='#0f172a'; btn.style.transform='scale(1)'; };
        btn.onclick = function(){
            var native = doc.querySelector('[data-testid="collapsedControl"]') ||
                         doc.querySelector('[data-testid="stSidebarCollapseButton"] button');
            if (native) native.click();
        };
        doc.body.appendChild(btn);
    }
    setTimeout(injectBtn, 600);

    // Ocultar texto de ícono roto en expanders (p.ej. "_arrow_right", "expand_more")
    function fixExpanderIcons() {
        var doc = window.parent.document;
        doc.querySelectorAll('[data-testid="stExpander"] summary').forEach(function(s) {
            Array.from(s.children).forEach(function(el) {
                var t = (el.textContent || '').trim();
                // Texto que parece nombre de ícono: solo letras minúsculas, guiones bajos, sin espacios
                if (t && !t.includes(' ') && /^[_a-z][a-z_]+$/.test(t) && t.length < 50) {
                    el.style.setProperty('font-size',   '0',             'important');
                    el.style.setProperty('line-height', '0',             'important');
                    el.style.setProperty('width',       '0',             'important');
                    el.style.setProperty('overflow',    'hidden',        'important');
                    el.style.setProperty('display',     'inline-block',  'important');
                }
            });
        });
    }
    setTimeout(fixExpanderIcons, 900);
    setInterval(fixExpanderIcons, 3000);
})();
</script>
""", height=0)


# =============================================================================
#  HELPERS VISUALES
# =============================================================================

def kpi_card(label, value, icon, color, delta=None, delta_dir=None):
    d = ""
    if delta:
        cls   = {"up": "kpi-delta-up", "down": "kpi-delta-down"}.get(delta_dir, "kpi-delta-flat")
        arrow = {"up": "▲ ", "down": "▼ "}.get(delta_dir, "")
        d = f'<div class="{cls}">{arrow}{delta}</div>'
    return (
        f'<div class="kpi-card" style="border-top-color:{color}">'
        f'<div class="kpi-icon">{icon}</div>'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>{d}</div>'
    )

def bar_colors(n, rgb=(37, 99, 235)):
    r, g, b = rgb
    return [f"rgba({r},{g},{b},{.3+.7*(i/max(n-1,1)):.2f})" for i in range(n)]

def plotly_layout(**ov):
    base = dict(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=11, color="#374151"),
        margin=dict(l=0, r=8, t=6, b=0), showlegend=False,
    )
    base.update(ov)
    return base

def zebra_style(df):
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    for i, idx in enumerate(df.index):
        styles.loc[idx] = "background-color:#f9fafb" if i % 2 == 0 else "background-color:white"
    return styles

def alerta_warn(titulo, cuerpo):
    st.markdown(
        f"<div class='alert-warn'>"
        f"<span style='font-size:1.3rem;line-height:1.2'>⚠️</span>"
        f"<div><div style='font-weight:700;color:#92400e;font-size:.85rem;margin-bottom:.2rem'>{titulo}</div>"
        f"<div style='color:#b45309;font-size:.78rem;line-height:1.55'>{cuerpo}</div></div></div>",
        unsafe_allow_html=True,
    )

def alerta_err(titulo, cuerpo):
    st.markdown(
        f"<div class='alert-err'>"
        f"<span style='font-size:1.3rem;line-height:1.2'>🔴</span>"
        f"<div><div style='font-weight:700;color:#991b1b;font-size:.85rem;margin-bottom:.2rem'>{titulo}</div>"
        f"<div style='color:#b91c1c;font-size:.78rem;line-height:1.55'>{cuerpo}</div></div></div>",
        unsafe_allow_html=True,
    )


# =============================================================================
#  CARGA DE DATOS
# =============================================================================

@st.cache_data(ttl=60, show_spinner=False)
def cargar_datos() -> pd.DataFrame:
    """Carga cotizaciones desde MySQL. Devuelve DataFrame vacío si hay error."""
    import mysql.connector
    try:
        conn = mysql.connector.connect(
            host=os.getenv("MYSQL_HOST", "localhost"),
            port=int(os.getenv("MYSQL_PORT", 3306)),
            user=os.getenv("MYSQL_USER", "root"),
            password=os.getenv("MYSQL_PASSWORD", ""),
            database=os.getenv("MYSQL_DATABASE", "cotizaciones_imss"),
            connection_timeout=5,
        )
        df = pd.read_sql(
            "SELECT gpo, gen, esp, dif, var, cantidad, asunto, remitente, fecha_correo "
            "FROM cotizaciones ORDER BY fecha_correo DESC",
            conn,
        )
        conn.close()
    except Exception as e:
        st.session_state["db_error"] = str(e)
        return pd.DataFrame()

    df["fecha_correo"]    = pd.to_datetime(df["fecha_correo"])
    df["fecha"]           = df["fecha_correo"].dt.date
    df["remitente_corto"] = df["remitente"].apply(lambda x: x.split("@")[0] if "@" in str(x) else str(x))
    df["clave_imss"]      = df["gpo"].astype(str) + "-" + df["gen"].astype(str) + "-" + df["esp"].astype(str)
    st.session_state.pop("db_error", None)
    return df


@st.cache_data(ttl=30, show_spinner=False)
def cargar_correos() -> pd.DataFrame:
    """Carga el historial de correos procesados desde MySQL."""
    import mysql.connector
    try:
        conn = mysql.connector.connect(
            host=os.getenv("MYSQL_HOST", "localhost"),
            port=int(os.getenv("MYSQL_PORT", 3306)),
            user=os.getenv("MYSQL_USER", "root"),
            password=os.getenv("MYSQL_PASSWORD", ""),
            database=os.getenv("MYSQL_DATABASE", "cotizaciones_imss"),
            connection_timeout=5,
        )
        df = pd.read_sql("""
            SELECT asunto, remitente, remitente_imss, fecha_correo, estado,
                   filas_detectadas, filas_insertadas, expediente,
                   intentos, error_mensaje, actualizado_en
            FROM correos_procesados
            ORDER BY fecha_correo DESC
        """, conn)
        conn.close()
    except Exception:
        return pd.DataFrame()
    df["fecha_correo"]         = pd.to_datetime(df["fecha_correo"])
    df["actualizado_en"]       = pd.to_datetime(df["actualizado_en"])
    df["fecha"]                = df["fecha_correo"].dt.date
    df["remitente_imss"]       = df["remitente_imss"].fillna(df["remitente"])
    df["remitente_corto_imss"] = df["remitente_imss"].apply(
        lambda x: x.split("@")[0] if "@" in str(x) else str(x)
    )
    return df


# ── Skeleton mientras carga ────────────────────────────────────────────────────
skel = st.empty()
skel.markdown("""
<div style="padding:.4rem 0">
    <div class="skel" style="height:100px;margin-bottom:1.4rem;border-radius:16px"></div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.9rem;margin-bottom:1.4rem">
        <div class="skel" style="height:105px;border-radius:14px"></div>
        <div class="skel" style="height:105px;border-radius:14px"></div>
        <div class="skel" style="height:105px;border-radius:14px"></div>
        <div class="skel" style="height:105px;border-radius:14px"></div>
    </div>
    <div class="skel" style="height:48px;margin-bottom:.9rem;border-radius:10px"></div>
    <div class="skel" style="height:300px;border-radius:12px"></div>
</div>
""", unsafe_allow_html=True)

df_all = cargar_datos()
skel.empty()


# =============================================================================
#  SIDEBAR
# =============================================================================

_es_admin_actual = _es_admin(st.session_state.get("auth_user", ""))

with st.sidebar:
    st.markdown(
        f"<div style='color:#475569;font-size:.62rem;font-weight:700;text-transform:uppercase;"
        f"letter-spacing:.12em;padding:.6rem .4rem .4rem'>Navegación</div>",
        unsafe_allow_html=True,
    )

    _nav_options = [
        "🏠  Inicio",
        "📈  Tendencias",
        "🏆  Top Productos",
        "📋  Registros",
        "📬  Cotizaciones",
        "💡  Resumen Ejecutivo",
        "👤  Análisis Remitentes",
        "📄  Cotizaciones PDF",
        "📥  Correos Procesados",
    ]
    if _es_admin_actual:
        _nav_options.append("⚙️  Usuarios")

    nav_page = st.radio("nav", _nav_options, label_visibility="collapsed", key="nav_radio")

    st.markdown("<hr style='border-color:#1e293b;margin:.6rem 0'>", unsafe_allow_html=True)

    # Estado del sistema
    if df_all.empty:
        st.error("Sin conexión MySQL")
    else:
        st.markdown(
            "<p class='pulse' style='color:#4ade80;font-size:.8rem;font-weight:600;margin:.2rem 0'>"
            "● Sistema activo</p>"
            f"<p style='color:#334155;font-size:.68rem;margin:0'>Última carga: "
            f"{datetime.now().strftime('%H:%M:%S')}</p>",
            unsafe_allow_html=True,
        )

    if st.button("↻  Actualizar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown(
        "<p style='color:#334155;font-size:.65rem;text-align:center;margin-top:.3rem'>"
        "Auto-actualización cada 2 min</p>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='sidebar-spacer'></div>", unsafe_allow_html=True)
    st.markdown("<hr style='border-color:#1e293b;margin:.4rem 0'>", unsafe_allow_html=True)

    nombre_s  = st.session_state.get("auth_name", "Usuario")
    usuario_s = st.session_state.get("auth_user", "")
    rol_badge = "ADMIN" if _es_admin_actual else "USER"
    st.markdown(
        f"<div style='background:#1e293b;border-radius:10px;padding:.7rem 1rem;margin:.3rem 0 .6rem'>"
        f"<div style='color:#475569;font-size:.6rem;font-weight:700;text-transform:uppercase;"
        f"letter-spacing:.1em;margin-bottom:.25rem'>Sesión activa</div>"
        f"<div style='color:#f1f5f9;font-size:.83rem;font-weight:600;margin-bottom:.1rem'>{nombre_s}</div>"
        f"<div style='display:flex;align-items:center;gap:.4rem'>"
        f"<span style='color:#475569;font-size:.67rem'>{usuario_s}</span>"
        f"<span style='background:rgba(37,99,235,0.3);color:#93c5fd;font-size:.55rem;"
        f"font-weight:700;padding:.1rem .35rem;border-radius:3px'>{rol_badge}</span></div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    if st.button("Cerrar sesión", use_container_width=True):
        st.session_state.update({"auth_ok": False, "auth_user": "", "auth_name": ""})
        st.rerun()


# =============================================================================
#  GUARD: Sin datos de MySQL
# =============================================================================

if df_all.empty:
    err = st.session_state.get("db_error", "")
    alerta_err(
        "Sin conexión a la base de datos",
        f"No se pudo cargar información. Verifica que MySQL esté activo y el script esté corriendo."
        + (f"<br><code style='font-size:.72rem'>{err}</code>" if err else ""),
    )
    st.stop()


# =============================================================================
#  CÁLCULOS GLOBALES
# =============================================================================

hoy       = datetime.now().date()
sem_ini   = hoy - timedelta(days=7)
sem_ant_i = hoy - timedelta(days=14)

df_hoy     = df_all[df_all["fecha"] == hoy]
df_sem     = df_all[df_all["fecha"] >= sem_ini]
df_sem_ant = df_all[(df_all["fecha"] >= sem_ant_i) & (df_all["fecha"] < sem_ini)]
delta_sem  = len(df_sem) - len(df_sem_ant)

# Delta vs ayer
ayer       = hoy - timedelta(days=1)
df_ayer    = df_all[df_all["fecha"] == ayer]
delta_hoy  = len(df_hoy) - len(df_ayer)

# Top productos (disponible globalmente)
top_prod = (
    df_all.groupby(["gpo", "gen", "esp"])
    .agg(cotizaciones=("clave_imss", "count"), total_piezas=("cantidad", "sum"))
    .sort_values("cotizaciones", ascending=False)
    .head(15)
    .reset_index()
)
top_prod["total_piezas"] = top_prod["total_piezas"].fillna(0).astype(int)
top_prod.columns = ["GPO", "GEN", "ESP", "# Cotizaciones", "Total Piezas"]


# =============================================================================
#  HEADER GLOBAL
# =============================================================================

h_left, h_right = st.columns([3, 1])
with h_left:
    st.markdown(
        f"<div style='margin-bottom:.1rem'>"
        f"<span style='font-size:1.5rem;font-weight:800;color:{NAVY};letter-spacing:-.03em'>"
        f"Monitor de Cotizaciones IMSS</span></div>"
        f"<div style='font-size:.78rem;color:{MUTED};font-weight:400'>"
        f"GADMAR SA DE CV &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        f"&nbsp;·&nbsp; <strong style='color:{NAVY}'>{len(df_all):,}</strong> registros</div>",
        unsafe_allow_html=True,
    )
with h_right:
    st.markdown(
        f"<div style='display:flex;gap:1.4rem;justify-content:flex-end;align-items:center;"
        f"height:100%;padding-top:.4rem'>"
        f"<div style='text-align:center'>"
        f"<div style='font-size:1.5rem;font-weight:800;color:{NAVY}'>{len(df_hoy):,}</div>"
        f"<div style='font-size:.62rem;color:{MUTED};text-transform:uppercase;letter-spacing:.08em'>Hoy</div>"
        f"</div>"
        f"<div style='width:1px;height:32px;background:rgba(0,0,0,0.1)'></div>"
        f"<div style='text-align:center'>"
        f"<div style='font-size:1.5rem;font-weight:800;color:{NAVY}'>{len(df_sem):,}</div>"
        f"<div style='font-size:.62rem;color:{MUTED};text-transform:uppercase;letter-spacing:.08em'>Semana</div>"
        f"</div>"
        f"<div style='width:1px;height:32px;background:rgba(0,0,0,0.1)'></div>"
        f"<div style='text-align:center'>"
        f"<div style='font-size:1.5rem;font-weight:800;color:{NAVY}'>{len(df_all):,}</div>"
        f"<div style='font-size:.62rem;color:{MUTED};text-transform:uppercase;letter-spacing:.08em'>Total</div>"
        f"</div></div>",
        unsafe_allow_html=True,
    )

st.markdown(
    "<hr style='border:none;border-top:1px solid rgba(0,0,0,0.08);margin:.8rem 0 1rem'>",
    unsafe_allow_html=True,
)

# Alerta si no hay datos hoy
if df_hoy.empty:
    alerta_warn(
        "Sin cotizaciones hoy",
        f"No se ha registrado ninguna partida el <strong>{hoy.strftime('%d/%m/%Y')}</strong>. "
        "Verifica que el script de correos esté activo y con conexión a MySQL.",
    )

# ── KPIs globales ─────────────────────────────────────────────────────────────
k1, k2, k3, k4 = st.columns(4)
d_dir = "up" if delta_sem > 0 else ("down" if delta_sem < 0 else None)
d_txt = f"{delta_sem:+,} vs sem. ant." if delta_sem else "Sin cambio"
h_dir = "up" if delta_hoy > 0 else ("down" if delta_hoy < 0 else None)
h_txt = f"{delta_hoy:+,} vs ayer" if delta_hoy else "Sin cambio"

with k1: st.markdown(kpi_card("Claves hoy",       f"{len(df_hoy):,}",               "📅", BLUE, h_txt, h_dir), unsafe_allow_html=True)
with k2: st.markdown(kpi_card("Esta semana",        f"{len(df_sem):,}",               "📊", SKY,  d_txt, d_dir), unsafe_allow_html=True)
with k3: st.markdown(kpi_card("Total histórico",    f"{len(df_all):,}",               "📁", GREEN), unsafe_allow_html=True)
with k4: st.markdown(kpi_card("Remitentes activos", f"{df_all['remitente'].nunique()}", "👤", AMBER), unsafe_allow_html=True)

st.markdown("<div style='height:1.2rem'></div>", unsafe_allow_html=True)


# =============================================================================
#  DIÁLOGO: VISTA PREVIA PDF
# =============================================================================

@st.dialog("Vista previa del PDF", width="large")
def _ver_pdf_dialog(path: str, nombre: str, expediente: str):
    st.markdown(
        f"<div style='font-size:.78rem;color:{MUTED};margin-bottom:.6rem'>"
        f"<strong style='color:{NAVY}'>{expediente}</strong> &nbsp;·&nbsp; {nombre}</div>",
        unsafe_allow_html=True,
    )
    try:
        with open(path, "rb") as fh:
            pdf_bytes = fh.read()
        b64 = base64.b64encode(pdf_bytes).decode()
        st.markdown(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            f'width="100%" height="720px" '
            f'style="border:1px solid rgba(0,0,0,0.08);border-radius:8px"></iframe>',
            unsafe_allow_html=True,
        )
        st.download_button(
            "⬇  Descargar PDF", pdf_bytes,
            file_name=nombre, mime="application/pdf",
            use_container_width=True, key=f"dl_dialog_{nombre}",
        )
    except FileNotFoundError:
        st.error(f"No se encontró el archivo: {path}")


# =============================================================================
#  PÁGINAS
# =============================================================================

# ── INICIO (Dashboard principal) ──────────────────────────────────────────────
if nav_page == "🏠  Inicio":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    ch1, ch2 = st.columns([3, 2], gap="medium")

    with ch1:
        st.markdown('<p class="sec">Volumen · últimos 30 días</p>', unsafe_allow_html=True)
        vol = (df_all[df_all["fecha"] >= hoy - timedelta(days=30)]
               .groupby("fecha").size().reset_index(name="Claves"))
        prom_line = vol["Claves"].mean()
        fig_home = go.Figure()
        fig_home.add_trace(go.Bar(
            x=vol["fecha"], y=vol["Claves"],
            marker_color=bar_colors(len(vol)), marker_line_width=0,
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y:,} claves<extra></extra>",
            name="Claves",
        ))
        fig_home.add_trace(go.Scatter(
            x=vol["fecha"], y=[prom_line] * len(vol),
            mode="lines", line=dict(color=AMBER, width=1.8, dash="dot"),
            hovertemplate=f"Promedio: {prom_line:.1f}<extra></extra>",
            name=f"Promedio ({prom_line:.1f})",
        ))
        fig_home.update_layout(**plotly_layout(
            height=280, bargap=.28, xaxis=dict(tickformat="%d %b"), yaxis_title="Claves",
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        ))
        st.plotly_chart(fig_home, use_container_width=True)

    with ch2:
        st.markdown('<p class="sec">Top 5 productos</p>', unsafe_allow_html=True)
        t5 = top_prod.head(5).copy()
        t5["Clave"] = t5["GPO"] + "-" + t5["GEN"] + "-" + t5["ESP"]
        fig_t5 = go.Figure(go.Bar(
            x=t5["# Cotizaciones"], y=t5["Clave"], orientation="h",
            marker_color=bar_colors(len(t5), (14, 165, 233)), marker_line_width=0,
            text=t5["# Cotizaciones"], textposition="outside",
            hovertemplate="<b>%{y}</b><br>%{x:,}<extra></extra>",
        ))
        fig_t5.update_layout(**plotly_layout(
            height=280, xaxis_title="Cotizaciones", yaxis=dict(autorange="reversed"),
        ))
        st.plotly_chart(fig_t5, use_container_width=True)

    # Insights operativos
    st.markdown('<p class="sec">Resumen operativo</p>', unsafe_allow_html=True)
    i1, i2, i3 = st.columns(3)

    dias_act  = df_all["fecha"].nunique()
    prom_dia  = round(len(df_all) / max(dias_act, 1), 1)
    rem_top   = df_all["remitente_corto"].value_counts().idxmax() if not df_all.empty else "—"
    rem_top_v = int(df_all["remitente_corto"].value_counts().max()) if not df_all.empty else 0

    if len(df_sem_ant) > 0:
        cp = round((len(df_sem) - len(df_sem_ant)) / len(df_sem_ant) * 100, 1)
        tag_s = (f'<span class="tag tag-up">▲ +{cp}%</span>' if cp > 0
                 else f'<span class="tag tag-down">▼ {cp}%</span>' if cp < 0
                 else '<span class="tag tag-flat">Sin cambio</span>')
    else:
        tag_s = '<span class="tag tag-flat">Sin datos previos</span>'

    with i1:
        st.markdown(
            f'<div class="insight-card">'
            f'<div class="insight-num">{prom_dia:,.1f}</div>'
            f'<div class="insight-label">Promedio diario de claves</div>'
            f'<div class="insight-text">{tag_s} vs semana anterior.<br>'
            f'<strong>{dias_act}</strong> días con actividad registrada.</div>'
            f'</div>', unsafe_allow_html=True)
    with i2:
        prod_top = top_prod.iloc[0] if not top_prod.empty else None
        if prod_top is not None:
            st.markdown(
                f'<div class="insight-card">'
                f'<div class="insight-num">{prod_top["GPO"]}-{prod_top["GEN"]}-{prod_top["ESP"]}</div>'
                f'<div class="insight-label">Producto más solicitado</div>'
                f'<div class="insight-text">'
                f'<strong>{prod_top["# Cotizaciones"]:,}</strong> cotizaciones &nbsp;·&nbsp; '
                f'<strong>{prod_top["Total Piezas"]:,}</strong> piezas totales.</div>'
                f'</div>', unsafe_allow_html=True)
    with i3:
        st.markdown(
            f'<div class="insight-card">'
            f'<div class="insight-num">{rem_top}</div>'
            f'<div class="insight-label">Remitente más activo</div>'
            f'<div class="insight-text">'
            f'<strong>{rem_top_v:,}</strong> claves enviadas en total.<br>'
            f'<strong>{df_all["remitente"].nunique()}</strong> remitentes registrados.</div>'
            f'</div>', unsafe_allow_html=True)


# ── TENDENCIAS ────────────────────────────────────────────────────────────────
elif nav_page == "📈  Tendencias":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    # Filtro de rango de fechas
    with st.expander("🗓  Filtrar rango de fechas", expanded=False):
        fc1, fc2 = st.columns(2)
        with fc1:
            fecha_desde = st.date_input("Desde", value=hoy - timedelta(days=30), key="tend_desde")
        with fc2:
            fecha_hasta = st.date_input("Hasta", value=hoy, key="tend_hasta")
    df_tend = df_all[(df_all["fecha"] >= fecha_desde) & (df_all["fecha"] <= fecha_hasta)]

    if df_tend.empty:
        st.info("No hay datos para el rango seleccionado.")
    else:
        cv, cr = st.columns([3, 2], gap="medium")
        with cv:
            st.markdown('<p class="sec">Volumen de claves por día</p>', unsafe_allow_html=True)
            vol = df_tend.groupby("fecha").size().reset_index(name="Claves")
            fig1 = go.Figure(go.Bar(
                x=vol["fecha"], y=vol["Claves"],
                marker_color=bar_colors(len(vol)), marker_line_width=0,
                hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y:,} claves<extra></extra>",
            ))
            fig1.update_layout(**plotly_layout(height=290, bargap=.28, xaxis=dict(tickformat="%d %b"), yaxis_title="Claves"))
            st.plotly_chart(fig1, use_container_width=True)

        with cr:
            st.markdown('<p class="sec">Top remitentes</p>', unsafe_allow_html=True)
            tr = df_tend["remitente_corto"].value_counts().head(6).reset_index()
            tr.columns = ["Remitente", "Claves"]
            fig2 = go.Figure(go.Bar(
                x=tr["Claves"], y=tr["Remitente"], orientation="h",
                marker_color=bar_colors(len(tr), (14, 165, 233)), marker_line_width=0,
                text=tr["Claves"], textposition="outside",
                hovertemplate="<b>%{y}</b><br>%{x:,}<extra></extra>",
            ))
            fig2.update_layout(**plotly_layout(height=290, xaxis_title="Claves", yaxis=dict(autorange="reversed")))
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown('<p class="sec">Tendencia acumulada</p>', unsafe_allow_html=True)
        acum = df_tend.groupby("fecha").size().reset_index(name="P").sort_values("fecha")
        acum["Acum"] = acum["P"].cumsum()
        fig3 = go.Figure(go.Scatter(
            x=acum["fecha"], y=acum["Acum"], mode="lines",
            line=dict(color=BLUE, width=2.5),
            fill="tozeroy", fillcolor="rgba(37,99,235,0.06)",
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y:,} acumuladas<extra></extra>",
        ))
        fig3.update_layout(**plotly_layout(height=210, xaxis=dict(tickformat="%d %b"), yaxis_title="Acumulado"))
        st.plotly_chart(fig3, use_container_width=True)


# ── TOP PRODUCTOS ─────────────────────────────────────────────────────────────
elif nav_page == "🏆  Top Productos":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)
    ct, cp = st.columns([3, 2], gap="medium")

    with ct:
        st.markdown('<p class="sec">Claves más frecuentes (Top 15)</p>', unsafe_allow_html=True)
        st.dataframe(top_prod, use_container_width=True, hide_index=True, height=420,
            column_config={
                "# Cotizaciones": st.column_config.ProgressColumn(
                    "# Cotizaciones", min_value=0,
                    max_value=int(top_prod["# Cotizaciones"].max()), format="%d"),
                "Total Piezas": st.column_config.NumberColumn("Total Piezas", format="%d"),
            })
    with cp:
        st.markdown('<p class="sec">Distribución top 8</p>', unsafe_allow_html=True)
        pd8 = top_prod.head(8).copy()
        pd8["Clave"] = pd8["GPO"] + "-" + pd8["GEN"] + "-" + pd8["ESP"]
        fig4 = go.Figure(go.Pie(
            labels=pd8["Clave"], values=pd8["# Cotizaciones"], hole=.52,
            textposition="outside", textinfo="label+percent",
            hovertemplate="<b>%{label}</b><br>%{value}<br>%{percent}<extra></extra>",
            marker=dict(colors=px.colors.qualitative.Bold[:len(pd8)]),
        ))
        fig4.update_layout(**plotly_layout(height=380, margin=dict(l=10, r=10, t=30, b=10)))
        st.plotly_chart(fig4, use_container_width=True)


# ── REGISTROS ─────────────────────────────────────────────────────────────────
elif nav_page == "📋  Registros":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    # Filtros
    with st.expander("🔍  Filtros", expanded=False):
        rf1, rf2, rf3 = st.columns(3)
        with rf1:
            busq_rem_r = st.text_input("Remitente", placeholder="compras@...", key="reg_rem")
        with rf2:
            busq_gpo_r = st.text_input("GPO", placeholder="010", key="reg_gpo")
        with rf3:
            busq_esp_r = st.text_input("ESP", placeholder="001", key="reg_esp")
        rd1, rd2 = st.columns(2)
        with rd1:
            reg_desde = st.date_input("Desde", value=hoy - timedelta(days=30), key="reg_desde")
        with rd2:
            reg_hasta = st.date_input("Hasta", value=hoy, key="reg_hasta")

    df_reg = df_all.copy()
    if busq_rem_r:
        df_reg = df_reg[df_reg["remitente"].str.contains(busq_rem_r, case=False, na=False)]
    if busq_gpo_r:
        df_reg = df_reg[df_reg["gpo"].astype(str).str.contains(busq_gpo_r, case=False, na=False)]
    if busq_esp_r:
        df_reg = df_reg[df_reg["esp"].astype(str).str.contains(busq_esp_r, case=False, na=False)]
    df_reg = df_reg[(df_reg["fecha"] >= reg_desde) & (df_reg["fecha"] <= reg_hasta)]

    ci, cd = st.columns([3, 1])
    with ci:
        st.caption(f"{len(df_reg):,} registros (de {len(df_all):,} totales)")
    with cd:
        csv_b = (df_reg[["fecha_correo", "remitente", "gpo", "gen", "esp", "dif", "var", "cantidad", "asunto"]]
                 .to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"))
        st.download_button("⬇  Exportar CSV", csv_b,
            f"cotizaciones_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv", use_container_width=True)

    df_disp = (df_reg[["fecha_correo", "remitente_corto", "gpo", "gen", "esp", "dif", "var", "cantidad", "asunto"]]
               .rename(columns={"fecha_correo": "Fecha", "remitente_corto": "Remitente",
                                "gpo": "GPO", "gen": "GEN", "esp": "ESP", "dif": "DIF", "var": "VAR",
                                "cantidad": "Cantidad", "asunto": "Asunto"})
               .reset_index(drop=True))

    total_regs = len(df_disp)
    pc1, pc2, pc3 = st.columns([1.4, 1.6, 2])
    with pc1:
        por_pagina = st.selectbox("Registros por página", [10, 25, 50, 100], index=1, key="tab3_por_pagina")
    total_pags = max(1, -(-total_regs // por_pagina))

    _fsig = f"{busq_rem_r}|{busq_gpo_r}|{busq_esp_r}|{reg_desde}|{reg_hasta}|{por_pagina}"
    if st.session_state.get("_tab3_fsig") != _fsig:
        st.session_state["_tab3_fsig"] = _fsig
        st.session_state["tab3_pagina"] = 1

    with pc2:
        pagina = st.number_input("Página", min_value=1, max_value=total_pags,
            value=st.session_state.get("tab3_pagina", 1), step=1, key="tab3_pagina")
    with pc3:
        st.markdown(
            f"<div style='padding-top:1.85rem;color:{MUTED};font-size:.78rem'>"
            f"de <strong>{total_pags}</strong> páginas &nbsp;·&nbsp; "
            f"<strong>{total_regs:,}</strong> registros</div>",
            unsafe_allow_html=True)

    inicio  = (pagina - 1) * por_pagina
    df_pag  = df_disp.iloc[inicio: inicio + por_pagina]
    st.dataframe(df_pag.style.apply(zebra_style, axis=None),
        use_container_width=True, height=480, hide_index=True,
        column_config={
            "Fecha":    st.column_config.DatetimeColumn("Fecha", format="DD/MM/YYYY HH:mm"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", format="%d"),
        })

    bp, _, bn = st.columns([1, 4, 1])
    with bp:
        if pagina > 1:
            if st.button("← Anterior", use_container_width=True, key="tab3_prev"):
                st.session_state["tab3_pagina"] = pagina - 1
                st.rerun()
    with bn:
        if pagina < total_pags:
            if st.button("Siguiente →", use_container_width=True, key="tab3_next"):
                st.session_state["tab3_pagina"] = pagina + 1
                st.rerun()


# ── COTIZACIONES ──────────────────────────────────────────────────────────────
elif nav_page == "📬  Cotizaciones":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    df_cot = df_all.drop_duplicates(subset=["fecha", "remitente", "asunto"]).copy()
    df_cot["hora"] = df_cot["fecha_correo"].dt.hour
    claves_x_cot = (df_all.groupby(["fecha", "remitente", "asunto"])["clave_imss"]
                    .count().reset_index(name="n_claves"))
    df_cot = df_cot.merge(claves_x_cot, on=["fecha", "remitente", "asunto"], how="left")

    hoy_c    = datetime.now().date()
    cot_hoy  = df_cot[df_cot["fecha"] == hoy_c]
    cot_sem  = df_cot[df_cot["fecha"] >= hoy_c - timedelta(days=7)]
    dias_act = df_cot["fecha"].nunique()
    prom_dia = round(len(df_cot) / max(dias_act, 1), 1)
    rem_hoy  = cot_hoy["remitente"].nunique()

    ck1, ck2, ck3, ck4, ck5 = st.columns(5)
    with ck1: st.markdown(kpi_card("Cotizaciones hoy",  f"{len(cot_hoy):,}",  "📬", BLUE),  unsafe_allow_html=True)
    with ck2: st.markdown(kpi_card("Esta semana",        f"{len(cot_sem):,}",  "📅", SKY),   unsafe_allow_html=True)
    with ck3: st.markdown(kpi_card("Total histórico",    f"{len(df_cot):,}",   "📁", GREEN),  unsafe_allow_html=True)
    with ck4: st.markdown(kpi_card("Promedio / día",     f"{prom_dia}",        "📊", AMBER),  unsafe_allow_html=True)
    with ck5: st.markdown(kpi_card("Remitentes hoy",     f"{rem_hoy}",         "👤", NAVY),   unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    gc1, gc2 = st.columns([3, 2], gap="medium")
    with gc1:
        st.markdown('<p class="sec">Cotizaciones por día — últimos 30 días</p>', unsafe_allow_html=True)
        vol_c    = (df_cot[df_cot["fecha"] >= hoy_c - timedelta(days=30)]
                    .groupby("fecha").size().reset_index(name="Cotizaciones"))
        prom_line = vol_c["Cotizaciones"].mean()
        fig_vc   = go.Figure()
        fig_vc.add_trace(go.Bar(
            x=vol_c["fecha"], y=vol_c["Cotizaciones"],
            marker_color=bar_colors(len(vol_c)), marker_line_width=0,
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y} cotizaciones<extra></extra>",
            name="Cotizaciones",
        ))
        fig_vc.add_trace(go.Scatter(
            x=vol_c["fecha"], y=[prom_line] * len(vol_c),
            mode="lines", line=dict(color=AMBER, width=1.8, dash="dot"),
            hovertemplate=f"Promedio: {prom_line:.1f}<extra></extra>",
            name=f"Promedio ({prom_line:.1f})",
        ))
        fig_vc.update_layout(**plotly_layout(
            height=300, bargap=.28, xaxis=dict(tickformat="%d %b"), yaxis_title="Cotizaciones",
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        ))
        st.plotly_chart(fig_vc, use_container_width=True)

    with gc2:
        st.markdown('<p class="sec">Distribución por hora del día</p>', unsafe_allow_html=True)
        hora_dist = df_cot.groupby("hora").size().reset_index(name="Cotizaciones")
        hora_dist["hora_lbl"] = hora_dist["hora"].apply(lambda h: f"{h:02d}:00")
        fig_hora = go.Figure(go.Bar(
            x=hora_dist["hora_lbl"], y=hora_dist["Cotizaciones"],
            marker_color=bar_colors(len(hora_dist), (14, 165, 233)), marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>%{y} cotizaciones<extra></extra>",
        ))
        fig_hora.update_layout(**plotly_layout(height=300, bargap=.18, xaxis_title="Hora", yaxis_title="Cotizaciones"))
        st.plotly_chart(fig_hora, use_container_width=True)

    st.markdown('<p class="sec">Cotizaciones por remitente — últimos 30 días</p>', unsafe_allow_html=True)
    rem_dia = (df_cot[df_cot["fecha"] >= hoy_c - timedelta(days=30)]
               .groupby(["fecha", "remitente_corto"]).size().reset_index(name="Cotizaciones"))
    n_rem = rem_dia["remitente_corto"].nunique()
    pal_r = (px.colors.qualitative.Bold if n_rem <= 10 else px.colors.qualitative.Alphabet)[:n_rem]
    fig_rd = px.line(rem_dia.sort_values("fecha"), x="fecha", y="Cotizaciones",
                     color="remitente_corto", markers=True,
                     template="plotly_white", color_discrete_sequence=pal_r)
    fig_rd.update_layout(**plotly_layout(
        height=260, xaxis_title="", yaxis_title="Cotizaciones", showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        xaxis=dict(tickformat="%d %b"),
    ))
    fig_rd.update_traces(marker_size=5, line_width=2)
    st.plotly_chart(fig_rd, use_container_width=True)

    gr1, gr2 = st.columns([2, 3], gap="medium")
    with gr1:
        st.markdown('<p class="sec">Ranking de remitentes</p>', unsafe_allow_html=True)
        rank_r = df_cot["remitente_corto"].value_counts().head(10).reset_index()
        rank_r.columns = ["Remitente", "Cotizaciones"]
        fig_rk = go.Figure(go.Bar(
            x=rank_r["Cotizaciones"], y=rank_r["Remitente"], orientation="h",
            marker_color=bar_colors(len(rank_r), (5, 150, 105)), marker_line_width=0,
            text=rank_r["Cotizaciones"], textposition="outside",
            hovertemplate="<b>%{y}</b><br>%{x} cotizaciones<extra></extra>",
        ))
        fig_rk.update_layout(**plotly_layout(height=300, xaxis_title="Cotizaciones", yaxis=dict(autorange="reversed")))
        st.plotly_chart(fig_rk, use_container_width=True)

    with gr2:
        st.markdown('<p class="sec">Mapa de calor — hora × día de la semana</p>', unsafe_allow_html=True)
        df_cot["dow"] = df_cot["fecha_correo"].dt.dayofweek
        dias_nombres  = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        heat  = df_cot.groupby(["dow", "hora"]).size().reset_index(name="n")
        heat["dia_nombre"] = heat["dow"].map(lambda d: dias_nombres[d])
        pivot = heat.pivot(index="dia_nombre", columns="hora", values="n").fillna(0)
        pivot = pivot.reindex([d for d in dias_nombres if d in pivot.index])
        fig_heat = go.Figure(go.Heatmap(
            z=pivot.values,
            x=[f"{h:02d}h" for h in pivot.columns],
            y=pivot.index.tolist(),
            colorscale=[[0, "#f0f9ff"], [0.5, "#38bdf8"], [1, "#0369a1"]],
            hovertemplate="<b>%{y}</b> %{x}<br>%{z} cotizaciones<extra></extra>",
        ))
        fig_heat.update_layout(**plotly_layout(height=300, xaxis_title="Hora del día", margin=dict(l=0, r=8, t=6, b=0)))
        st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown('<p class="sec">Detalle de cotizaciones recientes</p>', unsafe_allow_html=True)
    df_tabla_cot = (df_cot[["fecha_correo", "remitente_corto", "asunto", "hora", "n_claves"]]
                    .sort_values("fecha_correo", ascending=False)
                    .rename(columns={"fecha_correo": "Fecha y hora", "remitente_corto": "Remitente",
                                     "asunto": "Asunto", "hora": "Hora", "n_claves": "# Claves"})
                    .reset_index(drop=True))
    st.dataframe(df_tabla_cot.style.apply(zebra_style, axis=None),
        use_container_width=True, height=420, hide_index=True,
        column_config={
            "Fecha y hora": st.column_config.DatetimeColumn("Fecha y hora", format="DD/MM/YYYY HH:mm"),
            "# Claves": st.column_config.NumberColumn("# Claves", format="%d"),
            "Hora":     st.column_config.NumberColumn("Hora", format="%d:00"),
        })


# ── RESUMEN EJECUTIVO ─────────────────────────────────────────────────────────
elif nav_page == "💡  Resumen Ejecutivo":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    prom_d   = round(len(df_all) / max(df_all["fecha"].nunique(), 1), 1)
    mej_dia  = df_all.groupby("fecha").size().idxmax()
    mej_val  = df_all.groupby("fecha").size().max()
    prod_top = top_prod.iloc[0] if not top_prod.empty else None
    rem_top  = df_all["remitente_corto"].value_counts().idxmax() if not df_all.empty else "—"
    rem_top_v = df_all["remitente_corto"].value_counts().max() if not df_all.empty else 0

    if len(df_sem_ant) > 0:
        cp = round((len(df_sem) - len(df_sem_ant)) / len(df_sem_ant) * 100, 1)
        tag_s = (f'<span class="tag tag-up">▲ +{cp}%</span>' if cp > 0
                 else f'<span class="tag tag-down">▼ {cp}%</span>' if cp < 0
                 else '<span class="tag tag-flat">Sin cambio</span>')
    else:
        tag_s = '<span class="tag tag-flat">Sin datos previos</span>'

    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown(
            f'<div class="insight-card"><div class="insight-num">{len(df_sem):,}</div>'
            f'<div class="insight-label">Claves esta semana</div>'
            f'<div class="insight-text">{tag_s} vs {len(df_sem_ant):,} claves semana anterior.</div>'
            f'</div>', unsafe_allow_html=True)
    with r2:
        st.markdown(
            f'<div class="insight-card"><div class="insight-num">{prom_d:,.1f}</div>'
            f'<div class="insight-label">Promedio diario</div>'
            f'<div class="insight-text">Día más activo: <strong>{mej_dia.strftime("%d/%m/%Y")}</strong>'
            f' con <strong>{mej_val:,}</strong> claves.</div></div>', unsafe_allow_html=True)
    with r3:
        st.markdown(
            f'<div class="insight-card"><div class="insight-num">{df_all["remitente"].nunique()}</div>'
            f'<div class="insight-label">Remitentes registrados</div>'
            f'<div class="insight-text">El más activo: <strong>{rem_top}</strong>'
            f' con <strong>{rem_top_v:,}</strong> claves.</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    r4, r5 = st.columns(2)
    with r4:
        if prod_top is not None:
            st.markdown(
                f'<div class="insight-card">'
                f'<div class="insight-num">{prod_top["GPO"]}-{prod_top["GEN"]}-{prod_top["ESP"]}</div>'
                f'<div class="insight-label">Producto más solicitado</div>'
                f'<div class="insight-text">En <strong>{prod_top["# Cotizaciones"]:,} cotizaciones</strong>'
                f' · <strong>{prod_top["Total Piezas"]:,} piezas</strong> totales.</div>'
                f'</div>', unsafe_allow_html=True)
    with r5:
        dias_a = df_all["fecha"].nunique()
        p_c    = df_all["fecha_correo"].min().strftime("%d/%m/%Y")
        u_c    = df_all["fecha_correo"].max().strftime("%d/%m/%Y")
        st.markdown(
            f'<div class="insight-card"><div class="insight-num">{dias_a}</div>'
            f'<div class="insight-label">Días con actividad</div>'
            f'<div class="insight-text">Desde <strong>{p_c}</strong> hasta <strong>{u_c}</strong>'
            f' — <strong>{len(df_all):,} claves</strong> capturadas automáticamente.</div>'
            f'</div>', unsafe_allow_html=True)

    st.markdown(
        f"<p style='color:#d1d5db;font-size:.68rem;text-align:center;margin-top:1.5rem'>"
        "GADMAR SA DE CV &nbsp;·&nbsp; Monitor Automático de Cotizaciones IMSS &nbsp;·&nbsp; v4.0</p>",
        unsafe_allow_html=True)


# ── ANÁLISIS REMITENTES ───────────────────────────────────────────────────────
elif nav_page == "👤  Análisis Remitentes":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    envios_dia = (df_all.groupby(["remitente", "fecha"])["asunto"].nunique().reset_index(name="envios"))
    resumen = (df_all.groupby("remitente")
               .agg(total_claves=("clave_imss", "count"), claves_unicas=("clave_imss", "nunique"),
                    dias_activos=("fecha", "nunique"), total_envios=("asunto", "nunique"))
               .reset_index())
    resumen["promedio_envios_dia"] = (resumen["total_envios"] / resumen["dias_activos"]).round(2)
    gpo_top = (df_all.groupby(["remitente", "gpo"]).size().reset_index(name="n")
               .sort_values("n", ascending=False).drop_duplicates("remitente")
               .rename(columns={"gpo": "gpo_favorito", "n": "veces_gpo"}))
    resumen = resumen.merge(gpo_top[["remitente", "gpo_favorito", "veces_gpo"]], on="remitente", how="left")
    resumen = resumen.sort_values("total_envios", ascending=False)
    resumen["remitente_corto"] = resumen["remitente"].apply(lambda x: x.split("@")[0] if "@" in str(x) else str(x))

    top_r  = resumen.iloc[0]
    p_env  = round(envios_dia["envios"].mean(), 1)
    max_d  = envios_dia.sort_values("envios", ascending=False).iloc[0]

    m1, m2, m3, m4 = st.columns(4)
    with m1: st.markdown(kpi_card("Remitente más activo", top_r["remitente_corto"], "🥇", BLUE, f"{int(top_r['total_envios'])} cotizaciones"), unsafe_allow_html=True)
    with m2: st.markdown(kpi_card("Promedio envíos/día",  f"{p_env}",               "📤", SKY),  unsafe_allow_html=True)
    with m3: st.markdown(kpi_card("Claves únicas",        f"{df_all['clave_imss'].nunique():,}", "🔑", GREEN), unsafe_allow_html=True)
    with m4: st.markdown(kpi_card("Día más activo",       max_d["fecha"].strftime("%d/%m/%Y"),   "📆", AMBER, f"{int(max_d['envios'])} envíos"), unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    st.markdown('<p class="sec">Resumen por remitente</p>', unsafe_allow_html=True)

    t_res = (resumen[["remitente", "total_envios", "dias_activos", "promedio_envios_dia", "claves_unicas", "gpo_favorito", "veces_gpo"]]
             .rename(columns={"remitente": "Remitente", "total_envios": "Total Cotizaciones",
                              "dias_activos": "Días Activos", "promedio_envios_dia": "Prom. Envíos/Día",
                              "claves_unicas": "Claves Únicas", "gpo_favorito": "GPO Favorito", "veces_gpo": "Veces GPO"})
             .reset_index(drop=True))
    st.dataframe(t_res.style.apply(zebra_style, axis=None), use_container_width=True, hide_index=True, height=270,
        column_config={
            "Total Cotizaciones": st.column_config.ProgressColumn("Total Cotizaciones",
                min_value=0, max_value=int(t_res["Total Cotizaciones"].max()), format="%d"),
            "Claves Únicas": st.column_config.NumberColumn("Claves Únicas", format="%d"),
            "Veces GPO":     st.column_config.NumberColumn("Veces GPO",     format="%d"),
            "Días Activos":  st.column_config.NumberColumn("Días Activos",  format="%d"),
        })

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    gc1, gc2 = st.columns([2, 3], gap="medium")

    with gc1:
        st.markdown('<p class="sec">Cotizaciones por remitente</p>', unsafe_allow_html=True)
        fig_r = go.Figure(go.Bar(
            x=resumen["total_envios"], y=resumen["remitente_corto"], orientation="h",
            marker_color=bar_colors(len(resumen)), marker_line_width=0,
            text=resumen["total_envios"], textposition="outside",
            hovertemplate="<b>%{y}</b><br>%{x}<extra></extra>",
        ))
        fig_r.update_layout(**plotly_layout(height=270, xaxis_title="Cotizaciones", yaxis=dict(autorange="reversed")))
        st.plotly_chart(fig_r, use_container_width=True)

    with gc2:
        st.markdown('<p class="sec">Envíos por día — últimos 30 días</p>', unsafe_allow_html=True)
        e30 = envios_dia[envios_dia["fecha"] >= hoy - timedelta(days=30)].copy()
        e30["remitente_corto"] = e30["remitente"].apply(lambda x: x.split("@")[0] if "@" in str(x) else str(x))
        n_pal = e30["remitente_corto"].nunique()
        pal   = (px.colors.qualitative.Bold if n_pal <= 10 else px.colors.qualitative.Alphabet)[:n_pal]
        fig_l = px.line(e30.sort_values("fecha"), x="fecha", y="envios",
                        color="remitente_corto", markers=True,
                        template="plotly_white", color_discrete_sequence=pal)
        fig_l.update_layout(**plotly_layout(
            height=270, xaxis_title="", yaxis_title="Envíos", showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
            xaxis=dict(tickformat="%d %b")))
        fig_l.update_traces(marker_size=5, line_width=2)
        st.plotly_chart(fig_l, use_container_width=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    st.markdown('<p class="sec">Detalle por remitente</p>', unsafe_allow_html=True)

    rem_det = st.selectbox("Selecciona un remitente",
        options=resumen["remitente"].tolist(),
        format_func=lambda x: x.split("@")[0] if "@" in x else x,
        key="rem_det")
    df_rem = df_all[df_all["remitente"] == rem_det]

    d1, d2, d3 = st.columns(3)
    with d1: st.markdown(kpi_card("Total cotizaciones",     f"{df_rem['asunto'].nunique():,}",   "📨", BLUE),  unsafe_allow_html=True)
    with d2: st.markdown(kpi_card("Total claves",          f"{len(df_rem):,}",                   "📊", SKY),   unsafe_allow_html=True)
    with d3: st.markdown(kpi_card("Claves únicas pedidas", f"{df_rem['clave_imss'].nunique():,}", "🔑", GREEN), unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    dc1, dc2 = st.columns(2)

    with dc1:
        st.markdown('<p class="sec">GPO más solicitados</p>', unsafe_allow_html=True)
        tg = df_rem["gpo"].value_counts().head(8).reset_index()
        tg.columns = ["GPO", "Veces"]
        fig_g = go.Figure(go.Bar(
            x=tg["Veces"], y=tg["GPO"], orientation="h",
            marker_color=bar_colors(len(tg), (5, 150, 105)), marker_line_width=0,
            text=tg["Veces"], textposition="outside",
            hovertemplate="GPO <b>%{y}</b><br>%{x} veces<extra></extra>",
        ))
        fig_g.update_layout(**plotly_layout(height=270, xaxis_title="Frecuencia",
            yaxis=dict(autorange="reversed", type="category")))
        st.plotly_chart(fig_g, use_container_width=True)

    with dc2:
        st.markdown('<p class="sec">Envíos por día</p>', unsafe_allow_html=True)
        er = df_rem.groupby("fecha")["asunto"].nunique().reset_index(name="Envíos")
        fig_e = go.Figure(go.Bar(
            x=er["fecha"], y=er["Envíos"],
            marker_color=bar_colors(len(er), (14, 165, 233)), marker_line_width=0,
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y} envíos<extra></extra>",
        ))
        fig_e.update_layout(**plotly_layout(height=270, bargap=.3, xaxis_title="", yaxis_title="Envíos",
            xaxis=dict(tickformat="%d %b")))
        st.plotly_chart(fig_e, use_container_width=True)


# ── COTIZACIONES PDF ──────────────────────────────────────────────────────────
elif nav_page == "📄  Cotizaciones PDF":

    @st.cache_data(ttl=30, show_spinner=False)
    def _cargar_pdfs(directorio: str):
        d = Path(directorio)
        if not d.exists():
            return []
        items = []
        for p in sorted(d.glob("*.pdf"), key=lambda x: x.stat().st_mtime, reverse=True):
            info = {"path": str(p), "nombre": p.name, "size_kb": p.stat().st_size / 1024}
            try:
                pts = p.stem.split("_")
                info["expediente"] = pts[0] if pts else p.stem
                info["fecha"] = datetime.strptime(f"{pts[1]}_{pts[2]}", "%Y%m%d_%H%M%S")
            except Exception:
                info["fecha"] = datetime.fromtimestamp(p.stat().st_mtime)
            jp = p.with_suffix(".json")
            if jp.exists():
                try:
                    with open(jp, encoding="utf-8") as f:
                        meta = _json_mod.load(f)
                    info.update(meta)
                    if "fecha_correo" in meta and meta["fecha_correo"]:
                        try:
                            info["fecha_correo_dt"] = datetime.fromisoformat(meta["fecha_correo"])
                        except Exception:
                            pass
                except Exception:
                    pass
            info.setdefault("expediente",           "—")
            info.setdefault("remitente",            "—")
            info.setdefault("asunto",               "—")
            info.setdefault("num_partidas",         "—")
            info.setdefault("estado_envio",         "desconocido")
            info.setdefault("destinatario_enviado", "—")
            info.setdefault("error",                "")
            info.setdefault("modo_prueba",          None)
            items.append(info)
        return items

    pdfs_raw = _cargar_pdfs(str(PDF_DIR))

    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)
    total_p  = len(pdfs_raw)
    exitosos = sum(1 for x in pdfs_raw if x["estado_envio"] == "exitoso")
    fallidos = sum(1 for x in pdfs_raw if x["estado_envio"] == "fallido")
    hoy_pdfs = sum(1 for x in pdfs_raw if x["fecha"].date() == hoy)
    size_tot = sum(x["size_kb"] for x in pdfs_raw)
    size_str = f"{size_tot/1024:.1f} MB" if size_tot > 1024 else f"{size_tot:.0f} KB"

    k1p, k2p, k3p, k4p, k5p = st.columns(5)
    with k1p: st.markdown(kpi_card("Total generadas",  f"{total_p:,}",  "📄", BLUE),       unsafe_allow_html=True)
    with k2p: st.markdown(kpi_card("Enviadas hoy",     f"{hoy_pdfs:,}", "📅", SKY),        unsafe_allow_html=True)
    with k3p: st.markdown(kpi_card("Envíos exitosos",  f"{exitosos:,}", "✅", GREEN),       unsafe_allow_html=True)
    with k4p: st.markdown(kpi_card("Envíos fallidos",  f"{fallidos:,}", "❌", RED),         unsafe_allow_html=True)
    with k5p: st.markdown(kpi_card("Tamaño en disco",  size_str,        "💾", AMBER),       unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    if not pdfs_raw:
        st.info("Aún no hay cotizaciones PDF generadas. Los PDFs aparecerán aquí automáticamente cuando lleguen correos.")
    else:
        fa, fb, fc, fd = st.columns([2, 2, 2, 1.2])
        with fa:
            busq_txt = st.text_input("🔍  Expediente / Nombre", placeholder="I-2026-...", key="pdf_busq")
        with fb:
            busq_rem = st.text_input("📧  Remitente", placeholder="compras@...", key="pdf_rem")
        with fc:
            estado_sel = st.selectbox("📊  Estado", ["Todos", "Exitoso", "Fallido", "Desconocido"], key="pdf_estado")
        with fd:
            if st.button("↺  Refrescar", use_container_width=True, key="pdf_refresh"):
                st.cache_data.clear()
                st.rerun()

        fe, ff = st.columns([1.5, 2.5])
        fechas_disp = sorted({x["fecha"].date() for x in pdfs_raw}, reverse=True)
        with fe:
            fecha_desde = st.date_input("📅  Desde", value=fechas_disp[-1] if fechas_disp else hoy, key="pdf_desde")
        with ff:
            fecha_hasta = st.date_input("📅  Hasta", value=hoy, key="pdf_hasta")

        pdfs = pdfs_raw
        if busq_txt:
            q = busq_txt.lower()
            pdfs = [x for x in pdfs if q in x["expediente"].lower() or q in x["nombre"].lower() or q in x["asunto"].lower()]
        if busq_rem:
            q = busq_rem.lower()
            pdfs = [x for x in pdfs if q in x["remitente"].lower() or q in x["destinatario_enviado"].lower()]
        if estado_sel != "Todos":
            mapa = {"Exitoso": "exitoso", "Fallido": "fallido", "Desconocido": "desconocido"}
            pdfs = [x for x in pdfs if x["estado_envio"] == mapa[estado_sel]]
        pdfs = [x for x in pdfs if fecha_desde <= x["fecha"].date() <= fecha_hasta]

        st.markdown(
            f"<div style='font-size:.75rem;color:{MUTED};margin:.4rem 0 .8rem'>"
            f"Mostrando <strong style='color:{NAVY}'>{len(pdfs)}</strong> de "
            f"<strong style='color:{NAVY}'>{total_p}</strong> cotizaciones</div>",
            unsafe_allow_html=True)

        pp1, pp2, pp3 = st.columns([1.4, 1.6, 2])
        with pp1:
            pdf_por_pag = st.selectbox("Registros por página", [10, 25, 50, 100], index=1, key="pdf_por_pagina")
        pdf_total_pags = max(1, -(-len(pdfs) // pdf_por_pag))

        _pdf_fsig = f"{busq_txt}|{busq_rem}|{estado_sel}|{fecha_desde}|{fecha_hasta}|{pdf_por_pag}"
        if st.session_state.get("_pdf_fsig") != _pdf_fsig:
            st.session_state["_pdf_fsig"] = _pdf_fsig
            st.session_state["pdf_pagina"] = 1

        with pp2:
            pdf_pagina = st.number_input("Página", min_value=1, max_value=pdf_total_pags,
                value=st.session_state.get("pdf_pagina", 1), step=1, key="pdf_pagina")
        with pp3:
            st.markdown(
                f"<div style='padding-top:1.85rem;color:{MUTED};font-size:.78rem'>"
                f"de <strong>{pdf_total_pags}</strong> páginas &nbsp;·&nbsp; "
                f"<strong>{len(pdfs):,}</strong> cotizaciones</div>",
                unsafe_allow_html=True)

        pdf_inicio = (pdf_pagina - 1) * pdf_por_pag
        pdfs_pag   = pdfs[pdf_inicio: pdf_inicio + pdf_por_pag]

        HEADS = ["#", "Expediente", "Remitente", "Fecha generación", "Asunto", "Claves", "Estado", "Ver", "Descargar"]
        COLS  = [.5, 1.7, 1.9, 1.5, 2.4, .55, .95, .8, .95]

        head_cols = st.columns(COLS)
        for col, txt in zip(head_cols, HEADS):
            col.markdown(
                f"<div class='sec' style='margin-bottom:.3rem;text-align:center'>{txt}</div>",
                unsafe_allow_html=True)

        for i, det in enumerate(pdfs_pag):
            bg  = "#f9fafb" if i % 2 == 0 else "#ffffff"
            pad = "padding:.55rem .55rem;border-radius:6px;"
            estado_color = "#166534" if det["estado_envio"] == "exitoso" else ("#991b1b" if det["estado_envio"] == "fallido" else "#64748b")
            estado_bg    = "#dcfce7" if det["estado_envio"] == "exitoso" else ("#fee2e2" if det["estado_envio"] == "fallido" else "#f1f5f9")

            rc = st.columns(COLS)
            rc[0].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{MUTED};text-align:center'>{pdf_inicio+i+1}</div>", unsafe_allow_html=True)
            rc[1].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{NAVY};font-weight:600'>{det['expediente']}</div>", unsafe_allow_html=True)
            rc[2].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:{MUTED};white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{det['remitente']}'>{det['remitente']}</div>", unsafe_allow_html=True)
            rc[3].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:#374151'>{det['fecha'].strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)
            rc[4].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:{MUTED};white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{det['asunto']}'>{det['asunto']}</div>", unsafe_allow_html=True)
            rc[5].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{NAVY};text-align:center;font-weight:600'>{det['num_partidas']}</div>", unsafe_allow_html=True)
            rc[6].markdown(
                f"<div style='background:{bg};{pad};text-align:center'>"
                f"<span style='background:{estado_bg};color:{estado_color};font-size:.68rem;"
                f"font-weight:700;padding:.18rem .55rem;border-radius:5px'>{det['estado_envio'].upper()}</span></div>",
                unsafe_allow_html=True)
            with rc[7]:
                if st.button("👁", key=f"ver_{i}", use_container_width=True, help="Vista previa PDF"):
                    _ver_pdf_dialog(det["path"], det["nombre"], det["expediente"])
            with rc[8]:
                try:
                    with open(det["path"], "rb") as fh:
                        pdf_b = fh.read()
                    st.download_button("⬇", pdf_b, file_name=det["nombre"],
                        mime="application/pdf", key=f"dl_{i}", use_container_width=True)
                except Exception:
                    st.markdown(f"<div style='background:{bg};{pad}font-size:.7rem;color:{MUTED};text-align:center'>N/D</div>", unsafe_allow_html=True)

        pbp, _, pbn = st.columns([1, 4, 1])
        with pbp:
            if pdf_pagina > 1:
                if st.button("← Anterior", use_container_width=True, key="pdf_prev"):
                    st.session_state["pdf_pagina"] = pdf_pagina - 1
                    st.rerun()
        with pbn:
            if pdf_pagina < pdf_total_pags:
                if st.button("Siguiente →", use_container_width=True, key="pdf_next"):
                    st.session_state["pdf_pagina"] = pdf_pagina + 1
                    st.rerun()

        # Detalle expandible del PDF seleccionado
        st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
        st.markdown('<p class="sec">Detalle del PDF seleccionado</p>', unsafe_allow_html=True)
        det_idx = st.selectbox(
            "Selecciona expediente",
            range(len(pdfs_pag)),
            format_func=lambda i: f"{pdfs_pag[i]['expediente']} — {pdfs_pag[i]['fecha'].strftime('%d/%m/%Y %H:%M')}",
            key="pdf_det_idx",
        )
        if pdfs_pag:
            det = pdfs_pag[det_idx]
            d1, d2, d3 = st.columns(3)
            estado_color = "#166534" if det["estado_envio"] == "exitoso" else ("#991b1b" if det["estado_envio"] == "fallido" else "#64748b")
            estado_bg    = "#dcfce7" if det["estado_envio"] == "exitoso" else ("#fee2e2" if det["estado_envio"] == "fallido" else "#f1f5f9")
            err_txt      = f"<br><b>Error:</b> <span style='color:#991b1b'>{det.get('error','—')}</span>" if det.get("error") else ""
            modo_txt     = "Prueba" if det.get("modo_prueba") else ("Producción" if det.get("modo_prueba") is False else "—")

            with d1:
                st.markdown(
                    f"<div class='insight-card'><div class='insight-label'>Cotización</div>"
                    f"<div style='margin:.4rem 0'></div>"
                    f"<div class='insight-text'><b>Asunto:</b> {det['asunto']}<br><b>Claves:</b> {det['num_partidas']}</div>"
                    f"</div>", unsafe_allow_html=True)
            with d2:
                st.markdown(
                    f"<div class='insight-card'><div class='insight-label'>Datos de envío</div>"
                    f"<div style='margin:.4rem 0'></div>"
                    f"<div class='insight-text'><b>Remitente:</b> {det['remitente']}<br>"
                    f"<b>Enviado a:</b> {det['destinatario_enviado']}<br><b>Modo:</b> {modo_txt}</div>"
                    f"</div>", unsafe_allow_html=True)
            with d3:
                st.markdown(
                    f"<div class='insight-card'><div class='insight-label'>Estado y archivo</div>"
                    f"<div style='margin:.4rem 0'>"
                    f"<span style='background:{estado_bg};color:{estado_color};font-size:.85rem;"
                    f"font-weight:700;padding:.3rem .8rem;border-radius:20px'>{det['estado_envio'].upper()}</span>"
                    f"</div>"
                    f"<div class='insight-text'><b>Generado:</b> {det['fecha'].strftime('%d/%m/%Y %H:%M:%S')}<br>"
                    f"<b>Archivo:</b> {det['nombre']}<br><b>Tamaño:</b> {det['size_kb']:.1f} KB{err_txt}</div>"
                    f"</div>", unsafe_allow_html=True)


# ── CORREOS PROCESADOS ────────────────────────────────────────────────────────
elif nav_page == "📥  Correos Procesados":
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    df_cp = cargar_correos()

    if df_cp.empty:
        st.info("Aún no hay correos registrados en la tabla de auditoría.")
    else:
        total_cp      = len(df_cp)
        complet_cp    = int((df_cp["estado"] == "completado").sum())
        sintabla_cp   = int((df_cp["estado"] == "sin_tabla").sum())
        error_cp      = int((df_cp["estado"] == "error").sum())
        procesando_cp = int((df_cp["estado"] == "procesando").sum())

        kcp1, kcp2, kcp3, kcp4, kcp5 = st.columns(5)
        with kcp1: st.markdown(kpi_card("Total correos",  f"{total_cp:,}",      "📧", NAVY),  unsafe_allow_html=True)
        with kcp2: st.markdown(kpi_card("Completados",    f"{complet_cp:,}",    "✅", GREEN), unsafe_allow_html=True)
        with kcp3: st.markdown(kpi_card("Sin tabla",      f"{sintabla_cp:,}",   "📭", AMBER), unsafe_allow_html=True)
        with kcp4: st.markdown(kpi_card("Errores",        f"{error_cp:,}",      "❌", RED),   unsafe_allow_html=True)
        with kcp5: st.markdown(kpi_card("En proceso",     f"{procesando_cp:,}", "⏳", SKY),   unsafe_allow_html=True)

        st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)

        # Filtros
        cf1, cf2, cf3 = st.columns([2.5, 2.5, 1.5])
        with cf1:
            cp_busq_asunto = st.text_input("🔍  Buscar asunto", placeholder="cotización de compra...", key="cp_asunto")
        with cf2:
            cp_busq_rem = st.text_input("👤  Remitente IMSS", placeholder="usuario@imss.gob.mx", key="cp_rem")
        with cf3:
            cp_estado_sel = st.selectbox("📊  Estado", ["Todos", "Completado", "Sin tabla", "Error", "Procesando"], key="cp_estado")
        cd1, cd2 = st.columns(2)
        with cd1:
            cp_fecha_desde = st.date_input("📅  Desde", value=hoy - timedelta(days=30), key="cp_desde")
        with cd2:
            cp_fecha_hasta = st.date_input("📅  Hasta", value=hoy, key="cp_hasta")

        df_f = df_cp.copy()
        if cp_busq_asunto:
            df_f = df_f[df_f["asunto"].str.contains(cp_busq_asunto, case=False, na=False)]
        if cp_busq_rem:
            df_f = df_f[df_f["remitente_imss"].str.contains(cp_busq_rem, case=False, na=False)]
        if cp_estado_sel != "Todos":
            _mapa_e = {"Completado": "completado", "Sin tabla": "sin_tabla", "Error": "error", "Procesando": "procesando"}
            df_f = df_f[df_f["estado"] == _mapa_e[cp_estado_sel]]
        df_f = df_f[(df_f["fecha"] >= cp_fecha_desde) & (df_f["fecha"] <= cp_fecha_hasta)].reset_index(drop=True)

        st.markdown(
            f"<div style='font-size:.75rem;color:{MUTED};margin:.4rem 0 .8rem'>"
            f"Mostrando <strong style='color:{NAVY}'>{len(df_f)}</strong> de "
            f"<strong style='color:{NAVY}'>{total_cp}</strong> correos</div>",
            unsafe_allow_html=True)

        # Resolver salto de página antes de instanciar el number_input
        if "_cp_pagina_jump" in st.session_state:
            st.session_state["cp_pagina"] = st.session_state.pop("_cp_pagina_jump")

        cp_pp1, cp_pp2, cp_pp3 = st.columns([1.4, 1.6, 2])
        with cp_pp1:
            cp_por_pag = st.selectbox("Registros por página", [10, 25, 50, 100], index=1, key="cp_por_pagina")
        cp_total_pags = max(1, -(-len(df_f) // cp_por_pag))
        _cp_fsig = f"{cp_busq_asunto}|{cp_busq_rem}|{cp_estado_sel}|{cp_fecha_desde}|{cp_fecha_hasta}|{cp_por_pag}"
        if st.session_state.get("_cp_fsig") != _cp_fsig:
            st.session_state["_cp_fsig"] = _cp_fsig
            st.session_state["_cp_pagina_jump"] = 1
        with cp_pp2:
            cp_pagina = st.number_input("Página", min_value=1, max_value=cp_total_pags,
                value=st.session_state.get("cp_pagina", 1), step=1, key="cp_pagina")
        with cp_pp3:
            st.markdown(
                f"<div style='padding-top:1.85rem;color:{MUTED};font-size:.78rem'>"
                f"de <strong>{cp_total_pags}</strong> páginas &nbsp;·&nbsp; "
                f"<strong>{len(df_f):,}</strong> correos</div>",
                unsafe_allow_html=True)

        cp_inicio  = (cp_pagina - 1) * cp_por_pag
        df_pag_cp  = df_f.iloc[cp_inicio: cp_inicio + cp_por_pag]

        _ESTADO_C = {
            "completado": ("#166534", "#dcfce7", "COMPLETADO"),
            "sin_tabla":  ("#92400e", "#fef3c7", "SIN TABLA"),
            "error":      ("#991b1b", "#fee2e2", "ERROR"),
            "procesando": ("#1e40af", "#dbeafe", "PROCESANDO"),
        }
        CP_HEADS = ["#", "Fecha",  "Remitente IMSS",  "Asunto",  "Estado",  "Claves",  "Expediente"]
        CP_COLS  = [.35,  1.3,      2.0,               2.8,       1.05,      .65,       1.4]

        head_cp = st.columns(CP_COLS)
        for col, txt in zip(head_cp, CP_HEADS):
            col.markdown(f"<div class='sec' style='margin-bottom:.3rem;text-align:center'>{txt}</div>", unsafe_allow_html=True)

        for i, (_, row) in enumerate(df_pag_cp.iterrows()):
            bg  = "#f9fafb" if i % 2 == 0 else "#ffffff"
            pad = "padding:.55rem .55rem;border-radius:6px;"
            ev  = str(row.get("estado", "")).lower()
            e_color, e_bg, e_lbl = _ESTADO_C.get(ev, ("#64748b", "#f1f5f9", ev.upper()))
            rem_txt  = str(row.get("remitente_imss", "—"))
            asunto_t = str(row.get("asunto", "—"))
            exp_txt  = str(row.get("expediente", "—") or "—")
            claves_n = int(row.get("filas_insertadas", 0) or 0)

            rc = st.columns(CP_COLS)
            rc[0].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{MUTED};text-align:center'>{cp_inicio+i+1}</div>", unsafe_allow_html=True)
            rc[1].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:#374151'>{row['fecha_correo'].strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)
            rc[2].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:{MUTED};white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{rem_txt}'>{rem_txt}</div>", unsafe_allow_html=True)
            rc[3].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:{NAVY};white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{asunto_t}'>{asunto_t}</div>", unsafe_allow_html=True)
            rc[4].markdown(
                f"<div style='background:{bg};{pad};text-align:center'>"
                f"<span style='background:{e_bg};color:{e_color};font-size:.67rem;font-weight:700;"
                f"padding:.18rem .5rem;border-radius:5px'>{e_lbl}</span></div>",
                unsafe_allow_html=True)
            rc[5].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{NAVY};text-align:center;font-weight:600'>{claves_n}</div>", unsafe_allow_html=True)
            rc[6].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:#374151'>{exp_txt}</div>", unsafe_allow_html=True)

            if ev == "error" and row.get("error_mensaje"):
                with st.expander(f"  ⚠  Error en correo {cp_inicio+i+1}", expanded=False):
                    st.code(str(row["error_mensaje"]), language=None)

        cbp, _, cbn = st.columns([1, 4, 1])
        with cbp:
            if cp_pagina > 1:
                if st.button("← Anterior", use_container_width=True, key="cp_prev"):
                    st.session_state["_cp_pagina_jump"] = cp_pagina - 1
                    st.rerun()
        with cbn:
            if cp_pagina < cp_total_pags:
                if st.button("Siguiente →", use_container_width=True, key="cp_next"):
                    st.session_state["_cp_pagina_jump"] = cp_pagina + 1
                    st.rerun()


# ── USUARIOS (solo admin) ─────────────────────────────────────────────────────
elif nav_page == "⚙️  Usuarios" and _es_admin_actual:
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    usuario_actual = st.session_state.get("auth_user", "")
    cfg_u         = _cargar_usuarios()
    usuarios_dict = cfg_u["credentials"]["usernames"]
    n_admins      = _contar_admins()

    total_u = len(usuarios_dict)
    n_users = total_u - n_admins
    ku1, ku2, ku3 = st.columns(3)
    with ku1: st.markdown(kpi_card("Total usuarios",    f"{total_u}",  "👥", BLUE),  unsafe_allow_html=True)
    with ku2: st.markdown(kpi_card("Administradores",   f"{n_admins}", "🛡️", GREEN), unsafe_allow_html=True)
    with ku3: st.markdown(kpi_card("Usuarios estándar", f"{n_users}",  "👤", SKY),   unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown('<p class="sec">Usuarios registrados</p>', unsafe_allow_html=True)

    UH = ["Usuario", "Nombre", "Email", "Rol", "Acciones"]
    UC = [2.4, 2.0, 2.4, 0.9, 1.4]
    head_u = st.columns(UC)
    for col, txt in zip(head_u, UH):
        col.markdown(f"<div class='sec' style='margin-bottom:.3rem;text-align:center'>{txt}</div>", unsafe_allow_html=True)

    for i, (uname, udata) in enumerate(usuarios_dict.items()):
        bg = "#f9fafb" if i % 2 == 0 else "#ffffff"
        pad = "padding:.55rem .55rem;border-radius:6px;"
        es_yo        = (uname == usuario_actual)
        es_ese_admin = udata.get("role") == "admin"
        rol_color = "#166534" if es_ese_admin else "#64748b"
        rol_bg    = "#dcfce7" if es_ese_admin else "#f1f5f9"
        rol_txt   = "ADMIN" if es_ese_admin else "USER"
        badge_yo  = (" <span style='background:#eff6ff;color:#2563eb;font-size:.6rem;font-weight:700;"
                     "padding:.1rem .4rem;border-radius:4px;margin-left:.3rem'>TÚ</span>") if es_yo else ""

        rc = st.columns(UC)
        rc[0].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:{NAVY};font-weight:600;"
                       f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{uname}'>"
                       f"{uname}{badge_yo}</div>", unsafe_allow_html=True)
        rc[1].markdown(f"<div style='background:{bg};{pad}font-size:.78rem;color:#374151'>{udata.get('name','—')}</div>", unsafe_allow_html=True)
        rc[2].markdown(f"<div style='background:{bg};{pad}font-size:.76rem;color:{MUTED};"
                       f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis' title='{udata.get('email','—')}'>"
                       f"{udata.get('email','—')}</div>", unsafe_allow_html=True)
        rc[3].markdown(f"<div style='background:{bg};{pad};text-align:center'>"
                       f"<span style='background:{rol_bg};color:{rol_color};font-size:.68rem;font-weight:700;"
                       f"padding:.18rem .55rem;border-radius:5px'>{rol_txt}</span></div>", unsafe_allow_html=True)
        with rc[4]:
            bc1, bc2 = st.columns(2)
            with bc1:
                nuevo_rol      = "user" if es_ese_admin else "admin"
                label_rol      = "↓" if es_ese_admin else "↑"
                deshabilitar_r = es_yo and es_ese_admin and n_admins == 1
                if st.button(label_rol, key=f"rol_{i}", use_container_width=True, disabled=deshabilitar_r,
                             help="No puedes quitarte admin: eres el único." if deshabilitar_r
                             else ("Bajar a USER" if es_ese_admin else "Promover a ADMIN")):
                    _cambiar_rol(uname, nuevo_rol)
                    st.rerun()
            with bc2:
                deshabilitar_d = es_yo or (es_ese_admin and n_admins == 1)
                confirm_key    = f"confirm_del_{uname}"
                if st.session_state.get(confirm_key):
                    if st.button("✓", key=f"del_ok_{i}", use_container_width=True, type="primary",
                                 help="Confirmar eliminación"):
                        _eliminar_usuario(uname)
                        st.session_state.pop(confirm_key, None)
                        st.rerun()
                else:
                    if st.button("🗑", key=f"del_{i}", use_container_width=True, disabled=deshabilitar_d,
                                 help=("No puedes eliminarte" if es_yo
                                       else ("No puedes eliminar al único admin" if deshabilitar_d
                                             else "Eliminar usuario"))):
                        st.session_state[confirm_key] = True
                        st.rerun()

    st.markdown("<div style='height:1.4rem'></div>", unsafe_allow_html=True)
    ac1, ac2 = st.columns(2, gap="large")

    with ac1:
        st.markdown('<p class="sec">Agregar nuevo usuario</p>', unsafe_allow_html=True)
        with st.form("form_agregar_usuario", clear_on_submit=True):
            nu_usuario = st.text_input("Usuario (correo)", placeholder="usuario@empresa.com")
            nu_nombre  = st.text_input("Nombre completo",  placeholder="Juan Pérez")
            nu_email   = st.text_input("Email (opcional)", placeholder="se usa el usuario si está vacío")
            nu_clave   = st.text_input("Contraseña",       type="password")
            nu_rol     = st.selectbox("Rol", ["user", "admin"])
            btn_agregar = st.form_submit_button("➕  Agregar usuario", use_container_width=True)
            if btn_agregar:
                if not nu_usuario or not nu_nombre or not nu_clave:
                    st.error("Usuario, nombre y contraseña son obligatorios.")
                elif nu_usuario in usuarios_dict:
                    st.error(f"El usuario '{nu_usuario}' ya existe.")
                elif len(nu_clave) < 6:
                    st.error("La contraseña debe tener al menos 6 caracteres.")
                else:
                    _agregar_usuario(nu_usuario, nu_nombre, nu_email, nu_clave, nu_rol)
                    st.success(f"Usuario '{nu_usuario}' agregado correctamente.")
                    st.rerun()

    with ac2:
        st.markdown('<p class="sec">Cambiar contraseña</p>', unsafe_allow_html=True)
        with st.form("form_cambiar_password", clear_on_submit=True):
            cp_usuario = st.selectbox("Usuario", list(usuarios_dict.keys()))
            cp_clave1  = st.text_input("Nueva contraseña",    type="password")
            cp_clave2  = st.text_input("Confirmar contraseña", type="password")
            btn_cambiar = st.form_submit_button("🔑  Cambiar contraseña", use_container_width=True)
            if btn_cambiar:
                if not cp_clave1 or not cp_clave2:
                    st.error("Ingresa la contraseña dos veces.")
                elif cp_clave1 != cp_clave2:
                    st.error("Las contraseñas no coinciden.")
                elif len(cp_clave1) < 6:
                    st.error("La contraseña debe tener al menos 6 caracteres.")
                else:
                    _cambiar_password(cp_usuario, cp_clave1)
                    if cp_usuario == usuario_actual:
                        _clear_remember()
                    st.success(f"Contraseña de '{cp_usuario}' actualizada.")
                    st.rerun()

    st.markdown(
        f"<p style='color:#d1d5db;font-size:.68rem;text-align:center;margin-top:1.5rem'>"
        "Panel de administración &nbsp;·&nbsp; Solo visible para usuarios con rol <strong>admin</strong></p>",
        unsafe_allow_html=True)

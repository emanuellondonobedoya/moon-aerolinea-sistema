# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Identidad visual y componentes de interfaz (ui_styles.py)

Azul marino #0B1D33/#10243A · Azul #356B9A · Dorado #C9A96E/#F4B41A · Manrope (títulos) y DM Sans (texto).
Incluye: CSS corporativo, logo, siluetas SVG de aeronaves, tarjetas KPI, stepper, cláusula de responsabilidad,
indicadores de "Dato manual / Fórmula pendiente", diagrama de gobernanza y estilos Plotly.
"""
from __future__ import annotations

import base64
from functools import lru_cache
from html import escape
from pathlib import Path
from textwrap import dedent

import pandas as pd
import plotly.express as px
import streamlit as st

NAVY, NAVY2, BLUE, SKY = "#0B1D33", "#10243A", "#356B9A", "#7FA9CF"
GOLD, AMBER = "#C9A96E", "#F4B41A"
INK, MUTED, BG = "#17283D", "#5F7087", "#F4F7FB"
OK, WARN, BAD = "#1F8A5B", "#C98A1B", "#B83A32"
PALETA = [BLUE, NAVY, GOLD, SKY, AMBER, "#4F6F91", "#E3D3B0", "#2E4A68", "#8FB3D1", "#9C7A3C"]
ESCALA = [[0, "#CFDDEB"], [1, NAVY]]
ESCALA_SUAVE = [[0, "#F1F5FA"], [1, BLUE]]
AMARILLO_DUDA = "#FFF3C4"
AMARILLO_MANUAL = "#FFF9E6"
PLOT_CFG = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"]}

_CSS = """
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');
html, body, .stApp, .stMarkdown, p, label, input, textarea, button, [data-testid="stSidebar"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: __BG__; color: __INK__; }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 2.2rem; max-width: 1380px; }
h1, h2, h3, h4 { font-family: 'Manrope', sans-serif !important; color: __NAVY__; letter-spacing: -.4px; }

/* Barra lateral */
[data-testid="stSidebar"] { background: linear-gradient(180deg, __NAVY__ 0%, __NAVY2__ 55%, #163C61 100%); }
[data-testid="stSidebar"] * { color: #E6EDF5 !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.14); }
[data-testid="stSidebar"] [role="radiogroup"] label { padding: 7px 10px; border-radius: 10px; margin-bottom: 2px; }
[data-testid="stSidebar"] [role="radiogroup"] label:hover { background: rgba(255,255,255,.08); }
[data-testid="stSidebar"] .stButton > button { background: rgba(255,255,255,.08) !important; border: 1px solid rgba(255,255,255,.28) !important; }
[data-testid="stSidebar"] .stButton > button:hover { background: rgba(244,180,26,.18) !important; border-color: __AMBER__ !important; }
.side-logo { background: #fff; border-radius: 14px; padding: 10px 14px; margin: 0 0 10px 0; text-align: center; }
.side-logo img { width: 100%; max-width: 190px; }
.side-user { font-size: 13px; line-height: 1.45; opacity: .95; }
.side-user b { color: __AMBER__ !important; font-weight: 700; }
.side-stat { font-size: 12px; opacity: .8; margin: 3px 0; }

/* Títulos */
.page-title { color: __NAVY__; font-family: 'Manrope', sans-serif; font-weight: 800; font-size: clamp(26px, 3vw, 34px);
  letter-spacing: -1.1px; line-height: 1.1; margin: 0 0 12px 0; }
.section-heading { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 17px; font-weight: 800; letter-spacing: -.3px; margin: 22px 0 2px 0; }
.section-sub { color: __MUTED__; font-size: 13px; margin: 0 0 10px 0; }

/* Hero con ruta de vuelo y avión */
.hero-panel { position: relative; overflow: hidden; border-radius: 20px; padding: 26px 30px; margin: 4px 0 20px 0;
  background: linear-gradient(115deg, __NAVY__ 0%, __NAVY2__ 55%, #1A4568 100%); box-shadow: 0 10px 25px rgba(11,29,51,.14); }
.hero-tag { display: inline-block; background: rgba(244,180,26,.16); color: __AMBER__; border: 1px solid rgba(244,180,26,.35);
  border-radius: 999px; padding: 3px 12px; font-size: 12px; font-weight: 700; margin-bottom: 10px; letter-spacing: .3px; }
.hero-title { color: #fff; font-family: 'Manrope', sans-serif; font-size: 25px; font-weight: 800; letter-spacing: -.6px; max-width: 620px; }
.hero-copy { color: #C3D0DE; font-size: 14px; margin-top: 8px; max-width: 600px; line-height: 1.6; }
.hero-route { position: absolute; right: 0; top: 0; height: 100%; width: 46%; pointer-events: none; }
.hero-tag, .hero-title, .hero-copy { position: relative; z-index: 1; }
.hero-title, .hero-copy { max-width: min(600px, 56%); }
@media (max-width: 820px) { .hero-title, .hero-copy { max-width: 100%; } }
@media (max-width: 820px) { .hero-route { display: none; } }

/* Tarjetas KPI */
.metric-card { background: #fff; border: 1px solid #E5EAF1; border-radius: 14px; padding: 15px 17px; margin-bottom: 12px; min-height: 96px; }
.metric-card.alerta { border-left: 4px solid __WARN__; }
.metric-card.bien { border-left: 4px solid __OK__; }
.metric-card.mal { border-left: 4px solid __BAD__; }
.metric-card.oro { border-left: 4px solid __AMBER__; }
.metric-label { color: __MUTED__; font-size: 12.5px; font-weight: 600; }
.metric-value { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 23px; font-weight: 800; margin-top: 7px; line-height: 1.15; }
.kpi-delta { display: inline-block; margin-top: 7px; font-size: 12px; font-weight: 700; color: __BLUE__; }
.kpi-delta.neg { color: __BAD__; }
.kpi-note { margin-top: 5px; font-size: 12px; color: __MUTED__; }

/* Stepper */
.stepper { display: flex; gap: 8px; margin: 0 0 14px 0; flex-wrap: wrap; }
.step { padding: 6px 14px; border-radius: 999px; background: #E6ECF4; color: __MUTED__; font-size: 13px; font-weight: 700; }
.step.on { background: __NAVY__; color: #fff; }
.step.done { background: #DCEBE3; color: __OK__; }

/* Confirmación */
@keyframes pop { 0% { transform: scale(.6); opacity: 0; } 70% { transform: scale(1.08); } 100% { transform: scale(1); opacity: 1; } }
.ok-card { display: flex; gap: 16px; align-items: center; background: #fff; border: 1px solid #CFE8DC; border-left: 5px solid __OK__;
  border-radius: 14px; padding: 18px 20px; margin: 8px 0 16px 0; }
.ok-icon { width: 44px; height: 44px; border-radius: 50%; background: __OK__; color: #fff; display: flex; align-items: center;
  justify-content: center; font-size: 22px; font-weight: 800; animation: pop .5s ease-out; flex: none; }
.ok-title { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 17px; font-weight: 800; }
.ok-copy { color: __MUTED__; font-size: 14px; margin-top: 2px; word-break: break-word; }

/* Cláusula de responsabilidad y pendientes */
.clausula { background: #F7F9FC; border: 1px solid #D9E2EC; border-left: 4px solid __GOLD__; border-radius: 12px;
  padding: 10px 14px; margin: 8px 0 6px 0; font-size: 13px; color: __INK__; }
.clausula b { color: __NAVY__; }
.manual { background: __AMARILLO_MANUAL__; border: 1px dashed #E0B94A; border-radius: 10px; padding: 8px 12px; font-size: 13px;
  color: #6B5215; margin: 6px 0; }
.dudoso { background: __AMARILLO_DUDA__; border: 1px solid #E9C94F; border-radius: 10px; padding: 8px 12px; font-size: 13px;
  color: #5C4610; margin: 6px 0; }
.chip { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; margin: 0 4px 4px 0; }
.chip.ok { background: #DCEBE3; color: __OK__; } .chip.warn { background: #FBEFD5; color: #8A5D0E; }
.chip.bad { background: #F6DEDC; color: __BAD__; } .chip.info { background: #E3ECF6; color: __BLUE__; }
.chip.navy { background: __NAVY__; color: #fff; }

/* Vacío */
.vacio { text-align: center; padding: 30px 20px; background: #fff; border: 1px dashed #C5D3E2; border-radius: 16px; color: __MUTED__; }
.vacio b { display: block; color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 16px; margin: 6px 0 4px 0; }
.vacio .ico { font-size: 30px; }

/* Norte estratégico */
.norte-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; margin: 6px 0 16px 0; }
.norte-card { background: #fff; border: 1px solid #E5EAF1; border-radius: 16px; padding: 18px 18px 16px 18px; position: relative; overflow: hidden; }
.norte-card:before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 5px; background: linear-gradient(180deg, __AMBER__, __GOLD__); }
.norte-card.navy { background: linear-gradient(135deg, __NAVY__, #163C61); border: none; }
.norte-card.navy .norte-k, .norte-card.navy .norte-t { color: #fff; }
.norte-k { font-family: 'Manrope', sans-serif; color: __GOLD__; font-weight: 800; font-size: 12px; letter-spacing: 1.4px; text-transform: uppercase; }
.norte-t { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 17px; font-weight: 700; line-height: 1.4; margin-top: 8px; }
.valores { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
.valor { background: #F3EEE2; color: #6E5524; border-radius: 999px; padding: 5px 12px; font-weight: 700; font-size: 13px; }

/* Diagrama de gobernanza */
.flujo { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); gap: 14px; margin: 8px 0 10px 0; }
@media (max-width: 1250px) { .flujo { grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); } }
.paso { position: relative; background: #fff; border: 1px solid #E5EAF1; border-top: 3px solid __NAVY__; border-radius: 14px; padding: 12px 12px; }
.paso:not(:last-child):after { content: "›"; position: absolute; right: -12px; top: 38%; color: __GOLD__; font-size: 22px; font-weight: 800; }
@media (max-width: 1250px) { .paso:after { display: none; } }
.paso .n { display: inline-flex; width: 24px; height: 24px; border-radius: 50%; background: __NAVY__; color: #fff; font-size: 12px; font-weight: 800;
  align-items: center; justify-content: center; margin-right: 6px; }
.paso .t { font-family: 'Manrope', sans-serif; font-weight: 800; color: __NAVY__; font-size: 14px; }
.paso .d { color: __MUTED__; font-size: 12.5px; margin-top: 6px; line-height: 1.45; }
.cimiento { background: linear-gradient(90deg, __NAVY__, #163C61); color: #E6EDF5; border-radius: 14px; padding: 12px 16px; font-size: 13px; margin-top: 4px; }
.cimiento b { color: __AMBER__; }

/* Flota */
.flota { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 12px; margin: 6px 0 14px 0; }
.avion { background: #fff; border: 1px solid #E5EAF1; border-radius: 16px; padding: 14px 16px; display: flex; gap: 12px; align-items: center; }
.avion svg { flex: none; }
.avion .m { font-family: 'Manrope', sans-serif; font-weight: 800; color: __NAVY__; font-size: 16px; }
.avion .s { color: __MUTED__; font-size: 12.5px; }
.sem { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 6px; vertical-align: middle; }

/* Login */
.login-panel { background: linear-gradient(140deg, __NAVY__ 0%, __NAVY2__ 60%, #1A4568 100%); border-radius: 22px; padding: 30px 30px; color: #fff;
  min-height: 430px; position: relative; overflow: hidden; }
.login-panel .lt { font-family: 'Manrope', sans-serif; font-size: 28px; font-weight: 800; letter-spacing: -.8px; margin-top: 18px; }
.login-panel .lc { color: #C3D0DE; font-size: 14px; line-height: 1.6; margin-top: 8px; max-width: 420px; }
.login-logo { background: #fff; display: inline-block; border-radius: 14px; padding: 10px 16px; }
.login-logo img { width: 210px; }
.login-avion { position: absolute; right: -10px; bottom: -10px; opacity: .9; }

/* Controles */
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid #DDE5EE; flex-wrap: wrap; }
.stTabs [data-baseweb="tab"] { height: 44px; padding: 0 14px; border-radius: 12px 12px 0 0; font-weight: 700; color: __MUTED__; }
.stTabs [aria-selected="true"] { color: __NAVY__; background: #fff; }
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {
  border-radius: 12px; font-weight: 700; transition: transform .2s ease, box-shadow .2s ease; }
.stButton > button:hover:enabled, .stDownloadButton > button:hover:enabled { transform: translateY(-1px); box-shadow: 0 6px 14px rgba(11,29,51,.14); }
button[kind="primary"], button[data-testid="stBaseButton-primary"], button[data-testid="stBaseButton-primaryFormSubmit"] {
  background: __NAVY__ !important; border-color: __NAVY__ !important; color: #fff !important; }
button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover { background: #163C61 !important; }
button:disabled { opacity: .55; }
.stFileUploader label, .stTextInput label, .stNumberInput label, .stSelectbox label, .stDateInput label, .stTextArea label {
  color: __NAVY__ !important; font-weight: 700 !important; font-size: 13.5px !important; }
[data-testid="stFileUploaderDropzone"] { background: #fff !important; border: 2px dashed #AABCD0 !important; border-radius: 16px !important; }
[data-testid="stFileUploaderDropzone"]:hover { border-color: __BLUE__ !important; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 16px; }
[data-testid="stExpander"] details { border-radius: 14px; background: #fff; }
/* La cláusula de responsabilidad debe leerse completa (sin cortar con …) */
.stCheckbox label, .stCheckbox label p, .stCheckbox [data-testid="stMarkdownContainer"] p,
.stCheckbox [data-testid="stWidgetLabel"] { white-space: normal !important; overflow: visible !important;
  text-overflow: clip !important; line-height: 1.45; }
"""


def inject_css():
    css = _CSS
    for k, v in {"__NAVY2__": NAVY2, "__NAVY__": NAVY, "__BLUE__": BLUE, "__GOLD__": GOLD, "__AMBER__": AMBER, "__INK__": INK,
                 "__MUTED__": MUTED, "__OK__": OK, "__WARN__": WARN, "__BAD__": BAD, "__BG__": BG,
                 "__AMARILLO_DUDA__": AMARILLO_DUDA, "__AMARILLO_MANUAL__": AMARILLO_MANUAL}.items():
        css = css.replace(k, v)
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def html(s: str):
    st.markdown(dedent(s).strip(), unsafe_allow_html=True)


def e(s) -> str:
    """Escapa texto de usuario antes de insertarlo en HTML (previene XSS)."""
    return escape("" if s is None else str(s))


def cop(v, vacio="—") -> str:
    try:
        if v is None or pd.isna(v):
            return vacio
        return f"$ {float(v):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return vacio


def num(v, dec=0, vacio="—") -> str:
    try:
        if v is None or pd.isna(v):
            return vacio
        s = f"{float(v):,.{dec}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return vacio


# ============================================================
# Logo y siluetas de aeronaves (SVG original, genérico)
# ============================================================
@lru_cache(maxsize=1)
def logo_b64() -> str | None:
    p = Path(__file__).resolve().parent / "assets" / "logo_moon.png"
    return base64.b64encode(p.read_bytes()).decode() if p.is_file() else None


def logo_img(ancho=190) -> str:
    b = logo_b64()
    if not b:
        return f'<div style="font-family:Manrope;font-weight:800;font-size:28px;color:{NAVY}">MOON</div>'
    return f'<img src="data:image/png;base64,{b}" alt="Moon Aerolínea" style="width:{ancho}px;max-width:100%">'


def svg_avion(tam=46, color=NAVY, rot=0, acento=AMBER) -> str:
    """Silueta vista superior de un bimotor turbohélice regional (diseño propio, genérico)."""
    return f"""<svg width="{tam}" height="{tam}" viewBox="0 0 64 64" aria-hidden="true" style="transform:rotate({rot}deg)">
<path d="M32 2 C35 2 36.2 7 36.2 12 L36.2 50 C36.2 55 34.2 61 32 62 C29.8 61 27.8 55 27.8 50 L27.8 12 C27.8 7 29 2 32 2 Z" fill="{color}"/>
<path d="M27.8 24 L3 32.5 L3 36 L27.8 31.5 Z M36.2 24 L61 32.5 L61 36 L36.2 31.5 Z" fill="{color}"/>
<rect x="14.2" y="22.5" width="4.2" height="11" rx="2" fill="{color}"/><rect x="45.6" y="22.5" width="4.2" height="11" rx="2" fill="{color}"/>
<rect x="10.8" y="21.4" width="11" height="1.7" rx=".85" fill="{acento}"/><rect x="42.2" y="21.4" width="11" height="1.7" rx=".85" fill="{acento}"/>
<path d="M27.8 52 L19 57.5 L19 59.5 L27.8 57 Z M36.2 52 L45 57.5 L45 59.5 L36.2 57 Z" fill="{color}"/>
<circle cx="32" cy="9" r="1.6" fill="{acento}"/></svg>"""


def svg_ruta() -> str:
    return f"""<svg class="hero-route" viewBox="0 0 420 170" preserveAspectRatio="xMaxYMid slice" aria-hidden="true">
<path d="M14 150 C 120 30, 270 8, 400 72" fill="none" stroke="{GOLD}" stroke-width="1.6" stroke-dasharray="3 7" stroke-linecap="round"/>
<circle cx="14" cy="150" r="4.5" fill="none" stroke="{GOLD}" stroke-width="1.6"/><circle cx="400" cy="72" r="5.5" fill="{AMBER}"/>
<circle cx="190" cy="44" r="3" fill="{GOLD}" opacity=".7"/>
<g transform="translate(262 22) rotate(118 18 18)">
<path d="M18 1.7 C19.5 1.7 20 4.5 20 7.3 L20 27.5 C20 30.3 19.1 33.7 18 34.3 C16.9 33.7 16 30.3 16 27.5 L16 7.3 C16 4.5 16.5 1.7 18 1.7 Z" fill="#fff"/>
<path d="M2.2 17.4 L16 14.3 L20 14.3 L33.8 17.4 L33.8 19.4 L20 18.3 L16 18.3 L2.2 19.4 Z" fill="#fff"/>
<rect x="8.4" y="12.4" width="2.5" height="7.3" rx="1.1" fill="#fff"/><rect x="25.1" y="12.4" width="2.5" height="7.3" rx="1.1" fill="#fff"/>
<path d="M11.8 31.2 L16.1 29.2 L19.9 29.2 L24.2 31.2 L24.2 32.6 L19.9 32 L16.1 32 L11.8 32.6 Z" fill="#fff"/></g></svg>"""


# ============================================================
# Componentes
# ============================================================
def hero(titulo_txt, texto, etiqueta=None):
    tag = f'<div class="hero-tag">{e(etiqueta)}</div>' if etiqueta else ""
    html(f'<div class="hero-panel">{svg_ruta()}{tag}<div class="hero-title">{e(titulo_txt)}</div>'
         f'<div class="hero-copy">{e(texto)}</div></div>')


def page_title(t):
    html(f'<h1 class="page-title">{e(t)}</h1>')


def kpi(label, value, delta=None, nota=None, tono=""):
    dlt = ""
    if delta is not None:
        neg = " neg" if delta < 0 else ""
        dlt = f'<div class="kpi-delta{neg}">{"▲" if delta >= 0 else "▼"} {abs(delta):.1f}% vs. periodo anterior</div>'
    nt = f'<div class="kpi-note">{e(nota)}</div>' if nota else ""
    html(f'<div class="metric-card {tono}"><div class="metric-label">{e(label)}</div>'
         f'<div class="metric-value">{e(value)}</div>{dlt}{nt}</div>')


def kpi_fila(items: list[dict]):
    cols = st.columns(len(items))
    for c, it in zip(cols, items):
        with c:
            kpi(it.get("label"), it.get("value"), it.get("delta"), it.get("nota"), it.get("tono", ""))


def titulo(t, sub=None):
    html(f'<div class="section-heading">{e(t)}</div>' + (f'<div class="section-sub">{e(sub)}</div>' if sub else ""))


def vacio(icono, t, texto):
    html(f'<div class="vacio"><span class="ico">{icono}</span><b>{e(t)}</b>{e(texto)}</div>')


def stepper(paso: int, pasos=("Foto", "Calidad", "Lectura", "Confrontación", "Guardado")):
    items = "".join(f'<div class="step {"done" if i + 1 < paso else ("on" if i + 1 == paso else "")}">{i + 1}. {e(t)}</div>'
                    for i, t in enumerate(pasos))
    html(f'<div class="stepper">{items}</div>')


def chip(texto, tono="info") -> str:
    return f'<span class="chip {tono}">{e(texto)}</span>'


def chips(lista: list[tuple[str, str]]):
    html("".join(chip(t, tn) for t, tn in lista))


def aviso_manual(texto="Dato manual / Fórmula pendiente: el sistema no bloquea; digite el valor si aplica."):
    html(f'<div class="manual">✍️ {e(texto)}</div>')


def aviso_dudoso(campos: list[str]):
    if campos:
        html(f'<div class="dudoso">🟡 <b>Revise con atención (lectura dudosa):</b> {e(", ".join(campos))}</div>')


def ok_card(titulo_txt, texto):
    html(f'<div class="ok-card"><div class="ok-icon">✓</div><div><div class="ok-title">{e(titulo_txt)}</div>'
         f'<div class="ok-copy">{e(texto)}</div></div></div>')


def texto_clausula(usuario: dict) -> str:
    return (f"Confirmo bajo mi responsabilidad (Cédula: {usuario.get('cedula')} · {usuario.get('nombre')}) que confronté "
            f"visualmente estos datos contra el soporte físico/imagen y certifico su veracidad.")


def clausula_responsabilidad(usuario: dict, key: str) -> bool:
    """Checkbox obligatorio. Los botones de guardar deben quedar deshabilitados mientras sea False."""
    html('<div class="clausula">🛡️ <b>Validación cruzada:</b> extracción automática + confrontación humana. '
         'Su cédula queda registrada en la bitácora inmutable.</div>')
    return st.checkbox(texto_clausula(usuario), key=key)


def campo_motivo(key: str, etiqueta="Motivo del cambio (obligatorio)") -> str:
    return st.text_input(etiqueta, key=key, max_chars=300, placeholder="Ej.: corrección según soporte físico")


# ============================================================
# Norte estratégico y gobernanza
# ============================================================
NORTE = {
    "Mega": "Ser el grupo aeronáutico regional más confiable de Colombia: sistematizado, auditable, rentable y diversificado.",
    "Propósito": "Unir con calidez y cumplimiento a las regiones que el país suele olvidar: el cielo nos une.",
    "Visión": "Ser el holding aeronáutico que todos quieren volar y donde todos quieren trabajar, por su cumplimiento y calidez.",
    "Misión": "Operar vuelos y servicios aeronáuticos seguros y puntuales, con equipos felices y datos confiables, creciendo de forma rentable.",
}
VALORES = ["Respeto", "Cumplimiento", "Seguridad", "Calidez humana", "Honestidad"]


def tarjetas_norte():
    tarjetas = "".join(f'<div class="norte-card{" navy" if i == 0 else ""}"><div class="norte-k">{e(k)}</div>'
                       f'<div class="norte-t">{e(v)}</div></div>' for i, (k, v) in enumerate(NORTE.items()))
    valores = "".join(f'<span class="valor">{e(v)}</span>' for v in VALORES)
    html(f'<div class="norte-grid">{tarjetas}<div class="norte-card"><div class="norte-k">Valores</div>'
         f'<div class="valores">{valores}</div></div></div>')


def diagrama_gobernanza():
    pasos = [
        ("Captura en pista", "El piloto fotografía tiquete, libro de vuelo y ficha. Contabilidad/Operaciones cargan respaldo."),
        ("Filtro de calidad", "Resolución ≥ 800 px, nitidez (Laplaciano), brillo y contraste. Foto borrosa = se repite."),
        ("Extracción IA", "Claude Vision lee por zonas ampliadas y marca en amarillo los campos dudosos."),
        ("Confrontación humana", "El responsable compara imagen vs. datos y firma con su cédula (cláusula)."),
        ("Carpetas seriadas", "AAAA/MM/DD/01-02-03 · PREFIJO_FECHA_HORA_MAT_Vnn_C|N · huella SHA-256."),
        ("Consolidados", "Cierre diario → histórico mensual → histórico anual, congelados con su hash."),
        ("Auditoría IA", "Revisoría fiscal: cruza imágenes y tablas, señala responsables y recomendaciones."),
    ]
    partes = []
    for i, (t, d) in enumerate(pasos, 1):
        partes.append(f'<div class="paso"><span class="n">{i}</span><span class="t">{e(t)}</span><div class="d">{e(d)}</div></div>')
    html(f'<div class="flujo">{"".join(partes)}</div>'
         '<div class="cimiento">🔒 <b>Cimientos:</b> bitácora <b>audit_logs</b> inmutable encadenada con SHA-256 · '
         'hash de integridad por registro · borrado lógico (nada se pierde) · respaldos automáticos diarios y antes de cada cierre · '
         'acceso por cédula y rol con contraseñas PBKDF2.</div>')


def tarjetas_flota(df: pd.DataFrame, umbral_amarillo=50.0, umbral_rojo=10.0):
    if df is None or df.empty:
        return
    items = []
    for _, a in df.iterrows():
        restante = None
        if pd.notna(a.get("proxima_inspeccion_horas")) and pd.notna(a.get("horas_totales")):
            restante = float(a["proxima_inspeccion_horas"]) - float(a["horas_totales"])
        if restante is None:
            color, txt = "#AAB7C6", "Sin datos de inspección"
        elif restante <= umbral_rojo:
            color, txt = BAD, f"{restante:,.1f} h para inspección"
        elif restante <= umbral_amarillo:
            color, txt = WARN, f"{restante:,.1f} h para inspección"
        else:
            color, txt = OK, f"{restante:,.1f} h para inspección"
        modelo = a.get("modelo_codigo") or "Sin modelo"
        items.append(f'<div class="avion">{svg_avion(46, NAVY, 90)}<div><div class="m">{e(a["matricula"])}</div>'
                     f'<div class="s">{e(modelo)}</div><div class="s"><span class="sem" style="background:{color}"></span>{e(txt)}</div>'
                     f'</div></div>')
    html(f'<div class="flota">{"".join(items)}</div>')


# ============================================================
# Plotly
# ============================================================
def estilo(fig, alto=360):
    fig.update_layout(height=alto, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", colorway=PALETA,
                      font=dict(family="DM Sans, sans-serif", color=MUTED, size=12),
                      title_font=dict(family="Manrope, sans-serif", size=15, color=NAVY), title_x=0,
                      margin=dict(l=10, r=10, t=54, b=10), legend=dict(orientation="h", y=-0.16, x=0, title_text=""),
                      hoverlabel=dict(bgcolor=NAVY, font_color="white", font_family="DM Sans, sans-serif"),
                      separators=",.")
    fig.update_xaxes(showgrid=False, linecolor="#DDE5EE", zeroline=False, title=None)
    fig.update_yaxes(gridcolor="#E9EEF5", zeroline=False, title=None)
    return fig


def mostrar(fig, vacio_msg="Aún no hay datos para este gráfico.", key=None):
    with st.container(border=True):
        if fig is None:
            st.caption(vacio_msg)
        else:
            st.plotly_chart(fig, width="stretch", config=PLOT_CFG, key=key)


def barras_h(d: pd.DataFrame, cat, val, titulo_txt, fmt=",.1f", top=12, color_val=True):
    if d is None or d.empty or cat not in d or val not in d:
        return None
    g = d.groupby(cat, as_index=False)[val].sum()
    g = g[g[val] > 0].nlargest(top, val).sort_values(val)
    if g.empty:
        return None
    fig = px.bar(g, x=val, y=cat, orientation="h", title=titulo_txt, text=val,
                 color=val if color_val else None, color_continuous_scale=ESCALA)
    fig.update_traces(texttemplate="%{text:" + fmt + "}", textposition="outside", cliponaxis=False,
                      marker_color=None if color_val else BLUE)
    fig.update_coloraxes(showscale=False)
    fig.update_xaxes(visible=False)
    fig.update_yaxes(title=None)
    fig = estilo(fig, max(280, 60 + 32 * len(g)))
    fig.update_layout(margin=dict(l=10, r=60, t=54, b=10))
    return fig


def dona(d: pd.DataFrame, cat, val, titulo_txt, centro=""):
    if d is None or d.empty:
        return None
    g = d.groupby(cat, as_index=False)[val].sum()
    g = g[g[val] > 0]
    if g.empty:
        return None
    fig = px.pie(g, names=cat, values=val, hole=0.6, title=titulo_txt, color_discrete_sequence=PALETA)
    fig.update_traces(textinfo="percent+label", textposition="outside", marker=dict(line=dict(color="white", width=2)))
    if centro:
        fig.add_annotation(text=centro, showarrow=False, font=dict(family="Manrope, sans-serif", size=15, color=NAVY))
    fig = estilo(fig)
    fig.update_layout(showlegend=False)
    return fig


# ============================================================
# Resaltado de celdas (Styler)
# ============================================================
def resaltar(df: pd.DataFrame, dudosas: set[tuple[int, str]] | None = None, pendientes_col: str | None = None,
             vacias_cols: list[str] | None = None):
    """Devuelve un Styler: amarillo intenso = lectura dudosa; amarillo suave = dato manual pendiente (vacío)."""
    dudosas = dudosas or set()
    vacias_cols = vacias_cols or []

    def _f(data):
        estilos = pd.DataFrame("", index=data.index, columns=data.columns)
        for i, idx in enumerate(data.index):
            for c in data.columns:
                v = data.at[idx, c]
                if (i, c) in dudosas:
                    estilos.at[idx, c] = f"background-color: {AMARILLO_DUDA}; font-weight: 700"
                elif c in vacias_cols and (v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip())):
                    estilos.at[idx, c] = f"background-color: {AMARILLO_MANUAL}"
        return estilos
    return df.style.apply(_f, axis=None)

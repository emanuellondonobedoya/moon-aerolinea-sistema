# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Núcleo Unificado de Operaciones, Contabilidad, Mantenimiento (OMA) y Gobernanza Gerencial

Ejecutar:   streamlit run app.py
Secretos:   ANTHROPIC_API_KEY, GERENCIA_CEDULA, GERENCIA_PASSWORD, MOON_DATA_DIR (opcional), ANTHROPIC_MODEL (opcional)
            en .streamlit/secrets.toml o variables de entorno.

Módulos (control estricto de permisos por rol en el servidor):
  0. Norte estratégico y gobernanza (todos)     3. Operaciones · archivo plano MO-A3-013 (OPERACIONES, GERENCIA)
  1. Pilotos · captura documental (PILOTO)       4. OMA · mantenimiento (OMA, GERENCIA)
  2. Contabilidad (CONTABILIDAD, GERENCIA)       5. Gerencia general (GERENCIA)
"""
from __future__ import annotations

import functools
import hashlib
import io
import json
import sqlite3
import time
import zipfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Moon Aerolínea · Núcleo Operacional", page_icon="🌙", layout="wide",
                   initial_sidebar_state="expanded")

import audit_ai as A  # noqa: E402
import database as db  # noqa: E402
import exports as X  # noqa: E402
import rules_engine as R  # noqa: E402
import ui_styles as U  # noqa: E402
import vision_ocr as V  # noqa: E402

U.inject_css()
ss = st.session_state
# Celdas vacías se ven en blanco (no "None"): coherente con la regla de "espacios en blanco editables"
_editor = functools.partial(st.data_editor, placeholder="")
_tabla = functools.partial(st.dataframe, placeholder="—")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

MODULOS = {
    "inicio": ("🌙 Norte estratégico", db.ROLES),
    "pilotos": ("✈️ Pilotos · Captura", ("PILOTO", "GERENCIA")),
    "contabilidad": ("💼 Contabilidad", ("CONTABILIDAD", "GERENCIA")),
    "operaciones": ("🛫 Operaciones", ("OPERACIONES", "GERENCIA")),
    "oma": ("🔧 OMA · Mantenimiento", ("OMA", "GERENCIA")),
    "gerencia": ("📊 Gerencia general", ("GERENCIA",)),
    "cuenta": ("👤 Mi cuenta", db.ROLES),
}


@st.cache_resource(show_spinner="Preparando la base de datos…")
def _arranque():
    return db.init_db()


try:
    _arranque()
except Exception as exc:  # pragma: no cover - fallo de entorno (disco de solo lectura, permisos)
    st.error(f"No se pudo inicializar la base de datos en {db.DATA_DIR}: {exc}")
    st.stop()


# ============================================================
# Utilidades de interfaz
# ============================================================
def flash(msg: str, errores: list | None = None):
    ss["flash"] = (msg, errores or [])


def mostrar_flash():
    if ss.get("flash"):
        msg, errores = ss.pop("flash")
        if msg:
            st.toast(msg, icon="✅")
        for e in errores:
            st.error(e)


def requiere_rol(*roles):
    u = ss.get("usuario")
    if not u or u.get("rol") not in roles:
        st.error("🔒 No tiene permisos para acceder a este módulo.")
        st.stop()


def rango_fechas(key, defecto="Este mes", opciones=("Hoy", "7 días", "30 días", "Este mes", "Mes anterior", "Este año",
                                                     "Personalizado")):
    hoy = db.hoy()
    c1, c2 = st.columns([2.4, 1.2], vertical_alignment="bottom")
    sel = c1.segmented_control("Periodo", list(opciones), default=defecto, key=f"{key}_preset", required=True) or defecto
    ini_mes = hoy.replace(day=1)
    fin_ant = ini_mes - timedelta(days=1)
    rangos = {"Hoy": (hoy, hoy), "7 días": (hoy - timedelta(days=6), hoy), "30 días": (hoy - timedelta(days=29), hoy),
              "Este mes": (ini_mes, hoy), "Mes": (ini_mes, hoy), "Mes anterior": (fin_ant.replace(day=1), fin_ant),
              "Este año": (date(hoy.year, 1, 1), hoy), "Año": (date(hoy.year, 1, 1), hoy),
              "Histórico completo": (date(2000, 1, 1), hoy)}
    if sel == "Personalizado":
        r = c2.date_input("Rango", value=(ini_mes, hoy), key=f"{key}_rango", format="DD/MM/YYYY")
        if isinstance(r, (list, tuple)):
            return (r[0], r[-1]) if len(r) else (ini_mes, hoy)
        return r, r
    d, h = rangos.get(sel, (ini_mes, hoy))
    c2.caption(f"Del {d:%d/%m/%Y} al {h:%d/%m/%Y}" if sel != "Histórico completo" else "Todo el histórico")
    return d, h


def selector_mes(key, etiqueta="Mes", n=24):
    hoy = db.hoy().replace(day=1)
    meses = []
    d = hoy
    for _ in range(n):
        meses.append(d.strftime("%Y-%m"))
        d = (d - timedelta(days=1)).replace(day=1)
    return st.selectbox(etiqueta, meses, key=key, format_func=lambda p: f"{R.MESES_ES[int(p[5:])]} {p[:4]}")


def limites_mes(periodo: str) -> tuple[date, date]:
    a, m = int(periodo[:4]), int(periodo[5:])
    ini = date(a, m, 1)
    fin = (date(a + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
    return ini, fin


def visor(ruta_rel, caption=None):
    raw = V.leer_evidencia(ruta_rel) if not R.es_vacio(ruta_rel) else None
    if not raw:
        U.vacio("🖼️", "Sin imagen de soporte", "Este registro no tiene evidencia asociada o el archivo no está disponible.")
        return
    try:
        V.abrir_imagen(raw)
    except V.ImagenInvalida:
        st.error("⚠️ La evidencia existe pero está dañada o fue alterada (no es una imagen válida). "
                 "Ejecute la auditoría general y restaure el archivo desde el respaldo.")
        return
    st.image(raw, caption=caption or Path(str(ruta_rel)).name, width="stretch")


def a_objeto(df: pd.DataFrame, cols) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(object).where(df[c].notna(), None)
    return df


def a_numero(df: pd.DataFrame, cols) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    return df


def guardar_con_clausula(usuario, key, etiqueta="💾 Guardar", motivo=True, extra_ok=True, tipo="primary") -> tuple[bool, str]:
    """Bloque estándar: (motivo) + cláusula + botón deshabilitado hasta cumplir. Devuelve (pulsado, motivo)."""
    txt = U.campo_motivo(f"{key}_mot") if motivo else "Registro nuevo"
    ok = U.clausula_responsabilidad(usuario, f"{key}_cl")
    habil = ok and extra_ok and (not motivo or len(txt.strip()) >= 5)
    pulsado = st.button(etiqueta, type=tipo, disabled=not habil, key=f"{key}_btn", width="stretch")
    if motivo and not len(txt.strip()) >= 5:
        st.caption("Escriba el motivo (mínimo 5 caracteres) y marque la cláusula para habilitar el botón.")
    return pulsado, txt


def editor_auditado(df_vista: pd.DataFrame, *, key: str, tabla: str, mapa: dict, modulo: str, usuario: dict,
                    column_config: dict | None = None, height=None, column_order=None):
    """Tabla editable con trazabilidad: solo columnas de 'mapa' son editables; exige motivo + cláusula."""
    version = ss.get(f"{key}_v", 0)
    k = f"{key}_{version}"
    deshab = [c for c in df_vista.columns if c not in mapa]
    _editor(df_vista, key=k, hide_index=True, width="stretch", num_rows="fixed", disabled=deshab,
                   column_config=column_config or {}, height=height or "auto", column_order=column_order)
    filas = (ss.get(k) or {}).get("edited_rows") or {}
    cambios: dict[int, dict] = {}
    for idx, cols in filas.items():
        rid = int(df_vista.iloc[int(idx)]["ID"])
        for col, val in cols.items():
            if col in mapa:
                cambios.setdefault(rid, {})[mapa[col]] = val
    if not cambios:
        st.caption("✏️ Doble clic en una celda editable para corregirla. Se pedirá motivo y cláusula de responsabilidad.")
        return
    n = sum(len(v) for v in cambios.values())
    with st.container(border=True):
        st.markdown(f"**{n} cambio(s) sin guardar en {len(cambios)} registro(s).**")
        motivo = U.campo_motivo(f"{k}_motivo")
        ok = U.clausula_responsabilidad(usuario, f"{k}_cl")
        c1, c2 = st.columns(2)
        guardar = c1.button("💾 Guardar cambios auditados", type="primary", key=f"{k}_save", width="stretch",
                            disabled=not (ok and len(motivo.strip()) >= 5))
        if c2.button("↩️ Descartar cambios", key=f"{k}_desc", width="stretch"):
            ss[f"{key}_v"] = version + 1
            st.rerun()
    if guardar:
        errores, total = [], 0
        for rid, camb in cambios.items():
            ok2, msg, n2 = db.actualizar_registro(tabla, rid, camb, usuario, motivo, modulo)
            if ok2:
                total += n2
            else:
                errores.append(msg)
        ss[f"{key}_v"] = version + 1
        flash(f"{total} campo(s) guardados con trazabilidad.", errores)
        st.rerun()


def bloque_eliminar(tabla, ids: list, etiquetas: dict, usuario, modulo, key):
    """Borrado lógico controlado (el dato queda en la base y en la bitácora)."""
    with st.expander("🗑️ Eliminar un registro (borrado lógico auditado)"):
        if not ids:
            st.caption("No hay registros.")
            return
        rid = st.selectbox("Registro", ids, format_func=lambda i: etiquetas.get(i, str(i)), key=f"{key}_sel")
        pulsado, motivo = guardar_con_clausula(usuario, f"{key}_del", "🗑️ Eliminar registro", tipo="secondary")
        if pulsado:
            ok, msg = db.eliminar_registro(tabla, int(rid), usuario, motivo, modulo)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()


# ============================================================
# Sesión, login y cambio obligatorio de clave
# ============================================================
def cerrar_sesion(motivo=None):
    u = ss.get("usuario")
    if u:
        try:
            db.auditar(u, "SEGURIDAD", "usuarios", u["cedula"], "LOGOUT", motivo=motivo or "Cierre de sesión")
        except Exception:
            pass
    for k in list(ss.keys()):
        del ss[k]
    if motivo:
        ss["aviso_login"] = motivo


def validar_sesion() -> dict | None:
    u = ss.get("usuario")
    if not u:
        return None
    limite = int(db.get_param("sesion_minutos", 30) or 30) * 60
    if time.time() - ss.get("ultimo_mov", time.time()) > limite:
        cerrar_sesion("La sesión se cerró por inactividad.")
        return None
    fresco = db.obtener_usuario(u["cedula"])
    if not fresco or not fresco.get("activo") or fresco.get("rol") != u.get("rol"):
        cerrar_sesion("Sus permisos cambiaron o la cuenta fue desactivada. Ingrese de nuevo.")
        return None
    ss["usuario"] = fresco
    ss["ultimo_mov"] = time.time()
    return fresco


def vista_login():
    c1, c2 = st.columns([1.15, 1], gap="large")
    with c1:
        U.html(f'<div class="login-panel"><div class="login-logo">{U.logo_img(210)}</div>'
               '<div class="lt">Núcleo Unificado de Operaciones</div>'
               '<div class="lc">Captura documental con IA, contabilidad, archivo plano Aerocivil, mantenimiento OMA y '
               'gobernanza gerencial en un solo lugar. Cada dato queda firmado con la cédula de quien lo confirma.</div>'
               f'<div class="login-avion">{U.svg_avion(210, "rgba(255,255,255,.10)", 35, "rgba(244,180,26,.55)")}</div></div>')
    with c2:
        U.html('<div style="height:26px"></div>')
        st.markdown("### Ingreso seguro")
        st.caption("Ingrese con su número de cédula y contraseña personal.")
        if ss.get("aviso_login"):
            st.info(ss.pop("aviso_login"))
        with st.form("login", border=True):
            cedula = st.text_input("Número de cédula", max_chars=12, placeholder="Solo números")
            clave = st.text_input("Contraseña", type="password", max_chars=128)
            entrar = st.form_submit_button("Ingresar", type="primary", width="stretch")
        if entrar:
            with st.spinner("Verificando credenciales…"):
                u, msg = db.autenticar(cedula, clave)
            if u:
                ss["usuario"], ss["ultimo_mov"] = u, time.time()
                st.rerun()
            st.error(msg)
        ok, msg_ia = V.ia_disponible()
        st.caption(("🤖 " if ok else "✍️ ") + msg_ia)


def vista_cambio_clave(u, obligatorio=True):
    if obligatorio:
        U.page_title("Establezca su contraseña personal")
        st.warning("Por seguridad debe cambiar la clave temporal antes de usar el sistema.")
    with st.form("cambio_clave", border=True):
        actual = st.text_input("Contraseña actual / temporal", type="password", max_chars=128)
        nueva = st.text_input("Nueva contraseña", type="password", max_chars=128)
        confirma = st.text_input("Confirmar nueva contraseña", type="password", max_chars=128)
        st.caption("Mínimo 10 caracteres con mayúscula, minúscula, número y símbolo. No use su cédula ni su nombre.")
        enviar = st.form_submit_button("Guardar contraseña", type="primary", width="stretch")
    if enviar:
        if nueva != confirma:
            st.error("La confirmación no coincide.")
            return
        ok, errores = db.cambiar_clave(u["cedula"], actual, nueva)
        if ok:
            ss["usuario"] = db.obtener_usuario(u["cedula"])
            flash("Contraseña actualizada.")
            st.rerun()
        for e in errores:
            st.error(e)


# ============================================================
# Componente reutilizable de captura documental con IA
# ============================================================
ETIQ_CAPTURA = {"COMB": "⛽ Tiquete de Combustible Terpel", "LV": "📘 Libro de Vuelo MO-A3-004",
                "FV": "🧾 Ficha de Vuelo / Cierre y Manifiesto P&B"}
CONSEJOS = {"COMB": "Encuadra el tiquete completo, sin dedos sobre los números de recibo y galones.",
            "LV": "Foto de frente a la hoja completa del libro, incluida la fila TOTALES y el bloque de mantenimiento.",
            "FV": "Incluye pasajeros, tripadis, bebés, tránsitos, carga y combustible básico del manifiesto."}
TRAMO_COLS = [("VUELO", "numero_vuelo", "t"), ("ORIGEN", "origen", "t"), ("DESTINO", "destino", "t"),
              ("GAL TANQ.", "galones_tanqueados", "n"), ("RECIBO COMB.", "recibo_combustible", "t"),
              ("LBS SALIDA", "lbs_comb_salida", "n"), ("PAX TOTAL", "pax_total", "n"), ("PAX TTO", "pax_transito", "n"),
              ("H. SALIDA", "hora_salida_min", "h"), ("H. DESPEGUE", "hora_despegue_min", "h"),
              ("H. ATERRIZAJE", "hora_aterrizaje_min", "h"), ("H. LLEGADA", "hora_llegada_min", "h"),
              ("T. VUELO", "tiempo_vuelo_min", "d"), ("T. NOCTURNO", "tiempo_nocturno_min", "d"),
              ("T. INSTRUM.", "tiempo_instrumentos_min", "d"), ("T. TOTAL", "tiempo_total_min", "d"),
              ("LBS LLEGADA", "lbs_comb_llegada", "n"), ("LBS USADO", "lbs_comb_usado", "n")]
MANT_COLS = [("ITEM", "item"), ("REPORTE", "reporte"), ("ACCIÓN CORRECTIVA", "accion_correctiva"), ("P/N ON", "pn_on"),
             ("S/N ON", "sn_on"), ("P/N OFF", "pn_off"), ("S/N OFF", "sn_off"), ("POSICIÓN", "posicion"), ("RII", "rii"),
             ("TÉCNICO", "tecnico_nombre"), ("LICENCIA TÉC.", "tecnico_licencia"), ("BASE", "base")]
FICHA_COLS = [("FECHA", "fecha", "f"), ("MATRÍCULA", "matricula", "t"), ("VUELO", "numero_vuelo", "t"), ("ORIGEN", "origen", "t"),
              ("DESTINO", "destino", "t"), ("HORA", "hora_min", "h"), ("PAX TOTAL", "pax_total", "n"), ("TRIPADIS", "tripadis", "n"),
              ("BEBÉS", "bebes", "n"), ("PAX TRÁNSITO", "pax_transito", "n"), ("CARGA A BORDO (KG)", "carga_a_bordo_kg", "n"),
              ("CARGA TRÁNSITO (KG)", "carga_transito_kg", "n"), ("COMB. BÁSICO (LBS)", "combustible_basico_lbs", "n"),
              ("PILOTO", "piloto_nombre", "t")]


def _fmt_celda(v, tipo):
    if tipo == "h":
        return R.fmt_hora(v) or None
    if tipo == "d":
        return R.fmt_duracion(v) or None
    if tipo == "f":
        d = R.parse_fecha(v)
        return d.strftime("%d/%m/%Y") if d else None
    return v


def _parse_celda(v, tipo):
    if tipo == "h":
        return R.parse_hora(v)
    if tipo == "d":
        return R.parse_duracion(v)
    if tipo == "f":
        return R.parse_fecha(v)
    if tipo == "n":
        return R.a_float(v)
    return R.texto(v) or None


def _df_desde_lista(filas, cols, dudosos_prefijo=None, dudosos=None):
    datos = {et: [_fmt_celda(f.get(campo), tp) for f in filas] for et, campo, tp in cols}
    df = pd.DataFrame(datos, columns=[c[0] for c in cols])
    numericas = [et for et, _, tp in cols if tp == "n"]
    df = a_numero(df, numericas)
    df = a_objeto(df, [et for et, _, tp in cols if tp != "n"])
    if dudosos_prefijo is not None:
        revisar = []
        for i in range(len(df)):
            campos = [d.split(".", 1)[1] for d in (dudosos or []) if d.startswith(f"{dudosos_prefijo}[{i}].")]
            etq = {campo: et for et, campo, _ in cols}
            revisar.append(", ".join(etq.get(c, c) for c in campos) or None)
        df["REVISAR"] = pd.Series(revisar, dtype=object)
    return df


def _lista_desde_df(df: pd.DataFrame, cols) -> list[dict]:
    out = []
    for _, r in df.iterrows():
        fila = {campo: _parse_celda(r.get(et), tp) for et, campo, tp in cols}
        if any(v is not None for v in fila.values()):
            out.append(fila)
    return out


def _cfg_editor(cols, editable=True) -> dict:
    cfg = {}
    for c in cols:
        et, campo, tp = c
        if tp == "n":
            cfg[et] = st.column_config.NumberColumn(et, format="%.1f" if "LBS" not in et and "PAX" not in et else "%d")
        elif tp in ("h", "d"):
            cfg[et] = st.column_config.TextColumn(et, help="Formato H:MM (acepta 0555, 5:55, 5h55)", max_chars=8)
        elif tp == "f":
            cfg[et] = st.column_config.TextColumn(et, help="DD/MM/AAAA", max_chars=10)
        else:
            cfg[et] = st.column_config.TextColumn(et)
    cfg["REVISAR"] = st.column_config.TextColumn("🟡 REVISAR", disabled=True, help="Campos leídos con duda por la IA")
    return cfg


def _etq(label, campo, dudosos, valor):
    if any(d == campo or d.endswith("." + campo) or d.startswith(campo) for d in dudosos):
        return f"🟡 {label}"
    if R.es_vacio(valor):
        return f"✍️ {label}"
    return label


def _correcciones(ia: dict | None, final: dict, prefijo="") -> list[dict]:
    if not ia:
        return []
    out = []
    for k, v in final.items():
        a = ia.get(k)
        if isinstance(a, date):
            a = a.isoformat()
        b = v.isoformat() if isinstance(v, date) else v
        if R.texto(a) != R.texto(b) and not (R.a_float(a) is not None and R.a_float(b) is not None and abs(R.a_float(a) - R.a_float(b)) < 1e-6):
            out.append({"campo": prefijo + k, "ia": a, "humano": b})
    return out


def _guardar_documento(tipo, estado, fecha, hora_min, matricula, charter, registrar_fn):
    """Seriado + escritura atómica + transacción en BD. Si la BD falla, se elimina el archivo (sin huérfanos)."""
    seq = db.siguiente_seq(tipo, fecha, matricula)
    ev = V.guardar_evidencia(estado["jpeg"], tipo, fecha, hora_min, matricula, seq, charter)
    cal = estado["calidad"]
    doc = {"tipo": tipo, "fecha_doc": fecha, "hora_min": hora_min, "matricula": matricula, "seq_dia": ev["seq"],
           "es_charter": charter, "ruta_archivo": ev["ruta_archivo"], "sha256": ev["sha256"], "ancho": cal["ancho"],
           "alto": cal["alto"], "nitidez": cal["nitidez"], "extraccion_ia": estado.get("crudo"),
           "campos_dudosos": estado.get("dudosos")}
    try:
        rid = registrar_fn(doc)
    except sqlite3.IntegrityError:
        V.eliminar_evidencia(ev["ruta_archivo"])
        raise ValueError("Esta imagen ya fue registrada anteriormente (huella SHA-256 duplicada).") from None
    except Exception:
        V.eliminar_evidencia(ev["ruta_archivo"])
        raise
    return rid, ev


def render_captura_documental(usuario: dict, tipos_permitidos: list[str], origen_carga: str, key_prefix: str):
    ia_ok, ia_msg = V.ia_disponible()
    if not ia_ok:
        st.warning("🤖 " + ia_msg + " El registro manual asistido sigue disponible: la operación no se detiene.")
    tabs = st.tabs([f"{i}. {ETIQ_CAPTURA[t]}" for i, t in enumerate(tipos_permitidos, 1)])
    for tab, tipo in zip(tabs, tipos_permitidos):
        with tab:
            _captura_tipo(usuario, tipo, origen_carga, f"{key_prefix}_{tipo}", ia_ok)


def _captura_tipo(usuario, tipo, origen_carga, k, ia_ok):
    sk = f"cap_{k}"
    estado = ss.get(sk) or {"n": 0}
    n = estado.get("n", 0)
    if estado.get("guardado"):
        U.stepper(5)
        g = estado["guardado"]
        U.ok_card(g["titulo"], g["texto"])
        for a in g.get("avisos", []):
            st.info(a)
        if st.button("Registrar otro documento", type="primary", key=f"{k}_otro"):
            ss[sk] = {"n": n + 1}
            st.rerun()
        return
    col_img, col_form = st.columns([1, 1.35], gap="large")
    with col_img:
        modo = st.segmented_control("Origen", ["📁 Subir archivo", "📷 Cámara"], default="📁 Subir archivo",
                                    key=f"{k}_origen_{n}", required=True, label_visibility="collapsed")
        if modo == "📷 Cámara":
            archivo = st.camera_input("Foto del documento", key=f"{k}_cam_{n}")
        else:
            archivo = st.file_uploader(ETIQ_CAPTURA[tipo], type=["jpg", "jpeg", "png", "webp"], key=f"{k}_up_{n}")
    if archivo is None:
        with col_form:
            U.stepper(1)
            U.vacio(V.TIPOS_DOC[tipo]["icono"], "Sube o toma la foto del documento", CONSEJOS[tipo])
        ss[sk] = {"n": n}
        return
    raw = archivo.getvalue()
    huella = hashlib.sha256(raw).hexdigest()
    if estado.get("huella") != huella:
        try:
            img = V.abrir_imagen(raw)
        except V.ImagenInvalida as exc:
            with col_form:
                st.error(f"📸 {exc}")
            return
        p = db.params_dict()
        calidad = V.evaluar_calidad(img, p.get("ia_min_lado_px") or 800, p.get("ia_umbral_nitidez") or 45)
        jpeg = V.jpeg_bytes(img)
        estado = {"n": n, "huella": huella, "jpeg": jpeg, "calidad": calidad, "datos": None, "ia": None, "dudosos": [],
                  "ronda": 0, "duplicado": db.documento_por_hash(hashlib.sha256(jpeg).hexdigest()),
                  "preview": V.jpeg_bytes(V.recorte_documento(img)[0], 1600, 85)}
        ss[sk] = estado
    with col_img:
        st.image(estado.get("preview") or estado["jpeg"], caption="Soporte para confrontación visual (se archiva la foto original)",
                 width="stretch")
    with col_form:
        cal = estado["calidad"]
        if estado.get("duplicado"):
            U.stepper(2)
            st.error(f"Esta misma imagen ya está registrada: {Path(estado['duplicado']['ruta_archivo']).name}. "
                     "No se permite duplicar soportes.")
            return
        if not cal["ok"]:
            U.stepper(2)
            st.error("📸 La foto no pasa el control de calidad. Repítela:")
            for m in cal["motivos"]:
                st.markdown(f"- {m}")
            st.caption(f"Nitidez {cal['nitidez']} · brillo {cal['brillo']} · contraste {cal['contraste']} · {cal['ancho']}×{cal['alto']} px")
            if st.button("📸 Tomar otra foto", key=f"{k}_retomar"):
                ss[sk] = {"n": n + 1}
                st.rerun()
            return
        if estado.get("datos") is None:
            U.stepper(3)
            st.success(f"✅ Calidad aprobada · nitidez {cal['nitidez']} · {cal['ancho']}×{cal['alto']} px")
            if estado.get("error_ia"):
                st.error(f"🤖 {estado['error_ia']}")
                if not estado.get("ia_fallo"):
                    if st.button("📸 Tomar otra foto", key=f"{k}_retomar2", type="primary"):
                        ss[sk] = {"n": n + 1}
                        st.rerun()
            b1, b2 = st.columns([1.3, 1])
            leer = b1.button("🔍 Leer con IA", type="primary", disabled=not ia_ok, key=f"{k}_leer", width="stretch")
            manual = b2.button("✍️ Registrar manualmente", key=f"{k}_manual", width="stretch")
            if manual:
                estado.update({"datos": _datos_vacios(tipo), "ia": None, "dudosos": [], "ronda": estado.get("ronda", 0) + 1,
                               "nota_ia": "Registro manual" + (" (IA reportó imagen ilegible)" if estado.get("error_ia") else "")})
                ss[sk] = estado
                st.rerun()
            if leer:
                with st.spinner("La IA está leyendo el documento por zonas…" if tipo == "LV" else "La IA está leyendo el documento…"):
                    if tipo == "COMB":
                        r = V.leer_tiquete(estado["jpeg"])
                    elif tipo == "LV":
                        r = V.leer_libro(V.abrir_imagen(estado["jpeg"]))
                    else:
                        r = V.leer_ficha(estado["jpeg"])
                if r.get("exito"):
                    datos = _datos_desde_ia(tipo, r)
                    estado.update({"datos": datos, "ia": json.loads(json.dumps(datos, default=str)), "dudosos": r.get("dudosos", []),
                                   "crudo": r.get("crudo"), "ronda": estado.get("ronda", 0) + 1, "error_ia": None})
                else:
                    estado.update({"error_ia": r.get("motivo_error"), "ia_fallo": r.get("ia_fallo", False)})
                ss[sk] = estado
                st.rerun()
            return
        U.stepper(4)
        U.aviso_dudoso(estado.get("dudosos") or [])
        if tipo == "COMB":
            _form_combustible(usuario, estado, sk, k, origen_carga)
        elif tipo == "LV":
            _form_libro(usuario, estado, sk, k, origen_carga)
        else:
            _form_ficha(usuario, estado, sk, k, origen_carga)


def _datos_vacios(tipo):
    if tipo == "COMB":
        return {"recibo": None, "numero_boleto": None, "fecha_consumo": db.hoy(), "hora_min": None, "estacion": None,
                "matricula": None, "producto": "JET A1", "destino": None, "numero_vuelo": None, "galones": None,
                "piloto_nombre": None, "piloto_cedula": None, "observacion": None}
    if tipo == "LV":
        return {"cabecera": {"numero_libro": None, "fecha": db.hoy(), "matricula": None, "piloto_nombre": None,
                             "piloto_licencia": None, "copiloto_nombre": None, "copiloto_licencia": None, "horas_totales": None,
                             "ciclos_totales": None, "proxima_inspeccion": None, "observaciones": None},
                "totales": {}, "tramos": [{}], "mantenimiento": [], "parametros_motor": []}
    return {"vuelos": [{"fecha": db.hoy()}], "observaciones": None}


def _datos_desde_ia(tipo, r):
    if tipo == "COMB":
        return r["datos"]
    if tipo == "LV":
        return {k: r.get(k) for k in ("cabecera", "totales", "tramos", "mantenimiento", "parametros_motor")}
    return {"vuelos": r.get("vuelos") or [], "observaciones": r.get("observaciones")}


def _form_combustible(usuario, estado, sk, k, origen_carga):
    d, dud, kk = estado["datos"], estado.get("dudosos") or [], f"{k}_{estado['ronda']}"
    U.titulo("Confronte los datos con el tiquete", "Amarillo 🟡 = lectura dudosa · ✍️ = dato vacío para digitar si aplica.")
    c1, c2 = st.columns(2)
    recibo = c1.text_input(_etq("Número de tiquete / recibo", "recibo", dud, d.get("recibo")), d.get("recibo") or "", key=f"{kk}_rec")
    boleto = c2.text_input(_etq("Número de boleto (medidor)", "numero_boleto", dud, d.get("numero_boleto")), d.get("numero_boleto") or "", key=f"{kk}_bol")
    fecha_ini = R.parse_fecha(d.get("fecha_consumo")) or db.hoy()
    fecha = c1.date_input(_etq("Fecha del suministro", "fecha", dud, d.get("fecha_consumo")), min(fecha_ini, db.hoy()),
                          min_value=db.hoy() - timedelta(days=400), max_value=db.hoy(), key=f"{kk}_fec", format="DD/MM/YYYY")
    hora = c2.text_input(_etq("Hora de inicio (HH:MM)", "hora_inicio", dud, d.get("hora_min")), R.fmt_hora(d.get("hora_min")), key=f"{kk}_hor")
    matricula = c1.text_input(_etq("Matrícula", "matricula", dud, d.get("matricula")), d.get("matricula") or "", key=f"{kk}_mat", placeholder="HK-4350")
    estacion = c2.text_input(_etq("Estación (IATA)", "estacion", dud, d.get("estacion")), d.get("estacion") or "", key=f"{kk}_est", placeholder="EOH")
    producto = c1.selectbox("Producto", R.PRODUCTOS, index=R.PRODUCTOS.index(d.get("producto") or "JET A1"), key=f"{kk}_prod")
    galones = c2.number_input(_etq("Galones entregados", "galones", dud, d.get("galones")), min_value=0.0, max_value=20000.0,
                              value=float(d.get("galones") or 0.0), step=0.1, format="%.1f", key=f"{kk}_gal")
    destino = c1.text_input(_etq("Destino", "destino", dud, d.get("destino")), d.get("destino") or "", key=f"{kk}_des")
    vuelo = c2.text_input(_etq("Número de vuelo", "numero_vuelo", dud, d.get("numero_vuelo")), d.get("numero_vuelo") or "", key=f"{kk}_vue")
    piloto = c1.text_input(_etq("Piloto o representante", "piloto_nombre", dud, d.get("piloto_nombre")),
                           d.get("piloto_nombre") or (usuario["nombre"] if usuario["rol"] == "PILOTO" else ""), key=f"{kk}_pil")
    ced_pil = c2.text_input("Identificación del piloto", d.get("piloto_cedula") or (usuario["cedula"] if usuario["rol"] == "PILOTO" else ""),
                            key=f"{kk}_cpi")
    obs = st.text_area("Observación", estado.get("nota_ia") or "", key=f"{kk}_obs", max_chars=500)
    mat_n = R.norm_matricula(matricula)
    errores = []
    if galones <= 0:
        errores.append("Los galones deben ser mayores que cero.")
    if not mat_n or not mat_n.startswith("HK-"):
        errores.append("Escriba una matrícula válida (HK-1234).")
    if not R.texto(recibo):
        errores.append("Escriba el número del tiquete.")
    if R.texto(hora) and R.parse_hora(hora) is None:
        errores.append("Hora no válida (use HH:MM).")
    if R.texto(recibo) and mat_n and db.consulta("SELECT id FROM combustible_tiquetes WHERE eliminado=0 AND REPLACE(recibo,' ','')=? "
                                                  "AND matricula=?", (recibo.replace(" ", ""), mat_n)):
        errores.append(f"El tiquete {recibo} ya está registrado para {mat_n}.")
    for e in errores:
        st.warning(e)
    ok = U.clausula_responsabilidad(usuario, f"{kk}_cl")
    if st.button("✅ Confirmar y Guardar", type="primary", disabled=not ok or bool(errores), key=f"{kk}_save", width="stretch"):
        final = {"recibo": R.texto(recibo), "numero_boleto": R.texto(boleto) or None, "fecha_consumo": fecha,
                 "hora_min": R.parse_hora(hora), "estacion": R.norm_iata(estacion) or None, "matricula": mat_n,
                 "producto": producto, "destino": R.norm_iata(destino) or None, "numero_vuelo": R.solo_digitos(vuelo) or None,
                 "galones": galones, "piloto_nombre": R.texto(piloto) or None, "piloto_cedula": R.solo_digitos(ced_pil) or None,
                 "observacion": R.texto(obs) or None}
        _, charter = R.normalizar_nro_vuelo(final["numero_vuelo"], mat_n)
        corr = _correcciones(estado.get("ia"), final)
        try:
            rid, ev = _guardar_documento("COMB", estado, fecha, final["hora_min"] or 0, mat_n, charter,
                                         lambda doc: db.registrar_tiquete(final, doc, usuario, origen_carga, corr))
        except ValueError as exc:
            st.error(str(exc))
            return
        estado["guardado"] = {"titulo": "Tiquete registrado con trazabilidad",
                              "texto": f"{mat_n} · {galones:,.1f} gal · recibo {final['recibo']} · archivo {ev['nombre']} · registro #{rid}"}
        ss[sk] = estado
        st.rerun()


def _form_libro(usuario, estado, sk, k, origen_carga):
    d, dud, kk = estado["datos"], estado.get("dudosos") or [], f"{k}_{estado['ronda']}"
    cab = d.get("cabecera") or {}
    U.titulo("Encabezado y tripulación", "Confronte cada casilla con la hoja del libro.")
    c1, c2, c3 = st.columns(3)
    numero = c1.text_input(_etq("Libro Nº", "numero_libro", dud, cab.get("numero_libro")), cab.get("numero_libro") or "", key=f"{kk}_num")
    fecha_ini = R.parse_fecha(cab.get("fecha")) or db.hoy()
    fecha = c2.date_input(_etq("Fecha", "fecha", dud, cab.get("fecha")), min(fecha_ini, db.hoy()), min_value=db.hoy() - timedelta(days=400),
                          max_value=db.hoy(), key=f"{kk}_fec", format="DD/MM/YYYY")
    matricula = c3.text_input(_etq("Matrícula (A/C REG)", "matricula", dud, cab.get("matricula")), cab.get("matricula") or "", key=f"{kk}_mat")
    c1, c2, c3 = st.columns(3)
    piloto = c1.text_input(_etq("Piloto al mando (PTL)", "piloto", dud, cab.get("piloto_nombre")),
                           cab.get("piloto_nombre") or (usuario["nombre"] if usuario["rol"] == "PILOTO" else ""), key=f"{kk}_pil")
    lic_p = c2.text_input("Licencia piloto", cab.get("piloto_licencia") or "", key=f"{kk}_lp")
    ced_p = c3.text_input("Cédula piloto (si difiere de la licencia)", usuario["cedula"] if usuario["rol"] == "PILOTO" else "", key=f"{kk}_cp")
    copiloto = c1.text_input(_etq("Copiloto (PCA)", "copiloto", dud, cab.get("copiloto_nombre")), cab.get("copiloto_nombre") or "", key=f"{kk}_cop")
    lic_c = c2.text_input("Licencia copiloto", cab.get("copiloto_licencia") or "", key=f"{kk}_lc")
    ced_c = c3.text_input("Cédula copiloto", "", key=f"{kk}_cc")
    with st.expander("Horas, ciclos y próxima inspección de la aeronave (opcional)"):
        h1, h2, h3 = st.columns(3)
        horas_tot = h1.text_input("Horas totales", R.texto(cab.get("horas_totales")), key=f"{kk}_ht")
        ciclos_tot = h2.text_input("Ciclos totales", R.texto(cab.get("ciclos_totales")), key=f"{kk}_ct")
        prox = h3.text_input("Próxima inspección", R.texto(cab.get("proxima_inspeccion")), key=f"{kk}_pi")

    U.titulo("Tramos del día", "Horas en H:MM (hora Colombia). Puede agregar o quitar filas. Las celdas vacías que se puedan deducir se calculan.")
    if f"{kk}_dft" not in estado:
        estado[f"{kk}_dft"] = _df_desde_lista(d.get("tramos") or [{}], TRAMO_COLS, "tramos", dud)
        ss[sk] = estado
    ed_tramos = _editor(estado[f"{kk}_dft"], key=f"{kk}_tramos", num_rows="dynamic", hide_index=True, width="stretch",
                               column_config=_cfg_editor(TRAMO_COLS))
    tramos = _lista_desde_df(ed_tramos.drop(columns=["REVISAR"], errors="ignore"), TRAMO_COLS)
    calculados = []
    avisos = []
    for i, t in enumerate(tramos):
        for c in R.completar_tramo(t):
            calculados.append(f"tramo {i + 1}: {c.replace('_min', '').replace('_', ' ')}")
        avisos += R.validar_tramo(t, i)
    if calculados:
        U.aviso_manual("Calculado automáticamente (celdas vacías deducibles): " + ", ".join(calculados))
    for a in avisos:
        st.warning(a)

    U.titulo("Validación matemática contra la fila TOTALES")
    tot = d.get("totales") or {}
    t1, t2, t3, t4, t5 = st.columns(5)
    tv = t1.text_input("TOTAL t. vuelo", R.fmt_duracion(tot.get("tot_tiempo_vuelo_min")), key=f"{kk}_tv")
    tn = t2.text_input("TOTAL t. nocturno", R.fmt_duracion(tot.get("tot_tiempo_nocturno_min")), key=f"{kk}_tn")
    ti = t3.text_input("TOTAL t. instrum.", R.fmt_duracion(tot.get("tot_tiempo_instrumentos_min")), key=f"{kk}_ti")
    tt = t4.text_input("TOTAL t. total", R.fmt_duracion(tot.get("tot_tiempo_total_min")), key=f"{kk}_tt")
    tl = t5.text_input("TOTAL lbs usado", R.texto(tot.get("tot_lbs_usado")), key=f"{kk}_tl")
    totales = {"tiempo_vuelo_min": R.parse_duracion(tv), "tiempo_nocturno_min": R.parse_duracion(tn),
               "tiempo_instrumentos_min": R.parse_duracion(ti), "tiempo_total_min": R.parse_duracion(tt),
               "lbs_comb_usado": R.a_float(tl)}
    validacion = R.validar_totales_libro(tramos, totales)
    if validacion:
        vdf = pd.DataFrame(validacion)[["CONCEPTO", "SUMA TRAMOS", "FILA TOTALES", "ESTADO"]]
        _tabla(vdf.style.apply(lambda r: [f"background-color: {'#F6DEDC' if r['ESTADO'] == 'DESCUADRE' else ('#DCEBE3' if r['ESTADO'] == 'OK' else U.AMARILLO_MANUAL)}"] * len(r), axis=1),
                     hide_index=True, width="stretch")
    descuadre = any(v["ESTADO"] == "DESCUADRE" for v in validacion)
    ack, obs_desc = True, ""
    if descuadre:
        st.error("La suma vertical de los tramos no coincide con la fila TOTALES. Corrija el dato mal leído/digitado. "
                 "Si el libro físico tiene el error, déjelo documentado para auditoría:")
        ack = st.checkbox("El libro físico presenta el descuadre: registrar la discrepancia para auditoría", key=f"{kk}_ack")
        obs_desc = st.text_input("Explique el descuadre (obligatorio)", key=f"{kk}_obsdesc") if ack else ""
        ack = ack and len(obs_desc.strip()) >= 5

    U.titulo("Mantenimiento (bloque OMA del libro)", "Deje la tabla vacía si no hubo reportes.")
    if f"{kk}_dfm" not in estado:
        filas = d.get("mantenimiento") or []
        estado[f"{kk}_dfm"] = a_objeto(pd.DataFrame({et: [f.get(c) for f in filas] for et, c in MANT_COLS},
                                                    columns=[c[0] for c in MANT_COLS]), [c[0] for c in MANT_COLS])
        ss[sk] = estado
    ed_mant = _editor(estado[f"{kk}_dfm"], key=f"{kk}_mant", num_rows="dynamic", hide_index=True, width="stretch",
                             column_config={"RII": st.column_config.SelectboxColumn("RII", options=["SI", "NO"])})
    obs = st.text_area("Observaciones de la tripulación", R.texto(cab.get("observaciones")) or estado.get("nota_ia") or "", key=f"{kk}_obs")

    mat_n = R.norm_matricula(matricula)
    errores = []
    if not mat_n.startswith("HK-"):
        errores.append("Escriba una matrícula válida (HK-1234).")
    if not tramos:
        errores.append("Registre al menos un tramo.")
    if not R.texto(piloto):
        errores.append("Escriba el piloto al mando.")
    if R.solo_digitos(numero) and mat_n and db.consulta("SELECT id FROM libros_vuelo WHERE eliminado=0 AND numero_libro=? AND matricula=?",
                                                        (R.solo_digitos(numero), mat_n)):
        errores.append(f"El libro Nº {numero} ya está registrado para {mat_n}.")
    for e in errores:
        st.warning(e)
    ok = U.clausula_responsabilidad(usuario, f"{kk}_cl")
    if st.button("✅ Confirmar y Guardar libro de vuelo", type="primary", disabled=not (ok and ack) or bool(errores),
                 key=f"{kk}_save", width="stretch"):
        sumas = {v["CAMPO"]: v["_suma"] for v in validacion}
        cabecera = {"numero_libro": numero, "fecha": fecha, "matricula": mat_n, "piloto_nombre": piloto, "piloto_licencia": lic_p,
                    "piloto_cedula": R.solo_digitos(ced_p) or None, "copiloto_nombre": copiloto, "copiloto_licencia": lic_c,
                    "copiloto_cedula": R.solo_digitos(ced_c) or None, "horas_totales": horas_tot, "ciclos_totales": ciclos_tot,
                    "proxima_inspeccion": prox,
                    "tot_tiempo_vuelo_min": totales["tiempo_vuelo_min"] if totales["tiempo_vuelo_min"] is not None else sumas.get("tiempo_vuelo_min"),
                    "tot_tiempo_nocturno_min": totales["tiempo_nocturno_min"] if totales["tiempo_nocturno_min"] is not None else sumas.get("tiempo_nocturno_min"),
                    "tot_tiempo_instrumentos_min": totales["tiempo_instrumentos_min"] if totales["tiempo_instrumentos_min"] is not None else sumas.get("tiempo_instrumentos_min"),
                    "tot_tiempo_total_min": totales["tiempo_total_min"] if totales["tiempo_total_min"] is not None else sumas.get("tiempo_total_min"),
                    "tot_lbs_usado": totales["lbs_comb_usado"] if totales["lbs_comb_usado"] is not None else sumas.get("lbs_comb_usado"),
                    "observaciones": " · ".join(x for x in [R.texto(obs), (f"DESCUADRE: {obs_desc}" if descuadre else "")] if x) or None,
                    "descuadre_totales": descuadre, "parametros_motor": d.get("parametros_motor")}
        mant = [{c: R.texto(r.get(et)) or None for et, c in MANT_COLS} for _, r in ed_mant.iterrows()]
        mant = [m for m in mant if any(m.values())]
        corr = _correcciones((estado.get("ia") or {}).get("cabecera"), {k2: cabecera[k2] for k2 in ("numero_libro", "fecha", "matricula", "piloto_nombre", "copiloto_nombre")})
        ia_tramos = (estado.get("ia") or {}).get("tramos") or []
        for i, (a, b) in enumerate(zip(ia_tramos, tramos)):
            corr += _correcciones(a, b, f"tramos[{i}].")
        hora0 = next((t.get("hora_salida_min") for t in tramos if t.get("hora_salida_min") is not None), 0)
        _, charter = R.normalizar_nro_vuelo(tramos[0].get("numero_vuelo"), mat_n)
        try:
            rid, ev = _guardar_documento("LV", estado, fecha, hora0, mat_n, charter,
                                         lambda doc: db.registrar_libro(cabecera, tramos, mant, doc, usuario, origen_carga, corr))
        except ValueError as exc:
            st.error(str(exc))
            return
        sync = db.sincronizar_plano(fecha, fecha, usuario)
        estado["guardado"] = {"titulo": "Libro de vuelo registrado con trazabilidad",
                              "texto": f"{mat_n} · Libro Nº {numero or 's/n'} · {len(tramos)} tramo(s) · archivo {ev['nombre']} · registro #{rid}",
                              "avisos": [f"Archivo plano actualizado: {sync['creadas']} vuelo(s) nuevos, {sync['actualizadas']} actualizado(s)."]}
        ss[sk] = estado
        st.rerun()


def _form_ficha(usuario, estado, sk, k, origen_carga):
    d, dud, kk = estado["datos"], estado.get("dudosos") or [], f"{k}_{estado['ronda']}"
    U.titulo("Vuelos de la ficha / cierre", "Pasajeros totales, tripadis, bebés (Nibra), tránsitos, carga y combustible básico del manifiesto P&B.")
    if f"{kk}_dff" not in estado:
        estado[f"{kk}_dff"] = _df_desde_lista(d.get("vuelos") or [{}], FICHA_COLS, "vuelos", dud)
        ss[sk] = estado
    ed = _editor(estado[f"{kk}_dff"], key=f"{kk}_fichas", num_rows="dynamic", hide_index=True, width="stretch",
                        column_config=_cfg_editor(FICHA_COLS))
    vuelos = _lista_desde_df(ed.drop(columns=["REVISAR"], errors="ignore"), FICHA_COLS)
    if vuelos:
        prev = pd.DataFrame([{"VUELO": v.get("numero_vuelo"), "RUTA": f"{v.get('origen') or ''}-{v.get('destino') or ''}",
                              "PAX A BORDO (cálculo)": R.pax_a_bordo(v.get("pax_total"), v.get("tripadis"), v.get("bebes"), v.get("pax_transito"))[0]}
                             for v in vuelos])
        st.caption("Pasajeros a bordo que irán al archivo plano = Pax total − Tripadis − Bebés − Tránsito")
        _tabla(prev, hide_index=True, width="stretch")
    errores = []
    for i, v in enumerate(vuelos):
        if not v.get("fecha"):
            errores.append(f"Fila {i + 1}: fecha no válida (DD/MM/AAAA).")
        if not R.norm_matricula(v.get("matricula")).startswith("HK-"):
            errores.append(f"Fila {i + 1}: matrícula no válida.")
        if v.get("fecha") and v["fecha"] > db.hoy():
            errores.append(f"Fila {i + 1}: la fecha no puede ser futura.")
    if not vuelos:
        errores.append("Registre al menos un vuelo.")
    for e in errores:
        st.warning(e)
    ok = U.clausula_responsabilidad(usuario, f"{kk}_cl")
    if st.button("✅ Confirmar y Guardar ficha", type="primary", disabled=not ok or bool(errores), key=f"{kk}_save", width="stretch"):
        for v in vuelos:
            v["matricula"] = R.norm_matricula(v.get("matricula"))
            v["observaciones"] = estado.get("nota_ia")
        corr = []
        for i, (a, b) in enumerate(zip((estado.get("ia") or {}).get("vuelos") or [], vuelos)):
            corr += _correcciones(a, b, f"vuelos[{i}].")
        v0 = vuelos[0]
        _, charter = R.normalizar_nro_vuelo(v0.get("numero_vuelo"), v0["matricula"])
        try:
            ids, ev = _guardar_documento("FV", estado, v0["fecha"], v0.get("hora_min") or 0, v0["matricula"], charter,
                                         lambda doc: db.registrar_fichas(vuelos, doc, usuario, origen_carga, corr))
        except ValueError as exc:
            st.error(str(exc))
            return
        fechas = sorted({v["fecha"] for v in vuelos})
        sync = db.sincronizar_plano(fechas[0], fechas[-1], usuario)
        estado["guardado"] = {"titulo": "Ficha de vuelo registrada con trazabilidad",
                              "texto": f"{len(ids)} vuelo(s) · archivo {ev['nombre']}",
                              "avisos": [f"Archivo plano actualizado: {sync['actualizadas']} vuelo(s) depurados con la ficha."]}
        ss[sk] = estado
        st.rerun()


# ============================================================
# Consolidados en cascada (componente compartido)
# ============================================================
def render_consolidados(area: str, usuario: dict, key: str):
    es_ger = usuario["rol"] == "GERENCIA"
    nivel = st.segmented_control("Nivel", ["📅 Diario", "📆 Histórico mensual", "🗓️ Histórico anual"], default="📅 Diario",
                                 key=f"{key}_nivel", required=True, label_visibility="collapsed")
    if nivel == "📅 Diario":
        c1, c2 = st.columns([1, 2])
        fecha = c1.date_input("Día", db.hoy() - timedelta(days=1), max_value=db.hoy(), key=f"{key}_dia", format="DD/MM/YYYY")
        estado = db.estado_dia(area, fecha)
        with c2:
            U.chips([(estado.replace("_", " "), "ok" if estado != "ABIERTO" else "warn"), (area, "navy")])
        hojas, meta = db.obtener_consolidado("DIARIO", area, fecha.isoformat())
        if estado == "ABIERTO" or hojas is None:
            hojas, meta = db.snapshot_area(area, fecha.isoformat(), fecha.isoformat()), None
            st.caption("Vista en vivo: el día sigue ABIERTO y puede cambiar.")
        else:
            st.caption(f"Instantánea congelada el {meta.get('fecha_cierre')} por {meta.get('cerrado_por_nombre')} · "
                       f"integridad {'verificada ✅' if meta.get('integro') else '⚠️ NO coincide'}")
        U.kpi_fila([{"label": n, "value": f"{len(df):,}"} for n, df in hojas.items()])
        st.download_button("⬇️ Descargar consolidado diario", X.excel_consolidado("DIARIO", area, fecha.isoformat(), hojas, meta, usuario),
                           X.nombre_consolidado("DIARIO", area, fecha.isoformat()), XLSX, key=f"{key}_dl_d", width="stretch")
        if estado == "ABIERTO":
            pulsado, _ = guardar_con_clausula(usuario, f"{key}_cd_{fecha}", f"🔒 Cerrar el día {fecha:%d/%m/%Y}", motivo=False)
            if pulsado:
                ok, msg = db.cerrar_dia(area, fecha, usuario)
                flash(msg if ok else "", [] if ok else [msg])
                st.rerun()
        elif estado == "CERRADO_DIARIO" and es_ger:
            with st.expander("🔓 Reabrir el día (solo Gerencia)"):
                pulsado, motivo = guardar_con_clausula(usuario, f"{key}_rd_{fecha}", "Reabrir día", tipo="secondary")
                if pulsado:
                    ok, msg = db.reabrir_dia(area, fecha, usuario, motivo)
                    flash(msg if ok else "", [] if ok else [msg])
                    st.rerun()
    elif nivel == "📆 Histórico mensual":
        periodo = selector_mes(f"{key}_mes")
        cerrado = db.periodo_mensual_cerrado(area, periodo)
        dias = db.dias_con_datos(area, periodo)
        estados = pd.DataFrame([{"DÍA": d, "ESTADO": db.estado_dia(area, d)} for d in dias])
        c1, c2, c3 = st.columns(3)
        with c1:
            U.kpi("Estado del mes", "CERRADO" if cerrado else "ABIERTO", tono="bien" if cerrado else "alerta")
        with c2:
            U.kpi("Días con datos", f"{len(dias)}")
        with c3:
            U.kpi("Días abiertos", f"{int((estados['ESTADO'] == 'ABIERTO').sum()) if not estados.empty else 0}")
        if not estados.empty:
            with st.expander("Estado por día"):
                _tabla(estados, hide_index=True, width="stretch")
        hojas, meta = db.obtener_consolidado("MENSUAL", area, periodo)
        if not cerrado or hojas is None:
            ini, fin = limites_mes(periodo)
            hojas, meta = db.snapshot_area(area, ini.isoformat(), fin.isoformat(), incluir_liquidacion=(area == "CONTABILIDAD")), None
            st.caption("Vista en vivo del mes (aún no congelada).")
        st.download_button("⬇️ Descargar histórico mensual", X.excel_consolidado("MENSUAL", area, periodo, hojas, meta, usuario),
                           X.nombre_consolidado("MENSUAL", area, periodo), XLSX, key=f"{key}_dl_m", width="stretch")
        if not cerrado:
            abiertos = int((estados["ESTADO"] == "ABIERTO").sum()) if not estados.empty else 0
            if abiertos:
                st.info(f"Al cerrar el mes se cerrarán también {abiertos} día(s) que siguen abiertos.")
            pulsado, _ = guardar_con_clausula(usuario, f"{key}_cm_{periodo}", f"🔒 Congelar cierre mensual {periodo}", motivo=False)
            if pulsado:
                with st.spinner("Congelando el mes…"):
                    ok, msg = db.cerrar_mes(area, periodo, usuario)
                flash(msg if ok else "", [] if ok else [msg])
                st.rerun()
        elif es_ger:
            with st.expander("🔓 Reabrir el mes (solo Gerencia)"):
                pulsado, motivo = guardar_con_clausula(usuario, f"{key}_rm_{periodo}", "Reabrir mes", tipo="secondary")
                if pulsado:
                    ok, msg = db.reabrir_mes(area, periodo, usuario, motivo)
                    flash(msg if ok else "", [] if ok else [msg])
                    st.rerun()
    else:
        anio = st.selectbox("Año", [str(db.hoy().year - i) for i in range(6)], key=f"{key}_anio")
        hojas, meta = db.obtener_consolidado("ANUAL", area, anio)
        if hojas is None:
            U.vacio("🗓️", "Sin meses congelados en este año", "El histórico anual se alimenta automáticamente con cada cierre mensual.")
            return
        meses = json.loads(meta.get("meses_incluidos") or "[]")
        U.chips([(meta.get("estado", ""), "ok" if meta.get("estado") == "CERRADO_ANUAL" else "info")] +
                [(m, "navy") for m in meses])
        U.kpi_fila([{"label": n, "value": f"{len(df):,}"} for n, df in hojas.items()])
        st.download_button("⬇️ Descargar histórico anual", X.excel_consolidado("ANUAL", area, anio, hojas, meta, usuario),
                           X.nombre_consolidado("ANUAL", area, anio), XLSX, key=f"{key}_dl_a", width="stretch")
        if es_ger and meta.get("estado") != "CERRADO_ANUAL":
            pulsado, _ = guardar_con_clausula(usuario, f"{key}_ca_{anio}", f"🔒 Cierre anual {anio}", motivo=False)
            if pulsado:
                ok, msg = db.cerrar_anio(area, anio, usuario)
                flash(msg if ok else "", [] if ok else [msg])
                st.rerun()
    with st.expander("Historial de cierres"):
        for niv in ("DIARIO", "MENSUAL", "ANUAL"):
            df = db.df_consolidados(niv, area)
            if not df.empty:
                st.markdown(f"**{niv.title()}**")
                _tabla(df.drop(columns=["resumen"], errors="ignore"), hide_index=True, width="stretch")


def render_auditoria_area(area: str, usuario: dict, key: str):
    etiqueta = {"CONTABILIDAD": "🔍 Auditar Área Contable con IA", "OPERACIONES": "🔍 Auditar Área de Operaciones con IA"}[area]
    st.caption("Revisa únicamente los datos y soportes de su área. No incluye evaluaciones de desempeño de empleados.")
    desde, hasta = rango_fechas(f"{key}_rg")
    if st.button(etiqueta, type="primary", key=f"{key}_run"):
        with st.spinner("Auditando…"):
            ss[f"{key}_res"] = A.ejecutar_auditoria(area, desde, hasta, usuario)
            ss[f"{key}_res"]["rango"] = (desde, hasta)
    res = ss.get(f"{key}_res")
    if not res:
        U.vacio("🛡️", "Ejecute la auditoría del periodo", "Se revisan diferencias, duplicados, celdas pendientes y documentos faltantes.")
        return
    h = res["hallazgos"]
    cuenta = h["SEVERIDAD"].value_counts() if not h.empty else {}
    U.kpi_fila([{"label": "Hallazgos", "value": f"{len(h)}"},
                {"label": "Altos", "value": f"{cuenta.get('Alta', 0)}", "tono": "mal" if cuenta.get("Alta", 0) else "bien"},
                {"label": "Medios", "value": f"{cuenta.get('Media', 0)}", "tono": "alerta" if cuenta.get("Media", 0) else ""},
                {"label": "Bajos", "value": f"{cuenta.get('Baja', 0)}"}])
    with st.container(border=True):
        st.markdown("**Informe" + (" (IA)" if res["uso_ia"] else " automático") + "**")
        st.text(res["resumen"])
    if h.empty:
        st.success("Sin hallazgos en el periodo.")
        return
    vista = h.drop(columns=["RESPONSABLE_CEDULA", "RESPONSABLE_NOMBRE", "ULTIMO_EDITOR_CEDULA", "ULTIMO_EDITOR_NOMBRE"], errors="ignore")
    sev = st.segmented_control("Severidad", ["Todos", "Alta", "Media", "Baja"], default="Todos", key=f"{key}_sev", required=True)
    if sev != "Todos":
        vista = vista[vista["SEVERIDAD"] == sev]
    _tabla(vista, hide_index=True, width="stretch", height=380)
    d, hh = res["rango"]
    st.download_button("⬇️ Descargar informe de auditoría", X.excel_auditoria(area, d, hh, vista, pd.DataFrame(), res["resumen"], usuario),
                       f"Moon_Auditoria_{area}_{d:%Y%m%d}_{hh:%Y%m%d}.xlsx", XLSX, key=f"{key}_dl")


# ============================================================
# MÓDULO 1 · Pilotos
# ============================================================
def modulo_pilotos(u):
    requiere_rol("PILOTO", "GERENCIA")
    U.page_title("Pilotos · Captura en pista")
    U.hero("Cada soporte, firmado y bajo control.",
           "Sube la foto del tiquete, del libro de vuelo o de la ficha. La IA lee los datos y tú confirmas que coincidan con el papel.",
           "Validación cruzada IA + humano")
    render_captura_documental(u, ["COMB", "LV", "FV"], "PILOTO", "pil")
    with st.expander("📂 Mis envíos de los últimos 45 días"):
        df = db.df_documentos(db.hoy() - timedelta(days=45), db.hoy(), cedula=u["cedula"])
        if df.empty:
            st.caption("Aún no ha enviado documentos.")
        else:
            _tabla(pd.DataFrame({"TIPO": df["tipo"].map(lambda t: V.TIPOS_DOC[t]["nombre"]), "FECHA": df["fecha_doc"],
                                       "MATRÍCULA": df["matricula"], "ARCHIVO": df["ruta_archivo"].map(lambda s: Path(s).name),
                                       "CARGADO": df["fecha_creacion"]}), hide_index=True, width="stretch")


# ============================================================
# MÓDULO 2 · Contabilidad
# ============================================================
def modulo_contabilidad(u):
    requiere_rol("CONTABILIDAD", "GERENCIA")
    U.page_title("Contabilidad")
    U.hero("Combustible, horas y honorarios con trazabilidad total.",
           "Concilia tiquetes contra libros de vuelo, liquida TIEMPO PILOTOS y consolida en cascada diario → mensual → anual.",
           "Fuentes: tiquetes + libros de vuelo")
    sec = st.segmented_control("Sección", ["⛽ Conciliación de combustible", "⏱️ Liquidación horas y pesos", "🗂️ Soportes y tabla auditada",
                                           "🔍 Auditoría IA contable", "📦 Consolidados", "📤 Carga de respaldo"],
                               default="⛽ Conciliación de combustible", key="cont_sec", required=True, label_visibility="collapsed")
    if sec.startswith("⛽"):
        tab_conciliacion(u)
    elif sec.startswith("⏱️"):
        tab_liquidacion(u)
    elif sec.startswith("🗂️"):
        tab_soportes_contables(u)
    elif sec.startswith("🔍"):
        render_auditoria_area("CONTABILIDAD", u, "audc")
    elif sec.startswith("📦"):
        render_consolidados("CONTABILIDAD", u, "conc")
    else:
        U.titulo("📤 Carga de Respaldo (Tiquete / Libro de Vuelo)", "Para soportes que el piloto no cargó. Queda trazado que lo cargó Contabilidad.")
        render_captura_documental(u, ["COMB", "LV"], "CONTABILIDAD", "contres")


def _vista_conciliacion(desde, hasta):
    tiq = db.df_tiquetes(desde, hasta)
    tramos = db.df_tramos(desde, hasta)
    p = db.params_dict()
    conc, sin_tiq = R.conciliar_combustible(db._lista(tiq), db._lista(tramos), p.get("tolerancia_galones_pct") or 2,
                                            p.get("tolerancia_galones_abs") or 2)
    cmap = {c["tiquete_id"]: c for c in conc}
    v = db.vista_tiquetes(tiq)
    if not v.empty:
        v.insert(9, "GAL LIBRO", v["ID"].map(lambda i: cmap.get(i, {}).get("gal_libro")).astype(float))
        v.insert(10, "DIF. GAL", v["ID"].map(lambda i: cmap.get(i, {}).get("dif_gal")).astype(float))
        v.insert(11, "CONCILIACIÓN", v["ID"].map(lambda i: cmap.get(i, {}).get("estado")))
        v.insert(12, "LIBRO Nº", v["ID"].map(lambda i: cmap.get(i, {}).get("numero_libro")))
    return tiq, tramos, v, cmap, sin_tiq


def tab_conciliacion(u):
    desde, hasta = rango_fechas("conc_rg")
    tiq, tramos, v, cmap, sin_tiq = _vista_conciliacion(desde, hasta)
    if v.empty:
        U.vacio("⛽", "Sin tiquetes en el periodo", "Cuando los pilotos carguen tiquetes aparecerán aquí para conciliar y valorizar.")
        return
    gal = v["GALONES"].sum()
    valor = v["VALOR TOTAL"].sum(min_count=1)
    sin_precio = int(v["VALOR UNIT"].isna().sum())
    concil = int((v["CONCILIACIÓN"] == "CONCILIADO").sum())
    U.kpi_fila([{"label": "Tiquetes", "value": f"{len(v):,}"}, {"label": "Galones", "value": U.num(gal, 1)},
                {"label": "Valorizado", "value": U.cop(valor)},
                {"label": "Sin precio", "value": f"{sin_precio}", "tono": "alerta" if sin_precio else "bien",
                 "nota": "Dato manual pendiente" if sin_precio else "Todo valorizado"},
                {"label": "Conciliados con libro", "value": f"{concil}/{len(v)}", "tono": "bien" if concil == len(v) else "alerta"}])
    g1, g2 = st.columns(2)
    with g1:
        U.mostrar(U.barras_h(v, "MATRICULA", "GALONES", "Galones por aeronave"))
    with g2:
        U.mostrar(U.dona(v.assign(N=1), "CONCILIACIÓN", "N", "Estado de conciliación", f"{len(v)}<br>tiquetes"))

    with st.container(border=True):
        U.titulo("💲 Valorizar tiquetes sin precio", "Aplica la tarifa vigente o un precio de factura. Si la factura no ha llegado, deje en blanco.")
        p = db.params_dict()
        productos = sorted(v.loc[v["VALOR UNIT"].isna(), "PRODUCTO"].dropna().unique().tolist())
        if not productos:
            st.caption("No hay tiquetes pendientes de valorizar en el periodo.")
        else:
            precios = {}
            cols = st.columns(len(productos))
            for c, prod in zip(cols, productos):
                vigente = p.get(f"precio_{prod}")
                precios[prod] = c.text_input(f"{prod} · COP/gal", "" if vigente is None else str(int(vigente)), key=f"conc_pr_{prod}",
                                             help="Tarifa vigente definida por Gerencia. En blanco = no se aplica (pendiente).")
                if vigente is None:
                    c.caption("✍️ Tarifa vigente sin parametrizar")
            validos = {k: R.a_float(x) for k, x in precios.items() if R.a_float(x)}
            pulsado, motivo = guardar_con_clausula(u, "conc_aplicar", "Aplicar precios a tiquetes sin valorizar", extra_ok=bool(validos))
            if pulsado:
                n, errores = 0, []
                for _, r in v[v["VALOR UNIT"].isna()].iterrows():
                    if r["PRODUCTO"] in validos:
                        ok, msg, k2 = db.actualizar_registro("combustible_tiquetes", int(r["ID"]), {"valor_unit": validos[r["PRODUCTO"]]},
                                                             u, motivo, "CONTABILIDAD")
                        n += 1 if ok and k2 else 0
                        if not ok:
                            errores.append(msg)
                flash(f"{n} tiquete(s) valorizados.", errores)
                st.rerun()

    U.titulo("Tabla de conciliación (editable y auditada)", "Precio en blanco = pendiente. VALOR TOTAL se recalcula: galones × valor unitario.")
    vista = a_objeto(v.copy(), ["HORA", "ESTACION", "MATRICULA", "RECIBO", "BOLETO", "PRODUCTO", "PILOTO", "OBSERVACION", "CONCILIACIÓN", "LIBRO Nº"])
    cfg = {"GALONES": st.column_config.NumberColumn(format="%.1f"), "GAL LIBRO": st.column_config.NumberColumn(format="%.1f"),
           "DIF. GAL": st.column_config.NumberColumn(format="%+.1f"), "VALOR UNIT": st.column_config.NumberColumn("✍️ VALOR UNIT", format="$ %d"),
           "VALOR TOTAL": st.column_config.NumberColumn(format="$ %d"),
           "PRODUCTO": st.column_config.SelectboxColumn(options=R.PRODUCTOS)}
    editor_auditado(vista, key="conc_ed", tabla="combustible_tiquetes", modulo="CONTABILIDAD", usuario=u, column_config=cfg,
                    mapa={"RECIBO": "recibo", "BOLETO": "numero_boleto", "FECHA": "fecha_consumo", "ESTACION": "estacion",
                          "MATRICULA": "matricula", "PRODUCTO": "producto", "GALONES": "galones", "VALOR UNIT": "valor_unit",
                          "PILOTO": "piloto_nombre", "OBSERVACION": "observacion"})
    if sin_tiq:
        st.warning(f"{len(sin_tiq)} tramo(s) del libro de vuelo registran tanqueo sin tiquete cargado (documento faltante).")
        _tabla(pd.DataFrame([{"FECHA": t["fecha"], "MATRÍCULA": t["matricula"], "LIBRO Nº": t.get("numero_libro"),
                                    "VUELO": t.get("numero_vuelo"), "GAL TANQ.": t.get("galones_tanqueados"),
                                    "RECIBO EN LIBRO": t.get("recibo_combustible"), "PILOTO": t.get("piloto_nombre")} for t in sin_tiq]),
                     hide_index=True, width="stretch")
    with st.expander("🔎 Confrontar tiquete y libro de vuelo lado a lado"):
        ids = v["ID"].tolist()
        etiquetas = {int(r.ID): f"#{r.ID} · {r.FECHA} · {r.MATRICULA} · {r.RECIBO} · {r.GALONES:,.1f} gal" for r in v.itertuples()}
        sel = st.selectbox("Tiquete", ids, format_func=lambda i: etiquetas.get(i, str(i)), key="conc_vis")
        fila = tiq[tiq["id"] == sel].iloc[0]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Tiquete de combustible**")
            visor(fila["ruta_archivo"])
        with c2:
            st.markdown("**Libro de vuelo conciliado**")
            c = cmap.get(sel, {})
            lib = tramos[tramos["libro_id"] == c.get("libro_id")] if c.get("libro_id") else pd.DataFrame()
            visor(lib.iloc[0]["ruta_archivo"] if not lib.empty else None)
    bloque_eliminar("combustible_tiquetes", v["ID"].tolist(),
                    {int(r.ID): f"#{r.ID} · {r.FECHA} · {r.MATRICULA} · {r.RECIBO}" for r in v.itertuples()}, u, "CONTABILIDAD", "conc_del")


def editor_tiempo_pilotos(u, key, modulo):
    p = db.params_dict()
    cfg = db.cfg_tiempo_pilotos(p)
    with st.expander(f"⚙️ Regla TIEMPO PILOTOS · modo actual: **{cfg['modo']}**"):
        st.caption("MANUAL = la columna queda en blanco para digitar fila por fila. FORMULA = expresión en minutos. "
                   "RANGOS = tabla por rangos de minutos. Un valor digitado a mano siempre prevalece.")
        modos = ["MANUAL", "FORMULA", "RANGOS"]
        modo = st.radio("Modo", modos, index=modos.index(cfg["modo"]) if cfg["modo"] in modos else 0, horizontal=True, key=f"{key}_modo")
        formula = st.text_input("Fórmula (resultado en minutos)", cfg.get("formula") or "", key=f"{key}_formula",
                                help=f"Variables: {', '.join(R.VARS_TIEMPO_PILOTOS)}. Funciones: si(cond, a, b), min, max, redondear, techo, piso.")
        err = R.validar_formula(formula, R.VARS_TIEMPO_PILOTOS) if R.texto(formula) else None
        if err:
            st.error(err)
        elif R.texto(formula):
            ej, _ = R.evaluar_formula(formula, R.variables_tramo({"tiempo_vuelo_min": 26, "tiempo_total_min": 34}, 0))
            st.caption(f"Ejemplo con vuelo 0:26 y bloque 0:34 → {R.fmt_duracion(ej) if ej is not None else 'sin resultado'}")
        base = st.selectbox("Comparar rangos contra", ["tiempo_vuelo", "tiempo_total"],
                            index=0 if cfg["base"] != "tiempo_total" else 1, key=f"{key}_base")
        rangos = pd.DataFrame(cfg.get("rangos") or [], columns=["desde", "hasta", "resultado"])
        rangos = a_numero(rangos, ["desde", "hasta"])
        rangos["resultado"] = rangos["resultado"].astype(object)
        ed = _editor(rangos, num_rows="dynamic", key=f"{key}_rangos", hide_index=True, width="stretch",
                            column_config={"desde": st.column_config.NumberColumn("Desde (min)"),
                                           "hasta": st.column_config.NumberColumn("Hasta (min)"),
                                           "resultado": st.column_config.TextColumn("Resultado (min o fórmula)")})
        lista, errores = [], []
        for _, r in ed.iterrows():
            if R.es_vacio(r["resultado"]) and R.es_vacio(r["desde"]) and R.es_vacio(r["hasta"]):
                continue
            e2 = R.validar_formula(str(r["resultado"]), R.VARS_TIEMPO_PILOTOS) if not R.es_vacio(r["resultado"]) else "Resultado vacío"
            if e2:
                errores.append(f"Rango {r['desde']}-{r['hasta']}: {e2}")
            lista.append({"desde": R.a_float(r["desde"]), "hasta": R.a_float(r["hasta"]), "resultado": R.texto(r["resultado"])})
        for e2 in errores:
            st.error(e2)
        ajuste = st.number_input("ajuste_min (variable disponible en fórmulas)", value=float(cfg.get("ajuste_min") or 0), step=1.0, key=f"{key}_aj")
        pulsado, motivo = guardar_con_clausula(u, f"{key}_save", "Guardar regla TIEMPO PILOTOS", extra_ok=not err and not errores)
        if pulsado:
            for clave, valor in (("tp_modo", modo), ("tp_formula", formula), ("tp_rangos", json.dumps(lista)),
                                 ("tp_rangos_base", base), ("tp_ajuste_min", ajuste)):
                db.set_param(clave, valor, u, motivo, modulo)
            flash("Regla TIEMPO PILOTOS actualizada.")
            st.rerun()


def tab_liquidacion(u):
    c1, c2 = st.columns([1, 3])
    with c1:
        periodo = selector_mes("liq_mes", "Periodo de liquidación")
    ini, fin = limites_mes(periodo)
    if db.periodo_mensual_cerrado("CONTABILIDAD", periodo):
        st.info(f"El periodo {periodo} está congelado: los cambios requieren que Gerencia lo reabra.")
    editor_tiempo_pilotos(u, "liq_tp", "CONTABILIDAD")
    liq, det = db.datos_liquidacion(ini, fin, periodo)
    if det.empty:
        U.vacio("⏱️", "Sin tramos de libro de vuelo en el periodo", "La liquidación se alimenta de los libros de vuelo confirmados.")
        return
    pend = int(det["_tp"].isna().sum())
    total = pd.to_numeric(liq["TOTAL A PAGAR (COP)"], errors="coerce").sum(min_count=1) if not liq.empty else None
    U.kpi_fila([{"label": "Tramos del periodo", "value": f"{len(det):,}"},
                {"label": "Tramos sin TIEMPO PILOTOS", "value": f"{pend}", "tono": "alerta" if pend else "bien",
                 "nota": "Dato manual / fórmula pendiente" if pend else "Completo"},
                {"label": "Tripulantes", "value": f"{len(liq)}"},
                {"label": "Total a pagar", "value": U.cop(total)}])

    U.titulo("Detalle de tramos · TIEMPO PILOTOS", "Edite la columna en H:MM. Vacíe una celda manual para volver a la regla configurada.")
    vista = det.drop(columns=[c for c in det.columns if c.startswith("_")] + ["LIBRO_ID"]).rename(columns={"TRAMO_ID": "ID"})
    vista = a_objeto(vista, list(vista.columns.drop("ID")))
    editor_auditado(vista, key=f"liq_det_{periodo}", tabla="tramos_vuelo", modulo="CONTABILIDAD", usuario=u,
                    mapa={"TIEMPO PILOTOS": "tiempo_pilotos_min"},
                    column_config={"TIEMPO PILOTOS": st.column_config.TextColumn("✍️ TIEMPO PILOTOS", help="H:MM")})

    U.titulo("Liquidación por tercero (cédula)", "PTL = piloto al mando · PCA = copiloto. Las celdas ✍️ admiten ajuste manual con motivo.")
    if liq.empty:
        st.caption("No hay tripulación identificada en los tramos.")
        return
    version = ss.get("liq_tab_v", 0)
    k = f"liq_tab_{periodo}_{version}"
    vliq = liq.drop(columns=["CLAVE"]).copy()
    vliq = a_numero(vliq, ["MIN SMLV", "VALOR HORA", "VALOR MINUTO", "TOTAL A PAGAR (COP)"])
    _editor(vliq, key=k, hide_index=True, width="stretch", disabled=[c for c in vliq.columns if c not in R.LIQ_EDITABLES],
                   column_config={"CATEGORIA": st.column_config.SelectboxColumn("✍️ CATEGORIA", options=["PTL", "PCA"]),
                                  "MIN PILOTO": st.column_config.NumberColumn("✍️ MIN PILOTO"),
                                  "MIN COPILOTO": st.column_config.NumberColumn("✍️ MIN COPILOTO"),
                                  "MIN SMLV": st.column_config.NumberColumn("✍️ MIN SMLV"),
                                  "VALOR HORA": st.column_config.NumberColumn("✍️ VALOR HORA", format="$ %d"),
                                  "VALOR MINUTO": st.column_config.NumberColumn(format="$ %.2f"),
                                  "TOTAL A PAGAR (COP)": st.column_config.NumberColumn("✍️ TOTAL A PAGAR (COP)", format="$ %d")})
    cambios = (ss.get(k) or {}).get("edited_rows") or {}
    if cambios:
        with st.container(border=True):
            pulsado, motivo = guardar_con_clausula(u, f"{k}_g", "💾 Guardar ajustes de liquidación")
            if pulsado:
                errores = []
                for idx, cols in cambios.items():
                    clave = liq.iloc[int(idx)]["CLAVE"]
                    for col, val in cols.items():
                        ok, msg = db.guardar_ajuste_liquidacion(periodo, clave, R.LIQ_EDITABLES[col], val, u, motivo)
                        if not ok:
                            errores.append(msg)
                ss["liq_tab_v"] = version + 1
                flash("Ajustes registrados en la bitácora.", errores)
                st.rerun()
    smlv = db.get_param("min_smlv")
    fig = go.Figure()
    fig.add_bar(x=liq["NOMBRE"], y=liq["MIN PILOTO"], name="Min. piloto (PTL)", marker_color=U.NAVY)
    fig.add_bar(x=liq["NOMBRE"], y=liq["MIN COPILOTO"], name="Min. copiloto (PCA)", marker_color=U.SKY)
    if smlv:
        fig.add_hline(y=float(smlv), line_dash="dash", line_color=U.AMBER, annotation_text=f"Mínimo SMLV {U.num(smlv)} min")
    fig.update_layout(barmode="stack", title="Minutos por tripulante vs. mínimo garantizado")
    U.mostrar(U.estilo(fig, 380))
    aud = db.df_auditoria(ini, fin)
    aud = aud[aud["tabla_afectada"].isin(["tramos_vuelo", "liquidacion_ajustes", "combustible_tiquetes", "libros_vuelo"]) &
              aud["accion"].isin(["EDITAR", "CORRECCION_IA", "ELIMINAR"])] if not aud.empty else aud
    st.download_button("⬇️ Descargar Excel de liquidación", X.excel_liquidacion(liq, det, db.vista_tiquetes(db.df_tiquetes(ini, fin)), aud, periodo, u),
                       f"Moon_Liquidacion_Pilotos_{periodo.replace('-', '_')}.xlsx", XLSX, key="liq_dl", type="primary")


def tab_soportes_contables(u):
    desde, hasta = rango_fechas("sop_rg")
    tipo = st.segmented_control("Ver", ["📘 Libros de vuelo", "⛽ Tiquetes"], default="📘 Libros de vuelo", key="sop_tipo", required=True)
    if tipo.startswith("📘"):
        libros = db.df_libros(desde, hasta)
        if libros.empty:
            U.vacio("📘", "Sin libros de vuelo en el periodo", "Use la carga de respaldo si el piloto no los envió.")
            return
        etiquetas = {int(r.id): f"{r.fecha} · {r.matricula} · Libro Nº {r.numero_libro or 's/n'} · {r.piloto_nombre or ''}"
                     for r in libros.itertuples()}
        sel = st.selectbox("Libro de vuelo", list(etiquetas), format_func=etiquetas.get, key="sop_libro")
        fila = libros[libros["id"] == sel].iloc[0]
        c1, c2 = st.columns([1, 1.4], gap="large")
        with c1:
            visor(fila["ruta_archivo"])
        with c2:
            U.chips([("Descuadre de totales", "bad")] if fila["descuadre_totales"] else [("Totales cuadrados", "ok")])
            cab = pd.DataFrame([{"ID": int(fila["id"]), "LIBRO Nº": fila["numero_libro"], "FECHA": fila["fecha"], "MATRICULA": fila["matricula"],
                                 "PILOTO": fila["piloto_nombre"], "LICENCIA PILOTO": fila["piloto_licencia"], "CÉDULA PILOTO": fila["piloto_cedula"],
                                 "COPILOTO": fila["copiloto_nombre"], "LICENCIA COPILOTO": fila["copiloto_licencia"],
                                 "CÉDULA COPILOTO": fila["copiloto_cedula"], "OBSERVACIONES": fila["observaciones"]}])
            cab = a_objeto(cab, list(cab.columns.drop("ID")))
            editor_auditado(cab, key=f"sop_cab_{sel}", tabla="libros_vuelo", modulo="CONTABILIDAD", usuario=u,
                            mapa={"LIBRO Nº": "numero_libro", "FECHA": "fecha", "MATRICULA": "matricula", "PILOTO": "piloto_nombre",
                                  "LICENCIA PILOTO": "piloto_licencia", "CÉDULA PILOTO": "piloto_cedula", "COPILOTO": "copiloto_nombre",
                                  "LICENCIA COPILOTO": "copiloto_licencia", "CÉDULA COPILOTO": "copiloto_cedula", "OBSERVACIONES": "observaciones"})
        tr = db.df_tramos(fila["fecha"], fila["fecha"])
        tr = tr[tr["libro_id"] == sel]
        vt = db.vista_tramos(tr).drop(columns=["FECHA", "MATRICULA", "LIBRO Nº", "PILOTO", "COPILOTO", "SOPORTE"])
        vt = a_objeto(vt, [c for c in vt.columns if c not in ("ID", "GAL TANQ.", "LBS SALIDA", "PAX TOTAL", "PAX TRÁNSITO", "LBS LLEGADA", "LBS USADO")])
        vt = a_numero(vt, ["GAL TANQ.", "LBS SALIDA", "PAX TOTAL", "PAX TRÁNSITO", "LBS LLEGADA", "LBS USADO"])
        U.titulo("Tramos del libro (editables con motivo)")
        editor_auditado(vt, key=f"sop_tr_{sel}", tabla="tramos_vuelo", modulo="CONTABILIDAD", usuario=u,
                        mapa={"VUELO": "numero_vuelo", "ORIGEN": "origen", "DESTINO": "destino", "GAL TANQ.": "galones_tanqueados",
                              "RECIBO COMB.": "recibo_combustible", "LBS SALIDA": "lbs_comb_salida", "PAX TOTAL": "pax_total",
                              "PAX TRÁNSITO": "pax_transito", "H. SALIDA": "hora_salida_min", "H. DESPEGUE": "hora_despegue_min",
                              "H. ATERRIZAJE": "hora_aterrizaje_min", "H. LLEGADA": "hora_llegada_min", "T. VUELO": "tiempo_vuelo_min",
                              "T. NOCTURNO": "tiempo_nocturno_min", "T. INSTRUM.": "tiempo_instrumentos_min", "T. TOTAL": "tiempo_total_min",
                              "LBS LLEGADA": "lbs_comb_llegada", "LBS USADO": "lbs_comb_usado", "TIEMPO PILOTOS": "tiempo_pilotos_min"})
        bloque_eliminar("libros_vuelo", [int(sel)], {int(sel): etiquetas[sel]}, u, "CONTABILIDAD", f"sop_dl_{sel}")
    else:
        tiq = db.df_tiquetes(desde, hasta)
        if tiq.empty:
            U.vacio("⛽", "Sin tiquetes en el periodo", "")
            return
        etiquetas = {int(r.id): f"{r.fecha_consumo} · {r.matricula} · {r.recibo} · {R.a_float(r.galones, 0):,.1f} gal" for r in tiq.itertuples()}
        sel = st.selectbox("Tiquete", list(etiquetas), format_func=etiquetas.get, key="sop_tiq")
        fila = tiq[tiq["id"] == sel].iloc[0]
        c1, c2 = st.columns([1, 1.4], gap="large")
        with c1:
            visor(fila["ruta_archivo"])
        with c2:
            vt = db.vista_tiquetes(tiq[tiq["id"] == sel])
            vt = a_objeto(vt, ["HORA", "ESTACION", "MATRICULA", "RECIBO", "BOLETO", "PRODUCTO", "PILOTO", "OBSERVACION"])
            _tabla(vt.T.rename(columns={0: "VALOR"}).astype(str), width="stretch")
            st.caption("Para corregir este tiquete use la tabla editable de la sección de conciliación.")
    p_ini, p_fin = desde.strftime("%Y%m%d"), hasta.strftime("%Y%m%d")
    liq, det = db.datos_liquidacion(desde, hasta, desde.strftime("%Y-%m"))
    aud = db.df_auditoria(desde, hasta)
    st.download_button("⬇️ Excel profesional (Liquidación · Detalle Vuelos · Combustible · Auditoría Cambios)",
                       X.excel_liquidacion(liq, det, db.vista_tiquetes(db.df_tiquetes(desde, hasta)), aud, f"{desde} a {hasta}", u),
                       f"Moon_Contabilidad_{p_ini}_{p_fin}.xlsx", XLSX, key="sop_dl")


# ============================================================
# MÓDULO 3 · Operaciones (archivo plano MO-A3-013)
# ============================================================
def modulo_operaciones(u):
    requiere_rol("OPERACIONES", "GERENCIA")
    U.page_title("Operaciones · Archivo plano MO-A3-013")
    U.hero("Del libro de vuelo al archivo plano, sin volver a digitar.",
           "Cruza libro de vuelo, ficha y combustible por Fecha + Matrícula + Hora + Origen + Destino. Lo que falte queda en "
           "blanco y editable: nunca se bloquea la operación.", "27 columnas A..AA · Aerocivil")
    sec = st.segmented_control("Sección", ["📄 Archivo plano", "🧾 Fichas y pasajeros", "🔍 Auditoría IA operativa", "📦 Consolidados",
                                           "📤 Carga de respaldo"], default="📄 Archivo plano", key="ops_sec", required=True,
                               label_visibility="collapsed")
    if sec.startswith("📄"):
        tab_plano(u)
    elif sec.startswith("🧾"):
        tab_fichas(u)
    elif sec.startswith("🔍"):
        render_auditoria_area("OPERACIONES", u, "audo")
    elif sec.startswith("📦"):
        render_consolidados("OPERACIONES", u, "opsc")
    else:
        U.titulo("📤 Carga de Respaldo (Tiquete / Ficha / Libro de Vuelo)", "Para soportes que el piloto no cargó. Queda trazado que lo cargó Operaciones.")
        render_captura_documental(u, ["COMB", "FV", "LV"], "OPERACIONES", "opsres")


PLANO_TIPO_EDITOR = {c: ("n" if c in R.PLANO_NUMERICAS else "t") for c in R.PLANO_CLAVES}


def _config_modelo_rapida(u, pl):
    aeronaves = db.aeronaves_dict()
    sin_modelo = sorted({R.norm_matricula(m) for m in pl["matricula"].dropna()
                         if not (aeronaves.get(R.norm_matricula(m)) or {}).get("modelo_codigo")})
    if not sin_modelo:
        return
    st.warning("Matrículas sin modelo configurado (equipo, sillas y carga quedan en blanco y editables): " + ", ".join(sin_modelo))
    with st.expander("✈️ Configurar modelo de estas aeronaves (se guarda en el catálogo de flota)"):
        modelos = db.df_modelos()
        c1, c2 = st.columns(2)
        mat = c1.selectbox("Matrícula", sin_modelo, key="ops_cfg_mat")
        opcion = c2.selectbox("Modelo", ["➕ Nuevo modelo"] + modelos["codigo"].tolist(), key="ops_cfg_mod")
        datos = {}
        if opcion == "➕ Nuevo modelo":
            st.caption("Puede dejar campos en blanco: quedarán como dato manual hasta que Gerencia los complete.")
            a, b, c = st.columns(3)
            datos = {"codigo": a.text_input("Código (ej. JS41)", key="ops_cfg_cod"), "nombre": b.text_input("Nombre", key="ops_cfg_nom"),
                     "sillas_ofrecidas": c.text_input("Sillas ofrecidas", key="ops_cfg_sil"),
                     "carga_ofrecida_kg": a.text_input("Carga ofrecida (kg)", key="ops_cfg_car"),
                     "tipo_trafico_default": b.text_input("Tipo tráfico", key="ops_cfg_tt"),
                     "uso_aeronave_default": c.text_input("Uso aeronave", key="ops_cfg_uso"),
                     "tipo_combustible_default": a.text_input("Tipo combustible", key="ops_cfg_tc"),
                     "metodo_medicion_default": b.text_input("Método medición", key="ops_cfg_mm"),
                     "reserva_combustible_lbs": c.text_input("Reserva fija (lbs)", key="ops_cfg_res"),
                     "factor_conversion_lbs_gal": a.text_input("Factor lbs/gal", key="ops_cfg_fac")}
        pulsado, motivo = guardar_con_clausula(u, "ops_cfg", "Guardar configuración de flota",
                                               extra_ok=(opcion != "➕ Nuevo modelo" or bool(R.texto(datos.get("codigo")))))
        if pulsado:
            errores = []
            codigo = opcion
            if opcion == "➕ Nuevo modelo":
                ok, msg = db.guardar_modelo(datos, u, motivo, "OPERACIONES")
                codigo = R.texto(datos["codigo"]).upper().replace(" ", "")
                if not ok:
                    errores.append(msg)
            if not errores:
                ok, msg = db.guardar_aeronave({"matricula": mat, "modelo_codigo": codigo}, u, motivo, "OPERACIONES")
                if not ok:
                    errores.append(msg)
            flash("Flota actualizada. Pulse 'Generar / actualizar' para completar el plano.", errores)
            st.rerun()


def tab_plano(u):
    desde, hasta = rango_fechas("pl_rg")
    b1, b2, b3 = st.columns([1.3, 1, 1])
    if b1.button("🔄 Generar / actualizar desde soportes", type="primary", key="pl_sync", width="stretch"):
        r = db.sincronizar_plano(desde, hasta, u)
        flash(f"Plano sincronizado: {r['creadas']} nuevas · {r['actualizadas']} actualizadas · {r['cerradas']} tramo(s) en días cerrados.")
        st.rerun()
    pl = db.plano_ordenado(desde, hasta)
    if pl.empty:
        U.vacio("📄", "Sin vuelos en el periodo", "Los vuelos aparecen al confirmar libros de vuelo, al importar la plantilla o al agregarlos manualmente.")
        _plano_extras(u, desde, hasta, pl)
        return
    f1, f2, f3 = st.columns([1.7, 2, 1])
    tipos = f1.pills("Tipo de vuelo", ["R", "A", "C"], selection_mode="multi", key="pl_tipos",
                     format_func=lambda t: f"{t} · {R.TIPOS_VUELO[t]}")
    mats = f2.multiselect("Matrícula", placeholder="Seleccione…", options=sorted(pl["matricula"].dropna().unique()), key="pl_mats")
    solo_pend = f3.toggle("Solo con pendientes", key="pl_pend")
    vis = pl.copy()
    if tipos:
        vis = vis[vis["tipo_vuelo"].isin(tipos)]
    if mats:
        vis = vis[vis["matricula"].isin(mats)]
    vis["_pend"] = vis["pendientes"].fillna("[]").map(lambda s: [p for p in json.loads(s) if p != "carga_real"])
    if solo_pend:
        vis = vis[vis["_pend"].map(bool)]
    con_pend = int(pl["pendientes"].fillna("[]").map(lambda s: bool([p for p in json.loads(s) if p != "carga_real"])).sum())
    U.kpi_fila([{"label": "Vuelos", "value": f"{len(pl):,}"},
                {"label": "Pasajeros a bordo", "value": U.num(pd.to_numeric(pl["pax_bordo"], errors="coerce").sum())},
                {"label": "Chárter", "value": f"{int((pl['tipo_vuelo'] == 'C').sum())}"},
                {"label": "Filas con celdas pendientes", "value": f"{con_pend}", "tono": "alerta" if con_pend else "bien",
                 "nota": "Dato manual pendiente" if con_pend else "Listo para Aerocivil"},
                {"label": "Sin ficha (pax sin depurar)", "value": f"{int(((pl['fuente'] == 'SOPORTES') & pl['ficha_id'].isna()).sum())}"}])
    _config_modelo_rapida(u, pl)

    U.titulo("Matriz MO-A3-013 (A..AA)", "Una sola ventana con regulares, adicionales y chárter, ordenada por fecha → matrícula → hora. "
                                         "✍️ = celdas de diligenciamiento manual. NRO se asigna al exportar.")
    v = db.vista_plano(vis)
    v["PENDIENTES"] = vis["_pend"].map(lambda lst: ", ".join(R.PLANO_HEADERS[p] for p in lst) or None).values
    v["FUENTE"] = vis["fuente"].values
    v = a_numero(v, [R.PLANO_HEADERS[c] for c in R.PLANO_NUMERICAS])
    v = a_objeto(v, [R.PLANO_HEADERS[c] for c in R.PLANO_TEXTO] + ["PENDIENTES", "FUENTE"])
    cfg = {R.PLANO_HEADERS[c]: st.column_config.NumberColumn(("✍️ " if c in R.PLANO_SIEMPRE_EDITABLES else "") + R.PLANO_HEADERS[c])
           for c in R.PLANO_NUMERICAS}
    cfg.update({R.PLANO_HEADERS["tipo_vuelo"]: st.column_config.SelectboxColumn(options=["R", "A", "C"]),
                "PENDIENTES": st.column_config.TextColumn("⚠️ PENDIENTES"), "NRO": st.column_config.NumberColumn(format="%d")})
    mapa = {R.PLANO_HEADERS[c]: c for c in R.PLANO_CLAVES if c != "nro"}
    editor_auditado(v, key=f"pl_ed_{desde}_{hasta}", tabla="operaciones_plano", mapa=mapa, modulo="OPERACIONES", usuario=u,
                    column_config=cfg, height=460)
    with st.expander("🟨 Mapa de celdas pendientes (amarillo = dato manual / fórmula pendiente)"):
        mapa_v = v.drop(columns=["ID", "PENDIENTES", "FUENTE"])
        _tabla(U.resaltar(mapa_v, vacias_cols=list(mapa_v.columns)), hide_index=True, width="stretch", height=360)

    with st.expander("🔎 Visor lateral de soportes del vuelo (libro · ficha · combustible)"):
        etiquetas = {int(r.id): f"{r.fecha_local} · {r.matricula} · {r.hora_utc or '----'} UTC · {r.origen}-{r.destino} · {r.nro_vuelo or ''}"
                     for r in vis.itertuples()}
        if etiquetas:
            sel = st.selectbox("Vuelo", list(etiquetas), format_func=etiquetas.get, key="pl_vis")
            fila = vis[vis["id"] == sel].iloc[0]
            lv = fv = None
            if not R.es_vacio(fila.get("tramo_id")):
                r = db.consulta("SELECT d.ruta_archivo FROM tramos_vuelo t JOIN libros_vuelo l ON l.id=t.libro_id "
                                "JOIN documentos d ON d.id=l.documento_id WHERE t.id=?", (int(fila["tramo_id"]),))
                lv = r[0]["ruta_archivo"] if r else None
            if not R.es_vacio(fila.get("ficha_id")):
                r = db.consulta("SELECT d.ruta_archivo FROM fichas_vuelo f JOIN documentos d ON d.id=f.documento_id WHERE f.id=?",
                                (int(fila["ficha_id"]),))
                fv = r[0]["ruta_archivo"] if r else None
            r = db.consulta("SELECT d.ruta_archivo FROM combustible_tiquetes t JOIN documentos d ON d.id=t.documento_id "
                            "WHERE t.eliminado=0 AND t.fecha_consumo=? AND t.matricula=? LIMIT 1",
                            (fila["fecha_local"], R.norm_matricula(fila["matricula"])))
            cb = r[0]["ruta_archivo"] if r else None
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("**Libro de vuelo**")
                visor(lv)
            with c2:
                st.markdown("**Ficha / cierre**")
                visor(fv)
            with c3:
                st.markdown("**Tiquete combustible**")
                visor(cb)
            if R.texto(fila.get("notas")):
                st.caption("Notas del motor de reglas: " + fila["notas"])

    U.titulo("Exportación")
    filas_exp = R.ordenar_plano(db._lista(vis.drop(columns=["_pend"])))
    df_exp = db.vista_plano(pd.DataFrame(filas_exp) if filas_exp else pd.DataFrame(), con_id=False)
    if tipos or mats or solo_pend:
        st.caption("La exportación respeta los filtros activos (NRO se renumera sobre las filas exportadas).")
    oficiales_pend = int(vis["_pend"].map(lambda lst: bool([p for p in lst if p in R.PLANO_OFICIALES])).sum())
    if oficiales_pend:
        st.warning(f"{oficiales_pend} fila(s) tienen celdas oficiales (A..S) en blanco. Complételas antes de enviar a la Aerocivil.")
    periodo_txt = f"{desde:%d/%m/%Y} - {hasta:%d/%m/%Y}"
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ Excel Control Interno (27 columnas A..AA)", X.excel_plano_control(df_exp, periodo_txt, db.get_param("plano_sigla", "3CA")),
                       f"Moon_Plano_Control_Interno_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx", XLSX, key="pl_dl1", width="stretch", type="primary")
    c2.download_button("⬇️ Archivo Plano Limpio Aerocivil (19 columnas A..S)", X.excel_plano_aerocivil(df_exp),
                       f"Moon_Plano_Aerocivil_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx", XLSX, key="pl_dl2", width="stretch")
    _plano_extras(u, desde, hasta, pl)


def _plano_extras(u, desde, hasta, pl):
    with st.expander("➕ Agregar vuelo manual (contingencia sin libro de vuelo)"):
        c1, c2, c3, c4 = st.columns(4)
        datos = {"fecha_local": c1.date_input("Fecha (hora Colombia)", db.hoy(), key="plm_f", format="DD/MM/YYYY"),
                 "hora_local": c2.text_input("Hora salida local (HH:MM)", key="plm_h"),
                 "matricula": c3.text_input("Matrícula", key="plm_m"), "nro_vuelo": c4.text_input("Nro vuelo", key="plm_n"),
                 "origen": c1.text_input("Origen", key="plm_o"), "destino": c2.text_input("Destino", key="plm_d"),
                 "pax_bordo": c3.text_input("Pasajeros a bordo", key="plm_p"), "horas_bloque": c4.text_input("Horas bloque (H:MM)", key="plm_hb"),
                 "horas_vuelo": c1.text_input("Horas vuelo (H:MM)", key="plm_hv"), "tipo_vuelo": c2.selectbox("Tipo", ["", "R", "A", "C"], key="plm_t"),
                 "carga_bordo": c3.text_input("Carga a bordo (kg)", key="plm_cb"), "carga_real": c4.text_input("Carga real (kg)", key="plm_cr")}
        modelo = db.modelo_de_matricula(datos["matricula"]) or {}
        nro, charter = R.normalizar_nro_vuelo(datos["nro_vuelo"], datos["matricula"])
        ruta = db.rutas_dict().get(nro) if nro else None
        datos.update({"tip": db.get_param("plano_tip", "TE"), "sigla": db.get_param("plano_sigla", "3CA"),
                      "tipo_equipo": modelo.get("codigo"), "sillas_ofrecidas": modelo.get("sillas_ofrecidas"),
                      "carga_ofrecida": modelo.get("carga_ofrecida_kg"), "tipo_trafico": modelo.get("tipo_trafico_default"),
                      "uso_aeronave": modelo.get("uso_aeronave_default"), "tipo_combustible": modelo.get("tipo_combustible_default"),
                      "metodo_medicion": modelo.get("metodo_medicion_default"), "correo_bordo": 0, "pax_transito": 0, "carga_transito": 0,
                      "nro_vuelo": nro, "tipo_vuelo": datos["tipo_vuelo"] or ("C" if charter else (ruta or {}).get("tipo_vuelo")),
                      "origen": datos["origen"] or (ruta or {}).get("origen"), "destino": datos["destino"] or (ruta or {}).get("destino"),
                      "combustible_consumido": (ruta or {}).get("combustible_estandar") if not charter else None})
        pulsado, _ = guardar_con_clausula(u, "plm", "Agregar vuelo al plano", motivo=False)
        if pulsado:
            ok, msg = db.crear_fila_plano_manual(datos, u)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
    with st.expander("⛽ Guardar ruta chárter frecuente (combustible estándar)"):
        st.caption("Para rutas chárter repetitivas: si el manifiesto no trae combustible básico, se usa este valor.")
        c1, c2, c3, c4 = st.columns(4)
        o, dd = c1.text_input("Origen", key="rc_o"), c2.text_input("Destino", key="rc_d")
        g = c3.text_input("Galones estándar", key="rc_g")
        mods = [""] + db.df_modelos()["codigo"].tolist()
        m = c4.selectbox("Modelo (opcional)", mods, key="rc_m")
        pulsado, motivo = guardar_con_clausula(u, "rc", "Guardar ruta chárter")
        if pulsado:
            ok, msg = db.guardar_ruta_charter(o, dd, g, m or None, u, motivo)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
        rc = db.df_rutas_charter()
        if not rc.empty:
            _tabla(rc[["origen", "destino", "modelo_codigo", "combustible_galones", "fecha_actualizacion"]], hide_index=True, width="stretch")
    with st.expander("📥 Importar histórico desde plantilla MO-A3-013 (hoja MES)"):
        st.caption("Carga los vuelos ya digitados en la plantilla Excel. No se duplican filas existentes.")
        archivo = st.file_uploader("Plantilla .xlsx", type=["xlsx"], key="pl_imp")
        local = db.buscar_plantilla_xlsx()
        if local and not archivo:
            st.caption(f"También se puede usar el archivo del servidor: {local.name}")
        pulsado, _ = guardar_con_clausula(u, "pl_imp_g", "Importar histórico", motivo=False, extra_ok=bool(archivo or local))
        if pulsado:
            fuente = archivo.getvalue() if archivo else local
            try:
                with st.spinner("Importando…"):
                    r = db.importar_mes_plano(fuente, u)
                flash(r["mensaje"])
            except Exception as exc:
                flash("", [f"No se pudo importar la plantilla: {exc}"])
            st.rerun()
    if not pl.empty:
        bloque_eliminar("operaciones_plano", pl["id"].astype(int).tolist(),
                        {int(r.id): f"#{r.id} · {r.fecha_local} · {r.matricula} · {r.origen}-{r.destino} · {r.nro_vuelo or ''}" for r in pl.itertuples()},
                        u, "OPERACIONES", "pl_del")


def tab_fichas(u):
    desde, hasta = rango_fechas("fi_rg")
    fichas = db.df_fichas(desde, hasta)
    if fichas.empty:
        U.vacio("🧾", "Sin fichas en el periodo", "Los pilotos o el despacho las cargan desde la captura documental.")
        return
    v = db.vista_fichas(fichas)
    v["PAX A BORDO (cálculo)"] = [R.pax_a_bordo(r["PAX TOTAL"], r["TRIPADIS"], r["BEBÉS"], r["PAX TRÁNSITO"])[0] for _, r in v.iterrows()]
    U.kpi_fila([{"label": "Fichas", "value": f"{len(v)}"}, {"label": "Pax totales", "value": U.num(v["PAX TOTAL"].sum())},
                {"label": "Tripadis", "value": U.num(v["TRIPADIS"].sum())}, {"label": "Bebés", "value": U.num(v["BEBÉS"].sum())},
                {"label": "Tránsitos", "value": U.num(v["PAX TRÁNSITO"].sum())}])
    v = a_objeto(v, ["MATRICULA", "VUELO", "ORIGEN", "DESTINO", "HORA", "PILOTO", "OBSERVACIONES", "SOPORTE"])
    editor_auditado(v, key="fi_ed", tabla="fichas_vuelo", modulo="OPERACIONES", usuario=u,
                    mapa={"FECHA": "fecha", "MATRICULA": "matricula", "VUELO": "numero_vuelo", "ORIGEN": "origen", "DESTINO": "destino",
                          "HORA": "hora_min", "PAX TOTAL": "pax_total", "TRIPADIS": "tripadis", "BEBÉS": "bebes",
                          "PAX TRÁNSITO": "pax_transito", "CARGA A BORDO (KG)": "carga_a_bordo_kg", "CARGA TRÁNSITO (KG)": "carga_transito_kg",
                          "COMB. BÁSICO (LBS)": "combustible_basico_lbs", "PILOTO": "piloto_nombre", "OBSERVACIONES": "observaciones"})
    st.caption("Después de corregir fichas, pulse 'Generar / actualizar desde soportes' en el archivo plano para recalcular pasajeros.")
    with st.expander("🔎 Ver soporte de una ficha"):
        etiquetas = {int(r.id): f"{r.fecha} · {r.matricula} · {r.origen}-{r.destino}" for r in fichas.itertuples()}
        sel = st.selectbox("Ficha", list(etiquetas), format_func=etiquetas.get, key="fi_vis")
        visor(fichas.loc[fichas["id"] == sel, "ruta_archivo"].iloc[0])
    bloque_eliminar("fichas_vuelo", fichas["id"].astype(int).tolist(),
                    {int(r.id): f"#{r.id} · {r.fecha} · {r.matricula} · {r.origen}-{r.destino}" for r in fichas.itertuples()}, u, "OPERACIONES", "fi_del")


# ============================================================
# MÓDULO 4 · OMA (mantenimiento)
# ============================================================
def modulo_oma(u):
    requiere_rol("OMA", "GERENCIA")
    U.page_title("OMA · Mantenimiento aéreo")
    U.hero("Aeronavegabilidad con semáforo preventivo.",
           "Reportes, acciones correctivas y trazabilidad de componentes tomados del bloque inferior del libro de vuelo MO-A3-004.",
           "Horas · ciclos · próxima inspección")
    p = db.params_dict()
    aer = db.df_aeronaves()
    mant = db.df_mantenimiento()
    abiertos = int((mant["estado"] == "ABIERTO").sum()) if not mant.empty else 0
    rii = int((mant["rii"] == "SI").sum()) if not mant.empty else 0
    criticas = 0
    if not aer.empty:
        rest = pd.to_numeric(aer["proxima_inspeccion_horas"], errors="coerce") - pd.to_numeric(aer["horas_totales"], errors="coerce")
        criticas = int((rest <= float(p.get("oma_umbral_amarillo_h") or 50)).sum())
    U.kpi_fila([{"label": "Aeronaves", "value": f"{len(aer)}"}, {"label": "Ítems abiertos", "value": f"{abiertos}", "tono": "alerta" if abiertos else "bien"},
                {"label": "Ítems con RII", "value": f"{rii}"},
                {"label": "Inspecciones próximas", "value": f"{criticas}", "tono": "mal" if criticas else "bien"}])
    U.titulo("Semáforo de flota", "Verde > umbral amarillo · Amarillo ≤ umbral · Rojo ≤ umbral rojo (horas restantes, parametrizable).")
    U.tarjetas_flota(aer, float(p.get("oma_umbral_amarillo_h") or 50), float(p.get("oma_umbral_rojo_h") or 10))
    sec = st.segmented_control("Sección", ["🛠️ Ítems de mantenimiento", "⏲️ Horas, ciclos e inspecciones", "➕ Registrar ítem",
                                           "🔩 Trazabilidad de componentes"], default="🛠️ Ítems de mantenimiento", key="oma_sec",
                               required=True, label_visibility="collapsed")
    if sec.startswith("🛠️"):
        c1, c2 = st.columns(2)
        mat = c1.selectbox("Matrícula", ["Todas"] + aer["matricula"].tolist(), key="oma_mat")
        est = c2.segmented_control("Estado", ["Todos", "ABIERTO", "CERRADO"], default="ABIERTO", key="oma_est", required=True)
        df = db.df_mantenimiento(None if mat == "Todas" else mat, None if est == "Todos" else est)
        if df.empty:
            U.vacio("🛠️", "Sin ítems", "Los reportes del libro de vuelo aparecen aquí automáticamente.")
            return
        v = pd.DataFrame({"ID": df["id"], "FECHA": df["fecha"], "MATRICULA": df["matricula"], "LIBRO Nº": df["numero_libro"],
                          "ITEM": df["item"], "REPORTE": df["reporte"], "ACCIÓN CORRECTIVA": df["accion_correctiva"],
                          "P/N ON": df["pn_on"], "S/N ON": df["sn_on"], "P/N OFF": df["pn_off"], "S/N OFF": df["sn_off"],
                          "POSICIÓN": df["posicion"], "RII": df["rii"], "TÉCNICO": df["tecnico_nombre"], "LICENCIA TÉC.": df["tecnico_licencia"],
                          "BASE": df["base"], "ESTADO": df["estado"]})
        v = a_objeto(v, list(v.columns.drop("ID")))
        editor_auditado(v, key="oma_ed", tabla="mantenimiento_oma", modulo="OMA", usuario=u,
                        column_config={"ESTADO": st.column_config.SelectboxColumn(options=["ABIERTO", "CERRADO"]),
                                       "RII": st.column_config.SelectboxColumn(options=["SI", "NO"])},
                        mapa={"FECHA": "fecha", "ITEM": "item", "REPORTE": "reporte", "ACCIÓN CORRECTIVA": "accion_correctiva",
                              "P/N ON": "pn_on", "S/N ON": "sn_on", "P/N OFF": "pn_off", "S/N OFF": "sn_off", "POSICIÓN": "posicion",
                              "RII": "rii", "TÉCNICO": "tecnico_nombre", "LICENCIA TÉC.": "tecnico_licencia", "BASE": "base", "ESTADO": "estado"})
        with st.expander("🔎 Ver la hoja del libro de vuelo del ítem"):
            etiquetas = {int(r.id): f"#{r.id} · {r.fecha} · {r.matricula} · {R.texto(r.reporte)[:60]}" for r in df.itertuples()}
            sel = st.selectbox("Ítem", list(etiquetas), format_func=etiquetas.get, key="oma_vis")
            visor(df.loc[df["id"] == sel, "ruta_archivo"].iloc[0])
    elif sec.startswith("⏲️"):
        if aer.empty:
            U.vacio("✈️", "Sin aeronaves registradas", "Gerencia registra la flota en Parametrización.")
            return
        mat = st.selectbox("Aeronave", aer["matricula"].tolist(), key="oma_hc_mat")
        a = aer[aer["matricula"] == mat].iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        datos = {"matricula": mat, "horas_totales": c1.text_input("Horas totales", R.texto(a["horas_totales"]), key=f"oma_ht_{mat}"),
                 "ciclos_totales": c2.text_input("Ciclos totales", R.texto(a["ciclos_totales"]), key=f"oma_ct_{mat}"),
                 "proxima_inspeccion_horas": c3.text_input("Próxima inspección (horas)", R.texto(a["proxima_inspeccion_horas"]), key=f"oma_pih_{mat}"),
                 "proxima_inspeccion_ciclos": c4.text_input("Próxima inspección (ciclos)", R.texto(a["proxima_inspeccion_ciclos"]), key=f"oma_pic_{mat}")}
        modelo = db.modelos_dict().get(a["modelo_codigo"]) or {}
        if modelo.get("ciclo_inspeccion_horas") and R.a_float(datos["horas_totales"]) and not R.a_float(datos["proxima_inspeccion_horas"]):
            st.info(f"Sugerencia: próxima inspección a {R.a_float(datos['horas_totales']) + modelo['ciclo_inspeccion_horas']:,.1f} h "
                    f"(ciclo del modelo {modelo['codigo']}: {modelo['ciclo_inspeccion_horas']:,.0f} h).")
        pulsado, motivo = guardar_con_clausula(u, f"oma_hc_{mat}", "Actualizar horas / ciclos / inspección")
        if pulsado:
            ok, msg = db.guardar_aeronave(datos, u, motivo, "OMA")
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
        rest = (pd.to_numeric(aer["proxima_inspeccion_horas"], errors="coerce") - pd.to_numeric(aer["horas_totales"], errors="coerce"))
        g = pd.DataFrame({"MATRICULA": aer["matricula"], "HORAS RESTANTES": rest}).dropna()
        if not g.empty:
            am, rj = float(p.get("oma_umbral_amarillo_h") or 50), float(p.get("oma_umbral_rojo_h") or 10)
            g["COLOR"] = g["HORAS RESTANTES"].map(lambda x: U.BAD if x <= rj else (U.WARN if x <= am else U.OK))
            fig = go.Figure(go.Bar(x=g["HORAS RESTANTES"], y=g["MATRICULA"], orientation="h", marker_color=g["COLOR"],
                                   text=g["HORAS RESTANTES"].map(lambda x: f"{x:,.1f} h"), textposition="outside"))
            fig.update_layout(title="Horas restantes para la próxima inspección")
            U.mostrar(U.estilo(fig, 320))
    elif sec.startswith("➕"):
        c1, c2, c3 = st.columns(3)
        datos = {"fecha": c1.date_input("Fecha", db.hoy(), key="omn_f", format="DD/MM/YYYY"),
                 "matricula": c2.selectbox("Matrícula", aer["matricula"].tolist() or [""], key="omn_m"),
                 "item": c3.text_input("Ítem", key="omn_i"), "reporte": st.text_area("Reporte", key="omn_r"),
                 "accion_correctiva": st.text_area("Acción correctiva", key="omn_a")}
        c1, c2, c3, c4 = st.columns(4)
        datos.update({"pn_on": c1.text_input("P/N ON", key="omn_pn1"), "sn_on": c2.text_input("S/N ON", key="omn_sn1"),
                      "pn_off": c3.text_input("P/N OFF", key="omn_pn2"), "sn_off": c4.text_input("S/N OFF", key="omn_sn2"),
                      "posicion": c1.selectbox("Posición", ["", "LH", "RH", "N/A"], key="omn_pos"),
                      "rii": c2.selectbox("RII", ["", "SI", "NO"], key="omn_rii"),
                      "tecnico_nombre": c3.text_input("Técnico", key="omn_tec"), "tecnico_licencia": c4.text_input("Licencia técnico", key="omn_lic"),
                      "base": c1.text_input("Base", key="omn_base"), "estado": c2.selectbox("Estado", ["ABIERTO", "CERRADO"], key="omn_est")})
        pulsado, _ = guardar_con_clausula(u, "omn", "Registrar ítem de mantenimiento", motivo=False,
                                          extra_ok=bool(R.texto(datos["reporte"]) and datos["matricula"]))
        if pulsado:
            rid = db.registrar_item_oma(datos, u)
            flash(f"Ítem #{rid} registrado.")
            st.rerun()
    else:
        buscar = st.text_input("Buscar P/N o S/N", key="oma_bus", placeholder="Ej. 124T01-2")
        df = db.df_mantenimiento()
        if not df.empty:
            df = df[df[["pn_on", "pn_off"]].notna().any(axis=1) | df[["sn_on", "sn_off"]].notna().any(axis=1)]
        if buscar and not df.empty:
            b = buscar.strip().upper()
            df = df[df[["pn_on", "sn_on", "pn_off", "sn_off"]].fillna("").apply(lambda r: any(b in str(x).upper() for x in r), axis=1)]
        if df.empty:
            U.vacio("🔩", "Sin cambios de componentes", "Los cambios de P/N y S/N aparecen al registrar libros de vuelo con mantenimiento.")
        else:
            _tabla(df[["fecha", "matricula", "numero_libro", "item", "pn_on", "sn_on", "pn_off", "sn_off", "posicion", "rii",
                             "tecnico_nombre", "estado"]].rename(columns=str.upper), hide_index=True, width="stretch")


# ============================================================
# MÓDULO 5 · Gerencia general
# ============================================================
def modulo_gerencia(u):
    requiere_rol("GERENCIA")
    U.page_title("Gerencia general")
    sec = st.segmented_control("Sección", ["📈 Dashboard ejecutivo", "🏅 Desempeño de empleados", "🛡️ Auditoría general IA",
                                           "⚙️ Parametrización", "📦 Consolidados", "📜 Bitácora inmutable"],
                               default="📈 Dashboard ejecutivo", key="ger_sec", required=True, label_visibility="collapsed")
    if sec.startswith("📈"):
        tab_dashboard(u)
    elif sec.startswith("🏅"):
        tab_desempeno(u)
    elif sec.startswith("🛡️"):
        tab_auditoria_general(u)
    elif sec.startswith("⚙️"):
        tab_parametrizacion(u)
    elif sec.startswith("📦"):
        area = st.segmented_control("Área", list(db.AREAS_CIERRE), default="CONTABILIDAD", key="ger_cons_area", required=True)
        render_consolidados(area, u, f"gc_{area}")
    else:
        tab_bitacora(u)


def _filtrar_vuelos(dv: pd.DataFrame, filtros: dict) -> pd.DataFrame:
    for col, vals in filtros.items():
        if vals and col in dv:
            dv = dv[dv[col].isin(vals)]
    return dv


def _evolucion(dv, dc, metrica):
    if metrica == "Costo combustible":
        if dc.empty or dc["VALOR_TOTAL"].notna().sum() == 0:
            return None
        g = dc.groupby("MES", as_index=False)["VALOR_TOTAL"].sum().rename(columns={"VALOR_TOTAL": "VALOR"})
    else:
        col = {"Galones": "GAL", "Vuelos": "VUELOS", "Pasajeros": "PAX"}[metrica]
        if dv.empty:
            return None
        g = dv.groupby("MES", as_index=False)[col].sum().rename(columns={col: "VALOR"})
    fig = px.bar(g, x="MES", y="VALOR", text="VALOR", title=f"Evolución intermensual · {metrica}", color_discrete_sequence=[U.BLUE])
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_xaxes(title=None, type="category")
    fig.update_yaxes(title=None)
    return U.estilo(fig, 360)


def _interanual(dv_all, dc_all, metrica):
    if metrica == "Costo combustible":
        if dc_all.empty:
            return None
        g = dc_all.groupby(["ANIO", "MES_NUM"], as_index=False)["VALOR_TOTAL"].sum().rename(columns={"VALOR_TOTAL": "VALOR"})
    else:
        col = {"Galones": "GAL", "Vuelos": "VUELOS", "Pasajeros": "PAX"}[metrica]
        if dv_all.empty:
            return None
        g = dv_all.groupby(["ANIO", "MES_NUM"], as_index=False)[col].sum().rename(columns={col: "VALOR"})
    if len(g) < 2:
        return None
    g["MES"] = g["MES_NUM"].map(lambda m: R.MESES_ES[int(m)][:3])
    fig = px.line(g.sort_values("MES_NUM"), x="MES", y="VALOR", color="ANIO", markers=True, title=f"Comparativo interanual · {metrica}",
                  color_discrete_sequence=[U.NAVY, U.AMBER, U.BLUE, U.SKY, U.GOLD])
    fig.update_xaxes(title=None, categoryorder="array", categoryarray=[R.MESES_ES[i][:3] for i in range(1, 13)])
    fig.update_yaxes(title=None)
    return U.estilo(fig, 360)


def tab_dashboard(u):
    desde, hasta = rango_fechas("dash_rg", defecto="Mes", opciones=("7 días", "30 días", "Mes", "Año", "Histórico completo", "Personalizado"))
    dv0 = db.dataset_vuelos(desde, hasta)
    dc0 = db.dataset_combustible(desde, hasta)
    if dv0.empty and dc0.empty:
        U.vacio("📈", "Sin datos en el periodo", "Importe la plantilla MO-A3-013 en Operaciones o registre soportes para alimentar el tablero.")
        return
    with st.container(border=True):
        f = st.columns(6)
        opciones = lambda c: sorted(dv0[c].dropna().unique()) if not dv0.empty and c in dv0 else []
        filtros = {"MODELO": f[0].multiselect("Modelo de avión", placeholder="Seleccione…", options=opciones("MODELO"), key="df_mod"),
                   "MATRICULA": f[1].multiselect("Matrícula", placeholder="Seleccione…", options=opciones("MATRICULA"), key="df_mat"),
                   "TIPO_VUELO": f[2].multiselect("Tipo de vuelo", placeholder="Seleccione…", options=opciones("TIPO_VUELO"), key="df_tipo"),
                   "RUTA": f[3].multiselect("Ruta", placeholder="Seleccione…", options=opciones("RUTA"), key="df_ruta"),
                   "ORIGEN": f[4].multiselect("Aeropuerto / base", placeholder="Seleccione…", options=opciones("ORIGEN"), key="df_base"),
                   "PILOTO": f[5].multiselect("Piloto", placeholder="Seleccione…", options=opciones("PILOTO"), key="df_pil")}
    dv = _filtrar_vuelos(dv0, filtros)
    dc = dc0.copy()
    if not dc.empty:
        if filtros["MATRICULA"]:
            dc = dc[dc["MATRICULA"].isin(filtros["MATRICULA"])]
        if filtros["ORIGEN"]:
            dc = dc[dc["ESTACION"].isin(filtros["ORIGEN"])]
        if filtros["PILOTO"]:
            dc = dc[dc["PILOTO"].isin(filtros["PILOTO"])]
    valorizados = dc[dc["VALOR_TOTAL"].notna()] if not dc.empty else dc
    precio = R.safe_div(valorizados["VALOR_TOTAL"].sum(), valorizados["GALONES"].sum(), None) if not valorizados.empty else None
    vuelos = len(dv)
    pax = dv["PAX"].sum() if not dv.empty else 0
    sillas = dv["SILLAS"].sum() if not dv.empty else 0
    ocup = R.safe_div(pax, sillas, None)
    gal = dv["GAL"].sum() if not dv.empty else 0
    costo = dc["VALOR_TOTAL"].sum(min_count=1) if not dc.empty else None
    horas = dv["MIN_BLOQUE"].sum() / 60 if not dv.empty else 0
    cop_pax = R.safe_div(gal * precio, pax, None) if precio else None

    previo = None
    if (hasta - desde).days < 3000:
        dias = (hasta - desde).days + 1
        previo = _filtrar_vuelos(db.dataset_vuelos(desde - timedelta(days=dias), desde - timedelta(days=1)), filtros)
    var = (lambda a, b: (a - b) / b * 100 if b else None)
    U.kpi_fila([{"label": "Vuelos", "value": f"{vuelos:,}", "delta": var(vuelos, len(previo)) if previo is not None else None},
                {"label": "Pasajeros a bordo", "value": U.num(pax), "delta": var(pax, previo["PAX"].sum()) if previo is not None and not previo.empty else None},
                {"label": "Factor de ocupación", "value": f"{U.num(ocup * 100, 1)} %" if ocup is not None else "—", "tono": "oro"},
                {"label": "Horas bloque", "value": U.num(horas, 1)}])
    U.kpi_fila([{"label": "Galones (real o estándar)", "value": U.num(gal, 0),
                 "delta": var(gal, previo["GAL"].sum()) if previo is not None and not previo.empty else None},
                {"label": "Costo combustible facturado", "value": U.cop(costo)},
                {"label": "Precio efectivo", "value": f"{U.cop(precio)}/gal" if precio else "Sin precios",
                 "nota": None if precio else "Valorice tiquetes en Contabilidad"},
                {"label": "Costo combustible por pasajero", "value": U.cop(cop_pax) if cop_pax else "—", "nota": "galones × precio efectivo ÷ pax"}])

    U.titulo("Evolución histórica")
    metrica = st.segmented_control("Métrica", ["Galones", "Costo combustible", "Vuelos", "Pasajeros"], default="Galones", key="dash_met",
                                   required=True, label_visibility="collapsed")
    dv_all = _filtrar_vuelos(db.dataset_vuelos(None, None), filtros)
    dc_all = db.dataset_combustible(None, None)
    g1, g2 = st.columns(2)
    with g1:
        U.mostrar(_evolucion(dv_all if (hasta - desde).days > 60 else dv, dc_all if (hasta - desde).days > 60 else dc, metrica),
                  "Sin datos para la métrica.", key="g_evo")
    with g2:
        U.mostrar(_interanual(dv_all, dc_all, metrica), "Se necesitan datos de más de un año para comparar.", key="g_inter")

    U.titulo("Combustible")
    g1, g2 = st.columns(2)
    with g1:
        U.mostrar(U.barras_h(dv[dv["PILOTO"] != "SIN DATO"] if not dv.empty else dv, "PILOTO", "GAL", "Pilotos que más combustible consumen (gal)",
                             ",.0f"), "Requiere libros de vuelo con piloto identificado.", key="g_pil")
    with g2:
        rv = dv.dropna(subset=["GAL_REAL", "GAL_ESTANDAR"]) if not dv.empty else dv
        fig = None
        if not rv.empty:
            g = rv.groupby("RUTA", as_index=False)[["GAL_REAL", "GAL_ESTANDAR"]].sum().nlargest(12, "GAL_ESTANDAR")
            g = g.melt(id_vars="RUTA", var_name="TIPO", value_name="GALONES").replace({"GAL_REAL": "Real (libro)", "GAL_ESTANDAR": "Estándar (BD)"})
            fig = U.estilo(px.bar(g, x="RUTA", y="GALONES", color="TIPO", barmode="group", title="Consumo real vs. estándar por ruta",
                                  color_discrete_sequence=[U.NAVY, U.AMBER]))
        U.mostrar(fig, "Requiere libros de vuelo con libras usadas y factor lbs/gal del modelo.", key="g_real")
    g1, g2 = st.columns(2)
    with g1:
        fig = None
        if not dv.empty and dv["MIN_VUELO"].sum() > 0:
            g = dv.groupby(["MATRICULA", "MODELO"], as_index=False).agg(GAL=("GAL", "sum"), MIN=("MIN_VUELO", "sum"))
            g = g[g["MIN"] > 0]
            g["GAL/H"] = g["GAL"] / (g["MIN"] / 60)
            fig = px.bar(g.sort_values("GAL/H"), x="GAL/H", y="MATRICULA", color="MODELO", orientation="h", text="GAL/H",
                         title="Eficiencia: galones por hora de vuelo", color_discrete_sequence=U.PALETA)
            fig.update_traces(texttemplate="%{text:,.1f}", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig)
        U.mostrar(fig, key="g_efi")
    with g2:
        fig = None
        if not valorizados.empty:
            g = valorizados.groupby("ESTACION", as_index=False).agg(V=("VALOR_TOTAL", "sum"), G=("GALONES", "sum"))
            g = g[g["G"] > 0]
            g["COP/GAL"] = g["V"] / g["G"]
            fig = px.bar(g.sort_values("COP/GAL"), x="COP/GAL", y="ESTACION", orientation="h", text="COP/GAL",
                         title="Precio efectivo COP/gal por estación", color_discrete_sequence=[U.GOLD])
            fig.update_traces(texttemplate="$ %{text:,.0f}", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig)
        U.mostrar(fig, "Valorice tiquetes en Contabilidad para ver precios por estación.", key="g_precio")

    U.titulo("Comercial y ocupación")
    dim = st.segmented_control("Pasajeros por", ["RUTA", "MATRICULA", "PILOTO"], default="RUTA", key="dash_dim", required=True,
                               format_func=lambda x: {"RUTA": "Ruta", "MATRICULA": "Aeronave", "PILOTO": "Piloto"}[x])
    g1, g2 = st.columns(2)
    with g1:
        base = dv[dv[dim] != "SIN DATO"] if not dv.empty else dv
        U.mostrar(U.barras_h(base, dim, "PAX", f"Pasajeros a bordo por {dim.lower()}", ",.0f", top=15), key="g_pax")
    with g2:
        fig = None
        if not dv.empty and dv["SILLAS"].sum() > 0:
            g = dv.groupby(["RUTA", "TIPO_VUELO"], as_index=False).agg(PAX=("PAX", "sum"), SILLAS=("SILLAS", "sum"), V=("VUELOS", "sum"))
            g = g[g["SILLAS"] > 0]
            top = g.groupby("RUTA")["V"].sum().nlargest(14).index
            g = g[g["RUTA"].isin(top)]
            g["OCUPACIÓN %"] = g["PAX"] / g["SILLAS"] * 100
            fig = px.bar(g, x="RUTA", y="OCUPACIÓN %", color="TIPO_VUELO", barmode="group", title="Factor de ocupación por ruta y tipo de vuelo",
                         color_discrete_map={"R": U.NAVY, "A": U.BLUE, "C": U.AMBER})
            fig.update_yaxes(ticksuffix=" %")
            fig = U.estilo(fig)
        U.mostrar(fig, key="g_ocup")
    g1, g2 = st.columns(2)
    with g1:
        fig = None
        if not dv.empty and dv["CARGA_OFRECIDA"].sum() > 0:
            g = dv.groupby("RUTA", as_index=False).agg(C=("CARGA_BORDO", "sum"), O=("CARGA_OFRECIDA", "sum"))
            g = g[g["O"] > 0]
            g["APROVECHAMIENTO %"] = g["C"] / g["O"] * 100
            g = g.nlargest(14, "APROVECHAMIENTO %").sort_values("APROVECHAMIENTO %")
            fig = px.bar(g, x="APROVECHAMIENTO %", y="RUTA", orientation="h", text="APROVECHAMIENTO %", title="Aprovechamiento de carga (%)",
                         color_discrete_sequence=[U.BLUE])
            fig.update_traces(texttemplate="%{text:.1f} %", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig)
        U.mostrar(fig, key="g_carga")
    with g2:
        fig = None
        if precio and not dv.empty:
            g = dv.groupby("RUTA", as_index=False).agg(GAL=("GAL", "sum"), PAX=("PAX", "sum"))
            g = g[g["PAX"] > 0]
            g["COP/PAX"] = g["GAL"] * precio / g["PAX"]
            g = g.nlargest(14, "COP/PAX").sort_values("COP/PAX")
            fig = px.bar(g, x="COP/PAX", y="RUTA", orientation="h", text="COP/PAX", title="Costo de combustible por pasajero (COP/pax)",
                         color_discrete_sequence=[U.NAVY])
            fig.update_traces(texttemplate="$ %{text:,.0f}", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig)
        U.mostrar(fig, "Requiere precios de combustible valorizados en el periodo.", key="g_coppax")

    U.titulo("Financiero tripulación", "Minutos liquidados (TIEMPO PILOTOS) vs. mínimo garantizado SMLV y total COP por tripulante.")
    periodo = selector_mes("dash_liq_mes", "Mes de liquidación")
    ini, fin = limites_mes(periodo)
    liq, det = db.datos_liquidacion(ini, fin, periodo)
    if liq.empty:
        st.caption("Sin liquidación para el mes (requiere libros de vuelo con tripulación y TIEMPO PILOTOS).")
    else:
        g1, g2 = st.columns(2)
        with g1:
            fig = go.Figure()
            fig.add_bar(x=liq["NOMBRE"], y=liq["MIN VOLADOS"], marker_color=U.NAVY, name="Minutos")
            smlv = db.get_param("min_smlv")
            if smlv:
                fig.add_hline(y=float(smlv), line_dash="dash", line_color=U.AMBER, annotation_text=f"SMLV {U.num(smlv)} min")
            fig.update_layout(title="Minutos por tripulante vs. mínimo garantizado")
            U.mostrar(U.estilo(fig), key="g_min")
        with g2:
            U.mostrar(U.barras_h(liq.assign(TOTAL=pd.to_numeric(liq["TOTAL A PAGAR (COP)"], errors="coerce").fillna(0)), "NOMBRE", "TOTAL",
                                 "Total COP liquidado por tripulante", "$,.0f"), "Aún no hay valores a pagar en el mes.", key="g_cop")


def tab_desempeno(u):
    st.caption("🔒 Exclusivo de Gerencia. La calidad usa los hallazgos de la auditoría más reciente: ejecute la Auditoría General para actualizarla.")
    desde, hasta = rango_fechas("des_rg")
    df = A.desempeno_empleados(desde, hasta)
    if df.empty:
        U.vacio("🏅", "Sin empleados evaluables", "Cree usuarios de Pilotos, Contabilidad y Operaciones en Parametrización.")
        return
    roles = st.pills("Rol", ["PILOTO", "CONTABILIDAD", "OPERACIONES"], selection_mode="multi", key="des_roles")
    if roles:
        df = df[df["ROL"].isin(roles)]
    activos = df[(df["REGISTROS CREADOS"] + df["EDICIONES"] + df["CIERRES"]) > 0]
    U.kpi_fila([{"label": "Empleados evaluados", "value": f"{len(df)}"}, {"label": "Con actividad", "value": f"{len(activos)}"},
                {"label": "Registros procesados", "value": U.num(df["REGISTROS CREADOS"].sum())},
                {"label": "% a tiempo (promedio)", "value": f"{U.num(df['% A TIEMPO'].dropna().mean(), 1)} %" if df["% A TIEMPO"].notna().any() else "—"},
                {"label": "Tasa de error (promedio)", "value": f"{U.num(df['TASA DE ERROR %'].dropna().mean(), 1)} %" if df["TASA DE ERROR %"].notna().any() else "—"}])
    g1, g2 = st.columns(2)
    with g1:
        m = df.melt(id_vars=["NOMBRE"], value_vars=["REGISTROS CREADOS", "EDICIONES", "CIERRES"], var_name="TIPO", value_name="N")
        m = m[m["N"] > 0]
        fig = px.bar(m, x="N", y="NOMBRE", color="TIPO", orientation="h", title="Productividad: registros, validaciones y cierres",
                     color_discrete_sequence=[U.NAVY, U.SKY, U.AMBER]) if not m.empty else None
        U.mostrar(U.estilo(fig, max(300, 70 + 30 * df["NOMBRE"].nunique())) if fig else None, key="d_prod")
    with g2:
        o = df.dropna(subset=["% A TIEMPO"])
        fig = None
        if not o.empty:
            fig = px.bar(o.sort_values("% A TIEMPO"), x="% A TIEMPO", y="NOMBRE", orientation="h", text="% A TIEMPO", color="ROL",
                         title="Oportunidad: % de registros y cierres a tiempo", color_discrete_sequence=U.PALETA)
            fig.update_traces(texttemplate="%{text:.0f} %", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig, max(300, 70 + 30 * len(o)))
        U.mostrar(fig, key="d_opo")
    g1, g2 = st.columns(2)
    with g1:
        c = df.dropna(subset=["TASA DE ERROR %"])
        fig = None
        if not c.empty:
            fig = px.bar(c.sort_values("TASA DE ERROR %"), x="TASA DE ERROR %", y="NOMBRE", orientation="h", text="TASA DE ERROR %",
                         title="Calidad: tasa de error detectada por la auditoría IA", color_discrete_sequence=[U.BAD])
            fig.update_traces(texttemplate="%{text:.1f} %", textposition="outside", cliponaxis=False)
            fig = U.estilo(fig, max(300, 70 + 30 * len(c)))
        U.mostrar(fig, key="d_cal")
    with g2:
        s = df[(df["REGISTROS CREADOS"] + df["EDICIONES"]) > 0]
        fig = None
        if not s.empty:
            fig = px.scatter(s, x="REGISTROS CREADOS", y="TASA DE ERROR %", size=s["REPROCESOS"] + 1, color="ROL", hover_name="NOMBRE",
                             title="Productividad vs. calidad (tamaño = reprocesos)", color_discrete_sequence=U.PALETA)
            fig = U.estilo(fig)
        U.mostrar(fig, key="d_scatter")
    _tabla(df, hide_index=True, width="stretch")
    st.download_button("⬇️ Descargar evaluación de desempeño", X.libro({"Desempeño": df}, {"Documento": "Evaluación de desempeño",
                                                                                        "Periodo": f"{desde} a {hasta}",
                                                                                        "Generado por": u["nombre"]}),
                       f"Moon_Desempeno_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx", XLSX, key="des_dl")
    with st.expander("🔎 Historial de acciones de un empleado"):
        sel = st.selectbox("Empleado", df["CEDULA"].tolist(), format_func=lambda c: f"{df.loc[df['CEDULA'] == c, 'NOMBRE'].iloc[0]} · {c}",
                           key="des_emp")
        _tabla(db.df_auditoria(desde, hasta, cedula=sel, limite=500), hide_index=True, width="stretch")


def ruta_evidencia_registro(tabla, rid) -> str | None:
    if R.es_vacio(rid) or not R.solo_digitos(rid):
        return None
    rid = int(R.solo_digitos(rid))
    q = {"combustible_tiquetes": "SELECT d.ruta_archivo r FROM combustible_tiquetes x JOIN documentos d ON d.id=x.documento_id WHERE x.id=?",
         "libros_vuelo": "SELECT d.ruta_archivo r FROM libros_vuelo x JOIN documentos d ON d.id=x.documento_id WHERE x.id=?",
         "tramos_vuelo": "SELECT d.ruta_archivo r FROM tramos_vuelo t JOIN libros_vuelo x ON x.id=t.libro_id JOIN documentos d ON d.id=x.documento_id WHERE t.id=?",
         "fichas_vuelo": "SELECT d.ruta_archivo r FROM fichas_vuelo x JOIN documentos d ON d.id=x.documento_id WHERE x.id=?",
         "operaciones_plano": "SELECT d.ruta_archivo r FROM operaciones_plano p JOIN tramos_vuelo t ON t.id=p.tramo_id JOIN libros_vuelo x "
                              "ON x.id=t.libro_id JOIN documentos d ON d.id=x.documento_id WHERE p.id=?",
         "documentos": "SELECT ruta_archivo r FROM documentos WHERE id=?"}.get(tabla)
    if not q:
        return None
    r = db.consulta(q, (rid,))
    return r[0]["r"] if r else None


def tab_auditoria_general(u):
    st.caption("Revisoría fiscal integral: imágenes de las 3 carpetas + tablas llenadas o editadas por pilotos y empleados.")
    desde, hasta = rango_fechas("ag_rg")
    ia_ok, ia_msg = V.ia_disponible()
    n = int(db.get_param("ia_max_imagenes_auditoria", 12) or 12)
    revisar = st.toggle(f"Re-leer imágenes originales con IA (hasta {n} por ejecución, priorizando las nunca auditadas)", value=ia_ok,
                        disabled=not ia_ok, key="ag_img")
    if not ia_ok:
        st.info("🤖 " + ia_msg + " Las reglas de auditoría y la verificación de integridad igual se ejecutan.")
    if st.button("🛡️ Ejecutar Auditoría General Integral IA", type="primary", key="ag_run"):
        barra = st.progress(0.0, "Auditando bases de datos…")
        res = A.ejecutar_auditoria("GENERAL", desde, hasta, u, revisar_imagenes=revisar,
                                   progreso=lambda f, t: barra.progress(min(1.0, float(f)), t))
        barra.empty()
        res["rango"] = (desde, hasta)
        ss["ag_res"] = res
    res = ss.get("ag_res")
    if res:
        h = res["hallazgos"]
        cuenta = h["SEVERIDAD"].value_counts() if not h.empty else {}
        U.kpi_fila([{"label": "Hallazgos", "value": f"{len(h)}"},
                    {"label": "Altos", "value": f"{cuenta.get('Alta', 0)}", "tono": "mal" if cuenta.get("Alta", 0) else "bien"},
                    {"label": "Medios", "value": f"{cuenta.get('Media', 0)}", "tono": "alerta" if cuenta.get("Media", 0) else ""},
                    {"label": "Bajos", "value": f"{cuenta.get('Baja', 0)}"},
                    {"label": "Imágenes re-leídas por IA", "value": f"{res.get('imagenes_revisadas', 0)}"}])
        with st.container(border=True):
            st.markdown("**Informe de revisoría" + (" (IA)" if res["uso_ia"] else " automático") + "**")
            st.text(res["resumen"])
        if not h.empty:
            g1, g2 = st.columns(2)
            with g1:
                g = h.groupby(["CATEGORIA", "SEVERIDAD"], as_index=False).size()
                fig = px.bar(g, x="size", y="CATEGORIA", color="SEVERIDAD", orientation="h", title="Hallazgos por categoría",
                             color_discrete_map={"Alta": U.BAD, "Media": U.WARN, "Baja": U.SKY})
                U.mostrar(U.estilo(fig, max(320, 60 + 28 * g["CATEGORIA"].nunique())), key="ag_cat")
            with g2:
                r = h.assign(RESP=h["RESPONSABLE_NOMBRE"].fillna("Sin responsable"), N=1)
                U.mostrar(U.barras_h(r, "RESP", "N", "Hallazgos por responsable del registro", ",.0f"), key="ag_resp")
            f1, f2 = st.columns(2)
            sev = f1.segmented_control("Severidad", ["Todos", "Alta", "Media", "Baja"], default="Todos", key="ag_sev", required=True)
            area = f2.segmented_control("Área", ["Todas", "CONTABILIDAD", "OPERACIONES", "GENERAL"], default="Todas", key="ag_area", required=True)
            vista = h if sev == "Todos" else h[h["SEVERIDAD"] == sev]
            vista = vista if area == "Todas" else vista[vista["AREA"] == area]
            _tabla(vista, hide_index=True, width="stretch", height=380)
            with st.expander("📜 Historial completo (audit_logs) de los registros con hallazgos"):
                _tabla(res["historial"], hide_index=True, width="stretch")
            with st.expander("🔎 Ver la evidencia de un hallazgo"):
                opciones = list(range(len(vista)))
                if opciones:
                    i = st.selectbox("Hallazgo", opciones, format_func=lambda j: f"{vista.iloc[j]['CATEGORIA']} · {vista.iloc[j]['TABLA']} "
                                                                                    f"#{vista.iloc[j]['REGISTRO_ID']}", key="ag_vis")
                    visor(ruta_evidencia_registro(vista.iloc[i]["TABLA"], vista.iloc[i]["REGISTRO_ID"]))
            d, hh = res["rango"]
            st.download_button("⬇️ Descargar informe de revisoría (Excel)", X.excel_auditoria("GENERAL", d, hh, vista, res["historial"], res["resumen"], u),
                               f"Moon_Revisoria_General_{d:%Y%m%d}_{hh:%Y%m%d}.xlsx", XLSX, key="ag_dl", type="primary")
        else:
            st.success("Sin hallazgos: bases de datos, evidencias y bitácora íntegras.")
    with st.expander("Auditorías anteriores"):
        _tabla(db.df_auditorias(), hide_index=True, width="stretch")


def tab_parametrizacion(u):
    sec = st.segmented_control("Parámetros", ["👥 Usuarios y roles", "✈️ Flota y modelos", "🗺️ Vuelos y rutas (BD)", "💲 Tarifas y fórmulas",
                                              "💾 Sistema y respaldos"], default="👥 Usuarios y roles", key="par_sec", required=True,
                               label_visibility="collapsed")
    if sec.startswith("👥"):
        _param_usuarios(u)
    elif sec.startswith("✈️"):
        _param_flota(u)
    elif sec.startswith("🗺️"):
        _param_rutas(u)
    elif sec.startswith("💲"):
        _param_tarifas(u)
    else:
        _param_sistema(u)


def _param_usuarios(u):
    if ss.get("clave_temporal"):
        ced, clave = ss["clave_temporal"]
        with st.container(border=True):
            st.success(f"Clave temporal para la cédula {ced}. Entréguela por un canal seguro: se pedirá cambiarla en el primer ingreso. "
                       "No se volverá a mostrar.")
            st.code(clave, language=None)
            if st.button("Ya la entregué · ocultar", key="par_ocultar"):
                del ss["clave_temporal"]
                st.rerun()
    df = db.df_usuarios()
    _tabla(df, hide_index=True, width="stretch", column_config={"activo": st.column_config.CheckboxColumn(),
                                                                       "forzar_cambio_clave": st.column_config.CheckboxColumn("debe cambiar clave")})
    with st.expander("➕ Crear usuario"):
        c1, c2, c3 = st.columns(3)
        ced = c1.text_input("Cédula", max_chars=12, key="nu_ced")
        nombre = c2.text_input("Nombre completo", key="nu_nom")
        rol = c3.selectbox("Rol", db.ROLES, key="nu_rol")
        cargo = c1.selectbox("Cargo de vuelo (pilotos)", ["", "PTL", "PCA"], key="nu_cargo", disabled=rol != "PILOTO")
        lic = c2.text_input("Licencia", key="nu_lic")
        email = c3.text_input("Correo (opcional)", key="nu_mail")
        pulsado, _ = guardar_con_clausula(u, "nu", "Crear usuario y generar clave temporal", motivo=False)
        if pulsado:
            ok, msg, temp = db.crear_usuario(ced, nombre, rol, u, cargo or None, lic, email)
            if ok:
                ss["clave_temporal"] = (R.solo_digitos(ced), temp)
                flash(msg)
            else:
                flash("", [msg])
            st.rerun()
    with st.expander("✏️ Editar, activar/desactivar o resetear clave"):
        if df.empty:
            return
        sel = st.selectbox("Usuario", df["cedula"].tolist(), format_func=lambda c: f"{df.loc[df['cedula'] == c, 'nombre'].iloc[0]} · {c}",
                           key="eu_sel")
        f = df[df["cedula"] == sel].iloc[0]
        c1, c2, c3 = st.columns(3)
        cambios = {"nombre": c1.text_input("Nombre", f["nombre"], key=f"eu_nom_{sel}"),
                   "rol": c2.selectbox("Rol", db.ROLES, index=db.ROLES.index(f["rol"]), key=f"eu_rol_{sel}"),
                   "cargo_vuelo": c3.selectbox("Cargo de vuelo", ["", "PTL", "PCA"], index=["", "PTL", "PCA"].index(R.texto(f["cargo_vuelo"]) if R.texto(f["cargo_vuelo"]) in ("PTL", "PCA") else ""),
                                               key=f"eu_cargo_{sel}") or None,
                   "licencia": c1.text_input("Licencia", R.texto(f["licencia"]), key=f"eu_lic_{sel}") or None,
                   "email": c2.text_input("Correo", R.texto(f["email"]), key=f"eu_mail_{sel}") or None,
                   "activo": 1 if c3.toggle("Cuenta activa", bool(R.a_int(f["activo"], 1)), key=f"eu_act_{sel}") else 0}
        pulsado, motivo = guardar_con_clausula(u, f"eu_{sel}", "Guardar cambios del usuario")
        if pulsado:
            ok, msg = db.actualizar_usuario(sel, cambios, u, motivo)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
        st.markdown("---")
        pulsado, motivo = guardar_con_clausula(u, f"rc_{sel}", "🔑 Resetear clave (genera temporal y desbloquea)", tipo="secondary")
        if pulsado:
            ok, msg, temp = db.resetear_clave(sel, u, motivo)
            if ok:
                ss["clave_temporal"] = (sel, temp)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()


def _texto_campo(col, etiqueta, valor, key):
    """Campo de parámetro que admite quedar en blanco (✍️ = dato manual / fórmula pendiente)."""
    inicial = "" if R.es_vacio(valor) else (str(valor).rstrip("0").rstrip(".") if isinstance(valor, float) else str(valor))
    return col.text_input(("✍️ " if not inicial else "") + etiqueta, inicial, key=key,
                          help="En blanco = dato manual / fórmula pendiente: el sistema no bloquea y pide el dato al usarlo.")


def _param_flota(u):
    U.titulo("Modelos de aeronave", "Los campos en blanco (✍️) se permiten: Operaciones los diligenciará a mano hasta que se parametricen.")
    modelos = db.df_modelos()
    _tabla(modelos, hide_index=True, width="stretch")
    sel = st.selectbox("Modelo", ["➕ Nuevo modelo"] + modelos["codigo"].tolist(), key="pf_mod")
    m = {} if sel == "➕ Nuevo modelo" else modelos[modelos["codigo"] == sel].iloc[0].to_dict()
    c1, c2, c3, c4 = st.columns(4)
    datos = {"codigo": c1.text_input("Código", R.texto(m.get("codigo")), key=f"pf_cod_{sel}", disabled=sel != "➕ Nuevo modelo"),
             "nombre": _texto_campo(c2, "Nombre", m.get("nombre"), f"pf_nom_{sel}"),
             "sillas_ofrecidas": _texto_campo(c3, "Sillas ofrecidas", m.get("sillas_ofrecidas"), f"pf_sil_{sel}"),
             "carga_ofrecida_kg": _texto_campo(c4, "Carga ofrecida (kg)", m.get("carga_ofrecida_kg"), f"pf_car_{sel}"),
             "tipo_trafico_default": _texto_campo(c1, "Tipo tráfico", m.get("tipo_trafico_default"), f"pf_tt_{sel}"),
             "uso_aeronave_default": _texto_campo(c2, "Uso aeronave", m.get("uso_aeronave_default"), f"pf_uso_{sel}"),
             "tipo_combustible_default": _texto_campo(c3, "Tipo combustible", m.get("tipo_combustible_default"), f"pf_tc_{sel}"),
             "metodo_medicion_default": _texto_campo(c4, "Método medición", m.get("metodo_medicion_default"), f"pf_mm_{sel}"),
             "reserva_combustible_lbs": _texto_campo(c1, "Reserva fija chárter (lbs)", m.get("reserva_combustible_lbs"), f"pf_res_{sel}"),
             "factor_conversion_lbs_gal": _texto_campo(c2, "Factor lbs/gal", m.get("factor_conversion_lbs_gal"), f"pf_fac_{sel}"),
             "ciclo_inspeccion_horas": _texto_campo(c3, "Ciclo de inspección (h)", m.get("ciclo_inspeccion_horas"), f"pf_ci_{sel}")}
    if sel == "➕ Nuevo modelo":
        datos["activo"] = 1
    st.caption("Fórmula chárter: Galones = (Combustible básico lbs del manifiesto + Reserva fija lbs) ÷ Factor lbs/gal. "
               "Referencia física: Jet A-1 ≈ 6,7 lb/gal (según API/temperatura del tiquete).")
    pulsado, motivo = guardar_con_clausula(u, f"pf_{sel}", "Guardar modelo", extra_ok=bool(R.texto(datos["codigo"])))
    if pulsado:
        ok, msg = db.guardar_modelo(datos, u, motivo)
        flash(msg if ok else "", [] if ok else [msg])
        st.rerun()

    U.titulo("Matrículas")
    aer = db.df_aeronaves()
    _tabla(aer, hide_index=True, width="stretch")
    sel = st.selectbox("Aeronave", ["➕ Nueva matrícula"] + aer["matricula"].tolist(), key="pf_aer")
    a = {} if sel == "➕ Nueva matrícula" else aer[aer["matricula"] == sel].iloc[0].to_dict()
    c1, c2, c3, c4 = st.columns(4)
    opciones_mod = [""] + modelos["codigo"].tolist()
    datos = {"matricula": c1.text_input("Matrícula", R.texto(a.get("matricula")), key=f"pa_mat_{sel}", disabled=sel != "➕ Nueva matrícula"),
             "modelo_codigo": c2.selectbox("Modelo", opciones_mod, index=opciones_mod.index(R.texto(a.get("modelo_codigo"))) if R.texto(a.get("modelo_codigo")) in opciones_mod else 0,
                                           key=f"pa_mod_{sel}") or None,
             "base": c3.text_input("Base", R.texto(a.get("base")), key=f"pa_base_{sel}"),
             "activo": 1 if c4.toggle("Activa", bool(R.a_int(a.get("activo"), 1)), key=f"pa_act_{sel}") else 0,
             "notas": st.text_input("Notas", R.texto(a.get("notas")), key=f"pa_not_{sel}")}
    pulsado, motivo = guardar_con_clausula(u, f"pa_{sel}", "Guardar aeronave", extra_ok=bool(R.texto(datos["matricula"])))
    if pulsado:
        ok, msg = db.guardar_aeronave(datos, u, motivo)
        flash(msg if ok else "", [] if ok else [msg])
        st.rerun()


def _param_rutas(u):
    rutas = db.df_rutas()
    st.caption(f"{len(rutas)} vuelos activos en el catálogo (equivalente a la hoja BD de la plantilla).")
    buscar = st.text_input("Buscar vuelo o ruta", key="pr_bus")
    vista = rutas
    if buscar:
        b = buscar.strip().upper()
        vista = rutas[rutas[["numero_vuelo", "ruta"]].fillna("").apply(lambda r: any(b in str(x).upper() for x in r), axis=1)]
    _tabla(vista[["numero_vuelo", "ruta", "origen", "destino", "hora", "tipo_vuelo", "combustible_estandar", "fuente", "fecha_actualizacion"]]
                 if not vista.empty else vista, hide_index=True, width="stretch", height=320)
    with st.expander("➕ Agregar o editar vuelo del catálogo"):
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        nro = c1.text_input("Número de vuelo", key="pr_nro")
        actual = rutas[rutas["numero_vuelo"] == R.solo_digitos(nro)].iloc[0].to_dict() if R.solo_digitos(nro) in rutas["numero_vuelo"].tolist() else {}
        datos = {"numero_vuelo": nro, "origen": c2.text_input("Origen", R.texto(actual.get("origen")), key=f"pr_o_{nro}"),
                 "destino": c3.text_input("Destino", R.texto(actual.get("destino")), key=f"pr_d_{nro}"),
                 "hora": c4.text_input("Hora (HH:MM)", R.texto(actual.get("hora")), key=f"pr_h_{nro}"),
                 "tipo_vuelo": c5.selectbox("Tipo", ["R", "A", "C"], index=["R", "A", "C"].index(R.texto(actual.get("tipo_vuelo")) if R.texto(actual.get("tipo_vuelo")) in ("R", "A", "C") else "R"), key=f"pr_t_{nro}"),
                 "combustible_estandar": c6.text_input("Fuel (gal)", "" if R.es_vacio(actual.get("combustible_estandar")) else str(actual.get("combustible_estandar")),
                                                       key=f"pr_f_{nro}")}
        if actual:
            st.caption("El vuelo existe: se actualizará con trazabilidad.")
        pulsado, motivo = guardar_con_clausula(u, "pr_g", "Guardar vuelo", extra_ok=bool(R.solo_digitos(nro)))
        if pulsado:
            ok, msg = db.guardar_ruta(datos, u, motivo)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
    with st.expander("🚫 Desactivar vuelo del catálogo"):
        nro = st.selectbox("Vuelo", rutas["numero_vuelo"].tolist(), key="pr_des") if not rutas.empty else None
        pulsado, motivo = guardar_con_clausula(u, "pr_d", "Desactivar", tipo="secondary", extra_ok=bool(nro))
        if pulsado:
            ok, msg = db.desactivar_ruta(nro, u, motivo)
            flash(msg if ok else "", [] if ok else [msg])
            st.rerun()
    with st.expander("📥 Importar / actualizar catálogo desde plantilla (hoja BD)"):
        archivo = st.file_uploader("Plantilla .xlsx con hoja BD", type=["xlsx"], key="pr_imp")
        reemplazar = st.toggle("Reemplazar vuelos existentes con los valores del archivo", key="pr_rep")
        pulsado, _ = guardar_con_clausula(u, "pr_i", "Importar catálogo", motivo=False, extra_ok=archivo is not None)
        if pulsado:
            try:
                n = db.importar_catalogo_bd(archivo.getvalue(), u, reemplazar)
                flash(f"{n} vuelo(s) importados/actualizados.")
            except Exception as exc:
                flash("", [f"No se pudo leer la hoja BD: {exc}"])
            st.rerun()
    U.titulo("Rutas chárter frecuentes")
    rc = db.df_rutas_charter()
    if rc.empty:
        st.caption("Sin rutas chárter guardadas. Operaciones puede agregarlas desde el archivo plano.")
    else:
        _tabla(rc, hide_index=True, width="stretch")


def _param_tarifas(u):
    p = db.params_dict()
    U.titulo("Tarifas y liquidación", "Deje en blanco lo que aún no esté definido: el sistema no bloquea y pide el dato manual.")
    c1, c2, c3 = st.columns(3)
    val = lambda k: "" if p.get(k) is None else (str(int(p[k])) if isinstance(p[k], float) and p[k].is_integer() else str(p[k]))
    nuevos = {"precio_JET A1": _texto_campo(c1, "Precio JET A1 (COP/gal)", val("precio_JET A1"), "pt_jet"),
              "precio_AVGAS 100LL": _texto_campo(c2, "Precio AVGAS 100LL (COP/gal)", val("precio_AVGAS 100LL"), "pt_avg"),
              "min_smlv": _texto_campo(c3, "Minutos mínimos garantizados SMLV", val("min_smlv"), "pt_smlv"),
              "tarifa_hora_ptl": _texto_campo(c1, "Valor hora piloto PTL (COP)", val("tarifa_hora_ptl"), "pt_ptl"),
              "tarifa_hora_pca": _texto_campo(c2, "Valor hora copiloto PCA (COP)", val("tarifa_hora_pca"), "pt_pca")}
    formula = c3.text_input("Fórmula TOTAL A PAGAR", p.get("liq_formula_total") or "", key="pt_form",
                            help="Variables: " + ", ".join(R.VARS_LIQUIDACION) + ". En blanco = digitación manual.")
    err = R.validar_formula(formula, R.VARS_LIQUIDACION) if R.texto(formula) else None
    if err:
        st.error(err)
    nuevos["liq_formula_total"] = formula
    pulsado, motivo = guardar_con_clausula(u, "pt", "Guardar tarifas y fórmula", extra_ok=not err)
    if pulsado:
        errores = []
        for k, v in nuevos.items():
            ok, msg = db.set_param(k, v, u, motivo)
            if not ok:
                errores.append(msg)
        flash("Parámetros guardados.", errores)
        st.rerun()
    U.titulo("Regla TIEMPO PILOTOS")
    editor_tiempo_pilotos(u, "ger_tp", "GERENCIA")
    U.titulo("Todos los parámetros (avanzado)", "Edite la columna 'valor'. Celda vacía = parámetro pendiente.")
    dfp = db.df_parametros()
    dfp = a_objeto(dfp, list(dfp.columns))
    k = f"par_all_{ss.get('par_all_v', 0)}"
    _editor(dfp, key=k, hide_index=True, width="stretch", disabled=[c for c in dfp.columns if c != "valor"], height=380)
    cambios = (ss.get(k) or {}).get("edited_rows") or {}
    if cambios:
        pend = {dfp.iloc[int(i)]["clave"]: c.get("valor") for i, c in cambios.items() if "valor" in c}
        errores = []
        for clave, v in pend.items():
            if clave in ("liq_formula_total",) and R.texto(v):
                e = R.validar_formula(v, R.VARS_LIQUIDACION)
                errores += [f"{clave}: {e}"] if e else []
            if clave == "tp_formula" and R.texto(v):
                e = R.validar_formula(v, R.VARS_TIEMPO_PILOTOS)
                errores += [f"{clave}: {e}"] if e else []
            if clave == "tp_modo" and R.texto(v).upper() not in ("MANUAL", "FORMULA", "RANGOS"):
                errores.append("tp_modo debe ser MANUAL, FORMULA o RANGOS.")
            if clave == "tp_rangos" and R.texto(v):
                try:
                    json.loads(v)
                except ValueError:
                    errores.append("tp_rangos debe ser JSON válido.")
        for e in errores:
            st.error(e)
        pulsado, motivo = guardar_con_clausula(u, f"{k}_g", "Guardar parámetros", extra_ok=not errores)
        if pulsado:
            errs = []
            for clave, v in pend.items():
                ok, msg = db.set_param(clave, R.texto(v).upper() if clave == "tp_modo" else v, u, motivo)
                if not ok:
                    errs.append(msg)
            ss["par_all_v"] = ss.get("par_all_v", 0) + 1
            flash("Parámetros guardados.", errs)
            st.rerun()


def _param_sistema(u):
    ia_ok, ia_msg = V.ia_disponible()
    tam = db.DB_PATH.stat().st_size / 1024 / 1024 if db.DB_PATH.exists() else 0
    n_ev = sum(1 for _ in db.EVIDENCIAS_DIR.rglob("*.jpg")) if db.EVIDENCIAS_DIR.exists() else 0
    U.kpi_fila([{"label": "Lectura IA", "value": "Activa" if ia_ok else "Manual", "tono": "bien" if ia_ok else "alerta", "nota": ia_msg[:80]},
                {"label": "Base de datos", "value": f"{tam:,.1f} MB"}, {"label": "Evidencias", "value": f"{n_ev:,}"},
                {"label": "Respaldos", "value": f"{len(db.listar_respaldos())}"}])
    st.caption(f"Carpeta de datos: {db.DATA_DIR}")
    st.warning("Para no perder datos, despliegue con un disco persistente (servidor propio, VPS o volumen Docker) y configure MOON_DATA_DIR. "
               "Los alojamientos gratuitos con disco efímero borran la base al reiniciarse: descargue respaldos con frecuencia.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔐 Verificar cadena de la bitácora inmutable", key="ps_cad", width="stretch"):
            r = db.verificar_cadena_auditoria()
            (st.success if r["ok"] else st.error)(r["mensaje"])
        if st.button("💾 Crear respaldo ahora", key="ps_resp", width="stretch"):
            p = db.respaldo_db("manual")
            db.auditar(u, "GERENCIA", None, None, "RESPALDO", nuevo=p.name, motivo="Respaldo manual")
            flash(f"Respaldo creado: {p.name}")
            st.rerun()
    with c2:
        with st.container(border=True):
            st.markdown("**Descargar copia completa de la base**")
            st.caption("Contiene datos sensibles (incluye hashes de contraseñas). Guárdela cifrada.")
            ok = U.clausula_responsabilidad(u, "ps_dl_cl")
            if st.button("Preparar copia para descarga", key="ps_prep", disabled=not ok, width="stretch"):
                ss["ps_respaldo"] = (db.bytes_respaldo_actual(), f"moon_respaldo_{db.ahora():%Y%m%d_%H%M}.db")
                db.auditar(u, "GERENCIA", None, None, "EXPORTAR", nuevo=ss["ps_respaldo"][1], motivo="Descarga de respaldo completo")
            if ok and ss.get("ps_respaldo"):
                st.download_button("⬇️ Descargar respaldo .db", ss["ps_respaldo"][0], ss["ps_respaldo"][1],
                                   "application/octet-stream", key="ps_dl", width="stretch")
    with st.expander("🗂️ Exportar evidencias de un mes (ZIP)"):
        periodo = selector_mes("ps_zip_mes")
        carpeta = db.EVIDENCIAS_DIR / periodo[:4] / f"{periodo[5:]}_{R.MESES_ES[int(periodo[5:])]}"
        archivos = sorted(carpeta.rglob("*.jpg")) if carpeta.exists() else []
        st.caption(f"{len(archivos)} imagen(es) en {carpeta.relative_to(db.DATA_DIR) if carpeta.exists() else 'carpeta inexistente'}.")
        if archivos and st.button("Preparar ZIP", key="ps_zip"):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
                for f in archivos:
                    z.write(f, f.relative_to(db.EVIDENCIAS_DIR))
            db.auditar(u, "GERENCIA", "documentos", None, "EXPORTAR", nuevo=f"{len(archivos)} evidencias {periodo}")
            st.download_button("⬇️ Descargar ZIP", buf.getvalue(), f"Moon_Evidencias_{periodo}.zip", "application/zip", key="ps_zip_dl")
    st.markdown("**Respaldos automáticos**")
    _tabla(pd.DataFrame(db.listar_respaldos()), hide_index=True, width="stretch", height=240)


def tab_bitacora(u):
    st.caption("Registro inmutable: nadie (ni Gerencia) puede editar o eliminar filas; cada evento queda encadenado con SHA-256.")
    desde, hasta = rango_fechas("bit_rg", defecto="7 días")
    c1, c2, c3 = st.columns(3)
    ced = c1.text_input("Cédula", key="bit_ced")
    tabla = c2.selectbox("Tabla", ["", "combustible_tiquetes", "libros_vuelo", "tramos_vuelo", "fichas_vuelo", "operaciones_plano",
                                   "mantenimiento_oma", "usuarios", "parametros", "documentos", "consolidados_diarios",
                                   "consolidados_mensuales", "consolidados_anuales", "liquidacion_ajustes", "aeronaves",
                                   "modelos_aeronave", "catalogo_rutas", "auditorias_ia"], key="bit_tab")
    accion = c3.selectbox("Acción", ["", "CREAR", "EDITAR", "ELIMINAR", "CORRECCION_IA", "SINCRONIZAR", "CIERRE_DIARIO", "CIERRE_MENSUAL",
                                     "CIERRE_ANUAL", "AUDITORIA_IA", "PARAMETRO", "LOGIN", "LOGIN_FALLIDO", "LOGOUT", "RESPALDO", "EXPORTAR"],
                          key="bit_acc")
    df = db.df_auditoria(desde, hasta, R.solo_digitos(ced) or (ced.strip() or None), tabla or None, accion or None)
    _tabla(df, hide_index=True, width="stretch", height=460)
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ Descargar bitácora filtrada", X.libro({"Bitácora": df}, {"Documento": "Bitácora inmutable audit_logs",
                                                                                   "Periodo": f"{desde} a {hasta}", "Generado por": u["nombre"]}),
                       f"Moon_Bitacora_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx", XLSX, key="bit_dl", width="stretch")
    if c2.button("🔐 Verificar integridad de la cadena", key="bit_ver", width="stretch"):
        r = db.verificar_cadena_auditoria()
        (st.success if r["ok"] else st.error)(r["mensaje"])


# ============================================================
# MÓDULO 0 · Norte estratégico y gobernanza
# ============================================================
def modulo_inicio(u):
    U.page_title("🌙 Moon Aerolínea · Norte Estratégico y Gobernanza")
    U.hero(f"Bienvenido(a), {u['nombre'].split(' ')[0]}.", "El cielo nos une: cada vuelo, cada galón y cada pasajero, "
           "con datos confiables, auditables y firmados por su responsable.", f"Rol: {u['rol'].title()}")
    U.tarjetas_norte()
    U.titulo("Gobernanza del dato", "Flujo de captura, validación cruzada, carpetas seriadas, consolidados en cascada y auditoría fiscal.")
    U.diagrama_gobernanza()
    hoy = db.hoy()
    ini = hoy.replace(day=1)
    dv = db.dataset_vuelos(ini, hoy)
    docs_hoy = db.df_documentos(hoy - timedelta(days=0), hoy)
    mant = db.df_mantenimiento(estado="ABIERTO")
    U.titulo("Pulso del mes")
    U.kpi_fila([{"label": "Vuelos del mes", "value": f"{len(dv):,}"}, {"label": "Pasajeros a bordo", "value": U.num(dv["PAX"].sum() if not dv.empty else 0)},
                {"label": "Soportes cargados hoy", "value": f"{len(docs_hoy)}"},
                {"label": "Ítems OMA abiertos", "value": f"{len(mant)}", "tono": "alerta" if len(mant) else "bien"}])
    U.titulo("Flota")
    p = db.params_dict()
    U.tarjetas_flota(db.df_aeronaves(), float(p.get("oma_umbral_amarillo_h") or 50), float(p.get("oma_umbral_rojo_h") or 10))
    accesos = [MODULOS[k][0] for k, (_, roles) in MODULOS.items() if u["rol"] in roles and k not in ("inicio", "cuenta")]
    st.caption("Sus módulos: " + " · ".join(accesos))


def modulo_cuenta(u):
    U.page_title("Mi cuenta")
    c1, c2 = st.columns([1, 1.3], gap="large")
    with c1:
        with st.container(border=True):
            st.markdown(f"**{u['nombre']}**")
            st.caption(f"C.C. {u['cedula']} · Rol {u['rol']}" + (f" · {u['cargo_vuelo']}" if u.get("cargo_vuelo") else ""))
            st.caption(f"Último acceso: {u.get('ultimo_acceso') or '—'} · Último cambio de clave: {u.get('fecha_cambio_clave') or '—'}")
    with c2:
        st.markdown("**Cambiar contraseña**")
        vista_cambio_clave(u, obligatorio=False)


# ============================================================
# Navegación principal
# ============================================================
INICIO_POR_ROL = {"PILOTO": "pilotos", "CONTABILIDAD": "contabilidad", "OPERACIONES": "operaciones", "OMA": "oma", "GERENCIA": "gerencia"}


def barra_lateral(u, forzado=False):
    with st.sidebar:
        U.html(f'<div class="side-logo">{U.logo_img(190)}</div>')
        U.html(f'<div class="side-user"><b>{U.e(u["nombre"])}</b><br>{U.e(u["rol"].title())} · C.C. {U.e(u["cedula"])}</div>')
        st.markdown("---")
        vista = None
        if not forzado:
            permitidos = [k for k, (_, roles) in MODULOS.items() if u["rol"] in roles]
            if ss.get("nav") not in permitidos:
                ss["nav"] = INICIO_POR_ROL.get(u["rol"], "inicio")
            vista = st.radio("Navegación", permitidos, format_func=lambda k: MODULOS[k][0], key="nav", label_visibility="collapsed")
            st.markdown("---")
        ok, _ = V.ia_disponible()
        U.html(f'<div class="side-stat">{"🤖 Lectura IA activa" if ok else "✍️ Modo manual asistido"}</div>'
               f'<div class="side-stat">🗓️ {db.ahora():%d/%m/%Y %H:%M} (hora Colombia)</div>')
        if st.button("Cerrar sesión", key="logout", width="stretch"):
            cerrar_sesion()
            st.rerun()
    return vista


def main():
    u = validar_sesion()
    if not u:
        vista_login()
        return
    if u.get("forzar_cambio_clave"):
        barra_lateral(u, forzado=True)
        vista_cambio_clave(u, obligatorio=True)
        return
    vista = barra_lateral(u)
    mostrar_flash()
    {"inicio": modulo_inicio, "pilotos": modulo_pilotos, "contabilidad": modulo_contabilidad, "operaciones": modulo_operaciones,
     "oma": modulo_oma, "gerencia": modulo_gerencia, "cuenta": modulo_cuenta}[vista](u)


main()

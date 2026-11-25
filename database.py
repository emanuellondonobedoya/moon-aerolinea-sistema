# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Capa de datos (database.py)

SQLite en modo WAL con:
  * Migraciones idempotentes (CREATE TABLE IF NOT EXISTS + PRAGMA table_info para columnas nuevas).
  * Usuarios por cédula y rol con PBKDF2-HMAC-SHA256 (600.000 iteraciones, salt aleatorio) y bloqueo por intentos.
  * Bitácora audit_logs inmutable (triggers que abortan UPDATE/DELETE) + cadena de hashes SHA-256.
  * hash_integridad por registro, borrado lógico (nunca se pierde un dato) y respaldos automáticos.
  * Consolidados en cascada: diario → mensual → anual con instantáneas congeladas y su hash.
  * Importadores: hoja BD (catálogo de rutas) y MES (archivo plano MO-A3-013) de NOVIEMBRE.xlsx, y base v2.
"""
from __future__ import annotations

import base64
import gzip
import hashlib
import hmac
import io
import json
import logging
import os
import re
import secrets
import shutil
import sqlite3
import string
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

import rules_engine as R

log = logging.getLogger("moon.db")

# ============================================================
# Configuración, rutas y hora oficial de Colombia (UTC-5, sin horario de verano)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
TZ_CO = timezone(timedelta(hours=-5), name="America/Bogota")


def secret(nombre: str, default=None):
    """Lee st.secrets y, si no existe, una variable de entorno. Nunca lanza excepción."""
    valor = None
    try:
        import streamlit as st
        valor = st.secrets.get(nombre)
    except Exception:
        valor = None
    return valor if valor not in (None, "") else os.environ.get(nombre, default)


DATA_DIR = Path(secret("MOON_DATA_DIR", str(BASE_DIR / "data"))).expanduser().resolve()
DB_PATH = DATA_DIR / "moon_aerolinea.db"
EVIDENCIAS_DIR = DATA_DIR / "evidencias"
RESPALDOS_DIR = DATA_DIR / "respaldos"
ROLES = ("PILOTO", "CONTABILIDAD", "OPERACIONES", "OMA", "GERENCIA")
AREAS_CIERRE = ("CONTABILIDAD", "OPERACIONES")
SISTEMA = {"cedula": "SISTEMA", "nombre": "Proceso automático", "rol": "SISTEMA"}
PBKDF2_ITER = 600_000


def ahora() -> datetime:
    return datetime.now(TZ_CO).replace(tzinfo=None)


def ahora_txt() -> str:
    return ahora().strftime("%Y-%m-%d %H:%M:%S")


def hoy() -> date:
    return ahora().date()


# ============================================================
# Conexión
# ============================================================
@contextmanager
def conexion(escritura: bool = False):
    """Conexión con claves foráneas, espera ante bloqueo y escritura sincrónica.
    escritura=True abre una transacción BEGIN IMMEDIATE (serializa escritores: sin condiciones de carrera)."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA busy_timeout=30000")
    con.execute("PRAGMA synchronous=FULL")
    try:
        if escritura:
            con.execute("BEGIN IMMEDIATE")
        yield con
        if escritura:
            con.execute("COMMIT")
    except Exception:
        if escritura and con.in_transaction:
            con.execute("ROLLBACK")
        raise
    finally:
        con.close()


def _filas(con, sql, params=()) -> list[dict]:
    cur = con.execute(sql, params)
    cols = [c[0] for c in cur.description] if cur.description else []
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def _fila(con, sql, params=()) -> dict | None:
    r = _filas(con, sql, params)
    return r[0] if r else None


def consulta(sql, params=()) -> list[dict]:
    with conexion() as con:
        return _filas(con, sql, params)


def df_sql(sql, params=()) -> pd.DataFrame:
    with conexion() as con:
        cur = con.execute(sql, params)
        cols = [c[0] for c in cur.description] if cur.description else []
        return pd.DataFrame(cur.fetchall(), columns=cols)


def _json(v) -> str | None:
    if R.es_vacio(v):
        return None
    if isinstance(v, (dict, list, tuple)):
        return json.dumps(v, ensure_ascii=False, default=str)[:20000]
    return str(v)[:4000]


# ============================================================
# Esquema
# ============================================================
_AUDIT_COLS = """
    origen_carga TEXT, creado_por_cedula TEXT, creado_por_nombre TEXT, creado_por_rol TEXT, fecha_creacion TEXT,
    validado_por_cedula TEXT, fecha_validacion TEXT, actualizado_por_cedula TEXT, fecha_actualizacion TEXT,
    hash_integridad TEXT, eliminado INTEGER NOT NULL DEFAULT 0, motivo_eliminacion TEXT"""

SCHEMA = [
    "CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT)",
    """CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT NOT NULL UNIQUE, nombre TEXT NOT NULL,
        rol TEXT NOT NULL CHECK (rol IN ('PILOTO','CONTABILIDAD','OPERACIONES','OMA','GERENCIA')),
        cargo_vuelo TEXT, licencia TEXT, email TEXT, password_hash TEXT NOT NULL,
        forzar_cambio_clave INTEGER NOT NULL DEFAULT 1, activo INTEGER NOT NULL DEFAULT 1,
        intentos_fallidos INTEGER NOT NULL DEFAULT 0, bloqueado_hasta TEXT, ultimo_acceso TEXT,
        fecha_cambio_clave TEXT, creado_por_cedula TEXT, fecha_creacion TEXT NOT NULL)""",
    """CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, fecha_hora TEXT NOT NULL, usuario_cedula TEXT, usuario_nombre TEXT,
        rol TEXT, modulo TEXT, tabla_afectada TEXT, registro_id TEXT, accion TEXT NOT NULL, campo_modificado TEXT,
        valor_anterior TEXT, valor_nuevo TEXT, motivo_cambio TEXT, hash_anterior TEXT, hash_registro TEXT)""",
    """CREATE TRIGGER IF NOT EXISTS audit_logs_inmutable_upd BEFORE UPDATE ON audit_logs
        BEGIN SELECT RAISE(ABORT, 'audit_logs es inmutable: no se permite modificar'); END""",
    """CREATE TRIGGER IF NOT EXISTS audit_logs_inmutable_del BEFORE DELETE ON audit_logs
        BEGIN SELECT RAISE(ABORT, 'audit_logs es inmutable: no se permite eliminar'); END""",
    """CREATE TABLE IF NOT EXISTS parametros (
        clave TEXT PRIMARY KEY, valor TEXT, tipo TEXT, grupo TEXT, descripcion TEXT,
        actualizado_por TEXT, fecha_actualizacion TEXT)""",
    """CREATE TABLE IF NOT EXISTS modelos_aeronave (
        codigo TEXT PRIMARY KEY, nombre TEXT, sillas_ofrecidas INTEGER, carga_ofrecida_kg REAL,
        tipo_trafico_default TEXT, uso_aeronave_default TEXT, tipo_combustible_default TEXT,
        metodo_medicion_default TEXT, reserva_combustible_lbs REAL, factor_conversion_lbs_gal REAL,
        ciclo_inspeccion_horas REAL, activo INTEGER NOT NULL DEFAULT 1, actualizado_por TEXT, fecha_actualizacion TEXT)""",
    """CREATE TABLE IF NOT EXISTS aeronaves (
        matricula TEXT PRIMARY KEY, modelo_codigo TEXT REFERENCES modelos_aeronave(codigo) ON UPDATE CASCADE,
        horas_totales REAL, ciclos_totales INTEGER, proxima_inspeccion_horas REAL, proxima_inspeccion_ciclos INTEGER,
        base TEXT, activo INTEGER NOT NULL DEFAULT 1, notas TEXT, actualizado_por TEXT, fecha_actualizacion TEXT)""",
    """CREATE TABLE IF NOT EXISTS catalogo_rutas (
        id INTEGER PRIMARY KEY AUTOINCREMENT, numero_vuelo TEXT NOT NULL, ruta TEXT, origen TEXT, destino TEXT,
        hora TEXT, tipo_vuelo TEXT, combustible_estandar REAL, activo INTEGER NOT NULL DEFAULT 1, fuente TEXT,
        actualizado_por TEXT, fecha_actualizacion TEXT)""",
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_ruta_vuelo_activo ON catalogo_rutas(numero_vuelo) WHERE activo=1",
    """CREATE TABLE IF NOT EXISTS rutas_charter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, origen TEXT NOT NULL, destino TEXT NOT NULL, modelo_codigo TEXT,
        combustible_galones REAL, activo INTEGER NOT NULL DEFAULT 1, actualizado_por TEXT, fecha_actualizacion TEXT)""",
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_ruta_charter ON rutas_charter(origen, destino, COALESCE(modelo_codigo,'')) WHERE activo=1",
    """CREATE TABLE IF NOT EXISTS documentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT NOT NULL CHECK (tipo IN ('COMB','LV','FV')),
        fecha_doc TEXT, hora_min INTEGER, matricula TEXT, seq_dia INTEGER, es_charter INTEGER DEFAULT 0,
        ruta_archivo TEXT NOT NULL, sha256 TEXT NOT NULL, ancho INTEGER, alto INTEGER, nitidez REAL,
        origen_carga TEXT, creado_por_cedula TEXT, creado_por_nombre TEXT, creado_por_rol TEXT, fecha_creacion TEXT,
        extraccion_ia TEXT, campos_dudosos TEXT, ultima_auditoria_ia TEXT, eliminado INTEGER NOT NULL DEFAULT 0)""",
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_doc_sha ON documentos(sha256) WHERE eliminado=0",
    "CREATE INDEX IF NOT EXISTS ix_doc_fecha ON documentos(fecha_doc, matricula, tipo)",
    f"""CREATE TABLE IF NOT EXISTS combustible_tiquetes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, documento_id INTEGER REFERENCES documentos(id),
        recibo TEXT, numero_boleto TEXT, fecha_consumo TEXT, hora_min INTEGER, estacion TEXT, matricula TEXT,
        producto TEXT, destino TEXT, numero_vuelo TEXT, galones REAL, valor_unit REAL, valor_total REAL,
        piloto_nombre TEXT, piloto_cedula TEXT, observacion TEXT, {_AUDIT_COLS})""",
    "CREATE INDEX IF NOT EXISTS ix_tiq_fecha ON combustible_tiquetes(fecha_consumo, matricula)",
    f"""CREATE TABLE IF NOT EXISTS libros_vuelo (
        id INTEGER PRIMARY KEY AUTOINCREMENT, documento_id INTEGER REFERENCES documentos(id),
        numero_libro TEXT, fecha TEXT, matricula TEXT, piloto_nombre TEXT, piloto_licencia TEXT, piloto_cedula TEXT,
        copiloto_nombre TEXT, copiloto_licencia TEXT, copiloto_cedula TEXT, horas_totales REAL, ciclos_totales INTEGER,
        proxima_inspeccion TEXT, tot_tiempo_vuelo_min INTEGER, tot_tiempo_nocturno_min INTEGER,
        tot_tiempo_instrumentos_min INTEGER, tot_tiempo_total_min INTEGER, tot_lbs_usado REAL,
        descuadre_totales INTEGER NOT NULL DEFAULT 0, parametros_motor TEXT, observaciones TEXT, {_AUDIT_COLS})""",
    "CREATE INDEX IF NOT EXISTS ix_libro_fecha ON libros_vuelo(fecha, matricula)",
    """CREATE TABLE IF NOT EXISTS tramos_vuelo (
        id INTEGER PRIMARY KEY AUTOINCREMENT, libro_id INTEGER NOT NULL REFERENCES libros_vuelo(id), orden INTEGER,
        numero_vuelo TEXT, nro_vuelo_plano TEXT, es_charter INTEGER DEFAULT 0, origen TEXT, destino TEXT,
        galones_tanqueados REAL, recibo_combustible TEXT, lbs_comb_salida REAL, pax_total INTEGER, pax_transito INTEGER,
        hora_salida_min INTEGER, hora_despegue_min INTEGER, hora_aterrizaje_min INTEGER, hora_llegada_min INTEGER,
        tiempo_vuelo_min INTEGER, tiempo_nocturno_min INTEGER, tiempo_instrumentos_min INTEGER, tiempo_total_min INTEGER,
        lbs_comb_llegada REAL, lbs_comb_usado REAL, tiempo_pilotos_min INTEGER, tiempo_pilotos_origen TEXT,
        observacion TEXT, hash_integridad TEXT, actualizado_por_cedula TEXT, fecha_actualizacion TEXT,
        eliminado INTEGER NOT NULL DEFAULT 0, motivo_eliminacion TEXT)""",
    "CREATE INDEX IF NOT EXISTS ix_tramo_libro ON tramos_vuelo(libro_id)",
    f"""CREATE TABLE IF NOT EXISTS fichas_vuelo (
        id INTEGER PRIMARY KEY AUTOINCREMENT, documento_id INTEGER REFERENCES documentos(id),
        fecha TEXT, matricula TEXT, numero_vuelo TEXT, origen TEXT, destino TEXT, hora_min INTEGER,
        pax_total INTEGER, tripadis INTEGER, bebes INTEGER, pax_transito INTEGER, carga_a_bordo_kg REAL,
        carga_transito_kg REAL, combustible_basico_lbs REAL, piloto_nombre TEXT, observaciones TEXT, {_AUDIT_COLS})""",
    "CREATE INDEX IF NOT EXISTS ix_ficha_fecha ON fichas_vuelo(fecha, matricula)",
    f"""CREATE TABLE IF NOT EXISTS mantenimiento_oma (
        id INTEGER PRIMARY KEY AUTOINCREMENT, libro_id INTEGER REFERENCES libros_vuelo(id),
        documento_id INTEGER REFERENCES documentos(id), fecha TEXT, matricula TEXT, item TEXT, reporte TEXT,
        accion_correctiva TEXT, pn_on TEXT, sn_on TEXT, pn_off TEXT, sn_off TEXT, posicion TEXT, rii TEXT,
        tecnico_nombre TEXT, tecnico_licencia TEXT, base TEXT, estado TEXT NOT NULL DEFAULT 'ABIERTO', {_AUDIT_COLS})""",
    f"""CREATE TABLE IF NOT EXISTS operaciones_plano (
        id INTEGER PRIMARY KEY AUTOINCREMENT, tramo_id INTEGER REFERENCES tramos_vuelo(id), ficha_id INTEGER,
        fuente TEXT NOT NULL DEFAULT 'SOPORTES', fecha_local TEXT, hora_local_min INTEGER, es_charter INTEGER DEFAULT 0,
        tip TEXT, sigla TEXT, fecha_mmaaaa TEXT, fecha_utc TEXT, hora_utc TEXT, origen TEXT, destino TEXT, matricula TEXT,
        tipo_equipo TEXT, nro_vuelo TEXT, sillas_ofrecidas INTEGER, carga_ofrecida REAL, pax_bordo INTEGER,
        carga_bordo REAL, correo_bordo REAL, pax_transito INTEGER, carga_transito REAL, horas_bloque TEXT,
        tipo_vuelo TEXT, tipo_trafico TEXT, uso_aeronave TEXT, tipo_combustible TEXT, metodo_medicion TEXT,
        combustible_consumido REAL, horas_vuelo TEXT, carga_real REAL, campos_manuales TEXT, pendientes TEXT,
        notas TEXT, {_AUDIT_COLS})""",
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_plano_tramo ON operaciones_plano(tramo_id) WHERE tramo_id IS NOT NULL AND eliminado=0",
    "CREATE INDEX IF NOT EXISTS ix_plano_fecha ON operaciones_plano(fecha_local, matricula)",
    """CREATE TABLE IF NOT EXISTS liquidacion_ajustes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, periodo TEXT NOT NULL, clave_persona TEXT NOT NULL, campo TEXT NOT NULL,
        valor TEXT, motivo TEXT, usuario_cedula TEXT, fecha TEXT)""",
    """CREATE TABLE IF NOT EXISTS consolidados_diarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT, area TEXT NOT NULL, fecha TEXT NOT NULL, estado TEXT NOT NULL,
        n_registros INTEGER, resumen TEXT, snapshot BLOB, sha256 TEXT, cerrado_por_cedula TEXT, cerrado_por_nombre TEXT,
        fecha_cierre TEXT, reabierto_por_cedula TEXT, fecha_reapertura TEXT, motivo_reapertura TEXT, UNIQUE(area, fecha))""",
    """CREATE TABLE IF NOT EXISTS consolidados_mensuales (
        id INTEGER PRIMARY KEY AUTOINCREMENT, area TEXT NOT NULL, periodo TEXT NOT NULL, estado TEXT NOT NULL,
        n_registros INTEGER, resumen TEXT, snapshot BLOB, sha256 TEXT, cerrado_por_cedula TEXT, cerrado_por_nombre TEXT,
        fecha_cierre TEXT, reabierto_por_cedula TEXT, fecha_reapertura TEXT, motivo_reapertura TEXT, UNIQUE(area, periodo))""",
    """CREATE TABLE IF NOT EXISTS consolidados_anuales (
        id INTEGER PRIMARY KEY AUTOINCREMENT, area TEXT NOT NULL, anio TEXT NOT NULL, estado TEXT NOT NULL,
        meses_incluidos TEXT, n_registros INTEGER, resumen TEXT, snapshot BLOB, sha256 TEXT,
        actualizado_por_cedula TEXT, fecha_actualizacion TEXT, cerrado_por_cedula TEXT, fecha_cierre TEXT, UNIQUE(area, anio))""",
    """CREATE TABLE IF NOT EXISTS auditorias_ia (
        id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT NOT NULL, ambito TEXT NOT NULL, desde TEXT, hasta TEXT,
        ejecutada_por_cedula TEXT, ejecutada_por_nombre TEXT, n_hallazgos INTEGER, resumen TEXT, uso_ia INTEGER DEFAULT 0)""",
    """CREATE TABLE IF NOT EXISTS hallazgos (
        id INTEGER PRIMARY KEY AUTOINCREMENT, auditoria_id INTEGER NOT NULL REFERENCES auditorias_ia(id),
        severidad TEXT, area TEXT, categoria TEXT, tabla TEXT, registro_id TEXT, fecha_ref TEXT, matricula TEXT,
        detalle TEXT, accion_sugerida TEXT, responsable_cedula TEXT, responsable_nombre TEXT,
        ultimo_editor_cedula TEXT, ultimo_editor_nombre TEXT)""",
    "CREATE INDEX IF NOT EXISTS ix_hallazgo_aud ON hallazgos(auditoria_id)",
    "CREATE INDEX IF NOT EXISTS ix_audit_registro ON audit_logs(tabla_afectada, registro_id)",
    "CREATE INDEX IF NOT EXISTS ix_audit_usuario ON audit_logs(usuario_cedula, fecha_hora)",
]

# Columnas esperadas por tabla (para migrar bases antiguas sin perder datos)
_COLUMNAS_ESPERADAS = {
    "usuarios": {"cargo_vuelo": "TEXT", "licencia": "TEXT", "email": "TEXT", "fecha_cambio_clave": "TEXT"},
    "documentos": {"ultima_auditoria_ia": "TEXT", "campos_dudosos": "TEXT", "creado_por_nombre": "TEXT"},
    "combustible_tiquetes": {"numero_boleto": "TEXT", "destino": "TEXT", "numero_vuelo": "TEXT", "hora_min": "INTEGER"},
    "tramos_vuelo": {"tiempo_pilotos_origen": "TEXT", "nro_vuelo_plano": "TEXT", "observacion": "TEXT"},
    "libros_vuelo": {"descuadre_totales": "INTEGER NOT NULL DEFAULT 0", "parametros_motor": "TEXT"},
    "operaciones_plano": {"notas": "TEXT", "es_charter": "INTEGER DEFAULT 0", "ficha_id": "INTEGER"},
    "aeronaves": {"proxima_inspeccion_ciclos": "INTEGER", "base": "TEXT"},
    "modelos_aeronave": {"ciclo_inspeccion_horas": "REAL"},
}

# Campos editables por tabla y su tipo de normalización (lista blanca: nada fuera de aquí se puede tocar)
CAMPOS = {
    "combustible_tiquetes": {
        "recibo": "texto", "numero_boleto": "texto", "fecha_consumo": "fecha", "hora_min": "hora", "estacion": "iata",
        "matricula": "matricula", "producto": "producto", "destino": "iata", "numero_vuelo": "digitos",
        "galones": "real", "valor_unit": "real", "valor_total": "real", "piloto_nombre": "texto",
        "piloto_cedula": "digitos", "observacion": "texto"},
    "libros_vuelo": {
        "numero_libro": "digitos", "fecha": "fecha", "matricula": "matricula", "piloto_nombre": "texto",
        "piloto_licencia": "texto", "piloto_cedula": "digitos", "copiloto_nombre": "texto", "copiloto_licencia": "texto",
        "copiloto_cedula": "digitos", "horas_totales": "real", "ciclos_totales": "entero", "proxima_inspeccion": "texto",
        "tot_tiempo_vuelo_min": "duracion", "tot_tiempo_nocturno_min": "duracion", "tot_tiempo_instrumentos_min": "duracion",
        "tot_tiempo_total_min": "duracion", "tot_lbs_usado": "real", "observaciones": "texto"},
    "tramos_vuelo": {
        "numero_vuelo": "digitos", "origen": "iata", "destino": "iata", "galones_tanqueados": "real",
        "recibo_combustible": "texto", "lbs_comb_salida": "real", "pax_total": "entero", "pax_transito": "entero",
        "hora_salida_min": "hora", "hora_despegue_min": "hora", "hora_aterrizaje_min": "hora", "hora_llegada_min": "hora",
        "tiempo_vuelo_min": "duracion", "tiempo_nocturno_min": "duracion", "tiempo_instrumentos_min": "duracion",
        "tiempo_total_min": "duracion", "lbs_comb_llegada": "real", "lbs_comb_usado": "real",
        "tiempo_pilotos_min": "duracion", "observacion": "texto"},
    "fichas_vuelo": {
        "fecha": "fecha", "matricula": "matricula", "numero_vuelo": "digitos", "origen": "iata", "destino": "iata",
        "hora_min": "hora", "pax_total": "entero", "tripadis": "entero", "bebes": "entero", "pax_transito": "entero",
        "carga_a_bordo_kg": "real", "carga_transito_kg": "real", "combustible_basico_lbs": "real",
        "piloto_nombre": "texto", "observaciones": "texto"},
    "mantenimiento_oma": {
        "fecha": "fecha", "matricula": "matricula", "item": "texto", "reporte": "texto", "accion_correctiva": "texto",
        "pn_on": "mayus", "sn_on": "mayus", "pn_off": "mayus", "sn_off": "mayus", "posicion": "mayus",
        "rii": "enum:SI|NO", "tecnico_nombre": "texto", "tecnico_licencia": "texto", "base": "iata",
        "estado": "enum:ABIERTO|CERRADO"},
    "operaciones_plano": {
        "tip": "mayus", "sigla": "mayus", "fecha_mmaaaa": "mmaaaa", "fecha_utc": "plano_fecha", "hora_utc": "plano_hora",
        "origen": "iata", "destino": "iata", "matricula": "matricula_compacta", "tipo_equipo": "mayus",
        "nro_vuelo": "digitos", "sillas_ofrecidas": "entero", "carga_ofrecida": "real", "pax_bordo": "entero",
        "carga_bordo": "real", "correo_bordo": "real", "pax_transito": "entero", "carga_transito": "real",
        "horas_bloque": "plano_hhmm", "tipo_vuelo": "enum:R|A|C", "tipo_trafico": "mayus", "uso_aeronave": "mayus",
        "tipo_combustible": "mayus", "metodo_medicion": "mayus", "combustible_consumido": "real",
        "horas_vuelo": "plano_hhmm", "carga_real": "real"},
}
AREA_TABLA = {"combustible_tiquetes": "CONTABILIDAD", "libros_vuelo": "CONTABILIDAD", "tramos_vuelo": "CONTABILIDAD",
              "fichas_vuelo": "OPERACIONES", "operaciones_plano": "OPERACIONES", "mantenimiento_oma": None}
FECHA_TABLA = {"combustible_tiquetes": "fecha_consumo", "libros_vuelo": "fecha", "fichas_vuelo": "fecha",
               "operaciones_plano": "fecha_local", "mantenimiento_oma": "fecha"}
ETIQUETA_TABLA = {"combustible_tiquetes": "Tiquete de combustible", "libros_vuelo": "Libro de vuelo",
                  "tramos_vuelo": "Tramo de vuelo", "fichas_vuelo": "Ficha de vuelo", "mantenimiento_oma": "Ítem OMA",
                  "operaciones_plano": "Fila archivo plano", "usuarios": "Usuario", "parametros": "Parámetro"}

# ============================================================
# Parámetros por defecto (solo los definidos por Moon; lo no definido queda en blanco)
# ============================================================
PARAMETROS_DEFECTO = [
    # clave, valor, tipo, grupo, descripción
    ("precio_JET A1", None, "real", "TARIFAS", "Precio vigente por galón JET A1 (COP). En blanco = se digita manualmente."),
    ("precio_AVGAS 100LL", None, "real", "TARIFAS", "Precio vigente por galón AVGAS 100LL (COP)."),
    ("tarifa_hora_ptl", "250000", "real", "LIQUIDACION", "Valor hora piloto al mando (PTL) en COP."),
    ("tarifa_hora_pca", "110000", "real", "LIQUIDACION", "Valor hora copiloto (PCA) en COP."),
    ("min_smlv", "1800", "real", "LIQUIDACION", "Minutos mínimos garantizados cubiertos por el SMLV."),
    ("liq_formula_total", R.FORMULA_LIQUIDACION_DEFECTO, "formula", "LIQUIDACION",
     "Fórmula TOTAL A PAGAR. En blanco = digitación manual."),
    ("tp_modo", "MANUAL", "texto", "TIEMPO_PILOTOS", "Modo de TIEMPO PILOTOS: MANUAL, FORMULA o RANGOS."),
    ("tp_formula", "tiempo_vuelo", "formula", "TIEMPO_PILOTOS", "Fórmula TIEMPO PILOTOS (minutos) para modo FORMULA."),
    ("tp_rangos", "[]", "json", "TIEMPO_PILOTOS", "Tabla de rangos [{desde, hasta, resultado}] para modo RANGOS."),
    ("tp_rangos_base", "tiempo_vuelo", "texto", "TIEMPO_PILOTOS", "Variable que se compara con los rangos."),
    ("tp_ajuste_min", "0", "real", "TIEMPO_PILOTOS", "Ajuste en minutos disponible como variable ajuste_min."),
    ("plano_tip", "TE", "texto", "PLANO", "Columna A · TIP."),
    ("plano_sigla", "3CA", "texto", "PLANO", "Columna C · SIGLA de la empresa."),
    ("offset_utc_horas", "5", "real", "PLANO", "Horas a sumar a la hora Colombia para obtener UTC."),
    ("correo_default", "0", "real", "PLANO", "Columna P · CORREO A BORDO (siempre 0)."),
    ("carga_transito_default", "0", "real", "PLANO", "Columna R por defecto cuando la ficha no reporta carga en tránsito."),
    ("combustible_decimales", "0", "entero", "PLANO", "Decimales del combustible consumido calculado."),
    ("tolerancia_galones_pct", "2", "real", "CONTROL", "Tolerancia % entre tiquete y libro de vuelo."),
    ("tolerancia_galones_abs", "2", "real", "CONTROL", "Tolerancia absoluta en galones entre tiquete y libro."),
    ("dias_oportunidad", "1", "entero", "CONTROL", "Días máximos entre el vuelo y su registro para considerarlo a tiempo."),
    ("ia_min_lado_px", "800", "entero", "VISION", "Resolución mínima (lado mayor) de las fotos."),
    ("ia_umbral_nitidez", "45", "real", "VISION", "Varianza Laplaciana mínima (nitidez) de las fotos."),
    ("ia_max_imagenes_auditoria", "12", "entero", "VISION", "Imágenes máximas a re-leer por auditoría integral."),
    ("sesion_minutos", "30", "entero", "SEGURIDAD", "Minutos de inactividad antes de cerrar la sesión."),
    ("max_intentos_login", "5", "entero", "SEGURIDAD", "Intentos fallidos antes del bloqueo temporal."),
    ("bloqueo_minutos", "15", "entero", "SEGURIDAD", "Minutos de bloqueo tras superar los intentos."),
    ("respaldo_retencion", "60", "entero", "SEGURIDAD", "Número de respaldos automáticos a conservar."),
    ("oma_umbral_amarillo_h", "50", "real", "OMA", "Horas restantes para semáforo amarillo."),
    ("oma_umbral_rojo_h", "10", "real", "OMA", "Horas restantes para semáforo rojo."),
]

MODELOS_DEFECTO = [
    # JS32 según operación actual (video de despacho): 19 sillas, 480 kg, D, M, JETA1, M3, reserva 50 lbs.
    {"codigo": "JS32", "nombre": "Jetstream 32", "sillas_ofrecidas": 19, "carga_ofrecida_kg": 480,
     "tipo_trafico_default": "D", "uso_aeronave_default": "M", "tipo_combustible_default": "JETA1",
     "metodo_medicion_default": "M3", "reserva_combustible_lbs": 50, "factor_conversion_lbs_gal": 6.7,
     "ciclo_inspeccion_horas": None},
    # JS41 anunciado sin configuración conocida: todo en blanco y editable.
    {"codigo": "JS41", "nombre": "Jetstream 41", "sillas_ofrecidas": None, "carga_ofrecida_kg": None,
     "tipo_trafico_default": None, "uso_aeronave_default": None, "tipo_combustible_default": None,
     "metodo_medicion_default": None, "reserva_combustible_lbs": None, "factor_conversion_lbs_gal": None,
     "ciclo_inspeccion_horas": None},
]


# ============================================================
# Inicialización y migraciones
# ============================================================
def init_db() -> dict:
    """Idempotente: crea carpetas, tablas, índices y triggers; migra columnas; siembra parámetros, modelos,
    usuario de Gerencia; importa NOVIEMBRE.xlsx (BD) y la base v2 si existen; hace el respaldo diario."""
    for d in (DATA_DIR, EVIDENCIAS_DIR, RESPALDOS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    info = {"catalogo_importado": 0, "aeronaves_importadas": 0, "legacy_importados": 0, "gerencia_creada": False}
    con = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    try:
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA foreign_keys=ON")
        con.execute("BEGIN IMMEDIATE")
        for stmt in SCHEMA:
            con.execute(stmt)
        for tabla, cols in _COLUMNAS_ESPERADAS.items():
            existentes = {r[1] for r in con.execute(f"PRAGMA table_info({tabla})").fetchall()}
            for col, tipo in cols.items():
                if col not in existentes:
                    con.execute(f"ALTER TABLE {tabla} ADD COLUMN {col} {tipo}")
        for clave, valor, tipo, grupo, desc in PARAMETROS_DEFECTO:
            con.execute("INSERT OR IGNORE INTO parametros (clave, valor, tipo, grupo, descripcion, actualizado_por, "
                        "fecha_actualizacion) VALUES (?,?,?,?,?,?,?)", (clave, valor, tipo, grupo, desc, "SISTEMA", ahora_txt()))
        for m in MODELOS_DEFECTO:
            cols = list(m)
            con.execute(f"INSERT OR IGNORE INTO modelos_aeronave ({','.join(cols)}, actualizado_por, fecha_actualizacion) "
                        f"VALUES ({','.join('?' * len(cols))},?,?)", [m[c] for c in cols] + ["SISTEMA", ahora_txt()])
        con.execute("PRAGMA user_version=3")
        con.execute("COMMIT")
    except Exception:
        if con.in_transaction:
            con.execute("ROLLBACK")
        raise
    finally:
        con.close()

    info["gerencia_creada"] = _sembrar_gerencia()
    if not consulta("SELECT 1 FROM catalogo_rutas LIMIT 1"):
        ruta = buscar_plantilla_xlsx()
        if ruta:
            try:
                info["catalogo_importado"] = importar_catalogo_bd(ruta, SISTEMA)
                info["aeronaves_importadas"] = importar_flota_desde_mes(ruta, SISTEMA)
            except Exception:
                log.exception("No se pudo importar la hoja BD de %s", ruta)
    try:
        info["legacy_importados"] = importar_legacy_v2()
    except Exception:
        log.exception("No se pudo importar la base v2")
    try:
        respaldo_diario_si_corresponde()
    except Exception:
        log.exception("No se pudo realizar el respaldo diario")
    return info


def buscar_plantilla_xlsx() -> Path | None:
    candidatos = [secret("MOON_PLANTILLA_XLSX"), DATA_DIR / "NOVIEMBRE.xlsx", BASE_DIR / "NOVIEMBRE.xlsx"]
    for c in candidatos:
        if c and Path(c).is_file():
            return Path(c)
    return None


def get_meta(clave, default=None):
    r = consulta("SELECT valor FROM meta WHERE clave=?", (clave,))
    return r[0]["valor"] if r else default


def set_meta(clave, valor):
    with conexion(True) as con:
        con.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (clave, str(valor)))


# ============================================================
# Bitácora inmutable con cadena de hashes
# ============================================================
def registrar_auditoria(con, usuario: dict, modulo: str, tabla: str | None, registro_id, accion: str,
                        campo: str | None = None, anterior=None, nuevo=None, motivo: str | None = None) -> None:
    """Inserta una fila en audit_logs. Debe llamarse dentro de una transacción de escritura (con)."""
    usuario = usuario or SISTEMA
    prev = con.execute("SELECT hash_registro FROM audit_logs ORDER BY id DESC LIMIT 1").fetchone()
    hash_prev = prev[0] if prev and prev[0] else "GENESIS"
    fila = {"fecha_hora": ahora_txt(), "usuario_cedula": str(usuario.get("cedula", "")),
            "usuario_nombre": str(usuario.get("nombre", "")), "rol": str(usuario.get("rol", "")),
            "modulo": modulo, "tabla_afectada": tabla, "registro_id": None if registro_id is None else str(registro_id),
            "accion": accion, "campo_modificado": campo, "valor_anterior": _json(anterior),
            "valor_nuevo": _json(nuevo), "motivo_cambio": (motivo or None)}
    h = hashlib.sha256((hash_prev + json.dumps(fila, ensure_ascii=False, sort_keys=True)).encode("utf-8")).hexdigest()
    cols = list(fila) + ["hash_anterior", "hash_registro"]
    con.execute(f"INSERT INTO audit_logs ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                list(fila.values()) + [hash_prev, h])


def auditar(usuario, modulo, tabla, registro_id, accion, campo=None, anterior=None, nuevo=None, motivo=None):
    with conexion(True) as con:
        registrar_auditoria(con, usuario, modulo, tabla, registro_id, accion, campo, anterior, nuevo, motivo)


def verificar_cadena_auditoria() -> dict:
    """Recalcula la cadena SHA-256 completa. Detecta filas alteradas o borradas fuera del sistema."""
    filas = consulta("SELECT * FROM audit_logs ORDER BY id")
    prev, ok, ids_prev = "GENESIS", True, None
    for f in filas:
        datos = {k: f[k] for k in ("fecha_hora", "usuario_cedula", "usuario_nombre", "rol", "modulo", "tabla_afectada",
                                   "registro_id", "accion", "campo_modificado", "valor_anterior", "valor_nuevo",
                                   "motivo_cambio")}
        esperado = hashlib.sha256((prev + json.dumps(datos, ensure_ascii=False, sort_keys=True)).encode("utf-8")).hexdigest()
        if f["hash_anterior"] != prev or f["hash_registro"] != esperado:
            return {"ok": False, "total": len(filas), "primer_error_id": f["id"],
                    "mensaje": f"La cadena se rompe en el registro #{f['id']} (posterior a #{ids_prev})."}
        prev, ids_prev = f["hash_registro"], f["id"]
    return {"ok": ok, "total": len(filas), "primer_error_id": None,
            "mensaje": f"Cadena íntegra: {len(filas):,} eventos verificados."}


def df_auditoria(desde=None, hasta=None, cedula=None, tabla=None, accion=None, registro_id=None, limite=20000) -> pd.DataFrame:
    q, p = "SELECT * FROM audit_logs WHERE 1=1", []
    if desde:
        q, p = q + " AND fecha_hora >= ?", p + [f"{desde} 00:00:00"]
    if hasta:
        q, p = q + " AND fecha_hora <= ?", p + [f"{hasta} 23:59:59"]
    if cedula:
        q, p = q + " AND usuario_cedula = ?", p + [str(cedula)]
    if tabla:
        q, p = q + " AND tabla_afectada = ?", p + [tabla]
    if accion:
        q, p = q + " AND accion = ?", p + [accion]
    if registro_id is not None:
        q, p = q + " AND registro_id = ?", p + [str(registro_id)]
    return df_sql(q + f" ORDER BY id DESC LIMIT {int(limite)}", tuple(p))


# ============================================================
# Parámetros
# ============================================================
def _tipar(valor, tipo):
    if R.es_vacio(valor):
        return None
    if tipo == "real":
        return R.a_float(valor)
    if tipo == "entero":
        return R.a_int(valor)
    if tipo == "json":
        try:
            return json.loads(valor)
        except (TypeError, ValueError):
            return None
    return str(valor)


def get_param(clave, default=None):
    r = consulta("SELECT valor, tipo FROM parametros WHERE clave=?", (clave,))
    if not r:
        return default
    v = _tipar(r[0]["valor"], r[0]["tipo"])
    return default if v is None and default is not None else v


def params_dict() -> dict:
    return {r["clave"]: _tipar(r["valor"], r["tipo"]) for r in consulta("SELECT clave, valor, tipo FROM parametros")}


def df_parametros() -> pd.DataFrame:
    return df_sql("SELECT clave, valor, tipo, grupo, descripcion, actualizado_por, fecha_actualizacion FROM parametros ORDER BY grupo, clave")


def set_param(clave, valor, usuario, motivo, modulo="GERENCIA", tipo=None, grupo=None, descripcion=None) -> tuple[bool, str]:
    actual = consulta("SELECT valor, tipo FROM parametros WHERE clave=?", (clave,))
    tipo = tipo or (actual[0]["tipo"] if actual else "texto")
    if tipo == "json" and not R.es_vacio(valor) and not isinstance(valor, str):
        valor = json.dumps(valor, ensure_ascii=False)
    nuevo = None if R.es_vacio(valor) else str(valor).strip()
    if tipo in ("real", "entero") and nuevo is not None and R.a_float(nuevo) is None:
        return False, f"El valor de {clave} debe ser numérico."
    anterior = actual[0]["valor"] if actual else None
    if anterior == nuevo:
        return True, "Sin cambios."
    with conexion(True) as con:
        if actual:
            con.execute("UPDATE parametros SET valor=?, actualizado_por=?, fecha_actualizacion=? WHERE clave=?",
                        (nuevo, usuario.get("cedula"), ahora_txt(), clave))
        else:
            con.execute("INSERT INTO parametros (clave, valor, tipo, grupo, descripcion, actualizado_por, fecha_actualizacion) "
                        "VALUES (?,?,?,?,?,?,?)", (clave, nuevo, tipo, grupo or "OTROS", descripcion, usuario.get("cedula"), ahora_txt()))
        registrar_auditoria(con, usuario, modulo, "parametros", clave, "PARAMETRO", clave, anterior, nuevo, motivo)
    return True, "Parámetro actualizado."


def cfg_tiempo_pilotos(p: dict | None = None) -> dict:
    p = p or params_dict()
    return {"modo": p.get("tp_modo") or "MANUAL", "formula": p.get("tp_formula"), "rangos": p.get("tp_rangos") or [],
            "base": p.get("tp_rangos_base") or "tiempo_vuelo", "ajuste_min": p.get("tp_ajuste_min") or 0}


# ============================================================
# Seguridad: contraseñas y usuarios
# ============================================================
def hash_password(clave: str) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), salt, PBKDF2_ITER)
    return f"pbkdf2_sha256${PBKDF2_ITER}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"


def verificar_password(clave: str, almacenado: str) -> bool:
    try:
        algoritmo, it, salt_b64, hash_b64 = almacenado.split("$")
        if algoritmo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", (clave or "").encode("utf-8"), base64.b64decode(salt_b64), int(it))
        return hmac.compare_digest(dk, base64.b64decode(hash_b64))
    except (ValueError, TypeError, AttributeError):
        return False


_HASH_SEÑUELO = None


def _hash_señuelo():
    global _HASH_SEÑUELO
    if _HASH_SEÑUELO is None:
        _HASH_SEÑUELO = hash_password(secrets.token_urlsafe(16))
    return _HASH_SEÑUELO


_CLAVES_COMUNES = {"moon2026*", "password123", "contraseña1", "123456789a", "qwerty12345", "moonflights1", "aerolinea2026"}


def politica_clave(clave: str, cedula: str = "", nombre: str = "") -> list[str]:
    errores = []
    if len(clave or "") < 10:
        errores.append("Mínimo 10 caracteres.")
    if not re.search(r"[A-Z]", clave or ""):
        errores.append("Al menos una letra mayúscula.")
    if not re.search(r"[a-z]", clave or ""):
        errores.append("Al menos una letra minúscula.")
    if not re.search(r"\d", clave or ""):
        errores.append("Al menos un número.")
    if not re.search(r"[^A-Za-z0-9]", clave or ""):
        errores.append("Al menos un símbolo (por ejemplo * # $ ! ?).")
    if cedula and str(cedula) in (clave or ""):
        errores.append("No puede contener su número de cédula.")
    if (clave or "").lower() in _CLAVES_COMUNES:
        errores.append("Es una contraseña demasiado común.")
    primer_nombre = R.nombre_clave(nombre).split(" ")[0].lower() if nombre else ""
    if len(primer_nombre) >= 4 and primer_nombre in R.sin_tildes(clave or "").lower():
        errores.append("No puede contener su nombre.")
    return errores


def generar_clave_temporal(n=12) -> str:
    alfabeto = string.ascii_letters + string.digits + "*#$!?"
    while True:
        c = "".join(secrets.choice(alfabeto) for _ in range(n))
        if not politica_clave(c):
            return c


def _sembrar_gerencia() -> bool:
    if consulta("SELECT 1 FROM usuarios LIMIT 1"):
        return False
    cedula = R.solo_digitos(secret("GERENCIA_CEDULA", "1000000000")) or "1000000000"
    clave = str(secret("GERENCIA_PASSWORD", "Moon2026*"))
    with conexion(True) as con:
        con.execute("INSERT INTO usuarios (cedula, nombre, rol, password_hash, forzar_cambio_clave, activo, "
                    "creado_por_cedula, fecha_creacion) VALUES (?,?,?,?,1,1,?,?)",
                    (cedula, "Gerencia General", "GERENCIA", hash_password(clave), "SISTEMA", ahora_txt()))
        registrar_auditoria(con, SISTEMA, "SEGURIDAD", "usuarios", cedula, "CREAR", None, None,
                            {"cedula": cedula, "rol": "GERENCIA"}, "Usuario semilla de Gerencia (debe cambiar la clave)")
    return True


def _usuario_publico(f: dict | None) -> dict | None:
    if not f:
        return None
    return {k: f.get(k) for k in ("id", "cedula", "nombre", "rol", "cargo_vuelo", "licencia", "email", "activo",
                                  "forzar_cambio_clave", "ultimo_acceso", "bloqueado_hasta", "intentos_fallidos",
                                  "fecha_creacion", "fecha_cambio_clave")}


def obtener_usuario(cedula) -> dict | None:
    return _usuario_publico(_fila_unica("SELECT * FROM usuarios WHERE cedula=?", (R.solo_digitos(cedula),)))


def _fila_unica(sql, params=()):
    r = consulta(sql, params)
    return r[0] if r else None


def autenticar(cedula, clave) -> tuple[dict | None, str]:
    ced = R.solo_digitos(cedula)
    generico = "Cédula o contraseña incorrectas."
    if not ced or not clave or len(str(clave)) > 256:
        return None, generico
    p = params_dict()
    max_int, bloqueo = int(p.get("max_intentos_login") or 5), int(p.get("bloqueo_minutos") or 15)
    u = _fila_unica("SELECT * FROM usuarios WHERE cedula=?", (ced,))
    if not u:
        verificar_password(str(clave), _hash_señuelo())  # tiempo constante: no revela si la cédula existe
        auditar({"cedula": ced, "nombre": "", "rol": ""}, "SEGURIDAD", "usuarios", ced, "LOGIN_FALLIDO", motivo="Cédula no registrada")
        return None, generico
    if u["bloqueado_hasta"] and u["bloqueado_hasta"] > ahora_txt():
        minutos = max(1, int((datetime.strptime(u["bloqueado_hasta"], "%Y-%m-%d %H:%M:%S") - ahora()).total_seconds() // 60) + 1)
        return None, f"Cuenta bloqueada temporalmente por intentos fallidos. Intente de nuevo en {minutos} minuto(s)."
    if not verificar_password(str(clave), u["password_hash"]):
        intentos = int(u["intentos_fallidos"] or 0) + 1
        hasta = (ahora() + timedelta(minutes=bloqueo)).strftime("%Y-%m-%d %H:%M:%S") if intentos >= max_int else None
        with conexion(True) as con:
            con.execute("UPDATE usuarios SET intentos_fallidos=?, bloqueado_hasta=? WHERE cedula=?",
                        (0 if hasta else intentos, hasta, ced))
            registrar_auditoria(con, _usuario_publico(u), "SEGURIDAD", "usuarios", ced, "LOGIN_FALLIDO",
                                motivo=f"Intento {intentos}" + (" · cuenta bloqueada" if hasta else ""))
        if hasta:
            return None, f"Demasiados intentos. La cuenta quedó bloqueada {bloqueo} minutos."
        return None, generico
    if not u["activo"]:
        auditar(_usuario_publico(u), "SEGURIDAD", "usuarios", ced, "LOGIN_FALLIDO", motivo="Cuenta inactiva")
        return None, "La cuenta está inactiva. Comuníquese con Gerencia."
    with conexion(True) as con:
        con.execute("UPDATE usuarios SET intentos_fallidos=0, bloqueado_hasta=NULL, ultimo_acceso=? WHERE cedula=?",
                    (ahora_txt(), ced))
        registrar_auditoria(con, _usuario_publico(u), "SEGURIDAD", "usuarios", ced, "LOGIN")
    # Re-hash transparente si las iteraciones cambiaron
    if not u["password_hash"].startswith(f"pbkdf2_sha256${PBKDF2_ITER}$"):
        with conexion(True) as con:
            con.execute("UPDATE usuarios SET password_hash=? WHERE cedula=?", (hash_password(str(clave)), ced))
    return obtener_usuario(ced), "OK"


def crear_usuario(cedula, nombre, rol, usuario, cargo_vuelo=None, licencia=None, email=None) -> tuple[bool, str, str | None]:
    ced = R.solo_digitos(cedula)
    if not (5 <= len(ced) <= 12):
        return False, "La cédula debe tener entre 5 y 12 dígitos.", None
    if rol not in ROLES:
        return False, "Rol no válido.", None
    if not R.texto(nombre):
        return False, "Escriba el nombre completo.", None
    if consulta("SELECT 1 FROM usuarios WHERE cedula=?", (ced,)):
        return False, "Ya existe un usuario con esa cédula.", None
    temporal = generar_clave_temporal()
    with conexion(True) as con:
        con.execute("INSERT INTO usuarios (cedula, nombre, rol, cargo_vuelo, licencia, email, password_hash, "
                    "forzar_cambio_clave, activo, creado_por_cedula, fecha_creacion) VALUES (?,?,?,?,?,?,?,1,1,?,?)",
                    (ced, R.texto(nombre)[:120], rol, (cargo_vuelo or None), R.texto(licencia) or None,
                     R.texto(email) or None, hash_password(temporal), usuario.get("cedula"), ahora_txt()))
        registrar_auditoria(con, usuario, "GERENCIA", "usuarios", ced, "CREAR", None, None,
                            {"cedula": ced, "nombre": nombre, "rol": rol, "cargo_vuelo": cargo_vuelo}, "Alta de usuario")
    return True, "Usuario creado. Entregue la clave temporal por un canal seguro.", temporal


def actualizar_usuario(cedula, cambios: dict, usuario, motivo) -> tuple[bool, str]:
    ced = R.solo_digitos(cedula)
    actual = _fila_unica("SELECT * FROM usuarios WHERE cedula=?", (ced,))
    if not actual:
        return False, "Usuario no encontrado."
    permitidos = {"nombre", "rol", "cargo_vuelo", "licencia", "email", "activo"}
    cambios = {k: v for k, v in cambios.items() if k in permitidos and v != actual.get(k)}
    if not cambios:
        return True, "Sin cambios."
    if "rol" in cambios and cambios["rol"] not in ROLES:
        return False, "Rol no válido."
    quita_gerencia = (actual["rol"] == "GERENCIA" and (cambios.get("rol", "GERENCIA") != "GERENCIA" or cambios.get("activo", 1) == 0))
    if quita_gerencia:
        n = consulta("SELECT COUNT(*) n FROM usuarios WHERE rol='GERENCIA' AND activo=1 AND cedula<>?", (ced,))[0]["n"]
        if n == 0:
            return False, "Debe existir al menos otro usuario de Gerencia activo."
    if ced == R.solo_digitos(usuario.get("cedula")) and (cambios.get("activo") == 0 or "rol" in cambios):
        return False, "No puede desactivarse ni cambiarse el rol a sí mismo."
    with conexion(True) as con:
        for k, v in cambios.items():
            con.execute(f"UPDATE usuarios SET {k}=? WHERE cedula=?", (v, ced))
            registrar_auditoria(con, usuario, "GERENCIA", "usuarios", ced, "EDITAR", k, actual.get(k), v, motivo)
    return True, "Usuario actualizado."


def resetear_clave(cedula, usuario, motivo) -> tuple[bool, str, str | None]:
    ced = R.solo_digitos(cedula)
    if not consulta("SELECT 1 FROM usuarios WHERE cedula=?", (ced,)):
        return False, "Usuario no encontrado.", None
    temporal = generar_clave_temporal()
    with conexion(True) as con:
        con.execute("UPDATE usuarios SET password_hash=?, forzar_cambio_clave=1, intentos_fallidos=0, bloqueado_hasta=NULL "
                    "WHERE cedula=?", (hash_password(temporal), ced))
        registrar_auditoria(con, usuario, "GERENCIA", "usuarios", ced, "EDITAR", "password", "***", "*** (temporal)", motivo)
    return True, "Clave temporal generada.", temporal


def cambiar_clave(cedula, actual, nueva) -> tuple[bool, list[str]]:
    ced = R.solo_digitos(cedula)
    u = _fila_unica("SELECT * FROM usuarios WHERE cedula=?", (ced,))
    if not u or not verificar_password(str(actual or ""), u["password_hash"]):
        return False, ["La contraseña actual no es correcta."]
    if actual == nueva:
        return False, ["La nueva contraseña debe ser diferente de la actual."]
    errores = politica_clave(nueva, ced, u["nombre"])
    if errores:
        return False, errores
    with conexion(True) as con:
        con.execute("UPDATE usuarios SET password_hash=?, forzar_cambio_clave=0, fecha_cambio_clave=? WHERE cedula=?",
                    (hash_password(nueva), ahora_txt(), ced))
        registrar_auditoria(con, _usuario_publico(u), "SEGURIDAD", "usuarios", ced, "EDITAR", "password", "***", "***",
                            "Cambio de contraseña por el titular")
    return True, []


def df_usuarios() -> pd.DataFrame:
    return df_sql("SELECT cedula, nombre, rol, cargo_vuelo, licencia, email, activo, forzar_cambio_clave, ultimo_acceso, "
                  "bloqueado_hasta, fecha_creacion FROM usuarios ORDER BY rol, nombre")


def mapa_categorias_pilotos() -> dict:
    return {r["cedula"]: r["cargo_vuelo"] for r in consulta("SELECT cedula, cargo_vuelo FROM usuarios WHERE cargo_vuelo IS NOT NULL")}


def buscar_cedula_por_nombre_o_licencia(nombre=None, licencia=None) -> str | None:
    """Intenta asociar el tripulante del libro con un usuario (por licencia = cédula o por nombre)."""
    dig = R.solo_digitos(licencia)
    if len(dig) >= 6:
        r = consulta("SELECT cedula FROM usuarios WHERE cedula=? OR REPLACE(REPLACE(licencia,' ',''),'-','') LIKE ?",
                     (dig, f"%{dig}%"))
        if r:
            return r[0]["cedula"]
    if R.texto(nombre):
        clave = R.nombre_clave(nombre)
        for u in consulta("SELECT cedula, nombre FROM usuarios WHERE rol='PILOTO'"):
            if R.nombre_clave(u["nombre"]) == clave:
                return u["cedula"]
    return dig if len(dig) >= 6 else None


# ============================================================
# Flota, modelos y catálogos
# ============================================================
CAMPOS_MODELO = ["nombre", "sillas_ofrecidas", "carga_ofrecida_kg", "tipo_trafico_default", "uso_aeronave_default",
                 "tipo_combustible_default", "metodo_medicion_default", "reserva_combustible_lbs",
                 "factor_conversion_lbs_gal", "ciclo_inspeccion_horas", "activo"]
CAMPOS_AERONAVE = ["modelo_codigo", "horas_totales", "ciclos_totales", "proxima_inspeccion_horas",
                   "proxima_inspeccion_ciclos", "base", "activo", "notas"]


def df_modelos() -> pd.DataFrame:
    return df_sql("SELECT * FROM modelos_aeronave ORDER BY codigo")


def df_aeronaves() -> pd.DataFrame:
    return df_sql("SELECT a.*, m.nombre AS modelo_nombre FROM aeronaves a LEFT JOIN modelos_aeronave m "
                  "ON m.codigo=a.modelo_codigo ORDER BY a.matricula")


def modelos_dict() -> dict:
    return {r["codigo"]: r for r in consulta("SELECT * FROM modelos_aeronave")}


def aeronaves_dict() -> dict:
    return {r["matricula"]: r for r in consulta("SELECT * FROM aeronaves")}


def modelo_de_matricula(matricula) -> dict | None:
    m = R.norm_matricula(matricula)
    r = consulta("SELECT m.* FROM aeronaves a JOIN modelos_aeronave m ON m.codigo=a.modelo_codigo WHERE a.matricula=?", (m,))
    return r[0] if r else None


def _limpiar_valor_flota(k, v):
    if R.es_vacio(v):
        return None
    if k in ("sillas_ofrecidas", "ciclos_totales", "proxima_inspeccion_ciclos", "activo"):
        return R.a_int(v)
    if k in ("carga_ofrecida_kg", "reserva_combustible_lbs", "factor_conversion_lbs_gal", "ciclo_inspeccion_horas",
             "horas_totales", "proxima_inspeccion_horas"):
        return R.a_float(v)
    if k in ("tipo_trafico_default", "uso_aeronave_default", "tipo_combustible_default", "metodo_medicion_default",
             "modelo_codigo"):
        return R.texto(v).upper().replace(" ", "")[:12] or None
    if k == "base":
        return R.norm_iata(v) or None
    return R.texto(v)[:200]


def guardar_modelo(datos: dict, usuario, motivo, modulo="GERENCIA") -> tuple[bool, str]:
    codigo = re.sub(r"[^A-Z0-9]", "", R.texto(datos.get("codigo")).upper())[:10]
    if not codigo:
        return False, "Escriba el código del modelo (ej. JS41)."
    actual = _fila_unica("SELECT * FROM modelos_aeronave WHERE codigo=?", (codigo,))
    nuevos = {k: _limpiar_valor_flota(k, datos.get(k)) for k in CAMPOS_MODELO if k in datos}
    if "activo" in nuevos and nuevos["activo"] is None:
        nuevos["activo"] = 1
    with conexion(True) as con:
        if not actual:
            cols = ["codigo"] + list(nuevos)
            con.execute(f"INSERT INTO modelos_aeronave ({','.join(cols)}, actualizado_por, fecha_actualizacion) "
                        f"VALUES ({','.join('?' * len(cols))},?,?)", [codigo] + list(nuevos.values()) + [usuario.get("cedula"), ahora_txt()])
            registrar_auditoria(con, usuario, modulo, "modelos_aeronave", codigo, "PARAMETRO", None, None, nuevos, motivo or "Nuevo modelo")
            return True, f"Modelo {codigo} creado."
        cambios = {k: v for k, v in nuevos.items() if str(v) != str(actual.get(k))}
        for k, v in cambios.items():
            con.execute(f"UPDATE modelos_aeronave SET {k}=?, actualizado_por=?, fecha_actualizacion=? WHERE codigo=?",
                        (v, usuario.get("cedula"), ahora_txt(), codigo))
            registrar_auditoria(con, usuario, modulo, "modelos_aeronave", codigo, "PARAMETRO", k, actual.get(k), v, motivo)
    return True, f"Modelo {codigo} actualizado." if cambios else "Sin cambios."


def guardar_aeronave(datos: dict, usuario, motivo, modulo="GERENCIA") -> tuple[bool, str]:
    mat = R.norm_matricula(datos.get("matricula"))
    if not re.fullmatch(r"HK-[A-Z0-9]{3,5}", mat or ""):
        return False, "Matrícula no válida (formato HK-1234)."
    nuevos = {k: _limpiar_valor_flota(k, datos.get(k)) for k in CAMPOS_AERONAVE if k in datos}
    if nuevos.get("modelo_codigo") and not consulta("SELECT 1 FROM modelos_aeronave WHERE codigo=?", (nuevos["modelo_codigo"],)):
        return False, f"El modelo {nuevos['modelo_codigo']} no existe. Créelo primero."
    actual = _fila_unica("SELECT * FROM aeronaves WHERE matricula=?", (mat,))
    with conexion(True) as con:
        if not actual:
            nuevos.setdefault("activo", 1)
            cols = ["matricula"] + list(nuevos)
            con.execute(f"INSERT INTO aeronaves ({','.join(cols)}, actualizado_por, fecha_actualizacion) "
                        f"VALUES ({','.join('?' * len(cols))},?,?)", [mat] + list(nuevos.values()) + [usuario.get("cedula"), ahora_txt()])
            registrar_auditoria(con, usuario, modulo, "aeronaves", mat, "PARAMETRO", None, None, nuevos, motivo or "Nueva aeronave")
            return True, f"Aeronave {mat} registrada."
        cambios = {k: v for k, v in nuevos.items() if str(v) != str(actual.get(k))}
        for k, v in cambios.items():
            con.execute(f"UPDATE aeronaves SET {k}=?, actualizado_por=?, fecha_actualizacion=? WHERE matricula=?",
                        (v, usuario.get("cedula"), ahora_txt(), mat))
            registrar_auditoria(con, usuario, modulo, "aeronaves", mat, "PARAMETRO", k, actual.get(k), v, motivo)
    return True, f"Aeronave {mat} actualizada." if cambios else "Sin cambios."


def _asegurar_aeronave(con, matricula, usuario, modelo=None):
    """Registra la matrícula si no existe (modelo en blanco si no se conoce): nunca bloquea la captura."""
    mat = R.norm_matricula(matricula)
    if not mat or con.execute("SELECT 1 FROM aeronaves WHERE matricula=?", (mat,)).fetchone():
        return
    if modelo and not con.execute("SELECT 1 FROM modelos_aeronave WHERE codigo=?", (modelo,)).fetchone():
        modelo = None
    con.execute("INSERT INTO aeronaves (matricula, modelo_codigo, activo, actualizado_por, fecha_actualizacion) "
                "VALUES (?,?,1,?,?)", (mat, modelo, usuario.get("cedula"), ahora_txt()))
    registrar_auditoria(con, usuario, "FLOTA", "aeronaves", mat, "CREAR", None, None, {"modelo_codigo": modelo},
                        "Matrícula detectada en un soporte")


def df_rutas(incluir_inactivas=False) -> pd.DataFrame:
    q = "SELECT * FROM catalogo_rutas" + ("" if incluir_inactivas else " WHERE activo=1")
    df = df_sql(q)
    if not df.empty:
        df["_n"] = pd.to_numeric(df["numero_vuelo"], errors="coerce")
        df = df.sort_values(["_n", "numero_vuelo"]).drop(columns="_n")
    return df


def rutas_dict() -> dict:
    return {r["numero_vuelo"]: r for r in consulta("SELECT * FROM catalogo_rutas WHERE activo=1")}


def guardar_ruta(datos: dict, usuario, motivo, modulo="GERENCIA") -> tuple[bool, str]:
    nro = R.solo_digitos(datos.get("numero_vuelo"))
    if not nro:
        return False, "Escriba el número de vuelo."
    o, d = R.norm_iata(datos.get("origen")), R.norm_iata(datos.get("destino"))
    if (not o or not d) and R.texto(datos.get("ruta")):
        o, d = R.separar_ruta(datos.get("ruta"))
    tipo = R.texto(datos.get("tipo_vuelo")).upper()[:1] or None
    if tipo and tipo not in R.TIPOS_VUELO:
        return False, "Tipo de vuelo debe ser R, A o C."
    hora = R.parse_hora(datos.get("hora"))
    nuevos = {"ruta": f"{o}-{d}" if o and d else R.texto(datos.get("ruta")) or None, "origen": o or None,
              "destino": d or None, "hora": R.fmt_hora(hora) or None, "tipo_vuelo": tipo,
              "combustible_estandar": R.a_float(datos.get("combustible_estandar"))}
    actual = _fila_unica("SELECT * FROM catalogo_rutas WHERE numero_vuelo=? AND activo=1", (nro,))
    with conexion(True) as con:
        if not actual:
            con.execute("INSERT INTO catalogo_rutas (numero_vuelo, ruta, origen, destino, hora, tipo_vuelo, combustible_estandar, "
                        "activo, fuente, actualizado_por, fecha_actualizacion) VALUES (?,?,?,?,?,?,?,1,?,?,?)",
                        (nro, nuevos["ruta"], nuevos["origen"], nuevos["destino"], nuevos["hora"], nuevos["tipo_vuelo"],
                         nuevos["combustible_estandar"], "MANUAL", usuario.get("cedula"), ahora_txt()))
            registrar_auditoria(con, usuario, modulo, "catalogo_rutas", nro, "PARAMETRO", None, None, nuevos, motivo or "Nueva ruta")
            return True, f"Vuelo {nro} agregado al catálogo."
        for k, v in nuevos.items():
            if str(v) != str(actual.get(k)):
                con.execute(f"UPDATE catalogo_rutas SET {k}=?, actualizado_por=?, fecha_actualizacion=? WHERE id=?",
                            (v, usuario.get("cedula"), ahora_txt(), actual["id"]))
                registrar_auditoria(con, usuario, modulo, "catalogo_rutas", nro, "PARAMETRO", k, actual.get(k), v, motivo)
    return True, f"Vuelo {nro} actualizado."


def desactivar_ruta(numero_vuelo, usuario, motivo) -> tuple[bool, str]:
    nro = R.solo_digitos(numero_vuelo)
    with conexion(True) as con:
        n = con.execute("UPDATE catalogo_rutas SET activo=0, actualizado_por=?, fecha_actualizacion=? WHERE numero_vuelo=? AND activo=1",
                        (usuario.get("cedula"), ahora_txt(), nro)).rowcount
        if n:
            registrar_auditoria(con, usuario, "GERENCIA", "catalogo_rutas", nro, "ELIMINAR", "activo", 1, 0, motivo)
    return (True, f"Vuelo {nro} desactivado.") if n else (False, "No se encontró el vuelo activo.")


def df_rutas_charter() -> pd.DataFrame:
    return df_sql("SELECT * FROM rutas_charter WHERE activo=1 ORDER BY origen, destino")


def rutas_charter_dict() -> dict:
    out = {}
    for r in consulta("SELECT * FROM rutas_charter WHERE activo=1"):
        out[(r["origen"], r["destino"], r["modelo_codigo"] or "")] = r
    return out


def guardar_ruta_charter(origen, destino, galones, modelo, usuario, motivo, modulo="OPERACIONES") -> tuple[bool, str]:
    o, d = R.norm_iata(origen), R.norm_iata(destino)
    g = R.a_float(galones)
    if not o or not d:
        return False, "Indique origen y destino."
    if g is None or g <= 0:
        return False, "Indique los galones estándar de la ruta."
    modelo = R.texto(modelo).upper() or None
    with conexion(True) as con:
        prev = con.execute("SELECT id, combustible_galones FROM rutas_charter WHERE origen=? AND destino=? "
                           "AND COALESCE(modelo_codigo,'')=? AND activo=1", (o, d, modelo or "")).fetchone()
        if prev:
            con.execute("UPDATE rutas_charter SET combustible_galones=?, actualizado_por=?, fecha_actualizacion=? WHERE id=?",
                        (g, usuario.get("cedula"), ahora_txt(), prev[0]))
            registrar_auditoria(con, usuario, modulo, "rutas_charter", prev[0], "PARAMETRO", "combustible_galones", prev[1], g, motivo)
        else:
            cur = con.execute("INSERT INTO rutas_charter (origen, destino, modelo_codigo, combustible_galones, activo, "
                              "actualizado_por, fecha_actualizacion) VALUES (?,?,?,?,1,?,?)",
                              (o, d, modelo, g, usuario.get("cedula"), ahora_txt()))
            registrar_auditoria(con, usuario, modulo, "rutas_charter", cur.lastrowid, "PARAMETRO", None, None,
                                {"origen": o, "destino": d, "galones": g, "modelo": modelo}, motivo)
    return True, f"Ruta chárter {o}-{d} guardada."


# ============================================================
# Importadores (NOVIEMBRE.xlsx y base v2)
# ============================================================
def _abrir_libro(fuente, data_only=True):
    import openpyxl
    if isinstance(fuente, (bytes, bytearray)):
        return openpyxl.load_workbook(io.BytesIO(fuente), data_only=data_only, read_only=True)
    return openpyxl.load_workbook(fuente, data_only=data_only, read_only=True)


def _leer_bd(wb) -> list[dict]:
    if "BD" not in wb.sheetnames:
        return []
    ws, filas = wb["BD"], []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if len(r) < 8:
            continue
        nro, ruta, o, d, hora, tipo, fuel = r[1], r[2], r[3], r[4], r[5], r[6], r[7]
        nro = R.solo_digitos(nro)
        if not nro:
            continue
        o2, d2 = R.separar_ruta(ruta)
        o = R.norm_iata(o) if isinstance(o, str) and not str(o).startswith("=") else o2
        d = R.norm_iata(d) if isinstance(d, str) and not str(d).startswith("=") else d2
        tipo = R.texto(tipo).upper()[:1] or None
        filas.append({"numero_vuelo": nro, "ruta": R.texto(ruta) or (f"{o}-{d}" if o and d else None),
                      "origen": o or o2 or None, "destino": d or d2 or None, "hora": R.fmt_hora(R.parse_hora(hora)) or None,
                      "tipo_vuelo": tipo if tipo in R.TIPOS_VUELO else None,
                      "combustible_estandar": R.a_float(fuel) if not str(fuel or "").startswith("=") else None})
    return filas


def importar_catalogo_bd(fuente, usuario, reemplazar=False) -> int:
    """Importa la hoja BD (NUMERO DE VUELO, RUTA, ORIGEN, DESTINO, HORA, TIPO VLOS, FUEL)."""
    wb = _abrir_libro(fuente)
    try:
        filas = _leer_bd(wb)
    finally:
        wb.close()
    n = 0
    with conexion(True) as con:
        for f in filas:
            existe = con.execute("SELECT id FROM catalogo_rutas WHERE numero_vuelo=? AND activo=1", (f["numero_vuelo"],)).fetchone()
            if existe and not reemplazar:
                continue
            if existe:
                con.execute("UPDATE catalogo_rutas SET activo=0 WHERE id=?", (existe[0],))
            con.execute("INSERT INTO catalogo_rutas (numero_vuelo, ruta, origen, destino, hora, tipo_vuelo, combustible_estandar, "
                        "activo, fuente, actualizado_por, fecha_actualizacion) VALUES (?,?,?,?,?,?,?,1,?,?,?)",
                        (f["numero_vuelo"], f["ruta"], f["origen"], f["destino"], f["hora"], f["tipo_vuelo"],
                         f["combustible_estandar"], "NOVIEMBRE.xlsx · BD", usuario.get("cedula"), ahora_txt()))
            n += 1
        registrar_auditoria(con, usuario, "PARAMETROS", "catalogo_rutas", None, "PARAMETRO", None, None,
                            {"filas_importadas": n}, "Importación hoja BD (catálogo de vuelos y rutas)")
    return n


def importar_flota_desde_mes(fuente, usuario) -> int:
    """Registra las matrículas que aparecen en la hoja MES con su tipo de equipo (sin sobrescribir)."""
    wb = _abrir_libro(fuente)
    pares = {}
    try:
        if "MES" not in wb.sheetnames:
            return 0
        for r in wb["MES"].iter_rows(min_row=6, values_only=True):
            if len(r) < 10 or R.es_vacio(r[4]):
                continue
            mat, eq = R.norm_matricula(r[8]), R.texto(r[9]).upper()
            if mat and not str(r[8]).startswith("="):
                pares.setdefault(mat, eq or None)
    finally:
        wb.close()
    n = 0
    with conexion(True) as con:
        for mat, eq in pares.items():
            if con.execute("SELECT 1 FROM aeronaves WHERE matricula=?", (mat,)).fetchone():
                continue
            if eq and not con.execute("SELECT 1 FROM modelos_aeronave WHERE codigo=?", (eq,)).fetchone():
                con.execute("INSERT INTO modelos_aeronave (codigo, nombre, activo, actualizado_por, fecha_actualizacion) "
                            "VALUES (?,?,1,?,?)", (eq, eq, usuario.get("cedula"), ahora_txt()))
            con.execute("INSERT INTO aeronaves (matricula, modelo_codigo, activo, actualizado_por, fecha_actualizacion) "
                        "VALUES (?,?,1,?,?)", (mat, eq, usuario.get("cedula"), ahora_txt()))
            n += 1
        if n:
            registrar_auditoria(con, usuario, "PARAMETROS", "aeronaves", None, "PARAMETRO", None, None,
                                {"matriculas": sorted(pares)}, "Flota detectada en la hoja MES")
    return n


def importar_mes_plano(fuente, usuario) -> dict:
    """Importa filas reales de la hoja MES (MO-A3-013) al archivo plano como histórico (fuente IMPORTADO).
    Resuelve las fórmulas XLOOKUP con la hoja BD cuando el valor calculado no está en caché. Sin duplicados."""
    wb = _abrir_libro(fuente, data_only=True)
    try:
        bd = {f["numero_vuelo"]: f for f in _leer_bd(wb)}
        if "MES" not in wb.sheetnames:
            return {"insertadas": 0, "duplicadas": 0, "omitidas": 0, "mensaje": "El archivo no tiene hoja MES."}
        filas = [r for r in wb["MES"].iter_rows(min_row=6, values_only=True) if len(r) >= 27 and not R.es_vacio(r[4])]
    finally:
        wb.close()
    p = params_dict()
    ins = dup = omit = 0
    with conexion(True) as con:
        existentes = {(r[0], r[1], r[2], r[3], r[4]) for r in con.execute(
            "SELECT fecha_utc, hora_utc, matricula, origen, destino FROM operaciones_plano WHERE eliminado=0").fetchall()}
        for r in filas:
            v = list(r[:27])
            v = [None if isinstance(x, str) and x.startswith("=") else x for x in v]
            nro = R.solo_digitos(v[10])
            ref = bd.get(nro, {})
            fila = {
                "tip": R.texto(v[0]) or R.texto(p.get("plano_tip"), "TE"), "sigla": R.texto(v[2]) or R.texto(p.get("plano_sigla"), "3CA"),
                "fecha_utc": R.fecha_ddmmaaaa(str(v[4]).strip().zfill(8)) if v[4] is not None else None,
                "hora_utc": R.normalizar_valor("plano_hora", str(v[5]).zfill(4) if v[5] is not None else None),
                "origen": R.norm_iata(v[6]) or ref.get("origen"), "destino": R.norm_iata(v[7]) or ref.get("destino"),
                "matricula": R.matricula_compacta(v[8]) or None, "tipo_equipo": R.texto(v[9]).upper() or None,
                "nro_vuelo": nro or None, "sillas_ofrecidas": R.a_int(v[11]), "carga_ofrecida": R.a_float(v[12]),
                "pax_bordo": R.a_int(v[13]), "carga_bordo": R.a_float(v[14]), "correo_bordo": R.a_float(v[15]),
                "pax_transito": R.a_int(v[16]), "carga_transito": R.a_float(v[17]),
                "horas_bloque": R.normalizar_valor("plano_hhmm", str(v[18]).zfill(4) if v[18] is not None else None),
                "tipo_vuelo": (R.texto(v[19]).upper() or ref.get("tipo_vuelo") or None),
                "tipo_trafico": R.texto(v[20]).upper() or None, "uso_aeronave": R.texto(v[21]).upper() or None,
                "tipo_combustible": R.texto(v[22]).upper() or None, "metodo_medicion": R.texto(v[23]).upper() or None,
                "combustible_consumido": R.a_float(v[24], ref.get("combustible_estandar")),
                "horas_vuelo": R.normalizar_valor("plano_hhmm", str(v[25]).zfill(4) if v[25] is not None else None),
                "carga_real": R.a_float(v[26]),
            }
            if not fila["fecha_utc"] or not fila["matricula"]:
                omit += 1
                continue
            fila["fecha_mmaaaa"] = R.normalizar_valor("mmaaaa", v[3]) if v[3] else None
            clave = (fila["fecha_utc"], fila["hora_utc"], fila["matricula"], fila["origen"], fila["destino"])
            if clave in existentes:
                dup += 1
                continue
            existentes.add(clave)
            f_local, h_local = R.desde_utc(fila["fecha_utc"], fila["hora_utc"], p.get("offset_utc_horas") or 5)
            if not fila["fecha_mmaaaa"] and f_local:
                fila["fecha_mmaaaa"] = f_local.strftime("%m%Y")
            nro_plano, charter = R.normalizar_nro_vuelo(fila["nro_vuelo"], fila["matricula"])
            meta = {"fuente": "IMPORTADO", "fecha_local": f_local.isoformat() if f_local else None, "hora_local_min": h_local,
                    "es_charter": 1 if (charter or fila["tipo_vuelo"] == "C") else 0,
                    "campos_manuales": json.dumps([c for c in R.PLANO_CLAVES if c != "nro"]),
                    "pendientes": json.dumps(R.pendientes_fila(fila)), "notas": "Importado de plantilla MO-A3-013",
                    "origen_carga": "IMPORTACION", "creado_por_cedula": usuario.get("cedula"),
                    "creado_por_nombre": usuario.get("nombre"), "creado_por_rol": usuario.get("rol"),
                    "fecha_creacion": ahora_txt(), "validado_por_cedula": usuario.get("cedula"), "fecha_validacion": ahora_txt()}
            datos = {k: fila.get(k) for k in CAMPOS["operaciones_plano"]}
            meta["hash_integridad"] = R.huella(datos)
            cols = list(datos) + list(meta)
            con.execute(f"INSERT INTO operaciones_plano ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                        list(datos.values()) + list(meta.values()))
            _asegurar_aeronave(con, fila["matricula"], usuario, fila["tipo_equipo"])
            ins += 1
        registrar_auditoria(con, usuario, "OPERACIONES", "operaciones_plano", None, "CREAR", None, None,
                            {"insertadas": ins, "duplicadas": dup, "omitidas": omit}, "Importación histórica hoja MES (MO-A3-013)")
    return {"insertadas": ins, "duplicadas": dup, "omitidas": omit,
            "mensaje": f"{ins} vuelos importados · {dup} ya existían · {omit} filas sin fecha/matrícula."}


def importar_legacy_v2() -> int:
    """Migra una sola vez los tiquetes de la versión 2 (moon_fuel.db, tabla registros) sin perder imágenes."""
    if get_meta("legacy_v2_importado"):
        return 0
    origen = next((p for p in (DATA_DIR / "moon_fuel.db", BASE_DIR / "moon_fuel.db") if p.is_file()), None)
    if not origen:
        return 0
    src = sqlite3.connect(f"file:{origen}?mode=ro", uri=True)
    try:
        if not src.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='registros'").fetchone():
            return 0
        filas = _filas(src, "SELECT * FROM registros ORDER BY id")
    finally:
        src.close()
    n = 0
    with conexion(True) as con:
        for f in filas:
            doc_id = None
            nombre_img = R.texto(f.get("imagen"))
            for base in (origen.parent / "imagenes_tiquetes", DATA_DIR / "imagenes_tiquetes"):
                p = base / nombre_img if nombre_img else None
                if p and p.is_file():
                    sha = hashlib.sha256(p.read_bytes()).hexdigest()
                    if not con.execute("SELECT 1 FROM documentos WHERE sha256=? AND eliminado=0", (sha,)).fetchone():
                        try:
                            rel = p.resolve().relative_to(DATA_DIR)
                        except ValueError:
                            destino = EVIDENCIAS_DIR / "migracion_v2" / p.name
                            destino.parent.mkdir(parents=True, exist_ok=True)
                            shutil.copy2(p, destino)
                            rel = destino.relative_to(DATA_DIR)
                        cur = con.execute("INSERT INTO documentos (tipo, fecha_doc, matricula, ruta_archivo, sha256, origen_carga, "
                                          "creado_por_cedula, creado_por_rol, fecha_creacion) VALUES ('COMB',?,?,?,?,?,?,?,?)",
                                          (R.normalizar_valor("fecha", f.get("fecha_consumo")), R.norm_matricula(f.get("matricula")),
                                           str(rel).replace("\\", "/"), sha, "MIGRACION_V2", "SISTEMA", "SISTEMA", ahora_txt()))
                        doc_id = cur.lastrowid
                    break
            datos = {"recibo": R.texto(f.get("recibo")) or None, "numero_boleto": None,
                     "fecha_consumo": R.normalizar_valor("fecha", f.get("fecha_consumo")), "hora_min": None,
                     "estacion": None, "matricula": R.norm_matricula(f.get("matricula")) or None,
                     "producto": R.norm_producto(f.get("producto")), "destino": None, "numero_vuelo": None,
                     "galones": R.a_float(f.get("consumo_gal")), "valor_unit": R.a_float(f.get("valor_unit")) or None,
                     "valor_total": R.a_float(f.get("valor_total")) or None, "piloto_nombre": R.texto(f.get("piloto")) or None,
                     "piloto_cedula": None, "observacion": R.texto(f.get("observacion")) or "Migrado de la versión 2"}
            meta = {"documento_id": doc_id, "origen_carga": "MIGRACION_V2", "creado_por_cedula": "SISTEMA",
                    "creado_por_nombre": "Migración v2", "creado_por_rol": "SISTEMA",
                    "fecha_creacion": R.texto(f.get("fecha_registro")) or ahora_txt(), "hash_integridad": R.huella(datos)}
            cols = list(datos) + list(meta)
            cur = con.execute(f"INSERT INTO combustible_tiquetes ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                              list(datos.values()) + list(meta.values()))
            _asegurar_aeronave(con, datos["matricula"], SISTEMA)
            n += 1
        registrar_auditoria(con, SISTEMA, "MIGRACION", "combustible_tiquetes", None, "CREAR", None, None,
                            {"tiquetes_migrados": n, "origen": origen.name}, "Migración automática desde la versión 2")
        con.execute("INSERT OR REPLACE INTO meta VALUES ('legacy_v2_importado', ?)", (ahora_txt(),))
    return n


# ============================================================
# Documentos (evidencias) y capturas
# ============================================================
def siguiente_seq(tipo, fecha, matricula) -> int:
    r = consulta("SELECT COUNT(*) n FROM documentos WHERE tipo=? AND fecha_doc=? AND matricula=? AND eliminado=0",
                 (tipo, R.normalizar_valor("fecha", fecha), R.norm_matricula(matricula)))
    return int(r[0]["n"]) + 1


def documento_por_hash(sha) -> dict | None:
    return _fila_unica("SELECT * FROM documentos WHERE sha256=? AND eliminado=0", (sha,))


def obtener_documento(doc_id) -> dict | None:
    if R.es_vacio(doc_id):
        return None
    return _fila_unica("SELECT * FROM documentos WHERE id=?", (int(doc_id),))


def _insertar_documento(con, doc: dict, usuario, origen_carga) -> int:
    cur = con.execute(
        "INSERT INTO documentos (tipo, fecha_doc, hora_min, matricula, seq_dia, es_charter, ruta_archivo, sha256, ancho, alto, "
        "nitidez, origen_carga, creado_por_cedula, creado_por_nombre, creado_por_rol, fecha_creacion, extraccion_ia, campos_dudosos) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (doc["tipo"], R.normalizar_valor("fecha", doc.get("fecha_doc")), doc.get("hora_min"), R.norm_matricula(doc.get("matricula")),
         doc.get("seq_dia"), 1 if doc.get("es_charter") else 0, doc["ruta_archivo"], doc["sha256"], doc.get("ancho"),
         doc.get("alto"), doc.get("nitidez"), origen_carga, usuario.get("cedula"), usuario.get("nombre"), usuario.get("rol"),
         ahora_txt(), _json(doc.get("extraccion_ia")), _json(doc.get("campos_dudosos"))))
    registrar_auditoria(con, usuario, origen_carga, "documentos", cur.lastrowid, "CREAR", None, None,
                        {"tipo": doc["tipo"], "archivo": doc["ruta_archivo"], "sha256": doc["sha256"]}, "Evidencia almacenada")
    return cur.lastrowid


def _meta_creacion(usuario, origen_carga) -> dict:
    t = ahora_txt()
    return {"origen_carga": origen_carga, "creado_por_cedula": usuario.get("cedula"), "creado_por_nombre": usuario.get("nombre"),
            "creado_por_rol": usuario.get("rol"), "fecha_creacion": t, "validado_por_cedula": usuario.get("cedula"),
            "fecha_validacion": t}


def _log_correcciones(con, usuario, modulo, tabla, rid, correcciones):
    for c in correcciones or []:
        registrar_auditoria(con, usuario, modulo, tabla, rid, "CORRECCION_IA", c.get("campo"), c.get("ia"), c.get("humano"),
                            "Corrección humana sobre la lectura automática")


def _nota_cierre(con, area, fecha) -> str:
    if area and fecha and not _periodo_editable_con(con, area, fecha):
        return " · Registro posterior al cierre del periodo"
    return ""


def registrar_tiquete(datos: dict, doc: dict | None, usuario: dict, origen_carga: str, correcciones=None) -> int:
    limpio = {k: R.normalizar_valor(t, datos.get(k)) for k, t in CAMPOS["combustible_tiquetes"].items()}
    if limpio["galones"] is not None and limpio["valor_unit"] is not None:
        limpio["valor_total"] = round(limpio["galones"] * limpio["valor_unit"])
    with conexion(True) as con:
        doc_id = _insertar_documento(con, doc, usuario, origen_carga) if doc else None
        meta = _meta_creacion(usuario, origen_carga)
        meta.update({"documento_id": doc_id, "hash_integridad": R.huella(limpio)})
        cols = list(limpio) + list(meta)
        cur = con.execute(f"INSERT INTO combustible_tiquetes ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                          list(limpio.values()) + list(meta.values()))
        rid = cur.lastrowid
        _asegurar_aeronave(con, limpio["matricula"], usuario)
        nota = _nota_cierre(con, "CONTABILIDAD", limpio["fecha_consumo"])
        registrar_auditoria(con, usuario, origen_carga, "combustible_tiquetes", rid, "CREAR", None, None, limpio,
                            "Registro confirmado con cláusula de responsabilidad" + nota)
        _log_correcciones(con, usuario, origen_carga, "combustible_tiquetes", rid, correcciones)
    return rid


def registrar_libro(cabecera: dict, tramos: list[dict], mantenimiento: list[dict], doc: dict | None, usuario: dict,
                    origen_carga: str, correcciones=None) -> int:
    cab = {k: R.normalizar_valor(t, cabecera.get(k)) for k, t in CAMPOS["libros_vuelo"].items()}
    mat = cab["matricula"]
    if not cab.get("piloto_cedula"):
        cab["piloto_cedula"] = buscar_cedula_por_nombre_o_licencia(cab.get("piloto_nombre"), cab.get("piloto_licencia"))
    if not cab.get("copiloto_cedula"):
        cab["copiloto_cedula"] = buscar_cedula_por_nombre_o_licencia(cab.get("copiloto_nombre"), cab.get("copiloto_licencia"))
    with conexion(True) as con:
        doc_id = _insertar_documento(con, doc, usuario, origen_carga) if doc else None
        meta = _meta_creacion(usuario, origen_carga)
        meta.update({"documento_id": doc_id, "hash_integridad": R.huella(cab),
                     "descuadre_totales": 1 if cabecera.get("descuadre_totales") else 0,
                     "parametros_motor": _json(cabecera.get("parametros_motor"))})
        cols = list(cab) + list(meta)
        cur = con.execute(f"INSERT INTO libros_vuelo ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                          list(cab.values()) + list(meta.values()))
        libro_id = cur.lastrowid
        _asegurar_aeronave(con, mat, usuario)
        nota = _nota_cierre(con, "OPERACIONES", cab["fecha"]) or _nota_cierre(con, "CONTABILIDAD", cab["fecha"])
        registrar_auditoria(con, usuario, origen_carga, "libros_vuelo", libro_id, "CREAR", None, None, cab,
                            "Libro de vuelo confirmado con cláusula de responsabilidad" + nota)
        for i, t in enumerate(tramos):
            tr = {k: R.normalizar_valor(tp, t.get(k)) for k, tp in CAMPOS["tramos_vuelo"].items()}
            nro_plano, charter = R.normalizar_nro_vuelo(tr["numero_vuelo"], mat)
            extra = {"libro_id": libro_id, "orden": i + 1, "nro_vuelo_plano": nro_plano, "es_charter": 1 if charter else 0,
                     "tiempo_pilotos_origen": "MANUAL" if tr.get("tiempo_pilotos_min") is not None else None,
                     "hash_integridad": R.huella(tr)}
            cols = list(tr) + list(extra)
            c2 = con.execute(f"INSERT INTO tramos_vuelo ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                             list(tr.values()) + list(extra.values()))
            registrar_auditoria(con, usuario, origen_carga, "tramos_vuelo", c2.lastrowid, "CREAR", None, None, tr,
                                f"Tramo {i + 1} del libro {cab.get('numero_libro') or libro_id}")
        for m in mantenimiento or []:
            item = {k: R.normalizar_valor(tp, m.get(k)) for k, tp in CAMPOS["mantenimiento_oma"].items()}
            if not any(item.get(k) for k in ("reporte", "accion_correctiva", "pn_on", "pn_off")):
                continue
            item["fecha"] = item.get("fecha") or cab["fecha"]
            item["matricula"] = item.get("matricula") or mat
            item["estado"] = item.get("estado") or ("CERRADO" if item.get("accion_correctiva") else "ABIERTO")
            meta_m = _meta_creacion(usuario, origen_carga)
            meta_m.update({"libro_id": libro_id, "documento_id": doc_id, "hash_integridad": R.huella(item)})
            cols = list(item) + list(meta_m)
            c3 = con.execute(f"INSERT INTO mantenimiento_oma ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                             list(item.values()) + list(meta_m.values()))
            registrar_auditoria(con, usuario, origen_carga, "mantenimiento_oma", c3.lastrowid, "CREAR", None, None, item,
                                "Ítem de mantenimiento tomado del libro de vuelo")
        # Horas/ciclos de la aeronave (solo si avanzan: nunca retroceden por error de lectura)
        a = con.execute("SELECT horas_totales, ciclos_totales FROM aeronaves WHERE matricula=?", (mat,)).fetchone()
        if a:
            if cab.get("horas_totales") is not None and (a[0] is None or cab["horas_totales"] > a[0]):
                con.execute("UPDATE aeronaves SET horas_totales=? WHERE matricula=?", (cab["horas_totales"], mat))
                registrar_auditoria(con, usuario, "OMA", "aeronaves", mat, "EDITAR", "horas_totales", a[0], cab["horas_totales"],
                                    "Actualizado desde libro de vuelo")
            if cab.get("ciclos_totales") is not None and (a[1] is None or cab["ciclos_totales"] > a[1]):
                con.execute("UPDATE aeronaves SET ciclos_totales=? WHERE matricula=?", (cab["ciclos_totales"], mat))
                registrar_auditoria(con, usuario, "OMA", "aeronaves", mat, "EDITAR", "ciclos_totales", a[1], cab["ciclos_totales"],
                                    "Actualizado desde libro de vuelo")
        _log_correcciones(con, usuario, origen_carga, "libros_vuelo", libro_id, correcciones)
    return libro_id


def registrar_fichas(fichas: list[dict], doc: dict | None, usuario: dict, origen_carga: str, correcciones=None) -> list[int]:
    ids = []
    with conexion(True) as con:
        doc_id = _insertar_documento(con, doc, usuario, origen_carga) if doc else None
        for f in fichas:
            datos = {k: R.normalizar_valor(t, f.get(k)) for k, t in CAMPOS["fichas_vuelo"].items()}
            if not datos.get("fecha") or not datos.get("matricula"):
                continue
            meta = _meta_creacion(usuario, origen_carga)
            meta.update({"documento_id": doc_id, "hash_integridad": R.huella(datos)})
            cols = list(datos) + list(meta)
            cur = con.execute(f"INSERT INTO fichas_vuelo ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                              list(datos.values()) + list(meta.values()))
            ids.append(cur.lastrowid)
            _asegurar_aeronave(con, datos["matricula"], usuario)
            nota = _nota_cierre(con, "OPERACIONES", datos["fecha"])
            registrar_auditoria(con, usuario, origen_carga, "fichas_vuelo", cur.lastrowid, "CREAR", None, None, datos,
                                "Ficha de vuelo confirmada con cláusula de responsabilidad" + nota)
            _log_correcciones(con, usuario, origen_carga, "fichas_vuelo", cur.lastrowid, correcciones)
    return ids


def registrar_item_oma(datos: dict, usuario: dict) -> int:
    item = {k: R.normalizar_valor(t, datos.get(k)) for k, t in CAMPOS["mantenimiento_oma"].items()}
    item["estado"] = item.get("estado") or "ABIERTO"
    with conexion(True) as con:
        meta = _meta_creacion(usuario, "OMA")
        meta["hash_integridad"] = R.huella(item)
        cols = list(item) + list(meta)
        cur = con.execute(f"INSERT INTO mantenimiento_oma ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                          list(item.values()) + list(meta.values()))
        _asegurar_aeronave(con, item.get("matricula"), usuario)
        registrar_auditoria(con, usuario, "OMA", "mantenimiento_oma", cur.lastrowid, "CREAR", None, None, item,
                            "Ítem registrado directamente por OMA")
    return cur.lastrowid


# ============================================================
# Edición auditada genérica y borrado lógico
# ============================================================
def _fecha_registro(con, tabla, rid) -> str | None:
    if tabla == "tramos_vuelo":
        r = con.execute("SELECT l.fecha FROM tramos_vuelo t JOIN libros_vuelo l ON l.id=t.libro_id WHERE t.id=?", (rid,)).fetchone()
        return r[0] if r else None
    col = FECHA_TABLA.get(tabla)
    if not col:
        return None
    r = con.execute(f"SELECT {col} FROM {tabla} WHERE id=?", (rid,)).fetchone()
    return r[0] if r else None


def actualizar_registro(tabla: str, registro_id: int, cambios: dict, usuario: dict, motivo: str,
                        modulo: str) -> tuple[bool, str, int]:
    """Actualiza campos de la lista blanca, con motivo obligatorio, bloqueo por periodo cerrado,
    bitácora campo a campo y recálculo de hash_integridad. Devuelve (ok, mensaje, campos_cambiados)."""
    if tabla not in CAMPOS:
        return False, "Tabla no editable.", 0
    if len(R.texto(motivo)) < 5:
        return False, "Escriba el motivo del cambio (mínimo 5 caracteres).", 0
    tipos = CAMPOS[tabla]
    desconocidos = set(cambios) - set(tipos)
    if desconocidos:
        return False, "Campos no editables: " + ", ".join(sorted(desconocidos)), 0
    with conexion(True) as con:
        actual = _fila(con, f"SELECT * FROM {tabla} WHERE id=?", (int(registro_id),))
        if not actual:
            return False, f"Registro {registro_id} no encontrado.", 0
        if actual.get("eliminado"):
            return False, "El registro está eliminado.", 0
        area = AREA_TABLA.get(tabla)
        fecha = _fecha_registro(con, tabla, int(registro_id))
        if area and fecha and not _periodo_editable_con(con, area, fecha):
            return False, (f"El {fecha} está cerrado para {area}. Gerencia debe reabrir el periodo para editar "
                           f"({ETIQUETA_TABLA.get(tabla, tabla)} #{registro_id})."), 0
        nuevos = {}
        for k, v in cambios.items():
            nv = R.normalizar_valor(tipos[k], v)
            av = actual.get(k)
            if str(nv if nv is not None else "") != str(av if av is not None else ""):
                if isinstance(nv, (int, float)) and R.a_float(av) is not None and abs(float(nv) - float(R.a_float(av))) < 1e-9:
                    continue
                nuevos[k] = nv
        if not nuevos:
            return True, "Sin cambios.", 0
        extra = {}
        if tabla == "combustible_tiquetes" and ({"galones", "valor_unit"} & set(nuevos)) and "valor_total" not in nuevos:
            g = nuevos.get("galones", actual.get("galones"))
            u = nuevos.get("valor_unit", actual.get("valor_unit"))
            vt = round(float(g) * float(u)) if g is not None and u is not None else None
            if str(vt) != str(actual.get("valor_total")):
                nuevos["valor_total"] = vt
        if tabla == "tramos_vuelo" and "tiempo_pilotos_min" in nuevos:
            extra["tiempo_pilotos_origen"] = "MANUAL" if nuevos["tiempo_pilotos_min"] is not None else None
        if tabla == "tramos_vuelo" and "numero_vuelo" in nuevos:
            mat = con.execute("SELECT l.matricula FROM libros_vuelo l JOIN tramos_vuelo t ON t.libro_id=l.id WHERE t.id=?",
                              (int(registro_id),)).fetchone()
            nro, ch = R.normalizar_nro_vuelo(nuevos["numero_vuelo"], mat[0] if mat else "")
            extra.update({"nro_vuelo_plano": nro, "es_charter": 1 if ch else 0})
        if tabla == "operaciones_plano":
            manuales = set(json.loads(actual.get("campos_manuales") or "[]"))
            for k, v in nuevos.items():
                (manuales.add if v is not None else manuales.discard)(k)
            fila_nueva = {**{k: actual.get(k) for k in R.PLANO_CLAVES if k != "nro"}, **nuevos}
            extra["campos_manuales"] = json.dumps(sorted(manuales))
            extra["pendientes"] = json.dumps(R.pendientes_fila(fila_nueva))
            if {"fecha_utc", "hora_utc"} & set(nuevos):
                fl, hl = R.desde_utc(fila_nueva.get("fecha_utc"), fila_nueva.get("hora_utc"), get_param("offset_utc_horas", 5))
                extra.update({"fecha_local": fl.isoformat() if fl else actual.get("fecha_local"), "hora_local_min": hl})
            if "tipo_vuelo" in nuevos:
                extra["es_charter"] = 1 if nuevos["tipo_vuelo"] == "C" else 0
        nuevo_registro = {**{k: actual.get(k) for k in tipos}, **nuevos}
        extra.update({"hash_integridad": R.huella({k: nuevo_registro.get(k) for k in tipos}),
                      "actualizado_por_cedula": usuario.get("cedula"), "fecha_actualizacion": ahora_txt()})
        todo = {**nuevos, **extra}
        con.execute(f"UPDATE {tabla} SET {','.join(k + '=?' for k in todo)} WHERE id=?", list(todo.values()) + [int(registro_id)])
        for k, v in nuevos.items():
            registrar_auditoria(con, usuario, modulo, tabla, registro_id, "EDITAR", k, actual.get(k), v, motivo)
    return True, f"{ETIQUETA_TABLA.get(tabla, tabla)} #{registro_id}: {len(nuevos)} campo(s) actualizado(s).", len(nuevos)


def eliminar_registro(tabla: str, registro_id: int, usuario: dict, motivo: str, modulo: str) -> tuple[bool, str]:
    """Borrado lógico (el dato se conserva para auditoría). En libros elimina también sus tramos, ítems OMA y filas del plano."""
    if tabla not in CAMPOS:
        return False, "Tabla no válida."
    if len(R.texto(motivo)) < 5:
        return False, "Escriba el motivo de la eliminación."
    with conexion(True) as con:
        actual = _fila(con, f"SELECT * FROM {tabla} WHERE id=?", (int(registro_id),))
        if not actual or actual.get("eliminado"):
            return False, "Registro no encontrado o ya eliminado."
        area, fecha = AREA_TABLA.get(tabla), _fecha_registro(con, tabla, int(registro_id))
        if area and fecha and not _periodo_editable_con(con, area, fecha):
            return False, f"El {fecha} está cerrado para {area}. Gerencia debe reabrir el periodo."
        con.execute(f"UPDATE {tabla} SET eliminado=1, motivo_eliminacion=? WHERE id=?", (motivo, int(registro_id)))
        registrar_auditoria(con, usuario, modulo, tabla, registro_id, "ELIMINAR", None, None, None, motivo)
        if tabla == "libros_vuelo":
            tramos = [r[0] for r in con.execute("SELECT id FROM tramos_vuelo WHERE libro_id=? AND eliminado=0", (int(registro_id),))]
            for tid in tramos:
                con.execute("UPDATE tramos_vuelo SET eliminado=1, motivo_eliminacion=? WHERE id=?", (motivo, tid))
                registrar_auditoria(con, usuario, modulo, "tramos_vuelo", tid, "ELIMINAR", None, None, None, "En cascada: " + motivo)
                for pid in [r[0] for r in con.execute("SELECT id FROM operaciones_plano WHERE tramo_id=? AND eliminado=0", (tid,))]:
                    con.execute("UPDATE operaciones_plano SET eliminado=1, motivo_eliminacion=? WHERE id=?", (motivo, pid))
                    registrar_auditoria(con, usuario, modulo, "operaciones_plano", pid, "ELIMINAR", None, None, None, "En cascada: " + motivo)
            for mid in [r[0] for r in con.execute("SELECT id FROM mantenimiento_oma WHERE libro_id=? AND eliminado=0", (int(registro_id),))]:
                con.execute("UPDATE mantenimiento_oma SET eliminado=1, motivo_eliminacion=? WHERE id=?", (motivo, mid))
                registrar_auditoria(con, usuario, modulo, "mantenimiento_oma", mid, "ELIMINAR", None, None, None, "En cascada: " + motivo)
        if tabla == "tramos_vuelo":
            for pid in [r[0] for r in con.execute("SELECT id FROM operaciones_plano WHERE tramo_id=? AND eliminado=0", (int(registro_id),))]:
                con.execute("UPDATE operaciones_plano SET eliminado=1, motivo_eliminacion=? WHERE id=?", (motivo, pid))
                registrar_auditoria(con, usuario, modulo, "operaciones_plano", pid, "ELIMINAR", None, None, None, "En cascada: " + motivo)
    return True, f"{ETIQUETA_TABLA.get(tabla, tabla)} #{registro_id} eliminado (borrado lógico, queda en auditoría)."


def guardar_ajuste_liquidacion(periodo, clave_persona, campo, valor, usuario, motivo) -> tuple[bool, str]:
    if campo not in R.LIQ_EDITABLES.values():
        return False, "Campo de liquidación no editable."
    if len(R.texto(motivo)) < 5:
        return False, "Escriba el motivo del ajuste."
    if periodo_mensual_cerrado("CONTABILIDAD", periodo):
        return False, f"El periodo {periodo} está cerrado. Gerencia debe reabrirlo."
    v = None if R.es_vacio(valor) else (R.texto(valor).upper() if campo == "categoria" else R.a_float(valor))
    if campo == "categoria" and v not in (None, "PTL", "PCA"):
        return False, "La categoría debe ser PTL o PCA."
    anterior = ajustes_liquidacion(periodo).get(clave_persona, {}).get(campo)
    with conexion(True) as con:
        con.execute("INSERT INTO liquidacion_ajustes (periodo, clave_persona, campo, valor, motivo, usuario_cedula, fecha) "
                    "VALUES (?,?,?,?,?,?,?)", (periodo, clave_persona, campo, None if v is None else str(v), motivo,
                                               usuario.get("cedula"), ahora_txt()))
        registrar_auditoria(con, usuario, "CONTABILIDAD", "liquidacion_ajustes", f"{periodo}|{clave_persona}", "EDITAR",
                            campo, anterior, v, motivo)
    return True, "Ajuste registrado."


def ajustes_liquidacion(periodo) -> dict:
    out: dict = {}
    for r in consulta("SELECT clave_persona, campo, valor FROM liquidacion_ajustes WHERE periodo=? ORDER BY id", (periodo,)):
        out.setdefault(r["clave_persona"], {})[r["campo"]] = r["valor"]
    return {k: {c: v for c, v in d.items() if v is not None} for k, d in out.items()}


# ============================================================
# Consultas de negocio (DataFrames)
# ============================================================
def _rango(desde, hasta):
    d = R.normalizar_valor("fecha", desde) or "1900-01-01"
    h = R.normalizar_valor("fecha", hasta) or "2999-12-31"
    return d, h


def df_tiquetes(desde=None, hasta=None, incluir_eliminados=False) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    q = ("SELECT t.*, doc.ruta_archivo, doc.sha256 AS doc_sha256 FROM combustible_tiquetes t "
         "LEFT JOIN documentos doc ON doc.id=t.documento_id WHERE COALESCE(t.fecha_consumo,'') BETWEEN ? AND ?")
    if not incluir_eliminados:
        q += " AND t.eliminado=0"
    return df_sql(q + " ORDER BY t.fecha_consumo, t.matricula, t.id", (d, h))


def df_libros(desde=None, hasta=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    return df_sql("SELECT l.*, doc.ruta_archivo, doc.sha256 AS doc_sha256 FROM libros_vuelo l "
                  "LEFT JOIN documentos doc ON doc.id=l.documento_id WHERE l.eliminado=0 AND l.fecha BETWEEN ? AND ? "
                  "ORDER BY l.fecha, l.matricula, l.id", (d, h))


def df_tramos(desde=None, hasta=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    return df_sql(
        "SELECT t.*, l.fecha, l.matricula, l.numero_libro, l.piloto_nombre, l.piloto_cedula, l.piloto_licencia, "
        "l.copiloto_nombre, l.copiloto_cedula, l.copiloto_licencia, l.documento_id, l.creado_por_cedula, "
        "l.creado_por_nombre, l.creado_por_rol, l.fecha_creacion, l.origen_carga, doc.ruta_archivo "
        "FROM tramos_vuelo t JOIN libros_vuelo l ON l.id=t.libro_id LEFT JOIN documentos doc ON doc.id=l.documento_id "
        "WHERE t.eliminado=0 AND l.eliminado=0 AND l.fecha BETWEEN ? AND ? ORDER BY l.fecha, l.matricula, "
        "COALESCE(t.hora_salida_min, t.hora_despegue_min, 0), t.orden", (d, h))


def df_fichas(desde=None, hasta=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    return df_sql("SELECT f.*, doc.ruta_archivo FROM fichas_vuelo f LEFT JOIN documentos doc ON doc.id=f.documento_id "
                  "WHERE f.eliminado=0 AND f.fecha BETWEEN ? AND ? ORDER BY f.fecha, f.matricula, f.hora_min", (d, h))


def df_mantenimiento(matricula=None, estado=None, desde=None, hasta=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    q, p = ("SELECT m.*, l.numero_libro, doc.ruta_archivo FROM mantenimiento_oma m LEFT JOIN libros_vuelo l ON l.id=m.libro_id "
            "LEFT JOIN documentos doc ON doc.id=m.documento_id WHERE m.eliminado=0 AND COALESCE(m.fecha,'') BETWEEN ? AND ?"), [d, h]
    if matricula:
        q, p = q + " AND m.matricula=?", p + [R.norm_matricula(matricula)]
    if estado:
        q, p = q + " AND m.estado=?", p + [estado]
    return df_sql(q + " ORDER BY m.fecha DESC, m.id DESC", tuple(p))


def df_plano(desde=None, hasta=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    return df_sql("SELECT * FROM operaciones_plano WHERE eliminado=0 AND COALESCE(fecha_local,'') BETWEEN ? AND ? "
                  "ORDER BY fecha_local, matricula, hora_local_min", (d, h))


def df_documentos(desde=None, hasta=None, cedula=None, tipo=None) -> pd.DataFrame:
    d, h = _rango(desde, hasta)
    q, p = "SELECT * FROM documentos WHERE eliminado=0 AND COALESCE(fecha_doc,'') BETWEEN ? AND ?", [d, h]
    if cedula:
        q, p = q + " AND creado_por_cedula=?", p + [str(cedula)]
    if tipo:
        q, p = q + " AND tipo=?", p + [tipo]
    return df_sql(q + " ORDER BY fecha_creacion DESC", tuple(p))


def _lista(df: pd.DataFrame) -> list[dict]:
    if df.empty:
        return []
    return [{k: (None if R.es_vacio(v) else (v.item() if hasattr(v, "item") else v)) for k, v in r.items()}
            for r in df.to_dict("records")]


# ============================================================
# Archivo plano: sincronización desde soportes
# ============================================================
def sincronizar_plano(desde, hasta, usuario=None) -> dict:
    """Crea/actualiza filas MO-A3-013 desde los tramos de los libros (+ fichas, flota y catálogo).
    Nunca pisa celdas editadas a mano ni días cerrados de OPERACIONES."""
    usuario = usuario or SISTEMA
    tramos = _lista(df_tramos(desde, hasta))
    fichas = _lista(df_fichas(desde, hasta))
    p, modelos, aeronaves, rutas, charter = params_dict(), modelos_dict(), aeronaves_dict(), rutas_dict(), rutas_charter_dict()
    pares = R.emparejar_fichas(tramos, fichas)
    fichas_id = {f["id"]: f for f in fichas}
    res = {"creadas": 0, "actualizadas": 0, "cerradas": 0, "eliminadas": 0}
    with conexion(True) as con:
        existentes = {r["tramo_id"]: r for r in _filas(con, "SELECT * FROM operaciones_plano WHERE tramo_id IS NOT NULL AND eliminado=0")}
        for t in tramos:
            if not _periodo_editable_con(con, "OPERACIONES", t["fecha"]):
                res["cerradas"] += 1
                continue
            aero = aeronaves.get(R.norm_matricula(t["matricula"])) or {}
            modelo = modelos.get(aero.get("modelo_codigo")) if aero.get("modelo_codigo") else None
            nro, _ = R.normalizar_nro_vuelo(t.get("numero_vuelo"), t["matricula"])
            ruta = rutas.get(nro) if nro else None
            ficha = fichas_id.get(pares.get(t["id"]))
            o, d = R.norm_iata(t.get("origen")), R.norm_iata(t.get("destino"))
            rc = charter.get((o, d, (modelo or {}).get("codigo", ""))) or charter.get((o, d, ""))
            fila, pend, notas = R.construir_fila_plano(t, modelo, ruta, ficha, rc, p)
            datos = {k: fila.get(k) for k in CAMPOS["operaciones_plano"]}
            meta = {"fecha_local": fila["fecha_local"], "hora_local_min": fila["hora_local_min"], "es_charter": fila["es_charter"],
                    "ficha_id": ficha["id"] if ficha else None, "notas": " · ".join(notas)[:1000]}
            previo = existentes.get(t["id"])
            if not previo:
                # La fila la genera el sistema a partir del libro: la responsabilidad del dato es de quien confirmó el libro.
                meta.update(_meta_creacion(SISTEMA, "SOPORTES"))
                meta.update({"tramo_id": t["id"], "fuente": "SOPORTES", "campos_manuales": "[]",
                             "pendientes": json.dumps(pend), "hash_integridad": R.huella(datos)})
                cols = list(datos) + list(meta)
                cur = con.execute(f"INSERT INTO operaciones_plano ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                                  list(datos.values()) + list(meta.values()))
                registrar_auditoria(con, SISTEMA, "OPERACIONES", "operaciones_plano", cur.lastrowid, "CREAR", None, None,
                                    {"tramo_id": t["id"], "pendientes": pend},
                                    f"Fila generada automáticamente desde soportes (ejecutado por {usuario.get('nombre')} · {usuario.get('cedula')})")
                res["creadas"] += 1
                continue
            manuales = set(json.loads(previo.get("campos_manuales") or "[]"))
            cambios = {}
            for k, v in datos.items():
                if k in manuales:
                    continue
                av = previo.get(k)
                if str(v if v is not None else "") != str(av if av is not None else ""):
                    if isinstance(v, (int, float)) and R.a_float(av) is not None and abs(float(v) - R.a_float(av)) < 1e-9:
                        continue
                    cambios[k] = v
            for k in ("fecha_local", "hora_local_min", "es_charter", "ficha_id", "notas"):
                if str(meta[k]) != str(previo.get(k)):
                    cambios[k] = meta[k]
            if cambios:
                final = {**{k: previo.get(k) for k in CAMPOS["operaciones_plano"]}, **{k: v for k, v in cambios.items() if k in datos}}
                cambios["pendientes"] = json.dumps(R.pendientes_fila(final))
                cambios["hash_integridad"] = R.huella(final)
                cambios["fecha_actualizacion"] = ahora_txt()
                con.execute(f"UPDATE operaciones_plano SET {','.join(k + '=?' for k in cambios)} WHERE id=?",
                            list(cambios.values()) + [previo["id"]])
                visibles = {k: v for k, v in cambios.items() if k in datos}
                if visibles:
                    registrar_auditoria(con, SISTEMA, "OPERACIONES", "operaciones_plano", previo["id"], "SINCRONIZAR", "varios",
                                        {k: previo.get(k) for k in visibles}, visibles,
                                        f"Actualización automática desde soportes (ejecutado por {usuario.get('nombre')} · {usuario.get('cedula')})")
                    res["actualizadas"] += 1
    return res


def crear_fila_plano_manual(datos: dict, usuario: dict) -> tuple[bool, str]:
    fila = {k: R.normalizar_valor(t, datos.get(k)) for k, t in CAMPOS["operaciones_plano"].items()}
    fecha_local, hora_local = R.parse_fecha(datos.get("fecha_local")), R.parse_hora(datos.get("hora_local"))
    if not fecha_local or not fila.get("matricula"):
        return False, "Indique al menos fecha y matrícula."
    if hora_local is not None and not fila.get("fecha_utc"):
        fila["fecha_utc"], fila["hora_utc"] = R.a_utc(fecha_local, hora_local, get_param("offset_utc_horas", 5))
    fila["fecha_mmaaaa"] = fila.get("fecha_mmaaaa") or fecha_local.strftime("%m%Y")
    with conexion(True) as con:
        if not _periodo_editable_con(con, "OPERACIONES", fecha_local.isoformat()):
            return False, f"El {fecha_local} está cerrado para OPERACIONES."
        meta = _meta_creacion(usuario, "OPERACIONES")
        meta.update({"fuente": "MANUAL", "fecha_local": fecha_local.isoformat(), "hora_local_min": hora_local,
                     "es_charter": 1 if fila.get("tipo_vuelo") == "C" else 0,
                     "campos_manuales": json.dumps(sorted(k for k, v in fila.items() if v is not None)),
                     "pendientes": json.dumps(R.pendientes_fila(fila)), "notas": "Vuelo registrado manualmente (contingencia)",
                     "hash_integridad": R.huella(fila)})
        cols = list(fila) + list(meta)
        cur = con.execute(f"INSERT INTO operaciones_plano ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                          list(fila.values()) + list(meta.values()))
        _asegurar_aeronave(con, fila["matricula"], usuario, fila.get("tipo_equipo"))
        registrar_auditoria(con, usuario, "OPERACIONES", "operaciones_plano", cur.lastrowid, "CREAR", None, None, fila,
                            "Vuelo manual (contingencia) con cláusula de responsabilidad")
    return True, "Vuelo agregado al archivo plano."


def plano_ordenado(desde, hasta) -> pd.DataFrame:
    """Filas del plano con NRO consecutivo (orden fecha → matrícula → hora)."""
    filas = R.ordenar_plano(_lista(df_plano(desde, hasta)))
    return pd.DataFrame(filas) if filas else pd.DataFrame(columns=["id", "nro"] + R.PLANO_CLAVES)


# ============================================================
# Liquidación (datos)
# ============================================================
def datos_liquidacion(desde, hasta, periodo_ajustes: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (liquidación por tercero, detalle de tramos con TIEMPO PILOTOS efectivo)."""
    p = params_dict()
    cfg = cfg_tiempo_pilotos(p)
    tramos = _lista(df_tramos(desde, hasta))
    asign, det = [], []
    for t in tramos:
        tp, origen = R.tiempo_pilotos_efectivo(t, cfg)
        det.append({"TRAMO_ID": t["id"], "LIBRO_ID": t["libro_id"], "FECHA": t["fecha"], "MATRICULA": t["matricula"],
                    "LIBRO Nº": t.get("numero_libro"), "VUELO": t.get("numero_vuelo"),
                    "RUTA": f"{t.get('origen') or ''}-{t.get('destino') or ''}", "PILOTO": t.get("piloto_nombre"),
                    "CEDULA PILOTO": t.get("piloto_cedula"), "COPILOTO": t.get("copiloto_nombre"),
                    "CEDULA COPILOTO": t.get("copiloto_cedula"), "TIEMPO VUELO": R.fmt_duracion(t.get("tiempo_vuelo_min")),
                    "TIEMPO TOTAL": R.fmt_duracion(t.get("tiempo_total_min")),
                    "TIEMPO PILOTOS": R.fmt_duracion(tp) if tp is not None else None, "ORIGEN T. PILOTOS": origen,
                    "_tp": tp, "_tv": t.get("tiempo_vuelo_min"), "_tt": t.get("tiempo_total_min")})
        if R.texto(t.get("piloto_nombre")) or R.texto(t.get("piloto_cedula")):
            asign.append({"cedula": t.get("piloto_cedula"), "nombre": t.get("piloto_nombre"), "rol_vuelo": "PTL", "minutos": tp})
        if R.texto(t.get("copiloto_nombre")) or R.texto(t.get("copiloto_cedula")):
            asign.append({"cedula": t.get("copiloto_cedula"), "nombre": t.get("copiloto_nombre"), "rol_vuelo": "PCA", "minutos": tp})
    params = {"min_smlv": p.get("min_smlv"), "tarifa_ptl": p.get("tarifa_hora_ptl"), "tarifa_pca": p.get("tarifa_hora_pca"),
              "formula_total": p.get("liq_formula_total") if p.get("liq_formula_total") is not None else ""}
    periodo = periodo_ajustes or (R.normalizar_valor("fecha", desde) or "")[:7]
    filas = R.liquidar_tripulacion(asign, params, ajustes_liquidacion(periodo), mapa_categorias_pilotos())
    liq = pd.DataFrame(filas, columns=["CLAVE"] + R.LIQ_COLUMNAS)
    return liq, pd.DataFrame(det)


# ============================================================
# Cierres y consolidados en cascada
# ============================================================
def _periodo_editable_con(con, area, fecha) -> bool:
    f = R.normalizar_valor("fecha", fecha)
    if not f or area not in AREAS_CIERRE:
        return True
    if con.execute("SELECT 1 FROM consolidados_mensuales WHERE area=? AND periodo=? AND estado='CERRADO_MENSUAL'",
                   (area, f[:7])).fetchone():
        return False
    return not con.execute("SELECT 1 FROM consolidados_diarios WHERE area=? AND fecha=? AND estado='CERRADO_DIARIO'",
                           (area, f)).fetchone()


def periodo_editable(area, fecha) -> bool:
    with conexion() as con:
        return _periodo_editable_con(con, area, fecha)


def periodo_mensual_cerrado(area, periodo) -> bool:
    return bool(consulta("SELECT 1 FROM consolidados_mensuales WHERE area=? AND periodo=? AND estado='CERRADO_MENSUAL'",
                         (area, periodo)))


def estado_dia(area, fecha) -> str:
    f = R.normalizar_valor("fecha", fecha)
    if not f:
        return "ABIERTO"
    if periodo_mensual_cerrado(area, f[:7]):
        return "CERRADO_MENSUAL"
    r = consulta("SELECT estado FROM consolidados_diarios WHERE area=? AND fecha=?", (area, f))
    return "CERRADO_DIARIO" if r and r[0]["estado"] == "CERRADO_DIARIO" else "ABIERTO"


def dias_con_datos(area, periodo) -> list[str]:
    ini = f"{periodo}-01"
    fin = f"{periodo}-31"
    if area == "CONTABILIDAD":
        q = ("SELECT DISTINCT fecha_consumo d FROM combustible_tiquetes WHERE eliminado=0 AND fecha_consumo BETWEEN ? AND ? "
             "UNION SELECT DISTINCT fecha FROM libros_vuelo WHERE eliminado=0 AND fecha BETWEEN ? AND ?")
        r = consulta(q, (ini, fin, ini, fin))
    else:
        q = ("SELECT DISTINCT fecha_local d FROM operaciones_plano WHERE eliminado=0 AND fecha_local BETWEEN ? AND ? "
             "UNION SELECT DISTINCT fecha FROM fichas_vuelo WHERE eliminado=0 AND fecha BETWEEN ? AND ?")
        r = consulta(q, (ini, fin, ini, fin))
    return sorted({x["d"] for x in r if x["d"]})


def vista_tiquetes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["ID", "FECHA", "HORA", "ESTACION", "MATRICULA", "RECIBO", "BOLETO", "PRODUCTO", "GALONES",
                                     "VALOR UNIT", "VALOR TOTAL", "PILOTO", "OBSERVACION", "CARGADO POR", "ORIGEN CARGA", "SOPORTE"])
    return pd.DataFrame({
        "ID": df["id"], "FECHA": df["fecha_consumo"], "HORA": df["hora_min"].map(R.fmt_hora), "ESTACION": df["estacion"],
        "MATRICULA": df["matricula"], "RECIBO": df["recibo"], "BOLETO": df["numero_boleto"], "PRODUCTO": df["producto"],
        "GALONES": pd.to_numeric(df["galones"], errors="coerce"), "VALOR UNIT": pd.to_numeric(df["valor_unit"], errors="coerce"),
        "VALOR TOTAL": pd.to_numeric(df["valor_total"], errors="coerce"), "PILOTO": df["piloto_nombre"],
        "OBSERVACION": df["observacion"], "CARGADO POR": df["creado_por_nombre"], "ORIGEN CARGA": df["origen_carga"],
        "SOPORTE": df["ruta_archivo"].fillna("").map(lambda s: Path(s).name if s else "")})


def vista_tramos(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["ID", "FECHA", "MATRICULA", "LIBRO Nº", "VUELO", "ORIGEN", "DESTINO", "GAL TANQ.", "RECIBO COMB.", "LBS SALIDA",
            "PAX TOTAL", "PAX TRÁNSITO", "H. SALIDA", "H. DESPEGUE", "H. ATERRIZAJE", "H. LLEGADA", "T. VUELO", "T. NOCTURNO",
            "T. INSTRUM.", "T. TOTAL", "LBS LLEGADA", "LBS USADO", "TIEMPO PILOTOS", "PILOTO", "COPILOTO", "SOPORTE"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame({
        "ID": df["id"], "FECHA": df["fecha"], "MATRICULA": df["matricula"], "LIBRO Nº": df["numero_libro"],
        "VUELO": df["numero_vuelo"], "ORIGEN": df["origen"], "DESTINO": df["destino"],
        "GAL TANQ.": pd.to_numeric(df["galones_tanqueados"], errors="coerce"), "RECIBO COMB.": df["recibo_combustible"],
        "LBS SALIDA": pd.to_numeric(df["lbs_comb_salida"], errors="coerce"), "PAX TOTAL": pd.to_numeric(df["pax_total"], errors="coerce"),
        "PAX TRÁNSITO": pd.to_numeric(df["pax_transito"], errors="coerce"),
        "H. SALIDA": df["hora_salida_min"].map(R.fmt_hora), "H. DESPEGUE": df["hora_despegue_min"].map(R.fmt_hora),
        "H. ATERRIZAJE": df["hora_aterrizaje_min"].map(R.fmt_hora), "H. LLEGADA": df["hora_llegada_min"].map(R.fmt_hora),
        "T. VUELO": df["tiempo_vuelo_min"].map(R.fmt_duracion), "T. NOCTURNO": df["tiempo_nocturno_min"].map(R.fmt_duracion),
        "T. INSTRUM.": df["tiempo_instrumentos_min"].map(R.fmt_duracion), "T. TOTAL": df["tiempo_total_min"].map(R.fmt_duracion),
        "LBS LLEGADA": pd.to_numeric(df["lbs_comb_llegada"], errors="coerce"),
        "LBS USADO": pd.to_numeric(df["lbs_comb_usado"], errors="coerce"),
        "TIEMPO PILOTOS": df["tiempo_pilotos_min"].map(R.fmt_duracion), "PILOTO": df["piloto_nombre"],
        "COPILOTO": df["copiloto_nombre"], "SOPORTE": df["ruta_archivo"].fillna("").map(lambda s: Path(s).name if s else "")})


def vista_fichas(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["ID", "FECHA", "MATRICULA", "VUELO", "ORIGEN", "DESTINO", "HORA", "PAX TOTAL", "TRIPADIS", "BEBÉS", "PAX TRÁNSITO",
            "CARGA A BORDO (KG)", "CARGA TRÁNSITO (KG)", "COMB. BÁSICO (LBS)", "PILOTO", "OBSERVACIONES", "SOPORTE"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame({
        "ID": df["id"], "FECHA": df["fecha"], "MATRICULA": df["matricula"], "VUELO": df["numero_vuelo"], "ORIGEN": df["origen"],
        "DESTINO": df["destino"], "HORA": df["hora_min"].map(R.fmt_hora), "PAX TOTAL": pd.to_numeric(df["pax_total"], errors="coerce"),
        "TRIPADIS": pd.to_numeric(df["tripadis"], errors="coerce"), "BEBÉS": pd.to_numeric(df["bebes"], errors="coerce"),
        "PAX TRÁNSITO": pd.to_numeric(df["pax_transito"], errors="coerce"),
        "CARGA A BORDO (KG)": pd.to_numeric(df["carga_a_bordo_kg"], errors="coerce"),
        "CARGA TRÁNSITO (KG)": pd.to_numeric(df["carga_transito_kg"], errors="coerce"),
        "COMB. BÁSICO (LBS)": pd.to_numeric(df["combustible_basico_lbs"], errors="coerce"), "PILOTO": df["piloto_nombre"],
        "OBSERVACIONES": df["observaciones"], "SOPORTE": df["ruta_archivo"].fillna("").map(lambda s: Path(s).name if s else "")})


def vista_plano(df: pd.DataFrame, con_id=True) -> pd.DataFrame:
    """Plano con encabezados oficiales (A..AA). df debe venir de plano_ordenado()."""
    columnas = (["ID"] if con_id else []) + [R.PLANO_HEADERS[c] for c in R.PLANO_CLAVES]
    if df.empty:
        return pd.DataFrame(columns=columnas)
    out = pd.DataFrame({R.PLANO_HEADERS[c]: df[c] if c in df else None for c in R.PLANO_CLAVES})
    if con_id:
        out.insert(0, "ID", df["id"].values)
    return out


def snapshot_area(area: str, desde: str, hasta: str, incluir_liquidacion=False) -> dict:
    if area == "CONTABILIDAD":
        hojas = {"Combustible": vista_tiquetes(df_tiquetes(desde, hasta)), "Detalle Vuelos": vista_tramos(df_tramos(desde, hasta))}
        if incluir_liquidacion:
            liq, _ = datos_liquidacion(desde, hasta, desde[:7])
            hojas["Liquidación Pilotos"] = liq.drop(columns=["CLAVE"], errors="ignore")
        return hojas
    return {"Plano MO-A3-013": vista_plano(plano_ordenado(desde, hasta), con_id=False),
            "Fichas de Vuelo": vista_fichas(df_fichas(desde, hasta))}


def _serializar(hojas: dict) -> tuple[bytes, str]:
    paquete = {n: {"columns": [str(c) for c in df.columns],
                   "data": [[None if R.es_vacio(v) else (v.item() if hasattr(v, "item") else v) for v in fila]
                            for fila in df.itertuples(index=False, name=None)]}
               for n, df in hojas.items()}
    crudo = json.dumps(paquete, ensure_ascii=False, default=str).encode("utf-8")
    return gzip.compress(crudo, compresslevel=6), hashlib.sha256(crudo).hexdigest()


def deserializar_snapshot(blob: bytes) -> dict:
    paquete = json.loads(gzip.decompress(blob).decode("utf-8"))
    return {n: pd.DataFrame(v["data"], columns=v["columns"]) for n, v in paquete.items()}


def _resumen_hojas(hojas: dict) -> dict:
    return {n: len(df) for n, df in hojas.items()}


def cerrar_dia(area, fecha, usuario) -> tuple[bool, str]:
    f = R.normalizar_valor("fecha", fecha)
    if area not in AREAS_CIERRE or not f:
        return False, "Área o fecha no válida."
    if estado_dia(area, f) != "ABIERTO":
        return False, f"El {f} ya está cerrado para {area}."
    if area == "OPERACIONES":
        sincronizar_plano(f, f, SISTEMA)
    respaldo_db(f"antes_cierre_diario_{area}_{f}")
    hojas = snapshot_area(area, f, f)
    blob, sha = _serializar(hojas)
    n = sum(len(df) for df in hojas.values())
    with conexion(True) as con:
        con.execute("INSERT INTO consolidados_diarios (area, fecha, estado, n_registros, resumen, snapshot, sha256, cerrado_por_cedula, "
                    "cerrado_por_nombre, fecha_cierre) VALUES (?,?,?,?,?,?,?,?,?,?) ON CONFLICT(area, fecha) DO UPDATE SET "
                    "estado=excluded.estado, n_registros=excluded.n_registros, resumen=excluded.resumen, snapshot=excluded.snapshot, "
                    "sha256=excluded.sha256, cerrado_por_cedula=excluded.cerrado_por_cedula, cerrado_por_nombre=excluded.cerrado_por_nombre, "
                    "fecha_cierre=excluded.fecha_cierre",
                    (area, f, "CERRADO_DIARIO", n, json.dumps(_resumen_hojas(hojas)), blob, sha, usuario.get("cedula"),
                     usuario.get("nombre"), ahora_txt()))
        registrar_auditoria(con, usuario, area, "consolidados_diarios", f"{area}|{f}", "CIERRE_DIARIO", None, "ABIERTO",
                            "CERRADO_DIARIO", f"Cierre diario con cláusula de responsabilidad · {n} registros · sha256 {sha[:16]}")
    return True, f"Día {f} cerrado para {area} ({n} registros)."


def reabrir_dia(area, fecha, usuario, motivo) -> tuple[bool, str]:
    f = R.normalizar_valor("fecha", fecha)
    if usuario.get("rol") != "GERENCIA":
        return False, "Solo Gerencia puede reabrir un día."
    if len(R.texto(motivo)) < 5:
        return False, "Escriba el motivo de la reapertura."
    if periodo_mensual_cerrado(area, f[:7]):
        return False, "El mes está cerrado: reabra primero el mes."
    with conexion(True) as con:
        n = con.execute("UPDATE consolidados_diarios SET estado='REABIERTO', reabierto_por_cedula=?, fecha_reapertura=?, "
                        "motivo_reapertura=? WHERE area=? AND fecha=? AND estado='CERRADO_DIARIO'",
                        (usuario.get("cedula"), ahora_txt(), motivo, area, f)).rowcount
        if n:
            registrar_auditoria(con, usuario, "GERENCIA", "consolidados_diarios", f"{area}|{f}", "EDITAR", "estado",
                                "CERRADO_DIARIO", "REABIERTO", motivo)
    return (True, f"Día {f} reabierto para {area}.") if n else (False, "El día no estaba cerrado.")


def cerrar_mes(area, periodo, usuario) -> tuple[bool, str]:
    if not re.fullmatch(r"\d{4}-\d{2}", periodo or ""):
        return False, "Periodo no válido (AAAA-MM)."
    if periodo_mensual_cerrado(area, periodo):
        return False, f"El mes {periodo} ya está cerrado para {area}."
    abiertos = [d for d in dias_con_datos(area, periodo) if estado_dia(area, d) == "ABIERTO"]
    for d in abiertos:
        ok, msg = cerrar_dia(area, d, usuario)
        if not ok:
            return False, f"No se pudo cerrar el día {d}: {msg}"
    anio, mes = int(periodo[:4]), int(periodo[5:])
    ultimo = (date(anio + (mes == 12), mes % 12 + 1, 1) - timedelta(days=1)).isoformat()
    desde = f"{periodo}-01"
    respaldo_db(f"antes_cierre_mensual_{area}_{periodo}")
    hojas = snapshot_area(area, desde, ultimo, incluir_liquidacion=(area == "CONTABILIDAD"))
    blob, sha = _serializar(hojas)
    n = sum(len(df) for df in hojas.values())
    with conexion(True) as con:
        con.execute("INSERT INTO consolidados_mensuales (area, periodo, estado, n_registros, resumen, snapshot, sha256, "
                    "cerrado_por_cedula, cerrado_por_nombre, fecha_cierre) VALUES (?,?,?,?,?,?,?,?,?,?) "
                    "ON CONFLICT(area, periodo) DO UPDATE SET estado=excluded.estado, n_registros=excluded.n_registros, "
                    "resumen=excluded.resumen, snapshot=excluded.snapshot, sha256=excluded.sha256, "
                    "cerrado_por_cedula=excluded.cerrado_por_cedula, cerrado_por_nombre=excluded.cerrado_por_nombre, "
                    "fecha_cierre=excluded.fecha_cierre",
                    (area, periodo, "CERRADO_MENSUAL", n, json.dumps(_resumen_hojas(hojas)), blob, sha, usuario.get("cedula"),
                     usuario.get("nombre"), ahora_txt()))
        registrar_auditoria(con, usuario, area, "consolidados_mensuales", f"{area}|{periodo}", "CIERRE_MENSUAL", None, "ABIERTO",
                            "CERRADO_MENSUAL", f"Cierre mensual · {len(abiertos)} día(s) cerrados en el proceso · sha256 {sha[:16]}")
    _acumular_anual(area, periodo[:4], usuario)
    extra = f" (se cerraron también {len(abiertos)} día(s) pendientes)" if abiertos else ""
    return True, f"Mes {periodo} cerrado para {area}{extra}."


def reabrir_mes(area, periodo, usuario, motivo) -> tuple[bool, str]:
    if usuario.get("rol") != "GERENCIA":
        return False, "Solo Gerencia puede reabrir un mes."
    if len(R.texto(motivo)) < 5:
        return False, "Escriba el motivo de la reapertura."
    anual = consulta("SELECT estado FROM consolidados_anuales WHERE area=? AND anio=?", (area, periodo[:4]))
    if anual and anual[0]["estado"] == "CERRADO_ANUAL":
        return False, "El año está cerrado: no se puede reabrir el mes."
    with conexion(True) as con:
        n = con.execute("UPDATE consolidados_mensuales SET estado='REABIERTO', reabierto_por_cedula=?, fecha_reapertura=?, "
                        "motivo_reapertura=? WHERE area=? AND periodo=? AND estado='CERRADO_MENSUAL'",
                        (usuario.get("cedula"), ahora_txt(), motivo, area, periodo)).rowcount
        if n:
            registrar_auditoria(con, usuario, "GERENCIA", "consolidados_mensuales", f"{area}|{periodo}", "EDITAR", "estado",
                                "CERRADO_MENSUAL", "REABIERTO", motivo)
    if n:
        _acumular_anual(area, periodo[:4], usuario)
    return (True, f"Mes {periodo} reabierto para {area}. Los días siguen cerrados: reabra los que necesite.") if n else \
        (False, "El mes no estaba cerrado.")


def _acumular_anual(area, anio, usuario):
    meses = consulta("SELECT periodo, snapshot FROM consolidados_mensuales WHERE area=? AND periodo LIKE ? AND "
                     "estado='CERRADO_MENSUAL' ORDER BY periodo", (area, f"{anio}-%"))
    hojas: dict[str, list] = {}
    for m in meses:
        for nombre, df in deserializar_snapshot(m["snapshot"]).items():
            df = df.copy()
            df.insert(0, "PERIODO", m["periodo"])
            hojas.setdefault(nombre, []).append(df)
    combinado = {n: pd.concat(lst, ignore_index=True) for n, lst in hojas.items()}
    blob, sha = _serializar(combinado) if combinado else (None, None)
    n = sum(len(df) for df in combinado.values())
    with conexion(True) as con:
        prev = con.execute("SELECT estado FROM consolidados_anuales WHERE area=? AND anio=?", (area, anio)).fetchone()
        if prev and prev[0] == "CERRADO_ANUAL":
            return
        con.execute("INSERT INTO consolidados_anuales (area, anio, estado, meses_incluidos, n_registros, resumen, snapshot, sha256, "
                    "actualizado_por_cedula, fecha_actualizacion) VALUES (?,?,?,?,?,?,?,?,?,?) ON CONFLICT(area, anio) DO UPDATE SET "
                    "meses_incluidos=excluded.meses_incluidos, n_registros=excluded.n_registros, resumen=excluded.resumen, "
                    "snapshot=excluded.snapshot, sha256=excluded.sha256, actualizado_por_cedula=excluded.actualizado_por_cedula, "
                    "fecha_actualizacion=excluded.fecha_actualizacion",
                    (area, anio, "ACUMULANDO", json.dumps([m["periodo"] for m in meses]), n,
                     json.dumps(_resumen_hojas(combinado)), blob, sha, usuario.get("cedula"), ahora_txt()))
        registrar_auditoria(con, usuario, area, "consolidados_anuales", f"{area}|{anio}", "EDITAR", "meses_incluidos", None,
                            [m["periodo"] for m in meses], "Acumulado anual actualizado")


def cerrar_anio(area, anio, usuario) -> tuple[bool, str]:
    if usuario.get("rol") != "GERENCIA":
        return False, "Solo Gerencia ejecuta el cierre anual."
    r = consulta("SELECT * FROM consolidados_anuales WHERE area=? AND anio=?", (area, str(anio)))
    if not r:
        return False, "Aún no hay meses cerrados en ese año."
    if r[0]["estado"] == "CERRADO_ANUAL":
        return False, "El año ya está cerrado."
    with conexion(True) as con:
        con.execute("UPDATE consolidados_anuales SET estado='CERRADO_ANUAL', cerrado_por_cedula=?, fecha_cierre=? WHERE area=? AND anio=?",
                    (usuario.get("cedula"), ahora_txt(), area, str(anio)))
        registrar_auditoria(con, usuario, "GERENCIA", "consolidados_anuales", f"{area}|{anio}", "CIERRE_ANUAL", None,
                            "ACUMULANDO", "CERRADO_ANUAL", f"Cierre anual · meses {r[0]['meses_incluidos']}")
    return True, f"Año {anio} cerrado para {area}."


def obtener_consolidado(nivel, area, periodo) -> tuple[dict | None, dict | None]:
    tabla, col = {"DIARIO": ("consolidados_diarios", "fecha"), "MENSUAL": ("consolidados_mensuales", "periodo"),
                  "ANUAL": ("consolidados_anuales", "anio")}[nivel]
    r = consulta(f"SELECT * FROM {tabla} WHERE area=? AND {col}=?", (area, str(periodo)))
    if not r or not r[0].get("snapshot"):
        return None, (r[0] if r else None)
    meta = {k: v for k, v in r[0].items() if k != "snapshot"}
    blob = r[0]["snapshot"]
    hojas = deserializar_snapshot(blob)
    crudo = gzip.decompress(blob)
    meta["integro"] = hashlib.sha256(crudo).hexdigest() == r[0]["sha256"]
    return hojas, meta


def df_consolidados(nivel, area=None) -> pd.DataFrame:
    tabla, col = {"DIARIO": ("consolidados_diarios", "fecha"), "MENSUAL": ("consolidados_mensuales", "periodo"),
                  "ANUAL": ("consolidados_anuales", "anio")}[nivel]
    q = f"SELECT * FROM {tabla}" + (" WHERE area=?" if area else "") + f" ORDER BY {col} DESC"
    df = df_sql(q, (area,) if area else ())
    return df.drop(columns=["snapshot"], errors="ignore")


# ============================================================
# Hallazgos de auditoría
# ============================================================
COLS_HALLAZGO = ["SEVERIDAD", "AREA", "CATEGORIA", "TABLA", "REGISTRO_ID", "FECHA", "MATRICULA", "DETALLE",
                 "ACCION_SUGERIDA", "RESPONSABLE_CEDULA", "RESPONSABLE_NOMBRE", "ULTIMO_EDITOR_CEDULA", "ULTIMO_EDITOR_NOMBRE"]


def guardar_auditoria(ambito, desde, hasta, hallazgos: pd.DataFrame, resumen, uso_ia, usuario) -> int:
    with conexion(True) as con:
        cur = con.execute("INSERT INTO auditorias_ia (fecha, ambito, desde, hasta, ejecutada_por_cedula, ejecutada_por_nombre, "
                          "n_hallazgos, resumen, uso_ia) VALUES (?,?,?,?,?,?,?,?,?)",
                          (ahora_txt(), ambito, str(desde), str(hasta), usuario.get("cedula"), usuario.get("nombre"),
                           len(hallazgos), R.texto(resumen)[:20000], 1 if uso_ia else 0))
        aid = cur.lastrowid
        for h in hallazgos.to_dict("records"):
            con.execute("INSERT INTO hallazgos (auditoria_id, severidad, area, categoria, tabla, registro_id, fecha_ref, matricula, "
                        "detalle, accion_sugerida, responsable_cedula, responsable_nombre, ultimo_editor_cedula, ultimo_editor_nombre) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        [aid] + [None if R.es_vacio(h.get(c)) else str(h.get(c)) for c in COLS_HALLAZGO])
        registrar_auditoria(con, usuario, ambito, "auditorias_ia", aid, "AUDITORIA_IA", None, None,
                            {"hallazgos": len(hallazgos), "desde": str(desde), "hasta": str(hasta), "uso_ia": bool(uso_ia)},
                            f"Auditoría {ambito}")
    return aid


def df_auditorias(ambito=None) -> pd.DataFrame:
    q = "SELECT * FROM auditorias_ia" + (" WHERE ambito=?" if ambito else "") + " ORDER BY id DESC"
    return df_sql(q, (ambito,) if ambito else ())


def df_hallazgos_recientes(desde=None, hasta=None) -> pd.DataFrame:
    """Hallazgos de la auditoría más reciente de cada ámbito (sin duplicar el mismo registro/categoría)."""
    ids = [r["id"] for r in consulta("SELECT MAX(id) id FROM auditorias_ia GROUP BY ambito") if r["id"]]
    if not ids:
        return pd.DataFrame(columns=["auditoria_id"] + [c.lower() for c in COLS_HALLAZGO])
    q = f"SELECT * FROM hallazgos WHERE auditoria_id IN ({','.join('?' * len(ids))})"
    df = df_sql(q, tuple(ids))
    if desde or hasta:
        d, h = _rango(desde, hasta)
        df = df[df["fecha_ref"].fillna("").between(d, h) | df["fecha_ref"].isna()]
    return df.drop_duplicates(subset=["categoria", "tabla", "registro_id"])


# ============================================================
# Respaldos
# ============================================================
def respaldo_db(motivo="manual") -> Path:
    RESPALDOS_DIR.mkdir(parents=True, exist_ok=True)
    seguro = re.sub(r"[^A-Za-z0-9_-]", "_", motivo)[:60]
    destino = RESPALDOS_DIR / f"moon_{ahora():%Y%m%d_%H%M%S}_{seguro}.db"
    src = sqlite3.connect(DB_PATH, timeout=30)
    dst = sqlite3.connect(destino)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    retener = int(get_param("respaldo_retencion", 60) or 60)
    archivos = sorted(RESPALDOS_DIR.glob("moon_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    for viejo in archivos[retener:]:
        try:
            viejo.unlink()
        except OSError:
            pass
    return destino


def respaldo_diario_si_corresponde():
    if get_meta("ultimo_respaldo_diario") != hoy().isoformat():
        respaldo_db("diario")
        set_meta("ultimo_respaldo_diario", hoy().isoformat())


def listar_respaldos() -> list[dict]:
    return [{"archivo": p.name, "tamano_kb": round(p.stat().st_size / 1024, 1),
             "fecha": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")}
            for p in sorted(RESPALDOS_DIR.glob("moon_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)]


def bytes_respaldo_actual() -> bytes:
    p = respaldo_db("descarga_gerencia")
    return p.read_bytes()


# ============================================================
# Datasets analíticos (Gerencia)
# ============================================================
def dataset_vuelos(desde=None, hasta=None) -> pd.DataFrame:
    """Un registro por vuelo (fila del plano) enriquecido con libro, tripulación, flota y libras usadas."""
    d, h = _rango(desde, hasta)
    df = df_sql(
        "SELECT p.id, p.fecha_local, p.hora_local_min, p.matricula, p.tipo_equipo, p.nro_vuelo, p.origen, p.destino, p.tipo_vuelo, "
        "p.pax_bordo, p.pax_transito, p.sillas_ofrecidas, p.carga_ofrecida, p.carga_bordo, p.carga_real, p.combustible_consumido, "
        "p.horas_bloque, p.horas_vuelo, p.fuente, t.lbs_comb_usado, t.tiempo_vuelo_min, t.tiempo_total_min, "
        "l.piloto_nombre, l.piloto_cedula, l.copiloto_nombre, l.copiloto_cedula, a.modelo_codigo, m.factor_conversion_lbs_gal "
        "FROM operaciones_plano p LEFT JOIN tramos_vuelo t ON t.id=p.tramo_id LEFT JOIN libros_vuelo l ON l.id=t.libro_id "
        "LEFT JOIN aeronaves a ON a.matricula = ('HK-' || SUBSTR(p.matricula, 3)) "
        "LEFT JOIN modelos_aeronave m ON m.codigo = COALESCE(a.modelo_codigo, p.tipo_equipo) "
        "WHERE p.eliminado=0 AND COALESCE(p.fecha_local,'') BETWEEN ? AND ?", (d, h))
    if df.empty:
        return df
    df["FECHA"] = pd.to_datetime(df["fecha_local"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()
    df["MES"] = df["FECHA"].dt.strftime("%Y-%m")
    df["ANIO"] = df["FECHA"].dt.year.astype(str)
    df["MES_NUM"] = df["FECHA"].dt.month
    df["MATRICULA"] = df["matricula"].fillna("").map(lambda m: R.norm_matricula(m) or "SIN DATO")
    df["MODELO"] = df["modelo_codigo"].fillna(df["tipo_equipo"]).fillna("SIN MODELO")
    df["TIPO_VUELO"] = df["tipo_vuelo"].fillna("SIN TIPO")
    df["ORIGEN"] = df["origen"].fillna("SIN DATO")
    df["RUTA"] = df["origen"].fillna("?") + "-" + df["destino"].fillna("?")
    df["NRO_VUELO"] = df["nro_vuelo"].fillna("")
    df["PILOTO"] = df["piloto_nombre"].fillna("").map(lambda s: R.texto(s).title() or "SIN DATO")
    df["COPILOTO"] = df["copiloto_nombre"].fillna("").map(lambda s: R.texto(s).title() or "SIN DATO")
    num = lambda c: pd.to_numeric(df[c], errors="coerce")
    df["PAX"] = num("pax_bordo").fillna(0)
    df["SILLAS"] = num("sillas_ofrecidas")
    df["CARGA_BORDO"] = num("carga_bordo").fillna(0)
    df["CARGA_OFRECIDA"] = num("carga_ofrecida")
    df["GAL_ESTANDAR"] = num("combustible_consumido")
    factor = num("factor_conversion_lbs_gal")
    lbs = num("lbs_comb_usado")
    df["GAL_REAL"] = (lbs / factor).where((factor > 0) & lbs.notna())
    df["MIN_VUELO"] = num("tiempo_vuelo_min").fillna(df["horas_vuelo"].map(R.parse_duracion))
    df["MIN_BLOQUE"] = num("tiempo_total_min").fillna(df["horas_bloque"].map(R.parse_duracion))
    df["MIN_VUELO"] = pd.to_numeric(df["MIN_VUELO"], errors="coerce").fillna(0)
    df["MIN_BLOQUE"] = pd.to_numeric(df["MIN_BLOQUE"], errors="coerce").fillna(0)
    df["GAL"] = df["GAL_REAL"].fillna(df["GAL_ESTANDAR"])
    df["VUELOS"] = 1
    return df


def dataset_combustible(desde=None, hasta=None) -> pd.DataFrame:
    df = df_tiquetes(desde, hasta)
    if df.empty:
        return df
    df["FECHA"] = pd.to_datetime(df["fecha_consumo"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()
    df["MES"] = df["FECHA"].dt.strftime("%Y-%m")
    df["ANIO"] = df["FECHA"].dt.year.astype(str)
    df["MES_NUM"] = df["FECHA"].dt.month
    df["ESTACION"] = df["estacion"].fillna("SIN DATO")
    df["MATRICULA"] = df["matricula"].fillna("SIN DATO")
    df["PILOTO"] = df["piloto_nombre"].fillna("").map(lambda s: R.texto(s).title() or "SIN DATO")
    df["GALONES"] = pd.to_numeric(df["galones"], errors="coerce").fillna(0)
    df["VALOR_TOTAL"] = pd.to_numeric(df["valor_total"], errors="coerce")
    df["VALOR_UNIT"] = pd.to_numeric(df["valor_unit"], errors="coerce")
    return df

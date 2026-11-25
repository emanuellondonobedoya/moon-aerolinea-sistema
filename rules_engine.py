# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Motor de reglas (rules_engine.py)

Lógica pura (sin base de datos ni interfaz) para:
  * Normalizadores tolerantes: matrículas, horas, duraciones, fechas, decimales.
  * Evaluador seguro de fórmulas (AST controlado, sin eval/exec).
  * Regla TIEMPO PILOTOS (FORMULA / RANGOS / MANUAL) y liquidación de horas y pesos.
  * Archivo plano Aerocivil MO-A3-013 (27 columnas A..AA), UTC, chárter, depuración de pasajeros.
  * Validaciones matemáticas del Libro de Vuelo MO-A3-004 y conciliación de combustible.

Regla de oro: si un parámetro no existe (None/blanco) NUNCA se lanza excepción; se devuelve None
y el campo queda pendiente para diligenciamiento manual.
"""
from __future__ import annotations

import ast
import hashlib
import json
import math
import operator
import re
import unicodedata
from datetime import date, datetime, time, timedelta

# ============================================================
# Constantes
# ============================================================
OFFSET_UTC_COLOMBIA_H = 5
LITROS_POR_GALON = 3.78541
PRODUCTOS = ["JET A1", "AVGAS 100LL", "OTRO"]
TIPOS_VUELO = {"R": "Regular", "A": "Adicional", "C": "Chárter"}
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio",
            8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
_MESES_TXT = {"ene": 1, "jan": 1, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "may": 5, "jun": 6, "jul": 7,
              "ago": 8, "aug": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12, "dec": 12}

# (letra, clave interna, encabezado oficial) — hoja MES de la plantilla MO-A3-013
PLANO_COLUMNAS = [
    ("A", "tip", "TIP"), ("B", "nro", "NRO"), ("C", "sigla", "SIGLA"), ("D", "fecha_mmaaaa", "FECHA"),
    ("E", "fecha_utc", "FECHA OPERA. UTC"), ("F", "hora_utc", "HORA OPERA. UTC"), ("G", "origen", "ORIGEN"),
    ("H", "destino", "DESTINO"), ("I", "matricula", "MATRÍCULA"), ("J", "tipo_equipo", "TIPO DE EQUIPO"),
    ("K", "nro_vuelo", "NRO VUELO"), ("L", "sillas_ofrecidas", "SILLA OFRECIDAS"),
    ("M", "carga_ofrecida", "CARGA OFRECIDA"), ("N", "pax_bordo", "PASAJEROS A BORDO"),
    ("O", "carga_bordo", "CARGA A BORDO"), ("P", "correo_bordo", "CORREO A BORDO"),
    ("Q", "pax_transito", "PASAJEROS TRÁNSIT."), ("R", "carga_transito", "CARGA TRÁNSIT."),
    ("S", "horas_bloque", "HORAS BLOQUE"), ("T", "tipo_vuelo", "TIPO VUELO"), ("U", "tipo_trafico", "TIPO TRÁFICO"),
    ("V", "uso_aeronave", "USO AERONAVE"), ("W", "tipo_combustible", "TIPO COMBUSTIBLE"),
    ("X", "metodo_medicion", "MÉTODO MEDICIÓN COMBUSTIBLE"), ("Y", "combustible_consumido", "COMBUSTIBLE CONSUMIDO"),
    ("Z", "horas_vuelo", "HORAS VUELO"), ("AA", "carga_real", "CARGA REAL"),
]
PLANO_CLAVES = [c for _, c, _ in PLANO_COLUMNAS]
PLANO_HEADERS = {c: h for _, c, h in PLANO_COLUMNAS}
PLANO_LETRAS = {c: l for l, c, _ in PLANO_COLUMNAS}
PLANO_OFICIALES = PLANO_CLAVES[:19]          # A..S  → archivo limpio Aerocivil
PLANO_INTERNAS = PLANO_CLAVES[19:]           # T..AA → control interno (rojo en la plantilla)
PLANO_TEXTO = {"tip", "sigla", "fecha_mmaaaa", "fecha_utc", "hora_utc", "origen", "destino", "matricula",
               "tipo_equipo", "nro_vuelo", "horas_bloque", "tipo_vuelo", "tipo_trafico", "uso_aeronave",
               "tipo_combustible", "metodo_medicion", "horas_vuelo"}
PLANO_NUMERICAS = [c for c in PLANO_CLAVES if c not in PLANO_TEXTO and c != "nro"]
# Celdas que el despacho diligencia a mano (grupo de despacho): nunca se bloquean.
PLANO_SIEMPRE_EDITABLES = {"carga_bordo", "carga_transito", "carga_real"}

VARS_TIEMPO_PILOTOS = ("tiempo_vuelo", "tiempo_total", "tiempo_nocturno", "tiempo_instrumentos", "ajuste_min", "pax_total")
VARS_LIQUIDACION = ("min_piloto", "min_copiloto", "min_volados", "min_smlv", "min_excedentes", "valor_hora",
                    "valor_minuto", "tarifa_ptl", "tarifa_pca", "horas_excedentes")
FORMULA_LIQUIDACION_DEFECTO = "min_excedentes * valor_minuto"

LIQ_COLUMNAS = ["CEDULA", "NOMBRE", "CATEGORIA", "MIN PILOTO", "MIN COPILOTO", "MIN VOLADOS", "MIN SMLV",
                "TOTAL MIN EXCEDENTES", "VALOR HORA", "VALOR MINUTO", "HORAS", "MIN", "TOTAL A PAGAR (COP)",
                "TRAMOS", "TRAMOS SIN TIEMPO PILOTOS", "AJUSTES MANUALES"]
LIQ_EDITABLES = {"CATEGORIA": "categoria", "MIN PILOTO": "min_piloto", "MIN COPILOTO": "min_copiloto",
                 "MIN SMLV": "min_smlv", "VALOR HORA": "valor_hora", "TOTAL A PAGAR (COP)": "total_pagar"}


# ============================================================
# Utilidades numéricas y de texto
# ============================================================
def es_vacio(v) -> bool:
    """True para None, NaN, NaT, pd.NA, cadenas vacías o guiones de formulario."""
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip().lower() in ("", "-", "--", "—", "n/a", "na", "nan", "none", "nat", "null", "<na>")
    if isinstance(v, (list, tuple, dict, set)):
        return len(v) == 0
    try:
        if isinstance(v, float) and math.isnan(v):
            return True
        import pandas as pd  # import local: el motor no depende de pandas para funcionar
        r = pd.isna(v)
        if isinstance(r, bool):
            return r
        if hasattr(r, "ndim") and r.ndim == 0:
            return bool(r)
        return False
    except Exception:
        return False


def safe_div(num, den, default=0.0):
    """División protegida: nunca lanza ZeroDivisionError ni devuelve NaN/inf."""
    try:
        n, d = float(num), float(den)
        if d == 0 or math.isnan(n) or math.isnan(d):
            return default
        r = n / d
        return r if math.isfinite(r) else default
    except (TypeError, ValueError):
        return default


def a_float(v, default=None):
    """Convierte '1.234,5', '1,234.5', '320,5', '320 gal', 12, None → float (o default)."""
    if es_vacio(v):
        return default
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        f = float(v)
        return f if math.isfinite(f) else default
    if hasattr(v, "item"):
        try:
            return a_float(v.item(), default)
        except Exception:
            return default
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if not s or s in ("-", ".", ","):
        return default
    neg = s.startswith("-")
    s = s.replace("-", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif "." in s and (s.count(".") > 1 or _parece_miles(s)):
        s = s.replace(".", "")
    try:
        f = float(s)
    except ValueError:
        return default
    f = -f if neg else f
    return f if math.isfinite(f) else default


def _parece_miles(s: str) -> bool:
    """'19.500' o '1.600' se interpretan como miles (formato colombiano); '320.5' o '6.70' no."""
    partes = s.split(".")
    return (len(partes) >= 2 and 1 <= len(partes[0]) <= 3 and not partes[0].startswith("0")
            and all(len(p) == 3 for p in partes[1:]))


def a_int(v, default=None):
    f = a_float(v, None)
    return default if f is None else int(round(f))


def texto(v, default="") -> str:
    return default if es_vacio(v) else str(v).strip()


def sin_tildes(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(s or "")) if not unicodedata.combining(c))


def nombre_clave(s) -> str:
    """Clave para comparar nombres de personas sin tildes, espacios dobles ni mayúsculas."""
    return re.sub(r"\s+", " ", sin_tildes(texto(s)).upper()).strip()


def solo_digitos(s) -> str:
    return re.sub(r"\D", "", texto(s))


def huella(datos: dict) -> str:
    """SHA-256 canónico (claves ordenadas, sin espacios) para hash_integridad."""
    def _norm(v):
        if es_vacio(v):
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, float):
            # 5.0 y 5 deben producir la misma huella (la base y pandas pueden devolver uno u otro)
            return int(v) if v.is_integer() else round(v, 6)
        if isinstance(v, (date, datetime, time)):
            return v.isoformat()
        if hasattr(v, "item"):
            try:
                return _norm(v.item())
            except Exception:
                return str(v)
        return v
    canon = json.dumps({k: _norm(datos[k]) for k in sorted(datos)}, ensure_ascii=False, sort_keys=True,
                       separators=(",", ":"), default=str)
    return hashlib.sha256(canon.encode("utf-8")).hexdigest()


# ============================================================
# Matrículas, aeropuertos y números de vuelo
# ============================================================
def norm_matricula(m) -> str:
    """'hk4350', 'HK 4350', 'HK-4350', '4350' → 'HK-4350'."""
    s = re.sub(r"[\s_.]+", "", texto(m).upper())
    if not s:
        return ""
    r = re.match(r"^HK-?([A-Z0-9]+)$", s)
    if r:
        return f"HK-{r.group(1)}"
    if re.fullmatch(r"\d{3,4}[A-Z]?", s):
        return f"HK-{s}"
    return s


def matricula_compacta(m) -> str:
    return norm_matricula(m).replace("-", "")


def digitos_matricula(m) -> str:
    return solo_digitos(norm_matricula(m))[:4]


def norm_iata(c) -> str:
    s = re.sub(r"[^A-Z]", "", sin_tildes(texto(c)).upper())
    return s[:4]


def separar_ruta(ruta) -> tuple[str, str]:
    partes = [norm_iata(p) for p in re.split(r"[-–/>\s]+", texto(ruta)) if norm_iata(p)]
    if len(partes) >= 2:
        return partes[0], partes[-1]
    return (partes[0], "") if partes else ("", "")


def normalizar_nro_vuelo(nro, matricula) -> tuple[str | None, bool]:
    """
    Devuelve (número para el plano, es_chárter).
    Chárter = el número de vuelo corresponde a los 4 dígitos de la matrícula (4947, o seriado 494700/494701).
    """
    s = solo_digitos(nro)
    if not s:
        return None, False
    dig = digitos_matricula(matricula)
    if dig and (s == dig or (len(s) == len(dig) + 2 and s.startswith(dig))):
        return dig, True
    return s, False


def norm_producto(p) -> str:
    s = texto(p).upper().replace("-", " ")
    if "JET" in s:
        return "JET A1"
    if "AVGAS" in s or "100LL" in s.replace(" ", ""):
        return "AVGAS 100LL"
    return "OTRO" if s else "JET A1"


# ============================================================
# Horas, duraciones y fechas
# ============================================================
def parse_hora(v) -> int | None:
    """Hora del día → minutos desde medianoche. Acepta HH:MM, H:MM, HHMM, HMM, '5h55', time, fracción Excel."""
    if es_vacio(v):
        return None
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, timedelta):
        m = int(round(v.total_seconds() / 60))
        return m if 0 <= m < 1440 else None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        f = float(v)
        if not math.isfinite(f) or f < 0:
            return None
        if 0 <= f < 1 and not float(f).is_integer():
            return int(round(f * 1440)) % 1440
        n = int(round(f))
        if n < 24:
            return n * 60
        h, m = divmod(n, 100)
        return h * 60 + m if h < 24 and m < 60 else None
    s = texto(v).lower().replace(" ", "")
    pm = s.endswith("pm") or s.endswith("p.m.")
    am = s.endswith("am") or s.endswith("a.m.")
    s = re.sub(r"(a\.?m\.?|p\.?m\.?|hrs?|horas?)$", "", s)
    s = s.replace("h", ":").replace(".", ":").replace(";", ":").replace(",", ":")
    m = re.fullmatch(r"(\d{1,2}):(\d{1,2})(?::\d{1,2})?:?", s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
    elif re.fullmatch(r"\d{3,4}", s):
        h, mi = divmod(int(s), 100)
    elif re.fullmatch(r"\d{1,2}", s):
        h, mi = int(s), 0
    else:
        return None
    if pm and h < 12:
        h += 12
    if am and h == 12:
        h = 0
    return h * 60 + mi if 0 <= h < 24 and 0 <= mi < 60 else None


def parse_duracion(v) -> int | None:
    """Duración → minutos. '1:55'→115, '0:26'→26, '0155'→115, '34'→34, '1h55'→115, timedelta, time."""
    if es_vacio(v):
        return None
    if isinstance(v, timedelta):
        m = int(round(v.total_seconds() / 60))
        return m if m >= 0 else None
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        f = float(v)
        if not math.isfinite(f) or f < 0:
            return None
        if 0 < f < 1 and not float(f).is_integer():
            return int(round(f * 1440))
        return int(round(f))
    s = texto(v).lower().replace(" ", "")
    s = re.sub(r"(min|mins|minutos)$", "", s)
    s = s.replace("h", ":").replace(";", ":").replace(",", ":")
    m = re.fullmatch(r"(\d{1,3}):(\d{1,2}):?", s)
    if m:
        mi = int(m.group(2))
        return int(m.group(1)) * 60 + mi if mi < 60 else None
    m = re.fullmatch(r"(\d{1,3})\.(\d{2})", s)
    if m and int(m.group(2)) < 60:
        return int(m.group(1)) * 60 + int(m.group(2))
    if re.fullmatch(r"\d{1,2}", s):
        return int(s)
    if re.fullmatch(r"\d{3,4}", s):
        h, mi = divmod(int(s), 100)
        return h * 60 + mi if mi < 60 else None
    return None


def diferencia_min(inicio, fin) -> int | None:
    """Minutos entre dos horas del día (tolera cruce de medianoche)."""
    if inicio is None or fin is None:
        return None
    return (int(fin) - int(inicio)) % 1440


def fmt_hora(m) -> str:
    if es_vacio(m):
        return ""
    m = int(m) % 1440
    return f"{m // 60:02d}:{m % 60:02d}"


def fmt_duracion(m) -> str:
    if es_vacio(m):
        return ""
    m = int(round(float(m)))
    signo = "-" if m < 0 else ""
    m = abs(m)
    return f"{signo}{m // 60}:{m % 60:02d}"


def a_plano_hhmm(m) -> str | None:
    """Minutos → 'HHMM' sin dos puntos (formato del archivo plano)."""
    if es_vacio(m):
        return None
    m = int(round(float(m)))
    if m < 0:
        return None
    return f"{m // 60:02d}{m % 60:02d}"


def plano_hhmm_a_min(s) -> int | None:
    return parse_duracion(s)


def parse_fecha(v) -> date | None:
    """Acepta date/datetime, 'DD/MM/YYYY', 'D-M-YY', 'YYYY-MM-DD', 'DDMMYYYY', 'YYYYMMDD', 'Oct 31, 2026',
    '31 oct 2026' y seriales de Excel. Día primero ante ambigüedad (Colombia)."""
    if es_vacio(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if hasattr(v, "to_pydatetime"):
        try:
            return v.to_pydatetime().date()
        except Exception:
            return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = int(v)
        if 20000 <= n <= 80000:
            return date(1899, 12, 30) + timedelta(days=n)
        return parse_fecha(str(n))
    s = texto(v).strip()
    s = re.sub(r"[T\s]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?.*$", "", s).strip()

    def _mk(y, mo, d):
        try:
            y = int(y)
            if y < 100:
                y += 2000
            return date(y, int(mo), int(d))
        except (ValueError, TypeError):
            return None

    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        return _mk(m.group(1), m.group(2), m.group(3))
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", s)
    if m:
        return _mk(m.group(3), m.group(2), m.group(1))
    if re.fullmatch(r"\d{8}", s):
        y4 = int(s[:4])
        if 1990 <= y4 <= 2100:
            r = _mk(s[:4], s[4:6], s[6:])
            if r:
                return r
        return _mk(s[4:], s[2:4], s[:2])
    if re.fullmatch(r"\d{6}", s):
        return _mk(s[4:], s[2:4], s[:2])
    low = sin_tildes(s).lower().replace(",", " ").replace(".", " ")
    toks = [t for t in re.split(r"[\s/-]+", low) if t and t != "de"]
    mes = next((_MESES_TXT[t[:3]] for t in toks if t[:3] in _MESES_TXT and not t.isdigit()), None)
    nums = [t for t in toks if t.isdigit()]
    if mes and len(nums) >= 2:
        anio = next((n for n in nums if len(n) == 4), nums[-1])
        dia = next((n for n in nums if n != anio and len(n) <= 2), None)
        if dia:
            return _mk(anio, mes, dia)
    return None


def parse_fecha_plano(v) -> date | None:
    """Fechas del archivo plano: 'DDMMAAAA' tiene prioridad (también 7 dígitos sin el cero inicial)."""
    if not isinstance(v, (date, datetime)) and not es_vacio(v):
        s = solo_digitos(v)
        if len(s) in (7, 8) and re.fullmatch(r"[\d\s./-]+", str(v).strip()):
            s = s.zfill(8)
            try:
                d = date(int(s[4:]), int(s[2:4]), int(s[:2]))
                if 1990 <= d.year <= 2100:
                    return d
            except ValueError:
                pass
    return parse_fecha(v)


def fecha_ddmmaaaa(d) -> str | None:
    d = parse_fecha_plano(d)
    return d.strftime("%d%m%Y") if d else None


def a_utc(fecha_local, hora_min, offset_h=OFFSET_UTC_COLOMBIA_H) -> tuple[str | None, str | None]:
    """Hora Colombia → (fecha UTC 'DDMMAAAA', hora UTC 'HHMM'). Si cruza medianoche cambia la fecha."""
    f = parse_fecha(fecha_local)
    if f is None or hora_min is None:
        return (f.strftime("%d%m%Y") if f else None), None
    dt = datetime.combine(f, time(0, 0)) + timedelta(minutes=int(hora_min)) + timedelta(hours=float(offset_h or 0))
    return dt.strftime("%d%m%Y"), dt.strftime("%H%M")


def desde_utc(fecha_utc, hora_utc, offset_h=OFFSET_UTC_COLOMBIA_H) -> tuple[date | None, int | None]:
    """Inverso de a_utc: ('01112026', '1100') → (date local, minutos locales)."""
    f = parse_fecha_plano(fecha_utc)
    h = parse_hora(hora_utc)
    if f is None:
        return None, None
    if h is None:
        return f, None
    dt = datetime.combine(f, time(0, 0)) + timedelta(minutes=h) - timedelta(hours=float(offset_h or 0))
    return dt.date(), dt.hour * 60 + dt.minute


# ============================================================
# Normalización genérica de valores por tipo de campo
# ============================================================
def normalizar_valor(tipo: str, v):
    """Normaliza un valor digitado según el tipo declarado del campo. Devuelve None si queda en blanco."""
    if es_vacio(v):
        return None
    t = (tipo or "texto").lower()
    if t == "texto":
        return texto(v)[:2000]
    if t == "mayus":
        return texto(v).upper()[:500]
    if t == "matricula":
        return norm_matricula(v) or None
    if t == "matricula_compacta":
        return matricula_compacta(v) or None
    if t == "iata":
        return norm_iata(v) or None
    if t == "fecha":
        d = parse_fecha(v)
        return d.isoformat() if d else None
    if t == "hora":
        # Convención interna: un entero ya es "minutos desde medianoche"; los textos se interpretan (05:55, 0555…)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 0 <= v < 1440:
            return int(v)
        return parse_hora(v)
    if t == "duracion":
        return parse_duracion(v)
    if t == "plano_hhmm":
        return a_plano_hhmm(parse_duracion(v))
    if t == "plano_hora":
        h = parse_hora(v)
        return f"{h // 60:02d}{h % 60:02d}" if h is not None else None
    if t == "plano_fecha":
        return fecha_ddmmaaaa(v)
    if t == "real":
        return a_float(v)
    if t == "entero":
        return a_int(v)
    if t == "producto":
        return norm_producto(v)
    if t == "digitos":
        return solo_digitos(v) or None
    if t == "mmaaaa":
        d = solo_digitos(v)
        d = d.zfill(6) if len(d) == 5 else d
        return d if len(d) == 6 and 1 <= int(d[:2]) <= 12 else None
    if t.startswith("enum:"):
        opciones = t.split(":", 1)[1].split("|")
        s = texto(v).upper()
        return s if s in opciones else None
    return texto(v)


# ============================================================
# Evaluador seguro de fórmulas (sin eval/exec)
# ============================================================
class ErrorFormula(ValueError):
    pass


def _si(cond, a, b):
    return a if cond else b


def _redondear(x, n=0):
    return round(float(x), int(n))


FUNCIONES_FORMULA = {
    "min": min, "max": max, "abs": abs, "round": _redondear, "redondear": _redondear,
    "ceil": math.ceil, "techo": math.ceil, "floor": math.floor, "piso": math.floor, "si": _si,
}
_OPS_BIN = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul}
_OPS_CMP = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt, ast.LtE: operator.le,
            ast.Gt: operator.gt, ast.GtE: operator.ge}
_MAX_NODOS = 160
_MAX_LARGO = 400


def _eval_nodo(nodo, variables):
    if isinstance(nodo, ast.Expression):
        return _eval_nodo(nodo.body, variables)
    if isinstance(nodo, ast.Constant):
        if isinstance(nodo.value, bool) or isinstance(nodo.value, (int, float)):
            return nodo.value
        raise ErrorFormula("Solo se permiten números en las fórmulas.")
    if isinstance(nodo, ast.Name):
        nombre = nodo.id
        if nombre in ("verdadero", "true", "True"):
            return True
        if nombre in ("falso", "false", "False"):
            return False
        if nombre not in variables:
            raise ErrorFormula(f"Variable no permitida: {nombre}")
        val = variables[nombre]
        if es_vacio(val):
            raise ErrorFormula(f"La variable '{nombre}' no tiene valor (dato pendiente).")
        return float(val) if not isinstance(val, bool) else val
    if isinstance(nodo, ast.BinOp):
        a = _eval_nodo(nodo.left, variables)
        b = _eval_nodo(nodo.right, variables)
        if type(nodo.op) in _OPS_BIN:
            return _OPS_BIN[type(nodo.op)](a, b)
        if isinstance(nodo.op, ast.Div):
            return safe_div(a, b, 0.0)
        if isinstance(nodo.op, ast.FloorDiv):
            return math.floor(safe_div(a, b, 0.0))
        if isinstance(nodo.op, ast.Mod):
            return 0.0 if float(b) == 0 else float(a) % float(b)
        if isinstance(nodo.op, ast.Pow):
            if abs(float(b)) > 10 or abs(float(a)) > 1e6:
                raise ErrorFormula("Potencia fuera de rango permitido.")
            return float(a) ** float(b)
        raise ErrorFormula("Operador no permitido.")
    if isinstance(nodo, ast.UnaryOp):
        v = _eval_nodo(nodo.operand, variables)
        if isinstance(nodo.op, ast.USub):
            return -v
        if isinstance(nodo.op, ast.UAdd):
            return +v
        if isinstance(nodo.op, ast.Not):
            return not v
        raise ErrorFormula("Operador unario no permitido.")
    if isinstance(nodo, ast.BoolOp):
        vals = [_eval_nodo(x, variables) for x in nodo.values]
        return all(vals) if isinstance(nodo.op, ast.And) else any(vals)
    if isinstance(nodo, ast.Compare):
        izq = _eval_nodo(nodo.left, variables)
        for op, comp in zip(nodo.ops, nodo.comparators):
            der = _eval_nodo(comp, variables)
            if type(op) not in _OPS_CMP:
                raise ErrorFormula("Comparación no permitida.")
            if not _OPS_CMP[type(op)](izq, der):
                return False
            izq = der
        return True
    if isinstance(nodo, ast.Tuple):
        raise ErrorFormula("Use punto decimal (1.5), no coma.")
    if isinstance(nodo, ast.IfExp):
        return _eval_nodo(nodo.body if _eval_nodo(nodo.test, variables) else nodo.orelse, variables)
    if isinstance(nodo, ast.Call):
        if not isinstance(nodo.func, ast.Name) or nodo.func.id not in FUNCIONES_FORMULA:
            raise ErrorFormula("Función no permitida. Use: " + ", ".join(sorted(FUNCIONES_FORMULA)))
        if nodo.keywords or len(nodo.args) > 12:
            raise ErrorFormula("Argumentos no permitidos en la función.")
        if nodo.func.id == "si":
            if len(nodo.args) != 3:
                raise ErrorFormula("si(condición, valor_si, valor_no) requiere 3 argumentos.")
            cond = _eval_nodo(nodo.args[0], variables)
            return _eval_nodo(nodo.args[1] if cond else nodo.args[2], variables)
        args = [_eval_nodo(a, variables) for a in nodo.args]
        if not args:
            raise ErrorFormula("La función requiere argumentos.")
        return FUNCIONES_FORMULA[nodo.func.id](*args)
    raise ErrorFormula(f"Elemento no permitido en la fórmula: {type(nodo).__name__}")


def _parsear_formula(expr: str):
    s = texto(expr)
    if len(s) > _MAX_LARGO:
        raise ErrorFormula(f"La fórmula supera {_MAX_LARGO} caracteres.")
    s = s.replace("^", "**").replace("×", "*").replace("÷", "/")
    if re.search(r"__|\bimport\b|\blambda\b|\[|\]|\{|\}|;|:=|\bexec\b|\beval\b", s):
        raise ErrorFormula("La fórmula contiene elementos no permitidos.")
    try:
        arbol = ast.parse(s, mode="eval")
    except SyntaxError:
        raise ErrorFormula("Sintaxis inválida. Use punto decimal (1.5) y operadores + - * / ( ).") from None
    if sum(1 for _ in ast.walk(arbol)) > _MAX_NODOS:
        raise ErrorFormula("La fórmula es demasiado compleja.")
    return arbol


def evaluar_formula(expr, variables: dict, default=None):
    """Evalúa una expresión aritmética controlada. Devuelve (valor, error). Nunca lanza excepción.
    Fórmula vacía = modo manual → (default, None)."""
    if es_vacio(expr):
        return default, None
    try:
        r = _eval_nodo(_parsear_formula(expr), variables or {})
        if isinstance(r, bool):
            r = 1.0 if r else 0.0
        r = float(r)
        if not math.isfinite(r):
            return default, "El resultado no es un número válido."
        return r, None
    except ErrorFormula as e:
        return default, str(e)
    except (TypeError, ValueError, OverflowError, RecursionError, MemoryError):
        return default, "No se pudo evaluar la fórmula con los datos disponibles."


def validar_formula(expr, permitidas) -> str | None:
    """Valida sintaxis y variables. Devuelve None si es válida o el mensaje de error."""
    if es_vacio(expr):
        return None
    try:
        arbol = _parsear_formula(expr)
    except ErrorFormula as e:
        return str(e)
    nombres = {n.id for n in ast.walk(arbol) if isinstance(n, ast.Name)}
    funciones = {n.func.id for n in ast.walk(arbol) if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)}
    desconocidas = nombres - set(permitidas) - set(FUNCIONES_FORMULA) - {"verdadero", "falso", "true", "false", "True", "False"}
    if desconocidas:
        return "Variables no permitidas: " + ", ".join(sorted(desconocidas)) + ". Disponibles: " + ", ".join(permitidas)
    _, err = evaluar_formula(expr, {v: 60.0 for v in permitidas})
    if err and "Función no permitida" in err:
        return err
    if funciones - set(FUNCIONES_FORMULA):
        return "Función no permitida."
    return err


# ============================================================
# Libro de Vuelo MO-A3-004: completar y validar
# ============================================================
TRAMO_CAMPOS = ["numero_vuelo", "origen", "destino", "galones_tanqueados", "recibo_combustible", "lbs_comb_salida",
                "pax_total", "pax_transito", "hora_salida_min", "hora_despegue_min", "hora_aterrizaje_min",
                "hora_llegada_min", "tiempo_vuelo_min", "tiempo_nocturno_min", "tiempo_instrumentos_min",
                "tiempo_total_min", "lbs_comb_llegada", "lbs_comb_usado"]


def completar_tramo(t: dict) -> set:
    """Rellena SOLO los campos vacíos que pueden deducirse de otros. Devuelve el conjunto de campos calculados."""
    calc = set()
    if es_vacio(t.get("tiempo_vuelo_min")):
        v = diferencia_min(t.get("hora_despegue_min"), t.get("hora_aterrizaje_min"))
        if v is not None:
            t["tiempo_vuelo_min"], _ = v, calc.add("tiempo_vuelo_min")
    if es_vacio(t.get("tiempo_total_min")):
        v = diferencia_min(t.get("hora_salida_min"), t.get("hora_llegada_min"))
        if v is not None:
            t["tiempo_total_min"], _ = v, calc.add("tiempo_total_min")
    if es_vacio(t.get("lbs_comb_usado")):
        a, b = a_float(t.get("lbs_comb_salida")), a_float(t.get("lbs_comb_llegada"))
        if a is not None and b is not None and a >= b:
            t["lbs_comb_usado"], _ = a - b, calc.add("lbs_comb_usado")
    return calc


def validar_tramo(t: dict, idx: int = 0) -> list[str]:
    """Coherencia interna del tramo. Devuelve advertencias (no bloquean)."""
    avisos, n = [], f"Tramo {idx + 1}"
    tv, tt = t.get("tiempo_vuelo_min"), t.get("tiempo_total_min")
    calc_v = diferencia_min(t.get("hora_despegue_min"), t.get("hora_aterrizaje_min"))
    calc_t = diferencia_min(t.get("hora_salida_min"), t.get("hora_llegada_min"))
    if not es_vacio(tv) and calc_v is not None and abs(int(tv) - calc_v) > 1:
        avisos.append(f"{n}: tiempo de vuelo {fmt_duracion(tv)} ≠ aterrizaje − despegue ({fmt_duracion(calc_v)}).")
    if not es_vacio(tt) and calc_t is not None and abs(int(tt) - calc_t) > 1:
        avisos.append(f"{n}: tiempo total {fmt_duracion(tt)} ≠ llegada − salida ({fmt_duracion(calc_t)}).")
    if not es_vacio(tv) and not es_vacio(tt) and int(tv) > int(tt):
        avisos.append(f"{n}: el tiempo de vuelo supera el tiempo total (bloque).")
    a, b, u = a_float(t.get("lbs_comb_salida")), a_float(t.get("lbs_comb_llegada")), a_float(t.get("lbs_comb_usado"))
    if None not in (a, b, u) and abs((a - b) - u) > 1:
        avisos.append(f"{n}: libras usadas {u:,.0f} ≠ salida − llegada ({a - b:,.0f}).")
    pt, ptr = a_int(t.get("pax_total")), a_int(t.get("pax_transito"))
    if pt is not None and ptr is not None and ptr > pt:
        avisos.append(f"{n}: pasajeros en tránsito mayores que el total.")
    if es_vacio(t.get("origen")) or es_vacio(t.get("destino")):
        avisos.append(f"{n}: falta origen o destino.")
    return avisos


def validar_totales_libro(tramos: list[dict], totales: dict) -> list[dict]:
    """Compara la suma vertical de los tramos contra la fila TOTALES del libro."""
    conceptos = [("Tiempo de vuelo", "tiempo_vuelo_min", True), ("Tiempo nocturno", "tiempo_nocturno_min", True),
                 ("Tiempo instrumentos", "tiempo_instrumentos_min", True), ("Tiempo total (bloque)", "tiempo_total_min", True),
                 ("Libras de combustible usadas", "lbs_comb_usado", False)]
    out = []
    for etiqueta, campo, es_tiempo in conceptos:
        vals = [a_float(t.get(campo)) for t in tramos]
        vals = [v for v in vals if v is not None]
        suma = sum(vals) if vals else None
        tot = a_float(totales.get(campo)) if totales else None
        if suma is None and tot is None:
            continue
        if tot is None:
            estado = "COMPLETAR"
        elif suma is None:
            estado = "SIN DETALLE"
        else:
            estado = "OK" if abs(suma - tot) <= (1 if es_tiempo else 1.0) else "DESCUADRE"
        fmt = fmt_duracion if es_tiempo else (lambda x: f"{x:,.0f}")
        out.append({"CONCEPTO": etiqueta, "CAMPO": campo, "SUMA TRAMOS": fmt(suma) if suma is not None else "",
                    "FILA TOTALES": fmt(tot) if tot is not None else "", "ESTADO": estado,
                    "_suma": suma, "_total": tot})
    return out


# ============================================================
# TIEMPO PILOTOS y liquidación de horas y pesos
# ============================================================
def variables_tramo(t: dict, ajuste_min=0) -> dict:
    return {"tiempo_vuelo": a_float(t.get("tiempo_vuelo_min")), "tiempo_total": a_float(t.get("tiempo_total_min")),
            "tiempo_nocturno": a_float(t.get("tiempo_nocturno_min"), 0.0),
            "tiempo_instrumentos": a_float(t.get("tiempo_instrumentos_min"), 0.0),
            "ajuste_min": a_float(ajuste_min, 0.0), "pax_total": a_float(t.get("pax_total"), 0.0)}


def tiempo_pilotos_efectivo(t: dict, cfg: dict) -> tuple[int | None, str]:
    """
    Valor de TIEMPO PILOTOS para un tramo y su origen:
      MANUAL (digitado por Contabilidad), FORMULA, RANGO o PENDIENTE (en blanco).
    cfg = {modo: MANUAL|FORMULA|RANGOS, formula, rangos:[{desde,hasta,resultado}], base, ajuste_min}
    """
    manual = a_float(t.get("tiempo_pilotos_min"))
    if manual is not None:
        return max(0, int(round(manual))), "MANUAL"
    cfg = cfg or {}
    modo = texto(cfg.get("modo"), "MANUAL").upper()
    variables = variables_tramo(t, cfg.get("ajuste_min"))
    if modo == "FORMULA":
        val, err = evaluar_formula(cfg.get("formula"), variables)
        if err or val is None:
            return None, "PENDIENTE"
        return max(0, int(round(val))), "FORMULA"
    if modo == "RANGOS":
        base = variables.get(texto(cfg.get("base"), "tiempo_vuelo"))
        if base is None:
            return None, "PENDIENTE"
        for r in cfg.get("rangos") or []:
            d, h = a_float(r.get("desde")), a_float(r.get("hasta"))
            if (d is None or base >= d) and (h is None or base <= h):
                val, err = evaluar_formula(texto(r.get("resultado")), variables)
                if err or val is None:
                    return None, "PENDIENTE"
                return max(0, int(round(val))), "RANGO"
        return None, "PENDIENTE"
    return None, "PENDIENTE"


def _clave_persona(cedula, nombre) -> str:
    ced = solo_digitos(cedula)
    return ced if len(ced) >= 5 else ("NOMBRE:" + nombre_clave(nombre) if texto(nombre) else "")


def liquidar_tripulacion(asignaciones: list[dict], params: dict, ajustes: dict | None = None,
                         categorias: dict | None = None) -> list[dict]:
    """
    asignaciones: [{cedula, nombre, rol_vuelo: 'PTL'|'PCA', minutos: int|None}]
    params: {min_smlv, tarifa_ptl, tarifa_pca, formula_total}
    ajustes: {clave_persona: {campo: valor}}  (overrides manuales auditados)
    categorias: {cedula: 'PTL'|'PCA'} desde usuarios
    """
    ajustes, categorias = ajustes or {}, categorias or {}
    grupos: dict[str, dict] = {}
    for a in asignaciones:
        clave = _clave_persona(a.get("cedula"), a.get("nombre"))
        if not clave:
            continue
        g = grupos.setdefault(clave, {"cedula": solo_digitos(a.get("cedula")) or "", "nombre": texto(a.get("nombre")),
                                      "min_piloto": 0, "min_copiloto": 0, "tramos": 0, "pendientes": 0,
                                      "n_ptl": 0, "n_pca": 0})
        if not g["nombre"] and texto(a.get("nombre")):
            g["nombre"] = texto(a.get("nombre"))
        g["tramos"] += 1
        g["n_pca" if texto(a.get("rol_vuelo")).upper() == "PCA" else "n_ptl"] += 1
        m = a_float(a.get("minutos"))
        if m is None:
            g["pendientes"] += 1
            continue
        if texto(a.get("rol_vuelo")).upper() == "PCA":
            g["min_copiloto"] += int(round(m))
        else:
            g["min_piloto"] += int(round(m))
    formula = params.get("formula_total")
    formula = FORMULA_LIQUIDACION_DEFECTO if formula is None else formula
    filas = []
    for clave, g in grupos.items():
        aj = {k: v for k, v in (ajustes.get(clave) or {}).items() if not es_vacio(v)}
        mp = a_int(aj.get("min_piloto"), g["min_piloto"])
        mc = a_int(aj.get("min_copiloto"), g["min_copiloto"])
        volados = mp + mc
        cat = texto(aj.get("categoria")).upper() or texto(categorias.get(g["cedula"])).upper()
        if cat not in ("PTL", "PCA"):
            cat = "PTL" if (mp > mc or (mp == mc and g["n_ptl"] >= g["n_pca"])) else "PCA"
        smlv = a_float(aj.get("min_smlv"), a_float(params.get("min_smlv")))
        tarifa = a_float(params.get("tarifa_ptl")) if cat == "PTL" else a_float(params.get("tarifa_pca"))
        valor_hora = a_float(aj.get("valor_hora"), tarifa)
        valor_min = round(safe_div(valor_hora, 60.0), 4) if valor_hora is not None else None
        exced = max(0, int(round(volados - smlv))) if smlv is not None else None
        horas, mins = divmod(exced, 60) if exced is not None else (None, None)
        total = a_float(aj.get("total_pagar"))
        if total is None:
            variables = {"min_piloto": mp, "min_copiloto": mc, "min_volados": volados, "min_smlv": smlv,
                         "min_excedentes": exced, "valor_hora": valor_hora, "valor_minuto": valor_min,
                         "tarifa_ptl": a_float(params.get("tarifa_ptl")), "tarifa_pca": a_float(params.get("tarifa_pca")),
                         "horas_excedentes": safe_div(exced, 60.0) if exced is not None else None}
            total, _ = evaluar_formula(formula, variables)
            total = round(total) if total is not None else None
        filas.append({"CLAVE": clave, "CEDULA": g["cedula"], "NOMBRE": g["nombre"], "CATEGORIA": cat,
                      "MIN PILOTO": mp, "MIN COPILOTO": mc, "MIN VOLADOS": volados, "MIN SMLV": smlv,
                      "TOTAL MIN EXCEDENTES": exced, "VALOR HORA": valor_hora, "VALOR MINUTO": valor_min,
                      "HORAS": horas, "MIN": mins, "TOTAL A PAGAR (COP)": total, "TRAMOS": g["tramos"],
                      "TRAMOS SIN TIEMPO PILOTOS": g["pendientes"], "AJUSTES MANUALES": ", ".join(sorted(aj)) or ""})
    filas.sort(key=lambda r: (-(r["MIN VOLADOS"] or 0), r["NOMBRE"]))
    return filas


# ============================================================
# Archivo plano Aerocivil MO-A3-013
# ============================================================
def galones_charter(basico_lbs, reserva_lbs, factor_lbs_gal, decimales=0) -> tuple[float | None, str]:
    """Galones = (Combustible básico lbs + Reserva fija lbs) / Factor lbs/gal. None si falta un parámetro."""
    b, r, f = a_float(basico_lbs), a_float(reserva_lbs), a_float(factor_lbs_gal)
    if b is None:
        return None, "Falta el combustible básico (lbs) del manifiesto de peso y balance."
    if r is None:
        return None, "Reserva fija (lbs) sin parametrizar para el modelo."
    if f is None or f <= 0:
        return None, "Factor de conversión lbs/gal sin parametrizar para el modelo."
    return round(safe_div(b + r, f), int(decimales or 0)), "Fórmula chárter (básico + reserva) / factor"


def pax_a_bordo(pax_total, tripadis=None, bebes=None, transito=None) -> tuple[int | None, str]:
    """PASAJEROS A BORDO = Pax total − Tripadis − Bebés (Nibra) − Pasajeros en tránsito (nunca negativo)."""
    t = a_int(pax_total)
    if t is None:
        return None, "Sin pasajeros totales."
    v = t - (a_int(tripadis, 0) or 0) - (a_int(bebes, 0) or 0) - (a_int(transito, 0) or 0)
    if v < 0:
        return 0, "Los descuentos superan el total de pasajeros: revisar."
    return v, ""


def construir_fila_plano(tramo: dict, modelo: dict | None, ruta: dict | None, ficha: dict | None,
                         ruta_charter: dict | None, params: dict) -> tuple[dict, list[str], list[str]]:
    """
    Construye una fila MO-A3-013 (27 columnas) a partir del tramo del Libro de Vuelo, su ficha (cierre/P&B),
    el modelo de aeronave y el catálogo de rutas (hoja BD). Devuelve (fila, pendientes, notas).
    Lo que no se pueda deducir queda en None (celda en blanco editable).
    """
    modelo, ruta, ficha, params = modelo or {}, ruta or {}, ficha or {}, params or {}
    notas = []
    matricula = norm_matricula(tramo.get("matricula"))
    fecha = parse_fecha(tramo.get("fecha"))
    hora = tramo.get("hora_salida_min")
    if es_vacio(hora):
        hora = tramo.get("hora_despegue_min")
    hora = None if es_vacio(hora) else int(hora)
    if hora is None and not es_vacio(ficha.get("hora_min")):
        hora = int(ficha.get("hora_min"))
    fecha_utc, hora_utc = a_utc(fecha, hora, a_float(params.get("offset_utc_horas"), OFFSET_UTC_COLOMBIA_H))

    nro_plano, es_charter = normalizar_nro_vuelo(tramo.get("numero_vuelo"), matricula)
    if nro_plano is None and texto(ficha.get("numero_vuelo")):
        nro_plano, es_charter = normalizar_nro_vuelo(ficha.get("numero_vuelo"), matricula)

    origen = norm_iata(tramo.get("origen")) or norm_iata(ruta.get("origen")) or norm_iata(ficha.get("origen")) or None
    destino = norm_iata(tramo.get("destino")) or norm_iata(ruta.get("destino")) or norm_iata(ficha.get("destino")) or None

    if es_charter:
        tipo_vuelo = "C"
    else:
        tipo_vuelo = texto(ruta.get("tipo_vuelo")).upper() or None

    # Combustible consumido
    combustible = None
    if not es_charter and ruta and not es_vacio(ruta.get("combustible_estandar")):
        combustible = a_float(ruta.get("combustible_estandar"))
        notas.append(f"Combustible estándar de la ruta {nro_plano} (catálogo BD).")
    elif es_charter or not ruta:
        val, motivo = galones_charter(ficha.get("combustible_basico_lbs"), modelo.get("reserva_combustible_lbs"),
                                      modelo.get("factor_conversion_lbs_gal"), a_int(params.get("combustible_decimales"), 0))
        if val is not None:
            combustible = val
            notas.append(motivo)
        elif ruta_charter and not es_vacio(ruta_charter.get("combustible_galones")):
            combustible = a_float(ruta_charter.get("combustible_galones"))
            notas.append("Combustible de ruta chárter frecuente guardada.")
        else:
            notas.append("Combustible: " + motivo)

    # Pasajeros (Libro) depurados con la ficha (tripadis / bebés / tránsito)
    pax_total = tramo.get("pax_total")
    if es_vacio(pax_total):
        pax_total = ficha.get("pax_total")
    transito = tramo.get("pax_transito")
    if es_vacio(transito):
        transito = ficha.get("pax_transito")
    transito = a_int(transito, 0) or 0
    pax, alerta = pax_a_bordo(pax_total, ficha.get("tripadis"), ficha.get("bebes"), transito)
    if alerta:
        notas.append(alerta)
    if not ficha:
        notas.append("Sin ficha de vuelo: no se descontaron tripadis ni bebés.")

    fila = {
        "tip": texto(params.get("plano_tip"), "TE") or None,
        "sigla": texto(params.get("plano_sigla"), "3CA") or None,
        "fecha_mmaaaa": fecha.strftime("%m%Y") if fecha else None,
        "fecha_utc": fecha_utc, "hora_utc": hora_utc,
        "origen": origen, "destino": destino,
        "matricula": matricula_compacta(matricula) or None,
        "tipo_equipo": texto(modelo.get("codigo")) or None,
        "nro_vuelo": nro_plano,
        "sillas_ofrecidas": a_int(modelo.get("sillas_ofrecidas")),
        "carga_ofrecida": a_float(modelo.get("carga_ofrecida_kg")),
        "pax_bordo": pax,
        "carga_bordo": a_float(ficha.get("carga_a_bordo_kg")),
        "correo_bordo": a_float(params.get("correo_default"), 0.0),
        "pax_transito": transito,
        "carga_transito": a_float(ficha.get("carga_transito_kg"), a_float(params.get("carga_transito_default"))),
        "horas_bloque": a_plano_hhmm(tramo.get("tiempo_total_min")),
        "tipo_vuelo": tipo_vuelo,
        "tipo_trafico": texto(modelo.get("tipo_trafico_default")) or None,
        "uso_aeronave": texto(modelo.get("uso_aeronave_default")) or None,
        "tipo_combustible": texto(modelo.get("tipo_combustible_default")) or None,
        "metodo_medicion": texto(modelo.get("metodo_medicion_default")) or None,
        "combustible_consumido": combustible,
        "horas_vuelo": a_plano_hhmm(tramo.get("tiempo_vuelo_min")),
        "carga_real": None,
        # metadatos (no salen en el plano)
        "fecha_local": fecha.isoformat() if fecha else None,
        "hora_local_min": hora,
        "es_charter": 1 if es_charter else 0,
    }
    pendientes = pendientes_fila(fila)
    if not modelo:
        notas.append("Matrícula sin modelo configurado: complete equipo, sillas y carga ofrecida.")
    return fila, pendientes, notas


def pendientes_fila(fila: dict) -> list[str]:
    return [c for c in PLANO_CLAVES if c != "nro" and es_vacio(fila.get(c))]


def clave_vuelo(fecha, matricula, hora_min, origen, destino) -> str:
    """Llave compuesta Fecha + Matrícula + Hora + Origen + Destino."""
    f = parse_fecha(fecha)
    return "|".join([f.isoformat() if f else "", matricula_compacta(matricula), fmt_hora(hora_min),
                     norm_iata(origen), norm_iata(destino)])


def emparejar_fichas(tramos: list[dict], fichas: list[dict], tolerancia_min=120) -> dict:
    """Empareja cada tramo con su ficha por fecha + matrícula + origen/destino + hora más cercana
    (o por número de vuelo). Cada ficha se usa una sola vez. Devuelve {tramo_id: ficha_id}."""
    candidatos = []
    for t in tramos:
        ft, mt = parse_fecha(t.get("fecha")), matricula_compacta(t.get("matricula"))
        ht = t.get("hora_salida_min") if not es_vacio(t.get("hora_salida_min")) else t.get("hora_despegue_min")
        nt, _ = normalizar_nro_vuelo(t.get("numero_vuelo"), t.get("matricula"))
        for f in fichas:
            if parse_fecha(f.get("fecha")) != ft or matricula_compacta(f.get("matricula")) != mt:
                continue
            score = 0.0
            mismo_od = norm_iata(f.get("origen")) == norm_iata(t.get("origen")) and norm_iata(f.get("destino")) == norm_iata(t.get("destino"))
            if mismo_od and norm_iata(t.get("origen")):
                score += 100
            nf, _ = normalizar_nro_vuelo(f.get("numero_vuelo"), f.get("matricula"))
            if nf and nt and nf == nt:
                score += 40
            hf = f.get("hora_min")
            if not es_vacio(hf) and not es_vacio(ht):
                dif = min(diferencia_min(ht, hf), diferencia_min(hf, ht))
                if dif > tolerancia_min and score < 100:
                    continue
                score += max(0, 60 - dif / 2)
            if score >= 40:
                candidatos.append((score, t.get("id"), f.get("id")))
    candidatos.sort(key=lambda x: -x[0])
    usados_t, usados_f, out = set(), set(), {}
    for score, tid, fid in candidatos:
        if tid in usados_t or fid in usados_f:
            continue
        out[tid] = fid
        usados_t.add(tid)
        usados_f.add(fid)
    return out


def ordenar_plano(filas: list[dict]) -> list[dict]:
    """Orden del despacho: fecha → matrícula → hora; asigna NRO consecutivo 1..n."""
    def _k(f):
        d = parse_fecha(f.get("fecha_local")) or desde_utc(f.get("fecha_utc"), f.get("hora_utc"))[0] or date.min
        h = f.get("hora_local_min")
        if es_vacio(h):
            h = parse_hora(f.get("hora_utc")) or 0
        return (d, texto(f.get("matricula")), int(h))
    out = sorted(filas, key=_k)
    for i, f in enumerate(out, 1):
        f["nro"] = i
    return out


# ============================================================
# Conciliación de combustible (tiquete ↔ libro de vuelo)
# ============================================================
def _ult_digitos(s, n=5) -> str:
    d = solo_digitos(s)
    return d[-n:] if len(d) >= 4 else ""


def conciliar_combustible(tiquetes: list[dict], tramos: list[dict], tol_pct=2.0, tol_abs=2.0):
    """
    Cruza cada tiquete contra los tramos con tanqueo del Libro de Vuelo (misma matrícula y fecha).
    Devuelve (resultado_por_tiquete, tramos_con_tanqueo_sin_tiquete).
    """
    con_tanqueo = [t for t in tramos if a_float(t.get("galones_tanqueados"), 0) > 0 or texto(t.get("recibo_combustible"))]
    usados, res = set(), []
    for q in tiquetes:
        fq, mq = parse_fecha(q.get("fecha_consumo")), matricula_compacta(q.get("matricula"))
        gq = a_float(q.get("galones"), 0.0)
        recibos_q = {_ult_digitos(q.get("recibo")), _ult_digitos(q.get("numero_boleto"))} - {""}
        mejor, mejor_s = None, 0.0
        for t in con_tanqueo:
            if t.get("id") in usados or matricula_compacta(t.get("matricula")) != mq:
                continue
            ft = parse_fecha(t.get("fecha"))
            if ft is None or fq is None or abs((ft - fq).days) > 1:
                continue
            s = 10.0 if ft == fq else 2.0
            rt = _ult_digitos(t.get("recibo_combustible"))
            if rt and rt in recibos_q:
                s += 100
            gt = a_float(t.get("galones_tanqueados"))
            if gt:
                s += max(0.0, 30 - abs(gt - gq))
            if norm_iata(q.get("estacion")) and norm_iata(q.get("estacion")) == norm_iata(t.get("origen")):
                s += 15
            if s > mejor_s:
                mejor, mejor_s = t, s
        fila = {"tiquete_id": q.get("id"), "tramo_id": None, "libro_id": None, "numero_libro": None,
                "gal_libro": None, "dif_gal": None, "estado": "SIN LIBRO DE VUELO"}
        if mejor is not None and mejor_s >= 20:
            usados.add(mejor.get("id"))
            gt = a_float(mejor.get("galones_tanqueados"))
            dif = round(gq - gt, 1) if gt is not None else None
            tol = max(a_float(tol_abs, 2.0), abs(gq) * a_float(tol_pct, 2.0) / 100.0)
            estado = "CONCILIADO" if dif is not None and abs(dif) <= tol else ("SIN GALONES EN LIBRO" if dif is None else "DIFERENCIA GALONES")
            fila.update({"tramo_id": mejor.get("id"), "libro_id": mejor.get("libro_id"), "numero_libro": mejor.get("numero_libro"),
                         "gal_libro": gt, "dif_gal": dif, "estado": estado})
        res.append(fila)
    sin_tiquete = [t for t in con_tanqueo if t.get("id") not in usados]
    return res, sin_tiquete

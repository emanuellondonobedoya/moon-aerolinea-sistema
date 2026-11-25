# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Exportaciones Excel (exports.py)

  * Libros con formato corporativo (encabezados azul marino, filtros, paneles fijos, anchos, formatos COP/galones).
  * Archivo plano MO-A3-013: Control interno (27 columnas A..AA, bloque T..AA en rojo) y archivo limpio Aerocivil
    (19 columnas A..S), con códigos como texto (DDMMAAAA, HHMM) para no perder ceros a la izquierda.
  * Consolidados diario / mensual / anual con portada de trazabilidad (hash SHA-256 de la instantánea).
  * Liquidación de pilotos y reportes de auditoría.
Protección contra inyección de fórmulas: todo texto que empiece por = + - @ se escribe como texto literal.
"""
from __future__ import annotations

import io
import math
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import rules_engine as R

NAVY, BLUE, GOLD, RED, GRIS = "0B1D33", "356B9A", "C9A96E", "B83A32", "F4F7FB"
_BORDE = Border(bottom=Side(style="thin", color="DDE5EE"))
_FMT_COP = '"$" #,##0'
_FMT_GAL = '#,##0.0'
_FMT_ENT = '#,##0'


def _valor_seguro(v):
    if v is None:
        return None
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):
        try:
            v = v.item()
        except (ValueError, AttributeError):
            pass
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (list, dict, tuple, set)):
        return str(v)
    return v


def _escribir_celda(ws, fila, col, v, texto=False):
    v = _valor_seguro(v)
    celda = ws.cell(row=fila, column=col)
    if isinstance(v, str):
        celda.value = v
        celda.data_type = "s"          # jamás se interpreta como fórmula (=, +, -, @)
        if texto:
            celda.number_format = "@"
    else:
        celda.value = v
    return celda


def _formato_columna(nombre: str) -> str | None:
    n = str(nombre).upper()
    if any(k in n for k in ("VALOR", "COSTO", "COP", "PAGAR", "PRECIO", "INVERSI")):
        return _FMT_COP
    if "GAL" in n or "LBS" in n or "CARGA" in n or "KG" in n:
        return _FMT_GAL
    if n.startswith("MIN") or n in ("HORAS", "MIN", "TRAMOS", "PAX", "PASAJEROS A BORDO"):
        return _FMT_ENT
    return None


def escribir_hoja(ws, df: pd.DataFrame, titulo: str | None = None, subtitulo: str | None = None,
                  colores_encabezado: dict | None = None, columnas_texto: set | None = None, fila_inicio: int = 1):
    """Escribe un DataFrame con formato corporativo. colores_encabezado: {nombre_columna: hex}."""
    columnas_texto = columnas_texto or set()
    colores_encabezado = colores_encabezado or {}
    f = fila_inicio
    if titulo:
        ws.cell(row=f, column=1, value=titulo).font = Font(bold=True, size=14, color=NAVY)
        f += 1
    if subtitulo:
        c = ws.cell(row=f, column=1)
        c.value, c.data_type = subtitulo, "s"
        c.font = Font(italic=True, size=10, color="5F7087")
        f += 1
    if titulo or subtitulo:
        f += 1
    cab = f
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(row=cab, column=j)
        c.value, c.data_type = str(col), "s"
        c.fill = PatternFill("solid", fgColor=colores_encabezado.get(col, NAVY))
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[cab].height = 32
    formatos = {j: _formato_columna(col) for j, col in enumerate(df.columns, 1)}
    for i, fila in enumerate(df.itertuples(index=False, name=None), cab + 1):
        for j, v in enumerate(fila, 1):
            col = df.columns[j - 1]
            celda = _escribir_celda(ws, i, j, v, texto=col in columnas_texto)
            if formatos[j] and isinstance(celda.value, (int, float)) and col not in columnas_texto:
                celda.number_format = formatos[j]
            celda.border = _BORDE
    ws.freeze_panes = ws.cell(row=cab + 1, column=1)
    if len(df.columns):
        ws.auto_filter.ref = f"A{cab}:{get_column_letter(len(df.columns))}{max(cab, cab + len(df))}"
    for j, col in enumerate(df.columns, 1):
        muestras = [len(str(col))] + [len(str(v)) for v in df.iloc[:300, j - 1].tolist() if v is not None]
        ws.column_dimensions[get_column_letter(j)].width = min(max(max(muestras) + 2, 9), 48)
    return cab


def libro(hojas: dict, portada: dict | None = None) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    if portada:
        _portada(wb, portada)
    for nombre, df in hojas.items():
        ws = wb.create_sheet(_nombre_hoja(nombre))
        if df is None or (isinstance(df, pd.DataFrame) and df.empty and len(df.columns) == 0):
            df = pd.DataFrame({"INFORMACIÓN": ["Sin registros en el periodo."]})
        escribir_hoja(ws, df)
    if not wb.sheetnames:
        wb.create_sheet("Vacío")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _nombre_hoja(nombre: str) -> str:
    limpio = "".join(ch for ch in str(nombre) if ch not in '[]:*?/\\')
    return limpio[:31] or "Hoja"


def _portada(wb, datos: dict):
    ws = wb.create_sheet("Portada", 0)
    ws.sheet_view.showGridLines = False
    logo = Path(__file__).resolve().parent / "assets" / "logo_moon.png"
    fila = 1
    if logo.is_file():
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(str(logo))
            img.width, img.height = 260, int(260 * img.height / img.width)
            ws.add_image(img, "B2")
            fila = 9
        except Exception:
            fila = 2
    ws.cell(row=fila, column=2, value="Moon Aerolínea · El cielo nos une").font = Font(bold=True, size=16, color=NAVY)
    fila += 2
    for k, v in datos.items():
        a = ws.cell(row=fila, column=2, value=str(k))
        a.font = Font(bold=True, color="FFFFFF")
        a.fill = PatternFill("solid", fgColor=NAVY)
        a.alignment = Alignment(vertical="center")
        _escribir_celda(ws, fila, 3, v if not isinstance(v, (date, datetime)) else str(v))
        ws.cell(row=fila, column=3).fill = PatternFill("solid", fgColor=GRIS)
        fila += 1
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 90


# ============================================================
# Archivo plano MO-A3-013
# ============================================================
_TEXTO_PLANO = {R.PLANO_HEADERS[c] for c in R.PLANO_TEXTO}


def _df_plano_export(df: pd.DataFrame, claves: list[str]) -> pd.DataFrame:
    out = pd.DataFrame()
    for c in claves:
        h = R.PLANO_HEADERS[c]
        serie = df[h] if h in df.columns else pd.Series([None] * len(df))
        if c == "nro_vuelo":
            serie = serie.map(lambda v: int(v) if R.solo_digitos(v) and str(R.solo_digitos(v)) == str(v).strip() else v)
        elif c in ("sillas_ofrecidas", "pax_bordo", "pax_transito", "nro"):
            serie = serie.map(lambda v: R.a_int(v))
        elif c in R.PLANO_NUMERICAS:
            serie = serie.map(lambda v: (lambda f: (int(f) if f is not None and float(f).is_integer() else f))(R.a_float(v)))
        out[h] = serie.values
    return out


def excel_plano_control(df: pd.DataFrame, periodo_txt: str, sigla="3CA") -> bytes:
    """27 columnas A..AA. df con encabezados oficiales (sin ID)."""
    datos = _df_plano_export(df, R.PLANO_CLAVES)
    wb = Workbook()
    ws = wb.active
    ws.title = "MES"
    ws.merge_cells("A1:S2")
    t = ws["A1"]
    t.value = "ALIMENTADOR RESUMEN DIARIO DE CUMPLIMIENTO SISTEMA PLANO"
    t.font = Font(bold=True, size=14, color=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    for celda, txt in (("T1", "MO-A3-013"), ("T2", "Control interno Moon Aerolínea"), ("T3", f"Periodo: {periodo_txt}")):
        ws[celda].value = txt
        ws[celda].font = Font(bold=True, color=NAVY)
    ws["A3"].value = f"Empresa: {sigla}"
    ws["A3"].font = Font(bold=True, color="5F7087")
    colores = {R.PLANO_HEADERS[c]: (NAVY if c in R.PLANO_OFICIALES else RED) for c in R.PLANO_CLAVES}
    escribir_hoja(ws, datos, colores_encabezado=colores, columnas_texto=_TEXTO_PLANO, fila_inicio=5)
    # Letras de columna como referencia
    for j, (letra, _, _) in enumerate(R.PLANO_COLUMNAS, 1):
        c = ws.cell(row=4, column=j, value=letra)
        c.font = Font(bold=True, color=RED if j > 19 else BLUE, size=9)
        c.alignment = Alignment(horizontal="center")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_plano_aerocivil(df: pd.DataFrame) -> bytes:
    """Archivo limpio: SOLO columnas oficiales A..S, encabezado en la fila 1, códigos como texto."""
    datos = _df_plano_export(df, R.PLANO_OFICIALES)
    wb = Workbook()
    ws = wb.active
    ws.title = "PLANO"
    for j, col in enumerate(datos.columns, 1):
        c = ws.cell(row=1, column=j)
        c.value, c.data_type = col, "s"
        c.font = Font(bold=True)
    for i, fila in enumerate(datos.itertuples(index=False, name=None), 2):
        for j, v in enumerate(fila, 1):
            _escribir_celda(ws, i, j, v, texto=datos.columns[j - 1] in _TEXTO_PLANO)
    for j, col in enumerate(datos.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = max(10, min(len(col) + 2, 26))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# Consolidados, liquidación y auditoría
# ============================================================
def nombre_consolidado(nivel: str, area: str, periodo: str) -> str:
    p = str(periodo)
    if nivel == "DIARIO":
        return f"Moon_Consolidado_Diario_{area}_{p.replace('-', '')}.xlsx"
    if nivel == "MENSUAL":
        return f"Moon_Historico_Mensual_{area}_{p.replace('-', '_')}.xlsx"
    return f"Moon_Historico_Anual_{area}_{p}.xlsx"


def excel_consolidado(nivel: str, area: str, periodo: str, hojas: dict, meta: dict | None, usuario: dict) -> bytes:
    meta = meta or {}
    portada = {"Documento": f"Consolidado {nivel.title()} · {area}", "Periodo": periodo,
               "Estado": meta.get("estado") or "ABIERTO (vista en vivo, no congelada)",
               "Registros": sum(len(d) for d in hojas.values()),
               "Cerrado por": f"{meta.get('cerrado_por_nombre') or ''} {meta.get('cerrado_por_cedula') or ''}".strip() or "—",
               "Fecha de cierre": meta.get("fecha_cierre") or meta.get("fecha_actualizacion") or "—",
               "Huella SHA-256 instantánea": meta.get("sha256") or "—",
               "Integridad verificada": ("Sí" if meta.get("integro") else "NO COINCIDE") if "integro" in meta else "—",
               "Generado por": f"{usuario.get('nombre')} · C.C. {usuario.get('cedula')}",
               "Generado el": datetime.now().strftime("%Y-%m-%d %H:%M")}
    if "Plano MO-A3-013" in hojas:
        hojas = dict(hojas)
        hojas["Plano MO-A3-013"] = _df_plano_export(hojas["Plano MO-A3-013"], R.PLANO_CLAVES)
    return libro(hojas, portada)


def excel_liquidacion(liq: pd.DataFrame, detalle: pd.DataFrame, combustible: pd.DataFrame, auditoria: pd.DataFrame,
                      periodo: str, usuario: dict) -> bytes:
    portada = {"Documento": "Liquidación de horas y honorarios de tripulación", "Periodo": periodo,
               "Generado por": f"{usuario.get('nombre')} · C.C. {usuario.get('cedula')}",
               "Generado el": datetime.now().strftime("%Y-%m-%d %H:%M"),
               "Nota": "Celdas en blanco = dato manual / fórmula pendiente. Ajustes manuales trazados en 'Auditoría Cambios'."}
    det = detalle.drop(columns=[c for c in detalle.columns if str(c).startswith("_")], errors="ignore")
    return libro({"Liquidación Pilotos": liq.drop(columns=["CLAVE"], errors="ignore"), "Detalle Vuelos": det,
                  "Combustible": combustible, "Auditoría Cambios": auditoria}, portada)


def excel_auditoria(ambito: str, desde, hasta, hallazgos: pd.DataFrame, historial: pd.DataFrame, resumen: str,
                    usuario: dict) -> bytes:
    portada = {"Documento": f"Informe de auditoría · {ambito}", "Periodo": f"{desde} a {hasta}",
               "Hallazgos": len(hallazgos), "Ejecutada por": f"{usuario.get('nombre')} · C.C. {usuario.get('cedula')}",
               "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M")}
    lineas = [l for l in (resumen or "").splitlines() if l.strip()] or ["Sin resumen."]
    return libro({"Resumen IA": pd.DataFrame({"RESUMEN Y RECOMENDACIONES": lineas}), "Hallazgos": hallazgos,
                  "Historial Responsables": historial}, portada)

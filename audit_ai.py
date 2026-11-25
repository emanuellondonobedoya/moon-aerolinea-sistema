# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Auditoría con IA y desempeño (audit_ai.py)

  * Auditoría segmentada por área:
      - CONTABILIDAD: tiquetes vs. libros, diferencias de galones, recibos duplicados, valores sin valorizar,
        TIEMPO PILOTOS pendiente/atípico, descuadres de totales y documentos faltantes de pilotos.
      - OPERACIONES: celdas pendientes del plano, pasajeros libro vs. ficha, sobreocupación, cargas, bloque vs. vuelo,
        saltos de consecutivo de libros, desequilibrio de combustible entre tramos y fichas/libros faltantes.
  * Auditoría GENERAL (revisoría fiscal): todo lo anterior + integridad (hash por registro, SHA-256 de evidencias,
    cadena de audit_logs) + re-lectura independiente de imágenes con Claude Vision y trazabilidad de responsables.
  * KPIs de desempeño de empleados (exclusivo Gerencia): productividad, oportunidad y calidad.
Si la IA no está disponible las reglas deterministas igual se ejecutan y se genera un resumen automático.
"""
from __future__ import annotations

import json
import logging

import pandas as pd

import database as db
import rules_engine as R
import vision_ocr as V

log = logging.getLogger("moon.audit")
TABLAS_NEGOCIO = ["combustible_tiquetes", "libros_vuelo", "tramos_vuelo", "fichas_vuelo", "operaciones_plano", "mantenimiento_oma"]
# Hallazgos que NO se cuentan como error de la persona en la evaluación de calidad (pendientes de flujo o
# intervenciones externas a la aplicación).
CATEGORIAS_SIN_CULPA = {"TIEMPO PILOTOS pendiente", "Tiquete sin valorizar", "Celdas pendientes en archivo plano",
                        "Registro alterado fuera del sistema", "Evidencia faltante", "Evidencia alterada",
                        "Bitácora de auditoría alterada"}
_SEV_ORDEN = {"Alta": 0, "Media": 1, "Baja": 2}


def _h(sev, area, cat, tabla, rid, fecha, mat, detalle, accion) -> dict:
    return {"SEVERIDAD": sev, "AREA": area, "CATEGORIA": cat, "TABLA": tabla, "REGISTRO_ID": None if rid is None else str(rid),
            "FECHA": R.normalizar_valor("fecha", fecha), "MATRICULA": R.norm_matricula(mat) or None, "DETALLE": detalle,
            "ACCION_SUGERIDA": accion}


def _vacio() -> pd.DataFrame:
    return pd.DataFrame(columns=db.COLS_HALLAZGO)


# ============================================================
# Reglas CONTABILIDAD
# ============================================================
def auditar_contabilidad(desde, hasta) -> list[dict]:
    p = db.params_dict()
    tiq = db._lista(db.df_tiquetes(desde, hasta))
    tramos = db._lista(db.df_tramos(desde, hasta))
    libros = db._lista(db.df_libros(desde, hasta))
    out = []
    # 1. Recibos duplicados
    vistos: dict = {}
    for t in tiq:
        clave = R.solo_digitos(t.get("recibo"))
        if clave:
            vistos.setdefault(clave, []).append(t)
    for clave, grupo in vistos.items():
        if len(grupo) > 1:
            for t in grupo:
                otros = ", ".join(f"#{g['id']}" for g in grupo if g["id"] != t["id"])
                out.append(_h("Alta", "CONTABILIDAD", "Recibo duplicado", "combustible_tiquetes", t["id"], t.get("fecha_consumo"),
                              t.get("matricula"), f"El recibo {t.get('recibo')} también aparece en {otros}.",
                              "Verificar el soporte físico y eliminar (borrado lógico) el registro repetido."))
    # 2. Conciliación tiquete ↔ libro
    conc, sin_tiquete = R.conciliar_combustible(tiq, tramos, p.get("tolerancia_galones_pct") or 2, p.get("tolerancia_galones_abs") or 2)
    tiq_id = {t["id"]: t for t in tiq}
    for c in conc:
        t = tiq_id[c["tiquete_id"]]
        piloto = R.texto(t.get("piloto_nombre"), "el piloto")
        if c["estado"] == "SIN LIBRO DE VUELO":
            out.append(_h("Media", "CONTABILIDAD", "Documento faltante: libro de vuelo", "combustible_tiquetes", t["id"],
                          t.get("fecha_consumo"), t.get("matricula"),
                          f"Tiquete {t.get('recibo') or 's/n'} ({R.a_float(t.get('galones'), 0):,.1f} gal) sin tramo de libro que lo respalde.",
                          f"Solicitar a {piloto} la foto del libro de vuelo del {t.get('fecha_consumo')}."))
        elif c["estado"] == "DIFERENCIA GALONES":
            out.append(_h("Alta", "CONTABILIDAD", "Diferencia de galones tiquete vs. libro", "combustible_tiquetes", t["id"],
                          t.get("fecha_consumo"), t.get("matricula"),
                          f"Tiquete {R.a_float(t.get('galones'), 0):,.1f} gal vs. libro Nº {c.get('numero_libro') or '?'} "
                          f"{R.a_float(c.get('gal_libro'), 0):,.1f} gal (dif. {c.get('dif_gal'):+.1f}).",
                          "Confrontar el tiquete físico con la casilla GLS del libro y corregir el dato erróneo."))
    for t in sin_tiquete:
        out.append(_h("Media", "CONTABILIDAD", "Documento faltante: tiquete de combustible", "tramos_vuelo", t["id"], t.get("fecha"),
                      t.get("matricula"), f"El libro Nº {t.get('numero_libro') or '?'} registra tanqueo "
                      f"({R.a_float(t.get('galones_tanqueados'), 0):,.1f} gal, recibo {t.get('recibo_combustible') or 's/n'}) sin tiquete cargado.",
                      f"Solicitar a {R.texto(t.get('piloto_nombre'), 'el piloto')} la foto del tiquete o cargarlo como respaldo."))
    # 3. Sin valorizar
    hoy = db.hoy()
    for t in tiq:
        if R.a_float(t.get("valor_unit")) is None:
            f = R.parse_fecha(t.get("fecha_consumo"))
            sev = "Media" if f and (hoy - f).days > 5 else "Baja"
            out.append(_h(sev, "CONTABILIDAD", "Tiquete sin valorizar", "combustible_tiquetes", t["id"], t.get("fecha_consumo"),
                          t.get("matricula"), "Precio por galón en blanco (dato manual pendiente).",
                          "Aplicar la tarifa vigente o digitar el precio de la factura."))
    # 4. Volúmenes atípicos por aeronave
    df = pd.DataFrame(tiq)
    if not df.empty:
        df["g"] = pd.to_numeric(df["galones"], errors="coerce")
        for mat, g in df.groupby("matricula"):
            s = g["g"].std()
            if len(g) >= 5 and s and s > 0:
                z = (g["g"] - g["g"].mean()) / s
                for _, r in g[z.abs() > 2.5].iterrows():
                    out.append(_h("Media", "CONTABILIDAD", "Volumen atípico", "combustible_tiquetes", r["id"], r.get("fecha_consumo"),
                                  mat, f"{r['g']:,.1f} gal frente a un promedio de {g['g'].mean():,.1f} gal en la aeronave.",
                                  "Verificar el tiquete físico (posible error de digitación o lectura)."))
        for _, r in df[(df["g"].isna()) | (df["g"] <= 0)].iterrows():
            out.append(_h("Alta", "CONTABILIDAD", "Galones inválidos", "combustible_tiquetes", r["id"], r.get("fecha_consumo"),
                          r.get("matricula"), "El tiquete no tiene galones válidos.", "Corregir con el soporte físico."))
    # 5. TIEMPO PILOTOS
    cfg = db.cfg_tiempo_pilotos(p)
    for t in tramos:
        tp, origen = R.tiempo_pilotos_efectivo(t, cfg)
        if tp is None:
            out.append(_h("Baja", "CONTABILIDAD", "TIEMPO PILOTOS pendiente", "tramos_vuelo", t["id"], t.get("fecha"), t.get("matricula"),
                          f"Tramo {t.get('numero_vuelo') or ''} {t.get('origen') or ''}-{t.get('destino') or ''} sin TIEMPO PILOTOS.",
                          "Digitar el tiempo en Liquidación o definir la fórmula/rangos en parámetros."))
        else:
            tt, tv = R.a_float(t.get("tiempo_total_min")), R.a_float(t.get("tiempo_vuelo_min"))
            if (tt and tp > tt + 60) or (tv and tp > 3 * tv):
                out.append(_h("Media", "CONTABILIDAD", "TIEMPO PILOTOS atípico", "tramos_vuelo", t["id"], t.get("fecha"),
                              t.get("matricula"), f"TIEMPO PILOTOS {R.fmt_duracion(tp)} ({origen}) vs. vuelo {R.fmt_duracion(tv)} / "
                              f"bloque {R.fmt_duracion(tt)}.", "Revisar la regla aplicada o el valor digitado."))
        if not (R.texto(t.get("piloto_cedula")) or R.texto(t.get("piloto_nombre"))):
            out.append(_h("Media", "CONTABILIDAD", "Tripulación sin identificar", "tramos_vuelo", t["id"], t.get("fecha"),
                          t.get("matricula"), "El tramo no tiene piloto al mando asociado.", "Completar piloto en el libro de vuelo."))
        if R.a_float(t.get("tiempo_vuelo_min")) and R.a_float(t.get("tiempo_total_min")) and \
                R.a_float(t["tiempo_vuelo_min"]) > R.a_float(t["tiempo_total_min"]):
            out.append(_h("Alta", "CONTABILIDAD", "Tiempo de vuelo mayor que bloque", "tramos_vuelo", t["id"], t.get("fecha"),
                          t.get("matricula"), f"Vuelo {R.fmt_duracion(t['tiempo_vuelo_min'])} > bloque {R.fmt_duracion(t['tiempo_total_min'])}.",
                          "Corregir horas de despegue/aterrizaje o calzos."))
    for l in libros:
        if l.get("descuadre_totales"):
            out.append(_h("Alta", "CONTABILIDAD", "Descuadre fila TOTALES del libro", "libros_vuelo", l["id"], l.get("fecha"),
                          l.get("matricula"), f"El libro Nº {l.get('numero_libro') or '?'} se guardó con descuadre entre tramos y TOTALES.",
                          "Confrontar el libro físico y corregir la suma o el tramo erróneo."))
    return out


# ============================================================
# Reglas OPERACIONES
# ============================================================
_ETIQ_PEND = {k: R.PLANO_HEADERS[k] for k in R.PLANO_CLAVES}


def auditar_operaciones(desde, hasta) -> list[dict]:
    plano = db._lista(db.df_plano(desde, hasta))
    tramos = {t["id"]: t for t in db._lista(db.df_tramos(desde, hasta))}
    fichas = {f["id"]: f for f in db._lista(db.df_fichas(desde, hasta))}
    libros = db._lista(db.df_libros(desde, hasta))
    modelos = db.modelos_dict()
    aeronaves = db.aeronaves_dict()
    out = []
    fichas_usadas = set()
    for r in plano:
        pend = [p for p in json.loads(r.get("pendientes") or "[]") if p not in ("carga_real",)]
        if pend:
            out.append(_h("Media", "OPERACIONES", "Celdas pendientes en archivo plano", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), "Pendiente: " + ", ".join(_ETIQ_PEND.get(p, p) for p in pend),
                          "Completar las celdas en blanco antes del cierre mensual."))
        if r.get("ficha_id"):
            fichas_usadas.add(r["ficha_id"])
        t = tramos.get(r.get("tramo_id"))
        f = fichas.get(r.get("ficha_id"))
        if t and not f and r.get("fuente") == "SOPORTES":
            out.append(_h("Baja", "OPERACIONES", "Documento faltante: ficha de vuelo", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"Vuelo {r.get('nro_vuelo') or ''} {r.get('origen')}-{r.get('destino')} sin ficha: "
                          "tripadis/bebés no verificados.", f"Solicitar la ficha/cierre a {R.texto(t.get('piloto_nombre'), 'el piloto')} o al despacho."))
        if t and f and R.a_int(t.get("pax_total")) is not None and R.a_int(f.get("pax_total")) is not None \
                and R.a_int(t["pax_total"]) != R.a_int(f["pax_total"]):
            out.append(_h("Alta", "OPERACIONES", "Pasajeros libro vs. ficha", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"Libro {t['pax_total']} pax vs. ficha {f['pax_total']} pax en {r.get('origen')}-{r.get('destino')}.",
                          "Confrontar con el cierre de la base; el dato de PASAJEROS A BORDO se cobra en Aerocivil."))
        sillas, pax = R.a_int(r.get("sillas_ofrecidas")), R.a_int(r.get("pax_bordo"))
        if sillas and pax is not None and pax > sillas:
            out.append(_h("Alta", "OPERACIONES", "Pasajeros superan sillas", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"{pax} pax a bordo con {sillas} sillas ofrecidas.", "Corregir pasajeros o configuración del modelo."))
        co, cb, cr = R.a_float(r.get("carga_ofrecida")), R.a_float(r.get("carga_bordo")), R.a_float(r.get("carga_real"))
        if co and cb is not None and cb > co:
            out.append(_h("Media", "OPERACIONES", "Carga a bordo supera la ofrecida", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"{cb:,.0f} kg a bordo vs. {co:,.0f} kg ofrecidos.", "Verificar el reporte del grupo de despacho."))
        if co and cr is not None and cr > co:
            out.append(_h("Media", "OPERACIONES", "Carga real supera la ofrecida", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"Carga real {cr:,.0f} kg vs. ofrecida {co:,.0f} kg.", "Revisar sobrepeso reportado por la base."))
        hb, hv = R.parse_duracion(r.get("horas_bloque")), R.parse_duracion(r.get("horas_vuelo"))
        if hb is not None and hv is not None and hv > hb:
            out.append(_h("Alta", "OPERACIONES", "Horas de vuelo mayores que bloque", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), f"Vuelo {r.get('horas_vuelo')} > bloque {r.get('horas_bloque')}.", "Corregir tiempos."))
        # Consumo real vs estándar
        if t and R.a_float(t.get("lbs_comb_usado")) and R.a_float(r.get("combustible_consumido")):
            aero = aeronaves.get(R.norm_matricula(r.get("matricula"))) or {}
            factor = R.a_float((modelos.get(aero.get("modelo_codigo")) or {}).get("factor_conversion_lbs_gal"))
            if factor:
                real = R.safe_div(R.a_float(t["lbs_comb_usado"]), factor)
                std = R.a_float(r["combustible_consumido"])
                if std and abs(real - std) / std > 0.35:
                    out.append(_h("Baja", "OPERACIONES", "Consumo real vs. estándar", "operaciones_plano", r["id"], r.get("fecha_local"),
                                  r.get("matricula"), f"Real {real:,.0f} gal (libro) vs. estándar {std:,.0f} gal.",
                                  "Revisar libras usadas del libro o el estándar de la ruta."))
        if "superan el total" in R.texto(r.get("notas")):
            out.append(_h("Alta", "OPERACIONES", "Descuentos de pasajeros inconsistentes", "operaciones_plano", r["id"], r.get("fecha_local"),
                          r.get("matricula"), "Tripadis + bebés + tránsito superan el total de pasajeros.", "Revisar ficha y libro."))
    for fid, f in fichas.items():
        if fid not in fichas_usadas:
            out.append(_h("Media", "OPERACIONES", "Documento faltante: libro de vuelo", "fichas_vuelo", fid, f.get("fecha"), f.get("matricula"),
                          f"Ficha {f.get('origen') or ''}-{f.get('destino') or ''} sin tramo de libro de vuelo asociado.",
                          f"Solicitar a {R.texto(f.get('piloto_nombre'), 'el piloto')} la foto del libro de vuelo."))
    # Consecutivos de libros por matrícula
    por_mat: dict = {}
    for l in db._lista(db.df_libros(None, None)):
        n = R.a_int(l.get("numero_libro"))
        if n is not None:
            por_mat.setdefault(l.get("matricula"), []).append((n, l))
    ids_rango = {l["id"] for l in libros}
    for mat, lst in por_mat.items():
        lst.sort(key=lambda x: x[0])
        for (a, la), (b, lb) in zip(lst, lst[1:]):
            if lb["id"] not in ids_rango:
                continue
            if b == a:
                out.append(_h("Alta", "OPERACIONES", "Consecutivo de libro duplicado", "libros_vuelo", lb["id"], lb.get("fecha"), mat,
                              f"El libro Nº {b} está registrado más de una vez.", "Verificar si es un doble registro."))
            elif b - a > 1:
                out.append(_h("Media", "OPERACIONES", "Salto de consecutivo de libro", "libros_vuelo", lb["id"], lb.get("fecha"), mat,
                              f"Del libro Nº {a} se pasa al Nº {b}: faltan {b - a - 1} hoja(s).", "Solicitar las hojas faltantes al piloto."))
    # Desequilibrio de combustible entre tramos del mismo libro
    por_libro: dict = {}
    for t in tramos.values():
        por_libro.setdefault(t["libro_id"], []).append(t)
    for lid, lst in por_libro.items():
        lst.sort(key=lambda x: (x.get("orden") or 0))
        mat = lst[0].get("matricula")
        factor = R.a_float((modelos.get((aeronaves.get(R.norm_matricula(mat)) or {}).get("modelo_codigo")) or {}).get("factor_conversion_lbs_gal"))
        for a, b in zip(lst, lst[1:]):
            lleg, sal = R.a_float(a.get("lbs_comb_llegada")), R.a_float(b.get("lbs_comb_salida"))
            if lleg is None or sal is None:
                continue
            tanq = R.a_float(b.get("galones_tanqueados"), 0) or 0
            esperado = lleg + (tanq * factor if factor else 0)
            tol = max(60.0, 0.1 * esperado)
            if (tanq == 0 and abs(sal - lleg) > 60) or (tanq > 0 and factor and abs(sal - esperado) > tol):
                out.append(_h("Media", "OPERACIONES", "Desequilibrio de combustible entre tramos", "tramos_vuelo", b["id"], b.get("fecha"), mat,
                              f"Llegada tramo {a.get('orden')}: {lleg:,.0f} lbs → salida tramo {b.get('orden')}: {sal:,.0f} lbs "
                              f"(tanqueo {tanq:,.1f} gal; esperado ≈ {esperado:,.0f} lbs).", "Revisar libras y tanqueo en el libro físico."))
        for t in lst:
            a, b, u = R.a_float(t.get("lbs_comb_salida")), R.a_float(t.get("lbs_comb_llegada")), R.a_float(t.get("lbs_comb_usado"))
            if None not in (a, b, u) and abs((a - b) - u) > 1:
                out.append(_h("Media", "OPERACIONES", "Libras usadas inconsistentes", "tramos_vuelo", t["id"], t.get("fecha"), mat,
                              f"Usado {u:,.0f} lbs ≠ salida − llegada ({a - b:,.0f} lbs).", "Corregir el tramo."))
    return out


# ============================================================
# Integridad e imágenes (revisoría fiscal)
# ============================================================
def auditar_integridad(desde, hasta) -> list[dict]:
    out = []
    consultas = {"combustible_tiquetes": db.df_tiquetes(desde, hasta), "libros_vuelo": db.df_libros(desde, hasta),
                 "tramos_vuelo": db.df_tramos(desde, hasta), "fichas_vuelo": db.df_fichas(desde, hasta),
                 "operaciones_plano": db.df_plano(desde, hasta)}
    for tabla, df in consultas.items():
        for r in db._lista(df):
            datos = {k: r.get(k) for k in db.CAMPOS[tabla]}
            if r.get("hash_integridad") and R.huella(datos) != r["hash_integridad"]:
                fecha = r.get("fecha_consumo") or r.get("fecha") or r.get("fecha_local")
                out.append(_h("Alta", "GENERAL", "Registro alterado fuera del sistema", tabla, r["id"], fecha, r.get("matricula"),
                              "El hash de integridad no coincide con los datos actuales (modificación directa sobre la base).",
                              "Comparar con el respaldo más reciente y con la bitácora; escalar a Gerencia."))
    for d in db._lista(db.df_documentos(desde, hasta)):
        estado = V.verificar_evidencia(d.get("ruta_archivo"), d.get("sha256"))
        if estado != "OK":
            out.append(_h("Alta", "GENERAL", f"Evidencia {estado.lower()}", "documentos", d["id"], d.get("fecha_doc"), d.get("matricula"),
                          f"Archivo {d.get('ruta_archivo')} {('no existe' if estado == 'FALTANTE' else 'fue modificado (SHA-256 distinto)')}.",
                          "Restaurar la evidencia desde el respaldo o solicitar nuevamente el soporte."))
    cadena = db.verificar_cadena_auditoria()
    if not cadena["ok"]:
        out.append(_h("Alta", "GENERAL", "Bitácora de auditoría alterada", "audit_logs", cadena["primer_error_id"], None, None,
                      cadena["mensaje"], "Escalar de inmediato: la bitácora inmutable fue intervenida fuera del sistema."))
    return out


def auditar_imagenes_ia(desde, hasta, n_max=12, progreso=None) -> tuple[list[dict], int]:
    """Re-lee con Claude Vision las evidencias menos auditadas y compara contra lo guardado."""
    ok, _ = V.ia_disponible()
    if not ok or n_max <= 0:
        return [], 0
    docs = db._lista(db.df_documentos(desde, hasta))
    docs.sort(key=lambda d: (d.get("ultima_auditoria_ia") or "", d.get("fecha_creacion") or ""))
    docs = docs[:int(n_max)]
    out, revisados = [], 0
    for i, d in enumerate(docs):
        if progreso:
            progreso((i + 1) / max(1, len(docs)), f"Re-leyendo {d.get('ruta_archivo', '').split('/')[-1]}")
        r = V.releer_documento(d["tipo"], d["ruta_archivo"])
        if not r.get("exito"):
            continue
        revisados += 1
        out.extend(_comparar_relectura(d, r))
        with db.conexion(True) as con:
            con.execute("UPDATE documentos SET ultima_auditoria_ia=? WHERE id=?", (db.ahora_txt(), d["id"]))
    return out, revisados


def _comparar_relectura(doc: dict, r: dict) -> list[dict]:
    out = []

    def dif(tabla, rid, fecha, mat, campo, img, base):
        out.append(_h("Alta", "GENERAL", "Discrepancia imagen vs. base (re-lectura IA)", tabla, rid, fecha, mat,
                      f"{campo}: imagen «{img}» · base «{base}».",
                      "Abrir la evidencia, confirmar el valor correcto y corregir con motivo (la IA también puede equivocarse)."))
    if doc["tipo"] == "COMB":
        t = db.consulta("SELECT * FROM combustible_tiquetes WHERE documento_id=? AND eliminado=0", (doc["id"],))
        if not t:
            return out
        t, d = t[0], r["datos"]
        if d.get("galones") is not None and R.a_float(t.get("galones")) is not None and abs(d["galones"] - t["galones"]) > 0.6:
            dif("combustible_tiquetes", t["id"], t["fecha_consumo"], t["matricula"], "Galones", d["galones"], t["galones"])
        if R.solo_digitos(d.get("recibo")) and R.solo_digitos(d.get("recibo")) != R.solo_digitos(t.get("recibo")):
            dif("combustible_tiquetes", t["id"], t["fecha_consumo"], t["matricula"], "Recibo", d.get("recibo"), t.get("recibo"))
        if d.get("matricula") and d["matricula"] != R.norm_matricula(t.get("matricula")):
            dif("combustible_tiquetes", t["id"], t["fecha_consumo"], t["matricula"], "Matrícula", d["matricula"], t["matricula"])
        if d.get("fecha_consumo") and d["fecha_consumo"].isoformat() != t.get("fecha_consumo"):
            dif("combustible_tiquetes", t["id"], t["fecha_consumo"], t["matricula"], "Fecha", d["fecha_consumo"], t["fecha_consumo"])
    elif doc["tipo"] == "LV":
        l = db.consulta("SELECT * FROM libros_vuelo WHERE documento_id=? AND eliminado=0", (doc["id"],))
        if not l:
            return out
        l = l[0]
        tramos = db.consulta("SELECT * FROM tramos_vuelo WHERE libro_id=? AND eliminado=0 ORDER BY orden", (l["id"],))
        cab = r.get("cabecera") or {}
        if cab.get("numero_libro") and cab["numero_libro"] != R.solo_digitos(l.get("numero_libro")):
            dif("libros_vuelo", l["id"], l["fecha"], l["matricula"], "Número de libro", cab["numero_libro"], l.get("numero_libro"))
        if len(r.get("tramos") or []) != len(tramos):
            dif("libros_vuelo", l["id"], l["fecha"], l["matricula"], "Cantidad de tramos", len(r.get("tramos") or []), len(tramos))
        for ti, tb in zip(r.get("tramos") or [], tramos):
            for campo, etq in (("pax_total", "Pax total"), ("tiempo_vuelo_min", "Tiempo de vuelo"), ("tiempo_total_min", "Tiempo total")):
                a, b = ti.get(campo), tb.get(campo)
                if a is not None and b is not None and int(a) != int(b):
                    fa = R.fmt_duracion(a) if "tiempo" in campo else a
                    fb = R.fmt_duracion(b) if "tiempo" in campo else b
                    dif("tramos_vuelo", tb["id"], l["fecha"], l["matricula"], f"{etq} (tramo {tb.get('orden')})", fa, fb)
    else:
        fichas = db.consulta("SELECT * FROM fichas_vuelo WHERE documento_id=? AND eliminado=0 ORDER BY id", (doc["id"],))
        for vi, fb in zip(r.get("vuelos") or [], fichas):
            for campo, etq in (("pax_total", "Pax total"), ("tripadis", "Tripadis"), ("bebes", "Bebés"), ("pax_transito", "Tránsito")):
                a, b = vi.get(campo), fb.get(campo)
                if a is not None and b is not None and int(a) != int(b):
                    dif("fichas_vuelo", fb["id"], fb["fecha"], fb["matricula"], etq, a, b)
    return out


# ============================================================
# Responsables, resumen y orquestación
# ============================================================
def _responsables(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    cre_ced, cre_nom, ed_ced, ed_nom = [], [], [], []
    cache: dict = {}
    for _, h in df.iterrows():
        tabla, rid = h.get("TABLA"), h.get("REGISTRO_ID")
        clave = (tabla, rid)
        if clave not in cache:
            c = {"cc": None, "cn": None, "ec": None, "en": None}
            if tabla in TABLAS_NEGOCIO and R.solo_digitos(rid):
                if tabla == "tramos_vuelo":
                    f = db.consulta("SELECT l.creado_por_cedula, l.creado_por_nombre FROM tramos_vuelo t JOIN libros_vuelo l "
                                    "ON l.id=t.libro_id WHERE t.id=?", (int(rid),))
                elif tabla == "operaciones_plano":
                    # Fila generada desde soportes: responde quien confirmó el libro; fila manual/importada: quien la creó.
                    f = db.consulta("SELECT COALESCE(l.creado_por_cedula, p.creado_por_cedula) creado_por_cedula, "
                                    "COALESCE(l.creado_por_nombre, p.creado_por_nombre) creado_por_nombre FROM operaciones_plano p "
                                    "LEFT JOIN tramos_vuelo t ON t.id=p.tramo_id LEFT JOIN libros_vuelo l ON l.id=t.libro_id "
                                    "WHERE p.id=?", (int(rid),))
                else:
                    f = db.consulta(f"SELECT creado_por_cedula, creado_por_nombre FROM {tabla} WHERE id=?", (int(rid),))
                if f:
                    c["cc"], c["cn"] = f[0]["creado_por_cedula"], f[0]["creado_por_nombre"]
                e = db.consulta("SELECT usuario_cedula, usuario_nombre FROM audit_logs WHERE tabla_afectada=? AND registro_id=? "
                                "AND accion IN ('EDITAR','CORRECCION_IA') ORDER BY id DESC LIMIT 1", (tabla, str(rid)))
                if e:
                    c["ec"], c["en"] = e[0]["usuario_cedula"], e[0]["usuario_nombre"]
            elif tabla == "documentos" and R.solo_digitos(rid):
                f = db.consulta("SELECT creado_por_cedula, creado_por_nombre FROM documentos WHERE id=?", (int(rid),))
                if f:
                    c["cc"], c["cn"] = f[0]["creado_por_cedula"], f[0]["creado_por_nombre"]
            cache[clave] = c
        c = cache[clave]
        cre_ced.append(c["cc"]); cre_nom.append(c["cn"]); ed_ced.append(c["ec"]); ed_nom.append(c["en"])
    df = df.copy()
    df["RESPONSABLE_CEDULA"], df["RESPONSABLE_NOMBRE"] = cre_ced, cre_nom
    df["ULTIMO_EDITOR_CEDULA"], df["ULTIMO_EDITOR_NOMBRE"] = ed_ced, ed_nom
    return df


def historial_responsables(hallazgos: pd.DataFrame, limite=5000) -> pd.DataFrame:
    if hallazgos is None or hallazgos.empty:
        return pd.DataFrame()
    pares = {(r["TABLA"], str(r["REGISTRO_ID"])) for _, r in hallazgos.iterrows() if r.get("TABLA") and r.get("REGISTRO_ID")}
    filas = []
    for tabla, rid in list(pares)[:1500]:
        filas.extend(db.consulta("SELECT id, fecha_hora, usuario_cedula, usuario_nombre, rol, modulo, tabla_afectada, registro_id, "
                                 "accion, campo_modificado, valor_anterior, valor_nuevo, motivo_cambio FROM audit_logs "
                                 "WHERE tabla_afectada=? AND registro_id=? ORDER BY id", (tabla, rid)))
    df = pd.DataFrame(filas[:limite])
    return df.sort_values(["tabla_afectada", "registro_id", "id"]) if not df.empty else df


def _resumen_determinista(ambito, df: pd.DataFrame, desde, hasta) -> str:
    if df.empty:
        return f"Auditoría {ambito} del {desde} al {hasta}: sin hallazgos. Todos los registros pasan las validaciones."
    lineas = [f"Auditoría {ambito} del {desde} al {hasta}: {len(df)} hallazgo(s) "
              f"({(df['SEVERIDAD'] == 'Alta').sum()} altos, {(df['SEVERIDAD'] == 'Media').sum()} medios, "
              f"{(df['SEVERIDAD'] == 'Baja').sum()} bajos).", "Principales categorías:"]
    for cat, n in df["CATEGORIA"].value_counts().head(8).items():
        lineas.append(f"• {cat}: {n}")
    lineas.append("Prioridad: atender primero los hallazgos de severidad Alta y los documentos faltantes antes del cierre mensual.")
    return "\n".join(lineas)


def resumen_con_ia(ambito, df: pd.DataFrame, desde, hasta, incluir_responsables: bool) -> tuple[str, bool]:
    base = _resumen_determinista(ambito, df, desde, hasta)
    ok, _ = V.ia_disponible()
    if not ok or df.empty:
        return base, False
    cols = ["SEVERIDAD", "CATEGORIA", "TABLA", "REGISTRO_ID", "FECHA", "MATRICULA", "DETALLE"]
    if incluir_responsables:
        cols += ["RESPONSABLE_NOMBRE", "ULTIMO_EDITOR_NOMBRE"]
    muestra = df.sort_values("SEVERIDAD", key=lambda s: s.map(_SEV_ORDEN)).head(60)[[c for c in cols if c in df.columns]]
    conteo = df.groupby(["SEVERIDAD", "CATEGORIA"]).size().reset_index(name="N").to_dict("records")
    restriccion = ("Incluye a quién llamar a corregir (por nombre) y recomendaciones por responsable." if incluir_responsables else
                   "NO evalúes ni califiques el desempeño de empleados; enfócate en datos, soportes y acciones del área.")
    prompt = (f"Ámbito: {ambito}. Periodo {desde} a {hasta}.\nConteo de hallazgos: {json.dumps(conteo, ensure_ascii=False)}\n"
              f"Muestra de hallazgos (máx. 60): {muestra.to_json(orient='records', force_ascii=False)}\n\n"
              "Redacta un informe ejecutivo breve (máx. 250 palabras) con: 1) diagnóstico, 2) causas probables, "
              f"3) acciones priorizadas para antes del cierre, 4) documentos faltantes a solicitar. {restriccion} "
              "Usa viñetas simples con '•'. No inventes cifras distintas a las entregadas.")
    texto, err = V.llamar_claude_texto(prompt)
    if err or not texto:
        return base + f"\n(Resumen IA no disponible: {err})", False
    return texto.strip() + "\n\n— Resumen cuantitativo —\n" + base, True


def ejecutar_auditoria(ambito: str, desde, hasta, usuario: dict, revisar_imagenes=False, progreso=None) -> dict:
    """ambito: CONTABILIDAD | OPERACIONES | GENERAL"""
    filas: list[dict] = []
    revisadas = 0
    if ambito in ("CONTABILIDAD", "GENERAL"):
        filas += auditar_contabilidad(desde, hasta)
    if ambito in ("OPERACIONES", "GENERAL"):
        filas += auditar_operaciones(desde, hasta)
    if ambito == "GENERAL":
        filas += auditar_integridad(desde, hasta)
        if revisar_imagenes:
            extra, revisadas = auditar_imagenes_ia(desde, hasta, int(db.get_param("ia_max_imagenes_auditoria", 12) or 12), progreso)
            filas += extra
    df = pd.DataFrame(filas, columns=[c for c in db.COLS_HALLAZGO if c not in ("RESPONSABLE_CEDULA", "RESPONSABLE_NOMBRE",
                                                                               "ULTIMO_EDITOR_CEDULA", "ULTIMO_EDITOR_NOMBRE")])
    df = _responsables(df) if not df.empty else _vacio()
    if not df.empty:
        df = df.sort_values(["SEVERIDAD", "FECHA"], key=lambda s: s.map(_SEV_ORDEN) if s.name == "SEVERIDAD" else s.fillna(""),
                            ascending=[True, False]).reset_index(drop=True)
    resumen, uso_ia = resumen_con_ia(ambito, df, desde, hasta, incluir_responsables=(ambito == "GENERAL"))
    aid = db.guardar_auditoria(ambito, desde, hasta, df, resumen, uso_ia, usuario)
    historial = historial_responsables(df) if ambito == "GENERAL" else pd.DataFrame()
    return {"hallazgos": df, "resumen": resumen, "uso_ia": uso_ia, "auditoria_id": aid, "historial": historial,
            "imagenes_revisadas": revisadas}


# ============================================================
# Desempeño de empleados (EXCLUSIVO GERENCIA)
# ============================================================
def desempeno_empleados(desde, hasta) -> pd.DataFrame:
    usuarios = db.df_usuarios()
    usuarios = usuarios[usuarios["rol"].isin(["PILOTO", "CONTABILIDAD", "OPERACIONES"])]
    cols = ["CEDULA", "NOMBRE", "ROL", "REGISTROS CREADOS", "EDICIONES", "CORRECCIONES A LA IA", "CIERRES", "REPROCESOS",
            "DEMORA PROMEDIO (DÍAS)", "% A TIEMPO", "HALLAZGOS ATRIBUIDOS", "TASA DE ERROR %"]
    if usuarios.empty:
        return pd.DataFrame(columns=cols)
    logs = db.df_auditoria(desde, hasta, limite=200000)
    negocio = logs[logs["tabla_afectada"].isin(TABLAS_NEGOCIO)] if not logs.empty else logs
    dias_ok = int(db.get_param("dias_oportunidad", 1) or 1)
    d, h = db._rango(desde, hasta)
    registros = []
    for tabla, fcol in (("combustible_tiquetes", "fecha_consumo"), ("libros_vuelo", "fecha"), ("fichas_vuelo", "fecha"),
                        ("operaciones_plano", "fecha_local")):
        registros += db.consulta(f"SELECT '{tabla}' tabla, id, creado_por_cedula, fecha_creacion, {fcol} fecha_ref FROM {tabla} "
                                 f"WHERE eliminado=0 AND substr(fecha_creacion,1,10) BETWEEN ? AND ?", (d, h))
    reg = pd.DataFrame(registros, columns=["tabla", "id", "creado_por_cedula", "fecha_creacion", "fecha_ref"])
    if not reg.empty:
        reg["demora"] = (pd.to_datetime(reg["fecha_creacion"].str[:10], errors="coerce") -
                         pd.to_datetime(reg["fecha_ref"], errors="coerce")).dt.days
    cierres = db.df_sql("SELECT cerrado_por_cedula, fecha, fecha_cierre FROM consolidados_diarios WHERE substr(fecha_cierre,1,10) BETWEEN ? AND ?", (d, h))
    if not cierres.empty:
        cierres["demora"] = (pd.to_datetime(cierres["fecha_cierre"].str[:10], errors="coerce") -
                             pd.to_datetime(cierres["fecha"], errors="coerce")).dt.days
    hall = db.df_hallazgos_recientes(desde, hasta)
    if not hall.empty:
        # Calidad = errores atribuibles a la persona. Se excluyen pendientes de flujo e intervenciones externas.
        hall = hall[~hall["categoria"].isin(CATEGORIAS_SIN_CULPA)]
    creadores = {}
    if not reg.empty:
        creadores = {(r.tabla, str(r.id)): r.creado_por_cedula for r in reg.itertuples()}
    reproc = {}
    if not negocio.empty:
        ed = negocio[negocio["accion"] == "EDITAR"]
        for r in ed.itertuples():
            c = creadores.get((r.tabla_afectada, str(r.registro_id)))
            if c:
                reproc[c] = reproc.get(c, 0) + 1
    filas = []
    for u in usuarios.itertuples():
        ced = u.cedula
        lu = negocio[negocio["usuario_cedula"] == ced] if not negocio.empty else negocio
        creados = int((lu["accion"] == "CREAR").sum()) if not lu.empty else 0
        ediciones = int((lu["accion"] == "EDITAR").sum()) if not lu.empty else 0
        correcciones = int((lu["accion"] == "CORRECCION_IA").sum()) if not lu.empty else 0
        n_cierres = int(logs[(logs["usuario_cedula"] == ced) & logs["accion"].str.startswith("CIERRE")].shape[0]) if not logs.empty else 0
        demoras = []
        if not reg.empty:
            demoras += reg.loc[reg["creado_por_cedula"] == ced, "demora"].dropna().tolist()
        if not cierres.empty:
            demoras += cierres.loc[cierres["cerrado_por_cedula"] == ced, "demora"].dropna().tolist()
        demora = round(sum(demoras) / len(demoras), 2) if demoras else None
        a_tiempo = round(100.0 * sum(1 for x in demoras if x <= dias_ok) / len(demoras), 1) if demoras else None
        n_hall = int(((hall["responsable_cedula"] == ced) | (hall["ultimo_editor_cedula"] == ced)).sum()) if not hall.empty else 0
        base = max(1, creados + ediciones)
        filas.append({"CEDULA": ced, "NOMBRE": u.nombre, "ROL": u.rol, "REGISTROS CREADOS": creados, "EDICIONES": ediciones,
                      "CORRECCIONES A LA IA": correcciones, "CIERRES": n_cierres, "REPROCESOS": reproc.get(ced, 0),
                      "DEMORA PROMEDIO (DÍAS)": demora, "% A TIEMPO": a_tiempo, "HALLAZGOS ATRIBUIDOS": n_hall,
                      "TASA DE ERROR %": round(100.0 * n_hall / base, 1) if (creados + ediciones) else None})
    return pd.DataFrame(filas, columns=cols)

# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Visión artificial y evidencias (vision_ocr.py)

  * Carga segura de imágenes (tamaño, formato, bombas de descompresión, rotación EXIF, re-codificación JPEG
    que elimina metadatos y cualquier contenido incrustado).
  * Filtro anti-borrosidad capa 1 (solo Pillow + NumPy): resolución mínima, varianza del Laplaciano 2D,
    brillo y contraste sobre el área del documento.
  * Capa 2: lectura con Claude Vision. El Libro de Vuelo MO-A3-004 se divide en 3 zonas ampliadas
    (encabezado/motores, tabla de tramos y tripulación/mantenimiento) que se leen en paralelo.
  * Almacenamiento seriado: evidencias/AAAA/MM_Mes/DD/{01_Combustible|02_Libros_Vuelo|03_Fichas_Vuelo}/
    PREFIJO_AAAAMMDD_HHMM_MATRICULA_Vnn_{C|N}.jpg, escritura atómica y verificación SHA-256.
Si ANTHROPIC_API_KEY no está configurada o falla la conexión, todo sigue funcionando en modo manual asistido.
"""
from __future__ import annotations

import base64
import hashlib
import importlib.util
import io
import json
import logging
import os
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import numpy as np
from PIL import Image, ImageOps, UnidentifiedImageError

import database as db
import rules_engine as R

log = logging.getLogger("moon.vision")

Image.MAX_IMAGE_PIXELS = 80_000_000          # protección contra bombas de descompresión
MAX_BYTES = 25 * 1024 * 1024
FORMATOS_PERMITIDOS = {"JPEG", "PNG", "WEBP", "MPO", "BMP", "TIFF"}

TIPOS_DOC = {
    "COMB": {"prefijo": "COMB", "carpeta": "01_Combustible", "nombre": "Tiquete de combustible (Terpel)", "icono": "⛽"},
    "LV": {"prefijo": "LV", "carpeta": "02_Libros_Vuelo", "nombre": "Libro de vuelo MO-A3-004", "icono": "📘"},
    "FV": {"prefijo": "FV", "carpeta": "03_Fichas_Vuelo", "nombre": "Ficha de vuelo / Cierre y Manifiesto P&B", "icono": "🧾"},
}
MODELOS_RESPALDO = ["claude-sonnet-5-5", "claude-opus-5-5", "claude-haiku-5-5"]


class ImagenInvalida(Exception):
    pass


# ============================================================
# Carga segura y calidad (capa 1: Pillow + NumPy, sin OpenCV)
# ============================================================
def abrir_imagen(raw: bytes) -> Image.Image:
    if not raw:
        raise ImagenInvalida("El archivo está vacío.")
    if len(raw) > MAX_BYTES:
        raise ImagenInvalida("La imagen supera 25 MB. Tómala de nuevo con menor resolución.")
    try:
        with Image.open(io.BytesIO(raw)) as prueba:
            formato = prueba.format
            prueba.verify()
        if formato not in FORMATOS_PERMITIDOS:
            raise ImagenInvalida(f"Formato {formato} no permitido. Usa JPG o PNG.")
        img = Image.open(io.BytesIO(raw))
        img.load()
        img = ImageOps.exif_transpose(img)
        if img.mode in ("RGBA", "LA", "P"):
            fondo = Image.new("RGB", img.size, (255, 255, 255))
            img = img.convert("RGBA")
            fondo.paste(img, mask=img.split()[-1])
            return fondo
        return img.convert("RGB")
    except ImagenInvalida:
        raise
    except Image.DecompressionBombError:
        raise ImagenInvalida("La imagen es demasiado grande para procesarla con seguridad.") from None
    except (UnidentifiedImageError, OSError, ValueError, SyntaxError):
        raise ImagenInvalida("No se pudo abrir la imagen. Usa una foto JPG o PNG.") from None


def recorte_documento(img: Image.Image) -> tuple[Image.Image, tuple]:
    """Elimina bordes oscuros (pantallazos, mesa, carpeta) para analizar solo el papel."""
    gris = np.asarray(img.convert("L").resize((max(1, img.width // 4), max(1, img.height // 4))), dtype=np.float32)
    filas = np.where(gris.mean(axis=1) > 110)[0]
    if len(filas) < gris.shape[0] * 0.2:
        return img, (0, 0, img.width, img.height)
    banda = gris[filas.min():filas.max() + 1, :]
    cols = np.where(banda.mean(axis=0) > 110)[0]
    if len(cols) < gris.shape[1] * 0.2:
        return img, (0, 0, img.width, img.height)
    caja = (int(cols.min() * 4), int(filas.min() * 4), min(img.width, int((cols.max() + 1) * 4)),
            min(img.height, int((filas.max() + 1) * 4)))
    if (caja[2] - caja[0]) * (caja[3] - caja[1]) < 0.25 * img.width * img.height:
        return img, (0, 0, img.width, img.height)
    return img.crop(caja), caja


def varianza_laplaciana(gris: np.ndarray) -> float:
    """Varianza del Laplaciano 2D (kernel 0,1,0 / 1,-4,1 / 0,1,0) con NumPy puro."""
    g = gris.astype(np.float32)
    lap = (-4.0 * g[1:-1, 1:-1] + g[:-2, 1:-1] + g[2:, 1:-1] + g[1:-1, :-2] + g[1:-1, 2:])
    return float(lap.var())


def evaluar_calidad(img: Image.Image, min_lado: int = 800, umbral_nitidez: float = 45.0) -> dict:
    ancho, alto = img.size
    doc, _ = recorte_documento(img)
    ref = doc.copy()
    ref.thumbnail((1400, 1400))                      # escala de referencia para que el umbral sea estable
    gris = np.asarray(ref.convert("L"), dtype=np.float32)
    nitidez = varianza_laplaciana(gris)
    brillo, contraste = float(gris.mean()), float(gris.std())
    motivos = []
    if max(ancho, alto) < int(min_lado):
        motivos.append(f"Resolución insuficiente ({max(ancho, alto)} px). Se requieren al menos {int(min_lado)} px en el lado mayor.")
    if nitidez < float(umbral_nitidez):
        motivos.append("La foto está borrosa o movida. Apoya los codos, enfoca el papel y repite la foto.")
    if brillo < 60:
        motivos.append("La foto está muy oscura. Busca más luz o activa el flash.")
    if brillo > 248:
        motivos.append("La foto está sobreexpuesta (demasiada luz). Evita reflejos directos.")
    if contraste < 18:
        motivos.append("La foto tiene muy poco contraste: el texto no se distingue del fondo.")
    return {"ok": not motivos, "motivos": motivos, "ancho": ancho, "alto": alto, "nitidez": round(nitidez, 1),
            "brillo": round(brillo, 1), "contraste": round(contraste, 1)}


def jpeg_bytes(img: Image.Image, lado_max: int = 2200, calidad: int = 90) -> bytes:
    """Re-codifica a JPEG limpio (sin EXIF ni datos incrustados)."""
    copia = img.copy()
    copia.thumbnail((lado_max, lado_max), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    copia.convert("RGB").save(buf, "JPEG", quality=calidad, optimize=True)
    return buf.getvalue()


def zonas_libro(img: Image.Image) -> dict:
    """Recorte inteligente por zonas del Libro de Vuelo MO-A3-004 (con solapes para no cortar filas)."""
    doc, _ = recorte_documento(img)
    w, h = doc.size
    definicion = {"encabezado": (0.0, 0.34), "tramos": (0.14, 0.56), "tripulacion_oma": (0.40, 1.0)}
    zonas = {}
    for nombre, (y0, y1) in definicion.items():
        z = doc.crop((0, int(h * y0), w, int(h * y1)))
        if z.width < 1800:
            factor = min(2.5, 1800 / max(1, z.width))
            z = z.resize((int(z.width * factor), int(z.height * factor)), Image.Resampling.LANCZOS)
        zonas[nombre] = z
    return zonas


# ============================================================
# Evidencias seriadas
# ============================================================
def carpeta_evidencia(tipo: str, fecha) -> Path:
    f = R.parse_fecha(fecha) or db.hoy()
    return db.EVIDENCIAS_DIR / f"{f.year:04d}" / f"{f.month:02d}_{R.MESES_ES[f.month]}" / f"{f.day:02d}" / TIPOS_DOC[tipo]["carpeta"]


def nombre_evidencia(tipo, fecha, hora_min, matricula, seq, charter) -> str:
    f = R.parse_fecha(fecha) or db.hoy()
    hhmm = f"{int(hora_min) // 60:02d}{int(hora_min) % 60:02d}" if not R.es_vacio(hora_min) else "0000"
    mat = re.sub(r"[^A-Z0-9]", "", R.matricula_compacta(matricula)) or "SINMAT"
    return f"{TIPOS_DOC[tipo]['prefijo']}_{f:%Y%m%d}_{hhmm}_{mat}_V{int(seq):02d}_{'C' if charter else 'N'}.jpg"


def ruta_segura(rel) -> Path | None:
    """Resuelve una ruta relativa SOLO si queda dentro de DATA_DIR (evita path traversal)."""
    if R.es_vacio(rel):
        return None
    try:
        p = (db.DATA_DIR / str(rel)).resolve()
        p.relative_to(db.DATA_DIR)
        return p
    except (ValueError, OSError):
        return None


def guardar_evidencia(jpeg: bytes, tipo: str, fecha, hora_min, matricula, seq: int, charter: bool) -> dict:
    carpeta = carpeta_evidencia(tipo, fecha)
    carpeta.mkdir(parents=True, exist_ok=True)
    seq = max(1, int(seq or 1))
    while True:
        nombre = nombre_evidencia(tipo, fecha, hora_min, matricula, seq, charter)
        destino = carpeta / nombre
        if not destino.exists():
            break
        seq += 1
    fd, tmp = tempfile.mkstemp(dir=carpeta, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(jpeg)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, destino)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    sha = hashlib.sha256(destino.read_bytes()).hexdigest()
    if sha != hashlib.sha256(jpeg).hexdigest():
        destino.unlink(missing_ok=True)
        raise IOError("La verificación de integridad del archivo falló. Intente de nuevo.")
    return {"ruta_archivo": str(destino.relative_to(db.DATA_DIR)).replace("\\", "/"), "sha256": sha, "nombre": nombre, "seq": seq}


def eliminar_evidencia(rel):
    p = ruta_segura(rel)
    if p and p.is_file():
        p.unlink(missing_ok=True)


def leer_evidencia(rel) -> bytes | None:
    p = ruta_segura(rel)
    return p.read_bytes() if p and p.is_file() else None


def verificar_evidencia(rel, sha) -> str:
    p = ruta_segura(rel)
    if not p or not p.is_file():
        return "FALTANTE"
    if sha and hashlib.sha256(p.read_bytes()).hexdigest() != sha:
        return "ALTERADA"
    return "OK"


# ============================================================
# Claude Vision (capa 2)
# ============================================================
def modelo_configurado() -> str:
    return str(db.secret("ANTHROPIC_MODEL", MODELOS_RESPALDO[0]))


def ia_disponible() -> tuple[bool, str]:
    if not db.secret("ANTHROPIC_API_KEY"):
        return False, "La lectura automática no está configurada (falta ANTHROPIC_API_KEY). Puedes registrar los datos a mano."
    if importlib.util.find_spec("anthropic") is None:
        return False, "Falta instalar la librería anthropic. Puedes registrar los datos a mano."
    return True, f"Lectura automática activa ({modelo_configurado()})."


def extraer_json(texto: str) -> dict:
    if not texto:
        return {}
    limpio = re.sub(r"```(?:json)?", "", texto)
    ini, fin = limpio.find("{"), limpio.rfind("}")
    if ini < 0 or fin <= ini:
        return {}
    try:
        r = json.loads(limpio[ini:fin + 1])
        return r if isinstance(r, dict) else {}
    except json.JSONDecodeError:
        return {}


def _llamar(contenido: list, max_tokens: int, system: str) -> tuple[str | None, str | None]:
    ok, msg = ia_disponible()
    if not ok:
        return None, msg
    import anthropic
    cliente = anthropic.Anthropic(api_key=db.secret("ANTHROPIC_API_KEY"), timeout=90.0, max_retries=2)
    modelos = [modelo_configurado()] + [m for m in MODELOS_RESPALDO if m != modelo_configurado()]
    ultimo_error = None
    for modelo in modelos:
        try:
            r = cliente.messages.create(model=modelo, max_tokens=max_tokens, system=system,
                                        messages=[{"role": "user", "content": contenido}])
            texto = "".join(getattr(b, "text", "") for b in r.content if getattr(b, "type", "") == "text")
            return texto, None
        except anthropic.NotFoundError:
            ultimo_error = f"Modelo {modelo} no disponible."
            continue
        except anthropic.AuthenticationError:
            return None, "La clave ANTHROPIC_API_KEY no es válida. Registra los datos a mano y avisa a Gerencia."
        except (anthropic.APIConnectionError, anthropic.APITimeoutError):
            return None, "No hubo conexión con la lectura automática. Intenta de nuevo o registra a mano."
        except anthropic.RateLimitError:
            return None, "La lectura automática está saturada en este momento. Intenta en un minuto o registra a mano."
        except anthropic.APIStatusError as e:
            log.warning("Error de API Claude: %s", e)
            ultimo_error = "El servicio de lectura respondió con un error. Registra los datos a mano."
            if getattr(e, "status_code", 500) == 400:
                continue
            return None, ultimo_error
        except Exception:
            log.exception("Fallo inesperado en Claude")
            return None, "Fallo inesperado en la lectura automática. Registra los datos a mano."
    return None, ultimo_error or "No fue posible usar la lectura automática."


SISTEMA_JSON = ("Eres un digitador experto en documentos aeronáuticos colombianos de Moon Aerolínea. "
                "Respondes ÚNICAMENTE con un objeto JSON válido, sin texto adicional ni bloques de código. "
                "Nunca inventas datos: si un valor no se ve o no aparece, usas null.")


def llamar_claude_json(imagenes: list[bytes], prompt: str, max_tokens: int = 4096) -> tuple[dict | None, str | None]:
    contenido = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                              "data": base64.b64encode(b).decode()}} for b in imagenes]
    contenido.append({"type": "text", "text": prompt})
    texto, err = _llamar(contenido, max_tokens, SISTEMA_JSON)
    if err:
        return None, err
    datos = extraer_json(texto)
    if not datos:
        return None, "No pude interpretar el documento. Repite la foto más recta y con buena luz."
    return datos, None


def llamar_claude_texto(prompt: str, max_tokens: int = 2500) -> tuple[str | None, str | None]:
    return _llamar([{"type": "text", "text": prompt}], max_tokens,
                   "Eres revisor fiscal y auditor aeronáutico de Moon Aerolínea. Escribes en español claro, "
                   "concreto y profesional, con recomendaciones accionables. No inventas cifras.")


_REGLAS = """
Reglas obligatorias:
1. No inventes datos. Si una casilla está vacía, tiene un guion "-" o no se lee, usa null.
2. Horas del día en formato "HH:MM" (24 h) tal como están escritas. Duraciones en "H:MM".
3. Números con punto decimal, sin separadores de miles.
4. "campos_dudosos": lista con los nombres exactos de los campos que no puedas leer con seguridad
   (para tramos usa la forma "tramos[0].hora_salida", índice desde 0).
5. Si la imagen está borrosa, cortada, es otro documento o no se pueden leer los datos esenciales responde
   {"exito": false, "motivo_error": "<mensaje breve y amable explicando qué falló y cómo repetir la foto>"}.
"""

PROMPT_COMBUSTIBLE = """Analiza el TIQUETE DE ENTREGA de combustible de aviación (proveedor habitual Terpel).
Devuelve:
{"exito": true, "motivo_error": "",
 "recibo": "número del 'TIQUETE DE ENTREGA No.' (ej. 02-1243808)",
 "numero_boleto": "NUMERO DE BOLETO impreso por el medidor, si aparece",
 "fecha": "AAAA-MM-DD (casillas DIA/MES/AÑO; confírmalo con la fecha impresa del medidor)",
 "hora_inicio": "HH:MM de INICIAR impreso, si aparece",
 "estacion": "código IATA del aeropuerto (casilla IATA, ej. EOH)",
 "ciudad": "casilla CIUDAD (ej. MDE)",
 "matricula": "MATRÍCULA AVIÓN (ej. HK-4803)",
 "tipo_aeronave": "TIPO DE AVIÓN (ej. J32)",
 "numero_vuelo": "VUELO No., si aparece",
 "destino": "DESTINO, si aparece",
 "producto": "JET A1 o AVGAS 100LL según la casilla marcada / producto impreso",
 "galones": número de VOLUMEN ENTREGADO (o BRUTO ENTREGADO GAL),
 "unidad": "GAL" o "LT",
 "api": número o null, "temperatura": número o null,
 "piloto_nombre": "NOMBRE DEL PILOTO O REPRESENTANTE",
 "piloto_identificacion": "NÚMERO DE IDENTIFICACIÓN",
 "operador": "NOMBRE OPERADOR",
 "campos_dudosos": []}
El volumen entregado es obligatorio: si no se lee con claridad, responde exito=false.
""" + _REGLAS

PROMPT_LV_ENCABEZADO = """Primera imagen: LIBRO DE VUELO MO-A3-004 completo (contexto). Segunda imagen: ZONA AMPLIADA del ENCABEZADO.
Extrae SOLO de la zona ampliada:
{"exito": true, "motivo_error": "",
 "numero_libro": "número rojo junto a Nº",
 "fecha": "AAAA-MM-DD (casilla FECHA, escrita a mano como DD/MM/AAAA)",
 "matricula": "A/C REG (ej. HK-4854)",
 "escrito_por": "ESCRITO POR",
 "parametros_motor": [{"motor": 1, "torque": n, "prop_rpm": n, "itt": n, "fuel_flow": n, "fuel_press": n, "oil_press": n, "oil_temp": n}],
 "oat": n, "altitud": n, "velocidad": "texto",
 "horas_totales": n, "ciclos_totales": n, "proxima_inspeccion": "texto",
 "campos_dudosos": []}
""" + _REGLAS

PROMPT_LV_TRAMOS = """Primera imagen: LIBRO DE VUELO MO-A3-004 completo (contexto). Segunda imagen: ZONA AMPLIADA de la TABLA DE TRAMOS.
Columnas de izquierda a derecha: VUELO NÚMERO, ORIGEN, DESTINO, GLS. COMB. (galones tanqueados), RECIBO COMBUSTIBLE,
LBS COMB TOTAL, PAX TOTAL, PAX TTO (tránsito), HORA SALIDA, HORA DESPEGUE, HORA ATERRIZAJE, HORA LLEGADA, TIEMPO VUELO,
TIEMPO NOCTURNO, TIEMPO INSTRUM., TIEMPO TOTAL, LBS COMB LLEGADA, LBS COMB USADO. Al final está la fila TOTALES.
Lee cada fila escrita a mano (ignora filas vacías o tachadas con una línea diagonal) y devuelve:
{"exito": true, "motivo_error": "",
 "tramos": [{"vuelo_numero": "4718", "origen": "EOH", "destino": "EBG", "gls_combustible": 90, "recibo_combustible": "EOH Nº56923",
   "lbs_comb_total": 1600, "pax_total": 8, "pax_transito": null, "hora_salida": "05:55", "hora_despegue": "06:00",
   "hora_aterrizaje": "06:26", "hora_llegada": "06:29", "tiempo_vuelo": "0:26", "tiempo_nocturno": "0:05",
   "tiempo_instrumentos": "0:15", "tiempo_total": "0:34", "lbs_comb_llegada": 1200, "lbs_comb_usado": 400}],
 "totales": {"tiempo_vuelo": "H:MM", "tiempo_nocturno": "H:MM", "tiempo_instrumentos": "H:MM", "tiempo_total": "H:MM",
   "lbs_comb_llegada": n, "lbs_comb_usado": n},
 "campos_dudosos": []}
(El ejemplo solo ilustra el formato; extrae los valores reales de la imagen.)
""" + _REGLAS

PROMPT_LV_TRIPULACION = """Primera imagen: LIBRO DE VUELO MO-A3-004 completo (contexto). Segunda imagen: ZONA AMPLIADA de TRIPULACIÓN y MANTENIMIENTO.
Extrae SOLO de la zona ampliada:
{"exito": true, "motivo_error": "",
 "piloto": {"nombre": "fila PILOTO/INSTRUCTOR", "licencia": "texto completo de la licencia (ej. PTL 1152888465)"},
 "copiloto": {"nombre": "fila COPILOTO/INSTRUCTOR", "licencia": "texto (ej. PCA 7912)"},
 "observaciones_tripulacion": "texto",
 "mantenimiento": [{"item": "ej. 1/M", "reporte": "texto REPORTE", "accion_correctiva": "texto ACCIÓN CORRECTIVA",
   "pn_on": "", "sn_on": "", "pn_off": "", "sn_off": "", "posicion": "LH/RH", "rii": "SI/NO/null",
   "tecnico_nombre": "", "tecnico_licencia": "", "base": ""}],
 "campos_dudosos": []}
Si el bloque de mantenimiento está vacío devuelve "mantenimiento": [].
""" + _REGLAS

PROMPT_FICHA = """Analiza la FICHA DE VUELO / CIERRE DE VUELO / MANIFIESTO DE PESO Y BALANCE de Moon Aerolínea.
Puede contener uno o varios vuelos. Para cada vuelo extrae:
{"exito": true, "motivo_error": "",
 "vuelos": [{"fecha": "AAAA-MM-DD", "matricula": "HK-....", "numero_vuelo": "", "origen": "IATA", "destino": "IATA",
   "hora_salida": "HH:MM", "pax_total": n, "tripadis": n (personal aeronáutico/tripulación adicional que NO paga tasa),
   "bebes": n (infantes / 'nibra'), "pax_transito": n ('exentos por tránsito'), "carga_a_bordo_kg": n,
   "carga_transito_kg": n, "combustible_basico_lbs": n (COMBUSTIBLE BÁSICO del manifiesto P&B), "piloto": "texto"}],
 "observaciones": "texto",
 "campos_dudosos": []}
Para campos_dudosos de vuelos usa la forma "vuelos[0].pax_total".
""" + _REGLAS


def _dudosos(d) -> list[str]:
    return [str(x) for x in (d.get("campos_dudosos") or []) if R.texto(x)]


def leer_tiquete(jpeg: bytes) -> dict:
    r, err = llamar_claude_json([jpeg], PROMPT_COMBUSTIBLE, 2048)
    if err:
        return {"exito": False, "motivo_error": err, "ia_fallo": True}
    if not r.get("exito"):
        return {"exito": False, "motivo_error": R.texto(r.get("motivo_error"), "La imagen no es clara. Repite la foto.")}
    dud = _dudosos(r)
    gal = R.a_float(r.get("galones"))
    if gal is not None and R.texto(r.get("unidad")).upper().startswith("L"):
        gal = round(gal / R.LITROS_POR_GALON, 1)
        dud.append("galones (convertido de litros)")
    ident = R.solo_digitos(r.get("piloto_identificacion"))
    datos = {"recibo": R.texto(r.get("recibo")) or None, "numero_boleto": R.texto(r.get("numero_boleto")) or None,
             "fecha_consumo": R.parse_fecha(r.get("fecha")), "hora_min": R.parse_hora(r.get("hora_inicio")),
             "estacion": R.norm_iata(r.get("estacion")) or None, "matricula": R.norm_matricula(r.get("matricula")) or None,
             "producto": R.norm_producto(r.get("producto")), "destino": R.norm_iata(r.get("destino")) or None,
             "numero_vuelo": R.solo_digitos(r.get("numero_vuelo")) or None, "galones": gal,
             "piloto_nombre": R.texto(r.get("piloto_nombre")) or None, "piloto_cedula": ident or None,
             "observacion": None}
    return {"exito": True, "motivo_error": "", "datos": datos, "dudosos": dud, "crudo": r}


_MAPA_TRAMO = {"vuelo_numero": "numero_vuelo", "origen": "origen", "destino": "destino",
               "gls_combustible": "galones_tanqueados", "recibo_combustible": "recibo_combustible",
               "lbs_comb_total": "lbs_comb_salida", "pax_total": "pax_total", "pax_transito": "pax_transito",
               "hora_salida": "hora_salida_min", "hora_despegue": "hora_despegue_min",
               "hora_aterrizaje": "hora_aterrizaje_min", "hora_llegada": "hora_llegada_min",
               "tiempo_vuelo": "tiempo_vuelo_min", "tiempo_nocturno": "tiempo_nocturno_min",
               "tiempo_instrumentos": "tiempo_instrumentos_min", "tiempo_total": "tiempo_total_min",
               "lbs_comb_llegada": "lbs_comb_llegada", "lbs_comb_usado": "lbs_comb_usado"}


def _norm_tramo_ia(t: dict) -> dict:
    out = {}
    for k_ia, k in _MAPA_TRAMO.items():
        v = t.get(k_ia)
        if k.startswith("hora_"):
            out[k] = R.parse_hora(v)
        elif k.startswith("tiempo_"):
            out[k] = R.parse_duracion(v)
        elif k in ("numero_vuelo",):
            out[k] = R.solo_digitos(v) or None
        elif k in ("origen", "destino"):
            out[k] = R.norm_iata(v) or None
        elif k == "recibo_combustible":
            out[k] = R.texto(v) or None
        elif k in ("pax_total", "pax_transito"):
            out[k] = R.a_int(v)
        else:
            out[k] = R.a_float(v)
    return out


def _traducir_dudoso(d: str) -> str:
    for k_ia, k in _MAPA_TRAMO.items():
        d = re.sub(rf"\b{k_ia}\b", k, d)
    return d


def leer_libro(img: Image.Image) -> dict:
    """Lectura del Libro de Vuelo por 3 zonas en paralelo y fusión del resultado."""
    contexto = jpeg_bytes(img, 1500, 85)
    zonas = {n: jpeg_bytes(z, 2400, 90) for n, z in zonas_libro(img).items()}
    tareas = {"encabezado": PROMPT_LV_ENCABEZADO, "tramos": PROMPT_LV_TRAMOS, "tripulacion_oma": PROMPT_LV_TRIPULACION}
    with ThreadPoolExecutor(max_workers=3) as ex:
        futuros = {n: ex.submit(llamar_claude_json, [contexto, zonas[n]], p, 4096) for n, p in tareas.items()}
        res = {n: f.result() for n, f in futuros.items()}
    errores = [e for _, e in res.values() if e]
    if len(errores) == 3:
        return {"exito": False, "motivo_error": errores[0], "ia_fallo": True}
    enc, _ = res["encabezado"]
    tra, _ = res["tramos"]
    tri, _ = res["tripulacion_oma"]
    enc, tra, tri = enc or {}, tra or {}, tri or {}
    if tra and not tra.get("exito", True) and not tra.get("tramos"):
        return {"exito": False, "motivo_error": R.texto(tra.get("motivo_error"),
                                                        "No se lee la tabla de tramos. Repite la foto de frente y con buena luz.")}
    dud = [f"encabezado.{x}" for x in _dudosos(enc)] + [_traducir_dudoso(x) for x in _dudosos(tra)] + \
          [f"tripulacion.{x}" for x in _dudosos(tri)]
    for nombre, (_, e) in res.items():
        if e:
            dud.append(f"zona {nombre}: no se pudo leer ({e})")
    tramos = [_norm_tramo_ia(t) for t in (tra.get("tramos") or []) if isinstance(t, dict)]
    tramos = [t for t in tramos if any(v is not None for v in t.values())]
    tot = tra.get("totales") or {}
    totales = {"tot_tiempo_vuelo_min": R.parse_duracion(tot.get("tiempo_vuelo")),
               "tot_tiempo_nocturno_min": R.parse_duracion(tot.get("tiempo_nocturno")),
               "tot_tiempo_instrumentos_min": R.parse_duracion(tot.get("tiempo_instrumentos")),
               "tot_tiempo_total_min": R.parse_duracion(tot.get("tiempo_total")),
               "tot_lbs_usado": R.a_float(tot.get("lbs_comb_usado"))}
    piloto, copiloto = tri.get("piloto") or {}, tri.get("copiloto") or {}
    cabecera = {"numero_libro": R.solo_digitos(enc.get("numero_libro")) or None, "fecha": R.parse_fecha(enc.get("fecha")),
                "matricula": R.norm_matricula(enc.get("matricula")) or None,
                "piloto_nombre": R.texto(piloto.get("nombre")) or None, "piloto_licencia": R.texto(piloto.get("licencia")) or None,
                "copiloto_nombre": R.texto(copiloto.get("nombre")) or None,
                "copiloto_licencia": R.texto(copiloto.get("licencia")) or None,
                "horas_totales": R.a_float(enc.get("horas_totales")), "ciclos_totales": R.a_int(enc.get("ciclos_totales")),
                "proxima_inspeccion": R.texto(enc.get("proxima_inspeccion")) or None,
                "observaciones": R.texto(tri.get("observaciones_tripulacion")) or None}
    mant = []
    for m in tri.get("mantenimiento") or []:
        if not isinstance(m, dict):
            continue
        item = {k: (R.texto(m.get(k)) or None) for k in ("item", "reporte", "accion_correctiva", "pn_on", "sn_on", "pn_off",
                                                         "sn_off", "posicion", "rii", "tecnico_nombre", "tecnico_licencia", "base")}
        if any(item.values()):
            mant.append(item)
    return {"exito": True, "motivo_error": "", "cabecera": cabecera, "totales": totales, "tramos": tramos,
            "mantenimiento": mant, "parametros_motor": enc.get("parametros_motor") or [], "dudosos": dud,
            "crudo": {"encabezado": enc, "tramos": tra, "tripulacion_oma": tri}}


def leer_ficha(jpeg: bytes) -> dict:
    r, err = llamar_claude_json([jpeg], PROMPT_FICHA, 3000)
    if err:
        return {"exito": False, "motivo_error": err, "ia_fallo": True}
    if not r.get("exito"):
        return {"exito": False, "motivo_error": R.texto(r.get("motivo_error"), "La ficha no es legible. Repite la foto.")}
    vuelos = []
    for v in r.get("vuelos") or []:
        if not isinstance(v, dict):
            continue
        vuelos.append({"fecha": R.parse_fecha(v.get("fecha")), "matricula": R.norm_matricula(v.get("matricula")) or None,
                       "numero_vuelo": R.solo_digitos(v.get("numero_vuelo")) or None,
                       "origen": R.norm_iata(v.get("origen")) or None, "destino": R.norm_iata(v.get("destino")) or None,
                       "hora_min": R.parse_hora(v.get("hora_salida")), "pax_total": R.a_int(v.get("pax_total")),
                       "tripadis": R.a_int(v.get("tripadis")), "bebes": R.a_int(v.get("bebes")),
                       "pax_transito": R.a_int(v.get("pax_transito")), "carga_a_bordo_kg": R.a_float(v.get("carga_a_bordo_kg")),
                       "carga_transito_kg": R.a_float(v.get("carga_transito_kg")),
                       "combustible_basico_lbs": R.a_float(v.get("combustible_basico_lbs")),
                       "piloto_nombre": R.texto(v.get("piloto")) or None, "observaciones": None})
    if not vuelos:
        return {"exito": False, "motivo_error": "No encontré vuelos en la ficha. Verifica que sea el documento correcto."}
    return {"exito": True, "motivo_error": "", "vuelos": vuelos, "observaciones": R.texto(r.get("observaciones")),
            "dudosos": _dudosos(r), "crudo": r}


def releer_documento(tipo: str, ruta_rel: str) -> dict:
    """Re-lectura independiente para auditoría (revisoría fiscal sobre la imagen original)."""
    raw = leer_evidencia(ruta_rel)
    if raw is None:
        return {"exito": False, "motivo_error": "Evidencia no encontrada."}
    try:
        img = abrir_imagen(raw)
    except ImagenInvalida as e:
        return {"exito": False, "motivo_error": str(e)}
    if tipo == "COMB":
        return leer_tiquete(jpeg_bytes(img))
    if tipo == "LV":
        return leer_libro(img)
    return leer_ficha(jpeg_bytes(img))

# -*- coding: utf-8 -*-
"""
Genera el portal comercial de liquidacion (index.html autonomo).
==================================================================
V4: agrega filtros de Marca del vehiculo / Modelo / Anio, apertura de
rangos de anios (solo para filtrar/buscar, nunca duplica stock ni
valorizacion), y usa el stock valorizado OFICIAL de "Base No Estrategicos
y obsoletos.xlsx" (hoja "Resumen (2)") en vez de calcularlo.

Reutiliza el procesador de datos ya probado (dashboard-obsolescencia/src).
La columna Zona del Excel viene corregida (una zona unica por tienda, sin
mezclas), asi que se usa directamente.

Ejecutar desde esta carpeta:
    python generar_portal.py

Salida:
    index.html   (listo para publicar en GitHub Pages o abrir localmente)
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

RAIZ_PORTAL = Path(__file__).resolve().parent
RAIZ_PROYECTO = RAIZ_PORTAL.parent
RAIZ_STREAMLIT = RAIZ_PROYECTO / "dashboard-obsolescencia"
sys.path.insert(0, str(RAIZ_STREAMLIT))

from src.services import cargador_excel as datos  # noqa: E402

import anios  # noqa: E402
import resumen_zona  # noqa: E402

PLANTILLA = RAIZ_PORTAL / "plantilla_portal.html"
SALIDA = RAIZ_PORTAL / "index.html"
LOGO_GRUPO_PLANET_B64 = (RAIZ_PORTAL / "logo_grupo_planet_b64.txt").read_text().strip()
LOGO_AUTOPLANET_B64 = (RAIZ_PORTAL / "logo_autoplanet_b64.txt").read_text().strip()
BANNER_HERO_B64 = (RAIZ_PORTAL / "banner_hero_b64.txt").read_text().strip()

# Archivo oficial de valorizacion por zona (ver seccion 3 del README).
ARCHIVO_RESUMEN_ZONA = RAIZ_PROYECTO / "Base No Estrategicos y obsoletos.xlsx"
HOJA_RESUMEN_ZONA = "Resumen (2)"

# Orden en que se ofrecen las zonas en el selector: geograficas de norte a
# sur primero, luego las no geograficas. Cualquier zona nueva que aparezca
# en el Excel y no este en esta lista se agrega al final, alfabetica -- no
# desaparece silenciosamente.
ORDEN_ZONAS = [
    "Zona Norte", "Zona RM 1", "Zona RM 2", "Zona RM 3", "Zona V Region",
    "Zona Centro Sur", "Zona Sur", "Agroplanet", "CD", "E-Commerce", "AP/SG",
]


class Diccionario:
    def __init__(self):
        self.mapa: dict[str, int] = {}
        self.lista: list[str] = []

    def idx(self, valor) -> int:
        texto = "" if valor is None else str(valor)
        if texto not in self.mapa:
            self.mapa[texto] = len(self.lista)
            self.lista.append(texto)
        return self.mapa[texto]


def idx_opcional(dic: Diccionario, valor) -> int:
    """-1 cuando el dato no existe, en vez de guardar una cadena vacia."""
    if valor is None or str(valor).strip() == "":
        return -1
    return dic.idx(valor)


def idx_opcional_marca_vehiculo(dic: Diccionario, valor) -> int:
    """
    Igual que idx_opcional, pero ademas trata "0" como vacio.

    En el Excel V4, 12 filas traen literalmente "0" en Marca del vehiculo
    (probablemente un valor de relleno de una formula de busqueda que no
    encontro coincidencia, no una marca real). Se documenta como hallazgo,
    no se inventa una marca para esas filas.
    """
    if valor is None or str(valor).strip() in ("", "0"):
        return -1
    return dic.idx(valor)


def limpio(valor):
    if valor is None:
        return None
    try:
        if valor != valor:  # NaN
            return None
    except TypeError:
        return None
    return int(round(float(valor)))


def main() -> int:
    advertencias_generales: list[str] = []

    print("Leyendo el Excel con el procesador ya probado...")
    resultado = datos.cargar()
    df = resultado.df
    print(f"  {len(df):,} productos con stock en la base")
    if resultado.columnas_faltantes:
        print(f"  columnas opcionales no encontradas: {resultado.columnas_faltantes}")

    # --- Verificacion: una tienda no deberia tener mas de una zona --------
    por_tienda = df.groupby("tienda_label")["zona"].nunique()
    conflictivas = por_tienda[por_tienda > 1]
    if len(conflictivas):
        msg = (f"{len(conflictivas)} tiendas aparecen en mas de una zona en el "
               "Excel (se usa la primera zona encontrada de cada una).")
        advertencias_generales.append(msg)
        print(f"\n  AVISO: {msg}")
        for t in conflictivas.index:
            print("   -", t, sorted(df.loc[df['tienda_label'] == t, 'zona'].unique()))

    # --- Validaciones de datos (seccion 22 del pedido) ---------------------
    # limpiar_texto() deja "" para vacios (no NaN), asi que se cuenta asi.
    def _vacios(col: str, tambien_cero: bool = False) -> int:
        if col not in df:
            return len(df)
        vacio = df[col].astype(str).str.strip() == ""
        if tambien_cero:
            vacio = vacio | (df[col].astype(str).str.strip() == "0")
        return int(vacio.sum())

    n_marca_veh = _vacios("marca_vehiculo", tambien_cero=True)
    if n_marca_veh:
        advertencias_generales.append(
            f"{n_marca_veh} productos sin marca del vehiculo (incluye 12 con valor "
            "literal \"0\" en el Excel, tratado como vacio; quedan igual, solo sin ese dato)."
            if (df["marca_vehiculo"].astype(str).str.strip() == "0").sum()
            else f"{n_marca_veh} productos sin marca del vehiculo (quedan igual, solo sin ese dato)."
        )
    for etiqueta, col in [("modelo", "app_modelo"), ("motor", "app_motor")]:
        n = _vacios(col)
        if n:
            advertencias_generales.append(f"{n} productos sin {etiqueta} (quedan igual, solo sin ese dato).")

    # --- Diccionarios de texto --------------------------------------------
    d_zona, d_tienda, d_mat, d_texto, d_marca = (Diccionario() for _ in range(5))
    d_marca_veh, d_cat, d_subcat, d_modelo, d_motor, d_anio = (Diccionario() for _ in range(6))

    filas = []
    for fila in df.itertuples(index=False):
        filas.append([
            d_zona.idx(fila.zona),
            d_tienda.idx(fila.tienda_label),
            d_mat.idx(fila.material),
            d_texto.idx(getattr(fila, "texto_breve", "")),
            d_marca.idx(getattr(fila, "marca", "")),
            idx_opcional_marca_vehiculo(d_marca_veh, getattr(fila, "marca_vehiculo", None)),
            d_cat.idx(getattr(fila, "categoria", "")),
            d_subcat.idx(fila.subcategoria),
            idx_opcional(d_modelo, getattr(fila, "app_modelo", None)),
            idx_opcional(d_motor, getattr(fila, "app_motor", None)),
            idx_opcional(d_anio, getattr(fila, "app_anios", None)),
            limpio(fila.stock),
            limpio(getattr(fila, "precio_normal", None)),
            limpio(fila.valor_remate),
        ])

    # --- Apertura de anios (SOLO para filtrar/buscar; ver anios.py) --------
    # Un arreglo paralelo a d_anio.lista: para cada string original de anio,
    # la lista de anios individuales que representa. No se crea ninguna fila
    # nueva por anio -- el producto sigue siendo una unica fila; el filtro de
    # anio consulta este arreglo por indice (f[AN]) y nunca duplica stock.
    anios_por_indice: list[list[int]] = []
    anios_invalidos = 0
    for texto_original in d_anio.lista:
        r = anios.parsear_anio(texto_original)
        anios_por_indice.append(sorted(r.anios))
        if not r.valido:
            anios_invalidos += 1
    if anios_invalidos:
        advertencias_generales.append(
            f"{anios_invalidos} valores de año no se pudieron interpretar del todo "
            "(quedan sin filtro de año pero se conservan para búsqueda)."
        )
    todos_los_anios = sorted({a for lista in anios_por_indice for a in lista})
    print(f"\n  Años: {len(todos_los_anios)} años distintos encontrados "
          f"({todos_los_anios[0]}-{todos_los_anios[-1]} si hay alguno)" if todos_los_anios else "\n  Años: ninguno")

    # --- Zonas en el orden fijo de ORDEN_ZONAS -----------------------------
    zonas_presentes = [z for z in ORDEN_ZONAS if z in d_zona.mapa]
    zonas_nuevas = sorted(z for z in d_zona.mapa if z not in ORDEN_ZONAS)
    if zonas_nuevas:
        print(f"\n  AVISO: zonas nuevas no listadas en ORDEN_ZONAS (se agregan "
              f"al final del selector): {zonas_nuevas}")
    zonas_presentes += zonas_nuevas
    orden_zona = {z: i for i, z in enumerate(zonas_presentes)}
    reindex = sorted(range(len(d_zona.lista)), key=lambda i: orden_zona.get(d_zona.lista[i], 999))
    permutacion = {viejo: nuevo for nuevo, viejo in enumerate(reindex)}
    d_zona.lista = [d_zona.lista[i] for i in reindex]
    for f in filas:
        f[0] = permutacion[f[0]]

    # --- Stock valorizado OFICIAL por zona (seccion 14 del pedido) ---------
    print(f"\nLeyendo stock valorizado oficial de "
          f"'{ARCHIVO_RESUMEN_ZONA.name}' -> hoja '{HOJA_RESUMEN_ZONA}'...")
    res_zona = resumen_zona.leer(ARCHIVO_RESUMEN_ZONA, HOJA_RESUMEN_ZONA)
    analisis = [
        {"zona": z, "valor": round(v)}
        for z, v, _pct in res_zona.bloque_usado.filas
        if v is not None
    ]
    analisis.sort(key=lambda x: -x["valor"])
    total_oficial = res_zona.bloque_usado.total_general
    print(f"  {len(analisis)} zonas leidas. Total oficial: ${total_oficial:,.0f}")
    for a in analisis:
        print(f"    {a['zona']:22s} ${a['valor']:>15,.0f}  {a['valor']/total_oficial*100:5.1f}%")
    for adv in res_zona.advertencias:
        print("  AVISO:", adv)
        advertencias_generales.append(adv)

    # --- Reporte final de advertencias -------------------------------------
    if advertencias_generales:
        print(f"\n{'='*70}\nADVERTENCIAS ({len(advertencias_generales)}):")
        for a in advertencias_generales:
            print("  -", a)
    else:
        print("\nSin advertencias.")

    marca_tiempo = resultado.archivo.stat().st_mtime
    payload = {
        "zonas": d_zona.lista,
        "tiendas": d_tienda.lista,
        "materiales": d_mat.lista,
        "textos": d_texto.lista,
        "marcas": d_marca.lista,
        "marcasVehiculo": d_marca_veh.lista,
        "categorias": d_cat.lista,
        "subcats": d_subcat.lista,
        "modelos": d_modelo.lista,
        "motores": d_motor.lista,
        "anios": d_anio.lista,
        "aniosPorIndice": anios_por_indice,
        "filas": filas,
        "analisis": analisis,
        "analisisTotalOficial": round(total_oficial) if total_oficial else None,
        "info": {
            "archivo": resultado.archivo.name,
            "archivoResumen": ARCHIVO_RESUMEN_ZONA.name,
            "fecha": datetime.fromtimestamp(marca_tiempo).strftime("%d-%m-%Y"),
            "marcaTiempo": int(marca_tiempo * 1000),
            "totalTiendas": int(df["tienda_label"].nunique()),
            "totalZonas": len(zonas_presentes),
            "totalSubcats": int(df["subcategoria"].nunique()),
            "totalAnios": len(todos_los_anios),
        },
    }

    html = PLANTILLA.read_text(encoding="utf-8")
    html = re.sub(
        r"\.normalize\('NFD'\)\.replace\(/\[[^\]]*\]/g, ''\)",
        r".normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')",
        html,
    )
    html = html.replace("__LOGO_GRUPO_PLANET__", LOGO_GRUPO_PLANET_B64)
    html = html.replace("__LOGO_AUTOPLANET__", LOGO_AUTOPLANET_B64)
    html = html.replace("__BANNER_HERO__", BANNER_HERO_B64)
    datos_js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    html = html.replace("/*__DATOS__*/null", datos_js)

    SALIDA.write_text(html, encoding="utf-8")
    mb = SALIDA.stat().st_size / 1048576
    print(f"\nListo -> {SALIDA.name}  ({mb:.2f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

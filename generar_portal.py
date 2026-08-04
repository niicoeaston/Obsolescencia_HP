# -*- coding: utf-8 -*-
"""
Genera el portal comercial de liquidacion (index.html autonomo).
==================================================================
Reutiliza el procesador de datos ya probado (dashboard-obsolescencia/src),
aplica el mapeo geografico real de zona_tienda.py (la columna Zona del
Excel no es confiable por tienda, ver ese archivo), agrega el stock
valorizado por zona para la seccion Analisis, e inyecta todo en la
plantilla HTML/CSS/JS.

Ejecutar desde esta carpeta:
    python generar_portal.py

Salida:
    index.html   (listo para publicar en GitHub Pages o abrir localmente)
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

RAIZ_PORTAL = Path(__file__).resolve().parent
RAIZ_STREAMLIT = RAIZ_PORTAL.parent / "dashboard-obsolescencia"
sys.path.insert(0, str(RAIZ_STREAMLIT))

from src.services import cargador_excel as datos  # noqa: E402

import zona_tienda  # noqa: E402

PLANTILLA = RAIZ_PORTAL / "plantilla_portal.html"
SALIDA = RAIZ_PORTAL / "index.html"
LOGO_B64 = (RAIZ_PORTAL / "logo_grupo_planet_b64.txt").read_text().strip()


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
    print("Leyendo el Excel con el procesador ya probado...")
    resultado = datos.cargar()
    df = resultado.df
    print(f"  {len(df):,} productos con stock en la base")

    # --- Zona geografica real (no la columna Zona del Excel) --------------
    sin_mapa = sorted(set(df["tienda_label"]) - set(zona_tienda.ZONA_PRINCIPAL))
    if sin_mapa:
        print("\n  AVISO: estas tiendas no estan en zona_tienda.py y quedaran "
              "sin zona geografica asignada (no apareceran en el selector de "
              "zona, pero si en la tienda si se filtra por otro medio):")
        for t in sin_mapa:
            print("   -", t)

    zona_real = df["tienda_label"].map(zona_tienda.ZONA_PRINCIPAL)
    df = df.assign(zona_real=zona_real)
    df = df[df["zona_real"].notna()]  # tiendas nuevas sin mapear quedan fuera, no se inventan

    # --- Diccionarios de texto --------------------------------------------
    d_zona, d_tienda, d_mat, d_texto, d_marca = (Diccionario() for _ in range(5))
    d_cat, d_subcat, d_modelo, d_motor, d_anio = (Diccionario() for _ in range(5))

    filas = []
    for fila in df.itertuples(index=False):
        filas.append([
            d_zona.idx(fila.zona_real),
            d_tienda.idx(fila.tienda_label),
            d_mat.idx(fila.material),
            d_texto.idx(getattr(fila, "texto_breve", "")),
            d_marca.idx(getattr(fila, "marca", "")),
            d_cat.idx(getattr(fila, "categoria", "")),
            d_subcat.idx(fila.subcategoria),
            idx_opcional(d_modelo, getattr(fila, "app_modelo", None)),
            idx_opcional(d_motor, getattr(fila, "app_motor", None)),
            idx_opcional(d_anio, getattr(fila, "app_anios", None)),
            limpio(getattr(fila, "precio_normal", None)),
            limpio(fila.valor_remate),
        ])

    # --- Zonas en el orden fijo definido en zona_tienda.py -----------------
    zonas_presentes = [z for z in zona_tienda.ORDEN_ZONAS if z in d_zona.mapa]
    orden_zona = {z: i for i, z in enumerate(zonas_presentes)}
    # Reordena el diccionario de zonas para que el frontend las liste en ese orden
    reindex = sorted(range(len(d_zona.lista)), key=lambda i: orden_zona.get(d_zona.lista[i], 999))
    permutacion = {viejo: nuevo for nuevo, viejo in enumerate(reindex)}
    d_zona.lista = [d_zona.lista[i] for i in reindex]
    for f in filas:
        f[0] = permutacion[f[0]]

    # --- Stock valorizado por zona (para la seccion Analisis) --------------
    valor_por_zona: dict[str, float] = defaultdict(float)
    for stock, remate, zona in zip(df["stock"], df["valor_remate"], df["zona_real"]):
        if stock is not None and remate is not None:
            valor_por_zona[zona] += stock * remate
    analisis = sorted(
        ({"zona": z, "valor": round(v)} for z, v in valor_por_zona.items()),
        key=lambda x: -x["valor"],
    )
    print("\n  Stock valorizado por zona:")
    total = sum(a["valor"] for a in analisis)
    for a in analisis:
        print(f"    {a['zona']:22s} ${a['valor']:>15,.0f}  {a['valor']/total*100:5.1f}%")

    marca_tiempo = resultado.archivo.stat().st_mtime
    payload = {
        "zonas": d_zona.lista,
        "tiendas": d_tienda.lista,
        "materiales": d_mat.lista,
        "textos": d_texto.lista,
        "marcas": d_marca.lista,
        "categorias": d_cat.lista,
        "subcats": d_subcat.lista,
        "modelos": d_modelo.lista,
        "motores": d_motor.lista,
        "anios": d_anio.lista,
        "filas": filas,
        "analisis": analisis,
        "info": {
            "archivo": resultado.archivo.name,
            "fecha": datetime.fromtimestamp(marca_tiempo).strftime("%d-%m-%Y"),
            "marcaTiempo": int(marca_tiempo * 1000),
            "totalTiendas": int(df["tienda_label"].nunique()),
            "totalZonas": len(zonas_presentes),
            "totalSubcats": int(df["subcategoria"].nunique()),
        },
    }

    html = PLANTILLA.read_text(encoding="utf-8")
    html = re.sub(
        r"\.normalize\('NFD'\)\.replace\(/\[[^\]]*\]/g, ''\)",
        r".normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')",
        html,
    )
    html = html.replace("__LOGO_GRUPO_PLANET__", LOGO_B64)
    datos_js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    html = html.replace("/*__DATOS__*/null", datos_js)

    SALIDA.write_text(html, encoding="utf-8")
    mb = SALIDA.stat().st_size / 1048576
    print(f"\nListo -> {SALIDA.name}  ({mb:.2f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

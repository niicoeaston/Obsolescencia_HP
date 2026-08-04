# -*- coding: utf-8 -*-
"""
Calcula la zona principal (moda) de cada tienda a partir de la base real y
la compara contra el mapeo manual guardado en zona_tienda.py.

Ejecutar:
    python calcular_zona_tienda.py

No escribe nada solo: imprime el diccionario calculado (para copiar a mano
a zona_tienda.py) y la lista de tiendas donde el mapeo manual difiere del
calculado, para revisar antes de aceptarlo.
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

RAIZ_PORTAL = Path(__file__).resolve().parent
RAIZ_STREAMLIT = RAIZ_PORTAL.parent / "dashboard-obsolescencia"
sys.path.insert(0, str(RAIZ_STREAMLIT))

from src.services import cargador_excel as datos  # noqa: E402

import zona_tienda  # noqa: E402  (el mapeo manual actual, si existe)

resultado = datos.cargar()
df = resultado.df

conteo: dict[str, Counter] = defaultdict(Counter)
for tienda, zona in zip(df["tienda_label"], df["zona"]):
    conteo[tienda][zona] += 1

print(f"{len(conteo)} tiendas encontradas en la base.\n")
print("dict calculado (moda = zona con mas filas por tienda):\n")

calculado = {}
diferencias = []
for tienda in sorted(conteo):
    c = conteo[tienda]
    zona_moda, n_moda = c.most_common(1)[0]
    calculado[tienda] = zona_moda
    total = sum(c.values())
    manual = zona_tienda.ZONA_PRINCIPAL.get(tienda)
    marca = ""
    if manual and manual != zona_moda:
        marca = f"   << MANUAL DICE '{manual}', DIFIERE"
        diferencias.append((tienda, manual, zona_moda, dict(c)))
    elif not manual:
        marca = "   << NO ESTA EN EL MANUAL"
    print(f'    "{tienda}": "{zona_moda}",'
          f'  # {n_moda}/{total} filas{marca}')

print(f"\n\nTotal tiendas: {len(calculado)}")
print(f"Tiendas ausentes del mapeo manual: "
      f"{sum(1 for t in calculado if t not in zona_tienda.ZONA_PRINCIPAL)}")
print(f"Tiendas donde el manual difiere de la moda calculada: {len(diferencias)}")
for tienda, manual, moda, detalle in diferencias:
    print(f"\n  {tienda}")
    print(f"    manual   = {manual}")
    print(f"    calculado= {moda}")
    print(f"    reparto  = {detalle}")

# Tiendas con reparto muy parejo (la moda gana por muy poco): vale la pena
# que un humano las revise aunque el manual coincida con la moda.
print("\n\nTiendas con reparto ajustado (moda gana con <60% de las filas):")
for tienda in sorted(conteo):
    c = conteo[tienda]
    zona_moda, n_moda = c.most_common(1)[0]
    total = sum(c.values())
    if len(c) > 1 and n_moda / total < 0.6:
        print(f"    {tienda}: {dict(c)}")

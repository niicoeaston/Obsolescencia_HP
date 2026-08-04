# -*- coding: utf-8 -*-
"""
Apertura de rangos de anios para el filtro de Anio.
=====================================================
IMPORTANTE: esto SOLO produce un conjunto de anios individuales para poder
filtrar/buscar por anio. Nunca crea filas nuevas, nunca multiplica stock ni
valorizacion -- cada registro original sigue siendo una sola fila; el
conjunto de anios queda "pegado" a esa misma fila como un dato mas (igual
que el material o la marca), no como copias del producto.

Formatos soportados (ver README seccion "Formatos de anio soportados"):
  - Anio unico:              "2015"            -> {2015}
  - Rango simple:            "2004-2016"       -> {2004..2016}
  - Rango con espacios:      "2004 - 2016"     -> {2004..2016}
  - Guiones largos:          "2004–2016" / "2004—2016" -> idem (se normalizan a "-")
  - Varios segmentos:        "2004-2008 / 2010-2014" o "2004-2008, 2010, 2012-2014"
  - Anios sueltos con coma:  "2004, 2006, 2008" -> {2004, 2006, 2008}

Vacios o invalidos: nunca se inventa un anio. Se conserva el string
original para mostrarlo y para busqueda de texto libre, pero el producto
queda fuera de cualquier filtro que pida un anio especifico.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime

# Guiones alternativos que deben tratarse igual que el guion normal.
_GUIONES = ["–", "—", "−", "‒", "―"]  # – — − ‒ ―

_ANIO_MIN = 1900
_ANIO_MAX = datetime.now().year + 3
_SPAN_MAXIMO = 60  # anios; un rango mas largo que esto se marca como sospechoso

_RE_SEGMENTO_RANGO = re.compile(r"^\s*(\d{4})\s*-\s*(\d{4})\s*$")
_RE_SEGMENTO_UNICO = re.compile(r"^\s*(\d{4})\s*$")


@dataclass
class ResultadoAnio:
    original: str                     # tal cual vino en el Excel, para mostrar/auditar
    anios: set[int] = field(default_factory=set)  # anios individuales validos
    valido: bool = True               # False si hubo algun segmento no interpretable
    advertencia: str | None = None    # motivo, si valido=False o hubo problemas parciales


def normalizar_guiones(texto: str) -> str:
    for g in _GUIONES:
        texto = texto.replace(g, "-")
    return texto


def _validar_rango(y1: int, y2: int) -> str | None:
    """Devuelve un mensaje de advertencia si el rango no es aceptable, o None si esta bien."""
    if y1 > y2:
        return f"rango invertido ({y1}-{y2}, el inicial es mayor que el final)"
    if y1 < _ANIO_MIN or y2 > _ANIO_MAX:
        return f"anio fuera de rango razonable ({y1}-{y2})"
    if (y2 - y1) > _SPAN_MAXIMO:
        return f"rango sospechosamente largo ({y1}-{y2}, {y2-y1} anios)"
    return None


def parsear_anio(valor) -> ResultadoAnio:
    """
    Interpreta el contenido de la columna Anio de una fila.

    Nunca lanza excepcion: ante cualquier formato no reconocido, devuelve
    un resultado con anios vacio, valido=False y una advertencia legible,
    conservando siempre el valor original.
    """
    original = "" if valor is None else str(valor).strip()
    if original == "" or original.lower() in {"nan", "none", "s/i", "n/a", "-"}:
        return ResultadoAnio(original=original, anios=set(), valido=True, advertencia=None)

    texto = normalizar_guiones(original)
    # separadores de multiples rangos/anios: coma o slash
    segmentos = [s.strip() for s in re.split(r"[,/]", texto) if s.strip()]
    if not segmentos:
        return ResultadoAnio(original=original, anios=set(), valido=False,
                              advertencia="celda no interpretable")

    anios: set[int] = set()
    problemas: list[str] = []

    for seg in segmentos:
        m_rango = _RE_SEGMENTO_RANGO.match(seg)
        m_unico = _RE_SEGMENTO_UNICO.match(seg)
        if m_rango:
            y1, y2 = int(m_rango.group(1)), int(m_rango.group(2))
            problema = _validar_rango(y1, y2)
            if problema:
                problemas.append(f"'{seg}': {problema}")
                continue
            anios.update(range(y1, y2 + 1))
        elif m_unico:
            y = int(m_unico.group(1))
            if y < _ANIO_MIN or y > _ANIO_MAX:
                problemas.append(f"'{seg}': anio fuera de rango razonable ({y})")
                continue
            anios.add(y)
        else:
            problemas.append(f"'{seg}': formato no reconocido")

    if not anios:
        return ResultadoAnio(
            original=original, anios=set(), valido=False,
            advertencia="; ".join(problemas) or "sin anios validos",
        )
    return ResultadoAnio(
        original=original, anios=anios, valido=not problemas,
        advertencia="; ".join(problemas) if problemas else None,
    )


def probar() -> None:
    """Pruebas rapidas de los formatos del enunciado. Ejecutar: python anios.py"""
    casos = [
        ("2015", {2015}),
        ("2004-2016", set(range(2004, 2017))),
        ("2004 - 2016", set(range(2004, 2017))),
        ("2004–2016", set(range(2004, 2017))),
        ("2004—2016", set(range(2004, 2017))),
        ("2004-2008 / 2010-2014", set(range(2004,2009)) | set(range(2010,2015))),
        ("2004, 2006, 2008", {2004, 2006, 2008}),
        ("2004-2008, 2010, 2012-2014", set(range(2004,2009)) | {2010} | set(range(2012,2015))),
        ("", set()),
        ("2016-2004", set()),        # invertido -> invalido
        ("2004-20160", set()),       # numero de 5 digitos -> no matchea, invalido
    ]
    ok = 0
    for texto, esperado in casos:
        r = parsear_anio(texto)
        estado = "OK" if r.anios == esperado else "FALLA"
        if estado == "OK":
            ok += 1
        print(f"[{estado}] {texto!r:35s} -> {sorted(r.anios)[:3]}{'...' if len(r.anios)>3 else ''} "
              f"(n={len(r.anios)}) valido={r.valido} advertencia={r.advertencia}")
    print(f"\n{ok}/{len(casos)} casos correctos")


if __name__ == "__main__":
    probar()

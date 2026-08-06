# -*- coding: utf-8 -*-
"""
Recorta iconos individuales desde las hojas de referencia que dejo el usuario
en Elementos/ (capturas de un set de iconos de linea tecnica). No se genera
nada con IA (sin credito disponible) -- se usan y limpian los graficos reales
provistos: se recorta cada icono de su celda de grilla, se autoajusta al
contenido real (sin la etiqueta de texto en ingles de abajo, que es solo
referencia) y se vuelve transparente el fondo blanco.
"""
from pathlib import Path
from PIL import Image
import numpy as np

RAIZ = Path(__file__).resolve().parent.parent.parent
ELEMENTOS = RAIZ / "Elementos"
SALIDA = Path(__file__).resolve().parent
SALIDA.mkdir(exist_ok=True)

UMBRAL_BLANCO = 245
PADDING = 6


def blanco_a_transparente(img: Image.Image) -> Image.Image:
    """Convierte el fondo casi-blanco en transparente, conservando el trazo."""
    img = img.convert("RGBA")
    datos = np.array(img)
    rgb = datos[:, :, :3].astype(int)
    blancura = rgb.min(axis=2)  # bajo = oscuro/con color, alto = blanco
    alpha = np.clip((255 - blancura) * 255 // (255 - UMBRAL_BLANCO), 0, 255)
    alpha = np.where(blancura >= UMBRAL_BLANCO, 0, 255).astype(np.uint8)
    # anti-alias suave en el borde: usa un segundo umbral para semi-transparencia
    zona_borde = (blancura >= UMBRAL_BLANCO - 25) & (blancura < UMBRAL_BLANCO)
    alpha_suave = np.clip((UMBRAL_BLANCO - blancura) * (255 // 25), 0, 255)
    alpha_final = np.where(zona_borde, alpha_suave, alpha).astype(np.uint8)
    datos[:, :, 3] = alpha_final
    return Image.fromarray(datos, "RGBA")


def autocrop(img: Image.Image, pad=PADDING) -> Image.Image:
    """Recorta al bounding box real del contenido (alpha > 0), con padding."""
    datos = np.array(img)
    alpha = datos[:, :, 3]
    filas = np.where(alpha.any(axis=1))[0]
    cols = np.where(alpha.any(axis=0))[0]
    if len(filas) == 0 or len(cols) == 0:
        return img
    y0, y1 = max(0, filas[0] - pad), min(img.height, filas[-1] + 1 + pad)
    x0, x1 = max(0, cols[0] - pad), min(img.width, cols[-1] + 1 + pad)
    return img.crop((x0, y0, x1, y1))


def guarda_optimizado(img: Image.Image, destino: Path):
    """Paleta indexada (Fast Octree, conserva alpha) -- linea/relleno plano
    comprime muchisimo mejor que RGBA de 32 bits sin perdida visible."""
    paleta = img.convert("RGBA").quantize(colors=64, method=2)
    paleta.save(destino, optimize=True)


def recorta_celda(ruta_origen: str, caja, nombre_salida: str):
    img = Image.open(ELEMENTOS / ruta_origen).convert("RGB")
    celda = img.crop(caja)
    celda_t = blanco_a_transparente(celda)
    celda_t = autocrop(celda_t)
    destino = SALIDA / f"{nombre_salida}.png"
    guarda_optimizado(celda_t, destino)
    print(f"  {nombre_salida}: {celda_t.size} <- {ruta_origen} {caja}")


def main():
    hoja_repuestos = "Captura de pantalla 2026-08-04 214439.png"
    # filas de icono (sin la etiqueta de texto de abajo), grilla de 4 columnas
    fr = {1: (25, 168), 2: (200, 328), 3: (363, 490), 4: (530, 658)}
    cw = 972 / 4
    fc = {c: (round(cw * (c - 1)), round(cw * c)) for c in range(1, 5)}

    def caja_repuesto(fila, col):
        y0, y1 = fr[fila]
        x0, x1 = fc[col]
        return (x0, y0, x1, y1)

    print("Repuestos (hoja de linea tecnica):")
    recorta_celda(hoja_repuestos, caja_repuesto(1, 1), "icono_kit_embrague")
    recorta_celda(hoja_repuestos, caja_repuesto(1, 2), "icono_disco_freno")
    recorta_celda(hoja_repuestos, caja_repuesto(1, 3), "icono_pastillas_freno")
    recorta_celda(hoja_repuestos, caja_repuesto(1, 4), "icono_faro")
    recorta_celda(hoja_repuestos, caja_repuesto(2, 2), "icono_bujia")
    recorta_celda(hoja_repuestos, caja_repuesto(2, 3), "icono_amortiguador")
    recorta_celda(hoja_repuestos, caja_repuesto(2, 4), "icono_filtro_aire")

    hoja_ui = "Captura de pantalla 2026-08-04 214451.png"
    frU = {1: (48, 149), 2: (221, 329), 3: (397, 499), 4: (557, 654)}
    cwU = 970 / 6
    fcU = {c: (round(cwU * (c - 1)), round(cwU * c)) for c in range(1, 7)}

    def caja_ui(fila, col):
        y0, y1 = frU[fila]
        x0, x1 = fcU[col]
        return (x0, y0, x1, y1)

    print("Iconos de interfaz:")
    recorta_celda(hoja_ui, caja_ui(1, 1), "icono_zona")
    recorta_celda(hoja_ui, caja_ui(1, 2), "icono_tienda")
    recorta_celda(hoja_ui, caja_ui(1, 3), "icono_subcategoria")
    recorta_celda(hoja_ui, caja_ui(1, 4), "icono_filtro")
    recorta_celda(hoja_ui, caja_ui(1, 5), "icono_buscar")
    recorta_celda(hoja_ui, caja_ui(1, 6), "icono_exportar")
    recorta_celda(hoja_ui, caja_ui(2, 1), "icono_consulta_rapida")
    recorta_celda(hoja_ui, caja_ui(2, 2), "icono_vehiculo")
    recorta_celda(hoja_ui, caja_ui(3, 3), "icono_vista")
    recorta_celda(hoja_ui, caja_ui(3, 4), "icono_favorito")
    recorta_celda(hoja_ui, caja_ui(4, 1), "icono_tema")

    # Ilustracion completa del hero (composicion ya lista, se usa tal cual).
    # Se recorta un margen fijo primero para excluir el borde redondeado de
    # la tarjeta de la captura original (que no es parte del dibujo).
    hero = Image.open(ELEMENTOS / "Captura de pantalla 2026-08-04 214139.png").convert("RGB")
    MARGEN_BORDE = 16
    hero = hero.crop((MARGEN_BORDE, MARGEN_BORDE, hero.width - MARGEN_BORDE, hero.height - MARGEN_BORDE))
    hero_t = blanco_a_transparente(hero)
    hero_t = autocrop(hero_t, pad=4)
    guarda_optimizado(hero_t, SALIDA / "hero_ilustracion.png")
    print(f"Hero: {hero_t.size}")


if __name__ == "__main__":
    main()

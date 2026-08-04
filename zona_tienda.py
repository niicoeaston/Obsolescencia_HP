# -*- coding: utf-8 -*-
"""
Configuracion editable: zona geografica real de cada tienda.
==============================================================
POR QUE EXISTE ESTE ARCHIVO
---------------------------
La columna "Zona" del Excel viene por FILA de producto, no por tienda, y al
calcular la moda (ver calcular_zona_tienda.py) el resultado no sirve: para
la enorme mayoria de tiendas del pais -incluidas Temuco, Concepcion,
Coquimbo, Osorno o Valdivia- la zona mas frecuente es "Zona V Region", que
geograficamente no tiene sentido. Esa columna parece marcar el origen o
lote del stock, no la ubicacion de la tienda (queda documentado en
DIAGNOSTICO_EXCEL.md seccion 3).

Para un selector publico Zona -> Tienda esa ambiguedad no es aceptable: un
supervisor de Temuco no deberia ver su tienda bajo "V Region". Por eso este
mapeo se construyo a mano, ciudad por ciudad, con la geografia real de Chile
(nombre de la tienda -> region real), no derivado de la columna Zona del
archivo.

COMO EDITARLO
-------------
Es un diccionario plano: "codigo - nombre" (tal como aparece en la base) a
una de las 5 zonas geograficas + 1 zona no geografica "Otras / Nacional".
Para corregir una tienda, cambia el valor de su linea. Para agregar una
tienda nueva que aparezca en un Excel futuro y no este aqui, el generador la
deja fuera del selector de zona y avisa en la consola (no la inventa).

SUPUESTO A REVISAR CON EL NEGOCIO
----------------------------------
Los codigos "Zona RM 1 / RM 2 / RM 3" que trae el Excel parecen ser rutas de
reparto dentro del Gran Santiago, pero el archivo no permite saber con
certeza que comuna cae en cada ruta. En vez de inventar ese reparto, todas
las tiendas de la Region Metropolitana se agrupan aqui en una sola zona
"Zona Metropolitana". Si Autoplanet tiene la division oficial de comunas por
ruta (RM1/RM2/RM3), se puede reemplazar facilmente: solo hay que cambiar el
valor "Zona Metropolitana" por "Zona RM 1", "Zona RM 2" o "Zona RM 3" en las
lineas que corresponda.
"""

ZONA_METROPOLITANA = "Zona Metropolitana"
ZONA_NORTE = "Zona Norte"
ZONA_V_REGION = "Zona Valparaíso"
ZONA_CENTRO_SUR = "Zona Centro Sur"
ZONA_SUR = "Zona Sur"
ZONA_OTRAS = "Otras / Nacional"  # ecommerce, centro de distribucion, discontinuados

# Orden en que se ofrecen en el selector (geograficas primero, de norte a sur)
ORDEN_ZONAS = [
    ZONA_NORTE,
    ZONA_METROPOLITANA,
    ZONA_V_REGION,
    ZONA_CENTRO_SUR,
    ZONA_SUR,
    ZONA_OTRAS,
]

# tienda (codigo - nombre, tal como aparece en la base) -> zona real
ZONA_PRINCIPAL: dict[str, str] = {
    # --- Zona Metropolitana (Gran Santiago) ---
    "AP0001 – La Florida": ZONA_METROPOLITANA,
    "AP0002 – Maipu": ZONA_METROPOLITANA,
    "AP0003 – LF Mirador": ZONA_METROPOLITANA,
    "AP0005 – Puente Alto": ZONA_METROPOLITANA,
    "AP0006 – Quilicura": ZONA_METROPOLITANA,
    "AP0007 – Gran Avenida P.30": ZONA_METROPOLITANA,
    "AP0009 – Maipu Norte": ZONA_METROPOLITANA,
    "AP0010 – Santa Rosa": ZONA_METROPOLITANA,
    "AP0013 – Vivaceta": ZONA_METROPOLITANA,
    "AP0016 – Macul": ZONA_METROPOLITANA,
    "AP0017 – Gabriela": ZONA_METROPOLITANA,  # revisar: nombre ambiguo, sin señal fuerte
    "AP0019 – 10 De Julio": ZONA_METROPOLITANA,
    "AP0024 – Recoleta": ZONA_METROPOLITANA,
    "AP0026 – Melipilla": ZONA_METROPOLITANA,
    "AP0029 – Cerrillos": ZONA_METROPOLITANA,
    "AP0030 – Las Condes": ZONA_METROPOLITANA,
    "AP0031 – Buin": ZONA_METROPOLITANA,
    "AP0037 – Departamental": ZONA_METROPOLITANA,
    "AP0038 – Lo Barnechea": ZONA_METROPOLITANA,
    "AP0044 – Gran Avenida P.20": ZONA_METROPOLITANA,
    "AP0047 – Talagante": ZONA_METROPOLITANA,
    "AP0048 – Rotonda Atenas": ZONA_METROPOLITANA,
    "AP0049 – Gran Avenida P.40": ZONA_METROPOLITANA,
    "AP0056 – San Pablo Pudahuel": ZONA_METROPOLITANA,
    "AP0060 – Penaflor": ZONA_METROPOLITANA,
    "AP0063 – Paseo Quilin": ZONA_METROPOLITANA,
    "AP0066 – San Pablo Quinta Normal": ZONA_METROPOLITANA,
    "AP0067 – Camino Melipilla": ZONA_METROPOLITANA,
    "AP0068 – Colina": ZONA_METROPOLITANA,
    "AP0069 – Camilo Henriquez": ZONA_METROPOLITANA,
    "AP0070 – Orientales": ZONA_METROPOLITANA,  # revisar: nombre ambiguo, sin señal fuerte
    "SG0033 – MELIPILLA": ZONA_METROPOLITANA,

    # --- Zona Norte (Arica a Coquimbo) ---
    "AP0036 – Ovalle": ZONA_NORTE,
    "AP0041 – Coquimbo": ZONA_NORTE,
    "AP0051 – La Serena": ZONA_NORTE,
    "AP0061 – Calama": ZONA_NORTE,
    "AP0065 – Copiapo": ZONA_NORTE,
    "SG0002 – COQUIMBO": ZONA_NORTE,
    "SG0003 – OVALLE": ZONA_NORTE,

    # --- Zona Valparaíso (V Región real) ---
    "AP0012 – Vina Del Mar": ZONA_V_REGION,
    "AP0014 – San Antonio": ZONA_V_REGION,
    "AP0015 – Valparaiso": ZONA_V_REGION,
    "AP0025 – Villa Alemana": ZONA_V_REGION,
    "AP0033 – Quilpue": ZONA_V_REGION,
    "AP0035 – Los Andes": ZONA_V_REGION,
    "AP0042 – San Felipe": ZONA_V_REGION,
    "AP0045 – Concon": ZONA_V_REGION,
    "AP0053 – La Calera": ZONA_V_REGION,
    "AP0057 – Quillota": ZONA_V_REGION,
    "SG0004 – QUILLOTA": ZONA_V_REGION,
    "SG0019 – CASABLANCA": ZONA_V_REGION,
    "SG0021 – VERGARA": ZONA_V_REGION,  # Av. Vergara, Viña del Mar
    "SG0025 – LOS ANDES": ZONA_V_REGION,
    "SG0039 – SAN FELIPE": ZONA_V_REGION,

    # --- Zona Centro Sur (O'Higgins, Maule, Ñuble) ---
    "AP0018 – Rancagua": ZONA_CENTRO_SUR,
    "AP0023 – Curico": ZONA_CENTRO_SUR,
    "AP0028 – Linares": ZONA_CENTRO_SUR,
    "AP0039 – San Fernando": ZONA_CENTRO_SUR,
    "AP0040 – Chillan": ZONA_CENTRO_SUR,
    "AP0050 – Talca Poniente": ZONA_CENTRO_SUR,
    "AP0054 – Talca Oriente": ZONA_CENTRO_SUR,
    "AP0071 – Parral": ZONA_CENTRO_SUR,
    "AP0073 – Santa Cruz": ZONA_CENTRO_SUR,
    "SG0005 – RANCAGUA": ZONA_CENTRO_SUR,
    "SG0006 – CURICO": ZONA_CENTRO_SUR,
    "SG0007 – TALCA": ZONA_CENTRO_SUR,
    "SG0008 – CHILLAN": ZONA_CENTRO_SUR,
    "SG0026 – LINARES": ZONA_CENTRO_SUR,
    "SG0029 – SAN FERNANDO": ZONA_CENTRO_SUR,

    # --- Zona Sur (Biobío, Araucanía, Los Ríos, Los Lagos) ---
    "AP0020 – Concepcion Paicavi": ZONA_SUR,
    "AP0022 – Concepcion Prat": ZONA_SUR,
    "AP0032 – Temuco Leon Gallo": ZONA_SUR,
    "AP0034 – Los Angeles": ZONA_SUR,
    "AP0043 – Temuco Caupolican": ZONA_SUR,
    "AP0046 – Talcahuano": ZONA_SUR,
    "AP0052 – Angol": ZONA_SUR,
    "AP0055 – Osorno": ZONA_SUR,
    "AP0062 – Valdivia": ZONA_SUR,
    "AP0064 – Coronel": ZONA_SUR,
    "AP0072 – Puerto Montt": ZONA_SUR,
    "SG0011 – TEMUCO": ZONA_SUR,
    "SG0028 – PUERTO VARAS": ZONA_SUR,
    "SG0031 – LOS ANGELES": ZONA_SUR,
    "SG0032 – OSORNO": ZONA_SUR,
    "SG0034 – VICTORIA": ZONA_SUR,

    # --- Otras / Nacional (no son tiendas fisicas de barrio) ---
    "0714": ZONA_OTRAS,  # centro de distribucion
    "AP0074 – Ecommerce Autoplanet": ZONA_OTRAS,
    "AP0077 – Descontinuados": ZONA_OTRAS,
}

# Tiendas de baja confianza que conviene confirmar con el negocio
# (nombre sin ciudad reconocible con certeza).
REVISAR_MANUALMENTE = ["AP0017 – Gabriela", "AP0070 – Orientales"]

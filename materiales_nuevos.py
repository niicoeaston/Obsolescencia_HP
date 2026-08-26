# -*- coding: utf-8 -*-
"""
Materiales marcados a mano como "nuevo producto" -- fuente unica compartida
entre generar_portal.py (Vista Rapida + filtro "Solo productos nuevos") y
actualizar_stock.py (agrega la disponibilidad real por tienda de estos
materiales cuando aparece en el archivo diario de stock, ver seccion 15 del
README).

Nunca se infiere automaticamente comparando versiones del Excel: se agrega
el material aqui solo cuando el usuario confirma que es un ingreso nuevo
real. Los lotes anteriores no se quitan (se acumulan) salvo que el usuario
pida explicitamente dejar de destacar alguno.
"""

# Agregado 2026-08-25 (V8): 9 parachoques nuevos disponibles en CD.
MATERIALES_NUEVOS_V8 = {
    "1094160", "1094162", "1094163", "1094154",
    "1107928", "1107905", "1107901", "1107937", "1107902",
}
# Agregado 2026-08-26 (V9): 93 SKU nuevos disponibles en CD (opticos, farolas,
# parachoques, espejos, neblineros).
MATERIALES_NUEVOS_V9 = {
    "1094139", "1094151", "1094152", "1094153", "1094161", "1094167",
    "1107920", "1107925", "1108399", "1108405", "1108406", "1108407",
    "1108408", "1108409", "1108411", "1108412", "1108413", "1108415",
    "1108416", "1108419", "1108421", "1108422", "1108423", "1108424",
    "1108436", "1108442", "1108445", "1108448", "1108451", "1108452",
    "1108453", "1108455", "1108456", "1108458", "1108460", "1108462",
    "1108463", "1108471", "1108473", "1108477", "1108478", "1108479",
    "1108481", "1108484", "1108494", "1108496", "1108499", "1108500",
    "1108501", "1108502", "1108503", "1108504", "1108510", "1108513",
    "1108693", "1108696", "1118582", "1118643", "1118911", "1118971",
    "1119406", "1119407", "1119561", "1119599", "1119602", "1119608",
    "1119769", "1119825", "1119828", "1119829", "1119983", "1120020",
    "1120023", "1120029", "1120076", "1120182", "1120183", "1135970",
    "1135979", "1135981", "1135986", "1135987", "1135999", "487307",
    "487333", "487352", "487358", "487382", "487466", "487494",
    "487515", "487543", "487568",
}
MATERIALES_NUEVOS = MATERIALES_NUEVOS_V8 | MATERIALES_NUEVOS_V9

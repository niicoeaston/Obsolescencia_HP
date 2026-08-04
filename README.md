# Portal de Liquidación de Productos en Obsolescencia — Autoplanet

**Sitio publicado:**
- https://niicoeaston.github.io/
- https://niicoeaston.github.io/Obsolescencia_HP/

(mismo sitio, mismos datos — se publican los dos porque así se pidió; hay
que actualizar ambos cuando cambien los datos, ver sección 2).

Página web comercial (no un dashboard) para que supervisores de tienda,
equipos comerciales y clientes de taller consulten qué productos están en
liquidación por obsolescencia y a qué precio. Recorrido: **Zona → listado de
productos** (de todas las tiendas de la zona), acotando opcionalmente por
**Tienda → Subcategoría → Marca del vehículo → Modelo → Año**, con descarga
del resultado filtrado. Solo la Zona es obligatoria; el resto de los
filtros son para acotar, no para desbloquear resultados.

---

## 1. Qué es cada archivo

| Archivo | Para qué sirve |
|---|---|
| **`index.html`** | El sitio publicado. Un solo archivo autónomo (datos incluidos). **No se edita a mano** — se regenera con el script. |
| `plantilla_portal.html` | La plantilla real: HTML + CSS + JS. Aquí se edita el diseño o el comportamiento. |
| `generar_portal.py` | Lee los dos Excel (listado de productos + resumen de zonas), arma los datos y produce `index.html` a partir de la plantilla. |
| `anios.py` | Interpreta la columna Año (rangos abiertos, años sueltos, formatos varios) sin nunca duplicar filas. Ver sección 3. |
| `resumen_zona.py` | Lee el stock valorizado **oficial** por zona desde `Base No Estrategicos y obsoletos.xlsx`. Ver sección 4. |
| `logo_grupo_planet_b64.txt` | Logo de Grupo Planet en base64 (footer). |
| `logo_autoplanet_b64.txt` | Logo real de Autoplanet en base64 (header). |
| `banner_hero_b64.txt` | Banner de Talleres Autoplanet en base64, re-comprimido a JPEG (fondo del hero). |

El proyecto reutiliza el procesador de datos ya construido y probado en
`../dashboard-obsolescencia/src/services/cargador_excel.py` (limpieza,
validación, conversión de números, consolidación de duplicados). No se
duplicó esa lógica.

---

## 2. Cómo actualizar los datos cada mes

```bash
cd portal-web
python generar_portal.py
```

Esto lee **dos** archivos y regenera `index.html`:

1. **`../dashboard-obsolescencia/data/Listado Obsolecencia.xlsx`** (el mismo
   archivo que usa el dashboard interno) — el listado de productos, ahora
   con las columnas de Marca del vehículo, Modelo, Motor y Año (formato V4).
2. **`Base No Estrategicos y obsoletos.xlsx`** (en la raíz del proyecto,
   hoja `Resumen (2)`) — de aquí sale el número **oficial** de stock
   valorizado por zona que se muestra en la pestaña Análisis (ver sección 4).

El script imprime en pantalla el reparto de stock valorizado por zona y
cualquier advertencia (productos sin marca/modelo/motor/año, discrepancias
en el resumen de zonas) para que puedas revisarlo antes de publicar.

**Para reemplazar el Excel mensual del listado:** sigue las instrucciones de
`../dashboard-obsolescencia/README.md` (reemplazar el archivo en esa carpeta
`data/`, manteniendo el nombre). Si también cambia el archivo de resumen de
zonas, reemplázalo en la raíz del proyecto manteniendo el nombre y la hoja
`Resumen (2)`. Luego vuelve a correr el comando de arriba desde `portal-web/`.

### Publicar la actualización

Este repo tiene **dos remotos** porque el sitio quedó publicado en dos URLs
(ver arriba). Hay que subir a los dos:

```bash
git add index.html
git commit -m "Actualiza datos del mes"
git push origin main            # niicoeaston.github.io
git push obsolecenciahp main    # niicoeaston.github.io/Obsolescencia_HP
```

Si tu sesión de Windows ya inició sesión en GitHub una vez (Git Credential
Manager), el `push` no vuelve a pedir credenciales. GitHub Pages tarda
entre 30 segundos y 2 minutos en publicar el cambio. Los links **no cambian**.

---

## 3. Filtros de aplicación vehicular (Marca, Modelo, Año)

Desde la versión V4 del Excel el listado trae, por producto, a qué vehículo
aplica: Marca del vehículo, Modelo, Motor y Año (o rango de años). Estos se
agregaron como filtros adicionales, en un bloque colapsable **"Filtros de
aplicación"** debajo de Zona/Tienda/Subcategoría (para no saturar la
pantalla principal), y también como columna nueva **"Ver detalle"** en el
listado de resultados (fila expandible, igual en escritorio y en móvil).

### La columna "Marca" duplicada

El Excel V4 trae **dos** columnas llamadas literalmente "Marca": la marca
del repuesto (ej. "KUBOSHI") y la marca del vehículo (ej. "MITSUBISHI").
`pandas` renombra automáticamente la segunda a `"Marca.1"` al leer el
archivo — se confirmó esto empíricamente (no se asumió) antes de mapearla.
Se agregó el alias `"marca_vehiculo": ["Marca.1", "Marca Vehiculo", "Marca
del Vehiculo"]` en `../dashboard-obsolescencia/src/config/settings.py` para
que el cargador ya existente la reconozca sin tocar su lógica general.

Algunas filas (12 en la V4 actual) traen el valor literal `"0"` en esa
columna — parece un residuo de una fórmula de búsqueda del Excel origen que
no encontró coincidencia, no una marca real. Se trata igual que un dato
vacío (no se muestra, no aparece como opción de filtro), y queda anotado en
las advertencias que imprime `generar_portal.py` al generar el sitio.

### Apertura de rangos de año — sin duplicar nada

La columna Año puede traer un año único (`"2015"`) o un rango abierto
(`"2004-2016"`, con distintos tipos de guion, con espacios, con varios
segmentos separados por coma o `/`). El filtro por año necesita poder
buscar un año específico dentro de ese rango — pero **un producto sigue
siendo un producto**: abrir el rango nunca debe multiplicar su stock, su
valorización, ni aparecer dos veces en el listado o en la descarga.

La solución (`anios.py`) no crea filas nuevas. Cada producto sigue
apuntando a un solo string de año (el original, tal cual viene en el
Excel, para mostrarlo sin alterar). Por separado, existe un diccionario
`aniosPorIndice` que —para cada string distinto de año— guarda el conjunto
de años individuales que representa. Filtrar por "2016" busca qué
productos tienen 2016 **dentro de ese conjunto**, sin tocar la fila del
producto. Por construcción, es imposible que esto duplique stock o
valorización: no hay filas nuevas, solo una forma distinta de buscar en las
que ya existen.

Formatos soportados y validados con 11 casos de prueba (`python anios.py`):
año único, rango simple, rango con espacios, guion normal/largo (`–`/`—`),
varios rangos o años separados por coma, rango invertido (inválido) y
formato irreconocible (inválido, se conserva el string original igual). En
los datos reales de la V4 solo aparecen los formatos simples (año único o
`AAAA-AAAA`) — los demás casos están cubiertos por las pruebas del parser,
no por datos reales existentes.

Filas sin año válido (132 en la V4 actual, aplicaciones vacías) quedan
fuera de cualquier filtro por año específico, pero siguen apareciendo
normalmente si no se filtra por año.

---

## 4. Stock valorizado oficial por zona (pestaña Análisis)

**Cambio importante en V4**: el número que se muestra en el donut y la
tabla de la pestaña Análisis ya **no se calcula** en este proyecto (antes
era `stock disponible × valor remate con IVA`, sumado producto por
producto). Ahora se lee **tal cual** desde el archivo oficial `Base No
Estrategicos y obsoletos.xlsx`, hoja `Resumen (2)`, en la tabla dinámica
titulada "Distribución por zona" (columna H en ese archivo).

Motivo del cambio: ese número calculado no coincidía con el número oficial
del negocio — se rastreó la diferencia hasta la fórmula real de esa hoja
(`Stock_Valorizado_AP = Stock AP × Costo unitario`, es decir, a **costo**,
no al precio de remate/liquidación). Son dos métricas legítimas y
distintas (costo vs. precio de venta de liquidación); esta sección del
portal ahora muestra la oficial del negocio en lugar de recalcular una
propia.

`resumen_zona.py` busca el título "Distribución por zona" en la hoja **sin
asumir una fila o columna fija** (por si el archivo cambia de forma), y
lee las filas hasta la primera vacía. Se detectó que la hoja `Resumen (2)`
tiene en realidad **dos bloques** con ese mismo título (columnas H y AK) y
sus totales **no coinciden** ($393.718.464 vs $309.057.215, diferencia de
$84.661.249). Se usa el bloque de la columna H (el pedido explícitamente) y
la diferencia con el otro bloque se deja registrada como advertencia en la
consola al generar el sitio — este proyecto no decide cuál "está bien",
eso le corresponde al negocio.

---

## 5. La columna Zona: historia y estado actual

**Ya está resuelto — usa la columna `Zona` del Excel tal cual.** Queda esta
sección para que quede constancia de por qué, si en algún momento vuelve a
romperse.

Versiones anteriores del Excel (`V2`, `V3`) traían la zona **por fila de
producto, no por tienda**: una misma tienda aparecía repartida hasta en 4
zonas distintas, y calculando la más frecuente el resultado no servía (para
casi todas las tiendas del país salía "Zona V Region", sin sentido
geográfico). En ese momento este proyecto usaba un archivo
`zona_tienda.py` con una clasificación manual por ciudad, como parche.

Esa columna **ya viene corregida en el Excel actual**: cada tienda tiene
exactamente una zona, y son las zonas reales del negocio (incluye el
detalle real Zona RM 1 / RM 2 / RM 3, que antes no se podía inferir). Por
eso `zona_tienda.py` y `calcular_zona_tienda.py` se eliminaron del
proyecto — ya no hacen falta. `generar_portal.py` usa `fila.zona`
directamente y solo imprime un aviso si en el futuro alguna tienda vuelve
a aparecer con más de una zona.

Las 11 zonas actuales, en el orden en que se muestran (definido en
`ORDEN_ZONAS` dentro de `generar_portal.py`): Zona Norte, Zona RM 1, Zona
RM 2, Zona RM 3, Zona V Region, Zona Centro Sur, Zona Sur, Agroplanet, CD,
E-Commerce, AP/SG. Si aparece una zona nueva que no está en esa lista, el
script la agrega igual al final del selector y avisa en la consola — nunca
desaparece en silencio.

---

## 6. Arquitectura y por qué

**Un solo archivo HTML estático**, sin backend, sin build, publicado en
**GitHub Pages**, gratuito, sin servidor que mantener.

- **Por qué no Streamlit** (como la versión interna en
  `dashboard-obsolescencia/`): Streamlit necesita un proceso corriendo; no
  se puede compartir como un link que funcione para cualquiera sin que tu
  computador esté encendido. Este portal es exactamente para eso: un link
  público, permanente, sin instalación.
- **Por qué GitHub Pages y no Claude Artifacts**: la primera versión se
  publicó como Artifact y **no cargaba en otros dispositivos** (los
  artifacts nacen privados; sin compartirlos explícitamente, un dispositivo
  sin sesión ve una pantalla en blanco). GitHub Pages es un sitio público
  real, con su propio dominio, sin ninguna dependencia de sesión.
- **Por qué los datos van embebidos en el HTML** y no en una base de datos:
  mantiene "actualizar = reemplazar un archivo y volver a generar", sin
  servidor ni API que mantener. El costo es un archivo de ~1.4 MB, aceptable
  para este volumen (17.594 productos + logos + banner).
- **Por qué el link no quedó como un dominio propio tipo "obsolecenciahp"**:
  GitHub Pages para cuentas personales publica en `usuario.github.io`, fijo
  al nombre de la cuenta de GitHub. Para tener un dominio completamente
  distinto hace falta comprar un dominio real y conectarlo por DNS. Lo que
  sí se puede (y se hizo) es publicar un segundo repo cuyo nombre aparece
  como sub-ruta: `niicoeaston.github.io/Obsolescencia_HP/`.

### Identidad visual

Colores muestreados directamente de **autoplanet.cl** (rojo del hero
`#FF061A`, naranja de la barra de categorías `#E65100`) y del **logo real
de Grupo Planet** (rojo `#D7141A`, usado en el footer). El sistema de
diseño se validó primero en Stitch (proyecto "Autoplanet Portal Comercial")
antes de programarlo.

El header usa el **logo real de Autoplanet** (`logo_autoplanet_b64.txt`).
El fondo del hero usa el **banner real de Talleres Autoplanet**
(`banner_hero_b64.txt`, comprimido a JPEG ~100 KB) con un degradado oscuro
encima para que el texto blanco siga siendo legible sobre cualquier parte
de la imagen. El recorte (`background-position` / `background-size` en
`plantilla_portal.html`) está ajustado a mano para que ni el logo
"Talleres Autoplanet" ni la franja naranja de categorías de la imagen
original queden visibles — si se cambia el banner por otro archivo con
proporciones distintas, hay que volver a ajustar esos valores a ojo.

---

## 7. Qué se dejó fuera de la vista principal (a propósito)

Sin tarjetas KPI, sin gráfico de Top 12 subcategorías/marcas, sin columna
de descuento % ni "Total remate" en el listado de productos. El **valor
remate con IVA** (mostrado como "Precio Liquidación" en el selector de
orden) es el precio protagonista, junto al precio normal tachado como
referencia (igual que una vitrina de liquidación real).

El **stock disponible sí se muestra** (columna en la tabla de escritorio,
línea destacada en la tarjeta móvil, y también en la descarga CSV) — es
información operativa importante para decidir si vale la pena ir a buscar
el producto a esa tienda.

Los filtros de aplicación vehicular (Marca, Modelo, Año) y el detalle de
aplicación (Marca del vehículo / Modelo / Motor / Años) van en un bloque
colapsable y en una fila expandible "Ver detalle" respectivamente — no en
la vista principal — para no convertir esto en un formulario largo (ver
sección 3).

La sección **Análisis** (pestaña aparte, no la vista principal) muestra el
stock valorizado **oficial** por zona (sección 4) en un donut + tabla de
participación. Con 11 zonas reales, el donut solo colorea distinto las 5
más grandes y agrupa el resto en un segmento gris "Otras zonas" (más de
~5-6 colores en un gráfico circular deja de ser legible y dos colores
empiezan a repetirse). La **tabla de abajo nunca agrupa**: siempre lista
las 11 zonas por separado con su valor exacto.

---

## 8. Seguridad y qué expone el sitio

El repositorio y el sitio son **públicos** (requisito del hosting gratuito
de GitHub Pages) — indexable por buscadores, sin login. Por diseño, el
listado de productos solo expone: material, texto breve, marca,
subcategoría, marca/modelo/motor/año del vehículo de aplicación, stock
disponible, precio normal (tachado) y valor remate con IVA. **No expone**
costos, márgenes, contribución ni datos de otras tiendas fuera de la
selección activa. Esto es equivalente a un catálogo de ofertas público, no
información comercial sensible.

---

## 9. Pruebas realizadas

- Recorrido completo Zona → Tienda → Subcategoría → Marca/Modelo/Año →
  resultados, con datos reales V4 (17.594 productos, 88 tiendas, 11 zonas
  reales, 178 subcategorías).
- **Resultados con solo la Zona seleccionada** (sin Tienda): se ven los
  productos de todas las tiendas de esa zona a la vez, con una columna
  "Tienda" nueva en la tabla (y una línea equivalente en la tarjeta móvil)
  para identificar de qué tienda es cada producto; los filtros de
  Marca/Modelo/Año y la descarga CSV funcionan igual de bien cruzando
  varias tiendas. Se corrigió un problema real de layout en tablet (768px)
  donde el nombre de la subcategoría se superponía con la columna Stock
  cuando Tienda y Subcategoría se mostraban a la vez sin acotar (causado
  por un orden de reglas CSS invertido); ahora ambas etiquetas se acotan en
  ancho y saltan de línea en vez de desbordar.
- Buscador, orden (incluido ordenar por Stock disponible), "Ver todas las
  subcategorías", descarga CSV con las columnas nuevas (Marca vehículo,
  Modelo, Motor, Rango de años original, Años aperturados, Stock
  disponible).
- **Sin duplicación al filtrar por año abierto**: se verificó
  explícitamente que filtrar por un año específico dentro de un rango
  (ej. 1996 dentro de "1993-2000") reduce el conteo de productos sin que
  ningún material se repita, y que stock/valorización bajan de forma
  consistente con el subconjunto filtrado (nunca se multiplican).
- Filtros Marca del vehículo / Modelo / Año: cascada de reseteo correcta
  (cambiar Zona limpia Tienda, Subcategoría y los tres filtros de
  aplicación); "Limpiar filtros" resetea los seis campos.
- "Ver detalle" (fila expandible) probado en escritorio y en móvil,
  mostrando el mismo contenido en ambos; producto sin aplicación muestra
  el texto de respaldo "Información no disponible" en vez de dejar el
  espacio vacío o mostrar "nan".
- Sección Análisis: donut agrupado (top 5 + "Otras zonas") + tabla completa
  sin agrupar, coinciden los números y suman el total oficial
  ($393.718.464) leído de `Resumen (2)`.
- Tema claro/oscuro, incluido el donut repintándose con paleta distinta.
- Anchos móviles 390×844 sin desborde horizontal, incluyendo el bloque
  colapsable de filtros de aplicación.
- Tablet (768px): tabla simplificada (sin columna Marca, padding reducido,
  botón "Ver detalle" compactado a solo el ícono) para que Stock, Precio y
  Ver detalle quepan sin scroll horizontal.
- **Verificado en el sitio publicado real** (los dos links), no solo en
  local.

## 10. Limitaciones conocidas

- La exportación es **solo CSV** (se abre perfecto en Excel en español, con
  `;` y BOM). No se generó un `.xlsx` real para evitar sumar una
  dependencia externa a un sitio que debe cargar rápido.
- El sitio es público; no hay una vista distinta para "cliente" vs
  "supervisor" en esta primera versión — ambos ven exactamente lo mismo
  (que ya excluye todo dato sensible).
- El recorte del banner del hero está ajustado a mano para esta imagen
  específica; si se reemplaza el archivo del banner por uno con otras
  proporciones, hay que reajustar `background-position`/`background-size`
  en `plantilla_portal.html`.
- Dos sitios publicados (misma cuenta) significa dos `git push` cada vez
  que se actualizan los datos — si se olvida uno, quedan desincronizados.
- La hoja `Resumen (2)` tiene dos bloques "Distribución por zona" con
  totales distintos (ver sección 4); este proyecto usa el de la columna H
  y deja la discrepancia como advertencia, sin resolverla — es una
  decisión que le corresponde al negocio, no a este portal.

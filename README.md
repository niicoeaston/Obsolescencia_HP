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
| `generar_portal.py` | Lee el Excel del listado de productos, arma los datos y produce `index.html` a partir de la plantilla. |
| `anios.py` | Interpreta la columna Año (rangos abiertos, años sueltos, formatos varios) sin nunca duplicar filas. Ver sección 3. |
| `logo_grupo_planet_b64.txt` | Logo de Grupo Planet en base64 (footer). |
| `logo_autoplanet_b64.txt` | Logo real de Autoplanet en base64 (header). |
| `banner_hero_b64.txt` | Banner de Talleres Autoplanet en base64, re-comprimido a JPEG (ya no se usa como fondo del hero desde el rediseño "north star", ver sección 10, pero se deja por si se necesita). |
| `assets_src/` | Ilustración del hero e íconos del rediseño visual (PNG con fondo transparente), y el script que los recorta desde las referencias del usuario. Ver sección 10. |
| `lib/xlsx.full.min.js` | Librería SheetJS (MIT, v0.18.5) vendorizada para generar el Excel de "Mi selección" en el navegador, sin backend. Ver sección 11. |

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

Esto lee **`../dashboard-obsolescencia/data/Listado Obsolecencia.xlsx`** (el
mismo archivo que usa el dashboard interno) — el listado de productos, con
las columnas de Marca del vehículo, Modelo, Motor y Año (formato V4) — y
regenera `index.html` con los datos nuevos.

El script imprime en pantalla cualquier advertencia (productos sin
marca/modelo/motor/año, tiendas con más de una zona, años no interpretables)
para que puedas revisarlo antes de publicar.

**Para reemplazar el Excel mensual:** sigue las instrucciones de
`../dashboard-obsolescencia/README.md` (reemplazar el archivo en esa carpeta
`data/`, manteniendo el nombre). Luego vuelve a correr el comando de arriba
desde `portal-web/`.

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

## 4. La columna Zona: historia y estado actual

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

## 5. Arquitectura y por qué

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

## 6. Qué se dejó fuera de la vista principal (a propósito)

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

No hay una sección de análisis/gráficos aparte: el portal es solo
consulta y descarga del listado de productos, sin una vista de agregados
por zona.

---

## 7. Seguridad y qué expone el sitio

El repositorio y el sitio son **públicos** (requisito del hosting gratuito
de GitHub Pages) — indexable por buscadores, sin login. Por diseño, el
listado de productos solo expone: material, texto breve, marca,
subcategoría, marca/modelo/motor/año del vehículo de aplicación, stock
disponible, precio normal (tachado) y valor remate con IVA. **No expone**
costos, márgenes, contribución ni datos de otras tiendas fuera de la
selección activa. Esto es equivalente a un catálogo de ofertas público, no
información comercial sensible.

---

## 8. Pruebas realizadas

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
- Tema claro/oscuro.
- Anchos móviles 390×844 sin desborde horizontal, incluyendo el bloque
  colapsable de filtros de aplicación.
- Tablet (768px): tabla simplificada (sin columna Marca, padding reducido,
  botón "Ver detalle" compactado a solo el ícono) para que Stock, Precio y
  Ver detalle quepan sin scroll horizontal.
- **Verificado en el sitio publicado real** (los dos links), no solo en
  local.

## 9. Limitaciones conocidas

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

---

## 10. Rediseño visual "north star"

El portal se rediseñó visualmente tomando como referencia una imagen
(`Elementos/north star.png`) que el usuario compartió: fondo blanco,
paleta AutoPlanet, ilustraciones vectoriales lineales, hero con
titular/CTA/beneficios, sección "Vista rápida" y panel de detalle lateral.
**Se mantuvo intacta toda la lógica funcional** (filtros, búsqueda, carga
de datos, CSV, responsive, publicación) — el cambio es de diseño, no de
arquitectura.

### Origen de los assets (sin IA — créditos agotados)

El generador de imágenes con IA (nano-banana) no tenía créditos de prepago
disponibles. El usuario dejó en `Elementos/` capturas de un set de íconos
de línea técnica ya diseñado (composición del hero, 16 íconos de
repuestos, 24 íconos de interfaz). `portal-web/assets_src/recortar_iconos.py`
recorta cada ícono individual de esas hojas de referencia (detectando la
grilla y el bounding box real del trazo, excluyendo las etiquetas de texto
en inglés que eran solo para identificar cada ícono), vuelve transparente
el fondo blanco y cuantiza a paleta indexada (Fast Octree) para que el
peso final sea manejable (~120 KB en total para 17 assets, incluida la
ilustración del hero). No se generó ni inventó ningún gráfico nuevo.

### Piezas nuevas

- **Hero**: fondo blanco, titular con jerarquía de color ("Productos" en
  rojo), ilustración de repuestos a la derecha (kit de embrague, discos de
  freno, pastillas, óptico, terminales de dirección, bujía, amortiguador,
  filtro de aire), CTA real (antes era solo una etiqueta de estado) y tres
  recuadros de beneficio.
- **Vista rápida**: sección nueva, **dinámica según la Zona filtrada**
  (recalculada en JS en cada `render()`, no precomputada en Python). Muestra
  5 categorías (Kit de Embrague, Disco de Freno, Bujía, Amortiguador, Filtro
  de Aire) — para cada una se elige el **producto real con mayor stock
  disponible dentro de la zona activa** (o a nivel nacional si no hay zona
  seleccionada); nunca se inventan productos ni precios, y si una categoría
  no tiene stock en la zona activa esa tarjeta simplemente se omite. Es un
  carrusel horizontal con snap. El mapeo categoría → subcategorías reales
  vive en `CATEGORIAS_DESTACADAS` dentro de `plantilla_portal.html`.
- **Panel de detalle lateral**: al hacer clic en una tarjeta de Vista
  Rápida se abre un panel (se convierte en pantalla completa en móvil) con
  aplicación vehicular, zona, tienda, subcategoría, stock y un botón "Ver
  resultados filtrados" que aplica esa zona/tienda/subcategoría a la
  consulta real y hace scroll a los resultados.
- **Favoritos**: la estrella de cada tarjeta de Vista Rápida se guarda en
  `localStorage` (no requiere cuenta ni backend).
- **Modo oscuro para los íconos nuevos**: como son imágenes (no SVG con
  `currentColor`), se invierten con `filter:invert(.9) hue-rotate(180deg)`
  en tema oscuro — esto vuelve el trazo negro a blanco y devuelve el rojo
  a rojo (en vez de invertirlo a cian), sin necesidad de generar una
  segunda versión de cada ícono.

### Ajustes tras la revisión del usuario

- El mensaje informativo ("Selecciona una zona...") usaba una paleta azul
  (`--azul-info`) que no pertenecía a la identidad AutoPlanet. Se renombró a
  `--info`/`--info-fondo`/`--info-borde` con tonos grises neutros,
  consistentes con el resto de la paleta.
- Varios íconos nuevos (el pin de Zona, el ícono de vehículo, el ojo de
  Vista Rápida, etc.) se veían desproporcionados: las reglas CSS los
  forzaban a un cuadrado exacto (`width` y `height` iguales) ignorando el
  aspecto real de cada recorte (por ejemplo, el pin de Zona es 71×101px,
  bastante más alto que ancho). Se corrigió a `max-width`/`max-height` +
  `width:auto;height:auto` en todos los íconos nuevos, preservando su
  proporción real.

### Qué NO cambió

Los filtros (Zona/Tienda/Subcategoría/Marca/Modelo/Año), la búsqueda, la
apertura de años, la tabla/tarjetas de resultados con "Ver detalle", la
descarga CSV y el pipeline de generación/publicación siguen exactamente
igual — solo se les actualizó el estilo visual (íconos junto a las
etiquetas, botón de descarga, "Limpiar filtros").

---

## 11. "Lista de selección" ("Mi selección")

Herramienta de trabajo interna para que un supervisor arme un listado de
SKU antes de solicitar una revisión o traslado. **No es un carrito, no
reserva stock, no modifica datos fuente ni el análisis** — vive solo en
`localStorage` del navegador, como pidió el usuario explícitamente. Primera
etapa: sin login, sin aprobaciones, sin checkout (ver el pedido original si
se retoma esto en una segunda etapa).

### Cambio de categorías en Vista Rápida

La imagen de referencia (`referencia_visual_final_lista_seleccion.png`)
muestra Vista Rápida con **Kit de embrague, Disco de freno, Pastillas de
freno, Faro delantero y Bujía de encendido** — dos categorías distintas a
las que se habían acordado en la ronda de rediseño anterior (que incluía
Amortiguador y Filtro de Aire). Como la imagen reemplaza cualquier diseño
previo por instrucción explícita, se actualizó `CATEGORIAS_DESTACADAS` en
`plantilla_portal.html` a estas 5. "Faro delantero" mapea a las 4
subcategorías reales que representan lo mismo en la base (`FAROL
DERECHO/IZQUIERDO`, `OPTICO DERECHO/IZQUIERDO`); "Pastillas de freno" a
`PASTILLA DE FRENO DELANTERO/TRASERO`. Los 2 íconos que faltaban (pastillas
de freno, faro) se recortaron de la misma hoja de referencia ya usada
antes — nada generado con IA.

### Adaptación visual señalada (sección 5 del pedido)

La referencia muestra cada línea de la selección como una fila de tabla
con 7 columnas fijas (Material/Descripción·SKU, Tienda·Zona, Cantidad,
Stock, Precio, Total, Observación). A los anchos reales de un drawer
(~420–560px en escritorio, 100% en móvil), una tabla literal de 7 columnas
se vuelve ilegible o fuerza scroll horizontal. Se implementó cada línea
como una **tarjeta compacta de 2 filas** que muestra exactamente los mismos
datos (ícono, nombre/SKU, tienda/zona con chip de color, cantidad,
stock, precio, total, observación, eliminar), reorganizados para que quepan
sin comprimir texto ni scroll horizontal. La intención visual (jerarquía,
colores, chips de zona, total en rojo) se mantiene igual a la referencia.

### Arquitectura

- **Estado**: `SELECCION = {lineas: [...]}`, un array de líneas en memoria,
  persistido completo en `localStorage['ap-seleccion-v1']` en cada cambio.
- **ID único de línea**: `material + código de tienda + zona` (sección 16
  del pedido) — el mismo SKU en tiendas distintas queda en líneas separadas;
  agregar el mismo registro dos veces incrementa la cantidad existente en
  vez de duplicar la línea.
- **Captura de precio/stock**: al agregar, se guarda el stock y precio
  *de ese momento* (`stockCapturado`, `precioCapturado`). El total de línea
  siempre usa el precio capturado (no cambia solo si el precio cambió en la
  base) — el `precio actual` se lee en vivo desde `D.filas` solo para
  comparar y advertir, nunca para sobrescribir silenciosamente.
- **Reglas de compatibilidad de zona**: `categoriaZona()` mapea las 11
  zonas reales de la base a las categorías de la matriz pedida (RM1/RM2/
  RM3/Norte/Sur/Ecommerce/CD). Las zonas reales que la matriz no menciona
  (*Zona V Region, Zona Centro Sur, Agroplanet, AP/SG*) caen en la regla
  "cualquier otra zona solo se combina consigo misma". La zona base es la
  primera línea que **no** sea Ecommerce/CD; si la selección arranca con
  Ecommerce y/o CD, la zona base queda "pendiente" hasta que se agregue la
  primera RM, que la fija de forma permanente (sección 20 del pedido).
  Toda la lógica vive en `validarCompatibilidadZona()`.
- **Excel**: se genera 100% en el navegador con **SheetJS** (vendorizada en
  `lib/xlsx.full.min.js`, cargada con un `<script src>` normal — no un CDN,
  la sirve GitHub Pages igual que cualquier archivo del repo). Dos hojas:
  "Selección" (una fila por línea + fila de totales) y "Resumen" (totales
  generales + desglose por zona/tienda/subcategoría).

### Limitación real de la librería de Excel

SheetJS Community Edition (la gratuita, MIT) **no escribe estilos de celda
completos** en la versión vendorizada aquí (colores de fondo/negrita en
encabezados es una función de la edición Pro). Se implementó lo que sí es
nativo de la edición gratuita: **autofiltro en los encabezados** y **anchos
de columna ajustados**. Los encabezados no salen en negrita — es una
limitación de la librería, no un pendiente de implementación.

### Reglas de cantidad y stock

El tope de cada línea es el **stock actual en vivo** (no el capturado) —
si el stock bajó desde que se agregó, no se puede subir la cantidad más
allá del nuevo tope, pero **la cantidad ya guardada no se recorta sola**
(sección 29 del pedido): se muestra tal cual con una advertencia, y el
usuario decide si la ajusta.

### Qué se guarda y qué se detecta al recargar

`localStorage` guarda cada línea completa (SKU, tienda, zona, cantidades,
observación, precio/stock capturados). Al cargar la página,
`revisarSeleccionAlCargar()` compara cada línea contra `D.filas` actual y
avisa (toast) si algún producto ya no existe en la base o si cambió de
stock/precio — sin corregir nada solo.

### Pruebas realizadas

Se automatizaron y verificaron (no solo visualmente, revisando el estado
real): agregar mismo registro dos veces → incrementa, no duplica; mismo
SKU en tiendas distintas → líneas separadas; RM1 + Ecommerce + CD →
permitido; RM1 + RM2 → bloqueado con el mensaje exacto pedido; Norte + Sur
→ bloqueado; iniciar con Ecommerce + CD, fijar base con RM2, intentar RM1
después → bloqueado; cantidad no puede superar el stock (clamps con
mensaje); observación se abre/edita/guarda; eliminar línea actualiza
contador y totales; `localStorage` persiste tras recargar la página;
exportación a Excel con las 2 hojas y datos numéricos reales (verificado
abriendo el `.xlsx` generado con `openpyxl`); responsive en escritorio,
tablet y 390px sin scroll horizontal; tema oscuro (chips de zona con
variante oscura propia, no solo invertida); Escape cierra el drawer y el
modal de confirmación de "Vaciar selección".

### Instrucciones de publicación

Sin cambios respecto a lo ya documentado (sección 2) — mismo `git push` a
los dos remotos. Lo único nuevo es que ahora hay que agregar también
`portal-web/lib/xlsx.full.min.js` al commit (una sola vez; no cambia salvo
que se actualice la versión de la librería).

---

## 12. Correcciones puntuales (línea separadora + Vista Rápida dinámica)

### Línea separadora de la tabla de resultados

**Causa identificada**: la última celda de cada fila (`<td class="der
acciones-fila">`, con los botones "Agregar" y "Ver detalle") tenía
`display:flex` aplicado **directamente sobre el `<td>`**. Un `<td>` con
`display:flex` deja de estirarse a la altura real de la fila (que el
layout de tabla sí le exige a las demás celdas) y en cambio se ajusta a la
altura de su contenido — los botones, más bajos que el texto envuelto en
otras columnas. El resultado: el `border-bottom` de esa celda quedaba
~10-14px más arriba que el de las demás celdas de la misma fila, generando
el "escalón" reportado.

Confirmado **midiendo el DOM real** (`getBoundingClientRect()` de cada
`<td>` de una misma fila), no solo a ojo: antes del fix, la celda de
acciones media 60.67px de alto contra 74.25px del resto de la fila en la
misma `<tr>`; después del fix, las 8 celdas de cada fila miden exactamente
lo mismo y su `border-bottom` cae en el mismo píxel.

**Corrección**: se movió `display:flex` del `<td>` a un `<div
class="acciones-fila">` **dentro** del `<td>` (`plantilla_portal.html`,
función `filaTablaHtml()`). El `<td>` vuelve a comportarse como celda de
tabla normal (hereda `vertical-align:middle` de la regla base `tbody td`)
y el `div` interno sigue centrando los dos botones con flexbox. Se
verificó que abrir/cerrar "Ver detalle" no altera la altura de la fila
principal (mismo alto antes/durante/después de expandir), y que el fix se
sostiene en escritorio y tablet (768px, con el botón compactado).

### Vista Rápida: de fija a 100% dinámica

**Antes**: `calcularDestacados()` tenía una lista fija de 5 categorías
(Kit de Embrague, Disco de Freno, Pastillas de Freno, Faro Delantero,
Bujía) con sus subcategorías asociadas, y filtraba **solo por Zona**,
ignorando Tienda/Subcategoría/Marca/Modelo/Año/búsqueda. Era, tal como se
describió, una segunda lógica de filtrado paralela a `filtrar()`.

**Ahora**: `render()` calcula `filtrar()` **una sola vez** y ese mismo
array (`filas`) se pasa tanto a `pintarResultados()` (tabla/tarjetas
principales) como a `pintarVistaRapida(filas)` — una única fuente de
verdad, sin lógica de filtrado duplicada.

**Función usada para el top por stock** — `seleccionarTopStock(filas,
limite)`:
1. Excluye registros con stock nulo, no numérico, negativo o cero
   (`Number.isFinite(f[ST]) && f[ST] > 0`).
2. Ordena una **copia** del array (`.slice().sort(...)`, nunca muta
   `filas` ni `D.filas`) de mayor a menor stock.
3. Desempate estable y predecible: stock descendente → material ascendente
   → código de tienda ascendente (nunca aleatorio).
4. Deduplica por el mismo identificador que usa la Lista de selección
   (`material + código de tienda + zona`, función `idLinea()`).
5. Devuelve los primeros `limite` (5) — sin forzar variedad de
   subcategoría: si los 5 productos con más stock son de la misma
   subcategoría, se muestran los 5.

Ya no hay ningún listado hardcodeado ni SKU de ejemplo — los íconos de
categoría (`ICONO_SUBCAT_MAPA`) se conservan solo como *estética* (mapean
~9 subcategorías reales a un ícono dedicado), con un ícono genérico de
respaldo para las otras ~169 subcategorías de la base.

**Reactividad**: como `pintarVistaRapida()` ahora se llama dentro de
`render()` con el resultado fresco de `filtrar()`, se actualiza
automáticamente ante cualquier cambio de Zona/Tienda/Subcategoría/Marca/
Modelo/Año/búsqueda — sin recargar la página ni acciones adicionales.

**Estado vacío**: si el filtro activo no deja ningún producto con stock
válido, se muestra "No encontramos productos disponibles para los filtros
seleccionados. Prueba modificando o limpiando algunos filtros." en vez de
tarjetas vacías o de volver a un contenido fijo.

**Validación de desarrollo** (sección 19 del pedido): `seleccionarTopStock`
y `pintarVistaRapida` tienen `console.debug` (resultados filtrados, top por
stock con su ID, registros excluidos por stock inválido, filtros activos)
detrás de la constante `MOSTRAR_DEBUG_VISTA_RAPIDA` — **queda en `false`**
en este commit; se puede poner en `true` temporalmente para depurar.

### Supuestos aplicados

- La cantidad de tarjetas visibles se mantuvo en 5 (la ya aprobada en el
  rediseño visual), tanto en escritorio como en móvil.
- El ícono por categoría (`ICONO_SUBCAT_MAPA`) se conserva únicamente como
  mejora visual para las ~9 subcategorías con ilustración dedicada — no
  como criterio de selección de productos. **Nota**: esta parte quedó
  obsoleta con el ajuste de la sección siguiente, que quitó el ícono de
  las tarjetas de Vista Rápida.
- El texto de contexto de Vista Rápida ("Los productos con más stock
  en…") ahora lista todos los filtros activos (no solo Zona), separados
  por "·", igual que el resumen de resultados principales.

### Archivos modificados

Solo `portal-web/plantilla_portal.html` (CSS y JS). No se tocó
`generar_portal.py`, el pipeline de datos, ni ningún archivo fuente.

### Ajuste posterior: tarjetas de Vista Rápida sin ícono

El usuario notó que, al mostrar el ícono dedicado solo en ~9 de las ~178
subcategorías reales, la mayoría de las tarjetas caían al ícono genérico
(un cuadro/portafolio) mientras unas pocas mostraban una ilustración de
línea real — inconsistencia visual entre tarjetas. Pedido: quitar la
imagen por completo y dejar únicamente el cuadro con texto, para todas
las tarjetas por igual.

**Cambio**: en `tarjetaDestacadoHtml(d, i)` se eliminó el `<div
class="vr-icono">${icono}</div>` (y el cálculo de `icono` que ya no se
usaba en esa función). La tarjeta ahora empieza directamente en el
nombre del producto, seguida de código+marca, tienda/zona, precio y
stock — igual para todas, sin importar si la subcategoría tenía o no un
ícono dedicado.

`iconoHtmlParaSubcategoria()` **no se tocó**: sigue usándose en el panel
de detalle (`abrirPanelDestacado`, `.panel-icono`) y en cada línea de la
Lista de selección (`.ms-linea-icono`), ninguno de los dos alcanzados por
este pedido (que se limitó explícitamente a "los productos de la vista
previa", es decir, las tarjetas de Vista Rápida).

Se ajustó también la regla CSS `.vr-nombre` (se le agregó
`padding-right:26px`) para que el título nunca quede debajo del botón de
favorito (⭐, posicionado `absolute` arriba a la derecha de la tarjeta)
ahora que no hay ícono que empuje el texto hacia abajo. Verificado con
Playwright en tema claro y oscuro, y en ancho móvil (390px): las 8+
tarjetas probadas quedan uniformes, sin superposición del título con el
botón de favorito y sin huecos donde antes iba el ícono.

## 13. Actualización diaria de stock ("Agotado" + `actualizar_stock.py`)

A partir del 2026-08-06 el usuario deja diariamente un archivo de stock real
(exportado desde SAP) en `Actualizacion de Stock/Stock_DD_MM.xlsx` (día/mes
de hoy), con el stock "Libre utilización" por producto y centro/tienda. Se
construyó un pipeline para mantener el "Stock AP" del listado maestro
(`dashboard-obsolescencia/data/Listado Obsolecencia.xlsx`, hoja `Base`) al
día con ese archivo, sin perder productos que llegan a stock 0.

### Por qué antes no se podía dejar un producto en stock 0

El cargador compartido (`dashboard-obsolescencia/src/services/cargador_excel.py`,
usado tanto por el portal como por el dashboard interno) descartaba del
listado cualquier fila con `stock > STOCK_MINIMO` estricto (STOCK_MINIMO=0),
es decir, **stock exactamente 0 se eliminaba del listado por completo** en
vez de mostrarse como agotado. Se cambió esa comparación a `>=`
(`_validar()`, línea ~315): con STOCK_MINIMO=0 el filtro ahora significa
literalmente "sin mínimo", y dejó de excluir los ceros. Esto afecta también
al dashboard interno (comparten el mismo cargador), pero es una corrección
de semántica, no un cambio de comportamiento deseado en otro lado.

### Cómo se ve "Agotado" en el portal

En `plantilla_portal.html`, tanto `filaTablaHtml()` (tabla de escritorio)
como `tarjetaProdHtml()` (tarjetas móviles) revisan `f[ST] <= 0`: si es
así, la celda/línea de stock muestra un badge `.badge-agotado` ("Agotado")
en vez del número, y el botón "Agregar a mi selección" queda `disabled`
(con `title="Producto agotado"`). `agregarASeleccion()` además valida
`item.stock <= 0` como segunda barrera (por si se llega a invocar por otra
vía) y muestra un mensaje de error en vez de agregar la línea. Los
productos agotados siguen apareciendo en los filtros normales (Zona,
Tienda, Subcategoría, búsqueda) — solo cambia cómo se muestra su stock y
que no se pueden agregar a la selección. Vista Rápida ya los excluía
automáticamente (`seleccionarTopStock` solo considera `stock > 0`), así
que no requirió cambios.

### `portal-web/actualizar_stock.py`

Script nuevo que cruza el archivo diario contra el listado maestro y
actualiza el stock en el Excel real (no en el portal generado). Reglas de
negocio, confirmadas explícitamente con el usuario el 2026-08-06:

- **Cruce Centro/Tienda**: la columna "Nombre 1" del archivo diario trae la
  misma etiqueta de tienda que usa el maestro (ej. "AP0001-La Florida"), así
  que el cruce es por nombre exacto — excepto el Centro de Distribución
  (Centro SAP `0714`), cuyo "Nombre 1" es un nombre de operador ("Ega-Kat")
  y no una tienda; para ese caso se usa el código fijo `"0714"`, que es
  como aparece la fila de CD en el maestro.
- **Solo almacén `1100`**: filas del archivo diario con otro código de
  almacén (ej. `1600`) se ignoran por completo, aunque la tienda sí
  aparezca ese día (elegido explícitamente por el usuario en vez de sumar
  todos los almacenes).
- **Tienda ausente = no se toca**: si una tienda del maestro no aparece en
  absoluto en el archivo del día (en ningún almacén), su stock se deja
  intacto y se reporta como "tienda ausente" — para no marcar como agotado
  un local completo por un posible error de exportación SAP.
- **Productos nuevos se ignoran y se reportan**: si el archivo diario trae
  un Material+Tienda que no existe en el listado maestro, no se agrega
  (falta precio, marca, categoría, aplicación vehicular — datos que el
  archivo diario no trae) pero se lista aparte para revisión manual.
- **Nunca se escribe stock negativo**: SAP puede traer "Libre utilización"
  negativa (ajustes de inventario). Escribirla tal cual hacía que esas
  filas fueran descartadas por completo por el filtro de stock del
  cargador (`stock >= 0`), en vez de quedar marcadas como Agotado. Se
  detectó con datos reales (7 filas el primer día) y se corrigió con
  `nuevo = max(0, ...)` antes de escribir.
- **Respaldo automático**: antes de sobrescribir, copia el listado maestro
  a `Actualizacion de Stock/respaldos/Listado Obsolecencia (antes de
  Stock_DD_MM).xlsx`.

Uso: `python actualizar_stock.py` (usa la fecha de hoy) o
`python actualizar_stock.py DD MM` (fuerza fecha, para pruebas/rezagados).
Imprime un reporte con tiendas cubiertas/ausentes y el detalle de productos
que subieron, bajaron, pasaron a agotado, se recuperaron o se ignoraron por
ser nuevos.

### Automatización diaria (cron de la sesión)

Se programó un cron (`CronCreate`, 10:04am todos los días) que: busca el
archivo del día, corre `actualizar_stock.py` y `generar_portal.py`, y le
muestra al usuario un resumen del reporte — **sin publicar** en GitHub
Pages hasta que el usuario confirme explícitamente (decisión tomada con el
usuario: prefiere revisar el resumen antes de publicar, al menos al
comienzo). Limitación importante que se le explicó al usuario: este cron
vive solo dentro de la sesión de Claude Code actual (no se guarda en
disco) y expira automáticamente a los 7 días — si se necesita algo
verdaderamente permanente (que sobreviva a cerrar la sesión), se necesita
una tarea programada del sistema operativo (Task Scheduler de Windows)
corriendo el script de forma independiente.

### Verificado

Con el archivo real `Stock_06_08.xlsx`: 89 tiendas cubiertas, 0 ausentes,
589 productos actualizados (13 subieron, 56 bajaron, 520 pasaron a
Agotado, 0 recuperados), 65 productos del día ignorados por no estar en el
listado. Confirmado con Playwright que un producto en stock 0 muestra el
badge "Agotado" y botón deshabilitado (escritorio y móvil, 390px) y que un
producto con stock normal no se ve afectado. Conteo total de productos en
el listado se mantuvo en 17.594 antes y después (ningún producto se perdió
por el cruce ni por el fix de stock negativo).

Archivo modificado: solo `portal-web/plantilla_portal.html`.

## 14. Limpieza de código (sin cambios de funcionalidad)

Pasada de limpieza pura sobre `plantilla_portal.html` y `generar_portal.py`,
sin tocar ningún comportamiento visible. Verificado con un análisis
automatizado (todas las clases CSS y funciones JS referenciadas al menos
una vez fuera de su propia definición) y con una regresión completa por
Playwright (filtros, Vista Rápida, panel de detalle, Lista de selección
—agregar, cantidad, exportar a Excel—, badge "Agotado" y tema oscuro): sin
errores de consola, sin diferencias de comportamiento.

- **`nombreDeTienda(tiendaCompleta)`** (nueva función, junto a
  `codigoDeTienda`): la lógica "sacar el nombre de tienda después del
  guion (–)" estaba duplicada de forma independiente 5 veces
  (`tiendaCorta`, `tarjetaDestacadoHtml`, `lineaSeleccionHtml` y dos veces
  en `exportarSeleccionExcel`), con dos implementaciones ligeramente
  distintas (`indexOf`+`slice` vs `includes`+`split`). Se unificó en una
  sola función y las 5 quedaron reducidas a una llamada.
- **`precioActualDeLinea(linea)`** (función eliminada): quedó sin ningún
  llamador tras un refactor anterior; `lineaSeleccionHtml()` ya leía
  `f[VR]` directamente.
- **`generar_portal.py`**: se quitó un `re.sub()` que reescribía el regex
  de `norm()` en el HTML de salida (buscaba reemplazar una clase de
  caracteres Unicode literal por su forma escapada, tipo `\uXXXX-\uYYYY`).
  Comparado contra el archivo real, el patrón ya no coincidía con nada —
  la plantilla ya trae directamente la forma escapada — así que era una
  sustitución inerte desde hace tiempo; quitarla no cambia el HTML
  generado (verificado byte a byte). También se quitaron del diccionario
  `ASSETS_VISUALES` tres entradas (`__ICONO_TEMA__`,
  `__ICONO_AMORTIGUADOR__`, `__ICONO_FILTRO_AIRE__`) cuyos marcadores no
  existen en la plantilla: cada build leía y codificaba en base64 esos
  tres archivos de ícono sin que el resultado se usara en ningún lado.
- Un análisis de todos los selectores CSS del archivo no encontró ninguna
  clase sin uso (el único "hallazgo" fue un falso positivo: la cadena
  "autoplanet.cl" dentro de un comentario).

Confirmado: `index.html` generado después de la limpieza tiene el mismo
conteo de productos (17.594) y el mismo tamaño (1.57 MB) que antes.

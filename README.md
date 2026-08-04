# Portal de Liquidación de Productos en Obsolescencia — Autoplanet

**Sitio publicado:**
- https://niicoeaston.github.io/
- https://niicoeaston.github.io/Obsolescencia_HP/

(mismo sitio, mismos datos — se publican los dos porque así se pidió; hay
que actualizar ambos cuando cambien los datos, ver sección 2).

Página web comercial (no un dashboard) para que supervisores de tienda,
equipos comerciales y clientes de taller consulten qué productos están en
liquidación por obsolescencia y a qué precio. Recorrido: **Zona → Tienda →
Subcategoría → listado de productos**, con descarga del resultado filtrado.

---

## 1. Qué es cada archivo

| Archivo | Para qué sirve |
|---|---|
| **`index.html`** | El sitio publicado. Un solo archivo autónomo (datos incluidos). **No se edita a mano** — se regenera con el script. |
| `plantilla_portal.html` | La plantilla real: HTML + CSS + JS. Aquí se edita el diseño o el comportamiento. |
| `generar_portal.py` | Lee el Excel, arma los datos y produce `index.html` a partir de la plantilla. |
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

Esto lee **`../dashboard-obsolescencia/data/Listado Obsolecencia.xlsx`**
(el mismo archivo que usa el dashboard interno), regenera `index.html` con
los datos nuevos, e imprime en pantalla el reparto de stock valorizado por
zona para que puedas revisarlo antes de publicar.

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

## 3. La columna Zona: historia y estado actual

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

## 4. Arquitectura y por qué

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

## 5. Qué se dejó fuera de la vista principal (a propósito)

Sin tarjetas KPI, sin gráfico de Top 12 subcategorías/marcas, sin columna
de descuento % ni "Total remate" en el listado de productos. El **valor
remate con IVA** (mostrado como "Precio Liquidación" en el selector de
orden) es el precio protagonista, junto al precio normal tachado como
referencia (igual que una vitrina de liquidación real).

El **stock disponible sí se muestra** (columna en la tabla de escritorio,
línea destacada en la tarjeta móvil) — es información operativa
importante para decidir si vale la pena ir a buscar el producto a esa
tienda.

La sección **Análisis** (pestaña aparte, no la vista principal) muestra el
stock valorizado por zona en un donut + tabla de participación. Con 11
zonas reales, el donut solo colorea distinto las 5 más grandes y agrupa el
resto en un segmento gris "Otras zonas" (más de ~5-6 colores en un gráfico
circular deja de ser legible y dos colores empiezan a repetirse). La
**tabla de abajo nunca agrupa**: siempre lista las 11 zonas por separado
con su valor exacto.

---

## 6. Seguridad y qué expone el sitio

El repositorio y el sitio son **públicos** (requisito del hosting gratuito
de GitHub Pages) — indexable por buscadores, sin login. Por diseño, el
listado de productos solo expone: material, texto breve, marca,
subcategoría, stock disponible, precio normal (tachado) y valor remate con
IVA. **No expone** costos, márgenes, contribución ni datos de otras tiendas
fuera de la selección activa. Esto es equivalente a un catálogo de ofertas
público, no información comercial sensible.

---

## 7. Pruebas realizadas

- Recorrido completo Zona → Tienda → Subcategoría → resultados, con datos
  reales (17.594 productos, 88 tiendas, 11 zonas reales, 178
  subcategorías).
- Buscador, orden (incluido ordenar por Stock disponible), "Ver todas las
  subcategorías", descarga CSV.
- Sección Análisis: donut agrupado (top 5 + "Otras zonas") + tabla completa
  sin agrupar, coinciden los números y suman el total nacional.
- Tema claro/oscuro, incluido el donut repintándose con paleta distinta.
- Anchos móviles 360/375/390/412/430px sin desborde horizontal.
- Tablet (768px): tabla simplificada (sin columna Marca, padding reducido)
  para que Stock y Precio se vean completos sin scroll horizontal.
- **Verificado en el sitio publicado real** (los dos links), no solo en
  local.

## 8. Limitaciones conocidas

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

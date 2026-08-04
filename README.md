# Portal de Liquidación de Productos en Obsolescencia — Autoplanet

**Sitio publicado:** https://niicoeaston.github.io/

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
| `zona_tienda.py` | **Mapeo editable** de cada tienda a su zona geográfica real (ver sección 3). |
| `calcular_zona_tienda.py` | Herramienta de diagnóstico para revisar ese mapeo contra la base. No se usa en la generación. |
| `logo_grupo_planet_b64.txt` | El logo de Grupo Planet ya convertido a base64 (se inserta en el footer). |

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

```bash
git add index.html
git commit -m "Actualiza datos del mes"
git push
```

Si tu sesión de Windows ya inició sesión en GitHub una vez (Git Credential
Manager), el `push` no vuelve a pedir credenciales. GitHub Pages tarda
entre 30 segundos y 2 minutos en publicar el cambio. El link **no cambia**.

---

## 3. Por qué existe `zona_tienda.py`

La columna `Zona` del Excel viene **por fila de producto, no por tienda**:
una misma tienda aparece repartida entre varias zonas (ver
`../dashboard-obsolescencia/DIAGNOSTICO_EXCEL.md`, sección 3). Al calcular
cuál es la zona más frecuente por tienda, el resultado no sirve: para casi
todas las tiendas del país —incluidas Temuco, Concepción, Coquimbo u
Osorno— la zona más frecuente en el Excel es "Zona V Region", que
geográficamente no tiene sentido.

Por eso **este portal no usa la columna Zona del Excel**. En su lugar,
`zona_tienda.py` asigna cada tienda a su región real de Chile a partir del
**nombre de la ciudad** (Temuco → Zona Sur, Valparaíso → Zona Valparaíso,
etc.), a mano y de forma revisable.

### Cómo corregir una tienda

Abre `zona_tienda.py` y cambia el valor de su línea, por ejemplo:

```python
"AP0043 – Temuco Caupolican": ZONA_SUR,
```

Los cambios se aplican la próxima vez que corras `generar_portal.py`.

### Simplificación pendiente de confirmar con el negocio

El Excel trae "Zona RM 1", "Zona RM 2" y "Zona RM 3" (probablemente rutas
de reparto dentro del Gran Santiago), pero no hay forma de saber con
certeza qué comuna cae en cada ruta. Por ahora, **todas las tiendas de la
Región Metropolitana se agrupan en una sola "Zona Metropolitana"**. Si
Autoplanet tiene la división oficial de comunas por ruta, se reemplaza
fácilmente en `zona_tienda.py`.

Dos tiendas quedaron marcadas como de **baja confianza** por nombre
ambiguo (`AP0017 – Gabriela`, `AP0070 – Orientales`) — están en
`REVISAR_MANUALMENTE` dentro del mismo archivo.

---

## 4. Arquitectura y por qué

**Un solo archivo HTML estático**, sin backend, sin build, publicado en
**GitHub Pages** (repositorio `niicoeaston.github.io`, gratuito, sin
servidor que mantener).

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
  servidor ni API que mantener. El costo es un archivo de ~1 MB, aceptable
  para este volumen (17.594 productos).

### Identidad visual

Colores muestreados directamente de **autoplanet.cl** (rojo del hero
`#FF061A`, naranja de la barra de categorías `#E65100`) y del **logo real
de Grupo Planet** (rojo `#D7141A`, usado en el footer). El sistema de
diseño se valido primero en Stitch (proyecto "Autoplanet Portal
Comercial") antes de programarlo. El hero es solo tipografía y color, sin
imagen decorativa — decisión explícita para no depender de una ilustración
inventada.

---

## 5. Qué se dejó fuera de la vista principal (a propósito)

Por pedido explícito: sin tarjetas KPI, sin gráfico de Top 12
subcategorías/marcas, sin columna de stock ni de descuento % en el
listado de productos, sin "Total remate". El **valor remate con IVA** es
el único precio protagonista, junto al precio normal tachado como
referencia (igual que una vitrina de liquidación real).

El stock sigue usándose puertas adentro (para filtrar solo productos
disponibles y para calcular el stock valorizado de la sección Análisis),
pero nunca se muestra como columna.

La sección **Análisis** (pestaña aparte, no la vista principal) muestra el
stock valorizado por zona en un donut + tabla de participación.

---

## 6. Seguridad y qué expone el sitio

El repositorio y el sitio son **públicos** (requisito del hosting gratuito
de GitHub Pages) — indexable por buscadores, sin login. Por diseño, el
listado de productos solo expone: material, texto breve, marca,
subcategoría, precio normal (tachado) y valor remate con IVA. **No expone**
costos, márgenes, contribución ni datos de otras tiendas fuera de la
selección activa. Esto es equivalente a un catálogo de ofertas público, no
información comercial sensible.

---

## 7. Pruebas realizadas

- Recorrido completo Zona → Tienda → Subcategoría → resultados, con datos
  reales (17.594 productos, 88 tiendas, 6 zonas geográficas, 178
  subcategorías).
- Buscador, orden, "Ver todas las subcategorías", descarga CSV.
- Sección Análisis: donut + tabla, coinciden los números.
- Tema claro/oscuro.
- Anchos móviles 360/375/390/412/430px sin desborde horizontal (incluida
  una subcategoría con nombre largo que sí desbordaba una tarjeta al
  principio — corregido).
- Tablet (768px): tabla simplificada sin necesitar scroll horizontal para
  ver el precio.
- **Verificado en el sitio publicado real** (no solo en local), incluyendo
  una comprobación explícita de que no fuera un problema de caché del
  navegador.

## 8. Limitaciones conocidas

- La exportación es **solo CSV** (se abre perfecto en Excel en español, con
  `;` y BOM). No se generó un `.xlsx` real para evitar sumar una
  dependencia externa (librería de escritura de Excel) a un sitio que debe
  cargar rápido y funcionar offline una vez cargado.
- El sitio es público; no hay una vista distinta para "cliente" vs
  "supervisor" en esta primera versión — ambos ven exactamente lo mismo
  (que ya excluye todo dato sensible).
- La zona Metropolitana no distingue RM1/RM2/RM3 (ver sección 3).

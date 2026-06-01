# Guía de implementación en Joomla
## Landing Pages — Sistemas y Soluciones Digitales

---

## Archivos entregados

| Archivo | Producto | URL sugerida |
|---|---|---|
| `lp-hr1804uv.html` | HUENU HR1804 UV | `/plotter-uv-hr1804uv` |
| `lp-hf6090.html` | HUENU HF-6090 | `/plotter-uv-cama-plana-hf6090` |
| `lp-teneth7090.html` | Mesa de corte Teneth FC7090 | `/mesa-de-corte-teneth-7090` |
| `lp-km-alquiler.html` | KM Alquiler | `/alquiler-konica-minolta` |

---

## Paso 1 — Elegir el método de publicación

Hay dos formas de publicar estas LPs en Joomla. Elegí la que mejor se adapte a tu instalación.

### Opción A: Artículo con template "en blanco" (recomendada)

Las LPs incluyen su propio header y footer para que funcionen sin el template del sitio.
Con esta opción, el artículo muestra SOLO el HTML de la LP, sin módulos laterales ni navegación del template.

### Opción B: Publicar como archivo PHP/HTML independiente

Subir el HTML como archivo estático dentro de la carpeta de Joomla y enlazarlo directamente.
Más simple pero más difícil de mantener desde el panel de Joomla.

---

## Paso 2 — Opción A: Artículo con template en blanco

### 2.1 Crear el override del template (una sola vez)

**En Joomla 4 (template Cassiopeia):**

1. Ir a **Sistema → Plantillas del sitio → Cassiopeia → Crear overrides**
2. O crear manualmente el archivo:

```
/templates/cassiopeia/html/layouts/joomla/system/message.php
```

La forma más simple es usar el parámetro `?tmpl=component` en la URL:
- URL de la LP: `www.sistemasysoluciones.com/plotter-uv-hr1804uv?tmpl=component`
- Esto elimina todos los módulos y muestra solo el contenido del artículo

**Para una URL limpia sin `?tmpl=component`**, crear un template style dedicado:

1. **Sistema → Estilos de plantillas → Nuevo**
2. Duplicar el estilo actual
3. Asignarle el nombre "Landing Pages - Sin módulos"
4. En la pestaña **Asignación de menú**: asignar solo los ítems de menú de las LPs
5. En la pestaña **Avanzado**: desactivar todos los módulos de posición

### 2.2 Crear el artículo

Para cada landing page:

1. **Contenido → Artículos → Nuevo artículo**
2. **Título:** (puede ser el nombre del producto — se ocultará)
3. **Categoría:** Crear una categoría "Landing Pages" (no visible en el menú principal)
4. En el editor: hacer clic en **"Fuente"** o **"Toggle editor"** (modo HTML)
5. Pegar el contenido completo del archivo `.html`
   - Si usás TinyMCE: puede limpiar algunas etiquetas. Ver punto 2.3.
6. Pestaña **Opciones del artículo**:
   - Mostrar título: **No**
   - Mostrar fecha: **No**
   - Mostrar autor: **No**
   - Mostrar categoría: **No**
7. Pestaña **Metadatos**:
   - Meta descripción: copiar del `<meta name="description">` del HTML
   - Keywords: palabras clave del producto

### 2.3 Evitar que TinyMCE limpie el HTML

TinyMCE elimina etiquetas como `<style>`, `<header>`, `<footer>` y `<section>`.
Dos soluciones:

**Solución 1 (recomendada): Deshabilitar el editor para estas páginas**
1. En el artículo → pestaña **Opciones** → Editor: seleccionar "Sin editor"
2. Pegar el HTML directamente en el textarea

**Solución 2: Configurar TinyMCE para que no limpie**
1. **Sistema → Plugins → Editors - TinyMCE**
2. En "Elementos extendidos válidos" agregar: `style,section,header,footer,nav`
3. En "Modo de verificación de entidades HTML" seleccionar: **Raw**

### 2.4 Crear el ítem de menú (URL amigable)

1. **Menus → Menú principal → Nuevo ítem** (o crear un menú "Hidden" oculto)
2. **Tipo:** Artículos → Artículo único
3. **Artículo:** seleccionar el artículo de la LP
4. **Alias:** `plotter-uv-hr1804uv` (sin espacios, en minúsculas, con guiones)
5. **Estado:** Publicado
6. **Opciones de visualización de página:** Ocultar del menú principal

---

## Paso 3 — Configurar el formulario

Los formularios tienen `action="ACTION_URL_AQUI"`. Reemplazá ese valor con una de estas opciones:

### Opción A: Formspree (más simple, sin backend)

1. Registrarse en https://formspree.io (plan gratuito: 50 envíos/mes)
2. Crear un formulario nuevo → obtener el endpoint (ej: `https://formspree.io/f/xyzabc`)
3. Reemplazar `ACTION_URL_AQUI` por ese endpoint en cada HTML

Agregar en cada formulario:
```html
<input type="hidden" name="_subject" value="Consulta LP - HUENU HR1804 UV">
<input type="hidden" name="_next" value="https://www.sistemasysoluciones.com/gracias">
```

### Opción B: RSForm! Pro (plugin Joomla, recomendado para integración nativa)

1. Instalar RSForm! Pro desde https://www.rsjoomla.com
2. Crear un formulario por cada producto con los mismos campos del HTML
3. Reemplazar el bloque `<form>...</form>` en el HTML por el shortcode de RSForm:
   `{rsform 1}` (donde 1 es el ID del formulario)

### Opción C: Componente de contacto nativo de Joomla

1. **Componentes → Contactos → Nuevo contacto**
2. Crear un contacto por cada producto (ej: "Consulta HR1804 UV")
3. Usar el link del componente de contacto y redirigir desde el botón de la LP

### Opción D: Google Forms embebido

1. Crear un formulario en Google Forms
2. Copiar el `<iframe>` de embed
3. Reemplazar el bloque `<form>` en el HTML por el iframe de Google Forms

---

## Paso 4 — Configuración de Google Analytics / GA4

Agregar el script de Google Analytics antes del `</head>` en cada artículo (o en el template):

```html
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-XXXXXXXXXX"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'G-XXXXXXXXXX');
</script>
```

### Configurar conversiones (eventos)

En cada formulario, agregar este script para registrar la conversión en Google Ads:

```html
<script>
document.querySelector('form').addEventListener('submit', function() {
  gtag('event', 'conversion', {
    'send_to': 'AW-XXXXXXXXXX/XXXXXXXXXX',  // Tu ID de conversión de Google Ads
    'value': 1.0,
    'currency': 'ARS'
  });
});
</script>
```

---

## Paso 5 — Parámetros UTM para las campañas

Usar estas URLs en los anuncios de Google Ads para trackear las conversiones por campaña:

| Producto | URL con UTM |
|---|---|
| HR1804 UV | `/plotter-uv-hr1804uv?utm_source=google&utm_medium=cpc&utm_campaign=huenu-hr1804uv` |
| HF-6090 | `/plotter-uv-cama-plana-hf6090?utm_source=google&utm_medium=cpc&utm_campaign=huenu-hf6090` |
| Teneth 7090 | `/mesa-de-corte-teneth-7090?utm_source=google&utm_medium=cpc&utm_campaign=teneth-fc7090` |
| KM Alquiler | `/alquiler-konica-minolta?utm_source=google&utm_medium=cpc&utm_campaign=km-alquiler` |

---

## Paso 6 — Checklist antes de activar las campañas

- [ ] Los 4 artículos están publicados y tienen URL amigable
- [ ] El formulario envía correctamente (hacer test con datos reales)
- [ ] El formulario llega al email correcto de la empresa
- [ ] El pixel de Google Ads está instalado y verificado
- [ ] Las conversiones están configuradas en Google Ads
- [ ] Las LPs se visualizan correctamente en mobile (probar en celular)
- [ ] Los botones de WhatsApp abren la conversación con el mensaje correcto
- [ ] El número de teléfono del header es clickeable en mobile
- [ ] Los UTM están correctamente configurados en los anuncios
- [ ] Se creó una página de "Gracias" o confirmación de envío

---

## Notas adicionales

**Velocidad de carga:** Las LPs tienen todo el CSS inline, por lo que cargan rápido sin dependencias externas.

**Imágenes:** Los archivos HTML no incluyen imágenes del producto para no asumir rutas. Se recomienda agregar una imagen real del equipo en la sección hero. Agregar dentro de `.hero-content` antes del botón:
```html
<img src="/images/productos/hr1804uv.jpg" alt="Plotter UV HUENU HR1804UV" style="max-width:100%;border-radius:8px;margin-bottom:16px">
```

**Favicon:** El favicon del sitio principal se hereda automáticamente si se usa el template de Joomla.

**Robots:** Asegurarse de que las LPs no estén bloqueadas en robots.txt. Deben ser indexables.

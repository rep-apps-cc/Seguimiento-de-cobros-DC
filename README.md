# Seguimiento de Cobros — NegoFIN S.A.E.C.A.

Aplicación web progresiva (PWA) para generar automáticamente enlaces de WhatsApp
a partir de una planilla Excel con nombres, montos y números de celular.

Todo el procesamiento ocurre **en el celular/navegador del usuario**: el Excel
nunca se sube a ningún servidor.

## Funcionalidades

- Subida de un archivo `.xlsx`, `.xls` o `.csv` (botón o arrastrar y soltar).
- Detección automática de las columnas de la planilla (nombre, cédula, montos,
  celular, etc.) — no importa si el Excel cambia de estructura.
- Si no puede identificar sola la columna del número de celular, te la pregunta.
- Genera un mensaje de cobranza por defecto y permite editarlo libremente,
  insertando cualquier columna de la planilla como campo (`{{CAMPO}}`).
- Normaliza los números de celular al formato internacional de Paraguay (595).
- Crea un enlace `https://wa.me/...` por cada contacto válido, con el mensaje
  ya redactado, listo para tocar y enviar.
- Marca los contactos ya contactados (se guarda en el propio celular).
- Permite descargar de nuevo el Excel, con una columna adicional de enlaces.
- Se puede instalar como app en Android/iOS y abre en pantalla completa.

## Cómo publicarla en GitHub Pages (gratis, sin servidor)

1. Creá un repositorio nuevo en GitHub, por ejemplo `seguimiento-de-cobros`.
2. Subí **todo el contenido de esta carpeta** (`index.html`, `manifest.json`,
   `service-worker.js`, `css/`, `js/`, `icons/`, `favicon.ico`) a la raíz del
   repositorio.
3. En el repositorio: **Settings → Pages**.
4. En "Source" elegí la rama `main` y la carpeta `/ (root)`. Guardá.
5. GitHub te va a dar una URL, algo como:
   `https://tu-usuario.github.io/seguimiento-de-cobros/`
6. Abrí esa URL en el celular con Chrome (Android).

> Importante: la PWA necesita HTTPS para poder instalarse. GitHub Pages ya
> entrega el sitio con HTTPS automáticamente, así que no necesitás configurar
> nada extra.

## Cómo instalarla en Android (pantalla completa)

1. Abrí la URL de GitHub Pages en Chrome.
2. Tocá el menú (⋮) → **"Instalar app"** o **"Agregar a pantalla de inicio"**.
   Chrome también puede mostrar un banner automático para instalarla.
3. Al abrirla desde el ícono instalado, se abre **sin la barra de Chrome**,
   en pantalla completa.

## Cómo usarla

1. Tocá "Cargar planilla" y elegí el Excel.
2. Si la app no reconoce sola la columna del celular, seleccionala de la lista.
3. Revisá el mensaje sugerido (tocá "✏️ Mensaje de WhatsApp" para editarlo).
   Podés tocar cualquier "chip" con el nombre de una columna para insertarla
   en el mensaje, por ejemplo `{{NOMBRE}}`, `{{CEDULA}}`, `{{MONTO}}`.
4. Tocá "⚡ Generar enlaces".
5. En la lista de contactos, tocá el botón verde de WhatsApp junto a cada
   persona para abrir la conversación con el mensaje ya escrito.
6. Opcional: tocá "⬇ Descargar Excel" para guardar una copia de la planilla
   con la columna de enlaces agregada.

## Estructura de archivos

```
├── index.html            Pantalla principal
├── manifest.json         Configuración de la PWA (nombre, ícono, pantalla completa)
├── service-worker.js     Cacheo para poder instalar la app
├── favicon.ico
├── css/
│   └── styles.css
├── js/
│   └── app.js            Toda la lógica: lectura de Excel, plantilla, enlaces
└── icons/
    ├── icon-192.png
    ├── icon-512.png
    ├── icon-maskable-192.png
    ├── icon-maskable-512.png
    ├── apple-touch-icon.png
    └── logo-negofin.png  Logo original de la empresa (referencia)
```

## Personalizar el isologo o los colores

Los colores principales están definidos como variables CSS al inicio de
`css/styles.css`:

```css
--bright: #00C080;  /* verde brillante del logo */
--deep:   #006040;  /* verde profundo del logo */
```

Los íconos (`icons/icon-*.png`) fueron generados a partir de un isologo
simplificado (la persona + la flecha ascendente del logo de NegoFIN). Si
querés reemplazarlos por tu propio diseño, generá los mismos tamaños
(192x192, 512x512, y versión "maskable" con margen de seguridad) y
reemplazá los archivos manteniendo los mismos nombres.

## Notas técnicas

- La lectura y escritura de Excel usa la librería [SheetJS](https://sheetjs.com/)
  cargada desde CDN (`xlsx.full.min.js`). El service worker la cachea para
  que la app siga abriendo sin conexión después del primer uso.
- No hay backend ni base de datos: todo vive en el navegador. El progreso de
  "enviados" se guarda con `localStorage`, ligado al nombre y tamaño del
  archivo cargado.

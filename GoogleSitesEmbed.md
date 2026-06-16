# Código de Integración para Google Sites

Esta guía te permite integrar la plataforma de la **U. E. Santo Domingo Sabio** en tu sitio web de **Google Sites**, asegurando que sea responsivo, que no se corte en teléfonos celulares y que cargue de forma ultra-rápida.

Existen dos métodos recomendados para realizar la inserción:

---

## Método 1: Insertar mediante Código HTML (Recomendado)

Este método añade un contenedor inteligente que adapta de forma automática el alto y el ancho de la aplicación, evitando las barras de desplazamiento molestas en dispositivos móviles.

### Pasos para insertarlo en Google Sites:
1. Ingresa al editor de tu **Google Site**.
2. En la barra lateral derecha, selecciona la pestaña **Insertar** y haz clic en **Insertar (< >)**.
3. Elige la pestaña **Insertar código**.
4. Copia y pega el siguiente código HTML, asegurándote de reemplazar la frase `TU_URL_DE_APPS_SCRIPT_AQUÍ` por el enlace oficial de tu Aplicación Web de Google Apps Script (que termina en `/exec`):

```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SADOSA 2026 - Google Sites Embed</title>
  <style>
    /* Reset de estilos para ajuste perfecto */
    html, body {
      margin: 0;
      padding: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      background-color: #f8fafc;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    /* Contenedor adaptativo */
    .sadosa-container {
      position: relative;
      width: 100%;
      height: 100%;
      display: flex;
      flex-direction: column;
    }

    /* Iframe con bordes redondeados y sombra sutil */
    .sadosa-iframe {
      flex: 1;
      width: 100%;
      border: none;
      background: transparent;
      transition: opacity 0.3s ease;
    }

    /* Barra informativa inferior */
    .sadosa-footer {
      background-color: #1e293b;
      color: #94a3b8;
      font-size: 11px;
      padding: 8px 16px;
      text-align: center;
      font-weight: 500;
      letter-spacing: 0.5px;
      border-top: 1px solid #334155;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 6px;
    }

    .sadosa-footer-accent {
      color: #f59e0b;
      font-weight: 700;
    }

    .sadosa-btn-external {
      background-color: #334155;
      color: #ffffff;
      text-decoration: none;
      padding: 3px 8px;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 600;
      transition: background-color 0.2s;
    }

    .sadosa-btn-external:hover {
      background-color: #475569;
    }

    /* Animación de carga */
    .sadosa-loader {
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      font-size: 13px;
      color: #64748b;
      display: flex;
      align-items: center;
      gap: 10px;
      pointer-events: none;
    }

    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #2563eb;
      border-radius: 50%;
      width: 18px;
      height: 18px;
      animation: spin 1s linear infinite;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>

  <div class="sadosa-container">
    <!-- Cargador visual rápido -->
    <div id="loader" class="sadosa-loader">
      <div class="spinner"></div>
      <span>Cargando planificador escolar...</span>
    </div>

    <!-- Iframe Seguro conectado a tu Google Sheet (Reemplaza con tu URL) -->
    <iframe 
      id="sadosaFrame"
      class="sadosa-iframe"
      src="TU_URL_DE_APPS_SCRIPT_AQUÍ" 
      allow="clipboard-write"
      sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-modals"
      onload="hideLoader()">
    </iframe>

    <!-- Pie de página integrado con botón de auxilio -->
    <div class="sadosa-footer">
      <div>🏫 <span class="sadosa-footer-accent">SADOSA 2026</span> • Primaria Santo Domingo Sabio</div>
      <div>
        <a id="extLink" href="TU_URL_DE_APPS_SCRIPT_AQUÍ" target="_blank" class="sadosa-btn-external">
          ↗️ Abrir en pantalla completa
        </a>
      </div>
    </div>
  </div>

  <script>
    function hideLoader() {
      const loader = document.getElementById('loader');
      if (loader) {
        loader.style.opacity = '0';
        setTimeout(() => loader.style.display = 'none', 300);
      }
    }

    // Configuración automática por si se olvida cambiar la URL en el link de enlace externo de abajo
    const appScriptUrl = "TU_URL_DE_APPS_SCRIPT_AQUÍ";
    if (appScriptUrl !== "TU_URL_DE_APPS_SCRIPT_AQUÍ") {
      document.getElementById('extLink').href = appScriptUrl;
    }
  </script>

</body>
</html>
```

5. Presiona **Siguiente** y luego **Insertar**. ¡Listo! Acomoda el elemento arrastrándolo para darle un buen espacio vertical en el lienzo de Google Sites.

---

## Método 2: Insertar mediante URL Directa

Este método es el más simple porque Google Sites gestiona toda la inserción, pero en ocasiones puede mostrar bordes grises o limitaciones de escalado en tabletas.

### Pasos:
1. En tu **Google Site**, haz clic en **Insertar (< >)**.
2. Selecciona la pestaña que dice **Por URL**.
3. Pega directamente tu enlace de Aplicación Web de Google Apps Script (que termina en `/exec`).
4. Google Sites generará una vista previa interactiva. Elige la opción **Página completa** en lugar de miniatura y presiona **Insertar**.
5. Amplía el recuadro verticalmente tirando de los círculos de redimensión del editor para que el planificador se vea completamente cómodo.

---

## 📌 Recomendaciones de Oro para Google Sites

1. **Permisos en la Hoja de Cálculo**: Para que los Padres de Familia puedan ingresar a ver el calendario, asegúrate de que el documento de Google Sheets tenga permisos de compartición configurados en **"Cualquier persona con el enlace puede leer/ver"**.
2. **Quién tiene Acceso a la App Web**: Cuando realices el despliegue de tu Aplicación Web de Apps Script (Nueva implementación), recuerda colocar en la casilla **Quién tiene acceso**: **"Cualquier persona" (Anyone)**. De lo contrario, Google Sites pedirá obligatoriamente iniciar sesión con cuenta de Google propia a cada padre que intente mirar el calendario.
3. **Guardado del ID de Google Sheets**: Si los padres de familia ven un cartel amarillo con un aviso de "error de sincronización" cuando entren desde Google Sites, abre tu archivo de Google Sheets y presiona en el menú de arriba: **SADOSA 2026 > Abrir Panel de Maestros**. Esto realiza un enlace manual del ID del archivo solucionándolo al instante.

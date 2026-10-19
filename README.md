# Add-in PFI

Skeleton inicial del add-in de Outlook para el proyecto de ciberdefensa con XAI.

## Requisitos
- Node.js 18+
- Outlook Desktop o Outlook Web (con sideload habilitado)

## Instalacion
1. npm install

## Ejecutar
- npm run start

## Backend y veredicto

El navegador llama a `/api/v1/analyze` en el mismo dominio del complemento. La función Node.js `api/v1/analyze.mjs` en Vercel reenvía el correo al backend AWS y agrega `X-API-Key` del lado del servidor. La clave y la dirección de AWS no se incorporan al bundle del navegador. El código del clasificador no cambia.

Antes de publicar, configurar en Vercel (Production y Preview, según corresponda):

- `API_BASE_URL`: origen HTTPS del backend, sin ruta, query ni credenciales; por ejemplo `https://api.example.com`.
- `API_KEY`: clave que acepta el backend. No hay clave de respaldo en la función publicada; si falta configuración, responde 503.

Volver a desplegar después de configurar las variables. Para revocar una clave incluida en versiones anteriores del navegador, coordinar su reemplazo en la variable `API_KEY` de AWS y en Vercel, y actualizar ambos servicios. Cambiar solo el cliente no revoca esa clave.

La función acepta JSON del mismo origen, limita la solicitud a 1.100.000 bytes y espera hasta 25 segundos por AWS (duración máxima configurada en Vercel: 30 segundos). No registra cuerpos ni claves ni reenvía errores internos. El correo pasa por Vercel antes de llegar a AWS; se debe reflejar este recorrido en la documentación de privacidad. El control de origen no autentica usuarios ni impide clientes externos que falsifiquen ese encabezado. Esta versión no incorpora SSO, tokens de un solo uso ni DynamoDB. Configurar una regla de límite de solicitudes para `POST /api/v1/analyze` en Vercel Firewall antes de abrir el piloto a terceros; el límite debe probarse con la concurrencia esperada, y no se implementa como contador en memoria de una función aislada.

En desarrollo, webpack reenvía `/api` al backend local `http://127.0.0.1:8001`, o a `API_BASE_URL` si se define, y agrega `API_KEY` del lado del servidor de desarrollo. Solo para ese entorno local se admite la clave predeterminada de desarrollo. Configurar las variables en la terminal antes de ejecutar `npm start`.

El panel respeta `is_phishing` enviado por el servidor y muestra `content_signals` aunque el SLM no devuelva JSON. No vuelve a clasificar un mensaje por superar 50% de score.

Verificación local:

```powershell
node --test scripts/test-analysis.cjs
node --test scripts/test-proxy.mjs
npm run build
```

## Notas
- El panel se carga en lectura de mensajes (Message Read).
- El manifest principal esta en manifest.xml.
 - Sideload manual: en Outlook, agrega manifest.xml desde Complementos.
 - Si tu entorno no confia en el certificado local, acepta el certificado del dev server.

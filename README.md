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

`API_BASE_URL` se incorpora durante la compilación. En desarrollo, si no se configura, el panel llama a `/api` en el servidor local HTTPS; el proxy de webpack reenvía internamente esas solicitudes al backend local. Para compilar la versión publicada, `API_BASE_URL` es obligatoria y debe apuntar a la URL HTTPS real del backend. Cuando el dominio y su certificado estén configurados, definir esa variable y volver a compilar/publicar.

El panel respeta `is_phishing` enviado por el servidor y muestra `content_signals` aunque el SLM no devuelva JSON. No vuelve a clasificar un mensaje por superar 50% de score.

Verificación local:

```powershell
node --test scripts/test-analysis.cjs
npm run build
```

## Notas
- El panel se carga en lectura de mensajes (Message Read).
- El manifest principal esta en manifest.xml.
 - Sideload manual: en Outlook, agrega manifest.xml desde Complementos.
 - Si tu entorno no confia en el certificado local, acepta el certificado del dev server.

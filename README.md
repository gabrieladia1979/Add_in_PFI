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

`API_BASE_URL` se incorpora durante la compilación. Si no se configura, se usa `http://127.0.0.1:8001` para desarrollo. Para el complemento publicado, configurar la URL HTTPS del backend actualizado en el entorno de compilación y volver a compilar/publicar. Ya no se utiliza un túnel temporal fijo dentro del código.

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

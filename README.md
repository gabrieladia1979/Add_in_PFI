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

Esta rama `codex/hybrid-xai-addin` llama a `POST /api/v1/analyze/hybrid` del backend `flujo_v2`, rama `codex/hybrid-xai`. No cambia el add-in de `main`. El backend debe tener `PHISHARG_HYBRID_MODEL_DIR` apuntando a un artefacto Hybrid entrenado y las dependencias de `requirements-hybrid.txt`. Si la ruta o el artefacto faltan, el panel muestra un error explícito y no usa el clasificador anterior en silencio.

`API_BASE_URL` se incorpora durante la compilación. En desarrollo, si no se configura, el panel llama a `/api` en el servidor local HTTPS; el proxy de webpack reenvía internamente esas solicitudes al backend local en `http://127.0.0.1:8001`. Para compilar una versión publicada, `API_BASE_URL` es obligatoria y debe apuntar a una URL HTTPS del backend que tenga desplegada la ruta Hybrid. La instancia anterior de AWS puede no tenerla. La clave `API_KEY` también se incorpora al JavaScript del add-in; la clave por defecto es solo para pruebas locales y no constituye un mecanismo de autenticación secreto en una publicación.

El panel respeta `is_phishing` enviado por el servidor y muestra `content_signals` aunque la explicación no devuelva JSON. No vuelve a clasificar un mensaje por superar 50/100 de índice de riesgo. Lee `hybrid_xai`: presenta variables técnicas atribuidas al modelo y palabras LIME únicamente cuando el backend las marca confiables y el veredicto bruto coincide con el final. Si LIME no supera la fidelidad, muestra esa limitación sin inventar palabras influyentes. El cálculo de LIME puede aumentar la latencia de la respuesta.

Verificación local:

```powershell
node --test scripts/test-analysis.cjs
npm run build:dev
# Para probar la compilación de producción, configurar API_BASE_URL con HTTPS.
```

Con el backend Hybrid ejecutándose localmente en el puerto 8001 y su artefacto configurado, `node scripts/check-hybrid-backend.cjs` realiza una llamada HTTP real y comprueba el contrato que usa el add-in. Si el backend está en otro puerto, definir `ANALYSIS_BACKEND_URL`; si usa otra clave, definir `API_KEY`. La prueba imprime solo el veredicto y estados de XAI, no el contenido del correo.

## Notas
- El panel se carga en lectura de mensajes (Message Read).
- El manifest principal esta en manifest.xml.
 - Sideload manual: en Outlook, agrega manifest.xml desde Complementos.
 - Si tu entorno no confia en el certificado local, acepta el certificado del dev server.

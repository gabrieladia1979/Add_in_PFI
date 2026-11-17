# PhishARG Hybrid Piloto — Add-in Outlook

Complemento de Outlook para **PhishARG Hybrid Piloto**, diseñado para interactuar directamente con el servicio de inferencia híbrida y explicabilidad XAI en AWS ECS Fargate.

Para una descripción detallada de la arquitectura en dos fases, el protocolo de autenticación con tokens HMAC y las configuraciones de entorno, consultar [docs/ARQUITECTURA_PILOTO_HIBRIDO.md](docs/ARQUITECTURA_PILOTO_HIBRIDO.md).

## Requisitos
- Node.js 18+
- Outlook Desktop o Outlook Web (con sideload habilitado)

## Instalacion
1. npm install

## Ejecutar
- npm run start

## Backend y veredicto

Esta rama `hybrid-aws-pilot` llama a `POST /api/v1/analyze/hybrid` del backend `flujo_v2`, rama `hybrid-xai-aws`. Su manifiesto registra **PhishARG Hybrid Piloto** con un ID propio para instalarlo junto al complemento de `main`. El manifiesto publicado apunta al dominio previsto `add-in-pfi-hybrid.vercel.app`. Confirmá el dominio asignado al proyecto separado de Vercel y actualizá el manifiesto si corresponde, antes de instalarlo.

En el proyecto separado de Vercel, configurá `API_BASE_URL` con el origen HTTPS del servicio Hybrid, `PHISHARG_TOKEN_KEY` con la clave de firma compartida con ese servicio y `PHISHARG_RATE_LIMIT_ID` con la regla de control de tasa de ese proyecto. La función del servidor emite los permisos de análisis y conserva la clave de firma. En ECS, configurá `PHISHARG_CORS_ORIGINS` con el origen del complemento piloto. En desarrollo, webpack reenvía `/api` al backend local en `http://127.0.0.1:8001`.

El panel respeta `is_phishing` enviado por el servidor y muestra `content_signals` aunque la explicación no devuelva JSON. No vuelve a clasificar un mensaje por superar 50/100 de índice de riesgo. Lee `hybrid_xai`: presenta variables técnicas atribuidas al modelo y palabras LIME únicamente cuando el backend las marca confiables y el veredicto bruto coincide con el final. Si LIME no supera la fidelidad, muestra esa limitación sin inventar palabras influyentes. El cálculo de LIME puede aumentar la latencia de la respuesta.

Verificación local:

```powershell
node --test scripts/test-analysis.cjs
node --test scripts/test-api-token.cjs
npm run build:dev
```

Con el backend Hybrid ejecutándose localmente en el puerto 8001 y su artefacto configurado, `node scripts/check-hybrid-backend.cjs` realiza una llamada HTTP real y comprueba el contrato que usa el add-in. Si el backend está en otro puerto, definir `ANALYSIS_BACKEND_URL`; si usa otra clave, definir `API_KEY`. La prueba imprime solo el veredicto y estados de XAI, no el contenido del correo.

## Notas
- El panel se carga en lectura de mensajes (Message Read).
- El manifest principal esta en manifest.xml.
 - Sideload manual: en Outlook, agrega manifest.xml desde Complementos.
 - Si tu entorno no confia en el certificado local, acepta el certificado del dev server.

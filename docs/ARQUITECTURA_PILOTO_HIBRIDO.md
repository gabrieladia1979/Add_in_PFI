# Arquitectura y Guía Técnica — PhishARG Hybrid Piloto (Add-in)

Este documento detalla la arquitectura, el modelo de seguridad criptográfica y las pautas operativas del complemento **PhishARG Hybrid Piloto** (`entrega` y `hybrid-aws-pilot`), diseñado para conectarse directamente con el backend de inferencia híbrida y explicabilidad XAI en AWS ECS Fargate.

---

## 1. Identidad Independiente y Coexistencia

Este complemento cuenta con una identidad propia en su manifiesto de Office.js.

* **ID del Complemento:** `f0355dbe-00bc-47de-8626-f157b7c59135`
* **Nombre Visible:** `PhishARG Hybrid Piloto`
* **Dominio Previsto:** `https://add-in-pfi-hybrid.vercel.app`
* **Beneficio Operativo:** Permite la instalación paralela en el mismo cliente de Outlook junto al add-in principal sin colisión de identificadores ni sobreescritura de configuraciones.

---

## 2. Flujo de Datos y Privacidad Criptográfica

En el despliegue piloto, Vercel emite el permiso de análisis y AWS recibe el correo.

```
[ Outlook Desktop / Web ]
       │
       ├──── 1. Solicita Token con Digest (Sin cuerpo de correo) ───► [ Vercel Serverless ]
       │        POST /api/v1/analyze/token { nonce, digest }     api/v1/analyze/token.mjs
       │        ◄─── 2. Retorna Token HMAC firmado y efímero ──────────────┘
       │
       └──── 3. Envío directo del correo con Token HMAC ────────────► [ AWS ECS Fargate ]
                POST /api/v1/analyze/hybrid (HTTPS)          flujo_v2 (hybrid-xai-aws)
                - Fase 1: Veredicto rápido inmediato
                - Fase 2: Explicabilidad enriquecida (SHAP + LIME)
```

### Autorización y análisis
1. El complemento envía a Vercel un nonce y un resumen SHA-256 del nonce y el cuerpo JSON. Para cada fase, el complemento solicita un permiso nuevo y envía el cuerpo directamente al origen HTTPS del piloto.
2. **Tokens de un solo uso:** La clave secreta de firma (`PHISHARG_TOKEN_KEY`) permanece aislada en el servidor y nunca se incluye en el código empaquetado del navegador.
3. **Análisis en dos fases:** Permite mostrar el veredicto temprano y actualizar el panel cuando llega la explicación.

---

## 3. Protocolo de Pruebas y Verificación

Las pruebas locales comprueban el contrato Hybrid y la emisión de permisos.

```powershell
npm ci
npm test
npm run build
```

---

## 4. Configuración de Entorno

* **Vercel (Proyecto Piloto):**
  * `API_BASE_URL`: Origen HTTPS del servicio AWS asignado al piloto.
  * `PHISHARG_TOKEN_KEY`: Secreto criptográfico compartido con el servicio en ECS.
  * `PHISHARG_RATE_LIMIT_ID`: Identificador de regla para mitigación de abusos.
* **AWS ECS (Servicio Híbrido):**
  * `PHISHARG_CORS_ORIGINS`: `https://add-in-pfi-hybrid.vercel.app`
  * `TOKEN_KEY_SECRET_ARN`: ARN del secreto en AWS Secrets Manager para validación de tokens.

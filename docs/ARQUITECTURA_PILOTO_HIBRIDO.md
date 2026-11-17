# Arquitectura y Guía Técnica — PhishARG Hybrid Piloto (Add-in)

Este documento detalla la arquitectura, el modelo de seguridad criptográfica y las pautas operativas del complemento **PhishARG Hybrid Piloto** (`hybrid-aws-pilot`), diseñado para conectarse directamente con el backend de inferencia híbrida y explicabilidad XAI en AWS ECS Fargate.

---

## 1. Identidad Independiente y Coexistencia

A diferencia del add-in base de entrega (que opera a través de una función proxy monolítica en Vercel), este complemento piloto cuenta con una identidad propia en su manifiesto de Office.js:

* **ID del Complemento:** `f0355dbe-00bc-47de-8626-f157b7c59135`
* **Nombre Visible:** `PhishARG Hybrid Piloto`
* **Dominio Previsto:** `https://add-in-pfi-hybrid.vercel.app`
* **Beneficio Operativo:** Permite la instalación paralela en el mismo cliente de Outlook junto al add-in principal sin colisión de identificadores ni sobreescritura de configuraciones.

---

## 2. Flujo de Datos y Privacidad Criptográfica

Para optimizar el rendimiento y garantizar la máxima privacidad del usuario corporativo, este complemento implementa un protocolo de autorización en dos etapas:

```
[ Outlook Desktop / Web ]
       │
       ├──── 1. Solicita Token con Digest (Sin cuerpo de correo) ───► [ Vercel Serverless ]
       │        GET /api/v1/analyze/token?digest=SHA256(correo + salt)     api/v1/analyze/token.mjs
       │        ◄─── 2. Retorna Token HMAC firmado y efímero ──────────────┘
       │
       └──── 3. Envío directo del correo con Token HMAC ────────────► [ AWS ECS Fargate ]
                POST /api/v1/analyze/hybrid (HTTPS / ACM TLS 1.3)          flujo_v2 (hybrid-xai-aws)
                - Fase 1: Veredicto rápido inmediato
                - Fase 2: Explicabilidad enriquecida (SHAP + LIME)
```

### Garantías de Privacidad y Rendimiento:
1. **Cero texto en Vercel:** El contenido del correo electrónico **nunca pasa por la infraestructura de Vercel**; únicamente se transmite un hash SHA-256 salado para validar el origen y la cuota de peticiones.
2. **Tokens de un solo uso:** La clave secreta de firma (`PHISHARG_TOKEN_KEY`) permanece aislada en el servidor y nunca se incluye en el código empaquetado del navegador.
3. **Análisis en dos fases:** Permite devolver al usuario una alerta visual instantánea sobre el riesgo del mensaje mientras se calculan las atribuciones matemáticas complejas en segundo plano.

---

## 3. Protocolo de Pruebas y Verificación

El complemento cuenta con una suite completa de pruebas unitarias y de integración que validan contratos, manejo de tokens y aislamiento:

```powershell
# 1. Pruebas de contrato y veredicto híbrido (11 pruebas)
node --test scripts/test-analysis.cjs

# 2. Pruebas de emisión y seguridad de tokens HMAC (4 pruebas)
node --test scripts/test-api-token.cjs

# 3. Compilación de producción
npm run build
```

---

## 4. Configuración de Entorno

* **Vercel (Proyecto Piloto):**
  * `API_BASE_URL`: Origen HTTPS de la API en AWS (`https://api.phisharg.com`).
  * `PHISHARG_TOKEN_KEY`: Secreto criptográfico compartido con el servicio en ECS.
  * `PHISHARG_RATE_LIMIT_ID`: Identificador de regla para mitigación de abusos.
* **AWS ECS (Servicio Híbrido):**
  * `PHISHARG_CORS_ORIGINS`: `https://add-in-pfi-hybrid.vercel.app`
  * `TOKEN_KEY_SECRET_ARN`: ARN del secreto en AWS Secrets Manager para validación de tokens.

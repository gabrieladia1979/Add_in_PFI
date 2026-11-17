# PhishARG — Complemento de Ciberdefensa para Microsoft Outlook

**Proyecto Final de Integración (PFI) — Ingeniería en Informática — UADE**  
**Complemento Oficial de Lectura de Correo (Office.js / React / TypeScript)**

---

## 1. Descripción General

**PhishARG** es una solución de ciberdefensa proactiva contra ingeniería social y ataques de phishing por correo electrónico. Este repositorio contiene el código fuente del complemento de lectura para **Microsoft Outlook** (escritorio y web), diseñado para conectarse de forma segura y directa con el clúster de inferencia híbrida y explicabilidad (XAI) alojado en **AWS ECS Fargate**.

La rama canónica de evaluación y entrega académica es **`entrega`**, homologada con el backend de inferencia en `flujo_v2/entrega`.

---

## 2. Arquitectura de Inferencia y Seguridad

El complemento implementa una estrategia de análisis en dos fases con aislamiento riguroso de privacidad corporativa:

```
[ Outlook Desktop / Web ]
         │
         ├─── (1) POST /api/v1/analyze/token ───► [ Vercel Serverless Function ]
         │    Envía únicamente SHA-256(nonce + correo)     - Valida origen y límite de tasa
         │    (El cuerpo del mensaje NUNCA va a Vercel)    - Firma token efímero HMAC SHA-256
         │                                                 - Devuelve token + api_base_url
         ▼
         └─── (2) POST /api/v1/analyze/hybrid ──► [ AWS ECS Fargate (TLS 1.3) ]
              Directo con cabecera Bearer <token>          - Inferencia MiniLM + XGBoost V4
              - Fase 1: quick=true (Veredicto rápido)      - TreeSHAP (Atribución log-odds)
              - Fase 2: Enriquecimiento XAI progresivo     - LIME acotado en texto
                                                           - Explicación SLM Qwen 2.5 3B
```

### Pilares de Seguridad:
* **Privacidad de Extremo a Extremo:** El contenido del correo viaja exclusivamente desde el cliente hacia AWS por conexión cifrada **TLS 1.3**. Vercel actúa únicamente como autoridad emisora de autorizaciones sin acceder al texto.
* **Tokens Efímeros HMAC:** Autorización criptográfica de 90 segundos atada al hash del correo y a un *nonce* único, previniendo ataques de repetición (*replay attacks*).
* **Análisis en Dos Fases (*Two-Phase*):** Presentación inmediata del veredicto mientras se computan en paralelo las perturbaciones de LIME y las explicaciones didácticas de Qwen 2.5 3B.

---

## 3. Requisitos del Sistema

* **Node.js:** Versión 18 o superior.
* **Gestor de paquetes:** `npm` 9+.
* **Cliente de Correo:** Microsoft Outlook para Windows/macOS o Outlook en la Web (con permiso de carga lateral / *sideloading* habilitado).

---

## 4. Instalación y Puesta en Marcha

### Instalación de dependencias:
```powershell
npm ci
```

### Entorno de desarrollo local:
```powershell
npm run start
```
*Inicia el servidor Webpack en `https://localhost:3000` con recarga en caliente (*hot-reload*).*

### Compilación para producción:
```powershell
npm run build
```
*Genera el bundle optimizado y minificado en la carpeta `dist/`.*

---

## 5. Configuración de Variables de Entorno (Vercel)

Para el despliegue del emisor de tokens en Vercel, configurar las siguientes variables de entorno:

| Variable | Descripción | Valor de Ejemplo |
| :--- | :--- | :--- |
| `API_BASE_URL` | Origen HTTPS del backend en AWS ALB (sin rutas ni credenciales) | `https://api.phisharg.com` |
| `PHISHARG_TOKEN_KEY` | Clave secreta compartida para la firma criptográfica HMAC (mínimo 32 caracteres) | `[Clave secreta configurada en ECS]` |
| `PHISHARG_RATE_LIMIT_ID` | Identificador de la regla de control de tasa en Vercel Firewall | `phisharg_rate_limit_rule` |

---

## 6. Suite de Pruebas Automatizadas

El proyecto cuenta con una batería de 15 pruebas unitarias y de integración que validan contratos de red, tokens y parsing de señales:

```powershell
npm test
```

### Puntos auditados por la suite (15 / 15 ✅):
1. Cumplimiento del contrato de veredicto predictivo devuelto por el backend.
2. Preservación del veredicto ante señales de contenido crítico y bypass seguro.
3. Despacho rápido en dos fases (`quick=true`).
4. Rechazo estricto de backends sin contrato XAI.
5. Emisión, validación y expiración de tokens HMAC efímeros.
6. Aislamiento de privacidad: Vercel solo recibe el digest criptográfico salado.
7. Límite de carga útil de 1.1 MB y cancelación de solicitudes lentas.

---

## 7. Despliegue e Instalación en Outlook

* **Producción:** Instalar [`manifest.xml`](manifest.xml), el cual apunta al despliegue seguro en Vercel (`https://add-in-pfi-hybrid.vercel.app`).
* **Desarrollo Local:** Cargar [`manifest-local.xml`](manifest-local.xml) desde el menú de complementos de Outlook apuntando a `https://localhost:3000`.

---

## 8. Documentación Técnica de Referencia

* [**Arquitectura del Piloto Híbrido**](docs/ARQUITECTURA_PILOTO_HIBRIDO.md): Detalle técnico del protocolo HMAC, orígenes cruzados (CORS) y flujo de dos fases.
* [**Memoria de Saneamiento y Entrega**](docs/RESUMEN_SANEAMIENTO_Y_ENTREGA.md): Registro histórico de saneamiento de ramas, descarte de artefactos y localización al español.

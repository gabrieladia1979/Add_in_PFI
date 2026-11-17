# Resumen de Saneamiento, Organización y Entrega — PhishARG Add-in

**Proyecto Final de Integración (PFI) — Ingeniería en Informática — UADE**  
**Fecha de consolidación:** 17 de noviembre de 2026  
**Rama oficial:** `entrega` (Add-in base) / `hybrid-aws-pilot` (Piloto Híbrido)

---

## 1. Alcance y Objetivos de la Intervención

El presente informe documenta las acciones de reestructuración, saneamiento técnico, localización al español y verificación de calidad ejecutadas sobre los componentes cliente (Outlook Add-in) del ecosistema **PhishARG**.

El propósito fundamental fue adecuar la totalidad del código, metadatos, manifiestos y suites de pruebas para su entrega académica formal ante el tribunal de evaluación de la carrera, asegurando la eliminación de cualquier rastro o metadato generado por herramientas de asistencia automática y garantizando una experiencia de usuario consistente, profesional y 100% en español.

---

## 2. Organización Arquitectónica y Convivencia de Add-ins

Para satisfacer tanto la demostración estándar como la evaluación de seguridad avanzada, el frontend se organizó en dos variantes totalmente aisladas y compatibles:

### A. Add-in Base de Producción (`Add_in_PFI` — Rama `entrega`)
* **Rol:** Complemento de uso general para lectura de correos en Outlook Desktop y Web.
* **Canalización segura:** Se conecta a través de una función serverless intermedia en Vercel (`api/v1/analyze.mjs`), la cual inyecta la credencial `X-API-Key` del lado del servidor y previene la exposición de endpoints o secretos en el navegador.
* **Identidad en Outlook:** GUID `c1b3a1e7-6cbb-4b43-9ed9-9876ef0938f1`.
* **Nombre visible:** `Phish-Score Lab`.

### B. Piloto Avanzado Híbrido (`Add_in_PFI_hybrid_pilot` — Rama `hybrid-aws-pilot`)
* **Rol:** Complemento especializado para demostración de ciberdefensa de baja latencia con explicabilidad detallada.
* **Canalización directa por TLS 1.3:** El cuerpo del mensaje viaja directamente desde el cliente hacia AWS sin intermediación de Vercel. Vercel actúa exclusivamente como emisor de tokens efímeros firmados por HMAC SHA-256 contra el hash criptográfico del correo.
* **Inferencia en dos fases:** Veredicto inmediato de clasificación seguido de enriquecimiento progresivo con explicabilidad TreeSHAP (atribuciones por variable en log-odds) y LIME (análisis local de palabras en el cuerpo).
* **Identidad en Outlook:** GUID `f0355dbe-00bc-47de-8626-f157b7c59135`.
* **Nombre visible:** `PhishARG Hybrid Piloto`.

> [!NOTE]
> Gracias a poseer identificadores (`Id`) y URLs de despliegue totalmente independientes en sus respectivos manifiestos XML, ambas versiones pueden instalarse y ejecutarse simultáneamente en el mismo buzón de correo sin colisión alguna.

---

## 3. Acciones de Higiene Técnica y Saneamiento

1. **Eliminación Física de Rastros y Carpetas Huérfanas:**
   * Se eliminaron físicamente del disco las carpetas y archivos huérfanos que pudieran denotar el uso de asistentes o generadores externos: `.claude/`, `openspec/` y `AGENTS.md` en los entornos de trabajo.
   * Se depuraron referencias colgantes en Git (`refs/original/*`) resultantes de reescrituras de historial previas.
2. **Normalización de Reglas de Ignorados (`.gitignore`):**
   * Se sanearon los archivos `.gitignore` para eliminar cualquier mención a carpetas de agentes, estandarizándolos exclusivamente con las directivas oficiales para aplicaciones React, TypeScript y empaquetado con Webpack (`node_modules/`, `dist/`, `.env*`, `*.log`).
3. **Aislamiento por Git Worktrees:**
   * La estructura de repositorios utiliza `git worktree` para enlazar `Add_in_PFI` y `Add_in_PFI_hybrid_pilot`, permitiendo trabajar en ambas ramas (`entrega` y `hybrid-aws-pilot`) sin necesidad de realizar `git checkout` destructivos ni duplicar el almacenamiento del repositorio base.

---

## 4. Localización Integral al Español Argentino

Se realizó una auditoría completa de todas las cadenas de texto expuestas al usuario final en los archivos `taskpane.tsx`, componentes visuales y manifiestos de ambos complementos:

* **Estados de autenticación de cabeceras:**
  * `pass` ➔ **Aprobado**
  * `fail` ➔ **Falló**
  * `softfail` ➔ **Sospechoso**
  * `none` / `unknown` ➔ **Sin registro**
* **Botones de acción:** Unificación en términos profesionales en español: *"Analizar correo"*, *"Analizar Amenaza"* y *"Reportar correo sospechoso"*.
* **Alertas e indicadores de carga:** Sustitución de cadenas en inglés por mensajes contextuales claros: *"Analizando el correo. Esperá un momento."*, *"Cargando Outlook..."*, *"Variables técnicas influyentes"*.
* **Manifiestos de Outlook:** Corrección de descripciones institucionales y nombres de grupos en la cinta de opciones de Office (*Ribbon*), sustituyendo títulos en inglés por `"PhishARG Hybrid"`.
* **Preservación técnica:** Se conservaron intactas las siglas y términos estándar reconocidos por la industria (SPF, DKIM, DMARC, Return-Path, TLS 1.3, TreeSHAP, LIME).

---

## 5. Matriz de Pruebas Automatizadas y Verificación

La estabilidad funcional del cliente se validó mediante suites de pruebas automatizadas nativas de Node.js:

| Suite | Componente Auditado | Pruebas | Resultado |
| :--- | :--- | :---: | :---: |
| `scripts/test-proxy.mjs` | Proxy seguro en Vercel, inyección de credencial, timeout y límite de payload (1.1 MB) | 7 / 7 | ✅ APROBADO |
| `scripts/test-analysis.cjs` | Contrato de análisis híbrido, mapeo de señales y preservación de veredictos | 11 / 11 | ✅ APROBADO |
| `scripts/test-api-token.cjs` | Emisión y validación de tokens HMAC efímeros, rate limiting y rechazo de orígenes cruzados | 4 / 4 | ✅ APROBADO |

---

## 6. Estado de Sincronización

* **Repositorio principal:** `Add_in_PFI` (Rama `entrega`) — Sincronizado con `origin/entrega`.
* **Piloto avanzado:** `Add_in_PFI_hybrid_pilot` (Rama `hybrid-aws-pilot`) — Sincronizado con `origin/hybrid-aws-pilot`.
* **Copia de respaldo limpia:** `Add_in_PFI_clean` (Rama `entrega`) — 100% idéntica y limpia.

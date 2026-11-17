# Resumen de Saneamiento, Arquitectura y Entrega — PhishARG Hybrid Add-in

**Proyecto Final de Integración (PFI) — Ingeniería en Informática — UADE**  
**Fecha de consolidación:** 17 de noviembre de 2026  
**Rama oficial:** `entrega` (Cliente Hybrid Definitivo para Producción y Evaluación)

---

## 1. Alcance y Objetivos de la Integración

El presente documento formaliza la unificación del complemento de lectura de Outlook de **PhishARG** sobre la rama canónica `entrega`.

A partir de esta intervención, la rama `entrega` de `Add_in_PFI` queda 100% sincronizada con el flujo de la tesis implementado en `flujo_v2/entrega`, consumiendo directamente el servicio **Hybrid XAI** en AWS ECS Fargate a través del endpoint `POST /api/v1/analyze/hybrid`.

---

## 2. Arquitectura del Cliente Hybrid y Despacho en AWS

El cliente implementa un modelo de ciberdefensa en dos fases con aislamiento estricto de privacidad:

```
[ Outlook Desktop / Web ]
         │
         ├─── (1) POST /api/v1/analyze/token ───► [ Vercel Serverless Function ]
         │    Envía únicamente SHA-256(nonce + correo)     - Valida origen y rate limiting
         │    (El texto del correo NUNCA va a Vercel)      - Firma token efímero HMAC SHA-256
         │                                                 - Devuelve token + api_base_url
         ▼
         └─── (2) POST /api/v1/analyze/hybrid ──► [ AWS ECS Fargate (TLS 1.3) ]
              Directo con cabecera Bearer <token>          - Inferencia MiniLM + XGBoost V4
              - Fase 1: quick=true (Veredicto rápido)      - TreeSHAP (Atribución log-odds)
              - Fase 2: Enriquecimiento XAI progresivo     - LIME acotado en texto
                                                           - Explicación SLM Qwen 2.5 3B
```

### Características de Seguridad y Desempeño:
* **Privacidad de extremo a extremo:** El cuerpo de los correos corporativos viaja directamente desde el buzón de Outlook hacia el clúster de AWS bajo cifrado obligatorio TLS 1.3 sin transitar por servidores de terceros.
* **Tokens efímeros acotados al payload:** La autorización de Vercel genera un token firmado válido por 90 segundos estrictamente vinculado a la huella criptográfica del correo y a un número de uso único (*nonce*), imposibilitando ataques de repetición.
* **Inferencia en dos tiempos (*Two-Phase Analysis*):** El usuario recibe el veredicto predictivo de forma instantánea mientras el clúster computa las perturbaciones locales de LIME y las explicaciones didácticas de Qwen 2.5 3B, evitando bloqueos en la interfaz.

---

## 3. Identidad del Manifiesto y Despliegue

* **Identificador de Aplicación (GUID):** `f0355dbe-00bc-47de-8626-f157b7c59135`.
* **Nombre institucional:** `PhishARG Hybrid Piloto`.
* **Grupo en Cinta de Opciones (*Ribbon*):** `PhishARG Hybrid`.
* **Botón principal:** *"Analizar correo"* / *"Analizar amenaza"*.
* **Íconos oficiales:** Emblema institucional de escudo PhishARG en resoluciones oficiales de 16, 32 y 80 píxeles (`assets/phisharg-shield-*.png`).
* **Entornos:**
  * `manifest.xml`: Despliegue de producción alojado en `https://add-in-pfi-hybrid.vercel.app`.
  * `manifest-local.xml`: Entorno local de desarrollo sobre `https://localhost:3000`.

---

## 4. Auditoría de Higiene y Localización

1. **Eliminación Física de Rastros:**
   * Se eliminaron exhaustivamente del repositorio todas las carpetas y archivos huérfanos generados por asistentes externos (`.claude/`, `openspec/`, `AGENTS.md`).
   * Se normalizó `.gitignore` conservando exclusivamente los estándares de empaquetado Node.js / React / Webpack.
2. **Localización Integral al Español Argentino:**
   * Todas las etiquetas de la interfaz (`taskpane.tsx`), badges de cabeceras (`Aprobado`, `Falló`, `Sospechoso`, `Sin registro`), mensajes de carga (*"Analizando el correo. Esperá un momento."*) y descripciones del manifiesto se encuentran en español formal.
   * Se preservaron intactas las nomenclaturas y estándares técnicos internacionales (SPF, DKIM, DMARC, Return-Path, TLS 1.3, TreeSHAP, LIME).

---

## 5. Matriz de Pruebas Automatizadas

La estabilidad de la solución se encuentra validada mediante 15 pruebas unitarias y de integración nativas en Node.js:

```powershell
node --test scripts/test-analysis.cjs scripts/test-api-token.cjs
```

Puntos auditados con éxito (15 / 15 ✅):
1. Cumplimiento estricto del contrato de veredicto devuelto por el backend.
2. Preservación del veredicto de phishing y mapeo de señales heurísticas de contenido.
3. Compatibilidad con la ruta rápida `quick=true` y rechazo seguro de backends no híbridos.
4. Restricción de tráfico: Vercel solo recibe el digest criptográfico salado del mensaje.
5. Emisión de tokens HMAC vinculados al digest con control de orígenes cruzados.
6. Protección contra denegación de servicio por limitación de tamaño de payload (1.1 MB).
7. Cancelación coordinada de solicitudes lentas y políticas de fallback seguro.

# Estructura del Proyecto y Guía de Entrega — PhishARG Add-in

Este documento describe la arquitectura, la organización del repositorio y las pautas de verificación del complemento de Outlook **PhishARG**, desarrollado como parte del Proyecto Final de Integración (PFI).

---

## 1. Arquitectura del Complemento

PhishARG opera como un complemento de lectura de correo (Office.js) articulado en un modelo de dos capas:

```
[ Outlook Desktop / Web ]
          │
          ▼
   (Office.js Panel)
  src/taskpane/taskpane.tsx
          │
          ▼ HTTPS (Mismo origen)
  api/v1/analyze.mjs  ─── (Función Vercel Proxy / Servidor)
          │            - Agrega X-API-Key en el servidor
          │            - Valida tamaño (máx. 1.1 MB) y timeout
          ▼ HTTPS
 [ Backend AWS / Local ]
```

* **Frontend seguro:** La clave de autenticación del backend (`API_KEY`) y la URL interna de AWS no se exponen al navegador del usuario final.
* **Separación de entornos:** `manifest.xml` apunta al despliegue productivo en Vercel, mientras que `manifest-local.xml` permite depuración sobre `https://localhost:3000`.

---

## 2. Mapa de Directorios

* `src/`: Código fuente de la interfaz de usuario y servicios del cliente.
  * `taskpane/`: Componentes React (`taskpane.tsx`, `RiskRing.tsx`), plantilla HTML y estilos CSS accesibles.
  * `services/`: Lógica de integración con Office.js (`mailService.ts`), parsing de resultados (`analysisResult.ts`), extracción de cabeceras de seguridad (`securityService.ts`) y comunicación con el proxy (`analysisService.ts`).
  * `types/`: Definiciones TypeScript de señales de seguridad y modelos de datos.
* `api/`: Funciones del servidor intermediario en Vercel.
  * `v1/analyze.mjs`: Endpoint proxy que valida la solicitud y canaliza la comunicación con AWS.
* `assets/`: Íconos oficiales de la aplicación en resoluciones estándar (16, 32 y 80 píxeles).
* `public/`: Páginas auxiliares estáticas para pruebas de conectividad fuera de Outlook (`comprobar-conexion.html`, `check-connection.js`).
* `scripts/`: Pruebas automatizadas de contrato y proxy (`test-proxy.mjs`, `test-analysis.cjs`).
* `manifest.xml`: Manifiesto de producción para sideload o publicación institucional.
* `manifest-local.xml`: Manifiesto para entorno local de desarrollo.

---

## 3. Protocolo de Pruebas y Verificación

Para validar la integridad técnica del complemento:

```powershell
# Pruebas automatizadas del proxy de red y validación de contrato
node --test scripts/test-proxy.mjs
```

Puntos verificados por la suite:
1. Reenvío seguro con inyección de credencial únicamente en el servidor.
2. Rechazo estricto de métodos HTTP no permitidos y orígenes cruzados no autorizados.
3. Validación y límite de tamaño de carga útil (máx. 1.1 MB).
4. No filtración de diagnósticos ni errores internos del backend.
5. Manejo de timeouts (25 segundos) y códigos de estado estándar (503 ante falta de configuración).

---

## 4. Convenciones de Despliegue y Rama de Entrega

* **Rama oficial de entrega:** `entrega`.
* **Políticas de entorno:** Configurar en Vercel:
  * `API_BASE_URL`: Origen HTTPS del clasificador en AWS.
  * `API_KEY`: Clave autorizada para invocar la API del clasificador.

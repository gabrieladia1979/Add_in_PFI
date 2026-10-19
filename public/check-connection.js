const button = document.getElementById("check");
const status = document.getElementById("status");

button.addEventListener("click", async () => {
  button.disabled = true;
  status.dataset.state = "pending";
  status.textContent = "Comprobando Vercel → AWS…";
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30_000);
  try {
    const response = await fetch("/api/v1/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        metadata: { asunto: "Prueba técnica de conectividad", remitente_email: "prueba@example.invalid", remitente_nombre: "Prueba técnica" },
        contenido: "Hola equipo. La reunión semanal será el viernes a las diez. Este mensaje es una prueba técnica sin datos personales.",
        security_features: null,
      }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok) {
      if (response.status === 503) throw new Error("Falta configurar API_BASE_URL o API_KEY en este despliegue, o el servicio no está disponible. Revisá los registros de Vercel.");
      throw new Error(typeof data?.detail === "string" ? `${data.detail} (HTTP ${response.status})` : `La prueba no pudo completarse (HTTP ${response.status}).`);
    }
    if (typeof data?.is_phishing !== "boolean" || !Number.isFinite(data?.risk_score) || typeof data?.reason !== "string") {
      throw new Error("El servicio devolvió una respuesta incompatible.");
    }
    status.dataset.state = "success";
    status.textContent = "Conexión verificada: Vercel recibió una respuesta válida de AWS. El nuevo recorrido funciona en este despliegue.";
  } catch (error) {
    status.dataset.state = "error";
    status.textContent = controller.signal.aborted ? "La prueba tardó demasiado. Intentá nuevamente." : error.message;
  } finally {
    clearTimeout(timer);
    button.disabled = false;
  }
});

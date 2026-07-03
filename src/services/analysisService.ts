export type AnalyzeResponse = {
  is_phishing: boolean;
  confidence_score: number;
  raw_label: string;
};

const API_BASE_URL = "http://127.0.0.1:8001";

export async function analyzeEmail(subject: string, body: string, sender: string): Promise<AnalyzeResponse> {
  const response = await fetch(`${API_BASE_URL}/api/v1/analyze`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      metadata: {
        asunto: subject,
        remitente_email: sender,
        remitente_nombre: ""
      },
      contenido: body
    }),
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || "Error en el servicio de analisis");
  }

  const data = await response.json();
  
  return {
    is_phishing: data.is_phishing,
    confidence_score: data.risk_score,
    raw_label: data.reason
  };
}

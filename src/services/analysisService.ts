import { SecurityFeatures } from "../types/securityTypes";

export type AnalyzeResponse = {
  is_phishing: boolean;
  confidence_score: number;
  raw_label: string;`n  slm_explanation?: string;
};

const API_BASE_URL = process.env.API_BASE_URL || "http://127.0.0.1:8001";

export async function analyzeEmail(subject: string, body: string, sender: string, securityFeatures: SecurityFeatures | null): Promise<AnalyzeResponse> {
  const response = await fetch(`${API_BASE_URL}/api/v1/analyze`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",`n      "X-API-Key": process.env.API_KEY || "PHISHARG_DEV_KEY_123",
    },
    body: JSON.stringify({
      metadata: {
        asunto: subject,
        remitente_email: sender,
        remitente_nombre: ""
      },
      contenido: body,
      security_features: securityFeatures
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
    raw_label: data.reason,`n    slm_explanation: data.slm_explanation
  };
}

import { SecurityFeatures } from "../types/securityTypes";

export type SlmReason = {
  evidence_id: string;
  text: string;
};

export type SlmAction = {
  code: string;
  text: string;
};

export type SlmExplanation = {
  summary: string;
  reasons: SlmReason[];
  recommended_actions: SlmAction[];
};

export type AnalyzeResponse = {
  is_phishing: boolean;
  confidence_score: number;
  raw_label: string;
  intent?: string;
  slm_explanation?: string;
  slm_parsed?: SlmExplanation | null;
};

const API_BASE_URL = process.env.API_BASE_URL || "http://localhost:8001";

export async function analyzeEmail(subject: string, body: string, sender: string, securityFeatures: SecurityFeatures | null): Promise<AnalyzeResponse> {
  const response = await fetch(`${API_BASE_URL}/api/v1/analyze`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-API-Key": process.env.API_KEY || "PHISHARG_DEV_KEY_123",
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

  // Intentar parsear el JSON del SLM
  let slmParsed: SlmExplanation | null = null;
  if (data.slm_explanation) {
    try {
      slmParsed = JSON.parse(data.slm_explanation);
    } catch {
      slmParsed = null;
    }
  }

  return {
    is_phishing: data.is_phishing,
    confidence_score: data.risk_score,
    raw_label: data.reason,
    intent: data.intent,
    slm_explanation: data.slm_explanation,
    slm_parsed: slmParsed,
  };
}

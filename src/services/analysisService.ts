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
  content_signals?: { rule: string; description: string }[];
};

// En desarrollo, la ruta relativa pasa por el proxy HTTPS del dev server.
const API_BASE_URL = process.env.API_BASE_URL || "";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseSlmExplanation(value: unknown): SlmExplanation | null {
  if (typeof value !== "string") return null;
  try {
    const parsed: unknown = JSON.parse(value);
    if (!isRecord(parsed) || typeof parsed.summary !== "string" ||
        !Array.isArray(parsed.reasons) || !Array.isArray(parsed.recommended_actions)) return null;
    if (!parsed.reasons.every((reason) => isRecord(reason) &&
        typeof reason.evidence_id === "string" && typeof reason.text === "string")) return null;
    if (!parsed.recommended_actions.every((action) => isRecord(action) &&
        typeof action.code === "string" && typeof action.text === "string")) return null;
    return parsed as SlmExplanation;
  } catch {
    return null;
  }
}

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

  const data: unknown = await response.json();
  if (!isRecord(data) || typeof data.is_phishing !== "boolean" ||
      typeof data.risk_score !== "number" || !Number.isFinite(data.risk_score) ||
      data.risk_score < 0 || data.risk_score > 1 || typeof data.reason !== "string") {
    throw new Error("El servicio devolvió un resultado de análisis incompatible.");
  }

  return {
    is_phishing: data.is_phishing,
    confidence_score: data.risk_score,
    raw_label: data.reason,
    intent: typeof data.intent === "string" ? data.intent : undefined,
    slm_explanation: typeof data.slm_explanation === "string" ? data.slm_explanation : undefined,
    slm_parsed: parseSlmExplanation(data.slm_explanation),
    content_signals: Array.isArray(data.content_signals)
      ? data.content_signals.filter((signal): signal is { rule: string; description: string } =>
          isRecord(signal) && typeof signal.rule === "string" && typeof signal.description === "string")
      : [],
  };
}

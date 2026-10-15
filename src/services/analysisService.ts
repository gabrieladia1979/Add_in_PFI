import { SecurityFeatures } from "../types/securityTypes";

export type SlmReason = { evidence_id: string; text: string };
export type SlmAction = { code: string; text: string };
export type SlmExplanation = {
  summary: string;
  reasons: SlmReason[];
  recommended_actions: SlmAction[];
};

export type HybridTechnicalAttribution = {
  feature: string;
  value: number;
  contribution_log_odds: number;
};
export type HybridWordAttribution = { word: string; weight: number };
export type HybridXai = {
  scope: "raw_hybrid_model_score";
  supports_final_verdict: boolean;
  technical_status: "available" | "skipped" | "error";
  technical_attributions: HybridTechnicalAttribution[];
  lime_status: "reliable" | "unreliable" | "skipped" | "error" | "disabled";
  lime_word_attributions: HybridWordAttribution[];
};

export type AnalyzeResponse = {
  is_phishing: boolean;
  confidence_score: number;
  raw_label: string;
  intent?: string;
  slm_explanation?: string;
  slm_parsed?: SlmExplanation | null;
  content_signals: { rule: string; description: string }[];
  decision_source?: string;
  hybrid_xai: HybridXai;
};

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

function parseHybridXai(value: unknown): HybridXai {
  if (!isRecord(value) || value.scope !== "raw_hybrid_model_score" ||
      typeof value.supports_final_verdict !== "boolean" ||
      !["available", "skipped", "error"].includes(String(value.technical_status)) ||
      !["reliable", "unreliable", "skipped", "error", "disabled"].includes(String(value.lime_status)) ||
      !Array.isArray(value.technical_attributions) || !Array.isArray(value.lime_word_attributions)) {
    throw new Error("El backend respondió sin el contrato XAI de Hybrid.");
  }
  const technical = value.technical_attributions.filter((item): item is HybridTechnicalAttribution =>
    isRecord(item) && typeof item.feature === "string" &&
    typeof item.value === "number" && Number.isFinite(item.value) &&
    typeof item.contribution_log_odds === "number" && Number.isFinite(item.contribution_log_odds));
  const words = value.lime_word_attributions.filter((item): item is HybridWordAttribution =>
    isRecord(item) && typeof item.word === "string" &&
    typeof item.weight === "number" && Number.isFinite(item.weight));
  return {
    scope: "raw_hybrid_model_score",
    supports_final_verdict: value.supports_final_verdict as boolean,
    technical_status: value.technical_status as HybridXai["technical_status"],
    technical_attributions: technical,
    lime_status: value.lime_status as HybridXai["lime_status"],
    lime_word_attributions: value.lime_status === "reliable" ? words : [],
  };
}

export async function analyzeEmail(subject: string, body: string, sender: string, securityFeatures: SecurityFeatures | null, quick = false, signal?: AbortSignal): Promise<AnalyzeResponse> {
  const payload = JSON.stringify({
    metadata: { asunto: subject, remitente_email: sender, remitente_nombre: "" },
    contenido: body,
    security_features: securityFeatures,
  });
  const route = quick ? "/api/v1/analyze/hybrid?quick=true" : "/api/v1/analyze/hybrid";
  let endpoint = route;
  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (!__LOCAL_DEV__) {
    if (new TextEncoder().encode(payload).byteLength > 1_100_000) {
      throw new Error("El correo es demasiado grande para analizarlo.");
    }
    const nonceBytes = crypto.getRandomValues(new Uint8Array(16));
    const nonce = Array.from(nonceBytes, (value) => value.toString(16).padStart(2, "0")).join("");
    const digestBytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`${nonce}.${payload}`));
    const digest = Array.from(new Uint8Array(digestBytes), (value) => value.toString(16).padStart(2, "0")).join("");
    const tokenResponse = await fetch("/api/v1/analyze/token", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ nonce, digest }),
      signal,
    });
    if (tokenResponse.status === 429) throw new Error("Se alcanzó el límite temporal de análisis. Intente más tarde.");
    if (!tokenResponse.ok) throw new Error("No se pudo autorizar el análisis del correo.");
    const authorization: unknown = await tokenResponse.json();
    if (!isRecord(authorization) || typeof authorization.token !== "string" ||
        typeof authorization.api_base_url !== "string") {
      throw new Error("La autorización del análisis es inválida.");
    }
    let backend: URL;
    try { backend = new URL(authorization.api_base_url); } catch {
      throw new Error("La dirección del backend es inválida.");
    }
    if (backend.protocol !== "https:" || backend.username || backend.password) {
      throw new Error("La dirección del backend debe usar HTTPS.");
    }
    endpoint = new URL(route, backend).href;
    headers.Authorization = `Bearer ${authorization.token}`;
  }
  const response = await fetch(endpoint, {
    method: "POST",
    headers,
    body: payload,
    signal,
  });

  if (!response.ok) {
    if (response.status === 413) throw new Error("El correo es demasiado grande para analizarlo.");
    if (response.status === 429) throw new Error("El servicio está ocupado. Intente nuevamente en unos segundos.");
    if (response.status === 404) throw new Error("El backend configurado no tiene la ruta Hybrid.");
    if (response.status === 503) throw new Error("El modelo Hybrid no está configurado en el backend.");
    if (response.status === 401) throw new Error("La clave de acceso al backend no es válida.");
    throw new Error(`No se pudo analizar el correo (HTTP ${response.status}).`);
  }

  const raw: unknown = await response.json();
  if (!isRecord(raw) || typeof raw.is_phishing !== "boolean" ||
      typeof raw.risk_score !== "number" || !Number.isFinite(raw.risk_score) ||
      raw.risk_score < 0 || raw.risk_score > 1 || typeof raw.reason !== "string") {
    throw new Error("El backend devolvió un resultado de análisis incompatible.");
  }

  const xai = parseHybridXai(raw.hybrid_xai);
  const signals = Array.isArray(raw.content_signals)
    ? raw.content_signals.filter((signal): signal is { rule: string; description: string } =>
        isRecord(signal) && typeof signal.rule === "string" && typeof signal.description === "string")
    : [];

  return {
    is_phishing: raw.is_phishing,
    confidence_score: raw.risk_score,
    raw_label: raw.reason,
    intent: typeof raw.intent === "string" ? raw.intent : undefined,
    slm_explanation: typeof raw.slm_explanation === "string" ? raw.slm_explanation : undefined,
    slm_parsed: parseSlmExplanation(raw.slm_explanation),
    content_signals: signals,
    decision_source: typeof raw.decision_source === "string" ? raw.decision_source : undefined,
    hybrid_xai: xai,
  };
}

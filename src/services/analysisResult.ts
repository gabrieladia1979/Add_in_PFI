import type { AnalyzeResponse, SlmExplanation } from "./analysisService";

export type AnalysisResult = {
  isPhishing: boolean;
  contentSignals: { rule: string; description: string }[];
  threatLevel: string;
  explanation: string;
  plainServiceExplanation?: string;
  phishScore: number;
  intent?: string;
  slmData?: SlmExplanation | null;
  explanationSource?: "slm" | "fallback";
};

export function buildAnalysisResult(result: AnalyzeResponse): AnalysisResult {
  const score = Math.round((result.confidence_score || 0) * 100);

  let threatLevel = "Bajo";
  if (result.is_phishing && score >= 75) {
    threatLevel = "Peligro Critico";
  } else if (result.is_phishing && score >= 50) {
    threatLevel = "Alto";
  } else if (result.is_phishing) {
    threatLevel = "Medio";
  }

  const explanation = result.raw_label || (result.is_phishing
    ? "Se detectaron patrones similares a phishing. Revisa remitente, enlaces y urgencia."
    : "No se detectaron patrones fuertes de phishing, pero revisa enlaces y remitente.");

  return {
    isPhishing: result.is_phishing,
    contentSignals: result.content_signals || [],
    threatLevel,
    explanation,
    plainServiceExplanation: !result.slm_parsed && result.slm_explanation?.trim() &&
      !/^[{\[]/.test(result.slm_explanation.trim())
      ? result.slm_explanation.trim()
      : undefined,
    phishScore: score,
    intent: result.intent,
    slmData: result.slm_parsed,
    explanationSource: result.explanation_source,
  };
}


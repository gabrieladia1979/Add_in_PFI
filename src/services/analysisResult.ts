import type { AnalyzeResponse, SlmAction, SlmExplanation } from "./analysisService";

// Baseline advice is owned by the add-in and is shown even without SLM output.
const BASELINE_RECOMMENDATIONS: SlmAction[] = [
  {
    code: "stay_cautious",
    text: "Mantené la precaución: un resultado de bajo riesgo no garantiza que el correo sea seguro.",
  },
  {
    code: "verify_via_official_channel",
    text: "Verificá las solicitudes inesperadas por un canal conocido: llamá a un número de confianza o ingresá al sitio oficial por tu cuenta.",
  },
  {
    code: "check_sender_domain",
    text: "Si no esperabas este correo, revisá la dirección completa y el dominio del remitente antes de abrir enlaces o adjuntos. No te guíes solo por el nombre visible.",
  },
];

export type AnalysisResult = {
  isPhishing: boolean;
  contentSignals: { rule: string; description: string }[];
  threatLevel: string;
  explanation: string;
  plainServiceExplanation?: string;
  plainServiceSummary?: string;
  plainServiceReasons?: string[];
  phishScore: number;
  intent?: string;
  slmData?: SlmExplanation | null;
  explanationSource?: "slm" | "fallback";
  recommendedActions: SlmAction[];
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

  // Classifier keywords are not verified excerpts or contextual phishing signals.
  // Keep them out of the user-facing explanation in both response formats.
  const slmData = result.slm_parsed ? {
    ...result.slm_parsed,
    reasons: result.slm_parsed.reasons.filter((reason) =>
      reason.evidence_id !== "evidence.keywords"
      && !/^\s*Palabras clave detectadas\s*:/i.test(reason.text)
    ),
  } : result.slm_parsed;
  const plainServiceExplanation = !slmData && result.slm_explanation?.trim() &&
    !/^[{\[]/.test(result.slm_explanation.trim())
    ? result.slm_explanation
      .replace(/(?:\d+\.\s*)?Palabras clave detectadas\s*:[^\r\n]*?(?=\r?\n|\s+\d+\.\s|\s+Acciones recomendadas:|$)/gi, "")
      .trim()
    : undefined;
  const reasonsText = plainServiceExplanation?.match(/(?:^|\s)Motivos:\s*(.*?)\s+Acciones recomendadas:/s)?.[1];
  const plainServiceReasons = reasonsText
    ?.split(/(?:^|\s)\d+\.\s+/)
    .map((reason) => reason.trim())
    .filter(Boolean)
    .filter((reason) => !(result.intent === "desviar_pago" && reason === "Se detectó contenido relacionado con pagos o información financiera."));

  const serviceActions = result.slm_parsed?.recommended_actions;
  const specificActions: SlmAction[] = serviceActions?.length
    ? serviceActions
    : result.is_phishing
      ? result.intent === "desviar_pago"
        ? [{ code: "do_not_pay", text: "No transfieras dinero a la cuenta indicada hasta confirmar el cambio." }]
        : [{ code: "do_not_click", text: "No abras los enlaces ni adjuntos hasta verificar el mensaje." }]
      : [];

  return {
    isPhishing: result.is_phishing,
    contentSignals: result.content_signals || [],
    threatLevel,
    explanation,
    plainServiceExplanation,
    plainServiceSummary: plainServiceReasons?.length
      ? plainServiceExplanation?.split(/\s+Motivos:/)[0].trim()
      : undefined,
    plainServiceReasons: plainServiceReasons?.length ? plainServiceReasons : undefined,
    phishScore: score,
    intent: result.intent,
    slmData,
    explanationSource: result.explanation_source,
    recommendedActions: [
      ...specificActions.filter(
        (action) => action.code !== "no_action_needed"
          && !BASELINE_RECOMMENDATIONS.some((baseline) => baseline.code === action.code)
      ),
      ...BASELINE_RECOMMENDATIONS,
    ],
  };
}


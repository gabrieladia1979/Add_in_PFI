import type { AnalyzeResponse as BaseAnalyzeResponse } from "./analysisService";

export type AnalyzeResponse = BaseAnalyzeResponse & {
  readonly explanations?: readonly string[];
};

type MockAnalysisFixture = {
  readonly subjectMarker: string;
  readonly response: AnalyzeResponse;
};

const MOCK_ANALYSIS_DELAY_MS = 700;

const MOCK_ANALYSIS_FIXTURES: readonly MockAnalysisFixture[] = [
  {
    subjectMarker: "[PFI-DEMO-01]",
    response: {
      is_phishing: true,
      confidence_score: 0.98,
      raw_label:
        "El correo utiliza urgencia y solicita validar credenciales desde un enlace externo.",
      explanations: [
        "El asunto comunica que la contraseña vence hoy para provocar una reacción inmediata.",
        "El mensaje solicita actualizar credenciales mediante un enlace incluido en el correo.",
        "El destino del enlace no coincide con el flujo habitual de administración de la cuenta.",
        "La combinación de urgencia, credenciales y navegación externa eleva el riesgo de phishing.",
      ],
    },
  },
  {
    subjectMarker: "[PFI-DEMO-02]",
    response: {
      is_phishing: true,
      confidence_score: 0.64,
      raw_label:
        "El mensaje presenta señales sospechosas en el remitente y en la solicitud de pago.",
      explanations: [
        "La identidad del remitente presenta inconsistencias con la organización que dice representar.",
        "El correo introduce una solicitud de pago que no forma parte de una conversación verificada.",
        "La operación propuesta intenta sacar la validación del circuito habitual de la organización.",
        "La combinación de remitente dudoso y pedido financiero es una señal frecuente de fraude.",
      ],
    },
  },
  {
    subjectMarker: "[PFI-DEMO-03]",
    response: {
      is_phishing: false,
      confidence_score: 0.08,
      raw_label:
        "No se detectaron patrones relevantes de phishing en el correo analizado.",
      explanations: [
        "El remitente coincide con la identidad esperada para este tipo de comunicación.",
        "El mensaje mantiene un tono informativo, sin presión ni urgencia manipulativa.",
        "No solicita contraseñas, datos personales ni pagos.",
        "No se observan enlaces externos sospechosos ni desvíos hacia sitios no relacionados.",
      ],
    },
  },
];

function waitForDemoDelay(): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, MOCK_ANALYSIS_DELAY_MS);
  });
}

export async function analyzeEmail(
  subject: string,
  _body: string,
  _sender: string,
): Promise<AnalyzeResponse> {
  await waitForDemoDelay();

  const normalizedSubject = subject.toUpperCase();
  const fixture = MOCK_ANALYSIS_FIXTURES.find(({ subjectMarker }) =>
    normalizedSubject.includes(subjectMarker),
  );

  if (!fixture) {
    throw new Error(
      "No existe una prediccion mock para el asunto del correo seleccionado",
    );
  }

  return {
    ...fixture.response,
    explanations: fixture.response.explanations
      ? [...fixture.response.explanations]
      : undefined,
  };
}

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
      status: "success",
      is_phishing: true,
      score: 0.98,
      class_id: 1,
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
      status: "success",
      is_phishing: true,
      score: 0.64,
      class_id: 1,
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
      status: "success",
      is_phishing: false,
      score: 0.08,
      class_id: 0,
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

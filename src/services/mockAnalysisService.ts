import type { AnalyzeResponse } from "./analysisService";

export type { AnalyzeResponse } from "./analysisService";

type MockAnalysisFixture = {
  subjectMarker: string;
  response: AnalyzeResponse;
};

const MOCK_ANALYSIS_DELAY_MS = 700;

const MOCK_ANALYSIS_FIXTURES: MockAnalysisFixture[] = [
  {
    subjectMarker: "[PFI-DEMO-01]",
    response: {
      is_phishing: true,
      confidence_score: 0.98,
      raw_label:
        "El correo utiliza urgencia y solicita validar credenciales desde un enlace externo.",
    },
  },
  {
    subjectMarker: "[PFI-DEMO-02]",
    response: {
      is_phishing: true,
      confidence_score: 0.64,
      raw_label:
        "El mensaje presenta señales sospechosas en el remitente y en la solicitud de pago.",
    },
  },
  {
    subjectMarker: "[PFI-DEMO-03]",
    response: {
      is_phishing: false,
      confidence_score: 0.08,
      raw_label:
        "No se detectaron patrones relevantes de phishing en el correo analizado.",
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

  return { ...fixture.response };
}

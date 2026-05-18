export type AnalyzeResponse = {
  is_phishing: boolean;
  confidence_score: number;
  raw_label: string;
};

const API_BASE_URL = "http://127.0.0.1:8000";

export async function analyzeEmail(subject: string, body: string): Promise<AnalyzeResponse> {
  const response = await fetch(`${API_BASE_URL}/api/analyze`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ subject, body }),
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || "Error en el servicio de analisis");
  }

  return response.json() as Promise<AnalyzeResponse>;
}

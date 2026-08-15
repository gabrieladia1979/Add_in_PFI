// Example: "https://abcd-123-456.trycloudflare.com/predict"
const API_URL: string =
  "https://underwear-suites-hiring-genre.trycloudflare.com/predict";
const API_TIMEOUT_MS = 15_000;

export type AnalyzeResponse = {
  status: "success";
  is_phishing: boolean;
  score: number;
  class_id: number;
};

function isAnalyzeResponse(value: unknown): value is AnalyzeResponse {
  if (!value || typeof value !== "object") {
    return false;
  }

  const result = value as Record<string, unknown>;
  return (
    result.status === "success" &&
    typeof result.is_phishing === "boolean" &&
    typeof result.score === "number" &&
    Number.isFinite(result.score) &&
    typeof result.class_id === "number" &&
    Number.isFinite(result.class_id)
  );
}

export async function analyzeEmail(
  subject: string,
  body: string,
  sender: string,
): Promise<AnalyzeResponse> {
  if (API_URL === "PEGAR_ACA_LA_URL_COMPLETA_DE_TRYCLOUDFLARE") {
    throw new Error(
      "Configura API_URL con la URL completa del endpoint de Cloudflare Tunnel",
    );
  }

  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), API_TIMEOUT_MS);

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ subject, body, sender }),
      signal: controller.signal,
    });

    if (!response.ok) {
      const message = await response.text();
      throw new Error(
        message || `La API respondio con HTTP ${response.status}`,
      );
    }

    const data: unknown = await response.json();
    if (!isAnalyzeResponse(data)) {
      throw new Error("La API devolvio una respuesta invalida");
    }

    return data;
  } catch (error) {
    const requestError =
      error instanceof Error && error.name === "AbortError"
        ? new Error(
            `La API no respondio dentro de ${API_TIMEOUT_MS / 1000} segundos`,
          )
        : error;
    console.error("Error al llamar a la API de analisis:", requestError);
    throw requestError;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

export type AnalyzeRequest = {
  subject: string;
  body: string;
  sender: string;
};

export type AnalyzeResponse = {
  status: "success";
  is_phishing: boolean;
  score: number;
  class_id: number;
};

const API_URL = "http://localhost:8000/api/v1/analyze";

function isAnalyzeResponse(value: unknown): value is AnalyzeResponse {
  if (typeof value !== "object" || value === null) {
    return false;
  }

  const data = value as Record<string, unknown>;
  return (
    data.status === "success" &&
    typeof data.is_phishing === "boolean" &&
    typeof data.score === "number" &&
    Number.isFinite(data.score) &&
    data.score >= 0 &&
    data.score <= 1 &&
    typeof data.class_id === "number" &&
    Number.isInteger(data.class_id) &&
    (data.class_id === 0 || data.class_id === 1) &&
    data.class_id === (data.is_phishing ? 1 : 0)
  );
}

export async function analyzeEmail(
  subject: string,
  body: string,
  sender: string,
): Promise<AnalyzeResponse> {
  const payload: AnalyzeRequest = { subject, body, sender };
  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || "The analysis service returned an error");
  }

  const data: unknown = await response.json();
  if (!isAnalyzeResponse(data)) {
    throw new Error("The analysis service returned an invalid response");
  }

  return data;
}

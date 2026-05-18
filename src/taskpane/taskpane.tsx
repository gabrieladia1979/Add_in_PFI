import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { getCurrentMessageMetadata, MailMetadata } from "../services/mailService";
import { analyzeEmail, AnalyzeResponse } from "../services/analysisService";
import "./taskpane.css";

type AnalysisResult = {
  threatLevel: string;
  explanation: string;
  phishScore: number;
};

function buildAnalysisResult(result: AnalyzeResponse): AnalysisResult {
  const score = Math.round((result.confidence_score || 0) * 100);

  let threatLevel = "Bajo";
  if (result.is_phishing && score >= 75) {
    threatLevel = "Peligro Critico";
  } else if (result.is_phishing && score >= 50) {
    threatLevel = "Alto";
  } else if (result.is_phishing) {
    threatLevel = "Medio";
  }

  const explanation = result.is_phishing
    ? "Se detectaron patrones similares a phishing. Revisa remitente, enlaces y urgencia."
    : "No se detectaron patrones fuertes de phishing, pero revisa enlaces y remitente.";

  return {
    threatLevel,
    explanation,
    phishScore: score,
  };
}

function TaskPane(): JSX.Element {
  const [officeReady, setOfficeReady] = useState(false);
  const [isLinked, setIsLinked] = useState(false);
  const [mailMetadata, setMailMetadata] = useState<MailMetadata | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [analysisError, setAnalysisError] = useState<string | null>(null);

  useEffect(() => {
    if (typeof Office === "undefined") {
      return;
    }

    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        setOfficeReady(true);
      }
    });
  }, []);

  useEffect(() => {
    if (!officeReady || !isLinked) {
      return;
    }

    let cancelled = false;
    setLoadError(null);
    setMailMetadata(null);

    getCurrentMessageMetadata()
      .then((metadata) => {
        if (cancelled) {
          return;
        }
        if (!metadata) {
          setLoadError("No hay mensaje abierto.");
          return;
        }
        setMailMetadata(metadata);
      })
      .catch(() => {
        if (!cancelled) {
          setLoadError("No se pudo leer el mensaje.");
        }
      });

    return () => {
      cancelled = true;
    };
  }, [officeReady, isLinked]);

  const handleLink = (): void => {
    setIsLinked(true);
  };

  const handleAnalyze = async (): Promise<void> => {
    if (!mailMetadata) {
      return;
    }

    setIsAnalyzing(true);
    setAnalysisResult(null);
    setAnalysisError(null);

    try {
      const response = await analyzeEmail(mailMetadata.subject, mailMetadata.body);
      setAnalysisResult(buildAnalysisResult(response));
    } catch (error) {
      setAnalysisError("No se pudo analizar el correo.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  if (!officeReady) {
    return (
      <div className="min-h-screen bg-slate-50 px-6 py-8 text-slate-900">
        <p className="text-sm text-slate-600">Cargando Outlook...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 px-6 py-8 text-slate-900">
      <header className="mb-8">
        <p className="text-xs font-semibold uppercase tracking-wide text-emerald-600">
          Phish-Score Lab
        </p>
        <h1 className="mt-2 text-2xl font-semibold">Panel de Defensa Proactiva</h1>
        <p className="mt-2 text-sm text-slate-600">
          Educa en tiempo real y explica con XAI por que un correo es sospechoso.
        </p>
      </header>

      {!isLinked ? (
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <h2 className="text-lg font-semibold">Bienvenido</h2>
          <p className="mt-2 text-sm text-slate-600">
            Vincula tu cuenta para comenzar el analisis contextual del correo abierto.
          </p>
          <button
            className="mt-6 w-full rounded-xl bg-emerald-600 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-emerald-700"
            type="button"
            onClick={handleLink}
          >
            Vincular Outlook
          </button>
        </section>
      ) : (
        <section className="space-y-6">
          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-lg font-semibold">Correo actual</h2>
            {loadError ? (
              <p className="mt-3 text-sm text-rose-600">{loadError}</p>
            ) : (
              <div className="mt-4 space-y-3 text-sm text-slate-700">
                <div>
                  <span className="text-xs font-semibold uppercase text-slate-400">
                    Remitente
                  </span>
                  <p className="mt-1 font-medium text-slate-900">
                    {mailMetadata?.sender || "-"}
                  </p>
                </div>
                <div>
                  <span className="text-xs font-semibold uppercase text-slate-400">
                    Asunto
                  </span>
                  <p className="mt-1 font-medium text-slate-900">
                    {mailMetadata?.subject || "-"}
                  </p>
                </div>
                <div>
                  <span className="text-xs font-semibold uppercase text-slate-400">
                    Cuerpo (extracto)
                  </span>
                  <p className="mt-1 text-slate-700">
                    {mailMetadata?.bodyPreview || "-"}
                  </p>
                </div>
              </div>
            )}
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-lg font-semibold">Accion principal</h2>
            <p className="mt-2 text-sm text-slate-600">
              Ejecuta el analisis simulado con IA para generar explicaciones XAI.
            </p>
            <button
              className="mt-4 w-full rounded-xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
              type="button"
              onClick={handleAnalyze}
              disabled={!mailMetadata || isAnalyzing}
            >
              {isAnalyzing ? "Analizando..." : "Analizar Amenaza con IA"}
            </button>
          </div>

          {(isAnalyzing || analysisResult || analysisError) && (
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <h2 className="text-lg font-semibold">Resultado de la simulacion</h2>
              {isAnalyzing ? (
                <p className="mt-3 text-sm text-slate-600">
                  Ejecutando modelos de riesgo...
                </p>
              ) : analysisError ? (
                <p className="mt-3 text-sm text-rose-600">{analysisError}</p>
              ) : (
                analysisResult && (
                  <div className="mt-4 space-y-4">
                    <div>
                      <span className="text-xs font-semibold uppercase text-slate-400">
                        Nivel de amenaza
                      </span>
                      <p className="mt-1 text-base font-semibold text-rose-600">
                        {analysisResult.threatLevel}
                      </p>
                    </div>
                    <div className="rounded-xl border border-emerald-100 bg-emerald-50 p-4 text-sm text-emerald-900">
                      {analysisResult.explanation}
                    </div>
                    <div>
                      <div className="flex items-center justify-between text-xs font-semibold uppercase text-slate-400">
                        <span>Phish-Score</span>
                        <span>{analysisResult.phishScore}/100</span>
                      </div>
                      <div className="mt-2 h-2 w-full rounded-full bg-slate-200">
                        <div
                          className="h-2 rounded-full bg-rose-500"
                          style={{ width: `${analysisResult.phishScore}%` }}
                        />
                      </div>
                    </div>
                  </div>
                )
              )}
            </div>
          )}
        </section>
      )}
    </div>
  );
}

const rootElement = document.getElementById("root");
if (rootElement) {
  createRoot(rootElement).render(<TaskPane />);
}

export default TaskPane;

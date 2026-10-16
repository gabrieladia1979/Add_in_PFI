import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { getCurrentMessageMetadata, MailMetadata } from "../services/mailService";
import { analyzeEmail, AnalyzeResponse } from "../services/analysisService";
import "./taskpane.css";

type AnalysisResult = {
  isPhishing: boolean;
  threatLevel: string;
  explanation: string;
  score: number;
  classId: number;
};

function buildAnalysisResult(result: AnalyzeResponse): AnalysisResult {
  const threatLevel = result.is_phishing ? "Alto" : "Bajo";

  const explanation = result.is_phishing
    ? "Se detectaron patrones similares a phishing. Revisa remitente, enlaces y urgencia."
    : "No se detectaron patrones fuertes de phishing, pero revisa enlaces y remitente.";

  return {
    isPhishing: result.is_phishing,
    threatLevel,
    explanation,
    score: result.score,
    classId: result.class_id,
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
  const [itemVersion, setItemVersion] = useState(0);
  const analysisRequestId = useRef(0);

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
    if (!officeReady || !Office.context.requirements.isSetSupported("Mailbox", "1.5")) return;
    const mailbox = Office.context.mailbox;
    const onItemChanged = (): void => {
      analysisRequestId.current += 1;
      setAnalysisResult(null);
      setAnalysisError(null);
      setIsAnalyzing(false);
      setItemVersion((version) => version + 1);
    };
    mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    return () => mailbox.removeHandlerAsync(Office.EventType.ItemChanged);
  }, [officeReady]);

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
  }, [officeReady, isLinked, itemVersion]);

  const handleLink = (): void => {
    setIsLinked(true);
  };

  const handleAnalyze = async (): Promise<void> => {
    if (!mailMetadata) {
      return;
    }

    const requestId = ++analysisRequestId.current;
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setAnalysisError(null);

    try {
      const response = await analyzeEmail(mailMetadata.subject, mailMetadata.body, mailMetadata.sender);
      if (requestId === analysisRequestId.current) setAnalysisResult(buildAnalysisResult(response));
    } catch (error) {
      console.error("No se pudo analizar el correo:", error);
      if (requestId === analysisRequestId.current) setAnalysisError("No se pudo analizar el correo.");
    } finally {
      if (requestId === analysisRequestId.current) setIsAnalyzing(false);
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
              Envia el correo abierto al servicio de clasificacion de amenazas.
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
              <h2 className="text-lg font-semibold">Resultado del analisis</h2>
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
                    <div className="rounded-xl border border-slate-200 bg-white p-4 text-sm text-slate-700">
                      <h3 className="font-semibold">Recomendación</h3>
                      <p className="mt-2">
                        {analysisResult.isPhishing
                          ? "No abras enlaces ni adjuntos hasta verificar el mensaje por un canal oficial iniciado por vos."
                          : "Si tenés dudas, comprobá el remitente y los enlaces por un canal oficial."}
                      </p>
                    </div>
                    <div>
                      <span className="text-xs font-semibold uppercase text-slate-400">
                        Puntaje del modelo
                      </span>
                      <p className="mt-1 font-medium text-slate-900">{analysisResult.score}</p>
                    </div>
                    <p className="text-xs text-slate-500">
                      Clase del modelo: {analysisResult.classId}
                    </p>
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

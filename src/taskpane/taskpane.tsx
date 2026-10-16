import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  getCurrentMessageMetadata,
  MailMetadata,
} from "../services/mailService";
import { analyzeEmail, AnalyzeResponse } from "../services/mockAnalysisService";
import "./taskpane.css";

type AnalysisExplanation = {
  id: string;
  text: string;
};

type AnalysisResult = {
  isPhishing: boolean;
  threatLevel: string;
  explanations: AnalysisExplanation[];
  phishScore: number;
};

function buildAnalysisResult(result: AnalyzeResponse): AnalysisResult {
  const score = Math.round((result.confidence_score || 0) * 100);

  let threatLevel = "bajo";
  if (result.is_phishing && score >= 75) {
    threatLevel = "crítico";
  } else if (result.is_phishing && score >= 50) {
    threatLevel = "alto";
  } else if (result.is_phishing) {
    threatLevel = "medio";
  }

  const explanation =
    result.raw_label ||
    (result.is_phishing
      ? "Se detectaron patrones similares a phishing. Revisa remitente, enlaces y urgencia."
      : "No se detectaron patrones fuertes de phishing, pero revisa enlaces y remitente.");
  const specificExplanations = result.explanations?.filter(
    (item) => item.trim().length > 0,
  );
  const explanationTexts =
    specificExplanations && specificExplanations.length > 0
      ? specificExplanations
      : [explanation];

  return {
    isPhishing: result.is_phishing,
    threatLevel,
    explanations: explanationTexts.map((text, index) => ({
      id: `xai-${index}`,
      text,
    })),
    phishScore: score,
  };
}

type AnalysisViewProps = {
  result: AnalysisResult;
  onBack: () => void;
};

function AnalysisView({ result, onBack }: AnalysisViewProps): JSX.Element {
  const { isPhishing, threatLevel, explanations, phishScore } = result;
  const headingRef = useRef<HTMLHeadingElement>(null);
  const recommendations = isPhishing
    ? [
        "No hagas clic en los enlaces del correo.",
        "Reportá el mensaje como phishing en Outlook.",
        "Notificá al equipo de seguridad de tu organización.",
      ]
    : [
        "Verificá que el dominio del remitente coincida con la organización que dice representar.",
        "Revisá el destino de los enlaces antes de abrirlos.",
      ];

  useEffect(() => {
    headingRef.current?.focus();
  }, []);

  return (
    <main
      className="min-h-screen bg-slate-100 px-5 py-6 text-slate-900"
      aria-labelledby="analysis-heading"
    >
      <div className="mx-auto max-w-md space-y-4">
        <section
          className={`rounded-xl border p-5 ${
            isPhishing
              ? "border-rose-300 bg-rose-50"
              : "border-emerald-300 bg-emerald-50"
          }`}
          aria-live="polite"
        >
          <div className="flex items-center gap-4">
            <span
              className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-full text-xl font-bold text-white ${
                isPhishing ? "bg-rose-600" : "bg-emerald-600"
              }`}
              aria-hidden="true"
            >
              {isPhishing ? "×" : "✓"}
            </span>
            <h1
              id="analysis-heading"
              ref={headingRef}
              tabIndex={-1}
              className={`text-xl font-bold ${
                isPhishing ? "text-rose-700" : "text-emerald-700"
              }`}
            >
              {isPhishing
                ? "Amenaza de phishing detectada"
                : "Sin amenaza detectada"}
            </h1>
          </div>
          <p className="mt-5 text-base font-medium text-slate-600">
            Puntuación de riesgo: {phishScore}/100
          </p>
          <p className="mt-3 text-xl font-bold text-slate-900">
            Riesgo {threatLevel}
          </p>
        </section>

        <section className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="text-base font-semibold">
            Análisis de IA ¿Por qué {isPhishing ? "sí" : "no"} es phishing?
          </h2>
          <ul className="mt-4 space-y-3 text-sm leading-6 text-slate-600">
            {explanations.map(({ id, text }) => (
              <li key={id} className="flex gap-3">
                <span
                  className={`mt-2 h-2 w-2 shrink-0 rounded-full ${
                    isPhishing ? "bg-rose-500" : "bg-emerald-500"
                  }`}
                  aria-hidden="true"
                />
                <span>{text}</span>
              </li>
            ))}
          </ul>
        </section>

        <section className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="text-base font-semibold">Acciones recomendadas</h2>
          <ul className="mt-3 space-y-2 text-sm leading-5 text-slate-600">
            {recommendations.map((recommendation) => (
              <li key={recommendation} className="flex gap-3">
                <span className="text-slate-400" aria-hidden="true">
                  •
                </span>
                <span>{recommendation}</span>
              </li>
            ))}
          </ul>
        </section>

        <button
          className="w-full rounded-xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900"
          type="button"
          onClick={onBack}
        >
          Volver al correo
        </button>
      </div>
    </main>
  );
}

function TaskPane(): JSX.Element {
  const [officeReady, setOfficeReady] = useState(false);
  const [isLinked, setIsLinked] = useState(false);
  const [mailMetadata, setMailMetadata] = useState<MailMetadata | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(
    null,
  );
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [itemVersion, setItemVersion] = useState(0);
  const analysisRequestId = useRef(0);
  const analyzeButtonRef = useRef<HTMLButtonElement>(null);
  const shouldRestoreAnalyzeFocus = useRef(false);

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

  useEffect(() => {
    if (!analysisResult && shouldRestoreAnalyzeFocus.current) {
      shouldRestoreAnalyzeFocus.current = false;
      analyzeButtonRef.current?.focus();
    }
  }, [analysisResult]);

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
      const response = await analyzeEmail(
        mailMetadata.subject,
        mailMetadata.body,
        mailMetadata.sender,
      );
      if (requestId === analysisRequestId.current) setAnalysisResult(buildAnalysisResult(response));
    } catch (error) {
      if (requestId === analysisRequestId.current) setAnalysisError("No se pudo analizar el correo.");
    } finally {
      if (requestId === analysisRequestId.current) setIsAnalyzing(false);
    }
  };

  const handleBackToMessage = (): void => {
    shouldRestoreAnalyzeFocus.current = true;
    setAnalysisResult(null);
    setAnalysisError(null);
  };

  if (!officeReady) {
    return (
      <div className="min-h-screen bg-slate-50 px-6 py-8 text-slate-900">
        <p className="text-sm text-slate-600">Cargando Outlook...</p>
      </div>
    );
  }

  if (analysisResult) {
    return (
      <AnalysisView result={analysisResult} onBack={handleBackToMessage} />
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 px-6 py-8 text-slate-900">
      <header className="mb-8">
        <p className="text-xs font-semibold uppercase tracking-wide text-emerald-600">
          Phish-Score Lab
        </p>
        <h1 className="mt-2 text-2xl font-semibold">
          Panel de Defensa Proactiva
        </h1>
        <p className="mt-2 text-sm text-slate-600">
          Educa en tiempo real y explica con XAI por que un correo es
          sospechoso.
        </p>
      </header>

      {!isLinked ? (
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <h2 className="text-lg font-semibold">Bienvenido</h2>
          <p className="mt-2 text-sm text-slate-600">
            Vincula tu cuenta para comenzar el analisis contextual del correo
            abierto.
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
              Ejecuta el analisis simulado con IA para generar explicaciones
              XAI.
            </p>
            <button
              ref={analyzeButtonRef}
              className="mt-4 w-full rounded-xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
              type="button"
              onClick={handleAnalyze}
              disabled={!mailMetadata || isAnalyzing}
              aria-busy={isAnalyzing}
            >
              {isAnalyzing ? "Analizando..." : "Analizar Amenaza con IA"}
            </button>
            {isAnalyzing && (
              <p className="mt-3 text-sm text-slate-600" role="status">
                Ejecutando modelos de riesgo...
              </p>
            )}
            {analysisError && (
              <p className="mt-3 text-sm text-rose-600" role="alert">
                {analysisError} Intentá nuevamente.
              </p>
            )}
          </div>
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

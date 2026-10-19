import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { getCurrentMessageMetadata, MailMetadata, getMailboxItem, prepareIncidentReport } from "../services/mailService";
import { analyzeEmail } from "../services/analysisService";
import { buildAnalysisResult, AnalysisResult } from "../services/analysisResult";
import { extractSecurityFeatures } from "../services/securityService";
import { SecurityFeatures } from "../types/securityTypes";
import { RiskRing } from "./RiskRing";
import "./taskpane.css";

const ANALYSIS_TIMEOUT_MS = 30_000;

function AuthBadge({ label, result }: { label: string; result: string }): JSX.Element {
  let colorClasses = "bg-slate-100 text-slate-600";
  let icon = "❓";

  if (result === "pass") {
    colorClasses = "bg-emerald-100 text-emerald-700";
    icon = "✅";
  } else if (result === "fail") {
    colorClasses = "bg-rose-100 text-rose-700";
    icon = "❌";
  } else if (result === "softfail" || result === "neutral") {
    colorClasses = "bg-amber-100 text-amber-700";
    icon = "⚠️";
  }

  return (
    <span className={`inline-flex items-center gap-1 rounded-full px-3 py-1 text-xs font-semibold ${colorClasses}`}>
      {icon} {label}: {result}
    </span>
  );
}

function TaskPane(): JSX.Element {
  const [officeReady, setOfficeReady] = useState(false);
  const [isLinked, setIsLinked] = useState(false);
  const [mailMetadata, setMailMetadata] = useState<MailMetadata | null>(null);
  const [securityFeatures, setSecurityFeatures] = useState<SecurityFeatures | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [reportRecipient, setReportRecipient] = useState("");
  const [reportStatus, setReportStatus] = useState<string | null>(null);
  const [itemVersion, setItemVersion] = useState(0);
  const analysisRequestId = useRef(0);
  const activeAnalysisController = useRef<AbortController | null>(null);

  useEffect(() => () => activeAnalysisController.current?.abort(), []);

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
      activeAnalysisController.current?.abort();
      setAnalysisResult(null);
      setAnalysisError(null);
      setIsAnalyzing(false);
      setReportStatus(null);
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
    setSecurityFeatures(null);

    const item = getMailboxItem();
    if (item) {
      extractSecurityFeatures(item)
        .then((features) => {
          if (!cancelled) setSecurityFeatures(features);
        })
        .catch(() => {});
    }


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

  const handleUnlink = (): void => {
    // Ignore responses from analyses started before this local UI reset.
    analysisRequestId.current += 1;
    activeAnalysisController.current?.abort();
    setIsLinked(false);
    setMailMetadata(null);
    setSecurityFeatures(null);
    setAnalysisResult(null);
    setLoadError(null);
    setAnalysisError(null);
    setIsAnalyzing(false);
    setReportRecipient("");
    setReportStatus(null);
  };

  const handleRefresh = (): void => {
    analysisRequestId.current += 1;
    activeAnalysisController.current?.abort();
    setAnalysisResult(null);
    setAnalysisError(null);
    setIsAnalyzing(false);
    setReportStatus(null);
    setItemVersion((version) => version + 1);
  };

  const handleCancel = (): void => {
    analysisRequestId.current += 1;
    activeAnalysisController.current?.abort();
    setIsAnalyzing(false);
    setAnalysisError("Análisis cancelado. Podés volver a intentarlo.");
  };

  const handleAnalyze = async (): Promise<void> => {
    if (!mailMetadata) {
      return;
    }

    const requestId = ++analysisRequestId.current;
    const controller = new AbortController();
    activeAnalysisController.current = controller;
    const timeoutId = window.setTimeout(() => controller.abort(), ANALYSIS_TIMEOUT_MS);
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setAnalysisError(null);

    try {
      const response = await analyzeEmail(mailMetadata.subject, mailMetadata.body, mailMetadata.senderEmail, securityFeatures, controller.signal);
      if (requestId === analysisRequestId.current) {
        setAnalysisResult(buildAnalysisResult(response));
      }
    } catch (error) {
      if (requestId === analysisRequestId.current) {
        setAnalysisError(error instanceof Error && error.name === "AbortError"
          ? "El análisis superó los 30 segundos. Podés volver a intentarlo."
          : "No se pudo analizar el correo.");
      }
    } finally {
      window.clearTimeout(timeoutId);
      if (activeAnalysisController.current === controller) activeAnalysisController.current = null;
      if (requestId === analysisRequestId.current) {
        setIsAnalyzing(false);
      }
    }
  };

  const handleReport = (): void => {
    setReportStatus(null);
    if (!mailMetadata) return;
    const recipient = reportRecipient.trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
      setReportStatus("Ingresá el correo del área de seguridad de tu organización.");
      return;
    }
    try {
      prepareIncidentReport(recipient, mailMetadata, analysisResult?.phishScore);
      setReportStatus("Se abrió un borrador en Outlook. Revisalo y pulsá Enviar para completar el reporte.");
    } catch (error) {
      setReportStatus(error instanceof Error ? error.message : "No se pudo preparar el reporte.");
    }
  };

  const recommendedActions = analysisResult?.slmData?.recommended_actions?.length
    ? analysisResult.slmData.recommended_actions
    : analysisResult?.isPhishing && analysisResult.intent === "desviar_pago"
      ? [
          { code: "do_not_pay", text: "No transfieras dinero a la cuenta indicada hasta confirmar el cambio." },
          { code: "verify_via_official_channel", text: "Consultá al proveedor por un canal oficial que ya conozcas." },
        ]
    : analysisResult?.isPhishing
      ? [
          { code: "do_not_click", text: "No abras los enlaces ni adjuntos hasta verificar el mensaje." },
          { code: "verify_via_official_channel", text: "Consultá al remitente por un canal oficial iniciado por vos." },
        ]
      : [{ code: "verify_via_official_channel", text: "Si tenés dudas, comprobá el remitente y los enlaces por un canal oficial." }];

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
            <div className="flex items-center justify-between gap-2">
              <h2 className="text-lg font-semibold">Correo actual</h2>
              <button type="button" onClick={handleRefresh} className="text-xs font-semibold text-blue-700 hover:underline">
                Actualizar correo
              </button>
            </div>
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

          {/* Security Indicators Section */}
          {securityFeatures && securityFeatures.has_headers && (
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <h2 className="text-lg font-semibold">Indicadores de Seguridad</h2>

              {/* Auth badges */}
              <div className="mt-4 flex flex-wrap gap-2">
                <AuthBadge label="SPF" result={securityFeatures.spf_result} />
                <AuthBadge label="DKIM" result={securityFeatures.dkim_result} />
                <AuthBadge label="DMARC" result={securityFeatures.dmarc_result} />
              </div>

              {/* Spoofing alerts */}
              {!securityFeatures.from_return_path_match && (
                <div className="mt-3 rounded-lg bg-rose-50 px-3 py-2 text-xs text-rose-700">
                  ⚠️ Return-Path no coincide con el remitente
                </div>
              )}
              {!securityFeatures.from_reply_to_match && (
                <div className="mt-3 rounded-lg bg-rose-50 px-3 py-2 text-xs text-rose-700">
                  ⚠️ Reply-To difiere del remitente
                </div>
              )}
              {!securityFeatures.sender_vs_from_match && (
                <div className="mt-3 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-700">
                  ℹ️ Email enviado en nombre de otra persona
                </div>
              )}

              {/* Origin info */}
              <div className="mt-4 grid grid-cols-2 gap-2 text-xs text-slate-600">
                {securityFeatures.originating_country && (
                  <div>
                    <span className="font-semibold text-slate-400">País:</span>{" "}
                    {securityFeatures.originating_country}
                  </div>
                )}
                {securityFeatures.originating_ip && (
                  <div>
                    <span className="font-semibold text-slate-400">IP:</span>{" "}
                    {securityFeatures.originating_ip}
                  </div>
                )}
                <div>
                  <span className="font-semibold text-slate-400">Auth:</span>{" "}
                  {securityFeatures.auth_as}
                </div>
                <div>
                  <span className="font-semibold text-slate-400">Hops:</span>{" "}
                  {securityFeatures.received_hop_count}
                </div>
              </div>

              {/* Attachments */}
              {securityFeatures.attachment_count > 0 && (
                <div className={`mt-3 rounded-lg px-3 py-2 text-xs ${
                  securityFeatures.has_executable_attachment
                    ? "bg-rose-50 text-rose-700"
                    : "bg-slate-50 text-slate-600"
                }`}>
                  📎 {securityFeatures.attachment_count} adjunto{securityFeatures.attachment_count !== 1 ? "s" : ""}
                  {securityFeatures.has_executable_attachment && " · ⚠️ Contiene ejecutable"}
                </div>
              )}
            </div>
          )}

          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-lg font-semibold">Accion principal</h2>
            <p className="mt-2 text-sm text-slate-600">
              Analizá el correo abierto y revisá las señales que justifican el resultado.
            </p>
            <button
              className="mt-4 w-full rounded-xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
              type="button"
              onClick={handleAnalyze}
              disabled={!mailMetadata || isAnalyzing}
            >
              {isAnalyzing ? "Analizando..." : "Analizar Amenaza con IA"}
            </button>
            {isAnalyzing && (
              <button type="button" onClick={handleCancel} className="mt-3 w-full rounded-xl border border-slate-300 px-4 py-2 text-sm font-semibold text-slate-700">
                Cancelar análisis
              </button>
            )}
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-lg font-semibold">Reportar correo sospechoso</h2>
            <p className="mt-2 text-sm text-slate-600">
              Prepará un mensaje para el área de seguridad con el correo original adjunto. Outlook te pedirá revisar y enviar el reporte.
            </p>
            <label className="mt-4 block text-xs font-semibold text-slate-600" htmlFor="report-recipient">
              Correo del área de seguridad
            </label>
            <input
              id="report-recipient"
              type="email"
              value={reportRecipient}
              onChange={(event) => setReportRecipient(event.target.value)}
              placeholder="seguridad@tuempresa.com"
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm"
            />
            <button
              className="mt-3 w-full rounded-xl bg-blue-700 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-800 disabled:cursor-not-allowed disabled:bg-slate-400"
              type="button"
              onClick={handleReport}
              disabled={!mailMetadata}
            >
              Preparar reporte en Outlook
            </button>
            {reportStatus && <p role="status" className="mt-3 text-sm text-slate-700">{reportStatus}</p>}
          </div>

          {(isAnalyzing || analysisResult || analysisError) && (
            <div className="space-y-4">
              {isAnalyzing ? (
                <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                  <p className="text-sm text-slate-600">
                    Ejecutando modelos de riesgo...
                  </p>
                </div>
              ) : analysisError ? (
                <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                  <p className="text-sm text-rose-600">{analysisError}</p>
                </div>
              ) : (
                analysisResult && (
                  <>
                    {/* Veredicto principal */}
                    <div className={`rounded-2xl border p-5 shadow-sm ${
                      analysisResult.isPhishing
                        ? "border-rose-200 bg-rose-50"
                        : "border-emerald-200 bg-emerald-50"
                    }`}>
                      <div className="flex items-center gap-3">
                        <div className={`flex h-12 w-12 items-center justify-center rounded-full text-2xl ${
                          analysisResult.isPhishing ? "bg-rose-100" : "bg-emerald-100"
                        }`}>
                          {analysisResult.isPhishing ? "🚨" : "✅"}
                        </div>
                        <div className="flex-1">
                          <p className={`text-lg font-bold ${
                            analysisResult.isPhishing ? "text-rose-700" : "text-emerald-700"
                          }`}>
                            {analysisResult.isPhishing ? "PHISHING DETECTADO" : "CORREO LEGÍTIMO"}
                          </p>
                          <p className="text-xs text-slate-500">
                            Riesgo: {analysisResult.phishScore}%
                            {analysisResult.intent && ` · Intención: ${analysisResult.intent.replace(/_/g, " ")}`}
                          </p>
                        </div>
                        <RiskRing score={analysisResult.phishScore} isPhishing={analysisResult.isPhishing} />
                      </div>
                      {analysisResult.slmData && (
                        <p className="mt-3 rounded-lg bg-white/60 p-3 text-sm text-slate-700">
                          {analysisResult.slmData.summary}
                        </p>
                      )}
                    </div>

                    {analysisResult.contentSignals.length > 0 && (
                      <div className="rounded-2xl border border-rose-200 bg-rose-50 p-5 shadow-sm">
                        <h3 className="text-sm font-semibold text-rose-800">Señales detectadas en el mensaje</h3>
                        <ul className="mt-3 list-disc space-y-2 pl-5 text-sm text-slate-700">
                          {analysisResult.contentSignals.map((signal) => (
                            <li key={signal.rule}>{signal.description}</li>
                          ))}
                        </ul>
                        <p className="mt-3 text-sm font-medium text-rose-800">
                          No compartas claves ni realices el pago solicitado. Verificá la solicitud por un canal conocido, iniciado por vos.
                        </p>
                      </div>
                    )}

                    {/* Razones del SLM */}
                    {analysisResult.slmData && analysisResult.slmData.reasons.length > 0 && (
                      <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                        <h3 className="text-sm font-semibold text-slate-800">
                          📋 {analysisResult.isPhishing ? "¿Por qué es sospechoso?" : "Verificaciones realizadas"}
                        </h3>
                        {analysisResult.explanationSource && (
                          <p className="mt-1 text-xs text-slate-500">
                            {analysisResult.explanationSource === "slm"
                              ? "Explicación generada por el SLM"
                              : "Explicación segura basada en señales verificadas"}
                          </p>
                        )}
                        <div className="mt-3 space-y-2.5">
                          {analysisResult.slmData.reasons.map((reason, idx) => {
                            const icons: Record<string, string> = {
                              "slot.urgencia": "⚡",
                              "slot.autoridad": "🏛️",
                              "slot.financiero": "💰",
                              "slot.amenaza": "⚠️",
                              "slot.llamado_accion": "👆",
                              "slot.premio_falso": "🎁",
                              "slot.solicitud_datos": "🔑",
                              "security.url_check": "🔗",
                              "security.spf_fail": "⚙️",
                              "security.ml_classifier": "🤖",
                              "security.all_checks_pass": "✅",
                              "security.intent_detection": "🎯",
                            };
                            const icon = icons[reason.evidence_id] || (analysisResult.isPhishing ? "🔴" : "✅");
                            return (
                              <div key={idx} className="flex items-start gap-2.5">
                                <span className="mt-0.5 text-base">{icon}</span>
                                <div>
                                  <p className="text-sm text-slate-700">{reason.text}</p>
                                  <p className="text-xs text-slate-400">{reason.evidence_id.replace(/\./g, " · ")}</p>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* Mostrar una recomendación segura aunque falte la explicación estructurada. */}
                    {recommendedActions.length > 0 && (
                      <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                        <h3 className="text-sm font-semibold text-slate-800">
                          🛡️ {analysisResult.isPhishing ? "¿Qué deberías hacer?" : "Recomendación"}
                        </h3>
                        <div className="mt-3 space-y-2">
                          {recommendedActions.map((action, idx) => {
                            const actionStyles: Record<string, { bg: string; icon: string }> = {
                              do_not_click: { bg: "bg-rose-50 border-rose-200", icon: "🚫" },
                              do_not_pay: { bg: "bg-rose-50 border-rose-200", icon: "💳" },
                              do_not_reply: { bg: "bg-orange-50 border-orange-200", icon: "🔇" },
                              do_not_share_sensitive_information: { bg: "bg-amber-50 border-amber-200", icon: "🔒" },
                              report_message: { bg: "bg-blue-50 border-blue-200", icon: "📢" },
                              verify_via_official_channel: { bg: "bg-indigo-50 border-indigo-200", icon: "📞" },
                              no_action_needed: { bg: "bg-emerald-50 border-emerald-200", icon: "👍" },
                            };
                            const style = actionStyles[action.code] || { bg: "bg-slate-50 border-slate-200", icon: "ℹ️" };
                            return (
                              <div key={idx} className={`flex items-center gap-3 rounded-xl border p-3 ${style.bg}`}>
                                <span className="text-xl">{style.icon}</span>
                                <p className="text-sm text-slate-700">{action.text}</p>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* El servicio puede devolver una explicación de respaldo como texto plano. */}
                    {!analysisResult.slmData && (
                      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                        <h3 className="mb-3 text-sm font-semibold text-slate-800">
                          Resumen del análisis
                        </h3>
                        {analysisResult.plainServiceReasons?.length ? (
                          <div className={`space-y-4 rounded-xl border p-4 text-sm leading-6 ${analysisResult.isPhishing
                            ? "border-rose-100 bg-rose-50 text-rose-900"
                            : "border-emerald-100 bg-emerald-50 text-emerald-900"}`}>
                            <div>
                              <p className="mb-1 font-semibold">Conclusión</p>
                              <p>{analysisResult.plainServiceSummary}</p>
                            </div>
                            <div>
                              <p className="mb-1 font-semibold">Motivos principales</p>
                              <div className="space-y-2">
                                {analysisResult.plainServiceReasons.map((reason, index) => (
                                  <p key={`${index}-${reason}`}>{reason}</p>
                                ))}
                              </div>
                            </div>
                          </div>
                        ) : (
                          <p className="whitespace-pre-wrap rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm leading-6 text-slate-700">
                            {analysisResult.plainServiceExplanation || analysisResult.explanation}
                          </p>
                        )}
                        <div className="mt-4">
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
                    )}
                  </>
                )
              )}
            </div>
          )}
          <div className="pt-4 border-t border-slate-200 flex justify-center">
            <button
              type="button"
              onClick={handleUnlink}
              className="text-xs font-semibold text-slate-500 hover:text-rose-600 transition"
            >
              Desvincular cuenta de Outlook
            </button>
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

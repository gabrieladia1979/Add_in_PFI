import React, { useEffect, useId, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { getCurrentMessageMetadata, MailMetadata, getMailboxItem, prepareIncidentReport } from "../services/mailService";
import { analyzeEmail } from "../services/analysisService";
import { buildAnalysisResult, AnalysisResult } from "../services/analysisResult";
import { extractSecurityFeatures } from "../services/securityService";
import { SecurityFeatures } from "../types/securityTypes";
import { RiskRing } from "./RiskRing";
import "./taskpane.css";

const ANALYSIS_TIMEOUT_MS = 30_000;

// Service explanations may contain decorative emoji. Keep the message text intact.
function explanationText(value: string | null | undefined): string {
  return (value || "").replace(/[\p{Extended_Pictographic}\uFE0F\u200D]/gu, "").replace(/^[ \t]+/gm, "").trim();
}

function Chevron(): JSX.Element {
  return <svg className="pa-chevron" aria-hidden="true" viewBox="0 0 20 20" fill="none"><path d="m6 8 4 4 4-4" /></svg>;
}

function StatusMark({ warning }: { warning: boolean }): JSX.Element {
  return <svg className="pa-status-mark" aria-hidden="true" viewBox="0 0 24 24" fill="none">
    <path d="M12 3 4.5 6v5.5c0 4.2 3.1 7.6 7.5 9.5 4.4-1.9 7.5-5.3 7.5-9.5V6L12 3Z" />
    {warning ? <path d="M12 8v5m0 3v.01" /> : <path d="m8.5 12 2.3 2.3 4.7-4.7" />}
  </svg>;
}

function DomainAuthCheck({ features }: { features: SecurityFeatures | null }): JSX.Element | null {
  const [expanded, setExpanded] = useState(false);
  const detailId = useId();

  if (!features) {
    return null;
  }

  const checks: { name: string; result: string }[] = [
    { name: "SPF", result: features.spf_result },
    { name: "DKIM", result: features.dkim_result },
    { name: "DMARC", result: features.dmarc_result },
  ];

  const hasHeaders = features.has_headers;
  const passed = checks.filter(({ result }) => result === "pass" || result === "bestguesspass").length;
  const unknown = checks.filter(({ result }) => !result || result === "unknown" || result === "none").length;

  const statusText = !hasHeaders
    ? "No disponible"
    : unknown === 3
    ? "Sin registros"
    : `${passed}/3 verificados`;

  return (
    <div className={`pa-auth-check${expanded ? " pa-auth-check--expanded" : ""}`}>
      <button
        type="button"
        className="pa-auth-check-trigger"
        aria-controls={detailId}
        aria-expanded={expanded}
        onClick={() => setExpanded((curr) => !curr)}
      >
        <div className="pa-auth-check-header">
          <svg className="pa-auth-check-icon" aria-hidden="true" viewBox="0 0 24 24" fill="none">
            <path d="M12 3 4.5 6v5.5c0 4.2 3.1 7.6 7.5 9.5 4.4-1.9 7.5-5.3 7.5-9.5V6L12 3Z" strokeLinecap="round" strokeLinejoin="round" />
            <path d="m9 12 2 2 4-4" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
          <span className="pa-auth-check-title">Autenticación técnica</span>
        </div>
        <div className="pa-auth-check-summary">
          <span className="pa-auth-check-count">{statusText}</span>
          <div className="pa-auth-check-meter" aria-hidden="true">
            {checks.map(({ result }, idx) => {
              const segClass = !hasHeaders || !result || result === "unknown" || result === "none"
                ? ""
                : result === "pass" || result === "bestguesspass"
                ? "pa-auth-check-segment--pass"
                : result === "softfail" || result === "neutral"
                ? "pa-auth-check-segment--amber"
                : "pa-auth-check-segment--fail";
              return <span key={idx} className={`pa-auth-check-segment ${segClass}`} />;
            })}
          </div>
          <svg className={`pa-chevron transition-transform ${expanded ? "rotate-180" : ""}`} aria-hidden="true" viewBox="0 0 20 20" fill="none">
            <path d="m6 8 4 4 4-4" />
          </svg>
        </div>
      </button>

      <div id={detailId} className="pa-auth-check-detail">
        <p className="pa-auth-check-caption">
          {hasHeaders
            ? "Comprobaciones de autenticación en las cabeceras del correo:"
            : "No se pudieron obtener las cabeceras de transporte del correo en este entorno."}
        </p>
        <ul className="pa-auth-check-list">
          {checks.map(({ name, result }) => {
            let label = "No consta";
            let badgeClass = "pa-auth-badge--neutral";

            if (result === "pass" || result === "bestguesspass") {
              label = "Aprobado";
              badgeClass = "pa-auth-badge--pass";
            } else if (result === "softfail") {
              label = "Fallo parcial";
              badgeClass = "pa-auth-badge--amber";
            } else if (result === "neutral") {
              label = "Neutro";
              badgeClass = "pa-auth-badge--amber";
            } else if (result === "fail") {
              label = "Falló";
              badgeClass = "pa-auth-badge--fail";
            } else if (result === "none") {
              label = "Sin política";
              badgeClass = "pa-auth-badge--neutral";
            }

            return (
              <li key={name} className="pa-auth-check-item">
                <span className="pa-auth-check-name">{name}</span>
                <span className={`pa-auth-badge ${badgeClass}`}>{label}</span>
              </li>
            );
          })}
        </ul>
        {hasHeaders && features.from_return_path_match === false && (
          <div className="mt-2.5 flex items-center justify-between text-xs p-2 rounded bg-amber-50 text-amber-900 border border-amber-200">
            <span>Discrepancia en Return-Path</span>
            <span className="pa-auth-badge pa-auth-badge--amber">Revisar</span>
          </div>
        )}
      </div>
    </div>
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
  const [analysisSecurityFeatures, setAnalysisSecurityFeatures] = useState<SecurityFeatures | null>(null);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [reportRecipient, setReportRecipient] = useState("");
  const [reportStatus, setReportStatus] = useState<string | null>(null);
  const [itemVersion, setItemVersion] = useState(0);
  const analysisRequestId = useRef(0);
  const activeItemVersion = useRef(0);
  const activeAnalysisController = useRef<AbortController | null>(null);
  const resultRef = useRef<HTMLDivElement | null>(null);
  const mailOverviewRef = useRef<HTMLDetailsElement | null>(null);

  useEffect(() => () => activeAnalysisController.current?.abort(), []);

  useEffect(() => {
    if (analysisResult) {
      resultRef.current?.focus({ preventScroll: true });
      resultRef.current?.scrollIntoView({ block: "start" });
    }
  }, [analysisResult]);

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
      activeItemVersion.current += 1;
      analysisRequestId.current += 1;
      activeAnalysisController.current?.abort();
      setAnalysisResult(null);
      setAnalysisSecurityFeatures(null);
      setAnalysisError(null);
      setIsAnalyzing(false);
      setReportStatus(null);
      setMailMetadata(null);
      setSecurityFeatures(null);
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
      const version = activeItemVersion.current;
      extractSecurityFeatures(item)
        .then((features) => {
          if (!cancelled && version === activeItemVersion.current) {
            setSecurityFeatures(features);
          }
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
    setAnalysisSecurityFeatures(null);
    setLoadError(null);
    setAnalysisError(null);
    setIsAnalyzing(false);
    setReportRecipient("");
    setReportStatus(null);
  };

  const handleBackToMail = (): void => {
    setAnalysisResult(null);
    setAnalysisSecurityFeatures(null);
    setAnalysisError(null);
    setReportStatus(null);
    window.requestAnimationFrame(() => {
      mailOverviewRef.current?.querySelector("summary")?.focus({ preventScroll: true });
      mailOverviewRef.current?.scrollIntoView({ block: "start" });
    });
  };

  const handleCancel = (): void => {
    analysisRequestId.current += 1;
    activeAnalysisController.current?.abort();
    setIsAnalyzing(false);
    setAnalysisError("Análisis cancelado. Podés volver a intentarlo.");
    setAnalysisSecurityFeatures(null);
  };

  const handleAnalyze = async (): Promise<void> => {
    if (!mailMetadata) {
      return;
    }

    const requestId = ++analysisRequestId.current;
    const controller = new AbortController();
    activeAnalysisController.current = controller;
    const timeoutId = window.setTimeout(() => controller.abort(), ANALYSIS_TIMEOUT_MS);
    const versionAtAnalysis = activeItemVersion.current;
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setAnalysisSecurityFeatures(null);
    setAnalysisError(null);

    try {
      if (requestId !== analysisRequestId.current || versionAtAnalysis !== activeItemVersion.current || controller.signal.aborted) return;
      const featuresAtAnalysis = securityFeatures ? { ...securityFeatures } : null;
      setAnalysisSecurityFeatures(featuresAtAnalysis);
      const response = await analyzeEmail(mailMetadata.subject, mailMetadata.body, mailMetadata.senderEmail, featuresAtAnalysis, controller.signal);
      if (requestId === analysisRequestId.current) {
        setAnalysisResult(buildAnalysisResult(response));
      }
    } catch (error) {
      if (requestId === analysisRequestId.current) {
        setAnalysisError(error instanceof Error && error.name === "AbortError"
          ? "El análisis superó los 30 segundos. Podés volver a intentarlo."
          : "No se pudo analizar el correo.");
        setAnalysisSecurityFeatures(null);
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
      setReportStatus("Borrador abierto en Outlook.");
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

  // Present each explanation once when the service repeats a detected signal.
  const explanationKey = (text: string): string => explanationText(text).replace(/\s+/g, " ").toLocaleLowerCase("es");
  const seenExplanationKeys = new Set((analysisResult?.contentSignals || []).map((signal) => explanationKey(signal.description)));
  const visibleReasons = (analysisResult?.slmData?.reasons || []).filter((reason) => {
    const key = explanationKey(reason.text);
    if (!key || seenExplanationKeys.has(key)) return false;
    seenExplanationKeys.add(key);
    return true;
  });

  const overviewFeatures = analysisResult ? analysisSecurityFeatures : securityFeatures;
  const mailOverview = (
    <details
      key={`mail-${itemVersion}-${analysisResult ? "result" : isAnalyzing ? "analyzing" : "preview"}`}
      ref={mailOverviewRef}
      open={!analysisResult && !isAnalyzing}
      className="pa-disclosure"
    >
      <summary className="pa-disclosure-summary">
        <div className="min-w-0 flex-1">
          <h2 className={analysisResult ? "text-sm font-semibold text-slate-900" : "text-xs font-semibold text-slate-500"}>
            {analysisResult ? "Correo analizado" : "Correo actual"}
          </h2>
          {!analysisResult && (
            <>
              <p
                role={loadError ? "alert" : undefined}
                className={`mt-1 truncate text-sm font-semibold ${loadError ? "text-rose-600" : "text-slate-900"}`}
                title={mailMetadata?.subject}
              >
                {loadError || mailMetadata?.subject || "Cargando correo..."}
              </p>
              {mailMetadata && <p className="mt-1 truncate text-xs text-slate-500">{mailMetadata.sender}</p>}
            </>
          )}
        </div>
        <Chevron />
      </summary>
      <div className="pa-disclosure-body space-y-4">
        {mailMetadata && (
          <dl className="space-y-3 text-sm text-slate-700">
            <div>
              <dt className="text-xs font-semibold text-slate-500">Remitente</dt>
              <dd className="mt-1 break-words">{mailMetadata.senderEmail || mailMetadata.sender}</dd>
            </div>
            <div>
              <dt className="text-xs font-semibold text-slate-500">Asunto</dt>
              <dd className="mt-1 break-words">{mailMetadata.subject}</dd>
            </div>
            <div>
              <dt className="text-xs font-semibold text-slate-500">Cuerpo (extracto)</dt>
              <dd className="mt-1 break-words">{mailMetadata.bodyPreview}</dd>
            </div>
          </dl>
        )}
        {/* Security Indicators & Technical Authentication (only shown after analysis in Correo analizado) */}
        {Boolean(analysisResult) && (
          <>
            {overviewFeatures?.has_headers && (
              <div className="border-t border-slate-100 pt-4">
                <h2 className="text-sm font-semibold">Indicadores de seguridad</h2>

                {/* Spoofing alerts */}
                {!overviewFeatures.from_return_path_match && (
                  <div className="mt-3 rounded-lg bg-rose-50 px-3 py-2 text-xs text-rose-700">
                    Return-Path no coincide con el remitente
                  </div>
                )}
                {!overviewFeatures.from_reply_to_match && (
                  <div className="mt-3 rounded-lg bg-rose-50 px-3 py-2 text-xs text-rose-700">
                    Reply-To difiere del remitente
                  </div>
                )}
                {!overviewFeatures.sender_vs_from_match && (
                  <div className="mt-3 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-700">
                    Email enviado en nombre de otra persona
                  </div>
                )}

                {/* Origin info */}
                <div className="mt-4 grid grid-cols-2 gap-2 text-xs text-slate-600">
                  {overviewFeatures.originating_country && (
                    <div>
                      <span className="font-semibold text-slate-400">País:</span>{" "}
                      {overviewFeatures.originating_country}
                    </div>
                  )}
                  {overviewFeatures.originating_ip && (
                    <div>
                      <span className="font-semibold text-slate-400">IP:</span>{" "}
                      {overviewFeatures.originating_ip}
                    </div>
                  )}
                  <div>
                    <span className="font-semibold text-slate-400">Hops:</span>{" "}
                    {overviewFeatures.received_hop_count}
                  </div>
                </div>

                {/* Attachments */}
                {overviewFeatures.attachment_count > 0 && (
                  <div className={`mt-3 rounded-lg px-3 py-2 text-xs ${
                    overviewFeatures.has_executable_attachment
                      ? "bg-rose-50 text-rose-700"
                      : "bg-slate-50 text-slate-600"
                  }`}>
                    {overviewFeatures.attachment_count} adjunto{overviewFeatures.attachment_count !== 1 ? "s" : ""}
                    {overviewFeatures.has_executable_attachment && " · Contiene ejecutable"}
                  </div>
                )}
              </div>
            )}
            <DomainAuthCheck features={overviewFeatures} />
          </>
        )}
      </div>
    </details>
  );

  if (!officeReady) {
    return (
      <div className="min-h-screen bg-slate-50 px-6 py-8 text-slate-900">
        <p className="text-sm text-slate-600">Cargando Outlook...</p>
      </div>
    );
  }

  return (
    <div className={isLinked ? `pa-panel${analysisResult?.isPhishing ? " pa-panel--warning" : ""}` : "min-h-screen bg-slate-50 px-6 py-8 text-slate-900"}>
      {!isLinked && <header className="mb-8">
        <p className="text-xs font-semibold tracking-wide text-emerald-600">
          PhishARG
        </p>
        <h1 className="mt-2 text-2xl font-semibold">Panel de Defensa Proactiva</h1>
        <p className="mt-2 text-sm text-slate-600">
          Educa en tiempo real y explica con XAI por que un correo es sospechoso.
        </p>
      </header>}

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
        <section className="pa-workspace">
          {!analysisResult && mailOverview}

          {!analysisResult && (
            <div className="pa-analysis-start">
              <h2 className="pa-section-title">Analizar correo</h2>
              <button
                className="pa-button pa-button--primary"
                type="button"
                onClick={handleAnalyze}
                disabled={!mailMetadata || isAnalyzing}
              >
                {isAnalyzing ? "Analizando..." : "Analizar amenaza con IA"}
              </button>
              {isAnalyzing && (
                <button type="button" onClick={handleCancel} className="pa-button pa-button--secondary">
                  Cancelar análisis
                </button>
              )}
            </div>
          )}

          {(isAnalyzing || analysisResult || analysisError) && (
            <div ref={resultRef} tabIndex={-1} role="region" aria-label="Resultado del análisis" className={`pa-result ${analysisResult ? "pa-result--ready" : ""}`}>
              {isAnalyzing ? (
                <div className="pa-feedback" role="status">
                  <span className="pa-spinner" aria-hidden="true" />
                  <p>Analizando el correo. Esperá un momento.</p>
                </div>
              ) : analysisError ? (
                <div className="pa-feedback pa-feedback--error" role="alert">
                  <p>{analysisError}</p>
                </div>
              ) : (
                analysisResult && (
                  <>
                    <button type="button" onClick={handleBackToMail} className="pa-text-button mb-3 inline-flex items-center gap-2">
                      <svg aria-hidden="true" width="16" height="16" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                        <path d="m8 4-6 6 6 6M2 10h15" />
                      </svg>
                      Volver al correo
                    </button>
                    <div className={`pa-verdict ${analysisResult.isPhishing ? "pa-verdict--warning" : "pa-verdict--safe"}`}>
                      <div className="pa-verdict-header">
                        <StatusMark warning={analysisResult.isPhishing} />
                        <div className="pa-verdict-heading">
                          <h1>{analysisResult.isPhishing ? "Phishing detectado" : "Correo legítimo"}</h1>
                          {analysisResult.intent && <p>Intención: {analysisResult.intent.replace(/_/g, " ")}</p>}
                        </div>
                        <RiskRing score={analysisResult.phishScore} isPhishing={analysisResult.isPhishing} />
                      </div>
                      {analysisResult.slmData && <p className="pa-verdict-summary">{explanationText(analysisResult.slmData.summary)}</p>}
                    </div>

                    {analysisResult.contentSignals.length > 0 && (
                      <div className="pa-section pa-signals">
                        <h3 className="pa-section-title">Señales detectadas en el mensaje</h3>
                        <ul className="mt-3 list-disc space-y-2 pl-5 text-sm text-slate-700">
                          {analysisResult.contentSignals.map((signal) => (
                            <li key={signal.rule}>{explanationText(signal.description)}</li>
                          ))}
                        </ul>
                        <p className="pa-guidance">
                          No compartas claves ni realices el pago solicitado. Verificá la solicitud por un canal conocido, iniciado por vos.
                        </p>
                      </div>
                    )}

                    {visibleReasons.length > 0 && (
                      <section className="pa-section">
                        <h2 className="pa-section-title">{analysisResult.isPhishing ? "¿Por qué es sospechoso?" : "Verificaciones realizadas"}</h2>
                        <ul className="pa-explanations">
                          {visibleReasons.map((reason, idx) => <li key={idx}>{explanationText(reason.text)}</li>)}
                        </ul>
                      </section>
                    )}

                    {recommendedActions.length > 0 && (
                      <section className="pa-section pa-recommendations">
                        <h2 className="pa-section-title">{analysisResult.isPhishing ? "¿Qué deberías hacer?" : "Recomendación"}</h2>
                        <ul className="pa-actions">
                          {recommendedActions.map((action, idx) => <li key={idx}>{explanationText(action.text)}</li>)}
                        </ul>
                      </section>
                    )}

                    {/* El servicio puede devolver una explicación de respaldo como texto plano. */}
                    {!analysisResult.slmData && (
                      <div className="pa-section">
                        <h3 className="pa-section-title">
                          Resumen del análisis
                        </h3>
                        {analysisResult.plainServiceReasons?.length ? (
                          <div className="mt-3 space-y-4 text-sm leading-6">
                            <div>
                              <p className="mb-1 font-semibold">Conclusión</p>
                              <p>{explanationText(analysisResult.plainServiceSummary)}</p>
                            </div>
                            <div>
                              <p className="mb-1 font-semibold">Motivos principales</p>
                              <div className="space-y-2">
                                {analysisResult.plainServiceReasons.map((reason, index) => (
                                  <p key={`${index}-${reason}`}>{explanationText(reason)}</p>
                                ))}
                              </div>
                            </div>
                          </div>
                        ) : (
                          <p className="mt-3 whitespace-pre-wrap text-sm leading-6">
                            {explanationText(analysisResult.plainServiceExplanation || analysisResult.explanation)}
                          </p>
                        )}

                      </div>
                    )}
                  </>
                )
              )}
            </div>
          )}
          {analysisResult && mailOverview}

          <details key={`report-${itemVersion}-${analysisResult ? "result" : "preview"}`} className="pa-disclosure">
            <summary className="pa-disclosure-summary">
              <h2>Reportar correo sospechoso</h2>
              <Chevron />
            </summary>
            <div className="pa-disclosure-body">
              <label className="mt-4 block text-xs font-semibold text-slate-600" htmlFor="report-recipient">
                Destinatario
              </label>
              <input
                id="report-recipient"
                type="email"
                value={reportRecipient}
                onChange={(event) => setReportRecipient(event.target.value)}
                placeholder="seguridad@tuempresa.com"
                className="pa-input"
              />
              <button
                className="pa-button pa-button--secondary"
                type="button"
                onClick={handleReport}
                disabled={!mailMetadata}
              >
                Preparar reporte en Outlook
              </button>
              {reportStatus && <p role="status" className="mt-3 text-sm text-slate-700">{reportStatus}</p>}
            </div>
          </details>

          {!analysisResult && <div className="pa-footer">
            <button
              type="button"
              onClick={handleUnlink}
              className="pa-text-button pa-exit"
            >
              Desvincular cuenta de Outlook
            </button>
          </div>}
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

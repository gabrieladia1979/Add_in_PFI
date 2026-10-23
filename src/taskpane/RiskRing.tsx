import React from "react";

type RiskRingProps = {
  score: number;
  isPhishing: boolean;
};

export function RiskRing({ score, isPhishing }: RiskRingProps): JSX.Element {
  const percentage = Number.isFinite(score) ? Math.min(100, Math.max(0, Math.round(score))) : 0;
  return (
    <div role="img" aria-label={`Riesgo estimado: ${percentage}%`} className={`pa-risk ${isPhishing ? "pa-risk--warning" : "pa-risk--safe"}`}>
      <svg aria-hidden="true" viewBox="0 0 56 56">
        <circle cx="28" cy="28" r="24" fill="none" strokeWidth="3" className="pa-risk-track" />
        <circle cx="28" cy="28" r="24" fill="none" strokeWidth="3" pathLength="100" strokeDasharray={`${percentage} ${100 - percentage}`} className="pa-risk-value" />
      </svg>
      <span aria-hidden="true" className="pa-risk-number">{percentage}%</span>
    </div>
  );
}

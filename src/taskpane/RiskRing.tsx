import React from "react";

type RiskRingProps = {
  score: number;
  isPhishing: boolean;
};

export function RiskRing({ score, isPhishing }: RiskRingProps): JSX.Element {
  const percentage = Number.isFinite(score) ? Math.min(100, Math.max(0, Math.round(score))) : 0;

  return (
    <div
      role="img"
      aria-label={`Riesgo ${percentage}%`}
      className="relative h-14 w-14 shrink-0"
    >
      <svg aria-hidden="true" viewBox="0 0 56 56" className="h-full w-full -rotate-90">
        <circle
          cx="28"
          cy="28"
          r="24"
          fill="none"
          strokeWidth="4"
          className={isPhishing ? "stroke-rose-200" : "stroke-emerald-200"}
        />
        <circle
          cx="28"
          cy="28"
          r="24"
          fill="none"
          strokeWidth="4"
          pathLength="100"
          strokeDasharray={`${percentage} ${100 - percentage}`}
          className={isPhishing ? "stroke-rose-500" : "stroke-emerald-500"}
        />
      </svg>
      <span
        aria-hidden="true"
        className={`absolute inset-0 flex items-center justify-center text-sm font-bold ${
          isPhishing ? "text-rose-600" : "text-emerald-600"
        }`}
      >
        {percentage}%
      </span>
    </div>
  );
}

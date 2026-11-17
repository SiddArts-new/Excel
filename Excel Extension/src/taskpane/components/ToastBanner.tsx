import React from "react";
import { CheckCircle2, Info, AlertTriangle, XCircle, X } from "lucide-react";
import type { ToastMessage } from "../types/excel";

interface ToastBannerProps {
  toast: ToastMessage | null;
  onDismiss: () => void;
}

const STYLE_MAP: Record<ToastMessage["type"], string> = {
  success: "border-emerald-200 bg-emerald-50 text-emerald-900",
  info: "border-sky-200 bg-sky-50 text-sky-900",
  warning: "border-amber-200 bg-amber-50 text-amber-900",
  error: "border-rose-200 bg-rose-50 text-rose-900"
};

export const ToastBanner: React.FC<ToastBannerProps> = ({ toast, onDismiss }) => {
  if (!toast) {
    return null;
  }

  const Icon =
    toast.type === "success"
      ? CheckCircle2
      : toast.type === "warning"
      ? AlertTriangle
      : toast.type === "error"
      ? XCircle
      : Info;

  return (
    <div
      role="status"
      className={`mx-3 mt-2 flex items-start justify-between gap-2 rounded-md border px-2.5 py-2 text-xs shadow-sm ${STYLE_MAP[toast.type]}`}
    >
      <div className="flex min-w-0 items-start gap-2">
        <Icon className="mt-0.5 h-3.5 w-3.5 shrink-0" />
        <p className="break-words leading-snug">{toast.message}</p>
      </div>
      <button
        type="button"
        onClick={onDismiss}
        aria-label="Dismiss notification"
        className="shrink-0 rounded p-0.5 opacity-70 transition hover:opacity-100"
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
};

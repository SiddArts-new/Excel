import React from "react";

import { AlertTriangle, X } from "lucide-react";

interface ConfirmModalProps {
  isOpen: boolean;
  title: string;
  description: string;
  confirmLabel?: string;
  isDanger?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export const ConfirmModal: React.FC<ConfirmModalProps> = ({
  isOpen,
  title,
  description,
  confirmLabel = "Confirm",
  isDanger = false,
  onConfirm,
  onCancel
}) => {
  if (!isOpen) {
    return null;
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-3">
      <div
        role="dialog"
        aria-modal="true"
        className="w-full max-w-xs rounded-lg border border-slate-200 bg-white p-3.5 shadow-xl"
      >
        <div className="flex items-start justify-between gap-2">
          <div className="flex items-center gap-2">
            <AlertTriangle
              className={`h-4 w-4 shrink-0 ${
                isDanger ? "text-rose-600" : "text-amber-600"
              }`}
            />
            <h2 className="text-sm font-semibold text-slate-900">{title}</h2>
          </div>
          <button
            type="button"
            onClick={onCancel}
            aria-label="Close modal"
            className="rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <p className="mt-2 text-xs leading-relaxed text-slate-600">
          {description}
        </p>

        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            onClick={onCancel}
            className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={onConfirm}
            className={`rounded-md px-2.5 py-1.5 text-xs font-semibold text-white ${
              isDanger
                ? "bg-rose-600 hover:bg-rose-700"
                : "bg-emerald-600 hover:bg-emerald-700"
            }`}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};

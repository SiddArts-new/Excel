import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import "./index.css";

let root: ReactDOM.Root | null = null;

function renderApp(isOfficeReady: boolean): void {
  const rootEl = document.getElementById("root");
  if (!rootEl) {
    return;
  }

  if (!root) {
    root = ReactDOM.createRoot(rootEl);
  }
  root.render(
    <React.StrictMode>
      <App isOfficeReady={isOfficeReady} />
    </React.StrictMode>
  );
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    const isExcel = info.host === Office.HostType.Excel;
    renderApp(isExcel);
  });
} else {
  renderApp(false);
}

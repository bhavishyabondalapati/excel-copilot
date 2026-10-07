// Entry point. Decides whether we're inside Excel or a plain browser tab.

import { createRoot } from "react-dom/client";
import { App } from "./App";
import { createDemoWorkbook } from "../demo/demoData";
import { ExcelGateway } from "../excel/excelGateway";
import "./styles.css";

function start(inExcel: boolean) {
  const root = createRoot(document.getElementById("root")!);
  if (inExcel) {
    root.render(<App workbook={new ExcelGateway()} />);
  } else {
    // Outside Excel: use an in-memory workbook so the UI can be tried in a browser.
    const demo = createDemoWorkbook();
    root.render(<App workbook={demo} demoWorkbook={demo} />);
  }
}

// office.js is loaded by index.html from Microsoft's CDN. Outside Office,
// Office.onReady still fires, with host === null.
if (typeof Office !== "undefined") {
  Office.onReady((info) => start(info.host === Office.HostType.Excel));
} else {
  start(false);
}

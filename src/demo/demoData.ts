// Sample workbook used when the task pane is opened in a normal browser
// (outside Excel), so the UI can be tried without Office or API keys.

import { MemoryWorkbook } from "../core/memoryWorkbook";

export function createDemoWorkbook(): MemoryWorkbook {
  return new MemoryWorkbook({
    sheets: {
      Sales: {
        A1: "Region", B1: "Units", C1: "Unit price", D1: "Revenue",
        A2: "North", B2: 120, C2: 9.5,
        A3: "South", B3: 95, C3: 10.25,
        A4: "East", B4: 143, C4: 8.75,
        A5: "West", B5: 77, C5: 11,
        A7: "Average price", C7: "=AVERAGE(C2:C5)",
      },
      Notes: { A1: "Scratch space" },
    },
    activeSheet: "Sales",
    selection: "C7",
  });
}

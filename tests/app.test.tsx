// @vitest-environment jsdom
// UI smoke test: drives the real React app against the in-memory demo workbook.
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { App } from "../src/taskpane/App";
import { createDemoWorkbook } from "../src/demo/demoData";

const SETTINGS_KEY = "excel-copilot.settings.v1";

// Gemini is the default provider; these tests pick Demo so no key or network is needed.
beforeEach(() => localStorage.setItem(SETTINGS_KEY, JSON.stringify({ provider: "demo" })));
afterEach(() => {
  cleanup();
  localStorage.clear();
});

describe("App defaults", () => {
  it("uses Gemini and asks for a key until one is saved", () => {
    localStorage.clear();
    const wb = createDemoWorkbook();
    render(<App workbook={wb} demoWorkbook={wb} />);
    expect(screen.getByText(/Gemini · gemini-/)).toBeTruthy();
    expect(screen.getByText("Add a Gemini API key in Settings.")).toBeTruthy();
    expect((screen.getByRole("button", { name: "Preview changes" }) as HTMLButtonElement).disabled).toBe(true);
  });
});

async function click(name: string | RegExp) {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
}

describe("App (demo mode)", () => {
  it("previews, applies and undoes a change", async () => {
    const wb = createDemoWorkbook();
    render(<App workbook={wb} demoWorkbook={wb} />);

    fireEvent.change(screen.getByLabelText("What should I do?"), { target: { value: "Fill in the Revenue column" } });
    await click("Preview changes");

    const table = await screen.findByRole("table", { name: "Cells that will change" });
    expect(table.textContent).toContain("=B2*C2");
    expect(wb.get("Sales", "D2").formula).toBeNull();

    await click("Apply");
    expect(wb.get("Sales", "D5").formula).toBe("=B5*C5");
    expect(screen.getByText(/Revenue = Units × Unit price/)).toBeTruthy();

    await click("Undo");
    expect(wb.get("Sales", "D5").formula).toBeNull();
    expect(screen.queryByRole("button", { name: "Undo" })).toBeNull();
  });

  it("explains the selected formula", async () => {
    const wb = createDemoWorkbook(); // C7 (=AVERAGE) is selected
    render(<App workbook={wb} demoWorkbook={wb} />);
    await click("Explain selected formula");
    expect((await screen.findByText(/About Sales!C7/)).textContent).toBeTruthy();
    expect(screen.getByText(/uses AVERAGE/)).toBeTruthy();
  });

  it("shows an error when the selected cell has no formula", async () => {
    const wb = createDemoWorkbook();
    wb.selection = "A1";
    render(<App workbook={wb} demoWorkbook={wb} />);
    await click("Explain selected formula");
    expect((await screen.findByRole("alert")).textContent).toContain("doesn't contain a formula");
  });
});

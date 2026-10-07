import { XlsxParser } from "../../../../dist/index.js";
import { base64ToArrayBuffer, sampleWorkbookBase64 } from "./sampleWorkbook";
import "./demo.css";
import { setError, setStatus } from './status';

type ParserOptions = {
  dense: boolean;
  styles: boolean;
  drawings: boolean;
  skipHiddenRows: boolean;
};

type DemoMode = "sheet" | "all";

type StoryArgs = ParserOptions & {
  mode: DemoMode;
};

type WorkbookLike = {
  workSheets: Array<{
    name: string;
    state?: string;
    data: unknown[];
    mergeCells: string[];
  }>;
};

const parser = new XlsxParser();

function createOption(
  key: keyof ParserOptions,
  label: string,
  description: string,
  options: ParserOptions,
): HTMLLabelElement {
  const row = document.createElement("label");
  row.className = "sb-demo__option";

  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = options[key];
  input.dataset.option = key;

  const text = document.createElement("div");
  const title = document.createElement("strong");
  title.textContent = label;
  const body = document.createElement("span");
  body.textContent = description;
  text.append(title, body);

  row.append(input, text);
  return row;
}

function createDownloadUrl(): string {
  const bytes = Uint8Array.from(atob(sampleWorkbookBase64), (char) => char.charCodeAt(0));
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return URL.createObjectURL(blob);
}

export function createXlsxDemo(args: StoryArgs): HTMLElement {
  const root = document.createElement("div");
  root.className = "sb-demo";

  const options: ParserOptions = {
    dense: args.dense,
    styles: args.styles,
    drawings: args.drawings,
    skipHiddenRows: args.skipHiddenRows,
  };

  let currentWorkbook: WorkbookLike | null = null;
  let currentSheetIndex = 0;
  let currentUrl = createDownloadUrl();

  root.innerHTML = `
    <div class="sb-demo__shell">
      <div class="sb-demo__grid">
        <aside class="sb-demo__panel">
          <section class="sb-demo__section">
            <h2>Workbook source</h2>
            <div class="sb-demo__actions">
              <button class="sb-demo__button" type="button" data-action="sample">Load sample workbook</button>
              <a class="sb-demo__download" data-action="download" download="storybook-sample.xlsx" href="${currentUrl}">
                Download sample workbook
              </a>
            </div>
            <input class="sb-demo__input" data-action="upload" type="file" accept=".xlsx" />
          </section>

          <section class="sb-demo__section">
            <h2>Parser options</h2>
            <div class="sb-demo__options"></div>
          </section>

          <section class="sb-demo__section">
            <h2>Status</h2>
            <div class="sb-demo__status" data-role="status"></div>
          </section>
        </aside>

        <section class="sb-demo__viewer">
          <div class="sb-demo__toolbar">
            <div>
              <h2>${args.mode === "sheet" ? "Single-sheet rendering" : "Full workbook HTML"}</h2>
              <div class="sb-demo__toolbar-copy">
                ${
                  args.mode === "sheet"
                    ? "Switch sheets and inspect the rendered output."
                    : "Render every sheet in one continuous HTML output."
                }
              </div>
            </div>
            <div class="sb-demo__meta" data-role="meta"></div>
          </div>
          <div class="sb-demo__tabs" data-role="tabs"></div>
          <div class="sb-demo__canvas ${args.mode === "all" ? "sb-demo__full-html" : ""}" data-role="canvas">
            <div class="sb-demo__empty">
              <div>
                <strong>No workbook loaded</strong>
                <span>Load the embedded sample or upload your own file to begin.</span>
              </div>
            </div>
          </div>
        </section>
      </div>
    </div>
  `;

  const optionsHost = root.querySelector(".sb-demo__options") as HTMLDivElement;
  const status = root.querySelector('[data-role="status"]') as HTMLDivElement;
  const meta = root.querySelector('[data-role="meta"]') as HTMLDivElement;
  const tabs = root.querySelector('[data-role="tabs"]') as HTMLDivElement;
  const canvas = root.querySelector('[data-role="canvas"]') as HTMLDivElement;
  const sampleButton = root.querySelector('[data-action="sample"]') as HTMLButtonElement;
  const uploadInput = root.querySelector('[data-action="upload"]') as HTMLInputElement;
  const downloadLink = root.querySelector('[data-action="download"]') as HTMLAnchorElement;

  [
    createOption("dense", "dense", "Keeps the cell matrix compact when the workbook contains empty cells.", options),
    createOption("styles", "styles", "Enables parsing for styles and colors when the workbook includes them.", options),
    createOption("drawings", "drawings", "Enables parsing for images and drawing objects when they exist in the file.", options),
    createOption(
      "skipHiddenRows",
      "skipHiddenRows",
      "Skips hidden rows when that metadata is present in the workbook.",
      options,
    ),
  ].forEach((item) => optionsHost.append(item));

  setStatus(
    status,
    "Waiting for a workbook",
    "The demo is ready. The embedded sample includes two sheets and one merged range.",
  );

  function updateMeta(workbook: WorkbookLike) {
    const totalSheets = workbook.workSheets.length;
    const totalRows = workbook.workSheets.reduce((sum, sheet) => sum + sheet.data.length, 0);
    const totalMerges = workbook.workSheets.reduce((sum, sheet) => sum + sheet.mergeCells.length, 0);

    meta.innerHTML = `
      <article><strong>Sheets</strong><span>${totalSheets}</span></article>
      <article><strong>Detected rows</strong><span>${totalRows}</span></article>
      <article><strong>Merged ranges</strong><span>${totalMerges}</span></article>
    `;
  }

  function renderSheetTabs(workbook: WorkbookLike) {
    if (args.mode !== "sheet") {
      tabs.innerHTML = "";
      return;
    }

    tabs.innerHTML = "";
    workbook.workSheets.forEach((sheet, index) => {
      if (sheet.state && sheet.state !== 'visible') return;
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = `sb-demo__tab${index === currentSheetIndex ? " is-active" : ""}`;
      tab.textContent = sheet.name || `Sheet ${index + 1}`;
      tab.addEventListener("click", () => {
        currentSheetIndex = index;
        if (!currentWorkbook) {
          return;
        }

        canvas.innerHTML = parser.toHTMLSheet(currentWorkbook as never, currentSheetIndex);
        renderSheetTabs(currentWorkbook);
      });
      tabs.append(tab);
    });
  }

  async function parseWorkbook(buffer: ArrayBuffer, sourceLabel: string) {
    setStatus(status, "Processing", `Reading ${sourceLabel} with the selected parser options.`);

    try {
      currentWorkbook = (await parser.readFile(buffer, options)) as WorkbookLike;
      currentSheetIndex = Math.max(0, currentWorkbook.workSheets.findIndex(sheet => !sheet.state || sheet.state === 'visible'));

      updateMeta(currentWorkbook);
      renderSheetTabs(currentWorkbook);
      canvas.innerHTML =
        args.mode === "sheet"
          ? parser.toHTMLSheet(currentWorkbook as never, currentSheetIndex)
          : parser.toHTML(currentWorkbook as never);

      setStatus(
        status,
        "Workbook loaded",
        `${sourceLabel} was parsed successfully. You can change options and load it again.`,
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      setStatus(status, "Parsing failed", message);
      setError(canvas, message);
      tabs.innerHTML = "";
      meta.innerHTML = "";
    }
  }

  sampleButton.addEventListener("click", () => {
    void parseWorkbook(base64ToArrayBuffer(sampleWorkbookBase64), "the embedded sample workbook");
  });

  uploadInput.addEventListener("change", async () => {
    const file = uploadInput.files?.[0];
    if (!file) {
      return;
    }

    const buffer = await file.arrayBuffer();
    void parseWorkbook(buffer, file.name);
  });

  optionsHost.querySelectorAll<HTMLInputElement>('input[type="checkbox"]').forEach((input) => {
    input.addEventListener("change", () => {
      const optionName = input.dataset.option as keyof ParserOptions;
      options[optionName] = input.checked;
      setStatus(
        status,
        "Options updated",
        "The new settings will be applied the next time you load a workbook.",
      );
    });
  });

  downloadLink.addEventListener("click", () => {
    URL.revokeObjectURL(currentUrl);
    currentUrl = createDownloadUrl();
    downloadLink.href = currentUrl;
  });

  return root;
}

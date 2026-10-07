import { XlsxParser, XlsxWorkerParser, PREVIEW_LIMITS, type Workbook, type XlsxSheetPage } from "../../../../dist/index.js";
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

export type XlsxDemoElement = HTMLElement & { dispose: () => void };

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

export function createXlsxDemo(args: StoryArgs): XlsxDemoElement {
  const root = Object.assign(document.createElement("div"), { dispose: () => {} });
  root.className = "sb-demo";

  const options: ParserOptions = {
    dense: args.dense,
    styles: args.styles,
    drawings: args.drawings,
    skipHiddenRows: args.skipHiddenRows,
  };

  let currentWorkbook: Workbook | null = null;
  const workerParser = new XlsxWorkerParser(() => new Worker(new URL('./parserWorker.ts', import.meta.url), { type: 'module' }));
  let loadGeneration = 0;
  let currentLoad: AbortController | undefined;
  let rowPage = 0, columnPage = 0;
  let currentPage: XlsxSheetPage | undefined;
  let disposed = false;
  let currentSheetIndex = 0;
  const currentUrl = createDownloadUrl();

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
            <button class="sb-demo__button" type="button" data-action="cancel" disabled>Cancel loading</button>
            <p class="sb-demo__small">Up to 10 MiB and 250,000 sheet positions across the workbook. Additional limits apply to XML, images, and merged cells.</p>
            <p class="sb-demo__small">Preview uses saved formula results. Some number formats, drawings, and Excel layout features have limited support.</p>
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
              <h2>${args.mode === "sheet" ? "Workbook preview" : "Workbook preview and export"}</h2>
              <div class="sb-demo__toolbar-copy">
                Browse one sheet at a time using row and column pages.
              </div>
            </div>
            <div class="sb-demo__meta" data-role="meta"></div>
          </div>
          <div class="sb-demo__tabs" data-role="tabs"></div>
          <nav class="sb-demo__pages" aria-label="Spreadsheet pages">
            <button type="button" data-action="row-prev" disabled>Previous rows</button>
            <button type="button" data-action="row-next" disabled>Next rows</button>
            <button type="button" data-action="col-prev" disabled>Previous columns</button>
            <button type="button" data-action="col-next" disabled>Next columns</button>
            <span data-role="page-info" aria-live="polite"></span>
            ${args.mode === 'all' ? '<button type="button" data-action="export" disabled>Download full HTML</button>' : ''}
          </nav>
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
  const cancelButton = root.querySelector('[data-action="cancel"]') as HTMLButtonElement;
  const pageInfo = root.querySelector('[data-role="page-info"]') as HTMLSpanElement;
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');

  [
    createOption("dense", "dense", "Allocates objects for empty cells; uses more memory on sparse sheets.", options),
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

  function updateMeta(workbook: Workbook) {
    const totalSheets = workbook.workSheets.length;
    const totalRows = workbook.workSheets.reduce((sum, sheet) => sum + sheet.data.length, 0);
    const totalMerges = workbook.workSheets.reduce((sum, sheet) => sum + sheet.mergeCells.length, 0);

    meta.innerHTML = `
      <article><strong>Sheets</strong><span>${totalSheets}</span></article>
      <article><strong>Detected rows</strong><span>${totalRows}</span></article>
      <article><strong>Merged ranges</strong><span>${totalMerges}</span></article>
    `;
  }

  function renderSheetTabs(workbook: Workbook) {

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

        rowPage = columnPage = 0;
        renderPage();
        renderSheetTabs(currentWorkbook);
      });
      tabs.append(tab);
    });
  }

  function renderPage() {
    if (!currentWorkbook) return false;
    const start = performance.now();
    try {
      currentPage = parser.toHTMLSheetPage(currentWorkbook, currentSheetIndex, {
        limits: PREVIEW_LIMITS, rowPage, columnPage, pageRows: 100, pageColumns: 50,
      });
      const generated = performance.now();
      canvas.innerHTML = currentPage.html;
      canvas.scrollTop = canvas.scrollLeft = 0;
      const inserted = performance.now();
      root.dataset.renderMetrics = JSON.stringify({ htmlMs: generated - start, domMs: inserted - generated, htmlLength: currentPage.html.length });
      const p = currentPage;
      pageInfo.textContent = p.totalRows && p.totalColumns
        ? `Rows ${p.rowStart}–${p.rowEnd} (page ${p.rowPage + 1}/${p.totalRowPages}); columns ${p.columnStart}–${p.columnEnd} (page ${p.columnPage + 1}/${p.totalColumnPages})`
        : 'No visible cells';
      for (const [action, enabled] of Object.entries({
        'row-prev': p.rowPage > 0, 'row-next': p.rowPage + 1 < p.totalRowPages,
        'col-prev': p.columnPage > 0, 'col-next': p.columnPage + 1 < p.totalColumnPages, export: true,
      })) {
        const button = root.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
        if (button) button.disabled = !enabled;
      }
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setError(canvas, message);
      setStatus(status, 'Preview unavailable', message);
      pageInfo.textContent = 'Preview unavailable';
      root.querySelectorAll<HTMLButtonElement>('.sb-demo__pages button').forEach(button => button.disabled = true);
      return false;
    }
  }

  function cancelLoad() {
    loadGeneration++;
    currentLoad?.abort();
    currentLoad = undefined;
    cancelButton.disabled = true;
    root.removeAttribute('aria-busy');
  }

  async function loadWorkbook(read: () => Promise<ArrayBuffer>, sourceLabel: string) {
    cancelLoad();
    const generation = loadGeneration;
    const controller = currentLoad = new AbortController();
    const snapshot = { ...options, limits: PREVIEW_LIMITS };
    const start = performance.now();
    cancelButton.disabled = false;
    delete root.dataset.firstDisplayMs;
    root.setAttribute('aria-busy', 'true');
    setStatus(status, 'Processing', `Reading ${sourceLabel}.`);
    try {
      const buffer = await read();
      if (generation !== loadGeneration || disposed) return;
      const parseStart = performance.now();
      const workbook = await workerParser.readFile(buffer, {
        ...snapshot, signal: controller.signal,
        onProgress: progress => {
          if (generation !== loadGeneration || disposed) return;
          const detail = progress.phase === 'worksheet'
            ? `Sheet ${progress.completedSheets + 1} of ${progress.totalSheets}: ${progress.sheetName}`
            : progress.phase === 'complete' ? 'Preparing preview' : 'Reading workbook metadata';
          setStatus(status, 'Processing', detail);
        },
      });
      if (generation !== loadGeneration || disposed) return;
      root.dataset.parseMs = String(performance.now() - parseStart);
      currentWorkbook = workbook;
      currentSheetIndex = workbook.workSheets.findIndex(sheet => !sheet.state || sheet.state === 'visible');
      rowPage = columnPage = 0;
      updateMeta(workbook);
      renderSheetTabs(workbook);
      let rendered = true;
      if (currentSheetIndex < 0) {
        canvas.textContent = 'This workbook has no visible sheets.';
        pageInfo.textContent = '';
        root.querySelectorAll<HTMLButtonElement>('.sb-demo__pages button').forEach(button => button.disabled = true);
      } else rendered = renderPage();
      root.dataset.loadMs = String(performance.now() - start);
      requestAnimationFrame(() => requestAnimationFrame(() => {
        if (generation === loadGeneration && !disposed) root.dataset.firstDisplayMs = String(performance.now() - start);
      }));
      if (rendered) setStatus(status, 'Workbook loaded', `${sourceLabel} is ready.`);
    } catch (error) {
      if (generation !== loadGeneration || disposed) return;
      currentWorkbook = null;
      const message = error instanceof Error ? error.message : String(error);
      setStatus(status, 'Loading failed', message);
      setError(canvas, message);
      tabs.replaceChildren();
      meta.replaceChildren();
      pageInfo.textContent = '';
      root.querySelectorAll<HTMLButtonElement>('.sb-demo__pages button').forEach(button => button.disabled = true);
    } finally {
      if (generation === loadGeneration && !disposed) {
        currentLoad = undefined;
        cancelButton.disabled = true;
        root.removeAttribute('aria-busy');
      }
    }
  }

  cancelButton.addEventListener('click', () => {
    cancelLoad();
    setStatus(status, 'Loading cancelled', 'You can load another workbook.');
  });
  sampleButton.addEventListener('click', () => {
    void loadWorkbook(async () => base64ToArrayBuffer(sampleWorkbookBase64), 'the embedded sample workbook');
  });
  uploadInput.addEventListener('change', () => {
    const file = uploadInput.files?.[0];
    if (!file) return;
    if (file.size > PREVIEW_LIMITS.maxFileBytes) {
      cancelLoad();
      setStatus(status, 'File too large', 'Choose an XLSX file no larger than 10 MiB.');
      return;
    }
    void loadWorkbook(() => file.arrayBuffer(), file.name);
    uploadInput.value = '';
  });
  for (const [action, dr, dc] of [['row-prev', -1, 0], ['row-next', 1, 0], ['col-prev', 0, -1], ['col-next', 0, 1]] as const) {
    root.querySelector(`[data-action="${action}"]`)?.addEventListener('click', () => {
      rowPage += dr;
      columnPage += dc;
      renderPage();
    });
  }
  root.querySelector('[data-action="export"]')?.addEventListener('click', () => {
    if (!currentWorkbook) return;
    try {
      const html = parser.toHTML(currentWorkbook, { limits: PREVIEW_LIMITS });
      const url = URL.createObjectURL(new Blob([html], { type: 'text/html' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = 'workbook.html';
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) { setStatus(status, 'Export unavailable', error instanceof Error ? error.message : String(error)); }
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

  root.dispose = () => {
    disposed = true;
    cancelLoad();
    URL.revokeObjectURL(currentUrl);
    currentWorkbook = null;
  };

  return root;
}

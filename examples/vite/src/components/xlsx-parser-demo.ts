import { LitElement, html } from "lit";
import { createXlsxDemo, type XlsxDemoElement } from "../storybook/demoApp";

type DemoMode = "sheet" | "all";

export class XlsxParserDemoElement extends LitElement {
  static properties = {
    dense: { type: Boolean },
    styles: { type: Boolean },
    drawings: { type: Boolean },
    skipHiddenRows: { type: Boolean, attribute: "skip-hidden-rows" },
    mode: { type: String },
  };

  declare dense: boolean;
  declare styles: boolean;
  declare drawings: boolean;
  declare skipHiddenRows: boolean;
  declare mode: DemoMode;

  constructor() {
    super();
    this.dense = false;
    this.styles = false;
    this.drawings = false;
    this.skipHiddenRows = true;
    this.mode = "sheet";
  }

  createRenderRoot() {
    return this;
  }

  connectedCallback() {
    super.connectedCallback();
    if (this.hasUpdated) this.mountDemo();
  }

  firstUpdated() {
    this.mountDemo();
  }

  updated(changed: Map<string, unknown>) {
    if (
      changed.has("dense") ||
      changed.has("styles") ||
      changed.has("drawings") ||
      changed.has("skipHiddenRows") ||
      changed.has("mode")
    ) {
      this.mountDemo();
    }
  }

  disconnectedCallback() {
    this.querySelector<XlsxDemoElement>('.sb-demo')?.dispose();
    super.disconnectedCallback();
  }

  private mountDemo() {
    const mount = this.querySelector("[data-demo-root]");
    if (!mount) {
      return;
    }

    mount.querySelector<XlsxDemoElement>('.sb-demo')?.dispose();
    mount.replaceChildren(
      createXlsxDemo({
        dense: this.dense,
        styles: this.styles,
        drawings: this.drawings,
        skipHiddenRows: this.skipHiddenRows,
        mode: this.mode,
      }),
    );
  }

  render() {
    return html`<div data-demo-root></div>`;
  }
}

customElements.define("xlsx-parser-demo", XlsxParserDemoElement);

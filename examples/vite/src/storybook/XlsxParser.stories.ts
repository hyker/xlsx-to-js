import type { Meta, StoryObj } from "@storybook/web-components-vite";
import { html } from "lit";
import "../components/xlsx-parser-demo";

type DemoArgs = {
  dense: boolean;
  styles: boolean;
  drawings: boolean;
  skipHiddenRows: boolean;
};

const meta = {
  title: "XlsxParser",
  component: "xlsx-parser-demo",
  argTypes: {
    dense: {
      control: "boolean",
      description: "Keeps sparse workbooks compact by skipping empty cells in the output matrix.",
      table: {
        category: "Parser options",
        defaultValue: { summary: "true" },
      },
    },
    styles: {
      control: "boolean",
      description: "Parses cell styles, fills, borders, and typography metadata when available.",
      table: {
        category: "Parser options",
        defaultValue: { summary: "false" },
      },
    },
    drawings: {
      control: "boolean",
      description: "Parses images and drawing objects, including shapes and text boxes.",
      table: {
        category: "Parser options",
        defaultValue: { summary: "false" },
      },
    },
    skipHiddenRows: {
      control: "boolean",
      description: "Ignores rows marked as hidden in the workbook metadata.",
      table: {
        category: "Parser options",
        defaultValue: { summary: "true" },
      },
    },
  },
} satisfies Meta<DemoArgs>;

export default meta;

type Story = StoryObj<DemoArgs>;

export const InteractiveSheetView: Story = {
  name: "Single-sheet rendering",
  args: {
    dense: true,
    styles: false,
    drawings: false,
    skipHiddenRows: true,
  },
  render: (args) => html`<xlsx-parser-demo
    ?dense=${args.dense}
    ?styles=${args.styles}
    ?drawings=${args.drawings}
    ?skip-hidden-rows=${args.skipHiddenRows}
    mode="sheet"
  ></xlsx-parser-demo>`,
};

export const FullWorkbookHtml: Story = {
  name: "Full workbook HTML",
  args: {
    dense: true,
    styles: false,
    drawings: false,
    skipHiddenRows: true,
  },
  render: (args) => html`<xlsx-parser-demo
    ?dense=${args.dense}
    ?styles=${args.styles}
    ?drawings=${args.drawings}
    ?skip-hidden-rows=${args.skipHiddenRows}
    mode="all"
  ></xlsx-parser-demo>`,
};

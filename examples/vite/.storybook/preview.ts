import type { Preview } from "@storybook/web-components-vite";

const preview: Preview = {
  parameters: {
    layout: "fullscreen",

    options: {
      storySort: {
        order: ["XlsxParser"],
      },
    },
  },
};

export default preview;

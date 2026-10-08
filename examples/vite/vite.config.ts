import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

export default defineConfig({
  server: {
    open: true,
  },
  // Pre-bundle so first use does not trigger a dependency reload mid-run,
  // which the memory check would record as a crash.
  optimizeDeps: {
    include: ["jszip", "@xmldom/xmldom"],
  },
  build: {
    rollupOptions: {
      // memory.html is the on-device memory check; see the README.
      input: {
        main: fileURLToPath(new URL("index.html", import.meta.url)),
        memory: fileURLToPath(new URL("memory.html", import.meta.url)),
      },
    },
  },
});

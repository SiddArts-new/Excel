import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";
import { resolve } from "path";

export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    port: 3000,
    strictPort: true
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html")
      }
    }
  }
});

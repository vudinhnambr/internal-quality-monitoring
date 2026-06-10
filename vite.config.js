import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/gdrive": {
        target: "https://docs.google.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/gdrive/, ""),
      },
    },
  },
});

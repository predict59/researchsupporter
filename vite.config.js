import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  base: "/researchsupporter/",
  plugins: [react()],
  resolve: {
    dedupe: ["react", "react-dom"],
    preserveSymlinks: true,
  },
  optimizeDeps: {
    esbuildOptions: {
      preserveSymlinks: true,
    },
  },
  server: {
    fs: {
      strict: false,
    },
  },
});

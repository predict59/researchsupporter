import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  // Predict GitHub Pages 경로이므로 KPRC 코드 동기화 시에도 바꾸지 않는다.
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

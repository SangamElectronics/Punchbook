import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Vercel builds and serves from the repo root, so no special `base` path
// is needed (unlike GitHub Pages, which needs base: "/repo-name/").
export default defineConfig({
  plugins: [react()],
});

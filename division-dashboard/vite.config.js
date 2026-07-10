import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base so the built site works when served from a subpath
// (e.g. GitHub Pages project sites) as well as from a domain root.
export default defineConfig({
  base: "./",
  plugins: [react()],
});

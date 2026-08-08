import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base so the built site works when served from a subpath
// (e.g. GitHub Pages project sites) as well as from a domain root.
export default defineConfig({
  base: "./",
  plugins: [react()],
  // Emit ASCII-only output (non-ASCII characters become \uXXXX escapes).
  // Renders identically at runtime and keeps the bundle free of raw
  // non-ASCII bytes.
  esbuild: { charset: "ascii" },
});

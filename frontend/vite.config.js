import { defineConfig } from "vite";

// Builds the SuperDoc editor page (HTML + bundled JS/CSS) into app/static/dist
// so it works fully offline (no CDN). FastAPI serves dist/editor.html at /editor
// and mounts the hashed assets at /assets.
export default defineConfig({
  root: __dirname,
  base: "/assets/",
  build: {
    outDir: "../app/static/dist",
    assetsDir: "",
    emptyOutDir: true,
    rollupOptions: {
      input: "editor.html",
    },
  },
});

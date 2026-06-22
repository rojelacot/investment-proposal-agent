import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Proxies /api/* to the Express backend (server.js, port 5174). Vite's `server`
// and `preview` blocks each need their OWN proxy config — `vite preview` (and
// just opening the built dist/ output) does NOT inherit `server.proxy`, so
// without this, every /api call silently 404s when the app is run via
// `npm run preview` instead of `npm run dev`.
const apiProxy = {
  "/api": {
    target: "http://localhost:5174",
    changeOrigin: true,
  },
};

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: apiProxy,
  },
  preview: {
    proxy: apiProxy,
  },
});

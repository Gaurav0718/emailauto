import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// For local dev with the Functions API, run `npm run dev:full` instead of
// `npm run dev` — it builds once then serves site + /api/* together via
// `wrangler pages dev`, matching production exactly.
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
  },
});

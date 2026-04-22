import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    allowedHosts: ["app.reshamsutra.com"]
  },
  preview: {
    allowedHosts: ["app.reshamsutra.com"]
  }
});

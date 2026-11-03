// Static site build:  npm run web:dev  (local)  ·  npm run web:build  -> dist/ (deployable as-is)
import { defineConfig, type Plugin } from "vite";

/**
 * The only hosts the page may talk to: itself and HUJI Moodle. Added at build time only, because the
 * dev server injects inline scripts and a websocket that a strict policy would block.
 */
export const CSP = [
  "default-src 'self'",
  "connect-src 'self' https://moodle.huji.ac.il",
  "img-src 'self' data:",
  "style-src 'self'",
  "script-src 'self'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

const csp = (): Plugin => ({
  name: "csp",
  apply: "build",
  transformIndexHtml: (html) =>
    html.replace("<meta charset=\"UTF-8\" />", `<meta charset="UTF-8" />\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`),
});

export default defineConfig({
  root: "web",
  base: "./", // works at any path, e.g. https://<user>.github.io/<repo>/
  plugins: [csp()],
  // assetsInlineLimit 0: small font files would otherwise be inlined as data: URIs, which the CSP blocks.
  build: { outDir: "../dist", emptyOutDir: true, target: "es2022", assetsInlineLimit: 0 },
  server: { port: 5173, strictPort: true },
  preview: { port: 4173, strictPort: true },
});

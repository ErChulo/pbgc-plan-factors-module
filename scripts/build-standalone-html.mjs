import { readFile, writeFile, mkdir } from "node:fs/promises";
import { resolve } from "node:path";

const PF_VERSION = "0.7.13";
const root = resolve(process.cwd());
const templatePath = resolve(root, "templates", `PF_template.v${PF_VERSION}.xlsx`);
const jszipPath = resolve(root, "node_modules", "jszip", "dist", "jszip.min.js");
const appTemplatePath = resolve(root, "standalone", "app.template.js");
const outPath = resolve(root, "standalone", "pf-builder-standalone.html");

const [templateBuffer, jszipRaw, appTemplateRaw] = await Promise.all([
  readFile(templatePath),
  readFile(jszipPath, "utf8"),
  readFile(appTemplatePath, "utf8")
]);

const templateBase64 = templateBuffer.toString("base64");
const jszipCode = jszipRaw.replace(/<\/script>/gi, "<\\/script>");
const appCode = appTemplateRaw
  .replaceAll("__PF_VERSION__", PF_VERSION)
  .replaceAll("__TEMPLATE_BASE64__", templateBase64)
  .replace(/<\/script>/gi, "<\\/script>");

const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>PF Builder Standalone v${PF_VERSION}</title>
  <style>
    :root { --bg:#f4f1e6; --panel:#fffef8; --ink:#1f2a2a; --muted:#586865; --accent:#0d5d44; --border:#d7d0c1; --danger:#9f1a1a; }
    * { box-sizing:border-box; }
    body { margin:0; font-family:"Segoe UI","Aptos",sans-serif; background:radial-gradient(circle at top left,#fff8ea,var(--bg)); color:var(--ink); }
    main { max-width:920px; margin:28px auto; border:1px solid var(--border); background:var(--panel); border-radius:12px; padding:20px; }
    h1 { margin:0 0 10px; font-size:1.6rem; }
    p { margin:0 0 14px; color:var(--muted); }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
    label { display:block; font-weight:700; }
    input[type=file] { width:100%; margin-top:6px; }
    button { margin-top:14px; border:none; border-radius:8px; background:var(--accent); color:white; padding:10px 14px; font-weight:700; cursor:pointer; }
    pre { white-space:pre-wrap; color:var(--danger); margin-top:12px; }
    section { margin-top:12px; border:1px solid var(--border); border-radius:8px; background:white; padding:10px; }
    @media (max-width:760px) { .grid { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <div id="app"></div>
  <script>${jszipCode}</script>
  <script>${appCode}</script>
</body>
</html>`;

await mkdir(resolve(root, "standalone"), { recursive: true });
await writeFile(outPath, html, "utf8");
console.log(`Generated ${outPath}`);

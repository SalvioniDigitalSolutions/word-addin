#!/usr/bin/env node
/**
 * Writes manifest.xml from manifest.template.xml using ADDIN_PUBLIC_URL.
 * Default https://localhost:3000 — use only with npm run dev.
 * For no local server: set ADDIN_PUBLIC_URL to your static site origin, npm run build, deploy dist/.
 */
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const templatePath = path.join(root, "manifest.template.xml");
const raw = (process.env.ADDIN_PUBLIC_URL || "https://localhost:3000").trim();
const base = raw.replace(/\/$/, "");

let origin;
try {
  origin = new URL(base).origin;
} catch {
  console.error("Invalid ADDIN_PUBLIC_URL (must be https URL, no path quirks):", raw);
  process.exit(1);
}

if (!/^https:/i.test(origin)) {
  console.error("ADDIN_PUBLIC_URL must use https (required for Word add-ins).");
  process.exit(1);
}

const tpl = fs.readFileSync(templatePath, "utf8");
const out = tpl.replaceAll("{{ORIGIN}}", origin).replaceAll("{{BASE_URL}}", base);

fs.writeFileSync(path.join(root, "manifest.xml"), out);

const distDir = path.join(root, "dist");
if (fs.existsSync(distDir)) {
  fs.writeFileSync(path.join(distDir, "manifest.xml"), out);
}

console.log(`manifest.xml → BASE_URL=${base} ORIGIN=${origin}`);

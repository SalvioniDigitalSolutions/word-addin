# Legal AI Assistant — Word add-in

Task pane add-in for Microsoft Word: **Tools** (anonymize, translate), **Chat** (natural-language requests with tracked changes), and **Settings** for **OpenAI**, **Anthropic**, **Azure OpenAI**, or a **custom OpenAI-compatible** endpoint.

API keys are stored in **browser storage** for the add-in’s origin.

**Request mode (Settings):**

- **Proxy** — The task pane `POST`s to your backend (`/api/legal-ai/complete`). For local development, **`npm run dev:all`** runs webpack (HTTPS, port 3000) and a small **Express proxy** on **127.0.0.1:3548**; webpack forwards `/api` to that server so provider APIs are called from Node (no browser CORS).
- **Direct** — The task pane calls the LLM provider with `fetch` from Word’s webview. There is no app server; some providers block this with CORS.

Word still loads the add-in UI from **HTTPS URLs** in the manifest. To use the add-in without keeping a dev server running on your machine, host the built `dist/` on static hosting and set **`ADDIN_PUBLIC_URL`** when you build (see below).

## Local development (proxy, recommended)

1. `npm install`
2. Trust dev certificates (once): `npx office-addin-dev-certs install`
3. **`npm run dev:all`** — starts the proxy on port **3548** and webpack dev server on **https://localhost:3000**.
4. Keep **Settings → Request mode** on **Proxy** (default); leave **Backend base URL** empty so calls go to same-origin `/api` (proxied to 3548).
5. **Sideload the manifest** (pick your OS — manifest points at `https://localhost:3000`):
   - **macOS:** `npm run sideload:mac` (if the ribbon button never appears, try `npm run sideload:mac:reset`, then quit Word fully and reopen).
   - **Windows:** `npm run sideload:windows`, then **fully close Word** and reopen; or skip the script and use **Insert → Add-ins → My Add-ins → Upload My Add-in** and choose the repo’s **`manifest.xml`**.
6. Open a **real document** (not only Word’s start screen). On **Insert** or **Home**, look for the **Legal AI** group.

`npm run dev` alone only starts webpack; **Proxy** mode will fail until something serves `POST /api/legal-ai/complete` (use `dev:all` or `npm run server` in another terminal).

### If you “can’t run it locally”

- **You need desktop Word** (or Word inside a VM that can reach your PC’s `localhost`). **Word in the browser** cannot load `https://localhost:3000` from your machine unless you use a tunnel (e.g. ngrok) and rebuild the manifest with that HTTPS URL (`ADDIN_PUBLIC_URL=… npm run build`).
- **Trust dev HTTPS (required):** `npx office-addin-dev-certs install` — on Windows you may need an **elevated** PowerShell once so the cert is trusted for Office’s webview.
- **Corporate / managed Word** sometimes **blocks sideloading** or self-signed certs — then use **host `dist/` on real HTTPS** (see below) and sideload that `manifest.xml`, or ask IT for an exception.
- **Still stuck:** use **Upload My Add-in** with `manifest.xml` after `dev:all` is running; confirm **https://localhost:3000/taskpane.html** opens in Edge/Chrome without certificate errors.

## Use without a local dev server

1. Set a public HTTPS base URL (no trailing slash), e.g. `https://you.github.io/legal-word-addin`.
2. `ADDIN_PUBLIC_URL='https://…' npm run build`
3. Upload everything inside **`dist/`** to that URL.
4. In the task pane **Settings**, choose **Direct** (or host a compatible **`POST /api/legal-ai/complete`** API and use **Proxy** with **Backend base URL**).
5. Sideload **`dist/manifest.xml`**.

Regenerate **`manifest.xml`** with `npm run manifest` (same `ADDIN_PUBLIC_URL` as for build). Source template: **`manifest.template.xml`**.

## CORS

In **Direct** mode, some providers block browser `fetch` from arbitrary origins. **Proxy** mode (local or your own HTTPS API) avoids that.

## Production build

- `npm run build` — webpack production bundle into **`dist/`**, then writes **`manifest.xml`** (root + `dist/`) from the template.

## Security and compliance

- Keys in Settings live in **browser storage** for that origin. The dev proxy receives keys in the JSON body over localhost only; do not expose that endpoint to the internet without authentication.
- This is a **tooling scaffold**, not legal advice.

## Scripts

| Script | Purpose |
|--------|---------|
| `npm run dev:all` | Local proxy (3548) + webpack HTTPS (3000) |
| `npm run dev` | Webpack dev server only |
| `npm run server` | Local proxy only (`server/proxy.mjs`) |
| `npm run build` | Production `dist/` + manifest |
| `npm run manifest` | Regenerate `manifest.xml` from template |
| `npm run sideload:mac` | Copy manifest into Word (macOS) |
| `npm run sideload:mac:reset` | Same + clear Wef cache |
| `npm run sideload:windows` | Copy manifest into Word Wef folder (Windows) |

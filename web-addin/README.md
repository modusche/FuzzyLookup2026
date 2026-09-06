# Fuzzy Lookup 2026 — Office Web Add-in

JavaScript/HTML version of the Fuzzy Lookup add-in. Works on Windows, Mac, and Excel Online.

## Quick Start (Sideload for testing)

### 1. Install dependencies

```bash
cd web-addin
npm install
```

### 2. Generate SSL certificate (Office Add-ins require HTTPS)

```bash
npx office-addin-dev-certs install
```

### 3. Start the dev server

```bash
npx office-addin-dev-certs verify && npx serve src --listen 3000 --ssl
```

Or use the Office Add-in CLI:

```bash
npx office-addin-debugging start manifest.xml
```

### 4. Sideload in Excel

**Desktop Excel:**
1. Open Excel
2. Go to **Insert → Add-ins → My Add-ins → Upload My Add-in**
3. Browse to `web-addin/manifest.xml`
4. Click **Upload**

**Excel Online:**
1. Open a workbook in Excel Online
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Browse to `manifest.xml`

## Publishing to AppSource

1. Register at [Microsoft Partner Center](https://partner.microsoft.com/)
2. Host the web app on a public HTTPS URL (Azure, Vercel, etc.)
3. Update all `localhost:3000` URLs in `manifest.xml` to your public URL
4. Submit for review

## Architecture

```
web-addin/
├── manifest.xml          # Office Add-in manifest
├── package.json
├── src/
│   ├── taskpane.html     # Task pane UI
│   ├── taskpane.js       # Excel interaction + matching engine
│   └── algorithms.js     # Levenshtein, Jaro-Winkler, Jaccard
```

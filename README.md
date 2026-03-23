# Word Add-in — Base Template

A reusable Word add-in scaffold with 3 built-in features. Clone & customise per client.

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Defines the add-in, ribbon buttons, and URLs |
| `commands.html` | Invisible page that loads ribbon button functions |
| `commands.js` | Functions for ribbon buttons (landscape page, insert table) |
| `taskpane.html` | Style Pane UI shown in the task pane |
| `taskpane.js` | Style Pane logic |
| `assets/` | Icon images (icon-16.png, icon-32.png, icon-64.png, icon-80.png) |

---

## Features

### 1. Style Pane (ribbon button → task pane)
Opens a side panel with one-click style buttons (Heading 1/2/3, Normal, Quote, etc).
To add more styles: edit `taskpane.html` and add a new `<button>` with `onclick="applyStyle('Style Name')"`.

### 2. Insert Landscape Page (ribbon button → function)
Inserts a landscape-oriented page at the cursor using a section break + OOXML.

### 3. Insert Table (ribbon button → function)
Inserts a formatted table with a styled header row.
To customise: edit `commands.js` and change `ROWS`, `COLS`, `HEADERS`.

---

## How to customise per client

1. **Copy** this folder and rename it for the client (e.g. `word-addin-clientA`)
2. In `manifest.xml`:
   - Change `<Id>` to a new unique GUID
   - Change `<DisplayName>` to the client's name
   - Replace all `YOUR_HOSTED_URL` with the actual hosted URL
3. Adjust buttons, styles, table defaults as needed
4. Deploy (see below)

---

## Hosting

These are static files — host them anywhere:
- **GitHub Pages** (free, easy)
- **Netlify / Vercel** (free tier, drag & drop deploy)
- **Azure Static Web Apps** (Microsoft-native, good choice)
- Your own web server

⚠️ Must be served over **HTTPS**.

---

## Installing for a client (sideloading)

### Word on Windows
1. Copy `manifest.xml` to a shared network folder
2. In Word: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → add the folder path
3. Restart Word → Insert → My Add-ins → Shared Folder → select the add-in

### Word on Mac
1. Copy `manifest.xml` to: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
2. Restart Word → Insert → My Add-ins → select it

### Word Online
1. Insert → Add-ins → Upload My Add-in → upload `manifest.xml`

---

## Icons needed (assets/ folder)
Add PNG icons at these sizes:
- `icon-16.png` — 16×16px
- `icon-32.png` — 32×32px
- `icon-64.png` — 64×64px
- `icon-80.png` — 80×80px

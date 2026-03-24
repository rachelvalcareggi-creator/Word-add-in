# AGENTS.md — Word Add-in Project

## Project Overview

This is a **Microsoft Word Add-in** project — a static HTML/JS application that extends Word with custom functionality. No build system; files are deployed directly to a web server.

## Project Structure

```
├── manifest.xml          # Add-in definition, ribbon buttons, URLs
├── commands.html/js      # Invisible page for ribbon button functions
├── taskpane.html/js      # Style Pane UI and logic
├── assets/               # Icon PNGs (16, 32, 64, 80px)
└── README.md             # Deployment and customization guide
```

## Build/Lint/Test Commands

### No Build System
This project uses **pure static files**. No npm, bundler, or transpilation.

### Testing
- **Manual testing only**: Load the add-in via Word's sideloading
- No automated tests exist

### Validation
- XML must be valid (use VS Code XML extension or `xmllint`)
- JS syntax errors can be checked with browser dev tools
- Test in Word desktop (Windows/Mac) and Word Online

## Code Style Guidelines

### General Conventions

- **Vanilla JavaScript** — no frameworks or transpilers
- **Office.js API** — all Word operations via `Word.run(context => {...})` pattern
- **Always call `context.sync()`** after loading properties or making changes
- **Always call `event.completed()`** in ribbon button handlers

### JavaScript Style

```javascript
// ✓ CORRECT: async/await with Word.run
async function insertLandscapePage(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      // ... operations ...
      await context.sync();
    });
  } catch (error) {
    console.error("insertLandscapePage error:", error);
  } finally {
    event.completed();
  }
}

// ✓ CORRECT: Load properties before accessing
paragraphs.load("items");
await context.sync();
paragraphs.items.forEach((para) => { /* ... */ });
await context.sync();
```

### HTML Conventions

- Use semantic HTML5 (`<!DOCTYPE html>`, `<html lang="en">`)
- Set `charset="UTF-8"` in meta
- Load Office.js from CDN: `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`
- Inline styles allowed for simple pages; extract to `<style>` block if complex

### XML Conventions (manifest.xml)

- Use consistent indentation (2 spaces)
- Resource IDs must match between `<Resources>` and usage
- Icon URLs must be HTTPS
- `xsi:type` attributes required on VersionOverrides elements

### Naming Conventions

| Element | Convention | Example |
|---------|------------|---------|
| Functions | camelCase | `insertLandscapePage` |
| Constants | UPPER_SNAKE | `ROWS`, `COLS` |
| HTML IDs | kebab-case | `style-pane-container` |
| XML IDs | PascalCase | `HomeOpenStylePane` |
| XML Resid | PascalCase | `OpenStylePane.Label` |

### Error Handling

```javascript
// Always wrap Word operations in try/catch
try {
  await Word.run(async (context) => {
    // ...
  });
} catch (error) {
  console.error("FunctionName error:", error);
  // Update UI to inform user
}
```

### Ribbon Button Functions

Ribbon button handlers MUST:
1. Be `async` functions
2. Accept `event` parameter
3. Call `event.completed()` in `finally` block
4. Register with `Office.actions.associate()` in `Office.onReady()`

```javascript
Office.onReady(() => {
  Office.actions.associate("insertTable", insertTable);
});
```

## Deployment Notes

- Files must be served over **HTTPS**
- Update `manifest.xml` URLs to match your hosted location
- Generate new GUID for `<Id>` when cloning for new client
- Test icons at all required sizes (16, 32, 64, 80px)

## Customization Guide

### Adding Styles to Style Pane
Edit `taskpane.html` and add:
```html
<button class="style-btn" onclick="applyStyle('Style Name')">
  Style Name <span class="preview">Description</span>
</button>
```

### Customizing Table
Edit `commands.js` constants:
```javascript
const ROWS = 4;
const COLS = 3;
const HEADERS = ["Column 1", "Column 2", "Column 3"];
```

### Adding New Ribbon Button
1. Add `<Control>` in `manifest.xml` with `ExecuteFunction` action
2. Create function in `commands.js` matching `<FunctionName>`
3. Register in `Office.actions.associate()`

## Git Workflow

- **Mai commit/push automatici**: non eseguire mai commit o push senza esplicita richiesta
- **Esecuzione esplicita**: quando l'utente richiede esplicitamente un commit/push, eseguirlo solo per quella volta specifica; l'autorizzazione non si applica automaticamente alle operazioni successive

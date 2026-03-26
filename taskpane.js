/* taskpane.js — Rachele Tools logic */

const STORAGE_KEY = "racheleToolsSetup";

let selectedCover = null;

Office.onReady(() => {
  initTabs();
  initCoverTab();
  initTableGrid();
});

/* ── TABS ── */
function initTabs() {
  const tabBtns = document.querySelectorAll(".tab-btn");
  tabBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const tabId = btn.getAttribute("data-tab");
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + tabId).classList.add("active");

      if (tabId === "cover") {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (!saved) {
          showSetupDialog();
        }
      }
    });
  });
}

function initCoverTab() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    document.getElementById("setupOverlay").style.display = "none";
    document.getElementById("tab-cover").innerHTML = `
      <h2>Cover Page Setup</h2>
      <button class="menu-btn" onclick="resetSetup()" style="margin: 16px 0; background: #fff3cd; border-color: #ffc107;">
        Reset Cover Setup
      </button>
      <p style="color: #6c757d; font-size: 12px;">Click "Reset Cover Setup" to create a new cover page.</p>
    `;
  }
}

function showSetupDialog() {
  document.getElementById("setupOverlay").classList.add("open");
  document.getElementById("inputDate").value = new Date().toLocaleDateString();
  selectedCover = 0;
  document.querySelectorAll(".cover-thumb").forEach((el) => el.classList.remove("selected"));
  document.querySelector(`[data-cover="0"]`).classList.add("selected");
}

function showMainContent() {
  document.getElementById("setupOverlay").classList.remove("open");
  document.getElementById("setupOverlay").style.display = "none";
  document.getElementById("tab-cover").innerHTML = `
    <h2>Cover Page Setup</h2>
    <button class="menu-btn" onclick="resetSetup()" style="margin: 16px 0; background: #fff3cd; border-color: #ffc107;">
      Reset Cover Setup
    </button>
    <p style="color: #6c757d; font-size: 12px;">Click "Reset Cover Setup" to create a new cover page.</p>
  `;
}

function selectCover(num) {
  selectedCover = num;
  document.querySelectorAll(".cover-thumb").forEach((el) => el.classList.remove("selected"));
  document.querySelector(`[data-cover="${num}"]`).classList.add("selected");
}

function handleOk() {
  const title = document.getElementById("inputTitle").value.trim();
  const subtitle = document.getElementById("inputSubtitle").value.trim();
  const date = document.getElementById("inputDate").value.trim();

  if (!title || !subtitle || !date) {
    setStatus("Please fill all fields");
    return;
  }

  if (selectedCover === 1) {
    createCoverWithImage();
  } else {
    createDocumentWithoutCover(title, subtitle, date);
  }
}

function createDocumentWithoutCover(title, subtitle, date) {
  setStatus("Creating document...");

  Word.run(async (context) => {
    const body = context.document.body;
    const firstPara = body.paragraphs.getFirst();
    const insertPoint = firstPara.getRange("start");

    const titleControl = insertPoint.insertContentControl();
    titleControl.type = "richText";
    titleControl.title = "Title";
    titleControl.tag = "title";
    titleControl.style = "Title";
    titleControl.insertText(title, "end");
    const titleRange = titleControl.getRange();
    titleRange.font.size = 40;
    titleRange.font.bold = true;

    const subtitleControl = titleRange.insertContentControl();
    subtitleControl.type = "richText";
    subtitleControl.title = "Subtitle";
    subtitleControl.tag = "subtitle";
    subtitleControl.insertText(subtitle, "end");
    const subtitleRange = subtitleControl.getRange();
    subtitleRange.font.size = 24;

    const dateControl = subtitleRange.insertContentControl();
    dateControl.type = "richText";
    dateControl.title = "Date";
    dateControl.tag = "date";
    dateControl.insertText(date, "end");
    const dateRange = dateControl.getRange();
    dateRange.font.size = 16;
    dateRange.font.color = "#6c757d";

    await context.sync();
  }).then(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ cover: "none" }));
    showMainContent();
    setStatus("Document created!");
  }).catch((error) => {
    console.error("createDocumentWithoutCover error:", error);
    setStatus("Error creating document");
  });
}

function createCoverWithImage() {
  const title = document.getElementById("inputTitle").value.trim();
  const subtitle = document.getElementById("inputSubtitle").value.trim();
  const date = document.getElementById("inputDate").value.trim();

  if (!title || !subtitle || !date) {
    setStatus("Please fill all fields");
    return;
  }

  setStatus("Creating cover...");

  const img = new Image();
  img.crossOrigin = "Anonymous";
  img.onload = function () {
    const canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    const imgBase64 = canvas.toDataURL("image/png").split(",")[1];

    Word.run(async (context) => {
      const body = context.document.body;
      const section = context.document.sections.getFirst();
      section.load("pageWidth, pageHeight");
      await context.sync();

      const pageWidth = section.pageWidth;
      const pageHeight = section.pageHeight;

      const firstPara = body.paragraphs.getFirst();
      const insertPoint = firstPara.getRange("start");

      const coverImg = insertPoint.insertInlinePictureFromBase64(imgBase64, "after");
      coverImg.width = pageWidth;
      coverImg.height = pageHeight;

      const imgEnd = coverImg.getRange("end");

      const placeholderPara = imgEnd.insertParagraph("", "after");
      placeholderPara.font.size = 48;
      placeholderPara.insertText("Click here to add your image", "end");

      const textStart = placeholderPara.getRange("end");

      const titleControl = textStart.insertContentControl();
      titleControl.type = "richText";
      titleControl.title = "Title";
      titleControl.tag = "title";
      titleControl.style = "Title";
      titleControl.insertText(title, "end");
      const titleRange = titleControl.getRange();
      titleRange.font.size = 40;
      titleRange.font.bold = true;

      const subtitleControl = titleRange.insertContentControl();
      subtitleControl.type = "richText";
      subtitleControl.title = "Subtitle";
      subtitleControl.tag = "subtitle";
      subtitleControl.insertText(subtitle, "end");
      const subtitleRange = subtitleControl.getRange();
      subtitleRange.font.size = 24;

      const dateControl = subtitleRange.insertContentControl();
      dateControl.type = "richText";
      dateControl.title = "Date";
      dateControl.tag = "date";
      dateControl.insertText(date, "end");
      const dateRange = dateControl.getRange();
      dateRange.font.size = 16;
      dateRange.font.color = "#6c757d";

      await context.sync();
    }).then(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ cover: "1" }));
      showMainContent();
      setStatus("Cover created!");
    }).catch((error) => {
      console.error("createCoverWithImage error:", error);
      setStatus("Error creating cover");
    });
  };
  img.onerror = function () {
    setStatus("Error loading cover image");
  };
  img.src = "assets/cover1.png";
}

function resetSetup() {
  localStorage.removeItem(STORAGE_KEY);
  selectedCover = null;
  showSetupDialog();
}

/* ── STYLES ── */
async function applyStyle(styleName) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      paragraphs.items.forEach((para) => {
        para.style = styleName;
      });

      await context.sync();
      setStatus(`Applied "${styleName}"`);
    });
  } catch (error) {
    logDebug(`applyStyle("${styleName}") failed`, error);
    setStatus(`Could not apply "${styleName}"`);
  }
}

/* ── TABLES ── */
async function applyTableStyle(style) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const tables = selection.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("No table selected");
        return;
      }

      for (const table of tables.items) {
        table.load("rows");
        await context.sync();

        if (style === "heading") {
          table.rows.getItem(0).cells.format.fill = "#4F46E5";
          table.rows.getItem(0).cells.items.forEach((cell) => {
            cell.body.paragraphs.getItem(0).font.color = "white";
            cell.body.paragraphs.getItem(0).font.bold = true;
          });
        } else if (style === "text") {
          table.rows.getItem(0).cells.items.forEach((cell) => {
            cell.body.paragraphs.getItem(0).font.bold = true;
            cell.body.paragraphs.getItem(0).font.color = "black";
          });
        } else if (style === "bullets") {
          const rowCount = table.rows.count;
          for (let i = 0; i < rowCount; i++) {
            table.getCell(i, 0).body.paragraphs.getItem(0).listFormat.applyDisc();
          }
          table.rows.getItem(0).cells.items.forEach((cell) => {
            cell.body.paragraphs.getItem(0).font.bold = true;
          });
        }
        await context.sync();
      }
      setStatus(`Applied "${style}"`);
    });
  } catch (error) {
    logDebug(`applyTableStyle("${style}") failed`, error);
    setStatus("Could not apply table style");
  }
}

/* ── TABLE GRID PICKER ── */
let currentTableSize = { rows: 0, cols: 0 };

function initTableGrid() {
  const grid = document.getElementById("tableGrid");
  grid.innerHTML = "";
  
  for (let row = 0; row < 10; row++) {
    for (let col = 0; col < 10; col++) {
      const cell = document.createElement("div");
      cell.className = "grid-cell";
      cell.dataset.row = row + 1;
      cell.dataset.col = col + 1;
      cell.addEventListener("mouseenter", function() {
        updateTableGridPreview(row + 1, col + 1);
      });
      cell.addEventListener("click", function(e) {
        insertTableFromGrid(row + 1, col + 1);
      });
      grid.appendChild(cell);
    }
  }
}

function showTableGrid() {
  initTableGrid();
  document.getElementById("tableGridOverlay").classList.add("open");
  document.getElementById("tableGridCounter").textContent = "";
}

function hideTableGrid(event) {
  if (event && event.target.id !== "tableGridOverlay") return;
  document.getElementById("tableGridOverlay").classList.remove("open");
}

function updateTableGridPreview(rows, cols) {
  currentTableSize = { rows, cols };
  document.getElementById("tableGridCounter").textContent = `${rows} × ${cols}`;
  
  document.querySelectorAll(".grid-cell").forEach(cell => {
    const cellRow = parseInt(cell.dataset.row);
    const cellCol = parseInt(cell.dataset.col);
    
    if (cellRow <= rows && cellCol <= cols) {
      cell.classList.add("selected");
    } else {
      cell.classList.remove("selected");
    }
  });
}

async function insertTableFromGrid(rows, cols) {
  hideTableGrid();
  
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      
      const tableData = [];
      for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
          row.push("");
        }
        tableData.push(row);
      }
      
      selection.insertTable(rows, cols, "after", tableData);
      await context.sync();
      setStatus("Table inserted!");
    });
  } catch (error) {
    logDebug(`insertTableFromGrid(${rows}, ${cols}) failed`, error);
    setStatus("Could not insert table");
  }
}

async function insertTableCaption() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText("Table ", "replace");
      await context.sync();
      
      const newSelection = context.document.getSelection();
      newSelection.insertText(": ", "end");
      await context.sync();
      
      setStatus("Table caption inserted!");
    });
  } catch (error) {
    logDebug("insertTableCaption failed", error);
    setStatus("Could not insert caption");
  }
}

function toggleDropdown(id) {
  const dropdown = document.getElementById(id);
  const isOpen = dropdown.classList.contains("open");
  document.querySelectorAll(".dropdown-content").forEach((d) => d.classList.remove("open"));
  if (!isOpen) {
    dropdown.classList.add("open");
  }
}

async function applyShadingColor(color) {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("No table found");
        return;
      }

      const lastTable = tables.items[tables.items.length - 1];
      lastTable.load(["rows/cells", "rows/cells/shadingColor"]);
      await context.sync();

      for (const row of lastTable.rows.items) {
        row.load(["cells", "cells/shadingColor"]);
        await context.sync();

        if (row.cells && row.cells.items) {
          for (const cell of row.cells.items) {
            cell.load("shadingColor");
            await context.sync();

            if (color === "no-fill") {
              cell.shadingColor = "NoFill";
            } else {
              cell.shadingColor = color;
            }
          }
        }
      }

      await context.sync();
      document.querySelectorAll(".dropdown-content").forEach((d) => d.classList.remove("open"));
      setStatus("Shading applied to table");
    });
  } catch (error) {
    logDebug(`applyShadingColor("${color}") failed`, error);
    setStatus("Could not apply shading");
  }
}

async function applyBorders(borderType) {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("No table found");
        return;
      }

      const lastTable = tables.items[tables.items.length - 1];
      lastTable.load("rows");
      await context.sync();

      for (let rowIdx = 0; rowIdx < lastTable.rows.items.length; rowIdx++) {
        const row = lastTable.rows.items[rowIdx];
        row.load("cells");
        await context.sync();

        for (let cellIdx = 0; cellIdx < row.cells.items.length; cellIdx++) {
          const cell = row.cells.items[cellIdx];
          const borders = cell.format.borders;

          if (borderType === "all") {
            borders.top.visible = !borders.top.visible;
            borders.bottom.visible = !borders.bottom.visible;
            borders.left.visible = !borders.left.visible;
            borders.right.visible = !borders.right.visible;
          } else if (borderType === "outside") {
            borders.top.visible = !borders.top.visible;
            borders.bottom.visible = !borders.bottom.visible;
            borders.left.visible = !borders.left.visible;
            borders.right.visible = !borders.right.visible;
            borders.insideHorizontal.visible = false;
            borders.insideVertical.visible = false;
          } else if (borderType === "inside") {
            borders.insideHorizontal.visible = !borders.insideHorizontal.visible;
            borders.insideVertical.visible = !borders.insideVertical.visible;
          } else if (borderType === "top") {
            borders.top.visible = !borders.top.visible;
          } else if (borderType === "bottom") {
            borders.bottom.visible = !borders.bottom.visible;
          } else if (borderType === "left") {
            borders.left.visible = !borders.left.visible;
          } else if (borderType === "right") {
            borders.right.visible = !borders.right.visible;
          } else if (borderType === "none") {
            borders.top.visible = false;
            borders.bottom.visible = false;
            borders.left.visible = false;
            borders.right.visible = false;
            borders.insideHorizontal.visible = false;
            borders.insideVertical.visible = false;
          }
        }
      }

      await context.sync();
      document.querySelectorAll(".dropdown-content").forEach((d) => d.classList.remove("open"));
      setStatus("Borders applied to table");
    });
  } catch (error) {
    logDebug(`applyBorders("${borderType}") failed`, error);
    setStatus("Could not apply borders");
  }
}

/* ── LAYOUTS ── */
async function insertPageBreak() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertBreak(Word.BreakType.page, "after");
      await context.sync();
      setStatus("Page break inserted!");
    });
  } catch (error) {
    logDebug("insertPageBreak failed", error);
    setStatus("Could not insert page break");
  }
}

function getPageDimensions(type) {
  const A4_WIDTH = 12240;
  const A4_HEIGHT = 15840;
  
  switch (type) {
    case "landscape":
      return { width: A4_HEIGHT, height: A4_WIDTH };
    case "portrait":
      return { width: A4_WIDTH, height: A4_HEIGHT };
    default:
      return { width: A4_HEIGHT, height: A4_WIDTH };
  }
}

async function insertPage(type) {
  const { width, height } = getPageDimensions(type);
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
      await context.sync();
      
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      
      const lastSection = sections.items[sections.items.length - 1];
      lastSection.load("pageSetup");
      await context.sync();
      
      lastSection.pageSetup.pageWidth = width;
      lastSection.pageSetup.pageHeight = height;
      await context.sync();
      
      setStatus("Page inserted!");
    });
  } catch (error) {
    logDebug(`insertPage("${type}") failed`, error);
    setStatus("Could not insert page");
  }
}

/* ── TOOLS ── */
let gridlinesVisible = false;
let paragraphMarksVisible = false;

async function pasteUnformatted() {
  setStatus("Use Ctrl+Shift+V for plain text paste");
}

async function toggleGridlines() {
  try {
    if (Office.actions && Office.actions.invoke) {
      await Office.actions.invoke("ToggleGridlines");
      gridlinesVisible = !gridlinesVisible;
      document.getElementById("gridlinesBtn").textContent = 
        gridlinesVisible ? "Hide Gridlines" : "Show Gridlines";
      setStatus(gridlinesVisible ? "Gridlines shown" : "Gridlines hidden");
    } else {
      setStatus("Gridlines not supported");
    }
  } catch (error) {
    logDebug("toggleGridlines failed", error);
    setStatus("Could not toggle gridlines");
  }
}

async function toggleParagraphMarks() {
  try {
    if (Office.actions && Office.actions.invoke) {
      await Office.actions.invoke("ShowAllFormattingMarks");
      paragraphMarksVisible = !paragraphMarksVisible;
      document.getElementById("paragraphMarksBtn").textContent = 
        paragraphMarksVisible ? "Hide Paragraph Marks" : "Show Paragraph Marks";
      setStatus(paragraphMarksVisible ? "Paragraph marks shown" : "Paragraph marks hidden");
    } else {
      setStatus("Paragraph marks not supported");
    }
  } catch (error) {
    logDebug("toggleParagraphMarks failed", error);
    setStatus("Could not toggle paragraph marks");
  }
}

async function updateTOC() {
  try {
    await Word.run(async (context) => {
      const fields = context.document.fields;
      fields.load("items");
      await context.sync();
      
      let tocCount = 0;
      fields.items.forEach((field) => {
        field.load("type");
        if (field.type === "TOC") {
          tocCount++;
        }
      });
      await context.sync();
      
      if (tocCount > 0) {
        context.document.fields.updateAll();
        await context.sync();
        setStatus(`Updated ${tocCount} table of contents!`);
      } else {
        setStatus("No TOC found in document");
      }
    });
  } catch (error) {
    logDebug("updateTOC failed", error);
    setStatus("Could not update TOC");
  }
}

/* ── DEBUG ── */
let debugEnabled = false;

function logDebug(message, error = null) {
  if (!debugEnabled) return;
  
  const debugContent = document.getElementById("debugContent");
  const debugPanel = document.getElementById("debugPanel");
  
  const entry = document.createElement("div");
  entry.className = "debug-entry";
  
  const time = new Date().toLocaleTimeString();
  let details = `<div class="debug-time">[${time}]</div>`;
  details += `<div class="debug-message">${message}</div>`;
  
  if (error) {
    details += `<div class="debug-details">`;
    details += `<strong>Error:</strong> ${error.message || error}<br>`;
    if (error.stack) {
      const stackLines = error.stack.split("\n").slice(0, 3).join("<br>");
      details += `<strong>Stack:</strong> ${stackLines}`;
    }
    details += `</div>`;
  }
  
  entry.innerHTML = details;
  debugContent.appendChild(entry);
  debugPanel.classList.add("visible");
  debugPanel.scrollTop = debugPanel.scrollHeight;
}

function clearDebug() {
  const debugContent = document.getElementById("debugContent");
  debugContent.innerHTML = "";
}

function toggleDebug() {
  debugEnabled = !debugEnabled;
  const debugPanel = document.getElementById("debugPanel");
  const btn = document.getElementById("toggleDebug");
  
  if (debugEnabled) {
    debugPanel.classList.add("visible");
    btn.style.background = "#dc2626";
    logDebug("Debug mode enabled");
  } else {
    debugPanel.classList.remove("visible");
    btn.style.background = "#6b7280";
    logDebug("Debug mode disabled");
  }
}

/* ── UTILS ── */
function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    setTimeout(() => { el.textContent = ""; }, 3000);
  }
}

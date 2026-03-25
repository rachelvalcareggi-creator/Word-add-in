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
    console.error("applyStyle error:", error);
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
    console.error("applyTableStyle error:", error);
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
    console.error("insertTableFromGrid error:", error);
    setStatus("Could not insert table");
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
      const selection = context.document.getSelection();
      let cellsToShade = [];

      try {
        if (selection.cells) {
          selection.cells.load("items");
          await context.sync();
          if (selection.cells.items && selection.cells.items.length > 0) {
            cellsToShade = selection.cells.items;
          }
        }
      } catch (e) {
        console.log("selection.cells not available");
      }

      if (cellsToShade.length === 0) {
        setStatus("Select table cells first");
        return;
      }

      cellsToShade.forEach((cell) => {
        if (color === "no-fill") {
          cell.format.fill = "NoFill";
        } else {
          cell.format.fill = color;
        }
      });

      await context.sync();
      document.querySelectorAll(".dropdown-content").forEach((d) => d.classList.remove("open"));
      setStatus("Shading applied");
    });
  } catch (error) {
    console.error("applyShadingColor error:", error);
    setStatus("Could not apply shading");
  }
}

async function applyBorders(borderType) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      let cellsToBorder = [];

      try {
        if (selection.cells) {
          selection.cells.load("items");
          await context.sync();
          if (selection.cells.items && selection.cells.items.length > 0) {
            cellsToBorder = selection.cells.items;
          }
        }
      } catch (e) {
        console.log("selection.cells not available");
      }

      if (cellsToBorder.length === 0) {
        setStatus("Select table cells first");
        return;
      }

      cellsToBorder.forEach((cell) => {
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
      });

      await context.sync();
      document.querySelectorAll(".dropdown-content").forEach((d) => d.classList.remove("open"));
      setStatus("Borders applied");
    });
  } catch (error) {
    console.error("applyBorders error:", error);
    setStatus("Could not apply borders");
  }
}

/* ── LAYOUTS ── */
async function insertPage(type) {
  const A4_WIDTH = 12240;
  const A4_HEIGHT = 15840;
  const A3_WIDTH = 15840;
  const A3_HEIGHT = 22320;

  let width, height;

  switch (type) {
    case "landscape":
      width = A4_HEIGHT;
      height = A4_WIDTH;
      break;
    case "portrait":
      width = A4_WIDTH;
      height = A4_HEIGHT;
      break;
    case "a3-landscape":
      width = A3_HEIGHT;
      height = A3_WIDTH;
      break;
    case "a3-portrait":
      width = A3_WIDTH;
      height = A3_HEIGHT;
      break;
    default:
      return;
  }

  try {
    await Word.run(async (context) => {
      const newSection = context.document.addSection();
      newSection.pageWidth = width;
      newSection.pageHeight = height;
      await context.sync();
      setStatus("Page inserted!");
    });
  } catch (error) {
    console.error("insertPage error:", error);
    setStatus("Could not insert page");
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

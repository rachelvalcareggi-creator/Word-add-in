/* taskpane.js — Rachele Tools logic */

const STORAGE_KEY = "racheleToolsSetup";

let selectedCover = null;

Office.onReady(() => {
  initTabs();
  initCoverTab();
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
    titleRange.font.name = "Century Gothic";
    titleRange.font.size = 40;
    titleRange.font.bold = true;

    const subtitleControl = titleRange.insertContentControl();
    subtitleControl.type = "richText";
    subtitleControl.title = "Subtitle";
    subtitleControl.tag = "subtitle";
    subtitleControl.insertText(subtitle, "end");
    const subtitleRange = subtitleControl.getRange();
    subtitleRange.font.name = "Century Gothic";
    subtitleRange.font.size = 24;

    const dateControl = subtitleRange.insertContentControl();
    dateControl.type = "richText";
    dateControl.title = "Date";
    dateControl.tag = "date";
    dateControl.insertText(date, "end");
    const dateRange = dateControl.getRange();
    dateRange.font.name = "Century Gothic";
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
      titleRange.font.name = "Century Gothic";
      titleRange.font.size = 40;
      titleRange.font.bold = true;

      const subtitleControl = titleRange.insertContentControl();
      subtitleControl.type = "richText";
      subtitleControl.title = "Subtitle";
      subtitleControl.tag = "subtitle";
      subtitleControl.insertText(subtitle, "end");
      const subtitleRange = subtitleControl.getRange();
      subtitleRange.font.name = "Century Gothic";
      subtitleRange.font.size = 24;

      const dateControl = subtitleRange.insertContentControl();
      dateControl.type = "richText";
      dateControl.title = "Date";
      dateControl.tag = "date";
      dateControl.insertText(date, "end");
      const dateRange = dateControl.getRange();
      dateRange.font.name = "Century Gothic";
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
        para.font.name = "Century Gothic";
        para.font.nameAscii = "Century Gothic";
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
async function insertCustomTable() {
  const rows = parseInt(document.getElementById("inputRows").value) || 4;
  const cols = parseInt(document.getElementById("inputColumns").value) || 3;
  const headerStyle = document.getElementById("headerStyle").value;

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const table = body.insertTable(rows, cols);
      table.styleBuiltIn = Word.BuiltInStyleName.table.grid;

      if (headerStyle === "heading") {
        table.rows.getItem(0).cells.format.fill = "#4F46E5";
        table.rows.getItem(0).cells.items.forEach((cell) => {
          cell.body.paragraphs.getItem(0).font.color = "white";
          cell.body.paragraphs.getItem(0).font.bold = true;
          cell.body.paragraphs.getItem(0).font.name = "Century Gothic";
        });
      } else if (headerStyle === "text") {
        table.rows.getItem(0).cells.items.forEach((cell) => {
          cell.body.paragraphs.getItem(0).font.bold = true;
          cell.body.paragraphs.getItem(0).font.name = "Century Gothic";
        });
      } else if (headerStyle === "bullets") {
        for (let i = 0; i < rows; i++) {
          table.getCell(i, 0).body.paragraphs.getItem(0).listFormat.apply("bullet");
        }
        table.rows.getItem(0).cells.items.forEach((cell) => {
          cell.body.paragraphs.getItem(0).font.bold = true;
        });
      }

      table.rows.items.forEach((row) => {
        row.cells.items.forEach((cell) => {
          if (headerStyle !== "bullets" || table.rows.indexOf(row) !== 0) {
            cell.body.paragraphs.getItem(0).font.name = "Century Gothic";
          }
        });
      });

      await context.sync();
      setStatus("Table inserted!");
    });
  } catch (error) {
    console.error("insertCustomTable error:", error);
    setStatus("Could not insert table");
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

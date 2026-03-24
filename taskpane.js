/* taskpane.js — Rachele Tools logic */

const STORAGE_KEY = "racheleToolsSetup";

let selectedCover = null;
let userImageBase64 = null;

Office.onReady(() => {
  checkSetup();
});

function checkSetup() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    showMainContent();
  } else {
    showSetupDialog();
  }
}

function showSetupDialog() {
  document.getElementById("setupOverlay").classList.add("open");
  document.getElementById("inputDate").value = new Date().toLocaleDateString();
  selectedCover = 0;
  document.querySelectorAll(".cover-thumb").forEach((el) => {
    el.classList.remove("selected");
  });
  document.querySelector(`[data-cover="0"]`).classList.add("selected");
}

function showMainContent() {
  const overlay = document.getElementById("setupOverlay");
  overlay.classList.remove("open");
  overlay.style.display = "none";
  document.getElementById("mainContent").classList.add("open");
}

function selectCover(num) {
  selectedCover = num;
  document.querySelectorAll(".cover-thumb").forEach((el) => {
    el.classList.remove("selected");
  });
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

    const titlePara = insertPoint.insertParagraph(title, "after");
    titlePara.style = "Title";
    titlePara.font.name = "Century Gothic";
    titlePara.font.size = 40;
    titlePara.font.bold = true;
    titlePara.alignment = Word.Alignment.left;

    const subtitlePara = titlePara.getRange("end").insertParagraph(subtitle, "after");
    subtitlePara.font.name = "Century Gothic";
    subtitlePara.font.size = 24;
    subtitlePara.alignment = Word.Alignment.left;

    const datePara = subtitlePara.getRange("end").insertParagraph(date, "after");
    datePara.font.name = "Century Gothic";
    datePara.font.size = 16;
    datePara.font.color = "#6c757d";
    datePara.alignment = Word.Alignment.left;

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

      const titlePara = insertPoint.insertParagraph(title, "after");
      titlePara.style = "Title";
      titlePara.font.name = "Century Gothic";
      titlePara.font.size = 40;
      titlePara.font.bold = true;
      titlePara.alignment = Word.Alignment.left;

      const subtitlePara = titlePara.getRange("end").insertParagraph(subtitle, "after");
      subtitlePara.font.name = "Century Gothic";
      subtitlePara.font.size = 24;
      subtitlePara.alignment = Word.Alignment.left;

      const datePara = subtitlePara.getRange("end").insertParagraph(date, "after");
      datePara.font.name = "Century Gothic";
      datePara.font.size = 16;
      datePara.font.color = "#6c757d";
      datePara.alignment = Word.Alignment.left;

      const textEnd = datePara.getRange("end");

      const imgInsert = textEnd.insertInlinePictureFromBase64(imgBase64, "after");
      imgInsert.width = pageWidth;
      imgInsert.height = pageHeight;
      imgInsert.lockAspectRatio = false;

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
  document.getElementById("inputTitle").value = "";
  document.getElementById("inputSubtitle").value = "";
  document.getElementById("inputDate").value = "";
  document.querySelectorAll(".cover-thumb").forEach((el) => el.classList.remove("selected"));
  document.getElementById("mainContent").classList.remove("open");
  showSetupDialog();
}

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

function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    setTimeout(() => { el.textContent = ""; }, 3000);
  }
}

async function insertLandscapePage() {
  try {
    await Word.run(async (context) => {
      const currentSection = context.document.sections.getFirst();
      currentSection.load("pageWidth, pageHeight");
      await context.sync();

      let newWidth = currentSection.pageWidth;
      let newHeight = currentSection.pageHeight;

      if (newWidth < newHeight) {
        [newWidth, newHeight] = [newHeight, newWidth];
      }

      const newSection = context.document.addSection();
      newSection.pageWidth = newWidth;
      newSection.pageHeight = newHeight;

      await context.sync();
      setStatus("Landscape page inserted");
    });
  } catch (error) {
    console.error("insertLandscapePage error:", error);
    setStatus("Could not insert landscape page");
  }
}

async function insertTable() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertTable(0, 4, 3);
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length > 0) {
        const table = tables.items[0];
        table.styleBuiltIn = Word.BuiltInStyleName.table.grid;
        table.getCell(0, 0).body.paragraphs.getItem(0).text = "Column 1";
        table.getCell(0, 1).body.paragraphs.getItem(0).text = "Column 2";
        table.getCell(0, 2).body.paragraphs.getItem(0).text = "Column 3";

        table.rows.getItem(0).cells.format.fill = "#4F46E5";
        table.rows.getItem(0).cells.items.forEach((cell) => {
          cell.body.paragraphs.getItem(0).font.color = "white";
          cell.body.paragraphs.getItem(0).font.bold = true;
          cell.body.paragraphs.getItem(0).font.name = "Century Gothic";
        });
        await context.sync();
      }

      setStatus("Table inserted");
    });
  } catch (error) {
    console.error("insertTable error:", error);
    setStatus("Could not insert table");
  }
}

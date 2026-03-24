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

    const titleControl = insertPoint.insertContentControl();
    titleControl.type = "richText";
    titleControl.title = "Title";
    titleControl.tag = "title";
    titleControl.insertText(title, "end");
    const titleRange = titleControl.getRange();
    titleRange.paragraphFormat.styleBuiltIn = "Title";
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

      const coverControl = insertPoint.insertContentControl();
      coverControl.type = "picture";
      coverControl.title = "Cover Image";
      coverControl.tag = "coverImage";
      coverControl.insertInlinePictureFromBase64(imgBase64, "end");
      const coverRange = coverControl.getRange();
      coverRange.inlinePicture.width = pageWidth;
      coverRange.inlinePicture.height = pageHeight;

      await context.sync();

      const placeholderControl = coverRange.insertContentControl();
      placeholderControl.type = "picture";
      placeholderControl.title = "Add Your Image";
      placeholderControl.tag = "userImagePlaceholder";
      placeholderControl.placeholderText = "Click here to add your image";

      const titleControl = placeholderControl.getRange().insertContentControl();
      titleControl.type = "richText";
      titleControl.title = "Title";
      titleControl.tag = "title";
      titleControl.insertText(title, "end");
      const titleRange = titleControl.getRange();
      titleRange.paragraphFormat.styleBuiltIn = "Title";
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
  document.getElementById("inputTitle").value = "";
  document.getElementById("inputSubtitle").value = "";
  document.getElementById("inputDate").value = "";
  document.querySelectorAll(".cover-thumb").forEach((el) => el.classList.remove("selected"));
  document.getElementById("mainContent").classList.remove("open");
  showSetupDialog();
}

function addImageToCover() {
  document.getElementById("coverImageInput").click();
}

function handleCoverImageUpload(input) {
  const file = input.files[0];
  if (!file) return;

  setStatus("Adding image...");

  const reader = new FileReader();
  reader.onload = function (e) {
    const base64 = e.target.result.split(",")[1];

    Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (paragraphs.items.length > 2) {
        const insertRange = paragraphs.items[1].getRange("end");
        const img = insertRange.insertInlinePictureFromBase64(base64, "after");
        img.width = 5040;
        img.height = 3360;
        img.lockAspectRatio = false;
        await context.sync();
        setStatus("Image added!");
      }
    }).catch((error) => {
      console.error("handleCoverImageUpload error:", error);
      setStatus("Error adding image");
    });
  };
  reader.readAsDataURL(file);
  input.value = "";
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

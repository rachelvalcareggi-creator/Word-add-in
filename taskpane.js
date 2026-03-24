/* taskpane.js — Rachele Tools logic */

Office.onReady(() => {
  // Ready
});

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
      const selection = context.document.getSelection();
      const table = selection.insertTable(4, 3);

      table.styleBuiltIn = Word.BuiltInStyleName.table.grid;
      table.getCell(0, 0).body.text = "Column 1";
      table.getCell(0, 1).body.text = "Column 2";
      table.getCell(0, 2).body.text = "Column 3";

      table.rows.getItem(0).cells.format.fill = "#4F46E5";
      table.rows.getItem(0).cells.items.forEach((cell) => {
        cell.body.paragraphs.getItem(0).font.color = "white";
        cell.body.paragraphs.getItem(0).font.bold = true;
      });

      await context.sync();
      setStatus("Table inserted");
    });
  } catch (error) {
    console.error("insertTable error:", error);
    setStatus("Could not insert table");
  }
}

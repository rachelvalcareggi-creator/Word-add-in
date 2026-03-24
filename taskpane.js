/* taskpane.js — Rachele Tools logic */

Office.onReady(() => {
  renderStylesDropdown();
});

function toggleStyles() {
  const btn = document.getElementById("stylesBtn");
  const dropdown = document.getElementById("stylesDropdown");
  
  btn.classList.toggle("active");
  dropdown.classList.toggle("open");
}

function renderStylesDropdown() {
  const stylesList = document.getElementById("stylesList");
  const loading = document.getElementById("stylesLoading");
  
  if (loading) loading.style.display = "none";
  
  stylesList.innerHTML = `
    <div class="dropdown-section">Headings</div>
    <button class="dropdown-item" onclick="applyStyle('Heading 1')">Heading 1</button>
    <button class="dropdown-item" onclick="applyStyle('Heading 2')">Heading 2</button>
    <button class="dropdown-item" onclick="applyStyle('Heading 3')">Heading 3</button>
    
    <div class="dropdown-section">Bullets</div>
    <button class="dropdown-item" onclick="applyStyle('List Bullet')">Bullet 1</button>
    <button class="dropdown-item" onclick="applyStyle('List Bullet 2')">Bullet 2</button>
    <button class="dropdown-item" onclick="applyStyle('List Bullet 3')">Bullet 3</button>
    
    <div class="dropdown-section">Numbering</div>
    <button class="dropdown-item" onclick="applyStyle('List Number')">Number 1</button>
    <button class="dropdown-item" onclick="applyStyle('List Number 2')">Number 2</button>
    <button class="dropdown-item" onclick="applyStyle('List Number 3')">Number 3</button>
  `;
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

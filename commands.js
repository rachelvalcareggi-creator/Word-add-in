/* commands.js — ribbon button functions */

Office.onReady(() => {
  // Register all ribbon button functions on the global scope
  Office.actions.associate("insertLandscapePage", insertLandscapePage);
  Office.actions.associate("insertTable", insertTable);
});

// ─────────────────────────────────────────────
// INSERT LANDSCAPE PAGE
// Inserts a section break before and after a
// landscape-oriented page at the cursor position.
// ─────────────────────────────────────────────
async function insertLandscapePage(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // Insert section break (next page) before cursor
      selection.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.before);

      // Get the paragraph after the break and set page orientation
      // We use OOXML to set landscape orientation for this section
      const landscapeOoxml = `
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
          <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData>
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
              </Relationships>
            </pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
            <pkg:xmlData>
              <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                  <w:p>
                    <w:pPr>
                      <w:sectPr>
                        <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
                        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
                      </w:sectPr>
                    </w:pPr>
                  </w:p>
                  <w:sectPr>
                    <w:pgSz w:w="12240" w:h="15840"/>
                  </w:sectPr>
                </w:body>
              </w:document>
            </pkg:xmlData>
          </pkg:part>
        </pkg:package>`;

      selection.insertOoxml(landscapeOoxml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error("insertLandscapePage error:", error);
  } finally {
    event.completed();
  }
}

// ─────────────────────────────────────────────
// INSERT TABLE
// Inserts a styled 3-column x 4-row table
// at the cursor. Easy to customise rows/cols.
// ─────────────────────────────────────────────
async function insertTable(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // ── Customise these ──
      const ROWS = 4;
      const COLS = 3;
      const HEADERS = ["Column 1", "Column 2", "Column 3"];
      // ─────────────────────

      // Build initial values array (rows x cols)
      const values = [];
      for (let r = 0; r < ROWS; r++) {
        const row = [];
        for (let c = 0; c < COLS; c++) {
          row.push(r === 0 ? HEADERS[c] || `Col ${c + 1}` : "");
        }
        values.push(row);
      }

      const table = selection.insertTable(ROWS, COLS, Word.InsertLocation.after, values);

      // Style the table
      table.styleBuiltIn = Word.Style.tableGrid;

      // Bold the header row
      table.getRow(0).font.bold = true;
      table.getRow(0).shadingColor = "#4F46E5"; // indigo header
      table.getRow(0).font.color = "#FFFFFF";

      await context.sync();
    });
  } catch (error) {
    console.error("insertTable error:", error);
  } finally {
    event.completed();
  }
}

/* commands.js — ribbon button functions */

Office.onReady(() => {
  Office.actions.associate("insertLandscapePage", insertLandscapePage);
  Office.actions.associate("insertTable", insertTable);
  Office.actions.associate("toggleGridlines", toggleGridlines);
  Office.actions.associate("toggleParagraphs", toggleParagraphs);
});

// ─────────────────────────────────────────────
// INSERT LANDSCAPE PAGE
// Inserts a single landscape page using OOXML
// section properties inserted at the cursor.
// ─────────────────────────────────────────────
async function insertLandscapePage(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // OOXML that inserts:
      //   1. A section break ending the current (portrait) section
      //   2. An empty landscape page
      //   3. A section break resuming portrait
      const ooxml = `
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels"
    pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml"
    pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <!-- Paragraph ending the portrait section -->
          <w:p>
            <w:pPr>
              <w:sectPr>
                <w:pgSz w:w="12240" w:h="15840"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                         w:header="708" w:footer="708" w:gutter="0"/>
                <w:type w:val="nextPage"/>
              </w:sectPr>
            </w:pPr>
          </w:p>
          <!-- Paragraph on the landscape page -->
          <w:p>
            <w:pPr>
              <w:sectPr>
                <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                         w:header="708" w:footer="708" w:gutter="0"/>
                <w:type w:val="nextPage"/>
              </w:sectPr>
            </w:pPr>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

      selection.insertOoxml(ooxml, Word.InsertLocation.after);
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
// Inserts a clean formatted table at the cursor.
// ─────────────────────────────────────────────
async function insertTable(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // ── Customise these per client ──
      const ROWS = 4;
      const COLS = 3;
      const HEADERS = ["Column 1", "Column 2", "Column 3"];
      // ────────────────────────────────

      // Build values array
      const values = [];
      for (let r = 0; r < ROWS; r++) {
        const row = [];
        for (let c = 0; c < COLS; c++) {
          row.push(r === 0 ? (HEADERS[c] || `Col ${c + 1}`) : "");
        }
        values.push(row);
      }

      // Insert the table after the selection
      const table = selection.insertTable(ROWS, COLS, Word.InsertLocation.after, values);
      table.load("style");
      await context.sync();
      
      // Apply Grid Table 4 - Accent 1 style
      table.style = "GridTable4_Accent1";
      await context.sync();

      // Style the header row after sync
      const headerRow = table.getRow(0);
      headerRow.load("cells");
      await context.sync();

      // Bold header text
      headerRow.font.bold = true;

      await context.sync();
    });
  } catch (error) {
    console.error("insertTable error:", error);
  } finally {
    event.completed();
  }
}

// ─────────────────────────────────────────────
// TOGGLE GRIDLINES
// Toggles table gridlines visibility.
// Works in Word Desktop (WordApiDesktop 1.4)
// Shows message in Word Online
// ─────────────────────────────────────────────
async function toggleGridlines(event) {
  try {
    if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.4")) {
      await Word.run(async (context) => {
        const view = context.document.getView();
        view.load("areTableGridlinesDisplayed");
        await context.sync();
        view.areTableGridlinesDisplayed = !view.areTableGridlinesDisplayed;
        await context.sync();
      });
    } else {
      showToastMessage("Gridlines not supported in Word Online");
    }
  } catch (error) {
    console.error("toggleGridlines error:", error);
  } finally {
    event.completed();
  }
}

// ─────────────────────────────────────────────
// TOGGLE PARAGRAPHS
// Toggles paragraph marks visibility.
// Works in Word Desktop (WordApiDesktop 1.4)
// Shows message in Word Online
// ─────────────────────────────────────────────
async function toggleParagraphs(event) {
  try {
    if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.4")) {
      await Word.run(async (context) => {
        const view = context.document.getView();
        view.load("areAllNonprintingCharactersDisplayed");
        await context.sync();
        view.areAllNonprintingCharactersDisplayed = !view.areAllNonprintingCharactersDisplayed;
        await context.sync();
      });
    } else {
      showToastMessage("Use Ctrl+Shift+8 for paragraph marks");
    }
  } catch (error) {
    console.error("toggleParagraphs error:", error);
  } finally {
    event.completed();
  }
}

// ─────────────────────────────────────────────
// TOAST MESSAGE HELPER
// Shows a toast notification (used by ribbon buttons)
// ─────────────────────────────────────────────
function showToastMessage(message) {
  if (typeof showToast === "function") {
    showToast(message);
  }
}

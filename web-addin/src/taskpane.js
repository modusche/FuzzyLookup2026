import {
  combinedSimilarity,
  cleanText,
  applyTransformations,
} from "./algorithms.js";

/* global Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnGo").addEventListener("click", runFuzzyLookup);
    document.getElementById("btnClose").addEventListener("click", () => {
      Office.addin.hide();
    });
    document.getElementById("scrThreshold").addEventListener("input", (e) => {
      document.getElementById("lblThreshVal").textContent = (
        e.target.value / 100
      ).toFixed(2);
    });
    populateTables();
    document
      .getElementById("cboLeftTable")
      .addEventListener("change", () =>
        populateColumns("cboLeftTable", "cboLeftMatchCol")
      );
    document
      .getElementById("cboRightTable")
      .addEventListener("change", () =>
        populateColumns("cboRightTable", "cboRightMatchCol")
      );
  }
});

async function populateTables() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const leftSel = document.getElementById("cboLeftTable");
      const rightSel = document.getElementById("cboRightTable");
      leftSel.innerHTML = "";
      rightSel.innerHTML = "";

      // Add tables
      for (const sheet of sheets.items) {
        const tables = sheet.tables;
        tables.load("items/name");
        await context.sync();

        for (const table of tables.items) {
          const opt1 = new Option(
            `${table.name}  [Table - ${sheet.name}]`,
            `table:${sheet.name}:${table.name}`
          );
          const opt2 = new Option(
            `${table.name}  [Table - ${sheet.name}]`,
            `table:${sheet.name}:${table.name}`
          );
          leftSel.add(opt1);
          rightSel.add(opt2);
        }
      }

      // Add sheets with data
      for (const sheet of sheets.items) {
        const used = sheet.getUsedRangeOrNullObject();
        used.load("rowCount");
        await context.sync();

        if (!used.isNullObject && used.rowCount > 1) {
          const opt1 = new Option(
            `${sheet.name}  [Sheet]`,
            `sheet:${sheet.name}`
          );
          const opt2 = new Option(
            `${sheet.name}  [Sheet]`,
            `sheet:${sheet.name}`
          );
          leftSel.add(opt1);
          rightSel.add(opt2);
        }
      }

      if (leftSel.options.length > 0) leftSel.selectedIndex = 0;
      if (rightSel.options.length > 1) rightSel.selectedIndex = 1;
      else if (rightSel.options.length > 0) rightSel.selectedIndex = 0;

      await populateColumns("cboLeftTable", "cboLeftMatchCol");
      await populateColumns("cboRightTable", "cboRightMatchCol");
    });
  } catch (err) {
    showStatus(`Error loading tables: ${err.message}`, true);
  }
}

async function getRange(value) {
  let range = null;
  await Excel.run(async (context) => {
    if (value.startsWith("table:")) {
      const [, sheetName, tableName] = value.split(":");
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const table = sheet.tables.getItem(tableName);
      range = table.getRange();
    } else if (value.startsWith("sheet:")) {
      const sheetName = value.split(":")[1];
      const sheet = context.workbook.worksheets.getItem(sheetName);
      range = sheet.getUsedRange();
    }
    if (range) {
      range.load("values,rowCount,columnCount");
      await context.sync();
    }
  });
  return range;
}

async function populateColumns(tableSelectId, colSelectId) {
  const tableSel = document.getElementById(tableSelectId);
  const colSel = document.getElementById(colSelectId);
  colSel.innerHTML = "";

  if (tableSel.selectedIndex < 0) return;

  try {
    await Excel.run(async (context) => {
      const value = tableSel.value;
      let range;

      if (value.startsWith("table:")) {
        const [, sheetName, tableName] = value.split(":");
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const table = sheet.tables.getItem(tableName);
        range = table.getRange();
      } else if (value.startsWith("sheet:")) {
        const sheetName = value.split(":")[1];
        const sheet = context.workbook.worksheets.getItem(sheetName);
        range = sheet.getUsedRange();
      }

      if (range) {
        range.load("values,columnCount");
        await context.sync();

        for (let c = 0; c < range.columnCount; c++) {
          const hdr = range.values[0][c] || `Column ${c + 1}`;
          colSel.add(new Option(String(hdr), c));
        }
        if (colSel.options.length > 0) colSel.selectedIndex = 0;
      }
    });
  } catch (err) {
    console.error("populateColumns error:", err);
  }
}

function parseTransformRules(text) {
  if (!text.trim()) return [];
  return text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.includes("=>"))
    .map((line) => {
      const [from, ...rest] = line.split("=>");
      return { from: from.trim(), to: rest.join("=>").trim() };
    });
}

async function runFuzzyLookup() {
  const leftTableVal = document.getElementById("cboLeftTable").value;
  const rightTableVal = document.getElementById("cboRightTable").value;
  const leftColIdx = parseInt(document.getElementById("cboLeftMatchCol").value);
  const rightColIdx = parseInt(
    document.getElementById("cboRightMatchCol").value
  );
  const threshold = parseInt(document.getElementById("scrThreshold").value) / 100;
  const maxMatches =
    Math.max(1, parseInt(document.getElementById("txtMaxMatches").value)) || 1;
  const outputName =
    document.getElementById("txtOutputSheet").value.trim() || "Fuzzy_Results";
  const doTrim = document.getElementById("chkTrim").checked;
  const doLower = document.getElementById("chkLower").checked;
  const doRemovePunct = document.getElementById("chkRemovePunct").checked;
  const rules = parseTransformRules(
    document.getElementById("txtTransforms").value
  );

  if (!leftTableVal || !rightTableVal) {
    showStatus("Select both tables.", true);
    return;
  }
  if (isNaN(leftColIdx) || isNaN(rightColIdx)) {
    showStatus("Select match columns.", true);
    return;
  }

  const btn = document.getElementById("btnGo");
  btn.disabled = true;
  btn.textContent = "Processing...";
  showStatus("Reading data...");

  try {
    await Excel.run(async (context) => {
      // Read left and right data
      let leftRange, rightRange;

      if (leftTableVal.startsWith("table:")) {
        const [, sn, tn] = leftTableVal.split(":");
        leftRange = context.workbook.worksheets
          .getItem(sn)
          .tables.getItem(tn)
          .getRange();
      } else {
        leftRange = context.workbook.worksheets
          .getItem(leftTableVal.split(":")[1])
          .getUsedRange();
      }

      if (rightTableVal.startsWith("table:")) {
        const [, sn, tn] = rightTableVal.split(":");
        rightRange = context.workbook.worksheets
          .getItem(sn)
          .tables.getItem(tn)
          .getRange();
      } else {
        rightRange = context.workbook.worksheets
          .getItem(rightTableVal.split(":")[1])
          .getUsedRange();
      }

      leftRange.load("values,rowCount,columnCount");
      rightRange.load("values,rowCount,columnCount");
      await context.sync();

      const leftData = leftRange.values;
      const rightData = rightRange.values;
      const totalLeft = leftData.length - 1;
      const totalRight = rightData.length - 1;
      const leftCols = leftData[0].length;
      const rightCols = rightData[0].length;

      if (totalLeft < 1 || totalRight < 1) {
        showStatus("Tables need at least a header and one data row.", true);
        return;
      }

      showStatus(`Preparing ${totalLeft} x ${totalRight} comparisons...`);
      await new Promise((r) => setTimeout(r, 10));

      // Pre-cache cleaned values
      const leftCache = new Array(totalLeft);
      const rightCache = new Array(totalRight);

      for (let i = 0; i < totalLeft; i++) {
        let val = String(leftData[i + 1][leftColIdx] ?? "");
        val = cleanText(val, doTrim, doLower, doRemovePunct);
        if (rules.length > 0) val = applyTransformations(val, rules);
        leftCache[i] = val;
      }
      for (let i = 0; i < totalRight; i++) {
        let val = String(rightData[i + 1][rightColIdx] ?? "");
        val = cleanText(val, doTrim, doLower, doRemovePunct);
        if (rules.length > 0) val = applyTransformations(val, rules);
        rightCache[i] = val;
      }

      // Match
      const startTime = performance.now();
      const results = [];

      for (let lr = 0; lr < totalLeft; lr++) {
        if (lr % 50 === 0) {
          showStatus(
            `Matching: ${lr}/${totalLeft} (${Math.round((lr / totalLeft) * 100)}%)`
          );
          await new Promise((r) => setTimeout(r, 0));
        }

        const leftVal = leftCache[lr];
        if (leftVal.length === 0) continue;

        const matches = [];
        for (let rr = 0; rr < totalRight; rr++) {
          if (rightCache[rr].length === 0) continue;
          const score = combinedSimilarity(leftVal, rightCache[rr]);
          if (score >= threshold) {
            matches.push({ row: rr, score });
          }
        }

        matches.sort((a, b) => b.score - a.score);
        const top = matches.slice(0, maxMatches);

        for (const m of top) {
          const row = [];
          for (let c = 0; c < leftCols; c++) row.push(leftData[lr + 1][c]);
          for (let c = 0; c < rightCols; c++) row.push(rightData[m.row + 1][c]);
          row.push(Math.round(m.score * 10000) / 10000);
          results.push(row);
        }
      }

      const elapsed = ((performance.now() - startTime) / 1000).toFixed(1);
      showStatus(`Writing ${results.length} results...`);

      // Write output
      let outSheet = context.workbook.worksheets.getItemOrNullObject(outputName);
      await context.sync();
      if (outSheet.isNullObject) {
        outSheet = context.workbook.worksheets.add(outputName);
      } else {
        outSheet.getRange().clear();
      }
      await context.sync();

      // Headers
      const totalOutCols = leftCols + rightCols + 1;
      const headers = [];
      for (let c = 0; c < leftCols; c++)
        headers.push(leftData[0][c] || `Left.Col${c + 1}`);
      for (let c = 0; c < rightCols; c++)
        headers.push(rightData[0][c] || `Right.Col${c + 1}`);
      headers.push("Similarity");

      const headerRange = outSheet.getRangeByIndexes(0, 0, 1, totalOutCols);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;

      // Style left headers
      const leftHdrRange = outSheet.getRangeByIndexes(0, 0, 1, leftCols);
      leftHdrRange.format.fill.color = "#4472C4";
      leftHdrRange.format.font.color = "#FFFFFF";

      // Style right headers
      const rightHdrRange = outSheet.getRangeByIndexes(0, leftCols, 1, rightCols);
      rightHdrRange.format.fill.color = "#2F75B5";
      rightHdrRange.format.font.color = "#FFFFFF";

      // Style similarity header
      const simHdrRange = outSheet.getRangeByIndexes(0, totalOutCols - 1, 1, 1);
      simHdrRange.format.fill.color = "#00B050";
      simHdrRange.format.font.color = "#FFFFFF";

      // Bulk write results
      if (results.length > 0) {
        const dataRange = outSheet.getRangeByIndexes(
          1, 0, results.length, totalOutCols
        );
        dataRange.values = results;

        // Format similarity column as percentage
        const simCol = outSheet.getRangeByIndexes(
          1, totalOutCols - 1, results.length, 1
        );
        simCol.numberFormat = [["0.00%"]];
        simCol.format.font.bold = true;
      }

      // Autofit & freeze
      outSheet.getUsedRange().format.autofitColumns();
      outSheet.freezePanes.freezeRows(1);
      outSheet.activate();

      await context.sync();

      showStatus(
        `Done! ${totalLeft} x ${totalRight} rows, ${results.length} matches in ${elapsed}s`
      );
    });
  } catch (err) {
    showStatus(`Error: ${err.message}`, true);
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Go!";
  }
}

function showStatus(msg, isError = false) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = isError ? "status error" : "status";
}

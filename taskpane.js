/* ============================================================
   Cell Protector — Taskpane Logic
   Office.js Excel Add-in
   Version: 1.0.0
   ============================================================ */

"use strict";

/* ─── Constants ─────────────────────────────────────────── */
const META_SHEET_NAME   = "_CellProtector";
const HIGHLIGHT_COLOR   = "#CCFFCC"; // Light green for editable cells
const META_KEY_HEADER   = "Address";
const META_VALUE_HEADER = "OriginalFill";

/* ─── DOM refs ──────────────────────────────────────────── */
let els = {};

/* ─── State ─────────────────────────────────────────────── */
let _unprotectNeedsPassword = false;

/* ============================================================
   Initialisation
   ============================================================ */
Office.onReady(() => {
  cacheElements();
  bindEvents();
});

/**
 * Cache all DOM elements once at startup.
 */
function cacheElements() {
  els = {
    // Protect card
    radioRange:            document.querySelector('input[name="protectionMode"][value="range"]'),
    radioFormulas:         document.querySelector('input[name="protectionMode"][value="formulas"]'),
    passwordToggle:        document.getElementById("passwordToggle"),
    passwordSection:       document.getElementById("passwordSection"),
    passwordInput:         document.getElementById("passwordInput"),
    confirmPasswordInput:  document.getElementById("confirmPasswordInput"),
    highlightCheckbox:     document.getElementById("highlightCheckbox"),
    protectBtn:            document.getElementById("protectBtn"),

    // Unprotect card
    detectProtectBtn:           document.getElementById("detectProtectBtn"),
    unprotectPasswordSection:   document.getElementById("unprotectPasswordSection"),
    unprotectPasswordInput:     document.getElementById("unprotectPasswordInput"),
    unprotectConfirmBtn:        document.getElementById("unprotectConfirmBtn"),

    // Status
    statusBanner:          document.getElementById("statusBanner"),
    statusIcon:            document.getElementById("statusIcon"),
    statusText:            document.getElementById("statusText"),
  };
}

/**
 * Attach all event listeners.
 */
function bindEvents() {
  // Password toggle visibility
  els.passwordToggle.addEventListener("change", () => {
    toggleVisibility(els.passwordSection, els.passwordToggle.checked);
    if (!els.passwordToggle.checked) {
      els.passwordInput.value        = "";
      els.confirmPasswordInput.value = "";
    }
  });

  // Main protect button
  els.protectBtn.addEventListener("click", handleProtect);

  // Detect & unprotect
  els.detectProtectBtn.addEventListener("click", handleDetectAndUnprotect);

  // Confirm unprotect with password
  els.unprotectConfirmBtn.addEventListener("click", handleUnprotectWithPassword);

  // Allow Enter key in unprotect password field
  els.unprotectPasswordInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleUnprotectWithPassword();
  });
}

/* ============================================================
   Protect Flow
   ============================================================ */

/**
 * Main protect handler — validates inputs then routes to the
 * correct protection strategy.
 */
async function handleProtect() {
  clearStatus();

  const mode          = els.radioFormulas.checked ? "formulas" : "range";
  const usePassword   = els.passwordToggle.checked;
  const doHighlight   = els.highlightCheckbox.checked;

  // Validate passwords if enabled
  if (usePassword) {
    const pw  = els.passwordInput.value;
    const cpw = els.confirmPasswordInput.value;
    if (!pw) return showStatus("error", "⚠", "Please enter a password.");
    if (pw !== cpw) return showStatus("error", "✕", "Passwords do not match.");
  }

  const password = usePassword ? els.passwordInput.value : null;

  setButtonLoading(els.protectBtn, true);

  try {
    if (mode === "range") {
      await protectRange(password, doHighlight);
    } else {
      await protectFormulas(password, doHighlight);
    }
  } catch (err) {
    handleOfficejsError(err, "protect");
  } finally {
    setButtonLoading(els.protectBtn, false);
  }
}

/**
 * Lock the user-selected range and unlock everything else.
 * @param {string|null} password
 * @param {boolean} doHighlight
 */
async function protectRange(password, doHighlight) {
  await Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getActiveWorksheet();
    const selection = context.workbook.getSelectedRange();

    selection.load(["address", "rowCount", "columnCount"]);
    sheet.load("protection/protected");
    await context.sync();

    if (sheet.protection.protected) {
      showStatus("warning", "⚠", "This worksheet is already protected. Unprotect it first.");
      return;
    }

    if (selection.rowCount === 0 || selection.columnCount === 0) {
      showStatus("error", "⚠", "Please select a range first.");
      return;
    }

    // 1. Unlock all cells in the sheet
    const usedRange = sheet.getUsedRange(true);
    usedRange.load(["address"]);
    await context.sync();

    usedRange.format.protection.locked = false;

    // 2. Lock the selected range
    selection.format.protection.locked = true;

    // 3. Highlight editable (unlocked) cells if requested
    if (doHighlight) {
      await storeAndHighlightEditable(context, sheet, selection, "range");
    }

    // 4. Protect the sheet
    sheet.protection.protect(buildProtectionOptions(password));

    await context.sync();

    showStatus("success", "✓", "Cells protected successfully.");
  });
}

/**
 * Lock all formula cells and unlock everything else.
 * @param {string|null} password
 * @param {boolean} doHighlight
 */
async function protectFormulas(password, doHighlight) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    await context.sync();

    if (sheet.protection.protected) {
      showStatus("warning", "⚠", "This worksheet is already protected. Unprotect it first.");
      return;
    }

    // Unlock everything first
    const usedRange = sheet.getUsedRange(true);
    usedRange.load(["address"]);
    await context.sync();

    usedRange.format.protection.locked = false;
    await context.sync();

    // Find formula cells using specialCells
    let formulaRange = null;
    try {
      formulaRange = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
      formulaRange.load(["address", "cellCount"]);
      await context.sync();
    } catch (_) {
      // No formula cells found
      showStatus("warning", "⚠", "No formulas were found in this worksheet.");
      return;
    }

    if (!formulaRange || formulaRange.cellCount === 0) {
      showStatus("warning", "⚠", "No formulas were found in this worksheet.");
      return;
    }

    // Lock formula cells
    formulaRange.format.protection.locked = true;

    // Highlight editable cells if requested
    if (doHighlight) {
      await storeAndHighlightEditable(context, sheet, formulaRange, "formulas");
    }

    // Protect
    sheet.protection.protect(buildProtectionOptions(password));

    await context.sync();

    showStatus("success", "✓", "Formula cells protected successfully.");
  });
}

/* ============================================================
   Highlight & Metadata
   ============================================================ */

/**
 * For each editable (unlocked) cell in the used range, store
 * its original fill color in the hidden metadata sheet, then
 * apply the highlight color.
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {Excel.Range} lockedRange  — the range that was locked
 * @param {"range"|"formulas"} mode
 */
async function storeAndHighlightEditable(context, sheet, lockedRange, mode) {
  const usedRange = sheet.getUsedRange(true);
  usedRange.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  // We'll iterate cell-by-cell (small sheets); for large sheets this
  // can be optimized further.
  const rows    = usedRange.rowCount;
  const cols    = usedRange.columnCount;
  const startRow = usedRange.rowIndex;
  const startCol = usedRange.columnIndex;

  // Collect cell data in batch
  const cells = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = sheet.getCell(startRow + r, startCol + c);
      cell.load(["address", "format/fill/color", "format/protection/locked"]);
      cells.push(cell);
    }
  }
  await context.sync();

  // Build metadata rows and apply highlights
  const metaRows = []; // [address, originalFill]
  for (const cell of cells) {
    if (!cell.format.protection.locked) {
      // Store original fill (empty string means no fill/transparent)
      const origFill = cell.format.fill.color || "";
      metaRows.push([cell.address, origFill]);
      // Apply highlight
      cell.format.fill.color = HIGHLIGHT_COLOR;
    }
  }

  await context.sync();

  if (metaRows.length === 0) return;

  // Write metadata to hidden sheet
  await writeMetadataSheet(context, sheet, metaRows);
}

/**
 * Create (or overwrite) the hidden _CellProtector sheet and write
 * the cell-address → original-fill mapping.
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} parentSheet
 * @param {Array<[string,string]>} rows
 */
async function writeMetadataSheet(context, parentSheet, rows) {
  // Remove old meta sheet if it exists
  await removeMetaSheet(context);

  const metaSheet = context.workbook.worksheets.add(META_SHEET_NAME);
  metaSheet.visibility = Excel.SheetVisibility.hidden;

  // Write headers
  metaSheet.getRange("A1").values = [[META_KEY_HEADER]];
  metaSheet.getRange("B1").values = [[META_VALUE_HEADER]];

  // Write data rows
  if (rows.length > 0) {
    const dataRange = metaSheet.getRangeByIndexes(1, 0, rows.length, 2);
    dataRange.values = rows;
  }

  await context.sync();
}

/**
 * Remove the hidden metadata sheet if it exists.
 */
async function removeMetaSheet(context) {
  try {
    const metaSheet = context.workbook.worksheets.getItem(META_SHEET_NAME);
    metaSheet.delete();
    await context.sync();
  } catch (_) {
    // Sheet didn't exist — no-op
  }
}

/* ============================================================
   Unprotect Flow
   ============================================================ */

/**
 * Detects sheet protection status and either unprotects directly
 * (no password) or prompts for a password.
 */
async function handleDetectAndUnprotect() {
  clearStatus();
  hideUnprotectPasswordSection();
  _unprotectNeedsPassword = false;
  setButtonLoading(els.detectProtectBtn, true);

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["protection/protected", "name"]);
      await context.sync();

      if (!sheet.protection.protected) {
        showStatus("warning", "⚠", "This worksheet is not protected.");
        return;
      }

      // Try unprotecting without a password first
      try {
        sheet.protection.unprotect();
        await context.sync();

        // Restore highlights if we have metadata
        await restoreHighlightsIfPresent(context, sheet);
        showStatus("success", "✓", "Worksheet unprotected.");
        hideUnprotectPasswordSection();
      } catch (err) {
        // If it fails, assume password is required
        _unprotectNeedsPassword = true;
        showUnprotectPasswordSection();
        showStatus("warning", "🔑", "This sheet has a password. Enter it below.");
      }
    });
  } catch (err) {
    handleOfficejsError(err, "detect");
  } finally {
    setButtonLoading(els.detectProtectBtn, false);
  }
}

/**
 * Unprotect using the password entered by the user.
 */
async function handleUnprotectWithPassword() {
  const password = els.unprotectPasswordInput.value;
  if (!password) return showStatus("error", "⚠", "Please enter the sheet password.");

  clearStatus();
  setButtonLoading(els.unprotectConfirmBtn, true);

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("protection/protected");
      await context.sync();

      try {
        sheet.protection.unprotect(password);
        await context.sync();
      } catch (_) {
        showStatus("error", "✕", "Incorrect password. Please try again.");
        return;
      }

      // Restore highlights
      await restoreHighlightsIfPresent(context, sheet);

      hideUnprotectPasswordSection();
      els.unprotectPasswordInput.value = "";
      showStatus("success", "✓", "Worksheet unprotected and editable cells restored.");
    });
  } catch (err) {
    handleOfficejsError(err, "unprotect");
  } finally {
    setButtonLoading(els.unprotectConfirmBtn, false);
  }
}

/**
 * If the metadata sheet exists, restore each cell's original fill color,
 * then delete the metadata sheet.
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 */
async function restoreHighlightsIfPresent(context, sheet) {
  let metaSheet;
  try {
    metaSheet = context.workbook.worksheets.getItem(META_SHEET_NAME);
    metaSheet.load("name");
    await context.sync();
  } catch (_) {
    return; // No metadata — nothing to restore
  }

  // Read the data rows from columns A & B (skip header row 1)
  const usedRange = metaSheet.getUsedRange();
  usedRange.load("values");
  await context.sync();

  const values = usedRange.values; // [[header_a, header_b], [addr, color], ...]

  if (!values || values.length <= 1) {
    metaSheet.delete();
    await context.sync();
    return;
  }

  // Apply stored colors (rows index 1 onwards)
  for (let i = 1; i < values.length; i++) {
    const [addr, origFill] = values[i];
    if (!addr) continue;
    try {
      const cell = sheet.getRange(addr);
      if (origFill && origFill !== "") {
        cell.format.fill.color = origFill;
      } else {
        cell.format.fill.clear();
      }
    } catch (_) {
      // Cell address might be invalid — skip
    }
  }

  // Delete the metadata sheet
  metaSheet.delete();
  await context.sync();

  showStatus("success", "✓", "Worksheet unprotected. Editable cells restored.");
}

/* ============================================================
   Helpers
   ============================================================ */

/**
 * Build the Excel worksheet protection options object.
 * @param {string|null} password
 * @returns {Excel.WorksheetProtectionOptions}
 */
function buildProtectionOptions(password) {
  const options = {
    allowFormatCells:   false,
    allowFormatColumns: false,
    allowFormatRows:    false,
    allowInsertColumns: false,
    allowInsertRows:    false,
    allowInsertHyperlinks: false,
    allowDeleteColumns: false,
    allowDeleteRows:    false,
    allowSort:          false,
    allowAutoFilter:    false,
    allowPivotTables:   false,
  };
  if (password) options.password = password;
  return options;
}

/**
 * Show / hide an element with an animation-friendly approach.
 * @param {HTMLElement} el
 * @param {boolean} visible
 */
function toggleVisibility(el, visible) {
  if (visible) {
    el.classList.remove("hidden");
  } else {
    el.classList.add("hidden");
  }
}

function showUnprotectPasswordSection() {
  toggleVisibility(els.unprotectPasswordSection, true);
}

function hideUnprotectPasswordSection() {
  toggleVisibility(els.unprotectPasswordSection, false);
  els.unprotectPasswordInput.value = "";
}

/**
 * Display a status banner.
 * @param {"success"|"error"|"warning"} type
 * @param {string} icon
 * @param {string} message
 */
function showStatus(type, icon, message) {
  els.statusBanner.className = `status-banner ${type}`;
  els.statusIcon.textContent  = icon;
  els.statusText.textContent  = message;
  els.statusBanner.classList.remove("hidden");

  // Auto-dismiss success messages after 5 seconds
  if (type === "success") {
    setTimeout(clearStatus, 5000);
  }
}

function clearStatus() {
  els.statusBanner.classList.add("hidden");
  els.statusBanner.className = "status-banner hidden";
}

/**
 * Put a button in a loading/disabled state.
 * @param {HTMLButtonElement} btn
 * @param {boolean} loading
 */
function setButtonLoading(btn, loading) {
  btn.disabled = loading;
  btn.classList.toggle("loading", loading);
}

/**
 * Friendly error handling for Office.js errors.
 * @param {Error} err
 * @param {"protect"|"detect"|"unprotect"} operation
 */
function handleOfficejsError(err, operation) {
  console.error(`Cell Protector [${operation}] error:`, err);

  let message = "An unexpected error occurred. Please try again.";

  if (err && err.code) {
    switch (err.code) {
      case "InvalidOperation":
        message = "This operation isn't supported on the current selection.";
        break;
      case "AccessDenied":
        message = "Access denied. Make sure the workbook isn't open in read-only mode.";
        break;
      case "ItemNotFound":
        if (operation === "protect") {
          message = "Could not find the selected range. Please re-select and try again.";
        } else {
          message = "No active worksheet found. Please open a workbook first.";
        }
        break;
      case "GeneralException":
        message = "Excel encountered an unexpected error. Try again or restart Excel.";
        break;
      default:
        message = err.message || message;
    }
  } else if (err && err.message) {
    message = err.message;
  }

  showStatus("error", "✕", message);
}

/* ============================================================
   Future-friendly extension points
   ============================================================
   The following stubs are intentionally left for future versions:

   - protectColumns(columnIds, password, doHighlight)
   - protectRows(rowIndices, password, doHighlight)
   - protectNamedRanges(names, password, doHighlight)
   - protectMultipleSheets(sheetNames, password, doHighlight)
   - applyProtectionTemplate(templateConfig)

   Each would follow the same pattern:
   1. Unlock all cells (sheet.getUsedRange().format.protection.locked = false)
   2. Lock the target range/cells
   3. Optionally highlight editable cells (storeAndHighlightEditable)
   4. sheet.protection.protect(buildProtectionOptions(password))
   ============================================================ */

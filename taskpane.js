/* ============================================================
   Cell Protector — Taskpane Logic  v1.2.0
   Modes: selected ranges (multi-capture) | all formula cells
   ============================================================ */

"use strict";

/* ─── State ─────────────────────────────────────────────── */
let capturedRanges = []; // used in "range" mode
let currentMode    = "range"; // "range" | "formulas"

/* ─── DOM refs ──────────────────────────────────────────── */
let els = {};

/* ============================================================
   Init
   ============================================================ */
Office.onReady(() => {
  cacheElements();
  bindEvents();
  syncModeUI(); // set initial state
});

function cacheElements() {
  els = {
    // Mode tabs
    tabRange:             document.getElementById("tabRange"),
    tabFormulas:          document.getElementById("tabFormulas"),
    radioRange:           document.querySelector('input[name="mode"][value="range"]'),
    radioFormulas:        document.querySelector('input[name="mode"][value="formulas"]'),
    panelRange:           document.getElementById("panelRange"),
    panelFormulas:        document.getElementById("panelFormulas"),

    // Range panel
    addRangeBtn:          document.getElementById("addRangeBtn"),
    rangeListWrap:        document.getElementById("rangeListWrap"),
    rangeList:            document.getElementById("rangeList"),
    emptyHint:            document.getElementById("emptyHint"),
    clearAllBtn:          document.getElementById("clearAllBtn"),

    // Password
    passwordToggle:       document.getElementById("passwordToggle"),
    passwordSection:      document.getElementById("passwordSection"),
    passwordInput:        document.getElementById("passwordInput"),
    confirmPasswordInput: document.getElementById("confirmPasswordInput"),

    // Lock
    lockBtn:              document.getElementById("lockBtn"),

    // Unprotect
    detectProtectBtn:         document.getElementById("detectProtectBtn"),
    unprotectPasswordSection: document.getElementById("unprotectPasswordSection"),
    unprotectPasswordInput:   document.getElementById("unprotectPasswordInput"),
    unprotectConfirmBtn:      document.getElementById("unprotectConfirmBtn"),

    // Status
    statusBanner: document.getElementById("statusBanner"),
    statusIcon:   document.getElementById("statusIcon"),
    statusText:   document.getElementById("statusText"),
  };
}

function bindEvents() {
  // Mode tabs
  els.tabRange.addEventListener("click",    () => setMode("range"));
  els.tabFormulas.addEventListener("click", () => setMode("formulas"));

  // Range panel
  els.addRangeBtn.addEventListener("click", handleAddRange);
  els.clearAllBtn.addEventListener("click", clearAllRanges);

  // Password toggle
  els.passwordToggle.addEventListener("change", () => {
    toggleEl(els.passwordSection, els.passwordToggle.checked);
    if (!els.passwordToggle.checked) {
      els.passwordInput.value        = "";
      els.confirmPasswordInput.value = "";
    }
  });

  // Lock
  els.lockBtn.addEventListener("click", handleLock);

  // Unprotect
  els.detectProtectBtn.addEventListener("click",   handleDetect);
  els.unprotectConfirmBtn.addEventListener("click", handleUnprotectWithPassword);
  els.unprotectPasswordInput.addEventListener("keydown", e => {
    if (e.key === "Enter") handleUnprotectWithPassword();
  });
}

/* ============================================================
   Mode switching
   ============================================================ */

function setMode(mode) {
  currentMode = mode;
  clearStatus();
  syncModeUI();
}

function syncModeUI() {
  const isRange    = currentMode === "range";
  const isFormulas = currentMode === "formulas";

  // Tab active states
  els.tabRange.classList.toggle("active", isRange);
  els.tabFormulas.classList.toggle("active", isFormulas);

  // Panels
  toggleEl(els.panelRange,    isRange);
  toggleEl(els.panelFormulas, isFormulas);

  // Lock button: always enabled in formula mode; in range mode only if list has items
  els.lockBtn.disabled = isRange ? capturedRanges.length === 0 : false;

  // Button label
  els.lockBtn.querySelector("span, svg + *") ; // keep icon
  // Update text node (last child of button)
  const textNode = [...els.lockBtn.childNodes].find(n => n.nodeType === Node.TEXT_NODE);
  if (textNode) textNode.textContent = isFormulas ? " Lock Formula Cells" : " Lock Ranges";
}

/* ============================================================
   Range Capture (range mode only)
   ============================================================ */

async function handleAddRange() {
  clearStatus();
  setButtonLoading(els.addRangeBtn, true);

  try {
    await Excel.run(async context => {
      const selection = context.workbook.getSelectedRange();
      selection.load(["address"]);
      await context.sync();

      if (!selection.address) {
        showStatus("error", "⚠", "Please select a range in Excel first.");
        return;
      }

      const addr = selection.address.toUpperCase();

      if (capturedRanges.includes(addr)) {
        showStatus("warning", "⚠", `${addr} is already in the list.`);
        return;
      }

      capturedRanges.push(addr);
      renderRangeList();
      showStatus("success", "✓", `Added: ${addr}`);
    });
  } catch (err) {
    handleError(err, "add range");
  } finally {
    setButtonLoading(els.addRangeBtn, false);
  }
}

function removeRange(index) {
  capturedRanges.splice(index, 1);
  renderRangeList();
  clearStatus();
}

function clearAllRanges() {
  capturedRanges = [];
  renderRangeList();
  clearStatus();
}

function renderRangeList() {
  const has = capturedRanges.length > 0;
  toggleEl(els.rangeListWrap, has);
  toggleEl(els.emptyHint,     !has);
  if (currentMode === "range") els.lockBtn.disabled = !has;

  els.rangeList.innerHTML = "";
  capturedRanges.forEach((addr, i) => {
    const li = document.createElement("li");
    li.className = "range-item";
    li.innerHTML = `
      <span class="range-badge">
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
          <rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>
        </svg>
        ${addr}
      </span>
      <button class="range-remove" aria-label="Remove ${addr}">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
        </svg>
      </button>`;
    li.querySelector(".range-remove").addEventListener("click", () => removeRange(i));
    els.rangeList.appendChild(li);
  });
}

/* ============================================================
   Lock Flow
   ============================================================ */

async function handleLock() {
  clearStatus();

  // Validate password
  const usePassword = els.passwordToggle.checked;
  if (usePassword) {
    const pw  = els.passwordInput.value;
    const cpw = els.confirmPasswordInput.value;
    if (!pw)        return showStatus("error", "⚠", "Please enter a password.");
    if (pw !== cpw) return showStatus("error", "✕", "Passwords do not match.");
  }

  const password = usePassword ? els.passwordInput.value : null;
  setButtonLoading(els.lockBtn, true);

  try {
    if (currentMode === "range") {
      await lockRanges(password);
    } else {
      await lockFormulas(password);
    }
  } catch (err) {
    handleError(err, "lock");
  } finally {
    setButtonLoading(els.lockBtn, false);
    // Re-enable lock btn for formulas mode (disabled during loading)
    if (currentMode === "formulas") els.lockBtn.disabled = false;
  }
}

/**
 * Lock every captured range; unlock everything else.
 */
async function lockRanges(password) {
  if (capturedRanges.length === 0) {
    showStatus("error", "⚠", "Add at least one range before locking.");
    return;
  }

  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    await context.sync();

    if (sheet.protection.protected) {
      showStatus("warning", "⚠", "This sheet is already protected. Remove protection first.");
      return;
    }

    // Unlock everything
    sheet.getUsedRange(true).format.protection.locked = false;

    // Lock each captured range
    for (const addr of capturedRanges) {
      try {
        const localAddr = addr.includes("!") ? addr.split("!").slice(1).join("!") : addr;
        sheet.getRange(localAddr).format.protection.locked = true;
      } catch (_) { /* skip invalid addresses */ }
    }

    sheet.protection.protect(buildProtectionOptions(password));
    await context.sync();

    const summary = capturedRanges.join(", ");
    showStatus("success", "✓", `Locked and protected: ${summary}`);

    capturedRanges = [];
    renderRangeList();
  });
}

/**
 * Auto-detect all formula cells; lock them; unlock everything else.
 */
async function lockFormulas(password) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    await context.sync();

    if (sheet.protection.protected) {
      showStatus("warning", "⚠", "This sheet is already protected. Remove protection first.");
      return;
    }

    // Unlock everything
    const usedRange = sheet.getUsedRange(true);
    usedRange.format.protection.locked = false;
    await context.sync();

    // Find formula cells
    let formulaRange;
    try {
      formulaRange = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
      formulaRange.load("cellCount");
      await context.sync();
    } catch (_) {
      showStatus("warning", "⚠", "No formulas were found on this worksheet.");
      return;
    }

    if (!formulaRange || formulaRange.cellCount === 0) {
      showStatus("warning", "⚠", "No formulas were found on this worksheet.");
      return;
    }

    formulaRange.format.protection.locked = true;
    sheet.protection.protect(buildProtectionOptions(password));
    await context.sync();

    showStatus("success", "✓", `Formula cells locked and sheet protected. (${formulaRange.cellCount} formula cell${formulaRange.cellCount !== 1 ? "s" : ""} locked)`);
  });
}

/* ============================================================
   Unprotect Flow
   ============================================================ */

async function handleDetect() {
  clearStatus();
  toggleEl(els.unprotectPasswordSection, false);
  setButtonLoading(els.detectProtectBtn, true);

  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("protection/protected");
      await context.sync();

      if (!sheet.protection.protected) {
        showStatus("warning", "⚠", "This worksheet is not protected.");
        return;
      }

      try {
        sheet.protection.unprotect();
        await context.sync();
        showStatus("success", "✓", "Worksheet unprotected successfully.");
      } catch (_) {
        toggleEl(els.unprotectPasswordSection, true);
        showStatus("warning", "🔑", "This sheet is password-protected. Enter the password below.");
      }
    });
  } catch (err) {
    handleError(err, "detect");
  } finally {
    setButtonLoading(els.detectProtectBtn, false);
  }
}

async function handleUnprotectWithPassword() {
  const password = els.unprotectPasswordInput.value;
  if (!password) return showStatus("error", "⚠", "Please enter the sheet password.");

  clearStatus();
  setButtonLoading(els.unprotectConfirmBtn, true);

  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      try {
        sheet.protection.unprotect(password);
        await context.sync();
      } catch (_) {
        showStatus("error", "✕", "Incorrect password. Please try again.");
        return;
      }

      toggleEl(els.unprotectPasswordSection, false);
      els.unprotectPasswordInput.value = "";
      showStatus("success", "✓", "Worksheet unprotected successfully.");
    });
  } catch (err) {
    handleError(err, "unprotect");
  } finally {
    setButtonLoading(els.unprotectConfirmBtn, false);
  }
}

/* ============================================================
   Helpers
   ============================================================ */

function buildProtectionOptions(password) {
  const opts = {
    allowFormatCells: false, allowFormatColumns: false, allowFormatRows: false,
    allowInsertColumns: false, allowInsertRows: false, allowInsertHyperlinks: false,
    allowDeleteColumns: false, allowDeleteRows: false,
    allowSort: false, allowAutoFilter: false, allowPivotTables: false,
  };
  if (password) opts.password = password;
  return opts;
}

function toggleEl(el, visible) { el.classList.toggle("hidden", !visible); }

function showStatus(type, icon, message) {
  els.statusBanner.className = `status-banner ${type}`;
  els.statusIcon.textContent  = icon;
  els.statusText.textContent  = message;
  els.statusBanner.classList.remove("hidden");
  if (type === "success") setTimeout(clearStatus, 6000);
}

function clearStatus() { els.statusBanner.classList.add("hidden"); }

function setButtonLoading(btn, loading) {
  btn.disabled = loading;
  btn.classList.toggle("loading", loading);
}

function handleError(err, operation) {
  console.error(`Cell Protector [${operation}]:`, err);
  let msg = "An unexpected error occurred. Please try again.";
  if (err?.code) {
    switch (err.code) {
      case "InvalidOperation": msg = "This operation isn't supported on the current selection."; break;
      case "AccessDenied":     msg = "Access denied. Make sure the workbook is not read-only."; break;
      case "ItemNotFound":     msg = "Could not find the range. Please re-select and try again."; break;
      default:                 msg = err.message || msg;
    }
  } else if (err?.message) { msg = err.message; }
  showStatus("error", "✕", msg);
}

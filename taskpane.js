/* ============================================================
   Cell Protector — Taskpane Logic  v1.1.0
   Multi-range capture → lock on demand
   ============================================================ */

"use strict";

/* ─── State ─────────────────────────────────────────────── */
// Array of address strings the user has captured, e.g. ["Sheet1!A1:B5", "Sheet1!D3"]
let capturedRanges = [];

/* ─── DOM refs ──────────────────────────────────────────── */
let els = {};

/* ============================================================
   Init
   ============================================================ */
Office.onReady(() => {
  cacheElements();
  bindEvents();
});

function cacheElements() {
  els = {
    // Protect card
    addRangeBtn:          document.getElementById("addRangeBtn"),
    rangeListWrap:        document.getElementById("rangeListWrap"),
    rangeList:            document.getElementById("rangeList"),
    emptyHint:            document.getElementById("emptyHint"),
    clearAllBtn:          document.getElementById("clearAllBtn"),
    passwordToggle:       document.getElementById("passwordToggle"),
    passwordSection:      document.getElementById("passwordSection"),
    passwordInput:        document.getElementById("passwordInput"),
    confirmPasswordInput: document.getElementById("confirmPasswordInput"),
    lockBtn:              document.getElementById("lockBtn"),

    // Unprotect card
    detectProtectBtn:          document.getElementById("detectProtectBtn"),
    unprotectPasswordSection:  document.getElementById("unprotectPasswordSection"),
    unprotectPasswordInput:    document.getElementById("unprotectPasswordInput"),
    unprotectConfirmBtn:       document.getElementById("unprotectConfirmBtn"),

    // Status
    statusBanner: document.getElementById("statusBanner"),
    statusIcon:   document.getElementById("statusIcon"),
    statusText:   document.getElementById("statusText"),
  };
}

function bindEvents() {
  els.addRangeBtn.addEventListener("click", handleAddRange);
  els.clearAllBtn.addEventListener("click", clearAllRanges);
  els.lockBtn.addEventListener("click", handleLock);

  els.passwordToggle.addEventListener("change", () => {
    toggleEl(els.passwordSection, els.passwordToggle.checked);
    if (!els.passwordToggle.checked) {
      els.passwordInput.value = "";
      els.confirmPasswordInput.value = "";
    }
  });

  els.detectProtectBtn.addEventListener("click", handleDetect);
  els.unprotectConfirmBtn.addEventListener("click", handleUnprotectWithPassword);
  els.unprotectPasswordInput.addEventListener("keydown", e => {
    if (e.key === "Enter") handleUnprotectWithPassword();
  });
}

/* ============================================================
   Range Capture
   ============================================================ */

/**
 * Read the current selection from Excel and add its address
 * to the captured list (deduplicated).
 */
async function handleAddRange() {
  clearStatus();
  setButtonLoading(els.addRangeBtn, true);

  try {
    await Excel.run(async context => {
      const selection = context.workbook.getSelectedRange();
      selection.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      if (!selection.address) {
        showStatus("error", "⚠", "Please select a range in Excel first.");
        return;
      }

      // Normalise address (strip leading sheet name duplicate, upper-case)
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

/**
 * Remove a single range from the list by index.
 */
function removeRange(index) {
  capturedRanges.splice(index, 1);
  renderRangeList();
  clearStatus();
}

/**
 * Remove all captured ranges.
 */
function clearAllRanges() {
  capturedRanges = [];
  renderRangeList();
  clearStatus();
}

/**
 * Re-render the captured ranges list and toggle the empty hint.
 */
function renderRangeList() {
  const hasList = capturedRanges.length > 0;

  toggleEl(els.rangeListWrap, hasList);
  toggleEl(els.emptyHint,     !hasList);
  els.lockBtn.disabled = !hasList;

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
      <button class="range-remove" aria-label="Remove ${addr}" title="Remove">
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

/**
 * Validate password inputs then lock all captured ranges on the
 * active worksheet.
 */
async function handleLock() {
  clearStatus();

  if (capturedRanges.length === 0) {
    showStatus("error", "⚠", "Add at least one range before locking.");
    return;
  }

  const usePassword = els.passwordToggle.checked;
  if (usePassword) {
    const pw  = els.passwordInput.value;
    const cpw = els.confirmPasswordInput.value;
    if (!pw)       return showStatus("error", "⚠", "Please enter a password.");
    if (pw !== cpw) return showStatus("error", "✕", "Passwords do not match.");
  }

  const password = usePassword ? els.passwordInput.value : null;
  setButtonLoading(els.lockBtn, true);

  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        showStatus("warning", "⚠", "This sheet is already protected. Remove protection first.");
        return;
      }

      // 1. Unlock every cell in the used range
      const usedRange = sheet.getUsedRange(true);
      usedRange.format.protection.locked = false;

      // 2. Lock only the captured ranges
      for (const addr of capturedRanges) {
        try {
          // Strip the sheet-name prefix if present (e.g. "Sheet1!A1:B5" → "A1:B5")
          const localAddr = addr.includes("!") ? addr.split("!").slice(1).join("!") : addr;
          const range = sheet.getRange(localAddr);
          range.format.protection.locked = true;
        } catch (_) {
          // Invalid address for this sheet — skip silently
        }
      }

      // 3. Protect the sheet
      sheet.protection.protect(buildProtectionOptions(password));

      await context.sync();

      // Build friendly summary
      const rangesSummary = capturedRanges.join(", ");
      showStatus(
        "success",
        "✓",
        `Locked and protected: ${rangesSummary}`
      );

      // Reset captured list after successful lock
      capturedRanges = [];
      renderRangeList();
    });
  } catch (err) {
    handleError(err, "lock");
  } finally {
    setButtonLoading(els.lockBtn, false);
  }
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

      // Try without a password first
      try {
        sheet.protection.unprotect();
        await context.sync();
        showStatus("success", "✓", "Worksheet unprotected successfully.");
      } catch (_) {
        // Password required
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
      sheet.load("protection/protected");
      await context.sync();

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
    allowFormatCells:      false,
    allowFormatColumns:    false,
    allowFormatRows:       false,
    allowInsertColumns:    false,
    allowInsertRows:       false,
    allowInsertHyperlinks: false,
    allowDeleteColumns:    false,
    allowDeleteRows:       false,
    allowSort:             false,
    allowAutoFilter:       false,
    allowPivotTables:      false,
  };
  if (password) opts.password = password;
  return opts;
}

function toggleEl(el, visible) {
  el.classList.toggle("hidden", !visible);
}

function showStatus(type, icon, message) {
  els.statusBanner.className = `status-banner ${type}`;
  els.statusIcon.textContent  = icon;
  els.statusText.textContent  = message;
  els.statusBanner.classList.remove("hidden");
  if (type === "success") setTimeout(clearStatus, 6000);
}

function clearStatus() {
  els.statusBanner.classList.add("hidden");
}

function setButtonLoading(btn, loading) {
  btn.disabled = loading || (btn === els.lockBtn && capturedRanges.length === 0);
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
  } else if (err?.message) {
    msg = err.message;
  }
  showStatus("error", "✕", msg);
}

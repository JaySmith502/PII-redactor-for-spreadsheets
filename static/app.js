(() => {
  "use strict";

  const form = document.getElementById("workbook-form");
  if (!form) return;

  const input = document.getElementById("workbook-input");
  const dropZone = document.getElementById("drop-zone");
  const summary = document.getElementById("file-summary");
  const fileName = document.getElementById("file-name");
  const fileSize = document.getElementById("file-size");
  const fileTypeLabel = document.getElementById("file-type-label");
  const fileError = document.getElementById("file-error");
  const removeButton = document.getElementById("remove-file");
  const submitButton = document.getElementById("submit-button");
  const replaceSection = document.getElementById("replace-section");
  const replaceRules = document.getElementById("replace-rules");
  const addReplaceRule = document.getElementById("add-replace-rule");
  const replaceError = document.getElementById("replace-error");

  const columnSection = document.getElementById("column-section");
  const columnLoading = document.getElementById("column-loading");
  const columnError = document.getElementById("column-error");
  const columnPanel = document.getElementById("column-panel");
  const columnCount = document.getElementById("column-count");
  const selectAllCols = document.getElementById("select-all-cols");
  const clearCols = document.getElementById("clear-cols");
  const colList = document.getElementById("col-list");

  const peekTokenInput = document.getElementById("peek-token-input");
  const selectedColsInput = document.getElementById("selected-columns-input");
  const replacementRulesInput = document.getElementById("replacement-rules-input");

  const maxBytes = 25 * 1024 * 1024;
  const excelSuffix = ".xlsx";
  const replaceSuffixes = [".xlsx", ".docx", ".pdf"];

  /* Tracks which File the current peek result belongs to, so re-rendering the
     form does not upload the same file over and over. */
  let peekedFile = null;
  let peekRequestId = 0;

  const formatBytes = (bytes) => {
    if (bytes < 1024 * 1024) return Math.max(1, Math.round(bytes / 1024)) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  };

  const getMode = () => {
    const checked = form.querySelector('input[name="mode"]:checked');
    return checked ? checked.value : "encrypt";
  };

  const fileSuffix = (file) => {
    const name = file ? file.name.toLowerCase() : "";
    const index = name.lastIndexOf(".");
    return index >= 0 ? name.slice(index) : "";
  };

  const validSuffixForMode = (file, mode) => {
    const suffix = fileSuffix(file);
    if (mode === "encrypt" || mode === "decrypt") return suffix === excelSuffix;
    if (mode === "replace") return replaceSuffixes.includes(suffix);
    return false;
  };

  const setMode = (mode) => {
    const radio = form.querySelector('input[name="mode"][value="' + mode + '"]');
    if (radio && !radio.disabled) radio.checked = true;
  };

  const maybeSwitchToReplaceMode = (file) => {
    const suffix = fileSuffix(file);
    if ((suffix === ".pdf" || suffix === ".docx") && getMode() !== "replace") {
      setMode("replace");
    }
  };

  /* Steps are numbered by what is actually on screen: the field picker only
     exists for encrypt, and find-and-replace only for encrypt and replace. */
  const renumberSteps = () => {
    let step = 0;
    Array.from(document.querySelectorAll(".step-number")).forEach((badge) => {
      if (badge.closest("[hidden]")) return;
      step += 1;
      badge.textContent = String(step);
    });
  };

  const getCheckedKeys = () =>
    Array.from(colList.querySelectorAll("input[type=checkbox]:checked")).map(
      (checkbox) => checkbox.value
    );

  const hasValidFile = () => {
    const file = input.files && input.files[0];
    return Boolean(file) && validSuffixForMode(file, getMode()) && file.size <= maxBytes;
  };

  const refreshSubmit = () => {
    if (!hasValidFile()) {
      submitButton.disabled = true;
      return;
    }
    if (getMode() === "encrypt") {
      submitButton.disabled = !peekTokenInput.value || getCheckedKeys().length === 0;
      return;
    }
    submitButton.disabled = false;
  };

  const updateColumnCount = () => {
    const total = colList.querySelectorAll("input[type=checkbox]").length;
    const selected = getCheckedKeys().length;
    if (!total) {
      columnCount.textContent = "No fields available";
    } else if (selected === 0) {
      columnCount.textContent = "No fields selected";
    } else {
      columnCount.textContent = selected + " of " + total + " fields selected";
    }
    columnCount.classList.toggle("is-active", selected > 0);
    refreshSubmit();
  };

  const showFileError = (message) => {
    fileError.textContent = message;
    fileError.hidden = false;
    summary.hidden = true;
    submitButton.disabled = true;
  };

  const resetColumns = () => {
    peekTokenInput.value = "";
    selectedColsInput.value = "";
    peekedFile = null;
    peekRequestId += 1;
    colList.innerHTML = "";
    columnPanel.hidden = true;
    columnLoading.hidden = true;
    columnError.hidden = true;
  };

  const hideColumnSection = () => {
    columnSection.hidden = true;
    renumberSteps();
  };

  const showColumnError = (message) => {
    columnLoading.hidden = true;
    columnPanel.hidden = true;
    columnError.textContent = message;
    columnError.hidden = false;
  };

  const renderColumns = (columns) => {
    const sheets = Object.keys(columns);
    colList.innerHTML = "";

    if (sheets.length === 0) {
      showColumnError("No fields were found in this file.");
      return;
    }

    sheets.forEach((sheetName) => {
      const group = document.createElement("div");
      group.className = "col-sheet-group";

      const label = document.createElement("p");
      label.className = "col-sheet-label";
      label.textContent = "Sheet: " + sheetName;
      group.appendChild(label);

      const items = document.createElement("div");
      items.className = "col-items";

      columns[sheetName].forEach((colInfo) => {
        const cardLabel = document.createElement("label");
        cardLabel.className = "col-card";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = colInfo.key;
        /* Deliberately opt-in: nothing is encrypted unless the user picks it. */
        checkbox.checked = false;
        checkbox.addEventListener("change", updateColumnCount);

        const inner = document.createElement("div");
        inner.className = "col-card-inner";

        const checkmark = document.createElement("span");
        checkmark.className = "col-card-check";
        checkmark.setAttribute("aria-hidden", "true");
        checkmark.textContent = "OK";

        const text = document.createElement("div");
        text.className = "col-card-text";

        const strong = document.createElement("strong");
        strong.textContent = colInfo.header || "Column " + colInfo.col;

        const small = document.createElement("small");
        small.textContent = "Column " + colInfo.col;

        text.append(strong, small);
        inner.append(checkmark, text);
        cardLabel.append(checkbox, inner);
        items.appendChild(cardLabel);
      });

      group.appendChild(items);
      colList.appendChild(group);
    });

    columnLoading.hidden = true;
    columnError.hidden = true;
    columnPanel.hidden = false;
    updateColumnCount();
  };

  /* Reads the file's structure as soon as it is chosen, so the field picker
     appears inline on the main screen instead of after Process is clicked. */
  const loadColumns = async (file) => {
    const requestId = ++peekRequestId;
    peekTokenInput.value = "";
    colList.innerHTML = "";
    columnPanel.hidden = true;
    columnError.hidden = true;
    columnLoading.hidden = false;
    refreshSubmit();

    const csrfToken = form.querySelector('input[name="csrf_token"]').value;
    const formData = new FormData();
    formData.append("workbook", file);
    formData.append("csrf_token", csrfToken);

    try {
      const response = await fetch("/peek", { method: "POST", body: formData });
      const data = await response.json();
      if (requestId !== peekRequestId) return;

      if (!response.ok || data.error) {
        showColumnError(data.error || "Failed to read the file structure.");
        refreshSubmit();
        return;
      }

      peekTokenInput.value = data.peek_token;
      peekedFile = file;
      renderColumns(data.columns);
    } catch (error) {
      if (requestId !== peekRequestId) return;
      showColumnError("Could not read the file. Remove it and try again.");
      refreshSubmit();
    }
  };

  const syncColumnSection = () => {
    const file = input.files && input.files[0];
    const wantsColumns = getMode() === "encrypt" && Boolean(file) && hasValidFile();

    if (!wantsColumns) {
      resetColumns();
      hideColumnSection();
      return;
    }

    columnSection.hidden = false;
    renumberSteps();
    if (peekedFile === file && peekTokenInput.value) return;
    loadColumns(file);
  };

  const updateFile = () => {
    fileError.hidden = true;

    const file = input.files && input.files[0];
    if (!file) {
      summary.hidden = true;
      resetColumns();
      hideColumnSection();
      refreshSubmit();
      return;
    }
    maybeSwitchToReplaceMode(file);
    const mode = getMode();
    if (!validSuffixForMode(file, mode)) {
      resetColumns();
      hideColumnSection();
      showFileError(
        mode === "replace"
          ? "Choose an .xlsx, .docx, or .pdf file."
          : "Choose a file with the .xlsx extension."
      );
      return;
    }
    if (file.size > maxBytes) {
      resetColumns();
      hideColumnSection();
      showFileError("This file is larger than the 25 MB limit.");
      return;
    }
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileTypeLabel.textContent = fileSuffix(file).replace(".", "").toUpperCase() || "FILE";
    summary.hidden = false;
    syncColumnSection();
    refreshSubmit();
  };

  const newReplacementRow = () => {
    const row = document.createElement("div");
    row.className = "replace-row";

    const find = document.createElement("input");
    find.type = "text";
    find.className = "replace-input";
    find.placeholder = "Sensitive value";
    find.autocomplete = "off";
    find.maxLength = 500;

    const replace = document.createElement("input");
    replace.type = "text";
    replace.className = "replace-input";
    replace.placeholder = "Safe label";
    replace.autocomplete = "off";
    replace.maxLength = 500;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "text-button remove-rule-button";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      row.remove();
      replaceError.hidden = true;
      if (!replaceRules.children.length) replaceRules.appendChild(newReplacementRow());
    });

    row.append(find, replace, remove);
    return row;
  };

  const collectReplacementRules = () => {
    const rules = [];
    replaceError.hidden = true;

    Array.from(replaceRules.querySelectorAll(".replace-row")).forEach((row) => {
      const inputs = row.querySelectorAll(".replace-input");
      const find = inputs[0].value.trim();
      const replace = inputs[1].value.trim();
      if (!find && !replace) return;
      if (!find || !replace) {
        throw new Error("Each replacement row needs both a sensitive value and a safe label.");
      }
      if (find === replace) return;
      rules.push({ find: find, replace: replace });
    });

    if (getMode() === "replace" && rules.length === 0) {
      throw new Error("Add at least one replacement before using Replace only.");
    }
    replacementRulesInput.value = JSON.stringify(rules);
  };

  const updateMode = () => {
    const mode = getMode();
    const canReplace = mode === "encrypt" || mode === "replace";
    replaceSection.hidden = !canReplace;
    if (!canReplace) {
      replacementRulesInput.value = "";
      replaceError.hidden = true;
    }
    updateFile();
    renumberSteps();
  };

  addReplaceRule.addEventListener("click", () => {
    replaceRules.appendChild(newReplacementRow());
    const lastRow = replaceRules.lastElementChild;
    if (lastRow) lastRow.querySelector("input").focus();
  });
  replaceRules.appendChild(newReplacementRow());

  selectAllCols.addEventListener("click", () => {
    colList.querySelectorAll("input[type=checkbox]").forEach((checkbox) => {
      checkbox.checked = true;
    });
    updateColumnCount();
  });

  clearCols.addEventListener("click", () => {
    colList.querySelectorAll("input[type=checkbox]").forEach((checkbox) => {
      checkbox.checked = false;
    });
    updateColumnCount();
  });

  const setWorkingState = () => {
    submitButton.disabled = true;
    submitButton.querySelector(".button-label").hidden = true;
    submitButton.querySelector(".button-working").hidden = false;
    form.setAttribute("aria-busy", "true");
  };

  form.addEventListener("submit", (event) => {
    if (!hasValidFile()) {
      event.preventDefault();
      updateFile();
      return;
    }

    const mode = getMode();
    if (mode === "encrypt" || mode === "replace") {
      try {
        collectReplacementRules();
      } catch (error) {
        event.preventDefault();
        replaceError.textContent = error.message;
        replaceError.hidden = false;
        return;
      }
    }

    if (mode === "encrypt") {
      const keys = getCheckedKeys();
      if (!peekTokenInput.value || keys.length === 0) {
        event.preventDefault();
        showColumnError("Select at least one field to encrypt.");
        return;
      }
      selectedColsInput.value = JSON.stringify(keys);
      /* The file already sits on the server under the peek token, so skip
         sending the bytes a second time. */
      input.disabled = true;
    }

    setWorkingState();
  });

  input.addEventListener("change", updateFile);

  form.querySelectorAll('input[name="mode"]').forEach((radio) => {
    radio.addEventListener("change", updateMode);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add("is-dragging");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.remove("is-dragging");
    });
  });

  dropZone.addEventListener("drop", (event) => {
    if (event.dataTransfer.files.length) {
      input.files = event.dataTransfer.files;
      input.disabled = false;
      updateFile();
    }
  });

  removeButton.addEventListener("click", () => {
    input.value = "";
    input.disabled = false;
    replacementRulesInput.value = "";
    updateFile();
    input.focus();
  });

  updateMode();
})();

/* Floating "Copy secret key" button.
   Separate from the IIFE above, which returns early on pages without the
   upload form. The key is fetched only after the user confirms the warning
   dialog, and it is placed on the clipboard only - never rendered into the
   page. */
(() => {
  "use strict";

  const button = document.getElementById("copy-key-button");
  const statusEl = document.getElementById("copy-key-status");
  const backdrop = document.getElementById("key-modal-backdrop");
  const confirmCheckbox = document.getElementById("key-confirm-checkbox");
  const confirmButton = document.getElementById("key-modal-confirm");
  const cancelButton = document.getElementById("key-modal-cancel");
  const closeButton = document.getElementById("key-modal-close");
  if (!button || !statusEl || !backdrop) return;

  let hideTimer = 0;
  let lastFocused = null;

  const showStatus = (message, kind) => {
    statusEl.textContent = message;
    statusEl.classList.remove("is-ok", "is-error");
    if (kind) statusEl.classList.add(kind);
    statusEl.hidden = false;
    window.clearTimeout(hideTimer);
    hideTimer = window.setTimeout(() => {
      statusEl.hidden = true;
    }, 8000);
  };

  const openModal = () => {
    lastFocused = document.activeElement;
    /* Never remembered between copies: the warning is confirmed every time. */
    confirmCheckbox.checked = false;
    confirmButton.disabled = true;
    confirmButton.textContent = "Copy key to clipboard";
    backdrop.hidden = false;
    backdrop.removeAttribute("aria-hidden");
    document.body.style.overflow = "hidden";
    closeButton.focus();
  };

  const closeModal = () => {
    backdrop.hidden = true;
    backdrop.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
    if (lastFocused && typeof lastFocused.focus === "function") lastFocused.focus();
  };

  const legacyCopy = (value) => {
    const holder = document.createElement("textarea");
    holder.value = value;
    holder.setAttribute("readonly", "");
    holder.style.position = "fixed";
    holder.style.top = "-1000px";
    holder.style.opacity = "0";
    document.body.appendChild(holder);
    holder.select();
    let copied = false;
    try {
      copied = document.execCommand("copy");
    } catch (error) {
      copied = false;
    }
    holder.value = "";
    document.body.removeChild(holder);
    return copied;
  };

  const copyToClipboard = async (value) => {
    if (navigator.clipboard && window.isSecureContext) {
      try {
        await navigator.clipboard.writeText(value);
        return true;
      } catch (error) {
        /* fall through to the textarea path */
      }
    }
    return legacyCopy(value);
  };

  button.addEventListener("click", openModal);
  cancelButton.addEventListener("click", closeModal);
  closeButton.addEventListener("click", closeModal);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) closeModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !backdrop.hidden) closeModal();
  });
  confirmCheckbox.addEventListener("change", () => {
    confirmButton.disabled = !confirmCheckbox.checked;
  });

  confirmButton.addEventListener("click", async () => {
    if (!confirmCheckbox.checked) return;

    const meta = document.querySelector('meta[name="csrf-token"]');
    const csrfToken = meta ? meta.getAttribute("content") || "" : "";

    confirmButton.disabled = true;
    confirmButton.textContent = "Copying...";
    try {
      const body = new FormData();
      body.append("csrf_token", csrfToken);

      const response = await fetch("/secret-key", {
        method: "POST",
        body: body,
        credentials: "same-origin",
      });

      let payload = null;
      try {
        payload = await response.json();
      } catch (error) {
        payload = null;
      }

      if (!response.ok || !payload || !payload.key) {
        const message = (payload && payload.error) || "Could not read the secret key.";
        closeModal();
        showStatus(message, "is-error");
        return;
      }

      const copied = await copyToClipboard(payload.key);
      closeModal();
      showStatus(
        copied
          ? "Secret key copied. Keep it private and clear your clipboard when you are done."
          : "Clipboard blocked by the browser. Copy it from .env instead.",
        copied ? "is-ok" : "is-error"
      );
    } catch (error) {
      closeModal();
      showStatus("Network error. Please try again.", "is-error");
    } finally {
      confirmButton.textContent = "Copy key to clipboard";
      confirmButton.disabled = !confirmCheckbox.checked;
    }
  });
})();

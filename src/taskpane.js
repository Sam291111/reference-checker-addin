import {
  buildLookup,
  detectHeaderRow,
  findReferenceMatch,
  findWorkDocMatch,
  isReferenceHeader,
  normalizeKey,
  recordsFromCsv
} from "./referenceMatcher.js";

const state = {
  lookup: null,
  queue: [],
  settings: {
    alertsEnabled: true,
    soundEnabled: true
  },
  nextSuggestionId: 1,
  toastTimer: null,
  audioContext: null,
  reviewDialog: null
};

const elements = {
  connectionStatus: document.getElementById("connectionStatus"),
  lookupBadge: document.getElementById("lookupBadge"),
  lookupFile: document.getElementById("lookupFile"),
  lookupSummary: document.getElementById("lookupSummary"),
  alertsEnabled: document.getElementById("alertsEnabled"),
  soundEnabled: document.getElementById("soundEnabled"),
  testSound: document.getElementById("testSound"),
  scanSheet: document.getElementById("scanSheet"),
  openReviewWindow: document.getElementById("openReviewWindow"),
  selectAll: document.getElementById("selectAll"),
  applySelected: document.getElementById("applySelected"),
  deleteRows: document.getElementById("deleteRows"),
  removeSelected: document.getElementById("removeSelected"),
  hidePane: document.getElementById("hidePane"),
  clearQueue: document.getElementById("clearQueue"),
  toast: document.getElementById("toast"),
  toastTitle: document.getElementById("toastTitle"),
  toastMessage: document.getElementById("toastMessage"),
  queueCount: document.getElementById("queueCount"),
  queueList: document.getElementById("queueList")
};

loadSettings();
loadCachedLookup();
wireUi();
renderQueue();

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    elements.connectionStatus.textContent = "Open this add-in from Excel.";
    elements.scanSheet.disabled = true;
    elements.openReviewWindow.disabled = true;
    elements.hidePane.disabled = true;
    return;
  }

  elements.connectionStatus.textContent = "Watching workbook edits.";
  await registerWorksheetChangeHandler();
});

Office.actions?.associate?.("openReviewWindow", async (event) => {
  await openReviewWindow();
  event.completed();
});

Office.actions?.associate?.("hideTaskPane", async (event) => {
  await hideTaskPane();
  event.completed();
});

function wireUi() {
  elements.lookupFile.addEventListener("change", importLookupFile);
  elements.testSound.addEventListener("click", () => testSound());
  elements.scanSheet.addEventListener("click", () => scanActiveSheet());
  elements.openReviewWindow.addEventListener("click", () => openReviewWindow());
  elements.selectAll.addEventListener("click", selectAllSuggestions);
  elements.applySelected.addEventListener("click", () => applySelectedSuggestions());
  elements.deleteRows.addEventListener("click", () => deleteSelectedWorkDocRows());
  elements.removeSelected.addEventListener("click", removeSelectedSuggestions);
  elements.hidePane.addEventListener("click", hideTaskPane);
  elements.clearQueue.addEventListener("click", clearQueue);
  elements.queueList.addEventListener("change", handleQueueSelection);

  for (const key of Object.keys(state.settings)) {
    elements[key].checked = state.settings[key];
    elements[key].addEventListener("change", () => {
      state.settings[key] = elements[key].checked;
      saveSettings();
      primeAudio();
    });
  }
}

async function importLookupFile(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  primeAudio();
  try {
    const buffer = await file.arrayBuffer();
    const text = decodeLookupCsv(buffer);
    const records = recordsFromCsv(text);
    setLookup(buildLookup(records), file.name);
    localStorage.setItem("referenceChecker.lookup", JSON.stringify({ fileName: file.name, records }));
    notify("Lookup ready", `${records.length} lookup rows loaded.`);
  } catch (error) {
    notify("Lookup import failed", error.message);
  }
}

function decodeLookupCsv(buffer) {
  try {
    return new TextDecoder("windows-1252", { fatal: false }).decode(buffer);
  } catch {
    return new TextDecoder("utf-8", { fatal: false }).decode(buffer);
  }
}

function setLookup(lookup, fileName = "lookup CSV") {
  state.lookup = lookup;
  elements.lookupBadge.textContent = "Lookup ready";
  elements.lookupBadge.className = "badge badge-ready";
  elements.lookupSummary.textContent = `${lookup.entries.length} lookup rows loaded from ${fileName}.`;
}

function loadCachedLookup() {
  const cached = localStorage.getItem("referenceChecker.lookup");
  if (!cached) {
    return;
  }

  try {
    const { fileName, records } = JSON.parse(cached);
    setLookup(buildLookup(records), fileName || "cached lookup");
  } catch {
    localStorage.removeItem("referenceChecker.lookup");
  }
}

function loadSettings() {
  const cached = localStorage.getItem("referenceChecker.settings");
  if (!cached) {
    return;
  }

  try {
    state.settings = { ...state.settings, ...JSON.parse(cached) };
  } catch {
    localStorage.removeItem("referenceChecker.settings");
  }
}

function saveSettings() {
  localStorage.setItem("referenceChecker.settings", JSON.stringify(state.settings));
}

async function registerWorksheetChangeHandler() {
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onChanged.add(handleWorksheetChange);
      await context.sync();
    });
  } catch (error) {
    elements.connectionStatus.textContent = "Could not attach live watcher. Use Check Active Sheet.";
    notify("Watcher setup issue", error.message, { sound: false });
  }
}

async function handleWorksheetChange(event) {
  if (!state.lookup) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(event.worksheetId);
      await processChangedRange(context, worksheet, event.address);
    });
  } catch (error) {
    notify("Change check failed", error.message, { sound: false });
  }
}

async function scanActiveSheet() {
  primeAudio();
  if (!state.lookup) {
    notify("Import lookup first", "Choose the lookup CSV before checking a sheet.", { sound: false });
    return;
  }

  elements.scanSheet.disabled = true;
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRangeOrNullObject();
      worksheet.load(["id", "name"]);
      usedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
      await context.sync();

      if (usedRange.isNullObject) {
        notify("No cells to check", "The active sheet is empty.", { sound: false });
        return;
      }

      await processValues(context, worksheet, usedRange, usedRange, true);
    });
  } catch (error) {
    notify("Sheet check failed", error.message, { sound: false });
  } finally {
    elements.scanSheet.disabled = false;
  }
}

async function processChangedRange(context, worksheet, address) {
  const changedRange = worksheet.getRange(address);
  const usedRange = worksheet.getUsedRangeOrNullObject();
  worksheet.load(["id", "name"]);
  changedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
  usedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  if (usedRange.isNullObject) {
    return;
  }

  await processValues(context, worksheet, changedRange, usedRange, false);
}

async function processValues(context, worksheet, targetRange, usedRange, isFullScan) {
  const sheetValues = usedRange.values;
  const headerRow = detectHeaderRow(sheetValues);
  const headers = sheetValues[headerRow].map((value) => normalizeKey(value));
  const citationColumns = headers
    .map((header, index) => (isReferenceHeader(header) || header === "author" ? index : -1))
    .filter((index) => index >= 0);
  const workDocColumns = makeColumnMap(headers, ["doi", "title", "authors", "author", "year"]);
  const origin = {
    row: usedRange.rowIndex,
    column: usedRange.columnIndex
  };

  let addedCount = 0;
  if (citationColumns.length > 0) {
    addedCount += queueCitationSuggestions(worksheet, targetRange, sheetValues, headers, headerRow, citationColumns, origin, isFullScan);
  }

  if (workDocColumns.doi >= 0 && workDocColumns.year >= 0) {
    addedCount += queueWorkDocSuggestions(worksheet, targetRange, sheetValues, headerRow, workDocColumns, origin, isFullScan);
  }

  if (addedCount > 0) {
    renderQueue();
    notify(
      `${addedCount} suggestion${addedCount === 1 ? "" : "s"} added`,
      "Open the queue when you are ready to review."
    );
  } else if (isFullScan) {
    notify("No lookup matches found", "No new suggestions were added.", { sound: false });
  }

  await context.sync();
}

function queueCitationSuggestions(worksheet, targetRange, sheetValues, headers, headerRow, citationColumns, origin, isFullScan) {
  const startRow = isFullScan ? headerRow + 1 : targetRange.rowIndex - origin.row;
  const endRow = isFullScan ? sheetValues.length - 1 : targetRange.rowIndex - origin.row + targetRange.rowCount - 1;
  const startCol = isFullScan ? 0 : targetRange.columnIndex - origin.column;
  const endCol = isFullScan ? sheetValues[headerRow].length - 1 : targetRange.columnIndex - origin.column + targetRange.columnCount - 1;
  let addedCount = 0;

  for (let rowIndex = Math.max(startRow, headerRow + 1); rowIndex <= endRow; rowIndex += 1) {
    for (const columnIndex of citationColumns) {
      if (columnIndex < startCol || columnIndex > endCol) {
        continue;
      }

      const currentValue = sheetValues[rowIndex]?.[columnIndex];
      if (!currentValue) {
        continue;
      }

      const match = findReferenceMatch(state.lookup, currentValue);
      if (!match.matched) {
        continue;
      }

      const absoluteRow = origin.row + rowIndex;
      const absoluteColumn = origin.column + columnIndex;
      const lookupEntry = match.entries[0];
      const suggestedValue = !match.ambiguous ? match.citations[0] : "";

      if (suggestedValue && normalizeKey(suggestedValue) === normalizeKey(currentValue)) {
        continue;
      }

      addedCount += addSuggestion({
        type: "citation",
        sheetId: worksheet.id,
        sheetName: worksheet.name,
        row: absoluteRow,
        column: absoluteColumn,
        location: `${worksheet.name}!R${absoluteRow + 1}C${absoluteColumn + 1}`,
        columnHeader: headers[columnIndex],
        currentValue: String(currentValue),
        suggestedValue,
        ambiguous: match.ambiguous,
        citations: match.citations,
        lookupAuthor: lookupEntry?.citation || match.citations.join(", "),
        doi: lookupEntry?.doi || "",
        title: lookupEntry?.title || "",
        reason: match.ambiguous ? "Multiple lookup citations share this author/year." : "Citation already exists in lookup."
      });
    }
  }

  return addedCount;
}

function queueWorkDocSuggestions(worksheet, targetRange, sheetValues, headerRow, columns, origin, isFullScan) {
  const startRow = isFullScan ? headerRow + 1 : targetRange.rowIndex - origin.row;
  const endRow = isFullScan ? sheetValues.length - 1 : targetRange.rowIndex - origin.row + targetRange.rowCount - 1;
  let addedCount = 0;

  for (let rowIndex = Math.max(startRow, headerRow + 1); rowIndex <= endRow; rowIndex += 1) {
    const values = sheetValues[rowIndex] || [];
    const row = {
      doi: values[columns.doi],
      title: values[columns.title],
      authors: columns.authors >= 0 ? values[columns.authors] : values[columns.author],
      year: values[columns.year]
    };
    const match = findWorkDocMatch(state.lookup, row);
    if (!match.matched) {
      continue;
    }

    const lookupEntry = match.entries[0];
    const absoluteRow = origin.row + rowIndex;
    const added = addSuggestion({
      type: "work-doc",
      sheetId: worksheet.id,
      sheetName: worksheet.name,
      row: absoluteRow,
      column: -1,
      location: `${worksheet.name}!row ${absoluteRow + 1}`,
      columnHeader: "work doc",
      currentValue: row.doi || row.title || row.authors || "",
      suggestedValue: "",
      ambiguous: match.ambiguous,
      citations: match.citations,
      lookupAuthor: lookupEntry?.citation || match.citations.join(", "),
      doi: lookupEntry?.doi || row.doi || "",
      title: lookupEntry?.title || row.title || "",
      reason: `Paper already appears in lookup by ${describeWorkDocMatch(match.kind)}.`
    });
    addedCount += added;

    if (added) {
      highlightWorkDocRow(worksheet, absoluteRow, origin.column, sheetValues[headerRow].length);
    }
  }

  return addedCount;
}

function describeWorkDocMatch(kind) {
  if (kind === "doi") {
    return "DOI";
  }
  if (kind === "title-year") {
    return "title and year";
  }
  if (kind === "author-year") {
    return "author and year";
  }
  return kind;
}

function highlightWorkDocRow(worksheet, absoluteRow, startColumn, columnCount) {
  const rowRange = worksheet.getRangeByIndexes(absoluteRow, startColumn, 1, columnCount);
  rowRange.format.fill.color = "#fff4d9";
}

function addSuggestion(suggestion) {
  const dedupeKey = makeSuggestionKey(suggestion);
  if (state.queue.some((item) => item.dedupeKey === dedupeKey && item.status !== "removed")) {
    return 0;
  }

  state.queue.unshift({
    ...suggestion,
    id: String(state.nextSuggestionId++),
    dedupeKey,
    selected: Boolean(suggestion.suggestedValue && !suggestion.ambiguous),
    status: "queued",
    createdAt: new Date().toISOString()
  });
  return 1;
}

function makeSuggestionKey(suggestion) {
  return [
    suggestion.type,
    suggestion.sheetId,
    suggestion.row,
    suggestion.column,
    normalizeKey(suggestion.currentValue),
    normalizeKey(suggestion.suggestedValue || suggestion.citations.join("|"))
  ].join("::");
}

function renderQueue() {
  const visibleQueue = state.queue.filter((item) => item.status !== "removed");
  elements.queueCount.textContent = String(visibleQueue.length);
  elements.queueList.replaceChildren(...visibleQueue.map(renderSuggestion));
  updateQueueActions();
  syncQueueToStorage();
}

function renderSuggestion(item) {
  const wrapper = document.createElement("article");
  wrapper.className = `suggestion ${item.type === "work-doc" ? "work-doc" : ""} ${item.ambiguous ? "ambiguous" : ""} ${item.status === "applied" ? "applied" : ""}`;
  wrapper.dataset.id = item.id;

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = item.selected;
  checkbox.setAttribute("aria-label", `Select suggestion for ${item.location}`);
  wrapper.append(checkbox);

  const content = document.createElement("div");
  content.append(
    makeSuggestionTitle(item),
    makeFieldGroup("Location", `${item.location} (${item.columnHeader})`, "suggestion-meta"),
    makePaperBlock(item),
    makeChangeBlock(item)
  );
  wrapper.append(content);
  return wrapper;
}

function makeSuggestionTitle(item) {
  const row = document.createElement("div");
  row.className = "suggestion-title";
  const title = document.createElement("span");
  title.textContent = item.type === "work-doc" ? "Already in lookup - skip this paper" : "Possible citation replacement";
  const status = document.createElement("span");
  status.textContent = item.status === "applied" ? "Applied" : item.ambiguous ? "Review" : "Ready";
  row.append(title, status);
  return row;
}

function makePaperBlock(item) {
  const block = document.createElement("div");
  block.className = "suggestion-paper";
  block.append(
    makeFieldGroup("Author", item.lookupAuthor || "Unknown"),
    makeFieldGroup("DOI", item.doi || "No DOI in lookup"),
    makeFieldGroup("Title", item.title || "No title in lookup")
  );
  return block;
}

function makeChangeBlock(item) {
  const block = document.createElement("div");
  block.className = "suggestion-change";
  block.append(makeFieldGroup("Current", item.currentValue || "Blank"));
  if (item.suggestedValue) {
    block.append(makeFieldGroup("Suggested", item.suggestedValue, "", "suggested-value"));
  } else if (item.type === "work-doc") {
    block.append(makeFieldGroup("Action", "Do not add/process this listing unless it is intentionally different.", "", "suggested-value"));
  } else {
    block.append(makeFieldGroup("Suggested", item.citations.join(", ") || "No automatic change", "", "suggested-value"));
  }
  block.append(makeFieldGroup("Reason", item.reason));
  return block;
}

function makeFieldGroup(label, value, className = "", valueClassName = "") {
  const group = document.createElement("div");
  group.className = className || "";
  const labelEl = document.createElement("span");
  labelEl.className = "field-label";
  labelEl.textContent = label;
  const valueEl = document.createElement("span");
  valueEl.className = `field-value ${valueClassName}`.trim();
  valueEl.textContent = value;
  group.append(labelEl, valueEl);
  return group;
}

function handleQueueSelection(event) {
  const checkbox = event.target.closest("input[type='checkbox']");
  if (!checkbox) {
    return;
  }

  const card = checkbox.closest(".suggestion");
  const item = state.queue.find((suggestion) => suggestion.id === card?.dataset.id);
  if (item) {
    item.selected = checkbox.checked;
    updateQueueActions();
  }
}

function updateQueueActions() {
  const visibleQueue = state.queue.filter((item) => item.status !== "removed");
  const selectedReadyCount = visibleQueue.filter((item) => item.selected && canApply(item)).length;
  const selectedDeletableCount = visibleQueue.filter((item) => item.selected && canDeleteRow(item)).length;
  const selectedCount = visibleQueue.filter((item) => item.selected).length;
  elements.selectAll.disabled = visibleQueue.length === 0;
  elements.applySelected.disabled = selectedReadyCount === 0;
  elements.deleteRows.disabled = selectedDeletableCount === 0;
  elements.removeSelected.disabled = selectedCount === 0;
  elements.clearQueue.disabled = visibleQueue.length === 0;
}

function canApply(item) {
  return item.type === "citation" && item.suggestedValue && !item.ambiguous && item.status === "queued";
}

function canDeleteRow(item) {
  return item.type === "work-doc" && item.status === "queued";
}

function selectAllSuggestions() {
  const visibleQueue = state.queue.filter((item) => item.status !== "removed");
  if (visibleQueue.length === 0) {
    notify("Nothing to select", "There are no suggestions in the queue.", { sound: false });
    return;
  }

  for (const item of visibleQueue) {
    item.selected = true;
  }
  renderQueue();
  notify("Suggestions selected", `${visibleQueue.length} item${visibleQueue.length === 1 ? "" : "s"} selected.`, { sound: false });
}

async function applySelectedSuggestions() {
  primeAudio();
  const ids = state.queue.filter((item) => item.selected && canApply(item)).map((item) => item.id);
  if (ids.length === 0) {
    notify("Nothing selected", "Select one or more ready suggestions first.", { sound: false });
    return;
  }

  await applySuggestionIds(ids);
}

async function applySuggestionIds(ids) {
  const selected = state.queue.filter((item) => ids.includes(item.id) && canApply(item));
  if (selected.length === 0) {
    notify("Nothing ready to apply", "Only unambiguous citation suggestions can be applied.", { sound: false });
    return;
  }

  elements.applySelected.disabled = true;
  try {
    await Excel.run(async (context) => {
      for (const item of selected) {
        const worksheet = context.workbook.worksheets.getItem(item.sheetName);
        const cell = worksheet.getCell(item.row, item.column);
        cell.values = [[item.suggestedValue]];
        item.status = "applied";
        item.selected = false;
      }
      await context.sync();
    });
    renderQueue();
    notify("Changes applied", `${selected.length} cell${selected.length === 1 ? "" : "s"} updated.`, { sound: false });
  } catch (error) {
    notify("Apply failed", error.message, { sound: false });
  } finally {
    renderQueue();
  }
}

async function deleteSelectedWorkDocRows() {
  primeAudio();
  const ids = state.queue.filter((item) => item.selected && canDeleteRow(item)).map((item) => item.id);
  if (ids.length === 0) {
    notify("No rows selected", "Select work-doc suggestions to remove their rows.", { sound: false });
    return;
  }

  await deleteWorkDocRowsByIds(ids);
}

async function deleteWorkDocRowsByIds(ids) {
  const selected = state.queue.filter((item) => ids.includes(item.id) && canDeleteRow(item));
  if (selected.length === 0) {
    notify("No rows ready", "Only work-doc suggestions can remove entire rows.", { sound: false });
    return;
  }

  elements.deleteRows.disabled = true;
  try {
    const bySheet = groupBySheetName(selected);
    const deletedRowsBySheet = new Map();

    await Excel.run(async (context) => {
      for (const [sheetName, items] of bySheet.entries()) {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const sortedRows = [...new Set(items.map((item) => item.row))].sort((a, b) => b - a);
        deletedRowsBySheet.set(sheetName, sortedRows);
        for (const row of sortedRows) {
          const rowRange = worksheet.getRangeByIndexes(row, 0, 1, 1).getEntireRow();
          rowRange.delete(Excel.DeleteShiftDirection.up);
        }
      }
      await context.sync();
    });

    reconcileQueueAfterRowDelete(deletedRowsBySheet, selected.map((item) => item.id));
    renderQueue();
    const deletedRowCount = [...deletedRowsBySheet.values()].reduce((total, rows) => total + rows.length, 0);
    notify("Rows removed", `${deletedRowCount} work-doc row${deletedRowCount === 1 ? "" : "s"} removed.`, { sound: false });
  } catch (error) {
    notify("Remove rows failed", error.message, { sound: false });
  } finally {
    renderQueue();
  }
}

function groupBySheetName(items) {
  const groups = new Map();
  for (const item of items) {
    if (!groups.has(item.sheetName)) {
      groups.set(item.sheetName, []);
    }
    groups.get(item.sheetName).push(item);
  }
  return groups;
}

function reconcileQueueAfterRowDelete(deletedRowsBySheet, removedIds) {
  const removedIdSet = new Set(removedIds);
  for (const item of state.queue) {
    const deletedRows = deletedRowsBySheet.get(item.sheetName);
    if (!deletedRows) {
      continue;
    }

    if (removedIdSet.has(item.id) || deletedRows.includes(item.row)) {
      item.status = "removed";
      item.selected = false;
      continue;
    }

    const shift = deletedRows.filter((row) => row < item.row).length;
    if (shift > 0) {
      item.row -= shift;
      item.location = item.column >= 0
        ? `${item.sheetName}!R${item.row + 1}C${item.column + 1}`
        : `${item.sheetName}!row ${item.row + 1}`;
      item.dedupeKey = makeSuggestionKey(item);
    }
  }
}

function removeSelectedSuggestions() {
  const ids = state.queue.filter((item) => item.selected || item.status === "applied").map((item) => item.id);
  if (ids.length === 0) {
    notify("Nothing selected", "Select suggestions to remove from the queue.", { sound: false });
    return;
  }

  removeSuggestionIds(ids);
}

function removeSuggestionIds(ids) {
  const selected = state.queue.filter((item) => ids.includes(item.id));
  if (selected.length === 0) {
    notify("Nothing selected", "Select suggestions to remove from the queue.", { sound: false });
    return;
  }

  for (const item of selected) {
    item.status = "removed";
    item.selected = false;
  }
  renderQueue();
  notify("Queue updated", `${selected.length} suggestion${selected.length === 1 ? "" : "s"} removed.`, { sound: false });
}

function clearQueue() {
  for (const item of state.queue) {
    item.status = "removed";
    item.selected = false;
  }
  renderQueue();
  notify("Queue cleared", "All suggestions removed.", { sound: false });
}

async function openReviewWindow() {
  primeAudio();
  syncQueueToStorage();

  if (!Office.context?.ui?.displayDialogAsync) {
    notify("Review window unavailable", "Open this from Excel to use the separate review window.", { sound: false });
    return;
  }

  if (state.reviewDialog) {
    notify("Review window already open", "Use the existing review window.", { sound: false });
    return;
  }

  const reviewUrl = new URL("./review.html", window.location.href).toString();
  Office.context.ui.displayDialogAsync(
    reviewUrl,
    { height: 72, width: 42, displayInIframe: false },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        notify("Could not open review window", result.error?.message || "Excel did not open the window.", { sound: false });
        return;
      }

      state.reviewDialog = result.value;
      state.reviewDialog.addEventHandler(Office.EventType.DialogMessageReceived, handleReviewWindowMessage);
      state.reviewDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        state.reviewDialog = null;
      });
      notify("Review window open", "You can move it to another screen.", { sound: false });
    }
  );
}

async function handleReviewWindowMessage(args) {
  let message;
  try {
    message = JSON.parse(args.message);
  } catch {
    return;
  }

  if (message.type === "requestState") {
    syncQueueToStorage();
    return;
  }

  if (message.type === "apply") {
    await applySuggestionIds(message.ids || []);
    return;
  }

  if (message.type === "deleteRows") {
    await deleteWorkDocRowsByIds(message.ids || []);
    return;
  }

  if (message.type === "remove") {
    removeSuggestionIds(message.ids || []);
  }
}

async function hideTaskPane() {
  if (!Office.addin?.hide) {
    notify("Hide pane unavailable", "This needs shared runtime support in Excel.", { sound: false });
    return;
  }

  try {
    await Office.addin.hide();
  } catch (error) {
    notify("Could not hide pane", error.message, { sound: false });
  }
}

function syncQueueToStorage() {
  localStorage.setItem("referenceChecker.queue", JSON.stringify(state.queue));
}

function makeColumnMap(headers, names) {
  return Object.fromEntries(names.map((name) => [name, headers.indexOf(name)]));
}

function notify(title, message, options = {}) {
  if (state.settings.alertsEnabled) {
    showToast(title, message);
  }
  if (options.sound !== false) {
    playWarningSound().catch(() => {});
  }
}

function showToast(title, message) {
  window.clearTimeout(state.toastTimer);
  elements.toastTitle.textContent = title;
  elements.toastMessage.textContent = message;
  elements.toast.hidden = false;
  state.toastTimer = window.setTimeout(() => {
    elements.toast.hidden = true;
  }, 4500);
}

function primeAudio() {
  if (!state.settings.soundEnabled) {
    return Promise.resolve(false);
  }
  const context = getAudioContext();
  if (context?.state === "suspended") {
    return context.resume().then(() => true).catch(() => false);
  }
  return Promise.resolve(Boolean(context));
}

async function testSound() {
  state.settings.soundEnabled = true;
  elements.soundEnabled.checked = true;
  saveSettings();
  const played = await playWarningSound({ force: true });
  if (played) {
    showToast("Sound enabled", "You should hear this alert when new suggestions appear.");
  } else {
    showToast("Sound blocked", "Click the review window and try Enable / Test Sound there too.");
  }
}

async function playWarningSound({ force = false } = {}) {
  if (!state.settings.soundEnabled) {
    return false;
  }

  const context = getAudioContext();
  if (!context) {
    return false;
  }
  if (context.state === "suspended") {
    try {
      await context.resume();
    } catch {
      return false;
    }
  }

  if (!force && context.state !== "running") {
    return false;
  }

  const now = context.currentTime;
  playTone(context, now, 660, 0.08, 0.045);
  playTone(context, now + 0.1, 880, 0.1, 0.04);
  return true;
}

function getAudioContext() {
  if (state.audioContext) {
    return state.audioContext;
  }

  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) {
    return null;
  }

  state.audioContext = new AudioContextClass();
  return state.audioContext;
}

function playTone(context, startTime, frequency, duration, volume) {
  const oscillator = context.createOscillator();
  const gain = context.createGain();
  oscillator.type = "sine";
  oscillator.frequency.setValueAtTime(frequency, startTime);
  gain.gain.setValueAtTime(0.0001, startTime);
  gain.gain.exponentialRampToValueAtTime(volume, startTime + 0.015);
  gain.gain.exponentialRampToValueAtTime(0.0001, startTime + duration);
  oscillator.connect(gain).connect(context.destination);
  oscillator.start(startTime);
  oscillator.stop(startTime + duration + 0.02);
}

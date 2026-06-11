const STORAGE_KEY = "referenceChecker.queue";
const SETTINGS_KEY = "referenceChecker.reviewSettings";

const state = {
  queue: [],
  selectedIds: new Set(),
  previousCount: 0,
  audioContext: null,
  toastTimer: null,
  soundEnabled: true
};

const elements = {
  status: document.getElementById("reviewStatus"),
  count: document.getElementById("reviewCount"),
  soundEnabled: document.getElementById("reviewSoundEnabled"),
  testSound: document.getElementById("reviewTestSound"),
  selectAll: document.getElementById("reviewSelectAll"),
  applySelected: document.getElementById("reviewApplySelected"),
  deleteRows: document.getElementById("reviewDeleteRows"),
  removeSelected: document.getElementById("reviewRemoveSelected"),
  queueList: document.getElementById("reviewQueueList"),
  toast: document.getElementById("reviewToast"),
  toastTitle: document.getElementById("reviewToastTitle"),
  toastMessage: document.getElementById("reviewToastMessage")
};

loadSettings();
wireUi();
Office.onReady(() => {
  elements.status.textContent = "Review window connected.";
  sendToParent({ type: "requestState" });
  refreshFromStorage({ notifyNew: false });
  window.setInterval(() => refreshFromStorage({ notifyNew: true }), 1000);
});

function wireUi() {
  elements.soundEnabled.checked = state.soundEnabled;
  elements.soundEnabled.addEventListener("change", () => {
    state.soundEnabled = elements.soundEnabled.checked;
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ soundEnabled: state.soundEnabled }));
    primeAudio();
  });
  elements.testSound.addEventListener("click", () => testSound());
  elements.selectAll.addEventListener("click", selectAllSuggestions);
  elements.applySelected.addEventListener("click", () => {
    const ids = selectedReadyIds();
    if (ids.length > 0) {
      sendToParent({ type: "apply", ids });
    }
  });
  elements.deleteRows.addEventListener("click", () => {
    const ids = selectedDeletableIds();
    if (ids.length > 0) {
      sendToParent({ type: "deleteRows", ids });
    }
  });
  elements.removeSelected.addEventListener("click", () => {
    const ids = [...state.selectedIds];
    if (ids.length > 0) {
      sendToParent({ type: "remove", ids });
      state.selectedIds.clear();
      updateActions();
    }
  });
  elements.queueList.addEventListener("change", (event) => {
    const checkbox = event.target.closest("input[type='checkbox']");
    if (!checkbox) {
      return;
    }
    const id = checkbox.closest(".suggestion")?.dataset.id;
    if (!id) {
      return;
    }
    if (checkbox.checked) {
      state.selectedIds.add(id);
    } else {
      state.selectedIds.delete(id);
    }
    updateActions();
  });
}

function loadSettings() {
  try {
    const cached = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    state.soundEnabled = cached.soundEnabled ?? true;
  } catch {
    state.soundEnabled = true;
  }
}

function refreshFromStorage({ notifyNew }) {
  const raw = localStorage.getItem(STORAGE_KEY);
  const nextQueue = raw ? JSON.parse(raw) : [];
  const nextVisible = nextQueue.filter((item) => item.status !== "removed");
  const newCount = Math.max(0, nextVisible.length - state.previousCount);
  state.queue = nextQueue;
  state.previousCount = nextVisible.length;
  renderQueue();

  if (notifyNew && newCount > 0) {
    notify(`${newCount} new suggestion${newCount === 1 ? "" : "s"}`, "Review when ready.");
  }
}

function renderQueue() {
  const visibleQueue = state.queue.filter((item) => item.status !== "removed");
  const visibleIds = new Set(visibleQueue.map((item) => item.id));
  for (const id of state.selectedIds) {
    if (!visibleIds.has(id)) {
      state.selectedIds.delete(id);
    }
  }
  elements.count.textContent = String(visibleQueue.length);
  elements.count.className = visibleQueue.length > 0 ? "badge badge-ready" : "badge badge-muted";
  elements.queueList.replaceChildren(...visibleQueue.map(renderSuggestion));
  updateActions();
}

function renderSuggestion(item) {
  const wrapper = document.createElement("article");
  wrapper.className = `suggestion ${item.type === "work-doc" ? "work-doc" : ""} ${item.ambiguous ? "ambiguous" : ""} ${item.status === "applied" ? "applied" : ""}`;
  wrapper.dataset.id = item.id;

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = state.selectedIds.has(item.id);
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

function selectedReadyIds() {
  return state.queue
    .filter((item) => state.selectedIds.has(item.id) && canApply(item))
    .map((item) => item.id);
}

function selectedDeletableIds() {
  return state.queue
    .filter((item) => state.selectedIds.has(item.id) && canDeleteRow(item))
    .map((item) => item.id);
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
    showToast("Nothing to select", "There are no suggestions in the queue.");
    return;
  }

  state.selectedIds = new Set(visibleQueue.map((item) => item.id));
  renderQueue();
  showToast("Suggestions selected", `${visibleQueue.length} item${visibleQueue.length === 1 ? "" : "s"} selected.`);
}

function updateActions() {
  const visibleQueue = state.queue.filter((item) => item.status !== "removed");
  elements.selectAll.disabled = visibleQueue.length === 0;
  elements.applySelected.disabled = selectedReadyIds().length === 0;
  elements.deleteRows.disabled = selectedDeletableIds().length === 0;
  elements.removeSelected.disabled = state.selectedIds.size === 0;
}

function sendToParent(message) {
  const payload = JSON.stringify(message);
  if (Office.context?.ui?.messageParent) {
    Office.context.ui.messageParent(payload);
  } else {
    console.log(payload);
  }
}

function notify(title, message) {
  showToast(title, message);
  playWarningSound().catch(() => {});
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
  if (!state.soundEnabled) {
    return Promise.resolve(false);
  }
  const context = getAudioContext();
  if (context?.state === "suspended") {
    return context.resume().then(() => true).catch(() => false);
  }
  return Promise.resolve(Boolean(context));
}

async function testSound() {
  state.soundEnabled = true;
  elements.soundEnabled.checked = true;
  localStorage.setItem(SETTINGS_KEY, JSON.stringify({ soundEnabled: state.soundEnabled }));
  const played = await playWarningSound({ force: true });
  if (played) {
    showToast("Sound enabled", "You should hear this alert when new suggestions appear.");
  } else {
    showToast("Sound blocked", "Click inside this window and try again.");
  }
}

async function playWarningSound({ force = false } = {}) {
  if (!state.soundEnabled) {
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

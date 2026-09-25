const STORAGE_KEY = 'safmeds.v1';
const HISTORY_LIMIT = 20;
const ROUND_SECONDS = 60;
const DEFAULT_DECK_FILES = ['spreadsheets/SAFMEDs 1.xlsx'];

const state = {
  decks: loadState(),
  currentDeckName: null,
  practice: null,
  preview: null,
};
const cardPointerStates = new WeakMap();

const screens = {
  setup: document.getElementById('setup-screen'),
  practice: document.getElementById('practice-screen'),
  preview: document.getElementById('preview-screen'),
  results: document.getElementById('results-screen'),
  progress: document.getElementById('progress-screen'),
};

const deckList = document.getElementById('deck-list');
const defaultDeckSelect = document.getElementById('default-deck-select');
const defaultDeckForm = document.getElementById('default-deck-form');
const loadDefaultDeckBtn = document.getElementById('load-default-deck');
const defaultDeckMessage = document.getElementById('default-deck-message');
const appScriptSrc = document.querySelector('script[src$="app.js"]')?.getAttribute('src') || 'app.js';
const appBaseUrl = new URL(appScriptSrc, window.location.href);
const appBaseDirUrl = new URL('.', appBaseUrl);

const practiceTitle = document.getElementById('practice-title');
const timerEl = document.getElementById('timer');
const practiceStats = document.getElementById('practice-stats');
const cardFace = document.getElementById('card-face');
const markCorrectBtn = document.getElementById('mark-correct');
const markIncorrectBtn = document.getElementById('mark-incorrect');
const endEarlyBtn = document.getElementById('end-early');

const previewTitle = document.getElementById('preview-title');
const previewStats = document.getElementById('preview-stats');
const previewCardFace = document.getElementById('preview-card-face');
const previewPreviousBtn = document.getElementById('preview-previous');
const previewNextBtn = document.getElementById('preview-next');
const previewEditToggleBtn = document.getElementById('preview-edit-toggle');
const previewStartPracticeBtn = document.getElementById('preview-start-practice');
const previewBackBtn = document.getElementById('preview-back');
const previewEditMessage = document.getElementById('preview-edit-message');
const previewEditPanel = document.getElementById('preview-edit-panel');
const previewEditTerm = document.getElementById('preview-edit-term');
const previewEditDefinition = document.getElementById('preview-edit-definition');
const previewHighlightSelectionBtn = document.getElementById('preview-highlight-selection');
const previewRemoveHighlightBtn = document.getElementById('preview-remove-highlight');
const previewSaveEditBtn = document.getElementById('preview-save-edit');
const previewCancelEditBtn = document.getElementById('preview-cancel-edit');
const previewDiscardConfirm = document.getElementById('preview-discard-confirm');
const previewConfirmDiscardBtn = document.getElementById('preview-confirm-discard');
const previewKeepEditingBtn = document.getElementById('preview-keep-editing');

const resultsTitle = document.getElementById('results-title');
const headlineRate = document.getElementById('headline-rate');
const secondaryStats = document.getElementById('secondary-stats');
const resultsChart = document.getElementById('results-chart');
const resultsChartEmpty = document.getElementById('results-chart-empty');
const resultsHistory = document.getElementById('results-history');

const progressTitle = document.getElementById('progress-title');
const progressChart = document.getElementById('progress-chart');
const progressChartEmpty = document.getElementById('progress-chart-empty');
const progressHistory = document.getElementById('progress-history');

bindEvents();
renderDefaultDeckOptions();
renderDecks();
showScreen('setup');

function bindEvents() {
  if (defaultDeckForm) {
    defaultDeckForm.addEventListener('submit', onLoadDefaultDeck);
  }
  if (defaultDeckSelect) {
    defaultDeckSelect.addEventListener('change', () => {
      defaultDeckSelect.setAttribute('aria-invalid', 'false');
    });
  }

  bindCardInteractions(cardFace, onCardFaceClick);
  markCorrectBtn.addEventListener('click', () => markAnswer(true));
  markIncorrectBtn.addEventListener('click', () => markAnswer(false));
  endEarlyBtn.addEventListener('click', () => finishPractice(true));

  bindCardInteractions(previewCardFace, onPreviewCardFaceClick);
  previewPreviousBtn.addEventListener('click', () => changePreviewCard(-1));
  previewNextBtn.addEventListener('click', () => changePreviewCard(1));
  previewEditToggleBtn.addEventListener('click', togglePreviewEditor);
  previewStartPracticeBtn.addEventListener('click', () => {
    if (state.currentDeckName) startPractice(state.currentDeckName);
  });
  previewBackBtn.addEventListener('click', closePreview);
  previewHighlightSelectionBtn.addEventListener('mousedown', (event) => event.preventDefault());
  previewRemoveHighlightBtn.addEventListener('mousedown', (event) => event.preventDefault());
  previewHighlightSelectionBtn.addEventListener('click', applyHighlightToPreviewSelection);
  previewRemoveHighlightBtn.addEventListener('click', removeHighlightFromPreviewSelection);
  previewSaveEditBtn.addEventListener('click', savePreviewEdits);
  previewCancelEditBtn.addEventListener('click', () => closePreviewEditor());
  previewConfirmDiscardBtn.addEventListener('click', confirmDiscardPreviewEdits);
  previewKeepEditingBtn.addEventListener('click', cancelDiscardPreviewEdits);
  previewEditTerm.addEventListener('focus', () => setPreviewEditField('term'));
  previewEditDefinition.addEventListener('focus', () => setPreviewEditField('definition'));
  previewEditTerm.addEventListener('input', cancelDiscardPreviewEdits);
  previewEditDefinition.addEventListener('input', cancelDiscardPreviewEdits);
  document.getElementById('practice-back').addEventListener('click', cancelPractice);
  document.getElementById('results-back').addEventListener('click', () => showScreen('setup'));
  document.getElementById('practice-again').addEventListener('click', () => {
    if (state.currentDeckName) startPractice(state.currentDeckName);
  });
  document.getElementById('progress-back').addEventListener('click', () => showScreen('setup'));
  document.getElementById('progress-preview').addEventListener('click', () => {
    if (state.currentDeckName) openPreview(state.currentDeckName, 'progress');
  });
  document.getElementById('progress-practice').addEventListener('click', () => {
    if (state.currentDeckName) startPractice(state.currentDeckName);
  });

  window.addEventListener('keydown', onPracticeHotkeys);
  window.addEventListener('keydown', onPreviewHotkeys);
}

function bindCardInteractions(cardEl, onClick) {
  cardPointerStates.set(cardEl, createCardPointerState());
  cardEl.addEventListener('pointerdown', onCardPointerDown);
  cardEl.addEventListener('pointermove', onCardPointerMove);
  cardEl.addEventListener('pointerup', onCardPointerUp);
  cardEl.addEventListener('pointercancel', onCardPointerCancel);
  cardEl.addEventListener('click', onClick);
}

function createCardPointerState() {
  return {
    pointerId: null,
    startX: 0,
    startY: 0,
    moved: false,
    suppressNextClick: false,
  };
}

function getCardPointerState(cardEl) {
  let pointerState = cardPointerStates.get(cardEl);
  if (!pointerState) {
    pointerState = createCardPointerState();
    cardPointerStates.set(cardEl, pointerState);
  }
  return pointerState;
}

function onCardPointerDown(event) {
  if (!event.isPrimary) return;
  if (event.pointerType === 'mouse' && event.button !== 0) return;
  const pointerState = getCardPointerState(event.currentTarget);
  pointerState.suppressNextClick = false;
  pointerState.pointerId = event.pointerId;
  pointerState.startX = event.clientX;
  pointerState.startY = event.clientY;
  pointerState.moved = false;
}

function onCardPointerMove(event) {
  const pointerState = getCardPointerState(event.currentTarget);
  if (!event.isPrimary) return;
  if (pointerState.pointerId !== event.pointerId) return;
  if (Math.abs(event.clientX - pointerState.startX) > 4 || Math.abs(event.clientY - pointerState.startY) > 4) {
    pointerState.moved = true;
  }
}

function onCardPointerUp(event) {
  const pointerState = getCardPointerState(event.currentTarget);
  if (!event.isPrimary) return;
  if (pointerState.pointerId !== event.pointerId) return;
  const shouldSuppress = pointerState.moved;
  resetCardPointerTracking(pointerState);
  pointerState.suppressNextClick = shouldSuppress;
}

function onCardPointerCancel(event) {
  resetCardPointerState(getCardPointerState(event.currentTarget));
}

function resetCardPointerTracking(pointerState) {
  pointerState.pointerId = null;
  pointerState.startX = 0;
  pointerState.startY = 0;
  pointerState.moved = false;
}

function resetCardPointerState(pointerState) {
  resetCardPointerTracking(pointerState);
  pointerState.suppressNextClick = false;
}

function onCardFaceClick(event) {
  if (shouldKeepCardSelection(event, cardFace)) return;
  flipCard();
}

function onPreviewCardFaceClick(event) {
  if (state.preview?.isEditing) return;
  if (shouldKeepCardSelection(event, previewCardFace)) return;
  flipPreviewCard();
}

function shouldKeepCardSelection(event, cardEl) {
  const pointerState = getCardPointerState(cardEl);
  const shouldKeep = hasTextSelectionWithin(cardEl) || (event.detail > 0 && pointerState.suppressNextClick);
  resetCardPointerState(pointerState);
  return shouldKeep;
}

function hasTextSelectionWithin(element) {
  const selection = window.getSelection?.();
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) return false;

  for (let i = 0; i < selection.rangeCount; i += 1) {
    const container = selection.getRangeAt(i).commonAncestorContainer;
    if (element.contains(container.nodeType === Node.ELEMENT_NODE ? container : container.parentNode)) {
      return true;
    }
  }

  return false;
}

function getTrimmedSelectionRange(input) {
  const start = input.selectionStart ?? 0;
  const end = input.selectionEnd ?? 0;
  if (end <= start) return null;

  const selectedText = input.value.slice(start, end);
  const leadingWhitespace = selectedText.match(/^\s*/u)?.[0].length || 0;
  const trailingWhitespace = selectedText.match(/\s*$/u)?.[0].length || 0;
  const trimmedStart = start + leadingWhitespace;
  const trimmedEnd = end - trailingWhitespace;
  return trimmedEnd > trimmedStart ? { start: trimmedStart, end: trimmedEnd } : null;
}

function isWrappedInHighlight(value, start, end) {
  return getSurroundingHighlightRange(value, start, end) !== null;
}

function getSurroundingHighlightRange(value, start, end) {
  if (start >= 2 && value.slice(start - 2, start) === '==' && value.slice(end, end + 2) === '==') {
    return { start: start - 2, end: end + 2 };
  }

  const selectedText = value.slice(start, end);
  if (selectedText.startsWith('==') && selectedText.endsWith('==') && selectedText.length > 4) {
    return { start, end };
  }

  return null;
}

function renderDefaultDeckOptions() {
  if (!defaultDeckSelect || !loadDefaultDeckBtn) return;

  for (const fileName of DEFAULT_DECK_FILES) {
    const option = document.createElement('option');
    option.value = fileName;
    option.textContent = getDeckNameFromSource(fileName);
    defaultDeckSelect.append(option);
  }

  loadDefaultDeckBtn.disabled = DEFAULT_DECK_FILES.length === 0;
}

async function onLoadDefaultDeck(event) {
  event?.preventDefault();
  const fileName = defaultDeckSelect?.value;
  if (!defaultDeckSelect?.checkValidity()) {
    defaultDeckSelect?.setAttribute('aria-invalid', 'true');
    defaultDeckSelect?.reportValidity();
    setMessage(defaultDeckMessage, 'Please select a default deck before loading.', 'error');
    return;
  }
  defaultDeckSelect?.setAttribute('aria-invalid', 'false');

  setMessage(defaultDeckMessage, '');
  loadDefaultDeckBtn.disabled = true;
  try {
    const fileUrl = new URL(fileName, appBaseDirUrl).toString();
    const response = await fetch(fileUrl);
    if (!response.ok) {
      throw new Error(`Could not load ${fileName}.`);
    }
    const fileData = await response.arrayBuffer();
    const { deckName, cardCount } = importDeckFromArrayBuffer(fileData, fileName);
    setMessage(defaultDeckMessage, `Saved "${deckName}" with ${cardCount} cards.`, 'success');
  } catch (error) {
    setMessage(defaultDeckMessage, `Default deck load failed: ${error.message || 'Unable to read spreadsheet.'}`, 'error');
  } finally {
    loadDefaultDeckBtn.disabled = DEFAULT_DECK_FILES.length === 0;
  }
}

function loadState() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    return parsed && typeof parsed === 'object' ? normalizeDecks(parsed) : {};
  } catch {
    return {};
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.decks));
}

function normalizeDecks(rawDecks) {
  return Object.fromEntries(
    Object.entries(rawDecks).map(([deckKey, rawDeck]) => {
      const baseCards = cloneCards(rawDeck?.baseCards?.length ? rawDeck.baseCards : rawDeck?.cards || []);
      const overrides = normalizeOverrides(rawDeck?.overrides);
      const cards = applyCardOverrides(baseCards, overrides);
      return [deckKey, {
        name: rawDeck?.name || deckKey,
        baseCards,
        overrides,
        cards,
        history: Array.isArray(rawDeck?.history) ? rawDeck.history : [],
        signature: typeof rawDeck?.signature === 'string' ? rawDeck.signature : getDeckSignature(baseCards),
      }];
    })
  );
}

function cloneCards(cards) {
  return (Array.isArray(cards) ? cards : [])
    .map((card) => ({
      term: String(card?.term ?? '').trim(),
      definition: String(card?.definition ?? '').trim(),
    }))
    .filter((card) => card.term && card.definition);
}

function normalizeOverrides(rawOverrides) {
  if (!rawOverrides || typeof rawOverrides !== 'object') return {};

  return Object.fromEntries(
    Object.entries(rawOverrides)
      .filter(([key]) => /^\d+$/.test(key))
      .map(([key, card]) => [key, {
        term: String(card?.term ?? '').trim(),
        definition: String(card?.definition ?? '').trim(),
      }])
      .filter(([, card]) => card.term && card.definition)
  );
}

function applyCardOverrides(baseCards, overrides = {}) {
  return cloneCards(baseCards).map((card, index) => ({
    term: overrides[index]?.term ?? card.term,
    definition: overrides[index]?.definition ?? card.definition,
  }));
}

function importDeckFromArrayBuffer(fileData, sourceName) {
  const workbook = XLSX.read(fileData, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) throw new Error('No worksheets found in file.');

  const baseCards = parseCards(firstSheet);
  if (!baseCards.length) throw new Error('No valid term/definition rows were found.');

  const deckName = getDeckNameFromSource(sourceName);
  const signature = getDeckSignature(baseCards);
  const previousDeck = state.decks[deckName];
  const previousHistory = previousDeck?.signature === signature ? previousDeck.history || [] : [];
  const overrides = previousDeck?.signature === signature ? normalizeOverrides(previousDeck.overrides) : {};
  state.decks[deckName] = {
    name: deckName,
    baseCards,
    overrides,
    cards: applyCardOverrides(baseCards, overrides),
    history: previousHistory,
    signature,
  };
  saveState();
  renderDecks();
  return { deckName, cardCount: baseCards.length };
}

function getDeckNameFromSource(sourceName) {
  const sourceBaseName = String(sourceName || '').split(/[\\/]/).pop() || '';
  return sourceBaseName.replace(/\.[^.]+$/, '') || 'Imported Deck';
}

function getDeckSignature(cards) {
  let hash = 0;
  for (const card of cards) {
    const text = `${card.term}\u0000${card.definition}\u0001`;
    for (let i = 0; i < text.length; i += 1) {
      hash = (hash * 31 + text.charCodeAt(i)) >>> 0;
    }
  }
  return `${cards.length}:${hash}`;
}

function parseCards(sheet) {
  if (!sheet?.['!ref']) return [];

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const rows = [];
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    rows.push([
      getSheetCellText(sheet, rowIndex, 0),
      getSheetCellText(sheet, rowIndex, 1),
    ]);
  }
  if (!rows.length) return [];

  let startIndex = 0;
  const firstTerm = String(rows[0]?.[0] ?? '').trim().toLowerCase();
  const firstDef = String(rows[0]?.[1] ?? '').trim().toLowerCase();
  if (firstTerm === 'term' && firstDef === 'definition') {
    startIndex = 1;
  }

  const cards = [];
  for (let i = startIndex; i < rows.length; i += 1) {
    const term = String(rows[i]?.[0] ?? '').trim();
    const definition = String(rows[i]?.[1] ?? '').trim();
    if (!term || !definition) continue;
    cards.push({ term, definition });
  }
  return cards;
}

function getSheetCellText(sheet, rowIndex, colIndex) {
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[cellAddress];
  if (!cell) return '';

  const imageSource = parseSpreadsheetImageFormula(cell.f);
  if (imageSource) return imageSource;

  const rawValue = cell.w ?? cell.v ?? '';
  return String(rawValue ?? '').trim();
}

function parseSpreadsheetImageFormula(formula) {
  const formulaText = String(formula ?? '').trim();
  if (!formulaText) return '';

  const match = formulaText.match(/^(?:=)?(?:_xlfn\.)?IMAGE\(\s*"((?:[^"]|"")*)"/iu);
  if (!match) return '';

  const imageSource = match[1].replaceAll('""', '"').trim();
  return imageSource ? `![Spreadsheet image](${imageSource})` : '';
}

function updateDeckCard(deck, index, term, definition) {
  if (!deck || !Array.isArray(deck.baseCards) || !deck.baseCards.length) return;
  if (!deck.overrides || typeof deck.overrides !== 'object') {
    deck.overrides = {};
  }

  const nextCard = {
    term: String(term ?? '').trim(),
    definition: String(definition ?? '').trim(),
  };
  const baseCard = deck.baseCards[index];
  if (!baseCard) return;

  if (baseCard.term === nextCard.term && baseCard.definition === nextCard.definition) {
    delete deck.overrides[index];
  } else {
    deck.overrides[index] = nextCard;
  }

  deck.cards = applyCardOverrides(deck.baseCards, deck.overrides);
}

function setMessage(target, text, type = '') {
  if (!target) return;
  target.textContent = text;
  target.className = `message ${type}`.trim();
}

function renderDecks() {
  const names = Object.keys(state.decks).sort((a, b) => a.localeCompare(b));
  if (!names.length) {
    deckList.innerHTML = '<p class="empty">No decks yet. Load a spreadsheet to begin.</p>';
    return;
  }

  deckList.innerHTML = '';
  for (const name of names) {
    const deck = state.decks[name];
    const row = document.createElement('div');
    row.className = 'deck-row';
    if (name === state.currentDeckName) {
      row.classList.add('selected');
      row.setAttribute('aria-current', 'true');
    }

    const meta = document.createElement('div');
    meta.className = 'deck-meta';
    meta.innerHTML = `<strong>${escapeHtml(deck.name)}</strong><span>${deck.cards.length} cards · ${deck.history?.length || 0} timings</span>`;

    const actions = document.createElement('div');
    actions.className = 'controls';

    const previewBtn = createButton('Preview', () => openPreview(name));
    const progressBtn = createButton('Progress', () => openProgress(name));
    const practiceBtn = createButton('Practice', () => startPractice(name));
    const deleteBtn = createButton('Delete', () => deleteDeck(name), 'delete');

    actions.append(previewBtn, progressBtn, practiceBtn, deleteBtn);
    row.append(meta, actions);
    deckList.append(row);
  }
}

function createButton(text, onClick, extraClass = '') {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.textContent = text;
  if (extraClass) btn.classList.add(extraClass);
  btn.addEventListener('click', onClick);
  return btn;
}

function deleteDeck(name) {
  if (!window.confirm(`Delete deck "${name}" and all its history?`)) return;
  delete state.decks[name];
  saveState();
  if (state.currentDeckName === name) state.currentDeckName = null;
  renderDecks();
}

function openProgress(deckName) {
  state.currentDeckName = deckName;
  const deck = state.decks[deckName];
  progressTitle.textContent = `Progress: ${deckName}`;
  renderHistoryTable(progressHistory, deck.history || []);
  renderChart(progressChart, progressChartEmpty, deck.history || []);
  showScreen('progress');
}

function openPreview(deckName, returnScreen = 'setup') {
  const deck = state.decks[deckName];
  if (!deck || !deck.cards.length) return;

  state.currentDeckName = deckName;
  state.preview = {
    index: 0,
    showingDefinition: false,
    returnScreen,
    isEditing: false,
    editField: 'term',
    confirmingDiscard: false,
  };

  setMessage(previewEditMessage, '');
  syncPreviewEditor();
  previewTitle.textContent = `Preview: ${deckName}`;
  updatePreviewCard();
  updatePreviewStats();
  showScreen('preview');
}

function startPractice(deckName) {
  const deck = state.decks[deckName];
  if (!deck || !deck.cards.length) return;

  if (state.practice?.timerId) clearInterval(state.practice.timerId);
  setMessage(previewEditMessage, '');
  state.preview = null;

  state.currentDeckName = deckName;
  const round = shuffleDeck(deck.cards);
  state.practice = {
    round,
    index: 0,
    showingDefinition: false,
    correct: 0,
    incorrect: 0,
    seen: 0,
    startAt: Date.now(),
    endsAt: Date.now() + ROUND_SECONDS * 1000,
    timerId: null,
    finished: false,
  };

  practiceTitle.textContent = `Practice: ${deckName}`;
  updateCardFace();
  updatePracticeStats();
  updateTimer();

  state.practice.timerId = setInterval(() => {
    if (!state.practice || state.practice.finished) return;
    const remainingMs = state.practice.endsAt - Date.now();
    if (remainingMs <= 0) {
      timerEl.textContent = '0.0s';
      finishPractice(false);
      return;
    }
    updateTimer();
  }, 100);

  showScreen('practice');
}

function cancelPractice() {
  if (state.practice?.timerId) clearInterval(state.practice.timerId);
  state.practice = null;
  showScreen('setup');
}

function updateTimer() {
  const remaining = Math.max(0, (state.practice.endsAt - Date.now()) / 1000);
  timerEl.textContent = `${remaining.toFixed(1)}s`;
  timerEl.classList.toggle('urgent', remaining <= 10);
}

function updatePracticeStats() {
  const deck = state.decks[state.currentDeckName];
  const total = deck.cards.length;
  const cardNum = (state.practice.index % total) + 1;
  practiceStats.textContent = `Correct: ${state.practice.correct} · Incorrect: ${state.practice.incorrect} · Card ${cardNum} of ${total}`;
}

function updateCardFace() {
  const currentCard = getCurrentCard();
  if (!currentCard) {
    cardFace.textContent = 'No card';
    return;
  }
  const faceText = state.practice.showingDefinition ? currentCard.definition : currentCard.term;
  cardFace.innerHTML = formatCardText(faceText);
}

function updatePreviewCard() {
  const currentCard = getCurrentPreviewCard();
  if (!currentCard) {
    previewCardFace.innerHTML = '';
    previewCardFace.textContent = 'No card';
    return;
  }
  const faceText = state.preview.showingDefinition ? currentCard.definition : currentCard.term;
  previewCardFace.innerHTML = formatCardText(faceText);
}

function getCurrentCard() {
  if (!state.practice?.round.length) return null;
  return state.practice.round[state.practice.index];
}

function getCurrentPreviewCard() {
  const deck = state.decks[state.currentDeckName];
  if (!state.preview || !deck?.cards.length) return null;
  return deck.cards[state.preview.index];
}

function flipCard() {
  if (!state.practice || state.practice.finished) return;
  state.practice.showingDefinition = !state.practice.showingDefinition;
  updateCardFace();
}

function flipPreviewCard() {
  if (!state.preview || state.preview.isEditing) return;
  state.preview.showingDefinition = !state.preview.showingDefinition;
  updatePreviewCard();
  updatePreviewStats();
}

function changePreviewCard(direction) {
  const deck = state.decks[state.currentDeckName];
  if (!state.preview || state.preview.isEditing || !deck?.cards.length) return;
  const total = deck.cards.length;
  state.preview.index = (state.preview.index + direction + total) % total;
  state.preview.showingDefinition = false;
  setMessage(previewEditMessage, '');
  updatePreviewCard();
  updatePreviewStats();
}

function updatePreviewStats() {
  const deck = state.decks[state.currentDeckName];
  if (!state.preview || !deck?.cards.length) {
    previewStats.textContent = 'No card';
    return;
  }
  const cardNum = state.preview.index + 1;
  const faceLabel = state.preview.showingDefinition ? 'Definition' : 'Term';
  previewStats.textContent = `Card ${cardNum} of ${deck.cards.length} · ${faceLabel}`;
}

function togglePreviewEditor() {
  if (!state.preview) return;
  if (state.preview?.isEditing) {
    closePreviewEditor();
    return;
  }
  openPreviewEditor();
}

function openPreviewEditor() {
  const currentCard = getCurrentPreviewCard();
  if (!state.preview || !currentCard) return;
  state.preview.isEditing = true;
  state.preview.editField = state.preview.showingDefinition ? 'definition' : 'term';
  state.preview.confirmingDiscard = false;
  previewEditTerm.value = currentCard.term;
  previewEditDefinition.value = currentCard.definition;
  setMessage(previewEditMessage, 'Editing is saved locally on this device.', '');
  syncPreviewEditor();
  getActivePreviewEditInput()?.focus();
}

function closePreviewEditor({ force = false } = {}) {
  if (!state.preview?.isEditing) return true;
  if (!force && hasPreviewEditChanges()) {
    state.preview.confirmingDiscard = true;
    syncPreviewEditor();
    return false;
  }
  state.preview.isEditing = false;
  state.preview.confirmingDiscard = false;
  syncPreviewEditor();
  return true;
}

function syncPreviewEditor() {
  const isEditing = !!state.preview?.isEditing;
  const confirmingDiscard = !!state.preview?.confirmingDiscard;
  previewEditPanel.classList.toggle('hidden', !isEditing);
  previewDiscardConfirm.classList.toggle('hidden', !confirmingDiscard);
  previewEditToggleBtn.textContent = isEditing ? 'Close Editor' : 'Edit Card';
  previewCardFace.classList.toggle('editing', isEditing);
  previewCardFace.setAttribute('aria-disabled', isEditing ? 'true' : 'false');
  previewCardFace.tabIndex = isEditing ? -1 : 0;
  previewPreviousBtn.disabled = isEditing;
  previewNextBtn.disabled = isEditing;
  previewStartPracticeBtn.disabled = isEditing;
}

function hasPreviewEditChanges() {
  const currentCard = getCurrentPreviewCard();
  if (!state.preview?.isEditing || !currentCard) return false;
  return previewEditTerm.value !== currentCard.term || previewEditDefinition.value !== currentCard.definition;
}

function setPreviewEditField(field) {
  if (!state.preview) return;
  state.preview.editField = field === 'definition' ? 'definition' : 'term';
}

function getActivePreviewEditInput() {
  if (!state.preview) return null;
  return state.preview.editField === 'definition' ? previewEditDefinition : previewEditTerm;
}

function applyHighlightToPreviewSelection() {
  const input = getActivePreviewEditInput();
  if (!input) return;
  cancelDiscardPreviewEdits();
  const selection = getTrimmedSelectionRange(input);
  if (!selection) {
    setMessage(previewEditMessage, 'Select text in Term or Definition before highlighting.', 'error');
    return;
  }

  if (isWrappedInHighlight(input.value, selection.start, selection.end)) {
    setMessage(previewEditMessage, 'That text is already highlighted.', 'error');
    return;
  }

  input.value = `${input.value.slice(0, selection.start)}==${input.value.slice(selection.start, selection.end)}==${input.value.slice(selection.end)}`;
  input.focus();
  input.setSelectionRange(selection.start + 2, selection.end + 2);
  setMessage(previewEditMessage, 'Highlight markers added. Save edits to keep them on this device.', 'success');
}

function removeHighlightFromPreviewSelection() {
  const input = getActivePreviewEditInput();
  if (!input) return;
  cancelDiscardPreviewEdits();
  const selection = getTrimmedSelectionRange(input);
  if (!selection) {
    setMessage(previewEditMessage, 'Select highlighted text before removing the highlight.', 'error');
    return;
  }

  const surroundingRange = getSurroundingHighlightRange(input.value, selection.start, selection.end);
  if (!surroundingRange) {
    setMessage(previewEditMessage, 'The selected text is not currently wrapped in ==highlight markers==.', 'error');
    return;
  }

  input.value = `${input.value.slice(0, surroundingRange.start)}${input.value.slice(surroundingRange.start + 2, surroundingRange.end - 2)}${input.value.slice(surroundingRange.end)}`;
  input.focus();
  input.setSelectionRange(surroundingRange.start, surroundingRange.end - 4);
  setMessage(previewEditMessage, 'Highlight markers removed. Save edits to keep the change on this device.', 'success');
}

function savePreviewEdits() {
  const deck = state.decks[state.currentDeckName];
  if (!state.preview || !deck) return;
  cancelDiscardPreviewEdits();

  const term = previewEditTerm.value.trim();
  const definition = previewEditDefinition.value.trim();
  if (!term || !definition) {
    setMessage(previewEditMessage, 'Term and Definition both need text before saving.', 'error');
    return;
  }

  updateDeckCard(deck, state.preview.index, term, definition);
  saveState();
  closePreviewEditor({ force: true });
  updatePreviewCard();
  updatePreviewStats();
  setMessage(previewEditMessage, 'Saved locally on this device.', 'success');
}

function confirmDiscardPreviewEdits() {
  if (!state.preview) return;
  closePreviewEditor({ force: true });
  setMessage(previewEditMessage, 'Unsaved edits were discarded.', '');
}

function cancelDiscardPreviewEdits() {
  if (!state.preview?.confirmingDiscard) return;
  state.preview.confirmingDiscard = false;
  syncPreviewEditor();
}

function closePreview() {
  if (!state.preview) {
    showScreen('setup');
    return;
  }
  if (state.preview.isEditing && !closePreviewEditor()) return;
  const { returnScreen } = state.preview;
  setMessage(previewEditMessage, '');
  state.preview = null;
  if (returnScreen === 'progress' && state.currentDeckName) {
    openProgress(state.currentDeckName);
    return;
  }
  showScreen(returnScreen || 'setup');
}

function markAnswer(isCorrect) {
  if (!state.practice || state.practice.finished) return;
  const currentCard = getCurrentCard();
  if (isCorrect) state.practice.correct += 1;
  else state.practice.incorrect += 1;

  state.practice.seen += 1;
  state.practice.index += 1;
  state.practice.showingDefinition = false;

  if (state.practice.index >= state.practice.round.length) {
    state.practice.round = shuffleDeck(state.decks[state.currentDeckName].cards, currentCard);
    state.practice.index = 0;
  }

  updateCardFace();
  updatePracticeStats();
}

function finishPractice(endedEarly) {
  if (!state.practice || state.practice.finished) return;
  state.practice.finished = true;
  clearInterval(state.practice.timerId);

  const elapsed = Math.min(ROUND_SECONDS, Math.max(0, (Date.now() - state.practice.startAt) / 1000));
  const correct = state.practice.correct;
  const incorrect = state.practice.incorrect;

  const entry = {
    date: new Date().toISOString(),
    correct,
    incorrect,
    correctPerMin: ratePerMinute(correct, elapsed),
    incorrectPerMin: ratePerMinute(incorrect, elapsed),
  };

  const deck = state.decks[state.currentDeckName];
  deck.history = [entry, ...(deck.history || [])].slice(0, HISTORY_LIMIT);
  saveState();

  showResults(entry, state.practice.seen, endedEarly ? elapsed : ROUND_SECONDS);
  state.practice = null;
}

function showResults(entry, totalSeen, elapsed) {
  const deckName = state.currentDeckName;
  const deck = state.decks[deckName];

  resultsTitle.textContent = `Timing Results: ${deckName}`;
  headlineRate.textContent = `${entry.correctPerMin.toFixed(1)} correct/min`;
  secondaryStats.textContent = `Correct: ${entry.correct} · Incorrect: ${entry.incorrect} · Cards Seen: ${totalSeen} · Elapsed: ${elapsed.toFixed(1)}s`;

  renderHistoryTable(resultsHistory, deck.history || []);
  renderChart(resultsChart, resultsChartEmpty, deck.history || []);

  showScreen('results');
}

function renderHistoryTable(target, history) {
  if (!history.length) {
    target.innerHTML = '<tr><td colspan="5" class="empty">No timings yet.</td></tr>';
    return;
  }

  target.innerHTML = history
    .map((row) => {
      const date = new Date(row.date);
      const dateLabel = Number.isNaN(date.getTime()) ? row.date : date.toLocaleString();
      return `<tr>
        <td>${escapeHtml(dateLabel)}</td>
        <td>${row.correct}</td>
        <td>${row.incorrect}</td>
        <td>${Number(row.correctPerMin).toFixed(1)}</td>
        <td>${Number(row.incorrectPerMin).toFixed(1)}</td>
      </tr>`;
    })
    .join('');
}

function renderChart(svg, emptyEl, historyMostRecentFirst) {
  const history = [...historyMostRecentFirst].reverse();
  svg.innerHTML = '';

  if (history.length < 2) {
    svg.style.display = 'none';
    emptyEl.textContent = 'Add at least two timings to see a trend chart.';
    return;
  }

  svg.style.display = 'block';
  emptyEl.textContent = '';

  const width = 700;
  const height = 260;
  const margin = { top: 12, right: 14, bottom: 42, left: 42 };
  const chartW = width - margin.left - margin.right;
  const chartH = height - margin.top - margin.bottom;

  const maxVal = Math.max(1, ...history.map((h) => Math.max(h.correctPerMin || 0, h.incorrectPerMin || 0)));
  const niceMax = niceCeil(maxVal);

  const x = (i) => margin.left + (history.length <= 1 ? 0 : (i / (history.length - 1)) * chartW);
  const y = (value) => margin.top + chartH - (Math.min(value, niceMax) / niceMax) * chartH;

  const yTicks = 5;
  for (let i = 0; i <= yTicks; i += 1) {
    const value = (niceMax / yTicks) * i;
    const py = y(value);
    svg.append(
      createSvg('line', {
        x1: margin.left,
        y1: py,
        x2: width - margin.right,
        y2: py,
        stroke: 'var(--border)',
        'stroke-width': 1,
      })
    );
    svg.append(
      createSvg('text', {
        x: margin.left - 8,
        y: py + 4,
        'text-anchor': 'end',
        fill: 'var(--muted)',
        'font-size': 11,
      }, value.toFixed(1))
    );
  }

  const pointsCorrect = history.map((item, i) => `${x(i)},${y(item.correctPerMin || 0)}`).join(' ');
  const pointsIncorrect = history.map((item, i) => `${x(i)},${y(item.incorrectPerMin || 0)}`).join(' ');

  svg.append(createSvg('polyline', {
    points: pointsCorrect,
    fill: 'none',
    stroke: 'var(--correct)',
    'stroke-width': 2,
  }));
  svg.append(createSvg('polyline', {
    points: pointsIncorrect,
    fill: 'none',
    stroke: 'var(--incorrect)',
    'stroke-width': 2,
  }));

  history.forEach((item, i) => {
    const isLatest = i === history.length - 1;
    svg.append(createSvg('circle', {
      cx: x(i),
      cy: y(item.correctPerMin || 0),
      r: isLatest ? 4 : 3,
      fill: isLatest ? 'var(--correct)' : 'var(--panel)',
      stroke: 'var(--correct)',
      'stroke-width': 2,
    }));
    svg.append(createSvg('circle', {
      cx: x(i),
      cy: y(item.incorrectPerMin || 0),
      r: isLatest ? 4 : 3,
      fill: isLatest ? 'var(--incorrect)' : 'var(--panel)',
      stroke: 'var(--incorrect)',
      'stroke-width': 2,
    }));
  });

  renderDateLabels(svg, history, x, height, margin);
}

function renderDateLabels(svg, history, x, height, margin) {
  const maxLabels = 6;
  const wanted = new Set([0, history.length - 1]);
  const stride = Math.max(1, Math.ceil(history.length / maxLabels));
  for (let i = 0; i < history.length; i += stride) {
    wanted.add(i);
  }

  const sortedIndexes = [...wanted].sort((a, b) => a - b);
  sortedIndexes.forEach((i, idx) => {
    const date = new Date(history[i].date);
    const label = Number.isNaN(date.getTime()) ? String(history[i].date).slice(0, 10) : date.toLocaleDateString();
    const anchor = idx === 0 ? 'start' : idx === sortedIndexes.length - 1 ? 'end' : 'middle';
    const textX = idx === 0 ? margin.left : idx === sortedIndexes.length - 1 ? 700 - margin.right : x(i);
    svg.append(createSvg('text', {
      x: textX,
      y: height - 12,
      'text-anchor': anchor,
      fill: 'var(--muted)',
      'font-size': 11,
    }, label));
  });
}

function createSvg(tag, attrs, text = '') {
  const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
  Object.entries(attrs).forEach(([key, value]) => el.setAttribute(key, value));
  if (text) el.textContent = text;
  return el;
}

function ratePerMinute(count, elapsedSeconds) {
  if (!elapsedSeconds || elapsedSeconds <= 0) return 0;
  const perMinute = count / (elapsedSeconds / 60);
  return Math.round(perMinute * 10) / 10;
}

function showScreen(name) {
  if (name === 'setup') renderDecks();
  Object.entries(screens).forEach(([screenName, el]) => {
    el.classList.toggle('hidden', screenName !== name);
  });
}

function onPracticeHotkeys(event) {
  if (screens.practice.classList.contains('hidden') || !state.practice || state.practice.finished) return;

  const key = event.key;
  const code = event.code;

  if (code === 'Space') {
    event.preventDefault();
    flipCard();
    return;
  }

  if (key === 'ArrowRight' || key.toLowerCase() === 'y') {
    event.preventDefault();
    markAnswer(true);
    return;
  }

  if (key === 'ArrowLeft' || key.toLowerCase() === 'n') {
    event.preventDefault();
    markAnswer(false);
  }
}

function onPreviewHotkeys(event) {
  if (screens.preview.classList.contains('hidden') || !state.preview) return;

  const key = event.key;
  const code = event.code;

  if (state.preview.isEditing) {
    if (code === 'Escape') {
      event.preventDefault();
      closePreviewEditor();
    }
    return;
  }

  if (code === 'Space') {
    event.preventDefault();
    flipPreviewCard();
    return;
  }

  if (key === 'ArrowRight' || key.toLowerCase() === 'n') {
    event.preventDefault();
    changePreviewCard(1);
    return;
  }

  if (key === 'ArrowLeft' || key.toLowerCase() === 'p') {
    event.preventDefault();
    changePreviewCard(-1);
    return;
  }

  if (code === 'Enter' || code === 'NumpadEnter') {
    event.preventDefault();
    if (state.currentDeckName) startPractice(state.currentDeckName);
  }
}

function shuffle(items) {
  for (let i = items.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [items[i], items[j]] = [items[j], items[i]];
  }
  return items;
}

function shuffleDeck(cards, previousCard = null) {
  const shuffled = shuffle([...cards]);
  if (shuffled.length <= 1 || !previousCard) return shuffled;

  if (shuffled[0] === previousCard) {
    const randomIndex = 1 + Math.floor(Math.random() * (shuffled.length - 1));
    const [firstCard] = shuffled.splice(0, 1);
    shuffled.splice(randomIndex, 0, firstCard);
  }

  return shuffled;
}

function niceCeil(value) {
  if (value <= 1) return 1;
  const power = 10 ** Math.floor(Math.log10(value));
  const scaled = value / power;
  const step = scaled <= 1 ? 1 : scaled <= 2 ? 2 : scaled <= 5 ? 5 : 10;
  return step * power;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function formatCardText(text) {
  const rawText = String(text ?? '');
  const lines = rawText.replace(/\r\n?/gu, '\n').split('\n');
  return lines.map((line) => formatCardLine(line)).join('');
}

function formatCardLine(line) {
  const image = parseCardImage(line);
  if (image) {
    return `<img class="card-image" src="${escapeHtml(image.src)}" alt="${escapeHtml(image.alt)}" loading="lazy" decoding="async" />`;
  }
  if (!line.trim()) {
    return '<span class="card-text-line card-spacer" aria-hidden="true"></span>';
  }
  return `<span class="card-text-line">${formatHighlightedText(line)}</span>`;
}

function parseCardImage(line) {
  const rawLine = String(line ?? '').trim();
  if (!rawLine) return null;

  const markdownImage = parseMarkdownImage(rawLine);
  if (markdownImage) {
    return markdownImage;
  }

  const src = normalizeCardImageSource(rawLine);
  if (!src) return null;
  return {
    src,
    alt: getCardImageAltText(src),
  };
}

function parseMarkdownImage(line) {
  const markdownMatch = String(line ?? '').match(/^!\[(.*?)\]\((.+)\)$/u);
  if (!markdownMatch) return null;

  const src = normalizeCardImageSource(extractMarkdownImageSource(markdownMatch[2]), { allowExtensionless: true });
  if (!src) return null;
  return {
    src,
    alt: markdownMatch[1].trim() || 'Card image',
  };
}

function extractMarkdownImageSource(destination) {
  const trimmed = String(destination ?? '').trim();
  if (!trimmed) return '';

  if (trimmed.startsWith('<')) {
    const closingIndex = trimmed.indexOf('>');
    return closingIndex > 0 ? trimmed.slice(1, closingIndex).trim() : '';
  }

  const withoutTitle = trimmed.replace(/\s+(?:"[^"]*"|'[^']*'|\([^()]*\))\s*$/u, '');
  return withoutTitle.trim();
}

function normalizeCardImageSource(source, { allowExtensionless = false } = {}) {
  const value = String(source ?? '').trim();
  if (!value) return null;

  try {
    const resolvedUrl = new URL(value, appBaseDirUrl);
    const protocol = resolvedUrl.protocol.toLowerCase();
    const isRelativePath = isAppRelativeImagePath(value);
    if (!['http:', 'https:', 'data:'].includes(protocol) && !(protocol === 'file:' && isRelativePath)) {
      return null;
    }

    if (protocol === 'data:') {
      return /^data:image\//iu.test(value) ? value : null;
    }

    if (allowExtensionless || hasRecognizedImageFileName(resolvedUrl.pathname)) {
      return resolvedUrl.toString();
    }
    return null;
  } catch {
    return null;
  }
}

function isAppRelativeImagePath(source) {
  const value = String(source ?? '').trim();
  return !!value
    && !/^[a-z][a-z0-9+.-]*:/iu.test(value)
    && !value.startsWith('//')
    && !value.startsWith('/')
    && !value.startsWith('\\')
    && !/^[a-z]:[\\/]/iu.test(value);
}

function hasRecognizedImageFileName(pathname) {
  const fileName = String(pathname ?? '')
    .split('/')
    .filter(Boolean)
    .pop();
  return /\.(?:avif|bmp|gif|jpe?g|png|svg|webp)$/iu.test(fileName || '');
}

function getCardImageAltText(source) {
  const fallback = 'Card image';
  try {
    const url = new URL(String(source ?? '').trim(), appBaseDirUrl);
    const fileName = url.pathname.split('/').pop();
    return fileName ? decodeURIComponent(fileName) : fallback;
  } catch {
    return fallback;
  }
}

function formatHighlightedText(rawText) {
  let rendered = '';
  let lastIndex = 0;
  const highlightPattern = /(^|[^\p{L}\p{N}_])==(\S(?:[\s\S]*?\S)?)==(?=$|[^\p{L}\p{N}_])/gu;
  let match;

  while ((match = highlightPattern.exec(rawText)) !== null) {
    rendered += escapeHtml(rawText.slice(lastIndex, match.index));
    rendered += escapeHtml(match[1]);
    rendered += `<mark>${escapeHtml(match[2])}</mark>`;
    lastIndex = match.index + match[0].length;
  }

  rendered += escapeHtml(rawText.slice(lastIndex));
  return rendered;
}

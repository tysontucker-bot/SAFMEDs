const STORAGE_KEY = 'safmeds.v1';
const HISTORY_LIMIT = 20;
const ROUND_SECONDS = 60;

const state = {
  decks: loadState(),
  currentDeckName: null,
  practice: null,
};

const screens = {
  setup: document.getElementById('setup-screen'),
  practice: document.getElementById('practice-screen'),
  results: document.getElementById('results-screen'),
  progress: document.getElementById('progress-screen'),
};

const fileInput = document.getElementById('file-input');
const importMessage = document.getElementById('import-message');
const deckList = document.getElementById('deck-list');

const practiceTitle = document.getElementById('practice-title');
const timerEl = document.getElementById('timer');
const practiceStats = document.getElementById('practice-stats');
const cardFace = document.getElementById('card-face');
const markCorrectBtn = document.getElementById('mark-correct');
const markIncorrectBtn = document.getElementById('mark-incorrect');
const endEarlyBtn = document.getElementById('end-early');

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
renderDecks();
showScreen('setup');

function bindEvents() {
  fileInput.addEventListener('change', onImportFile);

  cardFace.addEventListener('click', flipCard);
  markCorrectBtn.addEventListener('click', () => markAnswer(true));
  markIncorrectBtn.addEventListener('click', () => markAnswer(false));
  endEarlyBtn.addEventListener('click', () => finishPractice(true));

  document.getElementById('practice-back').addEventListener('click', cancelPractice);
  document.getElementById('results-back').addEventListener('click', () => showScreen('setup'));
  document.getElementById('practice-again').addEventListener('click', () => {
    if (state.currentDeckName) startPractice(state.currentDeckName);
  });
  document.getElementById('progress-back').addEventListener('click', () => showScreen('setup'));
  document.getElementById('progress-practice').addEventListener('click', () => {
    if (state.currentDeckName) startPractice(state.currentDeckName);
  });

  window.addEventListener('keydown', onPracticeHotkeys);
}

function loadState() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.decks));
}

function onImportFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  setMessage('');
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const workbook = XLSX.read(reader.result, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!firstSheet) throw new Error('No worksheets found in file.');

      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, defval: '' });
      const cards = parseCards(rows);
      if (!cards.length) throw new Error('No valid term/definition rows were found.');

      const deckName = file.name.replace(/\.[^.]+$/, '') || 'Imported Deck';
      const previousHistory = state.decks[deckName]?.history || [];
      state.decks[deckName] = { name: deckName, cards, history: previousHistory };
      saveState();
      renderDecks();
      setMessage(`Saved "${deckName}" with ${cards.length} cards.`, 'success');
    } catch (error) {
      setMessage(`Import failed: ${error.message || 'Unable to read spreadsheet.'}`, 'error');
    } finally {
      fileInput.value = '';
    }
  };

  reader.onerror = () => {
    setMessage('Import failed: file could not be read.', 'error');
    fileInput.value = '';
  };

  reader.readAsArrayBuffer(file);
}

function parseCards(rows) {
  if (!Array.isArray(rows) || !rows.length) return [];

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

function setMessage(text, type = '') {
  importMessage.textContent = text;
  importMessage.className = `message ${type}`.trim();
}

function renderDecks() {
  const names = Object.keys(state.decks).sort((a, b) => a.localeCompare(b));
  if (!names.length) {
    deckList.innerHTML = '<p class="empty">No decks yet. Import a spreadsheet to begin.</p>';
    return;
  }

  deckList.innerHTML = '';
  for (const name of names) {
    const deck = state.decks[name];
    const row = document.createElement('div');
    row.className = 'deck-row';

    const meta = document.createElement('div');
    meta.className = 'deck-meta';
    meta.innerHTML = `<strong>${escapeHtml(deck.name)}</strong><span>${deck.cards.length} cards · ${deck.history?.length || 0} timings</span>`;

    const actions = document.createElement('div');
    actions.className = 'controls';

    const progressBtn = createButton('Progress', () => openProgress(name));
    const practiceBtn = createButton('Practice', () => startPractice(name));
    const deleteBtn = createButton('Delete', () => deleteDeck(name), 'delete');

    actions.append(progressBtn, practiceBtn, deleteBtn);
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

function startPractice(deckName) {
  const deck = state.decks[deckName];
  if (!deck || !deck.cards.length) return;

  if (state.practice?.timerId) clearInterval(state.practice.timerId);

  state.currentDeckName = deckName;
  const round = shuffle([...deck.cards]);
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
  cardFace.textContent = state.practice.showingDefinition ? currentCard.definition : currentCard.term;
}

function getCurrentCard() {
  if (!state.practice?.round.length) return null;
  return state.practice.round[state.practice.index];
}

function flipCard() {
  if (!state.practice || state.practice.finished) return;
  state.practice.showingDefinition = !state.practice.showingDefinition;
  updateCardFace();
}

function markAnswer(isCorrect) {
  if (!state.practice || state.practice.finished) return;
  if (isCorrect) state.practice.correct += 1;
  else state.practice.incorrect += 1;

  state.practice.seen += 1;
  state.practice.index += 1;
  state.practice.showingDefinition = false;

  if (state.practice.index >= state.practice.round.length) {
    state.practice.round = shuffle([...state.decks[state.currentDeckName].cards]);
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

function shuffle(items) {
  for (let i = items.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [items[i], items[j]] = [items[j], items[i]];
  }
  return items;
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

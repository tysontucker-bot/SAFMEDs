const STORAGE_KEY = 'safmeds-state-v1';
const HISTORY_LIMIT = 20;
const TIMING_SECONDS = 60;
const URGENT_THRESHOLD = 10;
const MAX_IMPORT_SIZE_BYTES = 5 * 1024 * 1024;

const state = {
  decks: {},
  view: 'decks',
  selectedDeck: null,
  notice: null,
  practice: null,
  lastResultDeck: null,
};

const app = document.getElementById('app');

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return;
    }
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object' && parsed.decks && typeof parsed.decks === 'object') {
      state.decks = parsed.decks;
    }
  } catch (error) {
    console.error('Failed to load stored SAFMEDs state', error);
  }
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ decks: state.decks }));
}

function setNotice(type, text) {
  state.notice = { type, text };
}

function clearNotice() {
  state.notice = null;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function getDeckArray() {
  return Object.values(state.decks).sort((left, right) => left.name.localeCompare(right.name));
}

function normalizeRows(rows) {
  if (!rows.length) {
    return [];
  }

  let startIndex = 0;
  const [firstCell = '', secondCell = ''] = rows[0] || [];
  if (
    String(firstCell).trim().toLowerCase() === 'term' &&
    String(secondCell).trim().toLowerCase() === 'definition'
  ) {
    startIndex = 1;
  }

  const cards = [];
  for (let index = startIndex; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const term = String(row[0] ?? '').trim();
    const definition = String(row[1] ?? '').trim();
    if (term && definition) {
      cards.push({ term, definition });
    }
  }
  return cards;
}

function saveDeck(fileName, cards) {
  const deckName = fileName.replace(/\.[^.]+$/, '') || 'Untitled deck';
  const existingHistory = state.decks[deckName]?.history || [];
  state.decks[deckName] = {
    name: deckName,
    cards,
    history: existingHistory.slice(0, HISTORY_LIMIT),
  };
  persistState();
  setNotice('success', `Saved ${cards.length} cards to “${deckName}”.`);
  render();
}

async function handleImport(event) {
  const [file] = event.target.files || [];
  event.target.value = '';
  if (!file) {
    return;
  }

  try {
    if (typeof XLSX === 'undefined') {
      throw new Error('The spreadsheet parser is not available.');
    }
    if (file.size > MAX_IMPORT_SIZE_BYTES) {
      throw new Error('The spreadsheet is too large. Please use a file smaller than 5 MB.');
    }
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    if (!firstSheetName || !worksheet) {
      throw new Error('The spreadsheet did not contain a readable worksheet.');
    }
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      blankrows: false,
      defval: '',
    });
    const cards = normalizeRows(rows);
    if (!cards.length) {
      throw new Error('The file did not contain any valid term/definition rows.');
    }
    saveDeck(file.name, cards);
  } catch (error) {
    console.error('Import failed', error);
    setNotice('error', error.message || 'The spreadsheet could not be read.');
    render();
  }
}

function deleteDeck(deckName) {
  const deck = state.decks[deckName];
  if (!deck) {
    return;
  }
  const confirmed = window.confirm(`Delete “${deckName}” and all of its history?`);
  if (!confirmed) {
    return;
  }
  if (state.practice?.deckName === deckName) {
    stopPracticeTimer();
    state.practice = null;
  }
  delete state.decks[deckName];
  if (state.selectedDeck === deckName || state.lastResultDeck === deckName) {
    state.selectedDeck = null;
    state.lastResultDeck = null;
    state.view = 'decks';
  }
  persistState();
  setNotice('success', `Deleted “${deckName}”.`);
  render();
}

function shuffleCards(cards) {
  const copy = cards.map((card) => ({ ...card }));
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [copy[index], copy[swapIndex]] = [copy[swapIndex], copy[index]];
  }
  return copy;
}

function stopPracticeTimer() {
  if (state.practice?.timerId) {
    window.clearInterval(state.practice.timerId);
    state.practice.timerId = null;
  }
}

function startPractice(deckName) {
  const deck = state.decks[deckName];
  if (!deck || !deck.cards.length) {
    setNotice('error', 'This deck has no cards to practice.');
    state.view = 'decks';
    render();
    return;
  }

  stopPracticeTimer();
  clearNotice();
  const startedAt = performance.now();
  state.selectedDeck = deckName;
  state.practice = {
    deckName,
    shuffledCards: shuffleCards(deck.cards),
    deckSize: deck.cards.length,
    currentIndex: 0,
    totalSeen: 0,
    correct: 0,
    incorrect: 0,
    showDefinition: false,
    startedAt,
    elapsedSeconds: 0,
    secondsLeft: TIMING_SECONDS,
    timerId: null,
    saved: false,
  };
  state.view = 'practice';
  state.practice.timerId = window.setInterval(updatePracticeTimer, 100);
  render();
}

function getCurrentCard() {
  const practice = state.practice;
  if (!practice) {
    return null;
  }
  return practice.shuffledCards[practice.currentIndex] || null;
}

function advanceCard() {
  const practice = state.practice;
  const deck = state.decks[practice.deckName];
  practice.currentIndex += 1;
  if (practice.currentIndex >= practice.shuffledCards.length) {
    practice.shuffledCards = shuffleCards(deck.cards);
    practice.currentIndex = 0;
  }
  practice.showDefinition = false;
}

function markAnswer(isCorrect) {
  const practice = state.practice;
  if (!practice) {
    return;
  }
  if (isCorrect) {
    practice.correct += 1;
  } else {
    practice.incorrect += 1;
  }
  practice.totalSeen += 1;
  advanceCard();
  render();
}

function roundRate(value, elapsedSeconds) {
  return Number((value / (elapsedSeconds / 60)).toFixed(1));
}

function finishPractice(endedEarly = false) {
  const practice = state.practice;
  if (!practice || practice.saved) {
    return;
  }

  stopPracticeTimer();
  const elapsedSeconds = endedEarly
    ? Math.max(0.1, (performance.now() - practice.startedAt) / 1000)
    : Math.max(0.1, Math.min(TIMING_SECONDS, practice.elapsedSeconds || TIMING_SECONDS));
  const entry = {
    date: new Date().toISOString(),
    correct: practice.correct,
    incorrect: practice.incorrect,
    correctPerMin: roundRate(practice.correct, elapsedSeconds),
    incorrectPerMin: roundRate(practice.incorrect, elapsedSeconds),
  };

  const deck = state.decks[practice.deckName];
  deck.history = [entry, ...(deck.history || [])].slice(0, HISTORY_LIMIT);
  practice.saved = true;
  state.lastResultDeck = practice.deckName;
  state.selectedDeck = practice.deckName;
  state.view = 'results';
  persistState();
  render();
}

function updatePracticeTimer() {
  const practice = state.practice;
  if (!practice) {
    return;
  }
  const elapsedSeconds = Math.min(TIMING_SECONDS, (performance.now() - practice.startedAt) / 1000);
  practice.elapsedSeconds = elapsedSeconds;
  practice.secondsLeft = Math.max(0, TIMING_SECONDS - elapsedSeconds);
  if (practice.secondsLeft <= 0) {
    finishPractice(false);
    return;
  }
  updatePracticeDisplay();
}

function openProgress(deckName) {
  state.selectedDeck = deckName;
  state.view = 'progress';
  clearNotice();
  render();
}

function backToDecks() {
  if (state.practice) {
    stopPracticeTimer();
    state.practice = null;
  }
  state.view = 'decks';
  render();
}

function formatDate(timestamp) {
  return new Date(timestamp).toLocaleString([], {
    dateStyle: 'medium',
    timeStyle: 'short',
  });
}

function buildChart(history) {
  if (history.length < 2) {
    return '<p class="chart-empty muted">Complete at least two timings to see the trend chart.</p>';
  }

  const ordered = [...history].reverse();
  const chartWidth = 760;
  const chartHeight = 280;
  const padding = { top: 20, right: 18, bottom: 54, left: 50 };
  const innerWidth = chartWidth - padding.left - padding.right;
  const innerHeight = chartHeight - padding.top - padding.bottom;
  const maxValue = Math.max(
    1,
    ...ordered.flatMap((entry) => [entry.correctPerMin, entry.incorrectPerMin])
  );
  const paddedMax = Math.max(5, Math.ceil(maxValue / 5) * 5);
  const gridlineCount = 5;
  const xStep = ordered.length > 1 ? innerWidth / (ordered.length - 1) : innerWidth;
  const xFor = (index) => padding.left + index * xStep;
  const yFor = (value) => padding.top + innerHeight - (value / paddedMax) * innerHeight;
  const pointString = (key) =>
    ordered
      .map((entry, index) => `${xFor(index)},${yFor(entry[key])}`)
      .join(' ');

  const labelIndexes = new Set([0, ordered.length - 1]);
  const every = Math.max(1, Math.ceil(ordered.length / 4));
  for (let index = 0; index < ordered.length; index += every) {
    labelIndexes.add(index);
  }

  const xLabels = ordered
    .map((entry, index) => {
      if (!labelIndexes.has(index)) {
        return '';
      }
      const anchor = index === 0 ? 'start' : index === ordered.length - 1 ? 'end' : 'middle';
      return `<text x="${xFor(index)}" y="${chartHeight - 18}" text-anchor="${anchor}" fill="currentColor" font-size="12" opacity="0.7">${escapeHtml(new Date(entry.date).toLocaleDateString())}</text>`;
    })
    .join('');

  const yGridlines = Array.from({ length: gridlineCount + 1 }, (_, index) => {
    const value = (paddedMax / gridlineCount) * index;
    const y = yFor(value);
    return `
      <line x1="${padding.left}" y1="${y}" x2="${chartWidth - padding.right}" y2="${y}" stroke="currentColor" opacity="0.12" />
      <text x="${padding.left - 10}" y="${y + 4}" text-anchor="end" fill="currentColor" font-size="12" opacity="0.7">${value.toFixed(0)}</text>
    `;
  }).join('');

  const renderPoints = (key, color) =>
    ordered
      .map((entry, index) => {
        const latest = index === ordered.length - 1;
        const fill = latest ? color : 'var(--panel)';
        const radius = latest ? 5 : 4;
        return `<circle cx="${xFor(index)}" cy="${yFor(entry[key])}" r="${radius}" fill="${fill}" stroke="${color}" stroke-width="2" />`;
      })
      .join('');

  return `
    <svg viewBox="0 0 ${chartWidth} ${chartHeight}" class="chart-svg" role="img" aria-label="Trend chart of correct and incorrect counts per minute over time">
      ${yGridlines}
      <line x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${chartWidth - padding.right}" y2="${padding.top + innerHeight}" stroke="currentColor" opacity="0.25" />
      <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}" stroke="currentColor" opacity="0.25" />
      <polyline fill="none" stroke="#2563eb" stroke-width="3" points="${pointString('correctPerMin')}" />
      <polyline fill="none" stroke="#dc2626" stroke-width="3" points="${pointString('incorrectPerMin')}" />
      ${renderPoints('correctPerMin', '#2563eb')}
      ${renderPoints('incorrectPerMin', '#dc2626')}
      ${xLabels}
    </svg>
  `;
}

function buildHistoryTable(history) {
  if (!history.length) {
    return '<p class="muted">No timings recorded yet.</p>';
  }

  const rows = history
    .map(
      (entry) => `
        <tr>
          <td>${escapeHtml(formatDate(entry.date))}</td>
          <td>${entry.correct}</td>
          <td>${entry.incorrect}</td>
          <td>${entry.correctPerMin.toFixed(1)}</td>
          <td>${entry.incorrectPerMin.toFixed(1)}</td>
        </tr>
      `
    )
    .join('');

  return `
    <div class="history-table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th scope="col">Date</th>
            <th scope="col">Correct</th>
            <th scope="col">Incorrect</th>
            <th scope="col">Correct/min</th>
            <th scope="col">Incorrect/min</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderSharedHistory(deck, options = {}) {
  const history = deck.history || [];
  return `
    <section class="panel stack">
      <div class="screen-header">
        <h3>Progress</h3>
        <p class="screen-subtitle">${options.subtitle || 'Track correct and incorrect rates across recent timings.'}</p>
      </div>
      <div class="chart-panel stack">
        <div class="legend" aria-label="Chart legend">
          <span class="legend-item"><span class="legend-swatch" style="background:#2563eb"></span>Correct/min</span>
          <span class="legend-item"><span class="legend-swatch" style="background:#dc2626"></span>Incorrect/min</span>
        </div>
        ${buildChart(history)}
      </div>
      <div class="stack">
        <h3>History</h3>
        ${buildHistoryTable(history)}
      </div>
    </section>
  `;
}

function renderDecksView() {
  const decks = getDeckArray();
  const notice = state.notice
    ? `<div class="message ${state.notice.type}">${escapeHtml(state.notice.text)}</div>`
    : '';

  const deckList = decks.length
    ? `<div class="deck-list">${decks
        .map(
          (deck) => `
            <article class="deck-item">
              <div class="deck-title-row">
                <h3>${escapeHtml(deck.name)}</h3>
                <span class="pill">${deck.cards.length} cards</span>
              </div>
              <div class="deck-meta muted">${deck.history?.length || 0} recorded timings</div>
              <div class="deck-actions">
                <button class="btn primary" data-action="practice" data-deck="${escapeHtml(deck.name)}">Practice</button>
                <button class="btn secondary" data-action="progress" data-deck="${escapeHtml(deck.name)}">Progress</button>
                <button class="btn danger" data-action="delete" data-deck="${escapeHtml(deck.name)}">Delete</button>
              </div>
            </article>
          `
        )
        .join('')}</div>`
    : `
      <div class="empty-state">
        <h3>No decks yet</h3>
        <p class="muted">Import a spreadsheet with Term and Definition columns to create your first SAFMEDs deck.</p>
      </div>
    `;

  app.innerHTML = `
    <section class="panel stack">
      <div class="screen-header">
        <h2>Deck setup</h2>
        <p class="screen-subtitle">Import .xlsx or .xls decks, then launch a timing or review each deck's progress.</p>
      </div>
      ${notice}
      <div class="import-row">
        <label class="file-input" for="deck-file">
          <span>Import spreadsheet (.xlsx, .xls)</span>
          <input id="deck-file" type="file" accept=".xlsx,.xls" />
        </label>
      </div>
    </section>
    <section class="panel stack">
      <div class="screen-header">
        <h2>Saved decks</h2>
        <p class="screen-subtitle">Deck data and history stay on this device only.</p>
      </div>
      ${deckList}
    </section>
  `;

  document.getElementById('deck-file')?.addEventListener('change', handleImport);
  bindDeckButtons();
}

function bindDeckButtons() {
  app.querySelectorAll('[data-action="practice"]').forEach((button) => {
    button.addEventListener('click', () => startPractice(button.dataset.deck));
  });
  app.querySelectorAll('[data-action="progress"]').forEach((button) => {
    button.addEventListener('click', () => openProgress(button.dataset.deck));
  });
  app.querySelectorAll('[data-action="delete"]').forEach((button) => {
    button.addEventListener('click', () => deleteDeck(button.dataset.deck));
  });
}

function renderPracticeView() {
  const practice = state.practice;
  const deck = practice ? state.decks[practice.deckName] : null;
  const card = getCurrentCard();
  if (!practice || !deck || !card) {
    backToDecks();
    return;
  }

  const secondsLeft = Math.ceil(practice.secondsLeft);
  const urgencyClass = secondsLeft <= URGENT_THRESHOLD ? 'timer-urgent' : '';
  const displayedText = practice.showDefinition ? card.definition : card.term;
  const displayedSide = practice.showDefinition ? 'Definition' : 'Term';

  app.innerHTML = `
    <section class="panel stack">
      <div class="screen-header">
        <h2>${escapeHtml(deck.name)}</h2>
        <p class="screen-subtitle">Keyboard: Space flips the card, Y / → marks correct, N / ← marks incorrect.</p>
      </div>
      <div class="practice-summary">
        <div class="practice-stats">
          <div class="stat-card">
            <p class="stat-label">Time left</p>
            <p class="stat-value ${urgencyClass}" id="practice-time-left">${secondsLeft}s</p>
          </div>
          <div class="stat-card">
            <p class="stat-label">Progress</p>
            <p class="stat-value" id="practice-progress">${practice.currentIndex + 1} / ${practice.deckSize}</p>
          </div>
          <div class="stat-card">
            <p class="stat-label">Correct</p>
            <p class="stat-value" id="practice-correct">${practice.correct}</p>
          </div>
          <div class="stat-card">
            <p class="stat-label">Incorrect</p>
            <p class="stat-value" id="practice-incorrect">${practice.incorrect}</p>
          </div>
        </div>
        <div class="practice-card">
          <button id="flip-card" class="flashcard-button" type="button">
            <span class="card-side" id="practice-card-side">${displayedSide}</span>
            <span class="card-face" id="practice-card-face">${escapeHtml(displayedText)}</span>
            <span class="shortcut-hint">Tap or press Space to flip.</span>
          </button>
        </div>
        <div class="button-row">
          <button class="btn success" id="mark-correct" type="button">Correct</button>
          <button class="btn danger" id="mark-incorrect" type="button">Incorrect</button>
          <button class="btn secondary" id="end-early" type="button">End Timing Early</button>
          <button class="btn secondary" id="back-decks" type="button">Back to Decks</button>
        </div>
      </div>
    </section>
  `;

  document.getElementById('flip-card')?.addEventListener('click', toggleCardFace);
  document.getElementById('mark-correct')?.addEventListener('click', () => markAnswer(true));
  document.getElementById('mark-incorrect')?.addEventListener('click', () => markAnswer(false));
  document.getElementById('end-early')?.addEventListener('click', () => finishPractice(true));
  document.getElementById('back-decks')?.addEventListener('click', backToDecks);
}

function updatePracticeDisplay() {
  const practice = state.practice;
  const card = getCurrentCard();
  if (!practice || !card || state.view !== 'practice') {
    return;
  }

  const secondsLeft = Math.ceil(practice.secondsLeft);
  const timeLeft = document.getElementById('practice-time-left');
  const progress = document.getElementById('practice-progress');
  const correct = document.getElementById('practice-correct');
  const incorrect = document.getElementById('practice-incorrect');
  const side = document.getElementById('practice-card-side');
  const face = document.getElementById('practice-card-face');

  if (!timeLeft || !progress || !correct || !incorrect || !side || !face) {
    render();
    return;
  }

  timeLeft.textContent = `${secondsLeft}s`;
  timeLeft.classList.toggle('timer-urgent', secondsLeft <= URGENT_THRESHOLD);
  progress.textContent = `${practice.currentIndex + 1} / ${practice.deckSize}`;
  correct.textContent = String(practice.correct);
  incorrect.textContent = String(practice.incorrect);
  side.textContent = practice.showDefinition ? 'Definition' : 'Term';
  face.textContent = practice.showDefinition ? card.definition : card.term;
}

function toggleCardFace() {
  if (!state.practice) {
    return;
  }
  state.practice.showDefinition = !state.practice.showDefinition;
  render();
}

function renderResultsView() {
  const deck = state.decks[state.lastResultDeck || state.selectedDeck];
  if (!deck) {
    state.view = 'decks';
    render();
    return;
  }
  const latest = deck.history?.[0];
  if (!latest) {
    state.view = 'progress';
    render();
    return;
  }

  app.innerHTML = `
    <section class="panel stack">
      <div class="screen-header">
        <h2>Timing results</h2>
        <p class="screen-subtitle">${escapeHtml(deck.name)}</p>
      </div>
      <div class="stats-grid">
        <div class="stat-card">
          <p class="stat-label">Correct/min</p>
          <p class="stat-value">${latest.correctPerMin.toFixed(1)}</p>
        </div>
        <div class="stat-card">
          <p class="stat-label">Correct</p>
          <p class="stat-value">${latest.correct}</p>
        </div>
        <div class="stat-card">
          <p class="stat-label">Incorrect</p>
          <p class="stat-value">${latest.incorrect}</p>
        </div>
        <div class="stat-card">
          <p class="stat-label">Cards seen</p>
          <p class="stat-value">${latest.correct + latest.incorrect}</p>
        </div>
      </div>
      <div class="button-row">
        <button class="btn secondary" id="results-back">Back to Decks</button>
        <button class="btn primary" id="practice-again">Practice Again</button>
      </div>
    </section>
    ${renderSharedHistory(deck, { subtitle: 'See progress immediately after each completed timing.' })}
  `;

  document.getElementById('results-back')?.addEventListener('click', () => {
    state.practice = null;
    backToDecks();
  });
  document.getElementById('practice-again')?.addEventListener('click', () => startPractice(deck.name));
}

function renderProgressView() {
  const deck = state.decks[state.selectedDeck];
  if (!deck) {
    state.view = 'decks';
    render();
    return;
  }

  app.innerHTML = `
    <section class="panel stack">
      <div class="screen-header">
        <h2>${escapeHtml(deck.name)} progress</h2>
        <p class="screen-subtitle">Review timing history and fluency trends for this deck.</p>
      </div>
      <div class="history-actions">
        <button class="btn secondary" id="progress-back">Back to Decks</button>
        <button class="btn primary" id="progress-practice">Practice This Deck</button>
      </div>
    </section>
    ${renderSharedHistory(deck)}
  `;

  document.getElementById('progress-back')?.addEventListener('click', backToDecks);
  document.getElementById('progress-practice')?.addEventListener('click', () => startPractice(deck.name));
}

function handlePracticeKeyboard(event) {
  if (state.view !== 'practice' || !state.practice) {
    return;
  }

  if (event.target instanceof HTMLInputElement) {
    return;
  }

  const key = event.key.toLowerCase();
  if (event.code === 'Space') {
    event.preventDefault();
    toggleCardFace();
    return;
  }
  if (key === 'y' || event.key === 'ArrowRight') {
    event.preventDefault();
    markAnswer(true);
    return;
  }
  if (key === 'n' || event.key === 'ArrowLeft') {
    event.preventDefault();
    markAnswer(false);
  }
}

function render() {
  switch (state.view) {
    case 'practice':
      renderPracticeView();
      break;
    case 'results':
      renderResultsView();
      break;
    case 'progress':
      renderProgressView();
      break;
    default:
      renderDecksView();
      break;
  }
}

window.addEventListener('keydown', handlePracticeKeyboard);
loadState();
render();

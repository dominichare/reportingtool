import {
  getSettings, saveSettings, sttReady, llmReady,
  getPeople, savePeople, rememberPerson, findPerson, topPeople,
} from './settings.js';
import { transcribe, writeNote, checkKey, STT_PROVIDERS, LLM_PROVIDERS } from './providers.js';
import { HELP_SECTIONS, HELP_TOPICS, itRequestText } from './help.js';
import {
  buildPrompt, noteDocx, transcriptDocx, noteFileName, transcriptFileName,
  guessDate, longDate, fmtTime, namedTranscript, participantNames,
  emailSubject, emailHtml, emailText, safeName,
} from './note.js';
import * as MS from './graph.js';

const DB = window.MinutesDB;

// =====================================================================
// Helpers
// =====================================================================

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const ME = '__me__';
const IGNORE = '__ignore__';

const I = {
  back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5l-7 7 7 7"/></svg>',
  chev: '<svg viewBox="0 0 12 20" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 2l7 8-7 8"/></svg>',
  gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3.2"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  wave: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M4 10v4M8 6v12M12 3v18M16 7v10M20 10v4"/></svg>',
  play: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5.5v13a1 1 0 0 0 1.5.86l10.6-6.5a1 1 0 0 0 0-1.72L9.5 4.64A1 1 0 0 0 8 5.5z"/></svg>',
  pause: '<svg viewBox="0 0 24 24" fill="currentColor"><rect x="6" y="5" width="4" height="14" rx="1.2"/><rect x="14" y="5" width="4" height="14" rx="1.2"/></svg>',
  check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
  up: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5M5.5 11.5L12 5l6.5 6.5"/></svg>',
  folder: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 6.5A2.5 2.5 0 0 1 5.5 4h3.6c.6 0 1.2.2 1.6.6L12 6h6.5A2.5 2.5 0 0 1 21 8.5v9a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5z"/></svg>',
  person: '<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-6.5 8-6.5s8 2.5 8 6.5z"/></svg>',
  mail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M3.5 6.5L12 13l8.5-6.5"/></svg>',
  doc: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/></svg>',
  cloud: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M7 19a5 5 0 0 1-.6-10A6.5 6.5 0 0 1 19 9.5a4.5 4.5 0 0 1-.5 9.5z"/></svg>',
  share: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7.5 7.5L12 3l4.5 4.5"/><path d="M5 12v6.5A2.5 2.5 0 0 0 7.5 21h9a2.5 2.5 0 0 0 2.5-2.5V12"/></svg>',
  copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="8" y="8" width="12" height="12" rx="2.5"/><path d="M16 8V6.5A2.5 2.5 0 0 0 13.5 4h-7A2.5 2.5 0 0 0 4 6.5v7A2.5 2.5 0 0 0 6.5 16H8"/></svg>',
  warn: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M10.3 3.9a2 2 0 0 1 3.4 0l8 13.6A2 2 0 0 1 20 20.5H4a2 2 0 0 1-1.7-3zM12 9a1 1 0 0 0-1 1v4a1 1 0 1 0 2 0v-4a1 1 0 0 0-1-1zm0 7.2a1.2 1.2 0 1 0 0 2.4 1.2 1.2 0 0 0 0-2.4z"/></svg>',
  info: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2a10 10 0 1 1 0 20 10 10 0 0 1 0-20zm0 8.2a1 1 0 0 0-1 1V17a1 1 0 1 0 2 0v-5.8a1 1 0 0 0-1-1zM12 6.5a1.3 1.3 0 1 0 0 2.6 1.3 1.3 0 0 0 0-2.6z"/></svg>',
  okc: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2a10 10 0 1 1 0 20 10 10 0 0 1 0-20zm4.3 6.3a1 1 0 0 0-1.4 0l-4.2 4.6-1.6-1.6a1 1 0 1 0-1.4 1.4l2.3 2.3a1 1 0 0 0 1.5 0l4.9-5.3a1 1 0 0 0-.1-1.4z"/></svg>',
  search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4.5 4.5"/></svg>',
  sparkle: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2.5l1.9 5.6 5.6 1.9-5.6 1.9L12 17.5l-1.9-5.6L4.5 10l5.6-1.9zM18.5 15l.9 2.6 2.6.9-2.6.9-.9 2.6-.9-2.6-2.6-.9 2.6-.9z"/></svg>',
  trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16M9 7V4.5h6V7M6.5 7l1 13h9l1-13"/></svg>',
};

function toast(msg) {
  const t = $('#toast');
  t.innerHTML = `<div>${esc(msg)}</div>`;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.innerHTML = ''; }, 3300);
}

function fmtSize(bytes) {
  return bytes > 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}
function fmtDur(ms) {
  const m = Math.round(ms / 60000);
  return m < 1 ? '<1 min' : `${m} min`;
}
const displayPath = p => String(p || '').split('/').filter(Boolean).join(' › ');

async function saveJob(job) {
  job.updatedAt = Date.now();
  await DB.put('jobs', job);
}

let wakeLock = null;
async function keepAwake(on) {
  try {
    if (on && 'wakeLock' in navigator && !wakeLock) wakeLock = await navigator.wakeLock.request('screen');
    if (!on && wakeLock) { await wakeLock.release(); wakeLock = null; }
  } catch {}
}

// =====================================================================
// Navigation (stack with iOS-style push / pop, Android back button aware)
// =====================================================================

const stack = []; // [{ render, el }]
let sheetOpen = null;

function mount(render, anim) {
  const el = document.createElement('section');
  el.className = `screen ${anim || ''}`;
  $('#app').appendChild(el);
  render(el);
  const content = $('.content', el);
  if (content) content.addEventListener('scroll', () => el.classList.toggle('scrolled', content.scrollTop > 28), { passive: true });
  return el;
}

function push(render) {
  const prev = stack[stack.length - 1];
  const el = mount(render, 'enter');
  stack.push({ render, el });
  history.pushState({ depth: stack.length }, '');
  if (prev) setTimeout(() => prev.el.classList.add('hidden'), 380);
}

function replace(render, anim = 'fade') {
  const old = stack.pop();
  const el = mount(render, anim);
  stack.push({ render, el });
  if (old) setTimeout(() => old.el.remove(), 260);
}

function resetTo(render) {
  stack.splice(0).forEach(s => s.el.remove());
  const el = mount(render, 'fade');
  stack.push({ render, el });
}

function back() { history.back(); }

window.addEventListener('popstate', e => {
  if (sheetOpen) sheetOpen.close(true);
  const depth = Math.max(1, e.state?.depth ?? 1);
  if (stack.length <= depth) return;
  while (stack.length > depth) stack.pop().el.remove();
  // Re-render the screen we are returning to so it reflects any changes.
  const top = stack[stack.length - 1];
  const fresh = mount(top.render, 'enter-back');
  top.el.replaceWith(fresh);
  top.el = fresh;
});

// Jump back to the first screen (used by "Done").
function backToHome() {
  if (sheetOpen) sheetOpen.close(true);
  const n = stack.length - 1;
  if (n > 0) history.go(-n);
}

function navBar({ title = '', back: showBack = false, backLabel = 'Back', right = '' }) {
  return `<header class="nav">
    <div class="nav-left">${showBack ? `<button class="nav-btn back" data-back>${I.back}<span>${esc(backLabel)}</span></button>` : ''}</div>
    <div class="nav-title">${esc(title)}</div>
    <div class="nav-right">${right}</div>
  </header>`;
}

function bindBack(el) {
  $$('[data-back]', el).forEach(b => b.addEventListener('click', back));
}

// ---------- Sheets ----------
function openSheet({ title, left = 'Cancel', right = '', onRight, build }) {
  const backdrop = document.createElement('div');
  backdrop.className = 'sheet-backdrop';
  const sheet = document.createElement('div');
  sheet.className = 'sheet';
  sheet.setAttribute('role', 'dialog');
  sheet.innerHTML = `<div class="sheet-grabber"></div>
    <div class="sheet-head">
      <div><button class="nav-btn" data-cancel>${esc(left)}</button></div>
      <div class="t">${esc(title)}</div>
      <div style="justify-self:end">${right ? `<button class="nav-btn bold" data-right>${esc(right)}</button>` : ''}</div>
    </div>
    <div class="sheet-body"></div>`;
  document.body.append(backdrop, sheet);
  history.pushState({ depth: stack.length, sheet: true }, '');

  const api = {
    el: sheet,
    body: $('.sheet-body', sheet),
    close(fromPop = false) {
      if (!sheetOpen) return;
      sheetOpen = null;
      sheet.classList.add('closing');
      backdrop.classList.add('closing');
      setTimeout(() => { sheet.remove(); backdrop.remove(); }, 240);
      if (!fromPop) history.back();
    },
  };
  sheetOpen = api;
  backdrop.addEventListener('click', () => api.close());
  $('[data-cancel]', sheet).addEventListener('click', () => api.close());
  if (right) $('[data-right]', sheet).addEventListener('click', () => onRight?.(api));
  build(api);
  return api;
}

function promptSheet({ title, value = '', placeholder = '', type = 'text', onSave }) {
  openSheet({
    title, right: 'Done',
    onRight: api => { onSave($('input', api.body).value.trim()); api.close(); },
    build: api => {
      api.body.innerHTML = `<div class="group mt"><div class="row"><input class="inline" style="text-align:left;color:var(--label)" type="${type}" value="${esc(value)}" placeholder="${esc(placeholder)}"></div></div>`;
      const inp = $('input', api.body);
      setTimeout(() => { inp.focus(); inp.select(); }, 350);
      inp.addEventListener('keydown', e => { if (e.key === 'Enter') { onSave(inp.value.trim()); api.close(); } });
    },
  });
}

// =====================================================================
// Home
// =====================================================================

function homeScreen(el) {
  const s = getSettings();
  const ready = sttReady(s) && llmReady(s) && s.myName;
  el.innerHTML = `
    ${navBar({ title: 'Minutes', right: `<button class="nav-btn" data-help>Help</button><button class="nav-btn" data-settings aria-label="Settings">${I.gear}</button>` })}
    <main class="content">
      <h1 class="large-title">Minutes</h1>
      ${ready ? '' : `<button class="notice warn" data-setup>${I.warn}<div><div class="t">Finish setting up</div><div class="s">Add your name and the AI services for transcription and note writing. Tap for a step-by-step guide.</div></div></button>`}
      <div class="hero">
        <div class="hero-icon">${I.wave}</div>
        <h2>New meeting note</h2>
        <p>In Easy Recorder, tap <b>Share</b> on a recording and choose <b>Minutes</b>. Or pick one here.</p>
        <button class="btn btn-primary" data-pick>${I.plus}<span>Choose recording</span></button>
        <input type="file" accept="audio/*,.m4a,.mp3,.wav,.aac" class="sr" data-file>
      </div>
      <div data-lists></div>
    </main>`;
  $('[data-settings]', el).addEventListener('click', () => push(settingsScreen));
  $('[data-setup]', el)?.addEventListener('click', () => push(e => helpTopicScreen(e, 'start')));
  $('[data-help]', el).addEventListener('click', () => push(helpScreen));
  const file = $('[data-file]', el);
  $('[data-pick]', el).addEventListener('click', () => file.click());
  file.addEventListener('change', () => { if (file.files[0]) startJob(file.files[0]); });
  renderJobLists($('[data-lists]', el));
}

async function renderJobLists(box) {
  const jobs = (await DB.all('jobs')).sort((a, b) => b.createdAt - a.createdAt);
  const active = jobs.filter(j => j.status !== 'complete');
  const done = jobs.filter(j => j.status === 'complete').slice(0, 25);
  const statusText = { new: 'Waiting to transcribe', transcribing: 'Transcribing', transcribed: 'Name the speakers', writing: 'Writing note', noted: 'Ready to review', saving: 'Saving' };
  let html = '';
  if (active.length) {
    html += `<div class="group-header">In progress</div><div class="group">${active.map(j => `
      <button class="row has-icon" data-job="${j.id}">
        <div class="row-icon" style="background:var(--orange)">${I.wave}</div>
        <div class="row-main"><div class="row-title">${esc(j.note?.title || j.fileName)}</div><div class="row-sub">${esc(statusText[j.status] || j.status)} · ${esc(longDate(j.meetingDate))}</div></div>
        <span class="chev">${I.chev}</span>
      </button>`).join('')}</div>`;
  }
  html += `<div class="group-header">Recent notes</div>`;
  html += done.length ? `<div class="group">${done.map(j => `
      <button class="row has-icon" data-job="${j.id}">
        <div class="row-icon" style="background:var(--accent)">${I.doc}</div>
        <div class="row-main"><div class="row-title">${esc(j.note?.title || j.fileName)}</div><div class="row-sub">${esc(longDate(j.meetingDate))}${j.save?.folder ? ` · ${esc(displayPath(j.save.folder))}` : ''}</div></div>
        <span class="chev">${I.chev}</span>
      </button>`).join('')}</div>`
    : `<div class="group"><div class="empty">Your meeting notes will appear here.</div></div>`;
  box.innerHTML = html;
  $$('[data-job]', box).forEach(b => b.addEventListener('click', async () => openJob(await DB.get('jobs', b.dataset.job))));
}

function openJob(job) {
  if (!job) return;
  const map = {
    new: processingScreen, transcribing: processingScreen, transcribed: speakersScreen,
    writing: speakersScreen, noted: reviewScreen, saving: saveScreen, complete: doneScreen,
  };
  push(el => map[job.status](el, job));
}

async function startJob(file) {
  const s = getSettings();
  if (!sttReady(s) || !llmReady(s)) {
    toast('Add your API keys in Settings first.');
    push(settingsScreen);
    return;
  }
  const job = {
    id: uid(),
    fileName: file.name || 'Recording.m4a',
    fileSize: file.size,
    lastModified: file.lastModified,
    meetingDate: guessDate(file.name, file.lastModified),
    createdAt: Date.now(),
    status: 'new',
    utterances: null,
    speakerNames: {},
    note: null,
  };
  await DB.put('audio', file, job.id);
  await saveJob(job);
  push(el => processingScreen(el, job));
}

// =====================================================================
// 1. Transcribing
// =====================================================================

function processingScreen(el, job) {
  el.innerHTML = `
    ${navBar({ title: 'Transcribing', back: true, backLabel: 'Minutes' })}
    <main class="content">
      <h1 class="large-title">Transcribing</h1>
      <p class="lede">${esc(job.fileName)} · ${fmtSize(job.fileSize)}</p>
      <div data-body></div>
    </main>`;
  bindBack(el);
  const body = $('[data-body]', el);

  const show = (state, msg = '') => {
    if (state === 'error') {
      body.innerHTML = `
        <div class="notice err">${I.warn}<div><div class="t">Something went wrong</div><div class="err-text">${esc(msg)}</div></div></div>
        <button class="btn btn-primary mt-l" data-retry>Try again</button>
        <button class="btn btn-plain mt" data-settings>Open Settings</button>`;
      $('[data-retry]', body).addEventListener('click', run);
      $('[data-settings]', body).addEventListener('click', () => push(settingsScreen));
      return;
    }
    body.innerHTML = `
      <div class="big-status"><div class="pulse">${I.wave}</div><div class="t" data-msg>${esc(msg || 'Starting…')}</div>
      <div class="s">Using ${esc(STT_PROVIDERS[getSettings().sttProvider].label)}. An hour-long meeting usually takes 1–3 minutes. Keep this screen open.</div></div>`;
  };

  async function run() {
    show('working', 'Preparing…');
    keepAwake(true);
    try {
      const blob = await DB.get('audio', job.id);
      if (!blob) throw new Error('The recording is no longer on this phone. Share it to Minutes again.');
      const file = blob instanceof File ? blob : new File([blob], job.fileName, { type: blob.type || 'audio/mp4' });
      job.status = 'transcribing';
      await saveJob(job);
      const r = await transcribe(file, msg => { const m = $('[data-msg]', body); if (m) m.textContent = msg; });
      job.utterances = r.utterances;
      job.durationMs = r.durationMs || (r.utterances.at(-1)?.end ?? 0);
      job.speakerNames = autoNameSpeakers(job);
      job.status = 'transcribed';
      await saveJob(job);
      keepAwake(false);
      if (el.isConnected) replace(e => speakersScreen(e, job), 'enter');
    } catch (e) {
      keepAwake(false);
      job.status = 'new';
      await saveJob(job);
      show('error', e.message || String(e));
    }
  }
  run();
}

function speakerStats(job) {
  const stats = {};
  for (const u of job.utterances) {
    const st = stats[u.speaker] ||= { speaker: u.speaker, ms: 0, turns: 0, first: u.start, samples: [] };
    st.ms += u.end - u.start;
    st.turns++;
    st.samples.push(u);
  }
  const total = Object.values(stats).reduce((n, s) => n + s.ms, 0) || 1;
  return Object.values(stats)
    .map(s => {
      // Best sample: a clear, longer turn (ideally 6–20 seconds).
      const scored = s.samples.map(u => {
        const d = u.end - u.start;
        return { u, score: (d >= 6000 && d <= 20000 ? 2 : 0) + Math.min(d, 20000) / 20000 + Math.min(u.text.length, 300) / 300 };
      }).sort((a, b) => b.score - a.score);
      return { ...s, share: s.ms / total, sample: scored[0].u, alt: scored.slice(1, 3).map(x => x.u) };
    })
    .sort((a, b) => b.ms - a.ms);
}

function autoNameSpeakers(job) {
  const names = {};
  const stats = speakerStats(job);
  if (stats.length) names[stats[0].speaker] = ME; // whoever talks most is assumed to be you
  return names;
}

// =====================================================================
// 2. Speakers
// =====================================================================

function speakersScreen(el, job) {
  const s = getSettings();
  const stats = speakerStats(job);
  const audioBlobP = DB.get('audio', job.id);
  let audio = null;
  let playingSpeaker = null;
  let stopAt = 0;

  const nameOf = sp => {
    const v = job.speakerNames[sp];
    return v === ME ? (s.myName || 'Me') : v;
  };

  el.innerHTML = `
    ${navBar({ title: 'Speakers', back: true, backLabel: 'Minutes' })}
    <main class="content with-toolbar">
      <h1 class="large-title">Who’s speaking?</h1>
      <p class="lede">${stats.length} voice${stats.length === 1 ? '' : 's'} found in ${fmtDur(job.durationMs)}. Name each one once. Tap ▶ to listen.</p>
      <div data-list></div>
    </main>
    <footer class="toolbar"><button class="btn btn-primary" data-next>${I.sparkle}<span>Write meeting note</span></button></footer>`;
  bindBack(el);
  const list = $('[data-list]', el);

  function render() {
    const used = new Set(Object.values(job.speakerNames).map(v => (v === ME ? s.myName : v)).filter(Boolean));
    list.innerHTML = stats.map((st, i) => {
      const v = job.speakerNames[st.speaker];
      const isMe = v === ME;
      const ignored = v === IGNORE;
      const value = isMe ? (s.myName || '') : (ignored ? '' : (v || ''));
      const suggestions = topPeople(8).filter(p => !used.has(p.name)).slice(0, 5);
      return `<div class="speaker" data-sp="${esc(st.speaker)}">
        <div class="speaker-top">
          <button class="play${playingSpeaker === st.speaker ? ' playing' : ''}" data-play aria-label="Play sample">${playingSpeaker === st.speaker ? I.pause : I.play}</button>
          <div style="flex:1;min-width:0">
            <div class="speaker-label">Voice ${i + 1} ${isMe ? '<span class="badge">Me</span>' : ''}${ignored ? '<span class="badge grey">Ignored</span>' : ''}</div>
            <div class="speaker-meta">${fmtDur(st.ms)} · ${Math.round(st.share * 100)}% of talking · ${st.turns} turn${st.turns === 1 ? '' : 's'}</div>
          </div>
        </div>
        <div class="quote">“${esc(st.sample.text)}”</div>
        ${ignored ? '' : `<div class="field">${I.person}<input data-name placeholder="${isMe && !s.myName ? 'Your name (set it in Settings)' : 'Name'}" value="${esc(value)}" autocomplete="off" autocapitalize="words" enterkeyhint="done"></div>`}
        <div class="chips">
          ${!isMe ? `<button class="chip ghost" data-me><span>It’s me</span></button>` : ''}
          ${!ignored && !isMe ? suggestions.map(p => `<button class="chip" data-pick="${esc(p.name)}"><span>${esc(p.name)}</span></button>`).join('') : ''}
          <button class="chip ghost" data-ignore><span>${ignored ? 'Include this voice' : 'Ignore (noise, TV…)'}</span></button>
        </div>
      </div>`;
    }).join('');

    $$('.speaker', list).forEach(card => {
      const sp = card.dataset.sp;
      $('[data-play]', card).addEventListener('click', () => togglePlay(sp));
      const inp = $('[data-name]', card);
      if (inp) attachPeopleTypeahead(inp, {
        onInput: v => {
          if (job.speakerNames[sp] === ME && v !== s.myName) job.speakerNames[sp] = v;
          else if (job.speakerNames[sp] !== ME) job.speakerNames[sp] = v;
          saveJob(job); updateNext();
        },
        onPick: p => {
          job.speakerNames[sp] = p.name;
          if (p.email) rememberPerson(p.name, p.email);
          saveJob(job); render();
        },
      });
      $('[data-me]', card)?.addEventListener('click', () => {
        for (const k of Object.keys(job.speakerNames)) if (job.speakerNames[k] === ME) delete job.speakerNames[k];
        job.speakerNames[sp] = ME; saveJob(job); render();
      });
      $$('[data-pick]', card).forEach(b => b.addEventListener('click', () => {
        job.speakerNames[sp] = b.dataset.pick; saveJob(job); render();
      }));
      $('[data-ignore]', card).addEventListener('click', () => {
        if (job.speakerNames[sp] === IGNORE) delete job.speakerNames[sp]; else job.speakerNames[sp] = IGNORE;
        saveJob(job); render();
      });
    });
    updateNext();
  }

  function updateNext() {
    const ok = stats.every(st => {
      const v = job.speakerNames[st.speaker];
      return v === IGNORE || (v === ME ? !!s.myName : !!(v && v.trim()));
    }) && stats.some(st => job.speakerNames[st.speaker] !== IGNORE);
    $('[data-next]', el).disabled = !ok;
  }

  async function togglePlay(sp) {
    const st = stats.find(x => x.speaker === sp);
    if (!audio) {
      const blob = await audioBlobP;
      if (!blob) { toast('The recording is no longer on this phone.'); return; }
      audio = new Audio(URL.createObjectURL(blob));
      audio.addEventListener('timeupdate', () => { if (audio.currentTime * 1000 >= stopAt) stop(); });
      audio.addEventListener('ended', stop);
    }
    if (playingSpeaker === sp) { stop(); return; }
    playingSpeaker = sp;
    // cycle through a few samples on repeated listens
    st._i = ((st._i ?? -1) + 1) % (1 + st.alt.length);
    const u = st._i === 0 ? st.sample : st.alt[st._i - 1];
    audio.currentTime = u.start / 1000;
    stopAt = Math.min(u.end, u.start + 15000);
    try { await audio.play(); } catch { toast('Couldn’t play the recording.'); }
    render();
  }
  function stop() {
    audio?.pause();
    playingSpeaker = null;
    $$('.play', list).forEach(b => { b.classList.remove('playing'); b.innerHTML = I.play; });
  }

  $('[data-next]', el).addEventListener('click', async () => {
    stop();
    // Tidy names and remember people for next time.
    for (const [sp, v] of Object.entries(job.speakerNames)) {
      if (v === ME) { job.speakerNames[sp] = s.myName; continue; }
      if (v === IGNORE) continue;
      job.speakerNames[sp] = v.trim().replace(/\s+/g, ' ');
      rememberPerson(job.speakerNames[sp]);
    }
    for (const st of stats) if (!job.speakerNames[st.speaker]) job.speakerNames[st.speaker] = IGNORE;
    job.status = 'writing';
    await saveJob(job);
    replace(e => writingScreen(e, job), 'enter');
  });

  render();
}

// Name input with suggestions from remembered people and (if signed in) Outlook contacts.
function attachPeopleTypeahead(input, { onInput, onPick }) {
  const field = input.closest('.field');
  let box = null;
  let timer = null;
  const close = () => { box?.remove(); box = null; };
  const show = items => {
    close();
    if (!items.length) return;
    box = document.createElement('div');
    box.className = 'suggest';
    box.innerHTML = items.map((p, i) => `<button data-i="${i}"><div>${esc(p.name)}</div>${p.email ? `<div class="e">${esc(p.email)}</div>` : ''}</button>`).join('');
    field.appendChild(box);
    $$('button', box).forEach(b => b.addEventListener('pointerdown', e => {
      e.preventDefault();
      const p = items[+b.dataset.i];
      input.value = p.name;
      close();
      onPick(p);
    }));
  };
  input.addEventListener('input', () => {
    const q = input.value.trim().toLowerCase();
    onInput(input.value);
    if (!q) { close(); return; }
    const local = getPeople().filter(p => p.name.toLowerCase().split(/\s+/).some(w => w.startsWith(q)) || p.name.toLowerCase().startsWith(q)).slice(0, 4);
    show(local);
    clearTimeout(timer);
    if (MS.msSignedIn() && q.length >= 2) {
      timer = setTimeout(async () => {
        try {
          const remote = await MS.searchPeople(input.value.trim());
          const names = new Set(local.map(p => p.name.toLowerCase()));
          if (document.activeElement === input) show([...local, ...remote.filter(p => !names.has(p.name.toLowerCase()))].slice(0, 6));
        } catch {}
      }, 350);
    }
  });
  input.addEventListener('blur', () => setTimeout(close, 150));
  input.addEventListener('keydown', e => { if (e.key === 'Enter') input.blur(); });
}

// =====================================================================
// 3. Writing the note
// =====================================================================

function cachedFolders() {
  try { return JSON.parse(localStorage.getItem('minutes.folders') || 'null'); } catch { return null; }
}

async function getFolders({ refresh = false } = {}) {
  const s = getSettings();
  if (MS.msSignedIn()) {
    const c = cachedFolders();
    if (c && !refresh && Date.now() - c.at < 6 * 3600e3) return c;
    try {
      const f = await MS.listFolders();
      const v = { ...f, at: Date.now(), live: true };
      localStorage.setItem('minutes.folders', JSON.stringify(v));
      return v;
    } catch (e) {
      if (c) return c;
    }
  }
  return {
    subject: s.manualFolders.split('\n').map(x => x.trim()).filter(Boolean),
    oneOnOnes: getPeople().map(p => p.name).sort(),
    live: false,
  };
}

function writingScreen(el, job, instruction = null) {
  el.innerHTML = `
    ${navBar({ title: 'Writing note', back: true, backLabel: 'Minutes' })}
    <main class="content">
      <h1 class="large-title">${instruction ? 'Revising note' : 'Writing note'}</h1>
      <div data-body></div>
    </main>`;
  bindBack(el);
  const body = $('[data-body]', el);
  const s = getSettings();

  async function run() {
    body.innerHTML = `<div class="big-status"><div class="pulse">${I.sparkle}</div><div class="t" data-msg>Reading the transcript…</div>
      <div class="s">Using ${esc(LLM_PROVIDERS[s.llmProvider].label)}. This usually takes under a minute.</div></div>`;
    keepAwake(true);
    try {
      const folders = await getFolders();
      const prompt = buildPrompt(job, { myName: s.myName, oneOnOnesFolder: s.oneOnOnesFolder, folders, instruction, previous: instruction ? job.note : null });
      const note = await writeNote(prompt, msg => { const m = $('[data-msg]', body); if (m) m.textContent = msg; });
      // Participants come from the names you gave, not from the AI.
      note.participants = participantNames(job);
      if (note.meeting_type === 'one_on_one' && note.one_on_one_with) {
        note.suggested_folder = `${s.oneOnOnesFolder}/${note.one_on_one_with}`;
      }
      job.note = note;
      job.save = null; // re-derive save choices from the new note
      job.status = 'noted';
      await saveJob(job);
      keepAwake(false);
      if (el.isConnected) replace(e => reviewScreen(e, job), 'enter');
    } catch (e) {
      keepAwake(false);
      body.innerHTML = `
        <div class="notice err">${I.warn}<div><div class="t">Couldn’t write the note</div><div class="err-text">${esc(e.message || e)}</div></div></div>
        <button class="btn btn-primary mt-l" data-retry>Try again</button>
        <button class="btn btn-plain mt" data-settings>Open Settings</button>`;
      $('[data-retry]', body).addEventListener('click', run);
      $('[data-settings]', body).addEventListener('click', () => push(settingsScreen));
    }
  }
  run();
}

// =====================================================================
// 4. Review
// =====================================================================

function reviewScreen(el, job) {
  const n = job.note;
  el.innerHTML = `
    ${navBar({ title: n.title, back: true, backLabel: 'Minutes', right: `<button class="nav-btn" data-transcript>Transcript</button>` })}
    <main class="content with-toolbar">
      <div class="paper">
        <div class="doc-title" contenteditable data-path="title">${esc(n.title)}</div>
        <div class="kind">${n.meeting_type === 'one_on_one' ? 'One-on-one meeting note' : 'Meeting note'}</div>
        <dl class="meta-grid">
          <dt>Date</dt><dd><input type="date" value="${esc(job.meetingDate)}" data-date></dd>
          <dt>Participants</dt><dd>${esc(n.participants.join(', '))}</dd>
        </dl>
        <h3>Key decisions</h3>
        ${n.key_decisions.length ? `<ul>${n.key_decisions.map((d, i) => `<li contenteditable data-path="key_decisions.${i}">${esc(d)}</li>`).join('')}</ul>` : '<p class="muted">No decisions recorded.</p>'}
        <h3>Discussion</h3>
        <ol>${n.narrative.map((p, i) => `<li><span class="np-head" contenteditable data-path="narrative.${i}.heading">${esc(p.heading)}</span>. <span contenteditable data-path="narrative.${i}.summary">${esc(p.summary)}</span></li>`).join('')}</ol>
        <h3>Actions</h3>
        ${n.actions.length ? `<table class="actions-table"><thead><tr><th>Action</th><th class="c-owner">Owner</th><th class="c-due">Due</th></tr></thead><tbody>
          ${n.actions.map((a, i) => `<tr><td contenteditable data-path="actions.${i}.action">${esc(a.action)}</td><td class="c-owner" contenteditable data-path="actions.${i}.owner">${esc(a.owner)}</td><td class="c-due" contenteditable data-path="actions.${i}.due">${esc(a.due || '—')}</td></tr>`).join('')}
        </tbody></table>` : '<p class="muted">No actions recorded.</p>'}
        <div style="height:14px"></div>
      </div>
      <p class="group-footer center mt">Tap any text to edit it.</p>
      <div class="group-header">Ask for changes</div>
      <div class="ask">
        <div class="field">${I.sparkle}<input data-ask placeholder="e.g. make it shorter" enterkeyhint="send"></div>
        <button class="icon-btn" data-send disabled aria-label="Revise">${I.up}</button>
      </div>
    </main>
    <footer class="toolbar"><button class="btn btn-primary" data-next>Continue to save &amp; send</button></footer>`;
  bindBack(el);

  $$('[contenteditable]', el).forEach(node => node.addEventListener('input', () => {
    const path = node.dataset.path.split('.');
    let obj = job.note;
    for (let i = 0; i < path.length - 1; i++) obj = obj[path[i]];
    let v = node.textContent.trim();
    if (path.at(-1) === 'due' && v === '—') v = '';
    obj[path.at(-1)] = v;
    if (node.dataset.path === 'title') $('.nav-title', el).textContent = v;
    saveJob(job);
  }));
  $('[data-date]', el).addEventListener('change', e => { if (e.target.value) { job.meetingDate = e.target.value; saveJob(job); } });

  const ask = $('[data-ask]', el);
  const send = $('[data-send]', el);
  ask.addEventListener('input', () => { send.disabled = !ask.value.trim(); });
  const revise = () => {
    const v = ask.value.trim();
    if (!v) return;
    replace(e => writingScreen(e, job, v), 'enter');
  };
  send.addEventListener('click', revise);
  ask.addEventListener('keydown', e => { if (e.key === 'Enter') revise(); });

  $('[data-transcript]', el).addEventListener('click', () => openSheet({
    title: 'Transcript', left: 'Close',
    build: api => {
      api.body.innerHTML = `<div class="group">${namedTranscript(job).map(u => `
        <div class="row row-stack"><div class="row-sub"><b style="color:var(--label)">${esc(u.name)}</b> · ${fmtTime(u.start)}</div><div style="font-size:16px">${esc(u.text)}</div></div>`).join('')}</div>`;
    },
  }));

  $('[data-next]', el).addEventListener('click', () => push(e => saveScreen(e, job)));
}

// =====================================================================
// 5. Save & send
// =====================================================================

async function initSave(job) {
  if (job.save) return job.save;
  const s = getSettings();
  const me = (s.myName || '').toLowerCase();
  const recipients = [];
  for (const name of job.note.participants) {
    if (name.toLowerCase() === me) continue;
    const p = findPerson(name);
    recipients.push({ name, email: p?.email || '' });
  }
  job.save = {
    folder: job.note.suggested_folder,
    reason: job.note.folder_reason,
    noteName: noteFileName(job),
    transcriptName: transcriptFileName(job),
    includeTranscript: true,
    spOn: !!s.spLastOn,
    spSite: s.spSite,
    spFolder: s.spFolder,
    emailOn: recipients.length > 0,
    recipients,
    done: {},
  };
  await saveJob(job);
  // Fill in missing email addresses from Outlook, if signed in.
  if (MS.msSignedIn()) {
    for (const r of job.save.recipients) {
      if (r.email) continue;
      try {
        const hits = await MS.searchPeople(r.name);
        const exact = hits.find(h => h.name.toLowerCase() === r.name.toLowerCase()) || hits[0];
        if (exact?.email) { r.email = exact.email; rememberPerson(r.name, r.email); }
      } catch {}
    }
    await saveJob(job);
  }
  return job.save;
}

function saveScreen(el, job) {
  const s = getSettings();
  const signedIn = MS.msSignedIn();
  el.innerHTML = `
    ${navBar({ title: 'Save & send', back: true, backLabel: 'Note' })}
    <main class="content with-toolbar"><h1 class="large-title">Save &amp; send</h1><div data-body><div class="big-status"><div class="spinner" style="margin:0 auto"></div></div></div></main>
    <footer class="toolbar"><button class="btn btn-primary" data-go disabled>${signedIn ? 'Save' : 'Next'}</button></footer>`;
  bindBack(el);
  const body = $('[data-body]', el);

  initSave(job).then(render);

  function render() {
    const sv = job.save;
    body.innerHTML = `
      ${signedIn ? '' : `<div class="notice info">${I.info}<div><div class="s">Not signed in to Microsoft 365. Next, you’ll share the files to the OneDrive and Outlook apps yourself.</div></div></div>`}
      <div class="group-header">OneDrive</div>
      <div class="group">
        <button class="row has-icon" data-folder>
          <div class="row-icon" style="background:var(--accent)">${I.folder}</div>
          <div class="row-main"><div class="row-title">${esc(displayPath(sv.folder) || 'Choose folder')}</div><div class="row-sub">Folder</div></div>
          <span class="chev">${I.chev}</span>
        </button>
        <button class="row has-icon" data-rename>
          <div class="row-icon" style="background:#5856d6">${I.doc}</div>
          <div class="row-main"><div class="row-title">${esc(sv.noteName)}</div><div class="row-sub">Meeting note (Word)</div></div>
          <span class="chev">${I.chev}</span>
        </button>
        <div class="row has-icon">
          <div class="row-icon" style="background:#8e8e93">${I.doc}</div>
          <div class="row-main"><div class="row-title">Save transcript</div><div class="row-sub">As a separate Word file</div></div>
          <label class="switch"><input type="checkbox" data-tr ${sv.includeTranscript ? 'checked' : ''}><span></span></label>
        </div>
      </div>
      ${sv.reason ? `<div class="group-footer">${I.sparkle.replace('<svg', '<svg style="display:inline;width:12px;height:12px;vertical-align:-1px;margin-right:4px"')}${esc(sv.reason)}</div>` : ''}

      <div class="group-header">SharePoint</div>
      <div class="group">
        <div class="row has-icon">
          <div class="row-icon" style="background:#038387">${I.cloud}</div>
          <div class="row-main"><div class="row-title">Also save to SharePoint</div></div>
          <label class="switch"><input type="checkbox" data-sp ${sv.spOn ? 'checked' : ''}><span></span></label>
        </div>
        ${sv.spOn && signedIn ? `
        <button class="row" data-site><div class="row-main"><div class="row-title">Site</div></div><div class="row-value">${esc(sv.spSite?.name || 'Choose')}</div><span class="chev">${I.chev}</span></button>
        <button class="row" data-spfolder ${sv.spSite ? '' : 'disabled style="opacity:.4"'}><div class="row-main"><div class="row-title">Folder</div></div><div class="row-value">${esc(sv.spFolder?.path || 'Documents')}</div><span class="chev">${I.chev}</span></button>` : ''}
      </div>
      ${sv.spOn && !signedIn ? `<div class="group-footer">You’ll share the files to SharePoint on the next screen.</div>` : sv.spOn ? `<div class="group-footer">Remembered for next time.</div>` : ''}

      <div class="group-header">Email draft</div>
      <div class="group">
        <div class="row has-icon">
          <div class="row-icon" style="background:#0a64d8">${I.mail}</div>
          <div class="row-main"><div class="row-title">Create email draft</div></div>
          <label class="switch"><input type="checkbox" data-email ${sv.emailOn ? 'checked' : ''}><span></span></label>
        </div>
        ${sv.emailOn ? `<div class="row row-stack">
          <div class="chips" style="margin-top:0">${sv.recipients.map((r, i) => `<button class="chip${r.email ? '' : ' warn'}" data-rcp="${i}"><span>${esc(r.name)}${r.email ? '' : ' · add email'}</span>${I.x}</button>`).join('')}</div>
          <div class="field">${I.plus}<input data-addr placeholder="Add a person" autocomplete="off" enterkeyhint="done"></div>
        </div>` : ''}
      </div>
      ${sv.emailOn ? `<div class="group-footer">${signedIn ? 'A draft is created in Outlook with the note attached, so you can check it before sending.' : 'You’ll share the note to Outlook on the next screen. The addresses are copied for you to paste.'} Tap a name to edit the address.</div>` : ''}
    `;
    const go = $('[data-go]', el);
    go.disabled = !sv.folder;

    $('[data-folder]', body).addEventListener('click', () => folderSheet(job, render));
    $('[data-rename]', body).addEventListener('click', () => promptSheet({
      title: 'File name', value: sv.noteName.replace(/\.docx$/, ''),
      onSave: v => {
        if (!v) return;
        sv.noteName = `${safeName(v)}.docx`;
        sv.transcriptName = `${safeName(v)} – Transcript.docx`;
        saveJob(job); render();
      },
    }));
    $('[data-tr]', body).addEventListener('change', e => { sv.includeTranscript = e.target.checked; saveJob(job); });
    $('[data-sp]', body).addEventListener('change', e => { sv.spOn = e.target.checked; saveSettings({ spLastOn: sv.spOn }); saveJob(job); render(); });
    $('[data-site]', body)?.addEventListener('click', () => siteSheet(job, render));
    $('[data-spfolder]', body)?.addEventListener('click', () => { if (sv.spSite) spFolderSheet(job, render); });
    $('[data-email]', body).addEventListener('change', e => { sv.emailOn = e.target.checked; saveJob(job); render(); });
    $$('[data-rcp]', body).forEach(b => b.addEventListener('click', ev => {
      const i = +b.dataset.rcp;
      const r = sv.recipients[i];
      // tapping the × removes; tapping the name edits the address
      if (ev.target.closest('svg')) { sv.recipients.splice(i, 1); saveJob(job); render(); return; }
      promptSheet({
        title: r.name, value: r.email, type: 'email', placeholder: 'name@company.com',
        onSave: v => { r.email = v; if (v) rememberPerson(r.name, v); saveJob(job); render(); },
      });
    }));
    const addr = $('[data-addr]', body);
    if (addr) {
      attachPeopleTypeahead(addr, {
        onInput: () => {},
        onPick: p => { sv.recipients.push({ name: p.name, email: p.email || findPerson(p.name)?.email || '' }); saveJob(job); render(); },
      });
      addr.addEventListener('keydown', e => {
        if (e.key !== 'Enter' || !addr.value.trim()) return;
        const v = addr.value.trim();
        const isEmail = /@/.test(v);
        sv.recipients.push(isEmail ? { name: v, email: v } : { name: v, email: findPerson(v)?.email || '' });
        saveJob(job); render();
      });
    }
  }

  $('[data-go]', el).addEventListener('click', async () => {
    const sv = job.save;
    if (sv.emailOn && signedIn && sv.recipients.some(r => !r.email)) {
      toast('Add an email address for everyone, or remove them.');
      return;
    }
    if (sv.spOn && signedIn && !sv.spSite) { toast('Choose a SharePoint site.'); return; }
    job.status = 'saving';
    await saveJob(job);
    replace(e => (signedIn ? savingScreen(e, job) : manualShareScreen(e, job)), 'enter');
  });
}

function folderSheet(job, onDone) {
  const s = getSettings();
  const sv = job.save;
  openSheet({
    title: 'Choose folder', left: 'Cancel',
    build: async api => {
      api.body.innerHTML = `<div class="field search">${I.search}<input placeholder="Search folders" data-q></div><div data-list><div class="big-status"><div class="spinner" style="margin:0 auto"></div></div></div>`;
      const folders = await getFolders();
      const oo = s.oneOnOnesFolder;
      const others = job.note.participants.filter(p => p.toLowerCase() !== (s.myName || '').toLowerCase());
      const ooNames = [...new Set([...others, ...folders.oneOnOnes])];
      const list = $('[data-list]', api.body);
      const q = $('[data-q]', api.body);
      const pick = path => { sv.folder = path; sv.reason = ''; saveJob(job); api.close(); onDone(); };

      const draw = () => {
        const f = q.value.trim().toLowerCase();
        const m = name => !f || name.toLowerCase().includes(f);
        const row = (path, label, sub, color, icon) => `<button class="row has-icon" data-path="${esc(path)}">
          <div class="row-icon" style="background:${color}">${icon}</div>
          <div class="row-main"><div class="row-title">${esc(label)}</div>${sub ? `<div class="row-sub">${esc(sub)}</div>` : ''}</div>
          ${path === sv.folder ? `<span style="color:var(--accent);width:20px">${I.check}</span>` : ''}</button>`;
        const sug = job.note.suggested_folder;
        const subj = folders.subject.filter(m);
        const ooList = ooNames.filter(m);
        let html = '';
        if (!f && sug) html += `<div class="group-header">Suggested</div><div class="group">${row(sug, displayPath(sug), job.note.folder_reason, 'var(--accent)', I.sparkle)}</div>`;
        html += `<div class="group-header">${esc(oo)}</div><div class="group">${ooList.map(n => row(`${oo}/${n}`, n, folders.oneOnOnes.includes(n) ? '' : 'New folder', '#ff9500', I.person)).join('') || '<div class="empty">No matches</div>'}</div>`;
        html += `<div class="group-header">Projects</div><div class="group">${subj.map(n => row(n, n, '', 'var(--accent)', I.folder)).join('') || `<div class="empty">${folders.live ? 'No matches' : 'Sign in to Microsoft 365, or add your folder names in Settings.'}</div>`}</div>`;
        html += `<div class="group mt-l"><button class="row link" data-other><div class="row-main"><div class="row-title">Type a folder path…</div></div></button></div>`;
        list.innerHTML = html;
        $$('[data-path]', list).forEach(b => b.addEventListener('click', () => pick(b.dataset.path)));
        $('[data-other]', list).addEventListener('click', () => {
          api.close();
          setTimeout(() => promptSheet({
            title: 'Folder path', value: sv.folder, placeholder: 'e.g. 2 Budget/Notes',
            onSave: v => { if (v) { sv.folder = v.replace(/^\/+|\/+$/g, ''); saveJob(job); onDone(); } },
          }), 300);
        });
      };
      q.addEventListener('input', draw);
      draw();
    },
  });
}

function siteSheet(job, onDone) {
  const sv = job.save;
  openSheet({
    title: 'SharePoint site',
    build: async api => {
      api.body.innerHTML = `<div class="field search">${I.search}<input placeholder="Search sites" data-q></div><div data-list><div class="big-status"><div class="spinner" style="margin:0 auto"></div></div></div>`;
      const list = $('[data-list]', api.body);
      const q = $('[data-q]', api.body);
      const draw = (sites, label) => {
        list.innerHTML = `<div class="group-header">${esc(label)}</div><div class="group">${sites.map((x, i) => `<button class="row" data-i="${i}"><div class="row-main"><div class="row-title">${esc(x.name)}</div><div class="row-sub">${esc((x.webUrl || '').replace(/^https:\/\//, ''))}</div></div></button>`).join('') || '<div class="empty">No sites found</div>'}</div>`;
        $$('[data-i]', list).forEach(b => b.addEventListener('click', () => {
          const site = sites[+b.dataset.i];
          sv.spSite = site; sv.spFolder = null;
          saveSettings({ spSite: site, spFolder: null });
          saveJob(job); api.close(); onDone();
        }));
      };
      try { draw(await MS.followedSites(), 'Sites you follow'); } catch (e) { list.innerHTML = `<div class="notice err">${I.warn}<div class="s">${esc(e.message)}</div></div>`; }
      let t;
      q.addEventListener('input', () => {
        clearTimeout(t);
        t = setTimeout(async () => {
          const v = q.value.trim();
          try { draw(v ? await MS.searchSites(v) : await MS.followedSites(), v ? 'Results' : 'Sites you follow'); } catch {}
        }, 350);
      });
    },
  });
}

function spFolderSheet(job, onDone) {
  const sv = job.save;
  const trail = [{ id: null, name: 'Documents' }];
  openSheet({
    title: 'SharePoint folder', right: 'Choose',
    onRight: api => {
      const cur = trail.at(-1);
      sv.spFolder = cur.id ? { id: cur.id, path: trail.slice(1).map(t => t.name).join(' › ') } : null;
      saveSettings({ spFolder: sv.spFolder });
      saveJob(job); api.close(); onDone();
    },
    build: api => {
      const draw = async () => {
        api.body.innerHTML = `<div class="group-header">${esc(trail.map(t => t.name).join(' › '))}</div><div data-list><div class="big-status"><div class="spinner" style="margin:0 auto"></div></div></div>`;
        const list = $('[data-list]', api.body);
        try {
          const kids = await MS.siteFolderChildren(sv.spSite.id, trail.at(-1).id);
          list.innerHTML = `<div class="group">
            ${trail.length > 1 ? `<button class="row link" data-up><div class="row-main"><div class="row-title">‹ Up one level</div></div></button>` : ''}
            ${kids.map((k, i) => `<button class="row has-icon" data-i="${i}"><div class="row-icon" style="background:#038387">${I.folder}</div><div class="row-main"><div class="row-title">${esc(k.name)}</div></div><span class="chev">${I.chev}</span></button>`).join('') || '<div class="empty">No subfolders. Tap Choose to save here.</div>'}
          </div>`;
          $('[data-up]', list)?.addEventListener('click', () => { trail.pop(); draw(); });
          $$('[data-i]', list).forEach(b => b.addEventListener('click', () => { trail.push(kids[+b.dataset.i]); draw(); }));
        } catch (e) {
          list.innerHTML = `<div class="notice err">${I.warn}<div class="s">${esc(e.message)}</div></div>`;
        }
      };
      draw();
    },
  });
}

// ---------- Saving (signed in: direct to Microsoft 365) ----------

function savingScreen(el, job) {
  const s = getSettings();
  const sv = job.save;
  const steps = [
    { key: 'note', label: 'Save note to OneDrive' },
    ...(sv.includeTranscript ? [{ key: 'transcript', label: 'Save transcript to OneDrive' }] : []),
    ...(sv.spOn ? [{ key: 'sp', label: 'Save to SharePoint' }] : []),
    ...(sv.emailOn && sv.recipients.length ? [{ key: 'email', label: 'Create Outlook draft' }] : []),
  ];
  el.innerHTML = `
    ${navBar({ title: 'Saving' })}
    <main class="content"><h1 class="large-title">Saving</h1><div class="group steps" data-steps></div><div data-err></div></main>`;
  const box = $('[data-steps]', el);
  let current = null;
  let failed = null;
  const draw = () => {
    box.innerHTML = steps.map(st => {
      const done = !!sv.done[st.key];
      const cls = done ? 'done' : st.key === failed ? 'error' : st.key === current ? 'active' : 'pending';
      const dot = done ? I.check : st.key === failed ? I.x : st.key === current ? '<div class="spinner"></div>' : '';
      return `<div class="row step ${cls}"><div class="dot">${dot}</div><div class="row-main"><div class="row-title">${esc(st.label)}</div></div></div>`;
    }).join('');
  };

  async function run() {
    failed = null;
    $('[data-err]', el).innerHTML = '';
    const author = s.myName;
    const note = noteDocx(job, author);
    const tr = sv.includeTranscript ? transcriptDocx(job, author) : null;
    try {
      for (const st of steps) {
        if (sv.done[st.key]) continue;
        current = st.key; draw();
        if (st.key === 'note') {
          const r = await MS.uploadToOneDrive(sv.folder, sv.noteName, note);
          sv.done.note = { webUrl: r.webUrl, name: r.name };
        } else if (st.key === 'transcript') {
          const r = await MS.uploadToOneDrive(sv.folder, sv.transcriptName, tr);
          sv.done.transcript = { webUrl: r.webUrl, name: r.name };
        } else if (st.key === 'sp') {
          const r = await MS.uploadToSharePoint(sv.spSite.id, sv.spFolder?.id, sv.noteName, note);
          if (tr) await MS.uploadToSharePoint(sv.spSite.id, sv.spFolder?.id, sv.transcriptName, tr);
          sv.done.sp = { webUrl: r.webUrl };
        } else if (st.key === 'email') {
          const r = await MS.createDraft({
            to: sv.recipients,
            subject: emailSubject(job),
            html: emailHtml(job, s.myName),
            attachments: [{ name: sv.noteName, blob: note }],
          });
          sv.done.email = { webLink: r.webLink };
        }
        await saveJob(job);
      }
      current = null; draw();
      await finishJob(job);
      replace(e => doneScreen(e, job), 'enter');
    } catch (e) {
      failed = current; current = null; draw();
      const signIn = e instanceof MS.SignInNeeded;
      $('[data-err]', el).innerHTML = `
        <div class="notice err">${I.warn}<div><div class="t">${signIn ? 'Sign in again' : 'Couldn’t finish saving'}</div><div class="err-text">${esc(e.message)}</div></div></div>
        <button class="btn btn-primary mt-l" data-retry>${signIn ? 'Sign in to Microsoft 365' : 'Try again'}</button>
        <button class="btn btn-plain mt" data-manual>Share the files instead</button>`;
      $('[data-retry]', el).addEventListener('click', () => {
        if (signIn) { sessionStorage.setItem('minutes.resume', job.id); MS.signIn(); } else run();
      });
      $('[data-manual]', el).addEventListener('click', () => replace(x => manualShareScreen(x, job), 'enter'));
    }
  }
  draw();
  run();
}

async function finishJob(job) {
  job.status = 'complete';
  job.completedAt = Date.now();
  await saveJob(job);
  await DB.del('audio', job.id); // the app's own copy of the recording is no longer needed
}

// ---------- Manual route (not signed in): Android Share sheet ----------

async function shareFiles(files, { title, text } = {}) {
  const data = { files, title, text };
  if (navigator.canShare && navigator.canShare({ files })) {
    try { await navigator.share(data); return true; } catch (e) { if (e.name === 'AbortError') return false; }
  }
  // Fallback: download
  for (const f of files) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(f);
    a.download = f.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  toast('Saved to Downloads.');
  return true;
}

async function copy(text, msg = 'Copied') {
  try { await navigator.clipboard.writeText(text); toast(msg); } catch { promptSheet({ title: 'Copy', value: text, onSave: () => {} }); }
}

function manualShareScreen(el, job) {
  const s = getSettings();
  const sv = job.save;
  const mk = (blob, name) => new File([blob], name, { type: blob.type });
  const noteFile = () => mk(noteDocx(job, s.myName), sv.noteName);
  const trFile = () => mk(transcriptDocx(job, s.myName), sv.transcriptName);
  const files = () => (sv.includeTranscript ? [noteFile(), trFile()] : [noteFile()]);
  sv.manual ||= {};

  const steps = [
    { key: 'od', title: 'Save to OneDrive', body: `Tap <b>Share</b>, choose <b>OneDrive</b>, then go to <span class="path">${esc(displayPath(sv.folder))}</span> and tap Upload.`, buttons: [['Share files', 'primary', () => shareFiles(files(), { title: job.note.title })], ['Copy folder', 'secondary', () => copy(displayPath(sv.folder).replace(/ › /g, '/'), 'Folder path copied')]] },
    ...(sv.spOn ? [{ key: 'sp', title: 'Save to SharePoint', body: 'Tap <b>Share</b>, choose <b>OneDrive</b>, then open <b>Shared libraries</b> and pick your SharePoint site and folder.', buttons: [['Share files', 'primary', () => shareFiles(files(), { title: job.note.title })]] }] : []),
    ...(sv.emailOn ? [{ key: 'em', title: 'Email the note', body: `Copy the addresses, then tap <b>Share</b> and choose <b>Outlook</b>. Paste the addresses into <b>To</b>.${sv.recipients.some(r => !r.email) ? ' Some people have no saved address, so add them in Outlook.' : ''}`, buttons: [['Copy addresses', 'secondary', () => copy(sv.recipients.filter(r => r.email).map(r => r.email).join('; '), 'Addresses copied')], ['Share to Outlook', 'primary', () => shareFiles([noteFile()], { title: emailSubject(job), text: emailText(job, s.myName) })]] }] : []),
  ];

  el.innerHTML = `
    ${navBar({ title: 'Share', back: true, backLabel: 'Back' })}
    <main class="content with-toolbar"><h1 class="large-title">Share</h1><p class="lede">Your files are ready. Complete each step, then tap Finish.</p><div data-steps></div>
      <div class="notice warn mt-l">${I.info}<div><div class="t">Then delete the recording</div><div class="s">Once it’s saved, delete <b>${esc(job.fileName)}</b> in Easy Recorder.</div></div></div>
    </main>
    <footer class="toolbar"><button class="btn btn-primary" data-finish>Finish</button></footer>`;
  bindBack(el);
  const box = $('[data-steps]', el);
  const draw = () => {
    box.innerHTML = steps.map((st, i) => `<div class="step-card${sv.manual[st.key] ? ' done' : ''}">
      <div class="h"><div class="n">${sv.manual[st.key] ? I.check.replace('<svg', '<svg style="width:14px;height:14px"') : i + 1}</div>${esc(st.title)}</div>
      <p>${st.body}</p>
      <div class="btn-row">${st.buttons.map(([label, kind], j) => `<button class="btn btn-${kind}" data-s="${i}" data-b="${j}">${kind === 'primary' ? I.share : I.copy}<span>${esc(label)}</span></button>`).join('')}</div>
    </div>`).join('');
    $$('[data-s]', box).forEach(b => b.addEventListener('click', async () => {
      const st = steps[+b.dataset.s];
      const [, kind, fn] = st.buttons[+b.dataset.b];
      const ok = await fn();
      if (kind === 'primary' && ok !== false) { sv.manual[st.key] = true; saveJob(job); draw(); }
    }));
  };
  draw();
  $('[data-finish]', el).addEventListener('click', async () => {
    sv.done.manual = true;
    await finishJob(job);
    replace(e => doneScreen(e, job), 'enter');
  });
}

// =====================================================================
// 6. Done
// =====================================================================

function doneScreen(el, job) {
  const s = getSettings();
  const sv = job.save || { done: {} };
  const d = sv.done || {};
  const rows = [];
  if (d.note) rows.push({ icon: I.doc, color: 'var(--accent)', title: d.note.name || sv.noteName, sub: `OneDrive › ${displayPath(sv.folder)}`, href: d.note.webUrl });
  if (d.transcript) rows.push({ icon: I.doc, color: '#8e8e93', title: d.transcript.name || sv.transcriptName, sub: 'Transcript', href: d.transcript.webUrl });
  if (d.sp) rows.push({ icon: I.cloud, color: '#038387', title: 'Saved to SharePoint', sub: `${sv.spSite?.name || ''}${sv.spFolder ? ` › ${sv.spFolder.path}` : ''}`, href: d.sp.webUrl });
  if (d.email) rows.push({ icon: I.mail, color: '#0a64d8', title: 'Email draft ready', sub: `To ${sv.recipients.map(r => r.name).join(', ')}. Open it to check and send.`, href: d.email.webLink });

  el.innerHTML = `
    ${navBar({ title: job.note?.title || 'Done', back: true, backLabel: 'Minutes' })}
    <main class="content">
      <div class="success"><div class="tick">${I.check}</div><div class="t">${d.manual ? 'All done' : 'Saved'}</div><div class="s">${esc(job.note?.title || '')} · ${esc(longDate(job.meetingDate))}</div></div>
      ${rows.length ? `<div class="group mt-l">${rows.map(r => `
        <a class="row has-icon" ${r.href ? `href="${esc(r.href)}" target="_blank" rel="noopener"` : ''}>
          <div class="row-icon" style="background:${r.color}">${r.icon}</div>
          <div class="row-main"><div class="row-title">${esc(r.title)}</div><div class="row-sub">${esc(r.sub)}</div></div>
          ${r.href ? `<span class="chev">${I.chev}</span>` : ''}
        </a>`).join('')}</div>` : ''}
      <div class="notice warn mt">${I.info}<div><div class="t">You can now delete the recording</div><div class="s">Open Easy Recorder and delete <b>${esc(job.fileName)}</b>. Minutes has already removed its own copy.</div></div></div>
      <div class="group-header">Files</div>
      <div class="group">
        <button class="row link" data-share><div class="row-main"><div class="row-title">Share the note again</div></div></button>
        <button class="row link" data-review><div class="row-main"><div class="row-title">View the note</div></div></button>
      </div>
      <button class="btn btn-primary mt-l" data-home>Done</button>
    </main>`;
  bindBack(el);
  $('[data-home]', el).addEventListener('click', backToHome);
  $('[data-share]', el).addEventListener('click', () => {
    const f = new File([noteDocx(job, s.myName)], sv.noteName || noteFileName(job), { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    shareFiles([f], { title: job.note.title });
  });
  $('[data-review]', el).addEventListener('click', () => openSheet({
    title: 'Note', left: 'Close',
    build: api => {
      const n = job.note;
      api.body.innerHTML = `<div class="paper"><div class="doc-title">${esc(n.title)}</div><div class="kind">${esc(longDate(job.meetingDate))}</div>
        <dl class="meta-grid"><dt>Participants</dt><dd>${esc(n.participants.join(', '))}</dd></dl>
        <h3>Key decisions</h3><ul>${n.key_decisions.map(x => `<li>${esc(x)}</li>`).join('')}</ul>
        <h3>Discussion</h3><ol>${n.narrative.map(p => `<li><span class="np-head">${esc(p.heading)}</span>. ${esc(p.summary)}</li>`).join('')}</ol>
        <h3>Actions</h3>${n.actions.length ? `<table class="actions-table"><thead><tr><th>Action</th><th class="c-owner">Owner</th><th class="c-due">Due</th></tr></thead><tbody>${n.actions.map(a => `<tr><td>${esc(a.action)}</td><td class="c-owner">${esc(a.owner)}</td><td class="c-due">${esc(a.due || '—')}</td></tr>`).join('')}</tbody></table>` : '<p class="muted">No actions recorded.</p>'}
        <div style="height:14px"></div></div>`;
    },
  }));
}

// =====================================================================
// Settings
// =====================================================================

function settingsScreen(el) {
  const s = getSettings();
  const acct = MS.msAccount();
  const input = (key, label, { type = 'text', placeholder = '', mode = '' } = {}) => `
    <div class="row"><div class="row-title" style="flex:none">${esc(label)}</div>
    <input class="inline" data-k="${key}" type="${type}" value="${esc(s[key] || '')}" placeholder="${esc(placeholder)}" autocomplete="off" autocapitalize="off" spellcheck="false" ${mode ? `inputmode="${mode}"` : ''}></div>`;
  const seg = (key, options) => `<div class="row"><div class="segmented" data-seg="${key}">${Object.entries(options).map(([k, v]) => `<button class="${s[key] === k ? 'on' : ''}" data-v="${k}">${esc(v.label)}</button>`).join('')}</div></div>`;

  const sttFields = {
    assemblyai: input('assemblyaiKey', 'API key', { type: 'password', placeholder: 'Paste key' }),
    openai: input('openaiKey', 'API key', { type: 'password', placeholder: 'sk-…' }),
    azure: input('azureSpeechKey', 'Key', { type: 'password', placeholder: 'Paste key' }) + input('azureSpeechRegion', 'Region', { placeholder: 'uksouth' }),
  }[s.sttProvider];
  const llmFields = {
    anthropic: input('anthropicKey', 'API key', { type: 'password', placeholder: 'sk-ant-…' }) + input('anthropicModel', 'Model'),
    openai: input('openaiKey', 'API key', { type: 'password', placeholder: 'sk-…' }) + input('openaiModel', 'Model'),
    azureopenai: input('azureOpenaiEndpoint', 'Endpoint', { placeholder: 'https://….openai.azure.com', mode: 'url' }) + input('azureOpenaiDeployment', 'Deployment') + input('azureOpenaiKey', 'Key', { type: 'password' }),
  }[s.llmProvider];

  el.innerHTML = `
    ${navBar({ title: 'Settings', back: true, backLabel: 'Minutes' })}
    <main class="content">
      <h1 class="large-title">Settings</h1>

      <div class="group mt">
        <button class="row has-icon" data-helpall><div class="row-icon" style="background:var(--accent)">${I.info}</div><div class="row-main"><div class="row-title">Help and setup guides</div></div><span class="chev">${I.chev}</span></button>
      </div>

      <div class="group-header">You</div>
      <div class="group">${input('myName', 'Your name', { placeholder: 'Firstname Surname' })}${input('myEmail', 'Email', { type: 'email', placeholder: 'Optional' })}</div>
      <div class="group-footer">The voice that talks most is labelled with your name automatically.</div>

      <div class="group-header">Transcription</div>
      <div class="group">${seg('sttProvider', STT_PROVIDERS)}${sttFields}</div>
      <div class="group-footer">${esc(STT_PROVIDERS[s.sttProvider].hint)} <a href="#" data-guide="${{ assemblyai: 'assemblyai', openai: 'openai-stt', azure: 'azure-speech' }[s.sttProvider]}">How to set up ${esc(STT_PROVIDERS[s.sttProvider].label)} ›</a></div>

      <div class="group-header">Note writing</div>
      <div class="group">${seg('llmProvider', LLM_PROVIDERS)}${llmFields}</div>
      <div class="group-footer">${esc(LLM_PROVIDERS[s.llmProvider].hint)} Keys are stored only on this phone. <a href="#" data-guide="${{ anthropic: 'claude', openai: 'openai-llm', azureopenai: 'azure-openai' }[s.llmProvider]}">How to set up ${esc(LLM_PROVIDERS[s.llmProvider].label)} ›</a></div>

      <div class="group-header">Microsoft 365</div>
      <div class="group">
        ${input('msClientId', 'Client ID', { placeholder: 'From your IT team' })}
        ${input('msTenant', 'Tenant', { placeholder: 'organizations' })}
        ${acct
          ? `<div class="row"><div class="row-main"><div class="row-title">${esc(acct.name || 'Signed in')}</div><div class="row-sub">${esc(acct.email || '')}</div></div></div>
             <button class="row destructive" data-signout><div class="row-main"><div class="row-title">Sign out</div></div></button>`
          : `<button class="row link" data-signin ${s.msClientId ? '' : 'disabled style="opacity:.4"'}><div class="row-main"><div class="row-title">Sign in</div></div></button>`}
      </div>
      <div class="group-footer">Optional. When signed in, notes save straight to OneDrive and SharePoint and the email draft is created for you. Without it, you share the files to the OneDrive and Outlook apps yourself. <a href="#" data-guide="microsoft365">What to ask IT ›</a></div>

      <div class="group-header">Folders</div>
      <div class="group">
        ${input('oneOnOnesFolder', 'One-on-ones')}
        ${input('subjectPrefixes', 'Project prefixes', { placeholder: '1,2,3' })}
        <button class="row" data-refresh ${acct ? '' : 'disabled style="opacity:.4"'}><div class="row-main"><div class="row-title">Refresh folder list</div></div><div class="row-value" data-fcount></div></button>
      </div>
      <div class="group-footer">Project folders are the top-level OneDrive folders starting with these prefixes.</div>
      ${acct ? '' : `
      <div class="group-header">Your project folders</div>
      <div class="group"><div class="row"><textarea data-k="manualFolders" placeholder="One per line, e.g.&#10;1 Budget Review&#10;2 Recruitment">${esc(s.manualFolders)}</textarea></div></div>
      <div class="group-footer">Used for suggestions until you sign in to Microsoft 365.</div>`}

      <div class="group-header">People</div>
      <div class="group"><button class="row" data-people><div class="row-main"><div class="row-title">Remembered people</div></div><div class="row-value">${getPeople().length}</div><span class="chev">${I.chev}</span></button></div>

      <div class="group-header">Language</div>
      <div class="group"><div class="row"><div class="row-title" style="flex:none">Recordings</div>
        <select class="inline" data-k="language" style="direction:rtl">
          ${[['en-GB', 'English (UK)'], ['en-US', 'English (US)'], ['en-IE', 'English (Ireland)'], ['fr-FR', 'French'], ['de-DE', 'German'], ['es-ES', 'Spanish']].map(([v, l]) => `<option value="${v}" ${s.language === v ? 'selected' : ''}>${l}</option>`).join('')}
        </select></div></div>
      <p class="group-footer center mt-l">Minutes · your recordings and transcripts go only to the services chosen above.</p>
    </main>`;
  bindBack(el);

  $$('[data-k]', el).forEach(inp => inp.addEventListener(inp.tagName === 'SELECT' ? 'change' : 'input', () => {
    saveSettings({ [inp.dataset.k]: inp.value.trim() });
    if (inp.dataset.k === 'msClientId') { const b = $('[data-signin]', el); if (b) { b.disabled = !inp.value.trim(); b.style.opacity = inp.value.trim() ? '' : '.4'; } }
  }));
  $$('[data-seg]', el).forEach(seg => $$('button', seg).forEach(b => b.addEventListener('click', () => {
    saveSettings({ [seg.dataset.seg]: b.dataset.v });
    const y = $('.content', el).scrollTop;
    settingsScreen(el);
    $('.content', el).scrollTop = y;
    el.classList.toggle('scrolled', y > 28);
    $('.content', el).addEventListener('scroll', () => el.classList.toggle('scrolled', $('.content', el).scrollTop > 28), { passive: true });
  })));
  $('[data-signin]', el)?.addEventListener('click', () => MS.signIn().catch(e => toast(e.message)));
  $('[data-signout]', el)?.addEventListener('click', () => { MS.signOut(); localStorage.removeItem('minutes.folders'); settingsScreen(el); });
  const fc = $('[data-fcount]', el);
  const c = cachedFolders();
  if (fc && c) fc.textContent = `${c.subject.length} projects`;
  $('[data-refresh]', el)?.addEventListener('click', async () => {
    fc.textContent = 'Loading…';
    try { const f = await getFolders({ refresh: true }); fc.textContent = `${f.subject.length} projects`; toast(`Found ${f.subject.length} project folders and ${f.oneOnOnes.length} one-on-one folders.`); } catch (e) { fc.textContent = ''; toast(e.message); }
  });
  $('[data-people]', el).addEventListener('click', () => push(peopleScreen));
  $('[data-helpall]', el).addEventListener('click', () => push(helpScreen));
  $$('[data-guide]', el).forEach(a => a.addEventListener('click', e => { e.preventDefault(); push(x => helpTopicScreen(x, a.dataset.guide)); }));
}

function peopleScreen(el) {
  const people = getPeople().sort((a, b) => a.name.localeCompare(b.name));
  el.innerHTML = `
    ${navBar({ title: 'People', back: true, backLabel: 'Settings' })}
    <main class="content">
      <h1 class="large-title">People</h1>
      <p class="lede">Names you’ve given speakers, offered as suggestions next time. Tap to add or change an email address.</p>
      <div class="group">${people.map((p, i) => `
        <div class="row"><button class="row-main" style="text-align:left" data-edit="${i}"><div class="row-title">${esc(p.name)}</div><div class="row-sub">${esc(p.email || 'No email')}</div></button>
        <button class="nav-btn" style="color:var(--red)" data-del="${i}" aria-label="Remove">${I.trash}</button></div>`).join('') || '<div class="empty">No one yet.</div>'}
      </div>
    </main>`;
  bindBack(el);
  $$('[data-edit]', el).forEach(b => b.addEventListener('click', () => {
    const p = people[+b.dataset.edit];
    promptSheet({ title: p.name, value: p.email, type: 'email', placeholder: 'name@company.com', onSave: v => { p.email = v; savePeople(people); peopleScreen(el); } });
  }));
  $$('[data-del]', el).forEach(b => b.addEventListener('click', () => {
    people.splice(+b.dataset.del, 1); savePeople(people); peopleScreen(el);
  }));
}


// =====================================================================
// Help
// =====================================================================

const appAddress = () => location.origin + location.pathname.replace(/index\.html$/, '');

function topicStatus(id) {
  const s = getSettings();
  const t = HELP_TOPICS[id];
  if (!t.use) {
    if (id === 'microsoft365') return MS.msSignedIn() ? 'Signed in' : '';
    return '';
  }
  const inUse = s[t.use.setting] === t.use.value;
  const filled = (t.fields || []).every(f => s[f.key]);
  return inUse && filled ? 'In use' : filled ? 'Set up' : '';
}

function helpScreen(el) {
  el.innerHTML = `
    ${navBar({ title: 'Help', back: true, backLabel: 'Back' })}
    <main class="content">
      <h1 class="large-title">Help</h1>
      <p class="lede">Step-by-step guides for setting up each service. Each guide lets you paste in your key and check it works.</p>
      ${HELP_SECTIONS.map(sec => `
        <div class="group-header">${esc(sec.title)}</div>
        <div class="group">${sec.topics.map(id => {
          const t = HELP_TOPICS[id];
          const st = topicStatus(id);
          return `<button class="row has-icon" data-topic="${id}">
            <div class="row-icon" style="background:${t.color}">${I[t.icon] || I.info}</div>
            <div class="row-main"><div class="row-title">${esc(t.title)}</div></div>
            ${st ? `<span class="status-pill${st === 'In use' || st === 'Signed in' ? ' on' : ''}">${esc(st)}</span>` : ''}
            <span class="chev">${I.chev}</span>
          </button>`;
        }).join('')}</div>`).join('')}
    </main>`;
  bindBack(el);
  $$('[data-topic]', el).forEach(b => b.addEventListener('click', () => push(x => helpTopicScreen(x, b.dataset.topic))));
}

function helpTable(tb) {
  return `<div class="group-header">${esc(tb.caption)}</div>
    <div class="group table-card"><table class="help-table">
      <colgroup><col style="width:27%"><col style="width:32%"><col style="width:23%"><col style="width:18%"></colgroup>
      <thead><tr>${tb.header.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${tb.rows.map(r => `<tr>${r.map((c, i) => i === 0 ? `<th scope="row">${esc(c)}</th>` : `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
}

function helpTopicScreen(el, id) {
  const t = HELP_TOPICS[id];
  const s = getSettings();
  const copyText = txt => txt.replace('{REDIRECT}', appAddress());

  el.innerHTML = `
    ${navBar({ title: t.title, back: true, backLabel: 'Help' })}
    <main class="content">
      ${t.kind ? `<div class="eyebrow">${esc(t.kind)}</div>` : ''}
      <h1 class="large-title">${esc(t.title)}</h1>
      <p class="lede">${esc(t.summary)}</p>

      ${t.facts ? `<div class="group">${t.facts.map(([k, v]) => `
        <div class="row row-stack"><div class="fact-label">${esc(k)}</div><div class="fact">${v}</div></div>`).join('')}</div>` : ''}

      ${t.table ? helpTable(t.table) : ''}
      ${t.table2 ? helpTable(t.table2) : ''}
      ${t.after ? `<div class="prose">${t.after}</div>` : ''}

      ${t.steps ? `<div class="group-header">${t.fields ? 'How to set it up' : 'Steps'}</div>
      ${t.steps.map((st, i) => `<div class="step-card">
        <div class="h"><div class="n">${i + 1}</div>${esc(st.title)}</div>
        <p>${st.body}</p>
        ${st.url || st.copy || st.link || st.action ? `<div class="btn-row">
          ${st.url ? `<a class="btn btn-secondary" href="${esc(st.url)}" target="_blank" rel="noopener">${I.share}<span>${esc(st.urlLabel || 'Open website')}</span></a>` : ''}
          ${st.copy ? `<button class="btn btn-secondary" data-copy="${i}">${I.copy}<span>${esc(st.copyLabel || 'Copy')}</span></button>` : ''}
          ${st.link ? `<button class="btn btn-secondary" data-link="${st.link.topic}"><span>${esc(st.link.label)}</span>${I.chev.replace('<svg', '<svg style="width:12px;height:12px"')}</button>` : ''}
          ${st.action === 'settings' ? `<button class="btn btn-secondary" data-opensettings>${I.gear}<span>Open Settings</span></button>` : ''}
        </div>` : ''}
      </div>`).join('')}` : ''}

      ${t.fields ? `
      <div class="group-header">Your details</div>
      <div class="group">${t.fields.map(f => `
        <div class="row"><div class="row-title" style="flex:none">${esc(f.label)}</div>
        <input class="inline" data-k="${f.key}" type="${f.type || 'text'}" value="${esc(s[f.key] || '')}" placeholder="${esc(f.placeholder || '')}" autocomplete="off" autocapitalize="off" spellcheck="false"></div>`).join('')}
      </div>
      <div class="group-footer">Saved on this phone as you type.</div>
      ${t.check ? `<button class="btn btn-secondary mt" data-check>${I.check}<span>Check</span></button><div data-result></div>` : ''}
      ${t.use ? `<button class="btn btn-primary mt" data-use></button>` : ''}
      ` : ''}
      <div style="height:24px"></div>
    </main>`;
  bindBack(el);

  $$('[data-k]', el).forEach(inp => inp.addEventListener('input', () => {
    saveSettings({ [inp.dataset.k]: inp.value.trim() });
    const r = $('[data-result]', el); if (r) r.innerHTML = '';
    drawUse();
  }));
  $$('[data-copy]', el).forEach(b => b.addEventListener('click', () => copy(copyText(t.steps[+b.dataset.copy].copy), 'Copied: paste it into an email')));
  $$('[data-link]', el).forEach(b => b.addEventListener('click', () => push(x => helpTopicScreen(x, b.dataset.link))));
  $$('[data-opensettings]', el).forEach(b => b.addEventListener('click', () => push(settingsScreen)));

  const check = $('[data-check]', el);
  check?.addEventListener('click', async () => {
    const r = $('[data-result]', el);
    check.disabled = true;
    r.innerHTML = `<div class="notice info">${I.info}<div><div class="s">Checking…</div></div></div>`;
    try {
      const msg = await checkKey(t.check);
      r.innerHTML = `<div class="notice ok">${I.okc}<div><div class="t">It works</div><div class="s">${esc(msg)}</div></div></div>`;
    } catch (e) {
      r.innerHTML = `<div class="notice err">${I.warn}<div><div class="t">Not working yet</div><div class="err-text">${esc(e.message || e)}</div></div></div>`;
    } finally {
      check.disabled = false;
    }
  });

  function drawUse() {
    const b = $('[data-use]', el);
    if (!b) return;
    const cur = getSettings();
    const inUse = cur[t.use.setting] === t.use.value;
    const filled = t.fields.every(f => cur[f.key]);
    b.innerHTML = inUse ? `${I.check}<span>In use for ${esc(t.kind.toLowerCase())}</span>` : `<span>${esc(t.use.label)}</span>`;
    b.classList.toggle('is-on', inUse);
    b.disabled = !inUse && !filled;
  }
  $('[data-use]', el)?.addEventListener('click', () => {
    saveSettings({ [t.use.setting]: t.use.value });
    drawUse();
    toast(`${t.title.replace(/ \(.*\)$/, '')} is now used for ${t.kind.toLowerCase()}.`);
  });
  drawUse();
}

// =====================================================================
// Start-up
// =====================================================================

async function boot() {
  if ('serviceWorker' in navigator) {
    try { await navigator.serviceWorker.register('./sw.js'); } catch {}
  }

  history.replaceState({ depth: 1 }, '');
  resetTo(homeScreen);

  // Finishing a Microsoft sign-in?
  try {
    const r = await MS.handleRedirect();
    if (r?.done) {
      toast('Signed in to Microsoft 365');
      const resumeId = sessionStorage.getItem('minutes.resume');
      sessionStorage.removeItem('minutes.resume');
      if (resumeId) {
        const job = await DB.get('jobs', resumeId);
        if (job) push(e => savingScreen(e, job));
      } else {
        getFolders({ refresh: true }).catch(() => {});
        push(settingsScreen);
      }
      return;
    }
  } catch (e) {
    toast(e.message);
  }

  // Arrived from the Android Share sheet?
  const params = new URLSearchParams(location.search);
  if (params.has('shared')) {
    history.replaceState({ depth: 1 }, '', location.pathname);
    const pending = await DB.get('kv', 'pendingShare');
    await DB.del('kv', 'pendingShare');
    if (pending?.file) startJob(pending.file);
    else if (params.get('shared') === 'none') toast('That share didn’t include an audio file.');
  }
}

boot();

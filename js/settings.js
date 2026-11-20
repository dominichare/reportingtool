// Settings live only on this phone (localStorage). Keys never leave the device
// except when sent directly to the provider you chose.

const KEY = 'minutes.settings.v1';

export const DEFAULTS = {
  myName: '',
  myEmail: '',

  // Transcription
  sttProvider: 'assemblyai', // assemblyai | openai | azure
  assemblyaiKey: '',
  openaiKey: '', // shared by OpenAI transcription + OpenAI note writing
  azureSpeechKey: '',
  azureSpeechRegion: 'uksouth',
  language: 'en-GB',

  // Note writing
  llmProvider: 'anthropic', // anthropic | openai | azureopenai
  anthropicKey: '',
  anthropicModel: 'claude-opus-5-5',
  openaiModel: 'gpt-5',
  azureOpenaiEndpoint: '',
  azureOpenaiKey: '',
  azureOpenaiDeployment: '',

  // Microsoft 365
  msClientId: '',
  msTenant: 'organizations',

  // Folders
  oneOnOnesFolder: 'One-on-Ones',
  subjectPrefixes: '1,2,3',
  manualFolders: '', // newline separated, used when not signed in to Microsoft

  // SharePoint (remembered)
  spSite: null, // { id, name, webUrl }
  spFolder: null, // { id, path }
  spLastOn: false,
};

let cache = null;

export function getSettings() {
  if (!cache) {
    let stored = {};
    try { stored = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch {}
    cache = { ...DEFAULTS, ...stored };
  }
  return cache;
}

export function saveSettings(patch) {
  cache = { ...getSettings(), ...patch };
  try { localStorage.setItem(KEY, JSON.stringify(cache)); } catch {}
  return cache;
}

export function sttReady(s = getSettings()) {
  if (s.sttProvider === 'assemblyai') return !!s.assemblyaiKey;
  if (s.sttProvider === 'openai') return !!s.openaiKey;
  if (s.sttProvider === 'azure') return !!(s.azureSpeechKey && s.azureSpeechRegion);
  return false;
}

export function llmReady(s = getSettings()) {
  if (s.llmProvider === 'anthropic') return !!s.anthropicKey;
  if (s.llmProvider === 'openai') return !!s.openaiKey;
  if (s.llmProvider === 'azureopenai') return !!(s.azureOpenaiEndpoint && s.azureOpenaiKey && s.azureOpenaiDeployment);
  return false;
}

// ---------- People (remembered speakers) ----------
const PEOPLE_KEY = 'minutes.people.v1';

export function getPeople() {
  try { return JSON.parse(localStorage.getItem(PEOPLE_KEY) || '[]'); } catch { return []; }
}

export function savePeople(list) {
  try { localStorage.setItem(PEOPLE_KEY, JSON.stringify(list)); } catch {}
}

export function rememberPerson(name, email) {
  name = (name || '').trim();
  if (!name) return;
  const list = getPeople();
  const found = list.find(p => p.name.toLowerCase() === name.toLowerCase());
  if (found) {
    found.count = (found.count || 0) + 1;
    found.last = Date.now();
    if (email) found.email = email;
  } else {
    list.push({ name, email: email || '', count: 1, last: Date.now() });
  }
  savePeople(list);
}

export function findPerson(name) {
  name = (name || '').trim().toLowerCase();
  return getPeople().find(p => p.name.toLowerCase() === name);
}

export function topPeople(n = 8) {
  return getPeople()
    .sort((a, b) => (b.count - a.count) || (b.last - a.last))
    .slice(0, n);
}

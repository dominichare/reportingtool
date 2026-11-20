// Switchable AI providers.
//   Transcription (speech → text with speaker separation): AssemblyAI, OpenAI, Azure AI Speech
//   Note writing (transcript → structured meeting note):   Claude, OpenAI, Azure OpenAI
// Everything is called directly from the phone with your own API keys.

import { getSettings } from './settings.js';

export const STT_PROVIDERS = {
  assemblyai: { label: 'AssemblyAI', hint: 'Best at telling speakers apart. Handles long recordings.' },
  openai: { label: 'OpenAI', hint: 'Simple, one key. Files up to 25 MB (about 25–50 minutes).' },
  azure: { label: 'Azure Speech', hint: 'Stays inside your company’s Microsoft cloud. Needs IT to set up.' },
};

export const LLM_PROVIDERS = {
  anthropic: { label: 'Claude', hint: 'Anthropic. Excellent, well-structured notes.' },
  openai: { label: 'OpenAI', hint: 'GPT via the OpenAI API. Uses the same key as OpenAI transcription.' },
  azureopenai: { label: 'Azure OpenAI', hint: 'Stays inside your company’s Microsoft cloud. Needs IT to set up.' },
};

class ProviderError extends Error {}

async function call(provider, url, opts) {
  let res;
  try {
    res = await fetch(url, opts);
  } catch (e) {
    throw new ProviderError(`Couldn’t reach ${provider}. Check your internet connection and try again.`);
  }
  if (!res.ok) {
    let detail = '';
    try {
      const j = await res.json();
      detail = j.error?.message || j.error || j.message || JSON.stringify(j);
    } catch { try { detail = await res.text(); } catch {} }
    if (res.status === 401 || res.status === 403) {
      throw new ProviderError(`${provider} didn’t accept the API key. Check it in Settings.${detail ? `\n\n${detail}` : ''}`);
    }
    if (res.status === 413) throw new ProviderError(`This recording is too large for ${provider}. Try another transcription service in Settings.`);
    if (res.status === 429) throw new ProviderError(`${provider} is busy or your account has run out of credit. Wait a minute and try again.${detail ? `\n\n${detail}` : ''}`);
    throw new ProviderError(`${provider} returned an error (${res.status}).${detail ? `\n\n${typeof detail === 'string' ? detail : JSON.stringify(detail)}` : ''}`);
  }
  return res;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Merge consecutive fragments by the same speaker into one utterance.
function mergeUtterances(list) {
  const out = [];
  for (const u of list) {
    const text = (u.text || '').trim();
    if (!text) continue;
    const prev = out[out.length - 1];
    if (prev && prev.speaker === u.speaker && u.start - prev.end < 2500) {
      prev.text += ' ' + text;
      prev.end = u.end;
    } else {
      out.push({ speaker: String(u.speaker), start: Math.round(u.start), end: Math.round(u.end), text });
    }
  }
  return out;
}

// ---------------- Transcription ----------------

async function transcribeAssemblyAI(file, s, onStatus) {
  const P = 'AssemblyAI';
  const auth = { authorization: s.assemblyaiKey };
  onStatus('Uploading recording…');
  const up = await (await call(P, 'https://api.assemblyai.com/v2/upload', {
    method: 'POST', headers: { ...auth, 'content-type': 'application/octet-stream' }, body: file,
  })).json();

  onStatus('Transcribing…');
  const lang = (s.language || 'en-GB').toLowerCase();
  const language_code = lang === 'en-gb' ? 'en_uk' : lang === 'en-us' ? 'en_us' : lang.split('-')[0];
  const job = await (await call(P, 'https://api.assemblyai.com/v2/transcript', {
    method: 'POST',
    headers: { ...auth, 'content-type': 'application/json' },
    body: JSON.stringify({ audio_url: up.upload_url, speaker_labels: true, language_code, punctuate: true, format_text: true }),
  })).json();

  const started = Date.now();
  for (;;) {
    await sleep(3000);
    const r = await (await call(P, `https://api.assemblyai.com/v2/transcript/${job.id}`, { headers: auth })).json();
    if (r.status === 'completed') {
      return {
        durationMs: Math.round((r.audio_duration || 0) * 1000),
        utterances: mergeUtterances((r.utterances || []).map(u => ({ speaker: u.speaker, start: u.start, end: u.end, text: u.text }))),
      };
    }
    if (r.status === 'error') throw new ProviderError(`AssemblyAI couldn’t transcribe this recording.\n\n${r.error || ''}`);
    const secs = Math.round((Date.now() - started) / 1000);
    onStatus(`Transcribing… ${secs}s`);
  }
}

async function transcribeOpenAI(file, s, onStatus) {
  const P = 'OpenAI';
  if (file.size > 25 * 1024 * 1024) {
    throw new ProviderError(`This recording is ${(file.size / 1048576).toFixed(0)} MB. OpenAI accepts up to 25 MB. Switch transcription to AssemblyAI or Azure Speech in Settings.`);
  }
  onStatus('Uploading and transcribing…');
  const fd = new FormData();
  fd.append('file', file, file.name || 'recording.m4a');
  fd.append('model', 'gpt-4o-transcribe-diarize');
  fd.append('response_format', 'diarized_json');
  fd.append('chunking_strategy', 'auto');
  const lang = (s.language || 'en-GB').split('-')[0];
  fd.append('language', lang);
  const r = await (await call(P, 'https://api.openai.com/v1/audio/transcriptions', {
    method: 'POST', headers: { authorization: `Bearer ${s.openaiKey}` }, body: fd,
  })).json();
  const segs = r.segments || [];
  return {
    durationMs: Math.round((r.duration || (segs.length ? segs[segs.length - 1].end : 0)) * 1000),
    utterances: mergeUtterances(segs.map(x => ({ speaker: x.speaker, start: x.start * 1000, end: x.end * 1000, text: x.text }))),
  };
}

async function transcribeAzure(file, s, onStatus) {
  const P = 'Azure Speech';
  onStatus('Uploading and transcribing…');
  const fd = new FormData();
  fd.append('audio', file, file.name || 'recording.m4a');
  fd.append('definition', JSON.stringify({
    locales: [s.language || 'en-GB'],
    diarization: { enabled: true, maxSpeakers: 8 },
  }));
  const url = `https://${encodeURIComponent(s.azureSpeechRegion)}.api.cognitive.microsoft.com/speechtotext/transcriptions:transcribe?api-version=2024-11-15`;
  const r = await (await call(P, url, {
    method: 'POST', headers: { 'Ocp-Apim-Subscription-Key': s.azureSpeechKey }, body: fd,
  })).json();
  const phrases = r.phrases || [];
  return {
    durationMs: r.durationMilliseconds || 0,
    utterances: mergeUtterances(phrases.map(p => ({
      speaker: p.speaker ?? 0,
      start: p.offsetMilliseconds,
      end: p.offsetMilliseconds + p.durationMilliseconds,
      text: p.text,
    }))),
  };
}

export async function transcribe(file, onStatus = () => {}) {
  const s = getSettings();
  const fn = { assemblyai: transcribeAssemblyAI, openai: transcribeOpenAI, azure: transcribeAzure }[s.sttProvider];
  if (!fn) throw new ProviderError('Choose a transcription service in Settings.');
  const result = await fn(file, s, onStatus);
  if (!result.utterances.length) throw new ProviderError('No speech was found in this recording.');
  return result;
}

// ---------------- Note writing ----------------

const FALLBACK_MODELS = new Set(['claude-opus-5-5', 'claude-opus-5', 'claude-sonnet-5-5', 'claude-fable-5-2', 'claude-fable-5-1']);

async function writeWithClaude({ system, user, schema }, s, onStatus) {
  onStatus('Writing the note with Claude…');
  const { default: Anthropic } = await import('../vendor/anthropic-sdk-1.6.0.mjs');
  const client = new Anthropic({ apiKey: s.anthropicKey, dangerouslyAllowBrowser: true, maxRetries: 2 });
  const model = s.anthropicModel || 'claude-opus-5-5';

  const base = {
    model,
    max_tokens: 32000,
    system,
    messages: [{ role: 'user', content: user }],
    output_config: { effort: 'medium', format: { type: 'json_schema', schema } },
  };
  // On a safety decline, let the API retry on a suitable model automatically.
  const withFallback = FALLBACK_MODELS.has(model)
    ? { ...base, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' }
    : base;

  let msg;
  try {
    msg = await client.beta.messages.stream(withFallback).finalMessage();
  } catch (e) {
    if (e instanceof Anthropic.AuthenticationError || e instanceof Anthropic.PermissionDeniedError) {
      throw new ProviderError('Claude didn’t accept the API key. Check it in Settings.');
    }
    if (e instanceof Anthropic.RateLimitError) {
      throw new ProviderError('Claude is busy or your account has run out of credit. Wait a minute and try again.');
    }
    if (e instanceof Anthropic.BadRequestError && withFallback !== base && /fallback/i.test(e.message)) {
      msg = await client.beta.messages.stream(base).finalMessage();
    } else if (e instanceof Anthropic.APIConnectionError) {
      throw new ProviderError('Couldn’t reach Claude. Check your internet connection and try again.');
    } else if (e instanceof Anthropic.APIError) {
      throw new ProviderError(`Claude returned an error (${e.status ?? 'unknown'}).\n\n${e.message}`);
    } else {
      throw e;
    }
  }

  if (msg.stop_reason === 'refusal') {
    throw new ProviderError('Claude declined to write a note for this transcript.');
  }
  if (msg.stop_reason === 'max_tokens') {
    throw new ProviderError('The note was cut off because it was too long. Try again.');
  }
  const text = msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
  return JSON.parse(text);
}

async function writeWithOpenAICompatible({ system, user, schema }, url, headers, model, P, onStatus) {
  onStatus(`Writing the note with ${P}…`);
  const body = {
    messages: [{ role: 'system', content: system }, { role: 'user', content: user }],
    response_format: { type: 'json_schema', json_schema: { name: 'meeting_note', strict: true, schema } },
  };
  if (model) body.model = model;
  const r = await (await call(P, url, {
    method: 'POST', headers: { ...headers, 'content-type': 'application/json' }, body: JSON.stringify(body),
  })).json();
  const choice = r.choices?.[0];
  if (choice?.message?.refusal) throw new ProviderError(`${P} declined to write a note for this transcript.`);
  if (choice?.finish_reason === 'length') throw new ProviderError('The note was cut off because it was too long. Try again.');
  return JSON.parse(choice.message.content);
}

export async function writeNote(prompt, onStatus = () => {}) {
  const s = getSettings();
  if (s.llmProvider === 'anthropic') return writeWithClaude(prompt, s, onStatus);
  if (s.llmProvider === 'openai') {
    return writeWithOpenAICompatible(prompt, 'https://api.openai.com/v1/chat/completions',
      { authorization: `Bearer ${s.openaiKey}` }, s.openaiModel || 'gpt-5', 'OpenAI', onStatus);
  }
  if (s.llmProvider === 'azureopenai') {
    const ep = s.azureOpenaiEndpoint.replace(/\/+$/, '');
    const url = `${ep}/openai/deployments/${encodeURIComponent(s.azureOpenaiDeployment)}/chat/completions?api-version=2024-10-21`;
    return writeWithOpenAICompatible(prompt, url, { 'api-key': s.azureOpenaiKey }, null, 'Azure OpenAI', onStatus);
  }
  throw new ProviderError('Choose a note-writing service in Settings.');
}

export { ProviderError };

// ---------------- Key checks (free or near-free) ----------------
// Returns a short success message, or throws a ProviderError with a friendly explanation.
export async function checkKey(kind) {
  const s = getSettings();
  const need = (v, what) => { if (!v) throw new ProviderError(`Enter the ${what} first.`); };
  switch (kind) {
    case 'assemblyai': {
      need(s.assemblyaiKey, 'API key');
      await call('AssemblyAI', 'https://api.assemblyai.com/v2/transcript?limit=1', { headers: { authorization: s.assemblyaiKey } });
      return 'AssemblyAI accepted the key.';
    }
    case 'openai': {
      need(s.openaiKey, 'API key');
      await call('OpenAI', 'https://api.openai.com/v1/models/gpt-4o-transcribe-diarize', { headers: { authorization: `Bearer ${s.openaiKey}` } });
      return 'OpenAI accepted the key, and the transcription model is available.';
    }
    case 'openai-llm': {
      need(s.openaiKey, 'API key');
      const model = s.openaiModel || 'gpt-5';
      const res = await fetch(`https://api.openai.com/v1/models/${encodeURIComponent(model)}`, { headers: { authorization: `Bearer ${s.openaiKey}` } }).catch(() => null);
      if (!res) throw new ProviderError('Couldn’t reach OpenAI. Check your internet connection.');
      if (res.status === 404) throw new ProviderError(`The key works, but the model “${model}” isn’t available to your account. Try another model name, e.g. gpt-5-mini.`);
      if (res.status === 401 || res.status === 403) throw new ProviderError('OpenAI didn’t accept the API key.');
      if (!res.ok) throw new ProviderError(`OpenAI returned an error (${res.status}).`);
      return `OpenAI accepted the key, and “${model}” is available.`;
    }
    case 'azure': {
      need(s.azureSpeechKey, 'key');
      need(s.azureSpeechRegion, 'region');
      const url = `https://${encodeURIComponent(s.azureSpeechRegion.trim().toLowerCase().replace(/\s+/g, ''))}.api.cognitive.microsoft.com/sts/v1.0/issueToken`;
      await call('Azure Speech', url, { method: 'POST', headers: { 'Ocp-Apim-Subscription-Key': s.azureSpeechKey } });
      return 'Azure Speech accepted the key for this region.';
    }
    case 'anthropic': {
      need(s.anthropicKey, 'API key');
      const { default: Anthropic } = await import('../vendor/anthropic-sdk-1.6.0.mjs');
      const client = new Anthropic({ apiKey: s.anthropicKey, dangerouslyAllowBrowser: true, maxRetries: 1 });
      const model = s.anthropicModel || 'claude-opus-5-5';
      try {
        await client.models.retrieve(model);
      } catch (e) {
        if (e instanceof Anthropic.NotFoundError) throw new ProviderError(`The key works, but the model “${model}” wasn’t found. Use claude-opus-5-5.`);
        if (e instanceof Anthropic.AuthenticationError || e instanceof Anthropic.PermissionDeniedError) throw new ProviderError('Claude didn’t accept the API key.');
        if (e instanceof Anthropic.APIConnectionError) throw new ProviderError('Couldn’t reach Claude. Check your internet connection.');
        if (e instanceof Anthropic.APIError) throw new ProviderError(`Claude returned an error (${e.status ?? 'unknown'}).`);
        throw e;
      }
      return `Claude accepted the key, and “${model}” is available.`;
    }
    case 'azureopenai': {
      need(s.azureOpenaiEndpoint, 'endpoint');
      need(s.azureOpenaiDeployment, 'deployment name');
      need(s.azureOpenaiKey, 'key');
      const ep = s.azureOpenaiEndpoint.replace(/\/+$/, '');
      const url = `${ep}/openai/deployments/${encodeURIComponent(s.azureOpenaiDeployment)}/chat/completions?api-version=2024-10-21`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'api-key': s.azureOpenaiKey, 'content-type': 'application/json' },
        body: JSON.stringify({ messages: [{ role: 'user', content: 'Reply with the single word OK.' }] }),
      }).catch(() => null);
      if (!res) throw new ProviderError('Couldn’t reach that endpoint. Check the endpoint address and your internet connection.');
      if (res.status === 404) throw new ProviderError('The endpoint was reached, but that deployment name wasn’t found. Check the deployment name with IT.');
      if (res.status === 401 || res.status === 403) throw new ProviderError('Azure OpenAI didn’t accept the key.');
      if (!res.ok) throw new ProviderError(`Azure OpenAI returned an error (${res.status}).`);
      return 'Azure OpenAI is working with this deployment.';
    }
  }
  throw new ProviderError('Nothing to check.');
}

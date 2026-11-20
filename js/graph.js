// Microsoft 365: sign-in (OAuth 2.0 auth-code + PKCE, as a single-page app) and
// Microsoft Graph calls for OneDrive, SharePoint, Outlook drafts and People search.
// Optional: if no Client ID is configured, the app uses Android's Share sheet instead.

import { getSettings } from './settings.js';

const TOKEN_KEY = 'minutes.mstoken.v1';
const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All Mail.ReadWrite People.Read';
const GRAPH = 'https://graph.microsoft.com/v1.0';

const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randomString = (n = 48) => b64url(crypto.getRandomValues(new Uint8Array(n)));

function redirectUri() {
  return location.origin + location.pathname.replace(/index\.html$/, '');
}

function authority() {
  const s = getSettings();
  return `https://login.microsoftonline.com/${encodeURIComponent(s.msTenant || 'organizations')}/oauth2/v2.0`;
}

export function msConfigured() {
  return !!getSettings().msClientId;
}

function loadToken() {
  try { return JSON.parse(localStorage.getItem(TOKEN_KEY) || 'null'); } catch { return null; }
}
function saveToken(t) {
  if (t) localStorage.setItem(TOKEN_KEY, JSON.stringify(t)); else localStorage.removeItem(TOKEN_KEY);
}

export function msAccount() {
  const t = loadToken();
  return t ? { name: t.name, email: t.email } : null;
}

export function msSignedIn() {
  return msConfigured() && !!loadToken();
}

export async function signIn() {
  const s = getSettings();
  if (!s.msClientId) throw new Error('Add your Microsoft app Client ID in Settings first.');
  const verifier = randomString(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = randomString(16);
  sessionStorage.setItem('minutes.pkce', JSON.stringify({ verifier, state, back: location.hash || '' }));
  const q = new URLSearchParams({
    client_id: s.msClientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'query',
    scope: SCOPES,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    state,
    prompt: 'select_account',
  });
  location.assign(`${authority()}/authorize?${q}`);
}

export function signOut() {
  saveToken(null);
}

// Call on start-up: completes a sign-in redirect if one is in progress.
// Returns { done: true, back } when a sign-in just finished.
export async function handleRedirect() {
  const p = new URLSearchParams(location.search);
  if (!p.has('code') && !p.has('error')) return null;
  const saved = JSON.parse(sessionStorage.getItem('minutes.pkce') || 'null');
  sessionStorage.removeItem('minutes.pkce');
  history.replaceState({ depth: 1 }, '', redirectUri() + (saved?.back || ''));
  if (p.has('error')) {
    const desc = p.get('error_description') || p.get('error');
    throw new Error(/AADSTS65001|AADSTS90094|admin/i.test(desc)
      ? 'Your organisation needs an administrator to approve this app before you can sign in. The Share-sheet route still works in the meantime.'
      : `Microsoft sign-in failed: ${desc}`);
  }
  if (!saved || saved.state !== p.get('state')) throw new Error('Sign-in was interrupted. Please try again.');
  await tokenRequest({
    grant_type: 'authorization_code',
    code: p.get('code'),
    redirect_uri: redirectUri(),
    code_verifier: saved.verifier,
  });
  try {
    const me = await graph('/me?$select=displayName,mail,userPrincipalName');
    const t = loadToken();
    saveToken({ ...t, name: me.displayName, email: me.mail || me.userPrincipalName });
  } catch {}
  return { done: true, back: saved.back };
}

async function tokenRequest(params) {
  const s = getSettings();
  const body = new URLSearchParams({ client_id: s.msClientId, scope: SCOPES, ...params });
  const res = await fetch(`${authority()}/token`, { method: 'POST', body });
  const j = await res.json();
  if (!res.ok) {
    if (params.grant_type === 'refresh_token') { saveToken(null); throw new SignInNeeded(); }
    throw new Error(j.error_description || 'Microsoft sign-in failed.');
  }
  const prev = loadToken() || {};
  saveToken({
    ...prev,
    access: j.access_token,
    refresh: j.refresh_token || prev.refresh,
    expires: Date.now() + (j.expires_in - 120) * 1000,
  });
}

export class SignInNeeded extends Error {
  constructor() { super('Your Microsoft sign-in has expired. Sign in again to continue.'); }
}

async function accessToken() {
  const t = loadToken();
  if (!t) throw new SignInNeeded();
  if (t.access && Date.now() < t.expires) return t.access;
  if (!t.refresh) { saveToken(null); throw new SignInNeeded(); }
  await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh });
  return loadToken().access;
}

export async function graph(path, { method = 'GET', body, headers = {}, raw = false } = {}) {
  const token = await accessToken();
  const isBlob = body instanceof Blob;
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    method,
    headers: {
      authorization: `Bearer ${token}`,
      ...(body && !isBlob ? { 'content-type': 'application/json' } : {}),
      ...headers,
    },
    body: body ? (isBlob ? body : JSON.stringify(body)) : undefined,
  });
  if (res.status === 401) { saveToken({ ...loadToken(), expires: 0 }); throw new SignInNeeded(); }
  if (!res.ok) {
    let msg = `Microsoft 365 returned an error (${res.status}).`;
    try { const j = await res.json(); msg += `\n\n${j.error?.message || ''}`; } catch {}
    if (res.status === 403) msg = 'Microsoft 365 refused this request. Your organisation may need to approve extra permissions for this app.';
    throw new Error(msg);
  }
  if (raw) return res;
  return res.status === 204 ? null : res.json();
}

const encPath = p => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

// ---------- OneDrive ----------

export async function listFolders() {
  const s = getSettings();
  const prefixes = s.subjectPrefixes.split(',').map(x => x.trim()).filter(Boolean);
  const root = await graph('/me/drive/root/children?$select=name,folder&$top=999');
  const subject = (root.value || [])
    .filter(i => i.folder && prefixes.some(p => i.name.startsWith(p)))
    .map(i => i.name)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  let oneOnOnes = [];
  try {
    const r = await graph(`/me/drive/root:/${encPath(s.oneOnOnesFolder)}:/children?$select=name,folder&$top=999`);
    oneOnOnes = (r.value || []).filter(i => i.folder).map(i => i.name).sort();
  } catch {}
  return { subject, oneOnOnes };
}

export async function uploadToOneDrive(folderPath, fileName, blob) {
  // Simple upload creates any missing folders in the path (e.g. a new person's one-on-one folder).
  const url = `/me/drive/root:/${encPath(folderPath)}/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=rename`;
  return graph(url, { method: 'PUT', body: blob, headers: { 'content-type': blob.type } });
}

// ---------- SharePoint ----------

export async function followedSites() {
  const r = await graph('/me/followedSites?$select=id,displayName,webUrl');
  return (r.value || []).map(x => ({ id: x.id, name: x.displayName, webUrl: x.webUrl }));
}

export async function searchSites(q) {
  const r = await graph(`/sites?search=${encodeURIComponent(q)}&$select=id,displayName,webUrl&$top=25`);
  return (r.value || []).map(x => ({ id: x.id, name: x.displayName, webUrl: x.webUrl }));
}

export async function siteFolderChildren(siteId, folderId) {
  const base = `/sites/${siteId}/drive`;
  const r = await graph(`${base}/${folderId ? `items/${folderId}` : 'root'}/children?$select=id,name,folder&$top=999`);
  return (r.value || []).filter(i => i.folder).map(i => ({ id: i.id, name: i.name }))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
}

export async function uploadToSharePoint(siteId, folderId, fileName, blob) {
  const target = folderId ? `items/${folderId}` : 'root';
  const url = `/sites/${siteId}/drive/${target}:/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=rename`;
  return graph(url, { method: 'PUT', body: blob, headers: { 'content-type': blob.type } });
}

// ---------- Outlook ----------

async function blobToBase64(blob) {
  const buf = new Uint8Array(await blob.arrayBuffer());
  let s = '';
  for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode(...buf.subarray(i, i + 0x8000));
  return btoa(s);
}

export async function createDraft({ to, subject, html, attachments }) {
  const msg = await graph('/me/messages', {
    method: 'POST',
    body: {
      subject,
      body: { contentType: 'HTML', content: html },
      toRecipients: to.map(r => ({ emailAddress: { address: r.email, name: r.name } })),
    },
  });
  for (const a of attachments) {
    await graph(`/me/messages/${msg.id}/attachments`, {
      method: 'POST',
      body: {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: a.name,
        contentType: a.blob.type,
        contentBytes: await blobToBase64(a.blob),
      },
    });
  }
  return msg; // includes webLink
}

// ---------- People ----------

export async function searchPeople(q) {
  const r = await graph(`/me/people?$search="${encodeURIComponent(q)}"&$top=6&$select=displayName,scoredEmailAddresses`);
  return (r.value || [])
    .map(p => ({ name: p.displayName, email: p.scoredEmailAddresses?.[0]?.address || '' }))
    .filter(p => p.name);
}

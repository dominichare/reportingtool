// Meeting-note template: the AI instructions, the output schema, and the Word layouts.

import { buildDocx } from './docx.js';

export const NOTE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['title', 'meeting_type', 'one_on_one_with', 'participants', 'key_decisions', 'narrative', 'actions', 'suggested_folder', 'folder_reason'],
  properties: {
    title: { type: 'string', description: 'Clear, specific meeting title, 3–8 words, no date.' },
    meeting_type: { type: 'string', enum: ['one_on_one', 'group'] },
    one_on_one_with: { type: 'string', description: 'For a one-on-one: the other person’s full name (Firstname Surname). Otherwise empty string.' },
    participants: { type: 'array', items: { type: 'string' } },
    key_decisions: { type: 'array', items: { type: 'string' }, description: 'Headline bullets of what was decided or agreed.' },
    narrative: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['heading', 'summary'],
        properties: {
          heading: { type: 'string' },
          summary: { type: 'string' },
        },
      },
    },
    actions: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['action', 'owner', 'due'],
        properties: {
          action: { type: 'string' },
          owner: { type: 'string' },
          due: { type: 'string', description: 'Due date or timeframe if stated, otherwise empty string.' },
        },
      },
    },
    suggested_folder: { type: 'string', description: 'Exactly one of the candidate folder paths provided.' },
    folder_reason: { type: 'string', description: 'One short sentence explaining the folder choice.' },
  },
};

const SYSTEM = `You turn meeting transcripts into concise, professional meeting notes written in British English.

Write for a busy senior reader who was not in the meeting:
- Title: clear and specific (e.g. "Q3 Budget Review", "Sarah Jones – Monthly Check-in"). No date in the title.
- Key decisions: short headline bullets of what was decided, agreed or concluded. If nothing was formally decided, give the main conclusions. Do not repeat actions here.
- Narrative: a concise, point-by-point account of what happened, in the order discussed. Each point has a short heading and 1–3 sentences. Attribute views to named people where it matters. Group related discussion; skip small talk and filler.
- Actions: every concrete follow-up, each with a named owner (use "Unassigned" if unclear) and a due date or timeframe only if one was stated.
- Use the speaker names exactly as given. Do not invent facts, names, numbers or dates.
- The transcript is machine-generated, so silently correct obvious mis-hearings when the meaning is clear.

Meeting type: "one_on_one" when the meeting is essentially between the note-taker and one other person; otherwise "group".

Folder: choose suggested_folder from the candidate list provided.
- For a one-on-one, choose "<one-on-ones folder>/<Firstname Surname>" for the other person, even if that subfolder is not yet listed.
- Otherwise choose the subject folder whose name best matches the topic. Folders starting with 1 are the highest priority projects, then 2, then 3. If nothing fits well, choose the closest match and say so in folder_reason.`;

export function fmtTime(ms) {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  const mm = String(m).padStart(h ? 2 : 1, '0');
  return h ? `${h}:${mm}:${String(sec).padStart(2, '0')}` : `${mm}:${String(sec).padStart(2, '0')}`;
}

export function namedTranscript(job) {
  return job.utterances.map(u => ({ ...u, name: job.speakerNames[u.speaker] || `Speaker ${u.speaker}` }))
    .filter(u => u.name !== '__ignore__');
}

export function participantNames(job) {
  const names = [];
  for (const u of namedTranscript(job)) if (!names.includes(u.name)) names.push(u.name);
  for (const extra of job.extraParticipants || []) if (extra && !names.includes(extra)) names.push(extra);
  return names;
}

export function buildPrompt(job, { myName, oneOnOnesFolder, folders, instruction, previous }) {
  const transcript = namedTranscript(job).map(u => `[${fmtTime(u.start)}] ${u.name}: ${u.text}`).join('\n');
  const people = participantNames(job);
  const candidates = [
    ...folders.subject,
    ...folders.oneOnOnes.map(n => `${oneOnOnesFolder}/${n}`),
  ];
  const parts = [
    `Note-taker (refer to them by name, not "me"): ${myName || 'the note-taker'}`,
    `Meeting date: ${longDate(job.meetingDate)}`,
    `Recording file: ${job.fileName}`,
    `Participants (speakers named by the note-taker): ${people.join(', ')}`,
    `One-on-ones folder: ${oneOnOnesFolder}`,
    `Candidate folders:\n${candidates.map(c => `- ${c}`).join('\n') || '- (none listed)'}`,
    `<transcript>\n${transcript}\n</transcript>`,
  ];
  if (previous && instruction) {
    parts.push(`Here is the current version of the note as JSON:\n${JSON.stringify(previous)}\n\nRevise it according to this request from the note-taker, keeping everything else the same: ${instruction}`);
  }
  return { system: SYSTEM, user: parts.join('\n\n'), schema: NOTE_SCHEMA };
}

// ---------- Dates & filenames ----------

export function longDate(iso) {
  if (!iso) return '';
  const d = new Date(iso + 'T12:00:00');
  return d.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
}

export function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|#%]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
}

export function noteFileName(job) {
  return `${job.meetingDate} – ${safeName(job.note.title)}.docx`;
}

export function transcriptFileName(job) {
  return `${job.meetingDate} – ${safeName(job.note.title)} – Transcript.docx`;
}

// Try to read the meeting date from the file name (e.g. "2026-11-20 10-31.m4a",
// "Recording_20261120_103100.m4a"), falling back to the file's modified date.
export function guessDate(fileName, lastModified) {
  const n = fileName || '';
  let m = n.match(/(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/);
  if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) return `${m[1]}-${m[2]}-${m[3]}`;
  m = n.match(/(\d{2})[-_.](\d{2})[-_.](20\d{2})/);
  if (m && +m[2] >= 1 && +m[2] <= 12) return `${m[3]}-${m[2]}-${m[1]}`;
  const d = lastModified ? new Date(lastModified) : new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ---------- Word documents ----------

export function noteDocx(job, author) {
  const n = job.note;
  const blocks = [
    { type: 'title', text: n.title },
    { type: 'subtitle', text: n.meeting_type === 'one_on_one' ? 'One-on-one meeting note' : 'Meeting note' },
    { type: 'meta', label: 'Date', value: longDate(job.meetingDate) },
    { type: 'meta', label: 'Participants', value: n.participants.join(', ') },
    { type: 'h1', text: 'Key decisions' },
    ...(n.key_decisions.length ? n.key_decisions.map(d => ({ type: 'bullet', runs: [d] })) : [{ type: 'p', runs: [{ text: 'No decisions recorded.', italic: true }] }]),
    { type: 'h1', text: 'Discussion' },
    ...n.narrative.map(p => ({ type: 'number', runs: [{ text: `${p.heading}. `, bold: true }, p.summary] })),
    { type: 'h1', text: 'Actions' },
  ];
  if (n.actions.length) {
    blocks.push({
      type: 'table',
      header: ['#', 'Action', 'Owner', 'Due'],
      widths: [0.06, 0.56, 0.2, 0.18],
      rows: n.actions.map((a, i) => [String(i + 1), a.action, a.owner, a.due || '—']),
    });
  } else {
    blocks.push({ type: 'p', runs: [{ text: 'No actions recorded.', italic: true }] });
  }
  return buildDocx(blocks, { title: n.title, author });
}

export function transcriptDocx(job, author) {
  const n = job.note;
  const blocks = [
    { type: 'title', text: `${n.title} – Transcript` },
    { type: 'subtitle', text: 'Machine transcription. May contain errors.' },
    { type: 'meta', label: 'Date', value: longDate(job.meetingDate) },
    { type: 'meta', label: 'Participants', value: n.participants.join(', ') },
    { type: 'meta', label: 'Recording', value: job.fileName },
    { type: 'h1', text: 'Transcript' },
    ...namedTranscript(job).map(u => ({
      type: 'p',
      runs: [{ text: `${u.name}  `, bold: true }, { text: `${fmtTime(u.start)}   `, color: '8E8E93' }, u.text],
    })),
  ];
  return buildDocx(blocks, { title: `${n.title} – Transcript`, author });
}

// ---------- Email ----------

export function emailSubject(job) {
  return `Meeting note: ${job.note.title} (${longDate(job.meetingDate)})`;
}

export function emailHtml(job, myName) {
  const e = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const n = job.note;
  const actions = n.actions.length
    ? `<p><b>Actions</b></p><ul>${n.actions.map(a => `<li>${e(a.action)} – <i>${e(a.owner)}</i>${a.due ? ` (${e(a.due)})` : ''}</li>`).join('')}</ul>`
    : '';
  return `<p>Hi all,</p><p>Please find attached the note from <b>${e(n.title)}</b> on ${e(longDate(job.meetingDate))}.</p>${actions}<p>Thanks,<br>${e(myName || '')}</p>`;
}

export function emailText(job, myName) {
  const n = job.note;
  const actions = n.actions.length ? `\n\nActions:\n${n.actions.map(a => `• ${a.action} – ${a.owner}${a.due ? ` (${a.due})` : ''}`).join('\n')}` : '';
  return `Hi all,\n\nPlease find attached the note from ${n.title} on ${longDate(job.meetingDate)}.${actions}\n\nThanks,\n${myName || ''}`;
}

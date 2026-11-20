# Minutes

Turn a meeting recording on your phone into a structured meeting note, then save it to OneDrive (and optionally SharePoint) and create an email draft. It runs entirely on an Android phone as an installable web app.

**Flow:** Easy Recorder → *Share* → **Minutes** → transcribe → name each voice once → AI writes the note → review/edit → choose folder → save & email draft → reminder to delete the recording.

The note has:
- title, date and participants at the top
- **Key decisions** as headline bullets
- **Discussion**, a concise point-by-point narrative
- **Actions** as a table (action, owner, due) at the end

The transcript is saved as a separate Word file next to the note.

---

## 1. Put it online (one-off)

The app is a set of static files. It needs to be served over HTTPS so Android can install it. The simplest host is **GitHub Pages**:

1. Merge this branch into `main`.
2. In the repository: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. The included workflow (`.github/workflows/pages.yml`) publishes the site to
   `https://<your-github-username>.github.io/reportingtool/`.

No keys or personal data are in the code. Keys are entered on the phone and stay on the phone.

## 2. Install on the phone

1. Open the address above in **Chrome** on the phone.
2. Tap **⋮ → Add to Home screen → Install**.
3. Open **Minutes** and tap **Help → Getting started**. It walks you through each step. Every service guide lets you paste your key and tap **Check** to confirm it works.

To process a recording, open it in Easy Recorder (or the My Files app), tap **Share**, and choose **Minutes**. You can also tap **Choose recording** in the app.

## 3. AI services (switchable in Settings)

You need one **transcription** service and one **note-writing** service. You can change either at any time.

| Transcription | Notes |
|---|---|
| **AssemblyAI** *(default)* | Very good at separating speakers and handles long recordings. Key from assemblyai.com. |
| **OpenAI** | One key for both jobs. Limited to 25 MB per file, which is roughly 25–50 minutes depending on recording quality. |
| **Azure AI Speech** | Runs inside your organisation's Microsoft cloud. IT provides the key and region. |

| Note writing | Notes |
|---|---|
| **Claude** *(default)* | Anthropic API key from console.anthropic.com. Model `claude-opus-5-5`. |
| **OpenAI** | Uses the same key as OpenAI transcription. The model is editable. |
| **Azure OpenAI** | Runs inside your organisation's Microsoft cloud. IT provides the endpoint, deployment and key. |

**Cost:** you pay the providers per use. That is typically well under £1 per hour of recording for transcription and note writing combined.

**Privacy:** audio and transcripts go only to the services you choose. Check your organisation's AI policy, particularly for one-on-ones. If the policy is "Microsoft only", use the two Azure options.

## 4. Microsoft 365 (optional, needs IT)

**Without sign-in**, Minutes uses Android's Share sheet:
- You share the files to the **OneDrive** app, which shows you the suggested folder.
- You share the note to **Outlook**. The addresses are copied for you to paste.

**With sign-in**, Minutes does all of this for you:
- saves straight to the right OneDrive folder, including creating a new person's one-on-one folder
- saves to SharePoint
- creates the Outlook draft with the note attached
- reads your real folder list
- fills in email addresses from your contacts

### What to send to IT

> Please register a single-page application in Microsoft Entra ID for a personal meeting-notes tool:
>
> - **Platform:** Single-page application (SPA)
> - **Redirect URI:** `https://<your-github-username>.github.io/reportingtool/`
> - **Supported account types:** this organisation only
> - **Delegated Microsoft Graph permissions:** `User.Read`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Mail.ReadWrite`, `People.Read`, `offline_access`
> - Grant admin consent if user consent is restricted.
>
> Please send me the **Application (client) ID** and the **Directory (tenant) ID**.

Enter both in **Settings → Microsoft 365** and tap **Sign in**.

Microsoft limits sign-ins for web apps to about 24 hours, so expect to tap **Sign in** again roughly once a day. Your work is kept while you do, and saving carries on afterwards.

## 5. Folders

- **One-on-ones:** `OneDrive/One-on-Ones/Firstname Surname`. A missing person folder is created on first save.
- **Projects:** top-level OneDrive folders starting with `1`, `2` or `3`. The AI suggests one and you confirm it with a tap.
- **File names:** `2026-11-20 – Meeting Title.docx` and `2026-11-20 – Meeting Title – Transcript.docx`.

Both the one-on-ones folder name and the prefixes can be changed in Settings. Until you sign in to Microsoft, list your project folders in Settings so the AI has something to choose from.

## Limits and trade-offs

- **It doesn't pick up recordings automatically.** A web app can't watch a folder, so you tap Share on each one. Only a native Android app could watch the folder.
- **It can't delete the original recording.** Android doesn't let a web app delete another app's files, so Minutes reminds you to delete it in Easy Recorder after saving. It deletes its own copy automatically.
- **It needs internet** while processing. Keep the app open during transcription; the screen is kept awake.

## Files

| Path | Purpose |
|---|---|
| `index.html`, `css/app.css` | The app shell and its iOS-style design (light and dark). |
| `js/app.js` | Screens and flow. |
| `js/providers.js` | Transcription and note-writing providers. |
| `js/note.js` | Note template, AI instructions and Word layout. |
| `js/help.js` | In-app Help: setup guides for every service, with key checks. |
| `js/docx.js` | Dependency-free Word (.docx) writer. |
| `js/graph.js` | Microsoft sign-in, OneDrive, SharePoint, Outlook and People. |
| `js/settings.js`, `js/db.js` | On-device settings, remembered people and job storage. |
| `sw.js`, `manifest.webmanifest` | Offline support, installability and the Android Share target. |
| `vendor/anthropic-sdk-1.6.0.mjs` | Anthropic TypeScript SDK bundled for the browser. |

# Minutes: pickup note

Your step-by-step routine to get Minutes running for the first time and check it works.
Do it on a computer first, then on the phone.

**Time needed:** about 1 hour in total.
**Cost:** a few pounds of prepaid credit, of which a test run uses only pennies.
**You'll need:** a payment card, a computer with Chrome (or Edge), and a short test recording.

---

## Before you start: the four pieces

| Piece | What it does | Where you get it |
|---|---|---|
| **The app's web address** | Where Minutes lives, so the phone and computer can open it | GitHub Pages, free (Step 1) |
| **AssemblyAI API key** | Turns the recording into text and works out who said what | assemblyai.com (Step 2) |
| **Claude API key** | Writes the structured meeting note from the transcript | console.anthropic.com (Step 3) |
| **Microsoft 365 sign-in** | *Optional, later.* Saves directly to OneDrive/SharePoint and creates Outlook drafts | Your IT team (Step 7) |

An **API key** is a long password that lets an app use a service and bills your account for it.
Treat it like a password:
- Never put it in GitHub, an email or a chat.
- Paste it **only** into Minutes' own Settings or Help screens.

---

## Step 1: Put the app online (10 minutes, free)

1. On your computer, go to **github.com** and sign in.
2. Open the repository **dominichare/reportingtool**.
3. Click **Settings** (top menu of the repository), then **Pages** (left-hand menu).
4. Under **Build and deployment**:
   - **Source:** choose **Deploy from a branch**.
   - **Branch:** choose **`ccr-66117013-brxjyu`**, and folder **`/ (root)`**.
   - Click **Save**.
5. Wait 1–3 minutes, then refresh the page. A box appears saying **"Your site is live at…"**.
6. Open the address: **https://dominichare.github.io/reportingtool/**

✅ **Success looks like:** a page titled **Minutes** with a blue **Choose recording** button and a **"Finish setting up"** notice.

> **Alternative:** in this Claude chat, say *"enable GitHub Pages for me"* and Claude can switch it on for you.

> The code is public, which is fine: it contains no keys or personal data. Your keys are stored only in the browser or phone where you type them.

---

## Step 2: AssemblyAI account and key (10 minutes)

AssemblyAI does the transcription and speaker separation. It costs roughly $0.15–$0.30 per hour of audio. New accounts usually get some free credit.

1. Go to **https://www.assemblyai.com/dashboard/signup** and create an account. You can use an email address or a Google account.
2. Confirm your email if asked.
3. In the dashboard, open **Billing** and add a payment card.
   - Billing is pay-as-you-go. Any free credit is used first.
4. In the dashboard, find **API Keys**. A key is usually created for you already. Click the **copy** icon next to it.
5. Keep that browser tab open. You'll paste the key into Minutes in Step 4.

✅ **Success looks like:** a key that is a long string of letters and numbers, copied to your clipboard.

---

## Step 3: Claude (Anthropic) account and key (10 minutes)

Claude writes the meeting note. That's typically 10–25p per hour-long meeting.

This is the **Claude developer console**, which is separate from the Claude chat app and any Claude subscription you may have.

1. Go to **https://console.anthropic.com/** and sign up. You may be asked to verify your email and phone number.
2. Open **Settings → Billing**, add a card and **buy credit**. **$10 is plenty** for testing and dozens of meetings.
3. *Recommended:* open **Settings → Limits** and set a monthly spend limit, e.g. $20, for peace of mind.
4. Open **Settings → API Keys** and click **Create Key**.
5. Name it **Minutes** and click **Create**.
6. Click **Copy**. ⚠️ **The key is shown only once.** If you lose it, delete it and create a new one.

✅ **Success looks like:** a key starting **`sk-ant-`** copied to your clipboard.

---

## Step 4: Set up Minutes and check both keys (5 minutes)

On the computer, open **https://dominichare.github.io/reportingtool/**. On a big screen the app shows as a phone-sized column in the middle.

1. Tap **Help** (top right) → **Getting started** and read it through.
2. Tap **⚙︎** (Settings) → **Your name**, and enter your name as you want it in notes, e.g. *Firstname Surname*.
3. Go back to **Help → AssemblyAI**:
   - Scroll to **Your details** and paste the AssemblyAI key.
   - Tap **Check**. You want a green **"It works"**.
   - The button below shows **"In use for transcription"**, because AssemblyAI is already the default.
4. **Help → Claude**:
   - Paste the Claude key and leave the model as **claude-opus-5-5**.
   - Tap **Check**. You want a green **"It works"**.
   - The button shows **"In use for note writing"**.
5. Go back to the home screen. The **"Finish setting up"** notice should have gone.

✅ **Success looks like:** two green ticks, and no setup warning on the home screen.

❌ **If Check fails:** see the **Troubleshooting** table at the end.

> Keys are saved **in that browser on that computer only**. When you move to the phone, you'll enter them again there (Step 6).
> To make that easy, keep both keys in a password manager or another safe place.

---

## Step 5: First test run on the computer (15 minutes)

### 5a. Get a test recording

You want **2–5 minutes, with 2 or 3 different voices**. Any of these will do:

- **Easiest:** record a short chat with a colleague or family member on the phone in Easy Recorder. Then get the `.m4a` file onto the computer by:
  - uploading it to OneDrive from the phone, or
  - emailing it to yourself (small files are fine), or
  - using a USB cable.
- **No one around?** Play a few minutes of a radio interview or a podcast with two presenters, and record it with Easy Recorder.

### 5b. Run it through

| # | Do this | What "working" looks like |
|---|---|---|
| 1 | On the home screen, click **Choose recording** and pick the file | The **Transcribing** screen with a pulsing icon |
| 2 | Wait | Usually under a minute for a short file. Keep the tab open. |
| 3 | **Who's speaking?** screen appears | The number of voices roughly matches reality. "Voice 1" (most talkative) is labelled **ME** with your name. |
| 4 | Click **▶** on each voice | You hear a short clip of that person. Click again to hear a different clip. |
| 5 | Type a name for each voice. If one isn't really a person (TV, noise), choose **Ignore** | **Write meeting note** turns bright blue |
| 6 | Click **Write meeting note** | "Writing note…" for up to a minute |
| 7 | Review the note | A sensible **title**, the **date**, the **participants**, then **Key decisions**, **Discussion** points and an **Actions** table |
| 8 | Click any sentence and edit it | Your edit sticks |
| 9 | In **Ask for changes**, type *"make it shorter"* and press **↑** | A revised note appears |
| 10 | Click **Transcript** (top right) | The full transcript with names and times |
| 11 | Click **Continue to save & send** | A suggested **folder** and a **file name**; the transcript switch is on; your participants are listed under **Email draft** |
| 12 | Tap the folder to see the choices, then **Cancel** | The suggestion, then One-on-Ones people, then Projects |
| 13 | Click **Next** → **Share files** | On a computer this usually **downloads** the Word files, or opens the Share window |
| 14 | Open the downloaded **.docx** in Word | A tidy note with headings, bullets, a numbered discussion and a proper **Actions table**. The transcript is a separate file. |
| 15 | Click **Finish**, then **Done** | Back on the home screen, with the note under **Recent notes** |

**Note:** until you sign in to Microsoft 365 (Step 7), "Project" folder suggestions come from a list you type in. Add a few of your real folder names, one per line, in **Settings → Your project folders**, e.g. `1 Budget Review`.

✅ **If all 15 rows behaved as described, the AI side is working.**

---

## Step 6: Install and test on the phone (10 minutes)

1. On the phone, open **Chrome** and go to **https://dominichare.github.io/reportingtool/**.
2. Tap **⋮** (top right) → **Add to Home screen** → **Install**.
3. Open **Minutes** from the home screen or app drawer. It opens full-screen, with no browser bar.
4. Enter your name and both keys again, using **Help → AssemblyAI** and **Help → Claude**, and **Check** each one.
5. Open **Easy Recorder**, choose a short recording, tap **Share**, and look for **Minutes** in the list.
   - If it isn't there, wait a minute after installing and try again.
6. Run through the same flow as Step 5b.
   - At **Share files**, pick **OneDrive** and save into the folder Minutes shows.
   - For email, tap **Copy addresses**, then **Share to Outlook**, and paste the addresses into **To**.
7. Check the files arrived in OneDrive and open them in Word on the phone.
8. Delete the test recording in Easy Recorder.

✅ **Success looks like:** you shared from Easy Recorder, the files landed in OneDrive, and an Outlook draft was ready to send.

**Then try a real meeting.** Tell people it's being recorded. If you've already saved person folders such as `One-on-Ones/Firstname Surname`, check the one-on-one folder suggestion.

---

## Step 7 (later): Microsoft 365 sign-in, via IT

This is optional. Everything works without it; signing in just removes the manual Share steps.

1. In Minutes: **Help → Microsoft 365 sign-in → Copy request to IT**, and email that text to IT.
   - It already includes the app's web address, which IT needs.
2. IT sends back two codes: an **Application (client) ID** and a **Directory (tenant) ID**. Paste both into the same Help page.
3. **Settings → Sign in**, using your work account.
4. **Settings → Refresh folder list**. Minutes should find your 1/2/3 project folders and your One-on-Ones people.
5. Run a test. **Save** now uploads directly, saves to SharePoint if switched on, and creates the Outlook draft.

---

## Troubleshooting

| You see | What to do |
|---|---|
| Check says **"didn't accept the API key"** | Re-copy the key. Make sure no spaces or missing characters crept in. For Claude, create a new key if needed. |
| **"run out of credit"** or **"busy"** | Add credit on that provider's Billing page, or wait a minute and try again. |
| Claude Check says the **model wasn't found** | Set the model to `claude-opus-5-5` exactly. |
| The site shows **404** after enabling Pages | Wait a few more minutes and refresh. Check that Settings → Pages shows the right branch and `/ (root)`. |
| **Minutes isn't in the phone's Share list** | It must be *installed* (Chrome → ⋮ → Add to Home screen → Install), not just open in a tab. |
| Too many voices (one person split in two) | Give both voices the same name; they're merged in the note. |
| Too few voices | Normal with short, overlapping recordings. Name the voice after whoever dominates the clips, and fix any wrong names in the note before saving. |
| Processing stopped part-way | Your progress is saved. Open Minutes and tap the meeting under **In progress**. |
| Anything else | Note the exact message (a screenshot is ideal) and bring it back to this Claude chat. |

---

## When you pick this chat up on the computer

Bring back:
- which steps worked and which didn't, plus any error messages or screenshots
- your impressions of the **note quality**, e.g. too long, too short, wrong emphasis, missing sections; the instructions to the AI can be tuned
- anything about the look or flow you'd like changed

Possible next steps when you're ready:
- **Import a transcript** from Samsung's Voice Recorder instead of uploading audio.
- **Note style options**, e.g. a short version vs a detailed version, or a different layout for one-on-ones.
- **Microsoft 365 sign-in**, once IT has registered the app.

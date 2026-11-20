// In-app Help: step-by-step guides for setting up each AI service and Microsoft 365.
// Each guide can include the settings fields it needs, a "Check" button and a "Use this" button,
// so you can finish setting up without leaving the guide.

export const HELP_SECTIONS = [
  { title: 'Start here', topics: ['start', 'choose', 'using'] },
  { title: 'Transcription (speech to text)', topics: ['assemblyai', 'openai-stt', 'azure-speech'] },
  { title: 'Note writing (AI)', topics: ['claude', 'openai-llm', 'azure-openai'] },
  { title: 'Saving and email', topics: ['microsoft365'] },
  { title: 'More', topics: ['privacy', 'troubleshooting'] },
];

const IT_REQUEST = `Please register a single-page application in Microsoft Entra ID for a personal meeting-notes tool:

- Platform: Single-page application (SPA)
- Redirect URI: {REDIRECT}
- Supported account types: this organisation only
- Delegated Microsoft Graph permissions: User.Read, Files.ReadWrite.All, Sites.ReadWrite.All, Mail.ReadWrite, People.Read, offline_access
- Please grant admin consent if user consent is restricted.

Please send me the Application (client) ID and the Directory (tenant) ID.`;

export const HELP_TOPICS = {
  start: {
    title: 'Getting started',
    color: 'var(--accent)', icon: 'sparkle',
    summary: 'Minutes needs two AI services: one to turn speech into text, and one to write the meeting note. Set up both once (about 15 minutes) and you’re done.',
    steps: [
      { title: 'Enter your name', body: 'Open <b>Settings</b> and type your name as you want it to appear in notes, e.g. <i>Firstname Surname</i>. The voice that talks most in a recording is labelled as you automatically.', action: 'settings' },
      { title: 'Choose a transcription service', body: 'This listens to the recording and works out who said what. <b>Recommended: AssemblyAI</b>. It’s the best at telling speakers apart and handles hour-long meetings. If your organisation only allows Microsoft services, use <b>Azure Speech</b> instead.', link: { topic: 'assemblyai', label: 'Set up AssemblyAI' } },
      { title: 'Choose a note-writing service', body: 'This reads the transcript and writes the structured note. <b>Recommended: Claude</b>, which writes clear, well-organised notes. For Microsoft-only organisations, use <b>Azure OpenAI</b>.', link: { topic: 'claude', label: 'Set up Claude' } },
      { title: 'Check both keys', body: 'Each guide has a <b>Check</b> button that confirms the key works, without using a recording or costing anything.' },
      { title: 'Optional: connect Microsoft 365', body: 'Without it, you share finished notes to the OneDrive and Outlook apps yourself (a couple of taps). With it, Minutes saves to the right folder and creates the email draft for you. It needs a one-off request to IT.', link: { topic: 'microsoft365', label: 'Microsoft 365 guide' } },
    ],
  },

  choose: {
    title: 'Which service should I choose?',
    color: '#5856d6', icon: 'info',
    summary: 'Any combination works. You can switch at any time in Settings, and nothing else changes.',
    table: {
      caption: 'Transcription',
      header: ['Service', 'Best for', 'Approx. cost per hour', 'Setup'],
      rows: [
        ['AssemblyAI', 'Best speaker separation; long meetings', '$0.15–$0.30', 'Easy: 5 min'],
        ['OpenAI', 'One key for everything', '≈ $0.36', 'Easy: 5 min'],
        ['Azure Speech', 'Data stays in your company’s Microsoft cloud', 'Varies by region', 'Needs IT'],
      ],
    },
    table2: {
      caption: 'Note writing',
      header: ['Service', 'Best for', 'Approx. cost per note', 'Setup'],
      rows: [
        ['Claude', 'Clearest, best-structured notes', '≈ 10–25p', 'Easy: 5 min'],
        ['OpenAI', 'One key for everything', '≈ 2–10p', 'Easy: 5 min'],
        ['Azure OpenAI', 'Data stays in your company’s Microsoft cloud', 'Set by your company', 'Needs IT'],
      ],
    },
    after: '<p><b>Simplest:</b> OpenAI for both (one account, one key). <b>Best quality:</b> AssemblyAI + Claude. <b>Strict company policy:</b> Azure Speech + Azure OpenAI.</p><p>Prices are approximate, in US dollars or pounds, and change over time. Check each provider’s pricing page.</p>',
  },

  using: {
    title: 'Recording to note, step by step',
    color: 'var(--green)', icon: 'wave',
    summary: 'What happens each time you turn a recording into a note.',
    steps: [
      { title: 'Install Minutes (once)', body: 'Open the Minutes web address in <b>Chrome</b>, tap <b>⋮</b> (top right), then <b>Add to Home screen → Install</b>. Minutes then appears in your app drawer and in Android’s Share menu.' },
      { title: 'Share the recording', body: 'In <b>Easy Recorder</b>, open the recording, tap <b>Share</b> and choose <b>Minutes</b>. Or open Minutes and tap <b>Choose recording</b>, then browse to <i>Recordings</i>.' },
      { title: 'Wait for the transcript', body: 'An hour-long meeting usually takes 1–3 minutes. Keep Minutes open; the screen stays on while it works.' },
      { title: 'Name the voices', body: 'Each voice gets a card with a sample of what they said. Tap <b>▶</b> to hear them (tap again for a different clip). Type a name once, or tap a name you’ve used before. Use <b>Ignore</b> for background noise or a TV.' },
      { title: 'Review the note', body: 'Tap any text to correct it. To change the whole note, type a request such as <i>“make it shorter”</i> or <i>“add more detail on the budget”</i> and tap <b>↑</b>.' },
      { title: 'Save and send', body: 'Check the suggested folder (tap to change), switch SharePoint on if needed, and check who gets the email draft. Tap <b>Save</b>.' },
      { title: 'Delete the recording', body: 'Minutes reminds you to delete the original in Easy Recorder. It has already removed its own copy.' },
    ],
  },

  assemblyai: {
    title: 'AssemblyAI',
    color: '#2545d3', icon: 'wave',
    kind: 'Transcription',
    summary: 'A specialist speech-to-text company. Very good at telling speakers apart, handles long recordings, and is inexpensive.',
    facts: [
      ['Cost', 'About $0.15–$0.30 per hour of audio, including speaker separation. New accounts usually get free credit to try it.'],
      ['Recording length', 'No practical limit for meetings'],
      ['Your data', 'Audio is sent to AssemblyAI (USA by default) for processing. Check this is acceptable under your organisation’s policy.'],
      ['You’ll need', 'An email address and a payment card'],
    ],
    steps: [
      { title: 'Create an account', body: 'Go to AssemblyAI and sign up with your email (or Google account).', url: 'https://www.assemblyai.com/dashboard/signup', urlLabel: 'Open AssemblyAI sign-up' },
      { title: 'Add a payment method', body: 'In the dashboard, open <b>Billing</b> and add a card. Pay-as-you-go: you’re charged only for what you transcribe. Any free credit is used first.', url: 'https://www.assemblyai.com/dashboard', urlLabel: 'Open dashboard' },
      { title: 'Copy your API key', body: 'In the dashboard, open <b>API Keys</b> and tap the copy icon next to your key. It’s a long string of letters and numbers.' },
      { title: 'Paste it below and tap Check', body: 'Paste the key into the box below. Tap <b>Check</b>; you should see a green tick. Then tap <b>Use AssemblyAI</b>.' },
    ],
    fields: [{ key: 'assemblyaiKey', label: 'API key', type: 'password', placeholder: 'Paste key' }],
    check: 'assemblyai',
    use: { setting: 'sttProvider', value: 'assemblyai', label: 'Use AssemblyAI for transcription' },
  },

  'openai-stt': {
    title: 'OpenAI (transcription)',
    color: '#10a37f', icon: 'wave',
    kind: 'Transcription',
    summary: 'The company behind ChatGPT. One OpenAI key can do both transcription and note writing.',
    facts: [
      ['Cost', 'About $0.006 per minute (≈ $0.36 per hour), including speaker separation.'],
      ['Recording length', 'Files up to 25 MB, which is about 25–50 minutes depending on recording quality. For longer meetings use AssemblyAI or Azure Speech.'],
      ['Your data', 'Sent to OpenAI (USA). By default, OpenAI does not use API data to train its models.'],
      ['Note', 'OpenAI may retire this speaker-separating model during 2027. If that happens, switch to AssemblyAI.'],
      ['You’ll need', 'An email address and a payment card. A ChatGPT subscription is <b>not</b> the same thing and isn’t needed.'],
    ],
    steps: [
      { title: 'Create an OpenAI platform account', body: 'Sign up (or log in) on the OpenAI <b>developer platform</b>, which is separate from ChatGPT.', url: 'https://platform.openai.com/signup', urlLabel: 'Open OpenAI platform' },
      { title: 'Add credit', body: 'Open <b>Settings → Billing</b>, add a card, and buy a small amount of credit (e.g. $10). You can also set a monthly limit there so you never spend more than you intend.', url: 'https://platform.openai.com/settings/organization/billing/overview', urlLabel: 'Open Billing' },
      { title: 'Create an API key', body: 'Open <b>API keys</b> and tap <b>Create new secret key</b>. Name it “Minutes”, then copy it. <b>It is shown only once</b>, so copy it before closing.', url: 'https://platform.openai.com/api-keys', urlLabel: 'Open API keys' },
      { title: 'Paste it below and tap Check', body: 'Paste the key (it starts with <i>sk-</i>), tap <b>Check</b>, then <b>Use OpenAI</b>. The same key is used if you also choose OpenAI for note writing.' },
    ],
    fields: [{ key: 'openaiKey', label: 'API key', type: 'password', placeholder: 'sk-…' }],
    check: 'openai',
    use: { setting: 'sttProvider', value: 'openai', label: 'Use OpenAI for transcription' },
  },

  'azure-speech': {
    title: 'Azure Speech',
    color: '#0078d4', icon: 'cloud',
    kind: 'Transcription',
    summary: 'Microsoft’s speech-to-text service, run inside your organisation’s own Azure account. Usually the easiest option to get approved under a strict company AI policy.',
    facts: [
      ['Cost', 'Billed to your organisation’s Azure account. Check the Azure pricing page for your region.'],
      ['Recording length', 'Up to about 2 hours / 300 MB per recording'],
      ['Your data', 'Stays within your organisation’s Microsoft Azure environment, in the region IT chooses (e.g. UK South).'],
      ['You’ll need', 'IT (or whoever manages your company’s Azure) to create the service and give you a key'],
    ],
    steps: [
      { title: 'Ask IT for a Speech resource', body: 'Send IT the request below. If you manage Azure yourself: in the Azure portal choose <b>Create a resource → Speech</b>, pick the <b>UK South</b> region and the <b>Standard S0</b> tier.', copy: 'Please could you create an Azure AI Speech resource for me (region: UK South, pricing tier: Standard S0) for transcribing meeting recordings with speaker separation, and send me Key 1 and the region name? It will be called directly from a web app on my phone using the fast transcription REST API (2024-11-15).', copyLabel: 'Copy request to IT', url: 'https://portal.azure.com/#create/Microsoft.CognitiveServicesSpeechServices', urlLabel: 'Open Azure portal' },
      { title: 'Find the key and region', body: 'In the Speech resource, open <b>Keys and Endpoint</b>. Copy <b>KEY 1</b> and note the <b>Location/Region</b> (e.g. <i>uksouth</i>, written without spaces).' },
      { title: 'Paste them below and tap Check', body: 'Enter the key and region, tap <b>Check</b>, then <b>Use Azure Speech</b>.' },
    ],
    fields: [
      { key: 'azureSpeechKey', label: 'Key', type: 'password', placeholder: 'Paste KEY 1' },
      { key: 'azureSpeechRegion', label: 'Region', placeholder: 'uksouth' },
    ],
    check: 'azure',
    use: { setting: 'sttProvider', value: 'azure', label: 'Use Azure Speech for transcription' },
  },

  claude: {
    title: 'Claude',
    color: '#d97757', icon: 'sparkle',
    kind: 'Note writing',
    summary: 'Anthropic’s AI. Writes clear, concise, well-structured meeting notes and follows the note template closely.',
    facts: [
      ['Cost', 'Typically 10–25p per hour-long meeting. Claude Opus 5.5 is priced at $4 per million tokens read and $20 per million written; an hour’s transcript is roughly 15,000 tokens.'],
      ['Model', '<i>claude-opus-5-5</i> (recommended). You can change it in Settings, e.g. to <i>claude-sonnet-5-5</i> for lower cost.'],
      ['Your data', 'Sent to Anthropic. By default, Anthropic does not train on API data.'],
      ['You’ll need', 'An email address and a payment card. A Claude.ai subscription is separate and isn’t needed.'],
    ],
    steps: [
      { title: 'Create an Anthropic Console account', body: 'Sign up on the Claude developer console (separate from the Claude chat app).', url: 'https://console.anthropic.com/', urlLabel: 'Open Claude Console' },
      { title: 'Buy credit', body: 'Open <b>Settings → Billing</b>, add a card and buy a small amount of credit (e.g. $10, enough for dozens of meetings). Under <b>Limits</b> you can set a monthly spend cap.', url: 'https://console.anthropic.com/settings/billing', urlLabel: 'Open Billing' },
      { title: 'Create an API key', body: 'Open <b>API Keys</b> and tap <b>Create Key</b>. Name it “Minutes” and copy it. <b>It is shown only once.</b>', url: 'https://console.anthropic.com/settings/keys', urlLabel: 'Open API Keys' },
      { title: 'Paste it below and tap Check', body: 'Paste the key (it starts with <i>sk-ant-</i>), tap <b>Check</b> (this also confirms the model name), then <b>Use Claude</b>.' },
    ],
    fields: [
      { key: 'anthropicKey', label: 'API key', type: 'password', placeholder: 'sk-ant-…' },
      { key: 'anthropicModel', label: 'Model', placeholder: 'claude-opus-5-5' },
    ],
    check: 'anthropic',
    use: { setting: 'llmProvider', value: 'anthropic', label: 'Use Claude for note writing' },
  },

  'openai-llm': {
    title: 'OpenAI (note writing)',
    color: '#10a37f', icon: 'sparkle',
    kind: 'Note writing',
    summary: 'GPT models from OpenAI. If you already use OpenAI for transcription, the same key works here.',
    facts: [
      ['Cost', 'Typically 2–10p per hour-long meeting, depending on the model.'],
      ['Model', '<i>gpt-5</i> by default. You can change it in Settings, e.g. <i>gpt-5-mini</i> for lower cost.'],
      ['Your data', 'Sent to OpenAI (USA). By default, OpenAI does not use API data to train its models.'],
      ['You’ll need', 'An OpenAI platform account with credit (see the OpenAI transcription guide)'],
    ],
    steps: [
      { title: 'Get an OpenAI API key', body: 'If you haven’t already, follow steps 1–3 of the OpenAI transcription guide: create a platform account, add credit and create an API key.', link: { topic: 'openai-stt', label: 'OpenAI account steps' } },
      { title: 'Paste it below and tap Check', body: 'Paste the key, check the model name, tap <b>Check</b> (this also confirms your account can use that model), then <b>Use OpenAI</b>.' },
    ],
    fields: [
      { key: 'openaiKey', label: 'API key', type: 'password', placeholder: 'sk-…' },
      { key: 'openaiModel', label: 'Model', placeholder: 'gpt-5' },
    ],
    check: 'openai-llm',
    use: { setting: 'llmProvider', value: 'openai', label: 'Use OpenAI for note writing' },
  },

  'azure-openai': {
    title: 'Azure OpenAI',
    color: '#0078d4', icon: 'cloud',
    kind: 'Note writing',
    summary: 'OpenAI’s GPT models, run inside your organisation’s own Azure account. Pairs with Azure Speech for a Microsoft-only setup.',
    facts: [
      ['Cost', 'Billed to your organisation’s Azure account.'],
      ['Your data', 'Stays within your organisation’s Azure environment, in the region IT chooses.'],
      ['You’ll need', 'IT to create an Azure OpenAI resource, deploy a model, and give you three things: endpoint, deployment name and key'],
    ],
    steps: [
      { title: 'Ask IT for an Azure OpenAI deployment', body: 'Send IT the request below. If you manage Azure yourself: create an <b>Azure OpenAI</b> resource (UK South), then in <b>Azure AI Foundry</b> deploy a chat model such as <i>gpt-4.1</i> or <i>gpt-5-mini</i>.', copy: 'Please could you set up an Azure OpenAI resource for me (region: UK South) with a deployed chat model that supports structured JSON output (e.g. gpt-4.1 or gpt-5-mini), for writing meeting notes from transcripts? Please send me the endpoint URL, the deployment name and Key 1. It will be called directly from a web app on my phone (chat completions API, version 2024-10-21).', copyLabel: 'Copy request to IT', url: 'https://ai.azure.com/', urlLabel: 'Open Azure AI Foundry' },
      { title: 'Collect the three details', body: '<b>Endpoint</b>: looks like <i>https://yourname.openai.azure.com</i>. <b>Deployment</b>: the name IT gave the model deployment (not necessarily the model name). <b>Key</b>: KEY 1 from <b>Keys and Endpoint</b>.' },
      { title: 'Paste them below and tap Check', body: 'Enter all three, tap <b>Check</b> (this sends a tiny test message), then <b>Use Azure OpenAI</b>.' },
    ],
    fields: [
      { key: 'azureOpenaiEndpoint', label: 'Endpoint', placeholder: 'https://….openai.azure.com' },
      { key: 'azureOpenaiDeployment', label: 'Deployment', placeholder: 'e.g. gpt-4-1' },
      { key: 'azureOpenaiKey', label: 'Key', type: 'password', placeholder: 'Paste KEY 1' },
    ],
    check: 'azureopenai',
    use: { setting: 'llmProvider', value: 'azureopenai', label: 'Use Azure OpenAI for note writing' },
  },

  microsoft365: {
    title: 'Microsoft 365 sign-in',
    color: '#0a64d8', icon: 'folder',
    summary: 'Optional. Lets Minutes save straight into OneDrive and SharePoint, read your real folder list, look up email addresses and create the Outlook draft for you.',
    facts: [
      ['Without it', 'You share the finished Word files to the OneDrive and Outlook apps yourself. Minutes shows the folder to pick and copies the email addresses.'],
      ['With it', 'One tap on <b>Save</b> does everything.'],
      ['You’ll need', 'IT to register Minutes as an app in Microsoft Entra ID (a standard, free, five-minute task)'],
      ['Sign-in', 'Microsoft limits web apps to about a day per sign-in, so expect to tap <b>Sign in</b> roughly once a day. Your work is kept while you do.'],
    ],
    steps: [
      { title: 'Send IT this request', body: 'Copy the request and email it to your IT team. It includes this app’s exact web address, which IT needs.', copy: IT_REQUEST, copyLabel: 'Copy request to IT' },
      { title: 'Enter the two IDs from IT', body: 'IT will send an <b>Application (client) ID</b> and a <b>Directory (tenant) ID</b>, both long codes like <i>1a2b3c4d-…</i>. Paste them below.' },
      { title: 'Sign in', body: 'Open <b>Settings</b> and tap <b>Sign in</b> under Microsoft 365. Use your work account. If you see “Need admin approval”, IT still needs to grant consent.', action: 'settings' },
      { title: 'Refresh your folders', body: 'In <b>Settings → Folders</b>, tap <b>Refresh folder list</b>. Minutes should report your 1/2/3 project folders and your One-on-Ones people.' },
    ],
    fields: [
      { key: 'msClientId', label: 'Client ID', placeholder: 'Application (client) ID' },
      { key: 'msTenant', label: 'Tenant ID', placeholder: 'Directory (tenant) ID' },
    ],
  },

  privacy: {
    title: 'Privacy and costs',
    color: '#8e8e93', icon: 'info',
    summary: 'Where your recordings go, and what you pay.',
    steps: [
      { title: 'On your phone', body: 'API keys, your settings and the list of people you’ve named are stored only on this phone, inside Minutes. They’re never sent anywhere except to the service they belong to.' },
      { title: 'During processing', body: 'The recording goes only to your chosen transcription service, and the transcript only to your chosen note-writing service. Nothing passes through any other server.' },
      { title: 'Afterwards', body: 'Minutes deletes its copy of the audio once the note is saved. The original stays in Easy Recorder until you delete it.' },
      { title: 'Company policy', body: 'Recordings of one-on-ones can be sensitive. If your organisation restricts AI tools, use <b>Azure Speech</b> and <b>Azure OpenAI</b>, which keep everything inside your company’s Microsoft environment. Always tell people a meeting is being recorded.' },
      { title: 'Costs', body: 'You pay each provider directly, per use, with no subscription. Typical total: <b>under £1 per hour of meeting</b>. Set monthly spending limits on each provider’s billing page for peace of mind.' },
    ],
  },

  troubleshooting: {
    title: 'Troubleshooting',
    color: 'var(--orange)', icon: 'warn',
    summary: 'Common messages and what to do.',
    steps: [
      { title: '“didn’t accept the API key”', body: 'The key was mistyped, revoked, or copied with a missing character. Create a new key on the provider’s site, paste it again, and tap <b>Check</b>.' },
      { title: '“busy or your account has run out of credit”', body: 'Add credit on the provider’s billing page, or wait a minute and tap <b>Try again</b>.' },
      { title: '“too large for OpenAI”', body: 'OpenAI accepts files up to 25 MB. Switch transcription to AssemblyAI or Azure Speech in Settings, then tap <b>Try again</b>.' },
      { title: 'Minutes isn’t in the Share menu', body: 'Minutes must be <b>installed</b> (Chrome → ⋮ → Add to Home screen → Install), not just opened in a tab. After installing, it can take a minute to appear.' },
      { title: 'Two voices are the same person', body: 'Give both the same name; they’ll be merged in the note.' },
      { title: 'One voice is really two people', body: 'Use the name of whoever speaks most in the samples. You can correct names in the note before saving.' },
      { title: '“Need admin approval” when signing in', body: 'Your IT team needs to grant consent to the app registration. Meanwhile, use the Share route: it works without signing in.' },
      { title: 'Processing stopped part-way', body: 'Your progress is saved. Open Minutes, tap the meeting under <b>In progress</b>, and continue.' },
    ],
  },
};

export function itRequestText(redirect) {
  return IT_REQUEST.replace('{REDIRECT}', redirect);
}

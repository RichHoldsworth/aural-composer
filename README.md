# Aural Composer

A browser-based tool for building Music listening-exam audio files. Stitch your own audio extracts, YouTube clips, and Spotify tracks together with examiner-style spoken announcements, plays/gaps, and reading time — then export the whole exam as a single WAV.

---

## Deploy this app to Vercel (no installation needed)

These steps assume you're on a school laptop without admin rights. Everything happens in a web browser.

### Step 1 — Create a free GitHub account (if you don't have one)

1. Open [github.com](https://github.com) in your browser
2. Click **Sign up** and follow the steps
3. Verify your email when GitHub asks

### Step 2 — Create a new repository

1. While logged in to GitHub, click the **+** in the top-right corner → **New repository**
2. Repository name: `aural-composer` (or anything you like)
3. Make it **Public** (so Vercel can read it without extra permissions)
4. **Don't** tick "Add a README" or any other initialisation options — leave them all unchecked
5. Click **Create repository**

You'll land on a page that says "Quick setup — if you've done this kind of thing before." Don't follow those instructions.

### Step 3 — Upload the project files

1. On that same page, look for the link **uploading an existing file** in the small grey text — click it
2. A drag-and-drop zone appears
3. From the folder I gave you (`aural-composer`), select **all the files and folders** and drag them into the drop zone:
   - `package.json`
   - `vite.config.js`
   - `tailwind.config.js`
   - `postcss.config.js`
   - `index.html`
   - `.gitignore`
   - The entire `src/` folder
4. Wait for the upload to finish (you'll see each file listed)
5. Scroll to the bottom, click **Commit changes**

### Step 4 — Create a free Vercel account

1. Open [vercel.com](https://vercel.com) in a new tab
2. Click **Sign Up**
3. Choose **Continue with GitHub** — this links the two accounts so Vercel can deploy directly from your repo
4. Approve the permissions Vercel asks for

### Step 5 — Deploy the app

1. On your Vercel dashboard, click **Add New… → Project**
2. You'll see a list of your GitHub repositories. Find `aural-composer` and click **Import**
3. Vercel auto-detects everything (it sees it's a Vite + React project)
4. **Don't change any settings.** Just click **Deploy**
5. Wait ~60 seconds while it builds. You'll see logs scrolling

When it's done, you'll get a URL like `https://aural-composer-xxx.vercel.app` — that's your app. Click it to open.

### Step 6 — Bookmark it

That URL is yours forever. Bookmark it on your school laptop, your home computer, your phone — anywhere you might need it.

---

## Updating the app later

If I send you new code (a new `App.jsx` for example):

1. Go to your GitHub repo → click `src/App.jsx`
2. Click the pencil icon (top right) to edit
3. Select all the existing code (Ctrl+A) and delete it
4. Paste the new code
5. Scroll down → click **Commit changes**
6. Vercel automatically detects the change and redeploys within a minute

---

## Hooking up Spotify

Once the app is deployed at e.g. `https://aural-composer-xxx.vercel.app`:

1. Open the app → **Voice & API** → scroll to **Spotify Connection**
2. Copy the Redirect URI it shows you (it'll be your Vercel URL)
3. Go to [developer.spotify.com/dashboard](https://developer.spotify.com/dashboard) → your app → **Settings**
4. Edit the **Redirect URIs**, add the one from step 2, save
5. Back in the app, paste your Client ID, click **Connect Spotify**

---

## Hooking up premium TTS (ElevenLabs / OpenAI)

For high-quality examiner voice baked into the WAV file:

- **ElevenLabs**: free tier gives ~10,000 characters/month — easily enough for one exam. Sign up at [elevenlabs.io](https://elevenlabs.io), copy your API key from settings, paste into the app under **Voice & API → ElevenLabs**.
- **OpenAI**: pay-as-you-go (about $0.015 per 1,000 characters with `tts-1-hd` — roughly 10–15p for a full exam). Sign up at [platform.openai.com](https://platform.openai.com), create an API key, paste it into the app.

Both keys are stored only in your browser's localStorage; they never leave your machine except when calling the provider directly.

---

## Running it locally instead (optional, requires Node.js)

If you ever want to run it on a machine where you DO have admin rights:

```bash
npm install
npm run dev
```

Then open `http://localhost:5173`.

---

## Project structure

```
aural-composer/
├── index.html              ← entry HTML
├── package.json            ← dependencies
├── vite.config.js          ← build tool config
├── tailwind.config.js      ← CSS framework config
├── postcss.config.js       ← CSS pipeline config
├── .gitignore              ← files to keep out of git
└── src/
    ├── main.jsx            ← React bootstrap
    ├── App.jsx             ← THE APP (all logic + UI)
    └── index.css           ← base styles
```

The vast majority of code lives in `src/App.jsx`. When iterating on features, that's almost always the only file you need to edit.

# RingCentral ACE Transcript Downloader
## SE How-To Guide — Mac Setup from Scratch

---

## What This Does

This python connects to RingCentral's API, pulls every recorded call across your entire account, fetches the AI-generated transcript, summary, and sentiment for each one.  This pyhton handles rate limits automatically, retries failed requests, and only downloads transcripts for calls where a RingSense license is assigned.  Downloads all RingSense call transcripts from a customer's RingCentral account and saves two files to your Mac:

  - **Excel spreadsheet** — all calls with metadata, AI summaries, and full transcripts
  - **PDF document** — formatted, branded transcript book with sentiment badges and AI summaries

Both files save to the same folder as the script. Run it once per customer.  Fetching transcripts Takes about 1.5 seconds per call -- please be patient. This program installs its own dependencies, then walks the SE through everything interactively:

SE will need three things from developers.ringcentral.com:

    Client ID      -- from your app Credentials tab
    Client Secret  -- from your app Credentials tab
    JWT Token      -- from your app Auth tab (Create JWT)

    These are only used during this session.
    They are never saved to disk.


After running python SE will be asked for the following;
  Customer name
  Client ID, Client Secret, JWT Token (password-masked so nothing shows on screen)
  Date range (menu — Last 7 / 30 / 90 days or custom)

---

## Before You Start

You will need:

1. A Mac (this guide is written for Mac)
2. The `download_transcripts.py` script (in this same folder)
3. A user account on the customer's RingCentral account (Super Admin role)
4. A RingCentral developer app with credentials (covered in Phase 1 below)

---

## Phase 1 — One-Time Customer Setup

Do this once per customer, with the customer's admin.

---

### Step 1 — Create an SE User on the Customer's Account

Ask the customer's RingCentral admin to:

1. Log into **admin.ringcentral.com**
2. Go to **Users** → **Add User**
3. Fill in your name, email, and set the role to **Super Admin**
4. Assign a **RingSense license** to the user (required to access transcripts)
5. Save — you will receive an email invitation

Accept the invitation and set your password.

> **Why Super Admin?** This role lets you pull call logs across all extensions,
> not just your own.

---

### Step 2 — Create a Developer App on the Customer's Account

1. Go to **developers.ringcentral.com**
2. Log in with the **SE user credentials** you just created
3. Click **Console** in the top right
4. Click **Register App**
5. Select **REST API** → click **Next**
6. Fill in:
   - **App Name:** ACE Transcript Downloader
   - **App Type:** Private
   - **Auth Type:** JWT auth flow
   - **Permissions:** check **Read Call Log**  and if available **Ringsense** ← (if Ringsense scope is not present see stp 5 below)
7. Upload any square image as the app icon (required)
8. Click **Register**

---

### Step 3 — Copy Your Client ID and Client Secret

After registering, you will see the app credentials page:

1. Copy your **Client ID** — paste it into a Notes doc to save it
2. Click **Show** next to **Client Secret** — copy and save it

> Keep these safe. Do not share them in email or Slack.

---

### Step 4 — Create a JWT Token

On the same app page:

1. Click the **Auth** tab
2. Click **Create JWT**
3. Name it: `ACE Downloader`
4. Click **Create**
5. **Copy the token immediately** — it is only shown once

> If you close the page without copying it, you will need to delete it and create a new one.

---

### Step 5 — Request the RingSense Scope

The RingSense scope is what gives the app access to transcripts.
It requires a separate approval from RingCentral.

1. Go to: **developers.ringcentral.com/api-products/ringsense**
2. Fill out the request form — include your app's Client ID
3. RingCentral approves within **1-2 business days**
4. You will receive a confirmation email

**Once approved — do these two things:**

1. Go back to your app in the Developer Console
2. Click **Settings** → find **Permissions** → check **RingSense** → save
3. Go to the **Auth** tab → delete your JWT token → create a new one

> **Important:** The JWT token must be regenerated **after** the RingSense scope
> is approved. An old token will not include the scope and transcripts will not download.

---

## Phase 2 — One-Time Mac Setup

Do this once on your Mac.

---

### Step 6 — Install Python

1. Open **Terminal**
   - Press **Command + Spacebar**
   - Type **Terminal**
   - Press **Enter**

2. Check if Python is already installed:
```
python3 --version
```

If you see `Python 3.x.x` — you are done with this step.

If you see `command not found`:

3. Go to **python.org/downloads**
4. Click the big yellow **Download Python** button
5. Open the downloaded file and follow the installer
6. When done, close and reopen Terminal
7. Run `python3 --version` again to confirm

---

### Step 7 — Put the Script on Your Mac

1. Download **download_transcripts.py** from wherever it was shared with you
2. Create a folder on your Desktop called **rcpoc**
3. Move the script into that folder

In Terminal:
```
mkdir ~/Desktop/download_transcripts
mv ~/Downloads/download_transcripts.py ~/Desktop/download_transcripts/
```

---

## Phase 3 — Running the Script

Do this for each customer POC.

---

### Step 8 — Run the Script

In Terminal:
```
cd ~/Desktop/download_transcripts
python3 download_transcripts.py
```

The first time you run it, it will install its dependencies automatically.
This takes about 30 seconds and only happens once.

---

### Step 9 — Answer the Prompts

The script will walk you through everything:

```
+======================================================+
|  RingCentral ACE Transcript Downloader               |
|  Downloads all call transcripts to Excel + PDF       |
+======================================================+

STEP 1 -- Customer & Credentials
--------------------------------------------------------
  Customer company name: Acme Corp
  Client ID: AbCdEfGhIjKlMnOp
  Client Secret: (you type this -- nothing shows on screen)
  JWT Token: (you type this -- nothing shows on screen)

STEP 2 -- Authenticating with RingCentral
  ->  Connecting...
  OK  Authenticated successfully

STEP 3 -- Date Range
  How far back do you want to pull transcripts?
    1. Last 7 days
    2. Last 30 days      <-- recommended for POC
    3. Last 90 days
    4. Custom range

  Enter number (1-4): 2

  Ready to run:
    Customer  : Acme Corp
    Date range: 2026-04-03 to 2026-05-03
    Output to : /Users/yourname/Desktop/rcpoc

  Start download? (y/n): y
```

---

### Step 10 — Wait for It to Run

You will see a progress bar as transcripts download:

```
STEP 5 -- Fetching RingSense Transcripts
  ->  Fetching transcripts for 312 calls...
  ->  About 1.5 seconds per call -- please be patient

    [########################################] 100%  312/312

  OK  247 of 312 calls have transcripts
```

**Do not close Terminal while it is running.**
It takes about 1.5 seconds per call, so 300 calls takes around 8 minutes.

If it pauses and shows:
```
  !!  Rate limit reached -- waiting 65 seconds...
```
That is normal — just leave it running.

---

### Step 11 — Open Your Files

When done you will see:

```
+======================================================+
|  COMPLETE                                            |
+======================================================+

  Customer    : Acme Corp
  Total calls : 312
  Transcripts : 247
  No transcript: 65 (unlicensed extensions)

  Files saved in the same folder as this script:

    [XLSX]  transcripts_acme_corp_20260503.xlsx        218 KB
    [PDF ]  transcripts_acme_corp_20260503.pdf           4.2 MB

  To open the folder:
    open "/Users/yourname/Desktop/download_transcripts"
```

Copy and paste that `open` command into Terminal to see your files.

---

## Running for a Different Customer

Just run the script again:

```
cd ~/Desktop/download_transcripts
python3 download_transcripts.py
```

Enter the new customer's credentials when prompted.
Each run creates new files with the customer name and date in the filename.
Previous files are never overwritten.

---

## What the Excel File Contains

Three tabs:

| Tab | Contents |
|---|---|
| **Summary** | Total calls, transcripts, talk time, date range |
| **All Calls** | One row per call — date, time, direction, duration, names, sentiment, AI summary |
| **Transcripts** | Full timestamped transcript for every call that has one |

Color coding in All Calls:
- Green rows = Positive sentiment
- Red rows = Negative sentiment
- Grey/white = Neutral

---

## What the PDF Contains

- Cover page with stats (total calls, transcripts, talk time, dates)
- One section per call with:
  - Direction badge (Inbound / Outbound)
  - Sentiment badge (Positive / Negative / Neutral)
  - AI-generated summary
  - Full timestamped transcript with real speaker names
- Page numbers and customer name on every page

---

## Troubleshooting

| What you see | What to do |
|---|---|
| `command not found: python3` | Install Python from python.org/downloads |
| `Authentication failed` | Check Client ID, Client Secret, and JWT Token — make sure they are from the SE user account you created on the customer's system |
| `RingSense scope not found` | Scope not yet approved, or JWT was not regenerated after approval — create a new JWT token |
| `0 calls have transcripts` | Extensions do not have RingSense licenses — ask the customer's admin to assign them in admin.ringcentral.com |
| `No recorded calls found` | Try a wider date range, or confirm that call recording is enabled |
| `Rate limit -- waiting 65 seconds` | Normal — the script handles this automatically, just wait |
| Script stops unexpectedly | Reopen Terminal, `cd ~/Desktop/rcpoc`, run again |

---

## Security Notes

- Your credentials are **never saved to disk**
- The JWT token, Client Secret, and Client ID are only used during the session
- When the POC is complete, ask the customer's admin to:
  - Delete the SE user from their account
  - Delete the developer app from their account

---

*RingCentral ACE Transcript Downloader · SE Internal Tool · Confidential*

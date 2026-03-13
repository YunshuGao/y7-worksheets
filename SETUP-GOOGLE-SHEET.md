# How to Set Up the Google Sheet Submission System

This takes about 5 minutes. You only need to do it once.

## Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) using your **personal Gmail** account
2. Create a new blank spreadsheet
3. Name it: **Write It Right — Student Submissions**
4. In Row 1, add these column headers (exactly as written):

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Timestamp | Student Name | Band | Diagnostic Score | M1 Word Power | M2 Text Types | M3 Reading | M4 TEEL | M5 Source Analysis | M6 Capstone | Responses JSON |

## Step 2 — Add the Apps Script

1. In your Google Sheet, click **Extensions → Apps Script**
2. Delete everything in the code editor
3. Paste this code:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date().toLocaleString('en-AU', {timeZone: 'Australia/Sydney'}),
      data.studentName || '',
      data.band || '',
      data.diagnosticScore || 0,
      data.moduleScores['1'] || 0,
      data.moduleScores['2'] || 0,
      data.moduleScores['3'] || 0,
      data.moduleScores['4'] || 0,
      data.moduleScores['5'] || 0,
      data.moduleScores['6'] || 0,
      JSON.stringify(data.responses || {})
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

4. Click **Save** (Ctrl+S), name the project "Write It Right Submissions"

## Step 3 — Deploy as Web App

1. Click **Deploy → New deployment**
2. Click the gear icon ⚙️ next to "Select type" → choose **Web app**
3. Set:
   - **Description:** Student submissions
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**
5. Click **Authorize access** → choose your personal Gmail → click **Advanced** → click **Go to Write It Right Submissions (unsafe)** → click **Allow**
6. **Copy the Web App URL** — it looks like: `https://script.google.com/macros/s/AKfycbx.../exec`

## Step 4 — Paste the URL into Write It Right

Open `Write-It-Right.html` and find this line near the top of the `<script>` section:

```javascript
const SUBMIT_URL = '';
```

Paste your URL between the quotes:

```javascript
const SUBMIT_URL = 'https://script.google.com/macros/s/AKfycbx.../exec';
```

Save, commit, and push. Students can now submit!

## How It Works

- Student clicks "Submit to Teacher" → their name, band, scores, and all text responses are sent to your Sheet
- Each submission is one row with a timestamp
- The "Responses JSON" column contains all their written answers (you can expand these later)
- Students can submit multiple times — each submission adds a new row
- No login required for students

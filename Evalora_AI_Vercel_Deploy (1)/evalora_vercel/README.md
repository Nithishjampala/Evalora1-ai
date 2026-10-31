# Evalora AI — Vercel-only deployment

This project deploys the four-phase Project Evaluation workflow entirely on Vercel. It uses:

- Static HTML frontend in `public/`
- Python FastAPI Function in `api/index.py`
- Node.js Vercel Function in `api/upload.js` to authorize secure client uploads
- **Private Vercel Blob** for uploaded guidelines and student workbook
- Gemini for the five-question evaluation per phase
- Google Drive API to read student-submitted Drive links
- Google Sheets API to append evaluation history

The browser uploads files directly to Vercel Blob. The Vercel Functions receive only small JSON metadata and Blob URLs, avoiding the Vercel Function 4.5 MB request-body limit. Vercel Blob private storage is used so uploaded student files are not publicly readable.

## 1. Create and connect Private Vercel Blob

1. Open your Vercel project.
2. Go to **Storage** → **Create Database** → **Blob**.
3. Select **Private** access (not Public).
4. Create the store and connect it to this Vercel project.
5. Confirm `BLOB_READ_WRITE_TOKEN` is present in the project's Environment Variables. Vercel normally adds it automatically.

## 2. Configure Google Cloud

1. Enable **Google Sheets API** and **Google Drive API**.
2. Create a Google service account and download its JSON key.
3. Share this spreadsheet with the service account's `client_email` as **Editor**:
   `https://docs.google.com/spreadsheets/d/1L8FuY_CFURGrRkWbEFAG_QsHTyF0oUJYfpCfNAl-sOw/edit`
4. Share student Google Drive documents (or their containing folder) with the same service-account email as Viewer.

## 3. Configure Vercel Environment Variables

Add these under **Project → Settings → Environment Variables**:

- `GEMINI_API_KEY`: your Gemini API key
- `GEMINI_MODEL`: one model name available to your account; default is `gemini-3.5-flash-lite`
- `GOOGLE_SHEET_ID`: `1L8FuY_CFURGrRkWbEFAG_QsHTyF0oUJYfpCfNAl-sOw`
- `GOOGLE_SERVICE_ACCOUNT_JSON`: the complete service-account JSON, pasted as one-line JSON
- `BLOB_READ_WRITE_TOKEN`: provided when you connect the Blob store; add manually only if it is not present

Never commit `.env` or service-account credentials to GitHub. `.env.example` is only a template.

## 4. Deploy

1. Upload the contents of this project folder to a GitHub repository.
2. Import the repository into Vercel.
3. Keep the project root as the repository root.
4. Deploy. Vercel installs the Node dependency from `package.json` and Python dependencies from `requirements.txt`.
5. After adding or changing environment variables, redeploy.
6. Open `https://YOUR-VERCEL-DOMAIN/api/health`. It should return JSON with `ok: true`, `blob_configured: true` and your configured model.

## 5. Student Excel format

Upload an `.xlsx` file. Required column: `Student Name` or `Name`.

Optional columns: `Project ID`, `Project Name`.

Phase link columns may be named `Ideation`, `Ideation Link`, `Phase 1`, and equivalent common names for Design, Build and Demo. Values must be Google Drive file URLs or file IDs. Empty links are skipped.

## 6. Google Sheet output

The application creates an `Evaluation History` tab if needed and appends 11 columns:

1. Name
2. Project ID
3. Project Name
4. Phase Type
5. Questions
6. AI Confidence Score
7. Feedback
8. Suggestions
9. Overall Score
10. Model Name
11. Answer Accuracy (%)

Each phase is evaluated using exactly five scored questions. Overall score is the arithmetic mean of those five ratings. Additional feedback is one overall candidate-facing feedback and does not affect the score. Answer Accuracy (%) is a model estimate of evidence support, not independently verified accuracy.

## 7. Supported uploads and limits

- Guideline files: PDF, DOCX, PPTX, TXT, MD, CSV
- Student workbook: XLSX
- Maximum individual upload configured by the upload-token route: 50 MB
- Files are uploaded to a **private** Blob store, then read by the Python Function with the Blob token.
- The Google Drive student links must be accessible to the configured service account.
- Scanned PDFs need OCR; this build extracts embedded text only.
- Very large student batches can exceed the maximum duration for your Vercel plan. Split large workbooks into smaller batches if needed.

## Current scope

This Vercel package implements the four-phase Project Evaluation and Google Sheets reporting workflow. The prior Gradio Video Analyzer is not part of this build. Video processing should be added as a separate flow after confirming the selected Gemini model's video support and the desired upload limits.

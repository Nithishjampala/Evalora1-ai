# Evalora AI — Vercel-only deployment (without Blob token)

This version removes Vercel Blob completely. It includes Project Evaluation (Ideation, Design, Build, Demo), Google Sheets output, and Video Analyzer.

## Important upload limits

Vercel Functions impose a request-body limit of approximately 4.5 MB. To stay below it, the app limits the combined five project-evaluation uploads to 4 MB. Uploaded video files are limited to 3 MB. For larger videos, use a supported public YouTube, Vimeo, or Dailymotion URL (subject to source availability and Vercel execution limits).

## Required Vercel environment variables

- `GEMINI_API_KEY` — Gemini API key
- `GEMINI_MODEL` — configured Gemini model (default: `gemini-3.5-flash-lite`)
- `GOOGLE_SHEET_ID` — `1L8FuY_CFURGrRkWbEFAG_QsHTyF0oUJYfpCfNAl-sOw`
- `GOOGLE_SERVICE_ACCOUNT_JSON` — complete service-account JSON on one line

No `BLOB_READ_WRITE_TOKEN`, Blob store, or `api/upload.js` is needed.

## Google setup

Enable Google Drive API and Google Sheets API in the service account's Google Cloud project. Share the spreadsheet with the service account email as Editor. Share student Drive documents with that service account with read access.

## Deploy

1. Push the repository root containing `pyproject.toml`, `vercel.json`, `requirements.txt`, `api/`, and `public/` to GitHub.
2. Import the repository into Vercel.
3. Add the four environment variables above in Vercel Project Settings → Environment Variables.
4. Redeploy.
5. Test `/api/health`.

## Google Sheets output

The backend appends to `Evaluation History` with 11 columns: Name, Project ID, Project Name, Phase Type, Questions, AI Confidence Score, Feedback, Suggestions, Overall Score, Model Name, Answer Accuracy (%). Answer Accuracy is an AI-estimated evidence-support score, not independent verification.

## Limitations

- All five project evaluation files must total 4 MB or less.
- Uploaded video files must be 3 MB or less; larger videos should be supplied through a supported public URL.
- Video processing requires Gemini model video support and must complete within Vercel's configured function duration.
- Scanned PDFs may need OCR before their text can be evaluated.

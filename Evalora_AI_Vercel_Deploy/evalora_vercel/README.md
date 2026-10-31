# Evalora AI — Vercel deployment package

This package deploys the four-phase project evaluation workflow to **Vercel only**. It uses a static HTML frontend and a Python FastAPI Vercel Function. It appends completed phase evaluations to the supplied Google Sheet.

## Included
- Ideation, Design, Build and Demo guideline uploads (all four required for each evaluation run)
- Student Excel upload with phase-specific Google Drive links
- Gemini evaluation using one configured model and no model fallback
- Google Drive document reading (PDF, DOCX, PPTX, TXT, MD, CSV)
- Google Sheets `Evaluation History` append with exactly 11 columns
- Empty phase links are skipped; individual failed links are reported

## Before deployment
1. In Google Cloud Console, enable **Google Sheets API** and **Google Drive API**.
2. Create a service account and download its JSON key.
3. Share the target spreadsheet with the service account's `client_email` as **Editor**.
4. Share every student Google Drive document with that same service-account email (Viewer is sufficient), or place them in a folder shared with it.
5. Ensure the Gemini API key has access to the configured model. This package deliberately does not fall back to another model.

## Deploy to Vercel
1. Extract this ZIP.
2. Push the contents of the `evalora_vercel` folder to a new GitHub repository.
3. In Vercel, choose **Add New → Project**, import that repository, and deploy.
4. In **Project → Settings → Environment Variables**, add:
   - `GEMINI_API_KEY` = your Gemini API key
   - `GEMINI_MODEL` = `gemini-3.5-flash-lite` (or another model you have verified is available; keep this to one model)
   - `GOOGLE_SHEET_ID` = `1L8FuY_CFURGrRkWbEFAG_QsHTyF0oUJYfpCfNAl-sOw`
   - `GOOGLE_SERVICE_ACCOUNT_JSON` = the complete service-account JSON as one-line JSON
5. Redeploy after adding variables.
6. Open `https://YOUR-VERCEL-DOMAIN/api/health`. It should return `ok: true`.

## Excel format
Required student name column: `Student Name` or `Name`.
Optional: `Project ID`, `Project Name`.
Phase link columns can be named `Ideation`, `Design`, `Build`, `Demo` (or common variants such as `Ideation Link`, `Phase 1`, etc.).

Each link must be a Google Drive file URL or file ID. The file must be accessible to the service account. Empty links are skipped.

## Google Sheet output
The backend creates an `Evaluation History` tab if it does not exist and writes these columns:

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

`Answer Accuracy (%)` is a model-estimated measure of how well the evaluation statements are supported by the submitted evidence. It is not independently verified accuracy.

## Important Vercel limits / current scope
- This package is designed for Vercel Functions, not the original Colab/Gradio runtime.
- The web request uploads four guideline files and one Excel file. Vercel Functions have a request-body size limit (commonly 4.5 MB); keep the combined multipart request below the limit. Large documents may need a Drive-based upload workflow. If uploads fail with a payload-size error, store the guideline documents in Drive and extend the UI to select Drive file IDs instead.
- Evaluations run in one request. Long batches may exceed your Vercel plan's maximum function duration. For large batches, split the student Excel into smaller batches.
- Vercel serverless local storage is temporary. This deployment writes evaluation output to Google Sheets, not a local SQLite database.
- The previous Gradio Video Analyzer is **not included** in this first Vercel package. Video uploads and long video processing need a separate architecture that respects Vercel request and runtime limits. The existing four-phase project evaluation workflow is included.
- The Gemini model name from the prior notebook is retained as the default, but availability and video/document capability depend on your Google AI account. Check `/api/health` and Vercel logs if the model is unavailable.

## Security
Never commit `.env` files, service-account keys, or Gemini API keys to GitHub. Configure secrets only in Vercel Environment Variables. If a service-account key is exposed, revoke and rotate it in Google Cloud Console.

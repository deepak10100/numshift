# NumShift

NumShift is a small Python web app for text and number replacement, country-code tools, emoji editing, document text extraction, and result export.

## Deploy on Render

This repository is set up for a Render **Web Service using Docker**.

1. Push the project files to a GitHub repository.
2. In Render, choose **New → Blueprint** and connect that repository, or create a Docker Web Service from it.
3. Render reads `render.yaml` and builds the included `Dockerfile`.
4. When the deploy is live, open the generated `onrender.com` URL.

The service binds to `0.0.0.0` and uses Render's `PORT` environment variable. The Docker image installs Poppler, Tesseract English/Spanish language data, and DejaVu Sans, which are needed for scanned-PDF OCR and PDF generation.

## Included files

- `index.html` — NumShift home page and editing tool.
- `features.html`, `how-it-works.html`, `faq.html`, `other-tools.html` — additional site pages.
- `app_server.py` — upload extraction and PDF export endpoints.
- `requirements.txt` — Python dependencies.
- `Dockerfile` and `render.yaml` — Render deployment configuration.

## Runtime notes

- Maximum upload size: 20 MB.
- Scanned PDF OCR: English and Spanish, at most the first 20 pages.
- Supported extraction formats: TXT, MD, CSV, JSON, HTML, PDF, DOCX and XLSX.
- Generated PDFs are currently text re-typesets; the original PDF's exact page design is not retained.
- Each generated PDF is written to a temporary export directory and served through a per-export unguessable link; generated PDFs are not exposed as public static files. Render is configured to use `/tmp/numshift-exports`, which is ephemeral and cleared when the service restarts or redeploys. Add a secured persistent disk and set `EXPORT_DIR` if exports must survive restarts.

## Run locally

```bash
python -m pip install -r requirements.txt
python app_server.py
```

Open `http://localhost:4173`. For OCR locally, install Poppler and Tesseract with `eng` and `spa` language data.

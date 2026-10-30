# 📊 SlideCheck AI – Smart Presentation Analyzer

SlideCheck AI is an AI-powered **presentation analysis tool**. You upload an already-created presentation (`.pptx` or `.pdf`) and the app evaluates its quality, gives scores out of 100, points out strengths and mistakes (with slide/page numbers), and suggests how to improve it. The result can be downloaded as a PDF report.

It is inspired by an *AI Resume Checker*, but it checks **presentations** instead of resumes.

> **Important:** SlideCheck AI is **only an analyzer**. It does **not** generate presentations, does **not** create or edit PowerPoint slides, and does **not** return a corrected file. It only gives feedback.

---

## ✨ Features

1. Upload PowerPoint (`.pptx`) and PDF (`.pdf`) files
2. Extract text from every slide / page
3. Analyze the presentation with **Google Gemini AI**
4. Overall presentation score out of **100**
5. Category-wise scores with progress bars
6. Ten analysis areas:
   - Content Quality
   - Grammar and Spelling
   - Slide Structure and Organization
   - Topic Relevance
   - Readability
   - Repeated Information
   - Missing Important Information
   - Text Density
   - Introduction and Conclusion Quality
   - Design and Formatting (only from facts that can really be inspected)
7. Strengths and weaknesses, with the relevant slide / page number
8. Clear improvement suggestions with priorities
9. Short overall analysis summary
10. Downloadable **PDF report** (ReportLab)
11. Honest limitations: the app clearly states what it could **not** inspect

---

## 🧠 How the scoring works (no random scores)

- The program first **measures** real facts itself: words per slide, crowded slides, almost-empty slides, small fonts, repeated lines, number of images/tables/charts.
- These facts and the extracted text are sent to Gemini, which returns a structured JSON evaluation for the ten categories.
- The app **validates** the JSON (scores are forced into 0–100) and calculates the **overall score as a weighted average** of the category scores.
- The *Text Density* score is the average of Gemini's judgement and a measured word-count formula.
- If there is too little information for *Design and Formatting*, that category is marked **"Not scored"** and left out of the overall score.

| Category | Weight |
|---|---|
| Content Quality | 15 |
| Slide Structure & Organization | 12 |
| Grammar & Spelling | 10 |
| Topic Relevance | 10 |
| Readability | 10 |
| Missing Important Information | 10 |
| Text Density | 10 |
| Repeated Information | 8 |
| Introduction & Conclusion | 8 |
| Design & Formatting | 7 |

---

## ⚠️ What the app can and cannot inspect

| Inspected | NOT inspected |
|---|---|
| Text on every slide/page | Colours, contrast, alignment, spacing |
| Word counts, repeated lines | Overlapping objects, image quality |
| Number of images, tables, charts (PPTX) | Text inside images / scanned pages (no OCR) |
| Font sizes and font names set in the file | Animations, transitions, audio, video |
| Slide titles and speaker notes (PPTX) | Tables/charts inside PDFs (look like normal text) |

Slides are **not rendered as pictures**, so visual design is judged only from structural facts, and the report says so.

---

## 🛠️ Technologies

| Technology | Purpose |
|---|---|
| Python 3.10+ | Programming language |
| Streamlit | Web interface |
| Google Gemini API (`google-genai`) | AI evaluation |
| python-pptx | Reading `.pptx` files |
| PyMuPDF | Reading `.pdf` files |
| ReportLab | Creating the PDF report |

---

## 📁 Project Structure

```
SlideCheck-AI/
├── app.py                          # Complete Streamlit application
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .gitignore                      # Keeps secrets out of Git
├── Project_Documentation.docx      # University project document
└── .streamlit/
    └── secrets.toml.example        # Template for the API key (safe to commit)
```

---

## 🚀 Installation (run locally)

**1. Install Python 3.10 or newer** from https://www.python.org

**2. Get a free Gemini API key** from https://aistudio.google.com/apikey

**3. Download the project and open a terminal inside the folder**

**4. (Recommended) create a virtual environment**
```bash
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS / Linux
```

**5. Install the dependencies**
```bash
pip install -r requirements.txt
```

**6. Configure the API key** (choose ONE option)

*Option A – Streamlit secrets file (recommended)*

Create a folder named `.streamlit` and inside it a file named `secrets.toml`:
```toml
GEMINI_API_KEY = "your-real-key-here"
```
(You can copy `.streamlit/secrets.toml.example` and edit it.)

*Option B – environment variable*
```bash
set GEMINI_API_KEY=your-real-key-here           # Windows (cmd)
$env:GEMINI_API_KEY="your-real-key-here"        # Windows (PowerShell)
export GEMINI_API_KEY="your-real-key-here"      # macOS / Linux
```

**7. Run the app**
```bash
streamlit run app.py
```
The app opens at http://localhost:8501

---

## 📖 How to use

1. Open the app.
2. Upload a `.pptx` or `.pdf` presentation (max 25 MB, first 100 slides/pages are analysed).
3. Optionally type the presentation topic and target audience (this improves the relevance check).
4. Click **Analyze Presentation**.
5. Read the overall score, category scores, strengths, weaknesses and suggestions.
6. Click **Download Report as PDF** to save the report.

---

## 🔐 API key security

- The key is **never written inside `app.py`**. The app reads it from Streamlit secrets or the `GEMINI_API_KEY` environment variable.
- `.gitignore` excludes `.streamlit/secrets.toml` and `.env` files, so they are not uploaded to GitHub.
- Never paste your key into a README, issue, screenshot or chat.
- If a key is ever exposed, delete it in Google AI Studio and create a new one.

---

## 🌐 Upload to GitHub

```bash
git init
git add .
git commit -m "SlideCheck AI - first version"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/SlideCheck-AI.git
git push -u origin main
```
Before pushing, run `git status` and make sure **`secrets.toml` is NOT listed**.

---

## ☁️ Deploy on Streamlit Community Cloud

1. Push the project to a GitHub repository.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app** → choose your repository, branch `main` and main file `app.py`.
4. Open **Advanced settings → Secrets** and paste:
   ```toml
   GEMINI_API_KEY = "your-real-key-here"
   ```
5. Click **Deploy**. After a few minutes your app gets a public link.

To change the key later: *App settings → Secrets*.

---

## 🧯 Troubleshooting

| Problem | Solution |
|---|---|
| "No Gemini API key was found" | Add `GEMINI_API_KEY` to secrets or environment variables, then restart. |
| "Gemini rejected the API key" | Check the key is copied correctly and active. |
| "Usage limit was reached" | Wait a minute (free tier limits) and try again. |
| "Almost no text could be extracted" | The file is probably made of images or scans. OCR is not supported. |
| Model not found | Set `GEMINI_MODEL` in secrets to a model available to your key. |

---

## 📌 Limitations

- AI feedback can contain mistakes; treat it as guidance, not a final grade.
- No OCR: text inside images is not read.
- Visual design is not rendered or inspected, only structural facts are used.
- Very long presentations are limited to the first 100 slides/pages.

---

## 🎓 Project Overview

SlideCheck AI is a Data Science mini-project that combines **data extraction** (python-pptx, PyMuPDF), **feature engineering** (word counts, density, repeated lines, font sizes), **AI-based text analysis** (Gemini), **score aggregation** (weighted average) and **report generation** (ReportLab) inside a simple web app (Streamlit).

**Project name:** SlideCheck AI – Smart Presentation Analyzer

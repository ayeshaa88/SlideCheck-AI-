"""
SlideCheck AI - Smart Presentation Analyzer
===========================================

A Streamlit app that ANALYZES an existing presentation (.pptx or .pdf)
and gives feedback. It does NOT create, edit or correct presentations.

How the program works (simple version):
    1. The user uploads a .pptx or .pdf file.
    2. We extract the text (and a few measurable facts) from every slide/page.
    3. We calculate simple statistics ourselves (word counts, repeated lines...).
    4. We send the extracted text + statistics to Google Gemini and ask for a
       structured JSON evaluation.
    5. We validate the JSON, calculate the overall score and show the results.
    6. The user can download the evaluation as a PDF report (ReportLab).

Run with:   streamlit run app.py
"""

import datetime
import html
import io
import json
import os
import re

import fitz  # PyMuPDF - reads PDF files
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation  # python-pptx - reads PowerPoint files
from pptx.enum.shapes import MSO_SHAPE_TYPE
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ---------------------------------------------------------------------------
# 1. SETTINGS
# ---------------------------------------------------------------------------
APP_NAME = "SlideCheck AI"
APP_SUBTITLE = "Smart Presentation Analyzer"

DEFAULT_MODEL = "gemini-2.5-flash"  # can be changed with the GEMINI_MODEL secret
MAX_FILE_MB = 25          # biggest file we accept
MAX_UNITS = 100           # maximum number of slides/pages analysed
MAX_PROMPT_CHARS = 60000  # maximum amount of slide text sent to Gemini
MIN_WORDS = 20            # fewer words than this = nothing to analyse
DENSE_WORDS = 120         # a slide/page with more words than this is "dense"
EMPTY_WORDS = 3           # a slide/page with fewer words than this is "empty"
SMALL_FONT_PT = 12        # fonts smaller than this are hard to read on a slide

# (key, label, weight, icon). Weights add up to 100.
CATEGORIES = [
    ("content_quality", "Content Quality", 15, "💡"),
    ("grammar_spelling", "Grammar & Spelling", 10, "✍️"),
    ("structure", "Slide Structure & Organization", 12, "🧱"),
    ("topic_relevance", "Topic Relevance", 10, "🎯"),
    ("readability", "Readability", 10, "👓"),
    ("repeated_information", "Repeated Information", 8, "🔁"),
    ("missing_information", "Missing Important Information", 10, "🧩"),
    ("text_density", "Text Density", 10, "📏"),
    ("intro_conclusion", "Introduction & Conclusion", 8, "🚪"),
    ("design_formatting", "Design & Formatting", 7, "🎨"),
]
CATEGORY_LABELS = {key: label for key, label, _, _ in CATEGORIES}


class AnalysisError(Exception):
    """An error whose message is safe and friendly to show to the user."""


class InvalidResponseError(AnalysisError):
    """Gemini answered, but the answer was not valid/usable JSON."""


# ---------------------------------------------------------------------------
# 2. SECRETS (API KEY) - never hard-coded
# ---------------------------------------------------------------------------
def get_secret(name, default=""):
    """Read a value from Streamlit secrets first, then environment variables."""
    value = ""
    try:
        value = st.secrets.get(name, "")
    except Exception:  # no secrets file exists - that is fine
        value = ""
    return str(value or os.environ.get(name, default) or default).strip()


# ---------------------------------------------------------------------------
# 3. TEXT EXTRACTION
# ---------------------------------------------------------------------------
def count_words(text):
    return len(re.findall(r"\b\w+\b", text))


def approx_min_font(sizes):
    """Approximate smallest font: the 10th percentile (ignores tiny footers)."""
    if not sizes:
        return None
    sizes = sorted(sizes)
    return round(sizes[len(sizes) // 10], 1)


def safe_shape_type(shape):
    try:
        return shape.shape_type
    except Exception:
        return None


def iter_shapes(shapes):
    """Yield every shape, including shapes that are inside groups."""
    for shape in shapes:
        yield shape
        if safe_shape_type(shape) == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def extract_pptx(file_bytes):
    """Read a .pptx file with python-pptx and collect text + simple facts."""
    try:
        prs = Presentation(io.BytesIO(file_bytes))
    except Exception:
        raise AnalysisError(
            "The PowerPoint file could not be opened. It may be corrupted, "
            "password-protected, or not a real .pptx file."
        )

    total_slides = len(prs.slides)
    units = []
    fonts_used = set()

    for number, slide in enumerate(prs.slides, start=1):
        if number > MAX_UNITS:
            break
        lines, sizes = [], []
        images = tables = charts = 0

        for shape in iter_shapes(slide.shapes):
            if safe_shape_type(shape) == MSO_SHAPE_TYPE.PICTURE:
                images += 1
            if getattr(shape, "has_chart", False) and shape.has_chart:
                charts += 1
            if getattr(shape, "has_table", False) and shape.has_table:
                tables += 1
                for row in shape.table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        lines.append(" | ".join(cells))
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        lines.append(text)
                    for run in paragraph.runs:
                        if run.font.size is not None and run.text.strip():
                            sizes.append(run.font.size.pt)
                        if run.font.name:
                            fonts_used.add(run.font.name)

        title = ""
        try:
            if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
                title = slide.shapes.title.text_frame.text.strip()
        except Exception:
            title = ""

        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            notes = ""

        text = "\n".join(lines)
        units.append({
            "number": number, "text": text, "title": title, "notes": notes,
            "word_count": count_words(text), "images": images,
            "tables": tables, "charts": charts, "min_font": approx_min_font(sizes),
        })

    return {
        "file_type": "PPTX", "unit": "Slide", "units": units,
        "total_units": total_slides, "fonts_used": sorted(fonts_used),
        "design_note": (
            "Design & Formatting was judged from structural facts only: number of pictures, "
            "tables and charts, slide titles, and font sizes/names that are explicitly set in "
            "the file. The slides were NOT rendered as images, so colours, contrast, alignment, "
            "spacing, overlapping objects and image quality were NOT inspected."
        ),
    }


def extract_pdf(file_bytes):
    """Read a .pdf file with PyMuPDF and collect text + simple facts."""
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception:
        raise AnalysisError("The PDF file could not be opened. It may be corrupted or not a real PDF.")
    if doc.needs_pass:
        doc.close()
        raise AnalysisError("This PDF is password-protected. Please remove the password and try again.")

    total_pages = doc.page_count
    units = []
    fonts_used = set()

    for index in range(min(total_pages, MAX_UNITS)):
        page = doc[index]
        text = page.get_text("text").strip()
        images = len(page.get_images(full=True))

        # Normalise font sizes to a standard 540 pt-high slide so that the
        # "small font" check works for PDFs exported from slide software.
        shortest_side = min(page.rect.width, page.rect.height) or 540
        scale = 540 / shortest_side
        sizes = []
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span.get("text", "").strip():
                        sizes.append(span["size"] * scale)
                        fonts_used.add(span.get("font", ""))

        first_line = text.splitlines()[0].strip() if text else ""
        units.append({
            "number": index + 1, "text": text, "title": first_line, "notes": "",
            "word_count": count_words(text), "images": images, "tables": 0,
            "charts": 0, "min_font": approx_min_font(sizes),
        })
    doc.close()

    return {
        "file_type": "PDF", "unit": "Page", "units": units,
        "total_units": total_pages, "fonts_used": sorted(f for f in fonts_used if f),
        "design_note": (
            "Design & Formatting was judged from page-level facts only: font sizes, number of "
            "font families, and number of images per page. The pages were NOT visually "
            "rendered or viewed, so colours, contrast, alignment, spacing, overlapping objects "
            "and image quality were NOT inspected. Tables and charts inside a PDF cannot be "
            "told apart from normal text."
        ),
    }


# ---------------------------------------------------------------------------
# 4. OUR OWN STATISTICS (calculated in Python, not by AI)
# ---------------------------------------------------------------------------
def compute_stats(extracted):
    units = extracted["units"]
    word_counts = [u["word_count"] for u in units]
    total_words = sum(word_counts)
    n = len(units)
    label = extracted["unit"]

    # Lines that appear on more than one slide/page (maybe repeated content).
    line_map = {}
    for u in units:
        seen = set()
        for line in u["text"].splitlines():
            norm = re.sub(r"\s+", " ", line.strip().lower())
            if len(norm) >= 20 and norm not in seen:
                seen.add(norm)
                line_map.setdefault(norm, []).append(u["number"])
    repeated = [(line, nums) for line, nums in line_map.items() if len(nums) >= 2]
    repeated.sort(key=lambda item: len(item[1]), reverse=True)

    return {
        "unit_count": n,
        "total_words": total_words,
        "avg_words": round(total_words / n, 1) if n else 0,
        "max_words": max(word_counts) if word_counts else 0,
        "dense_units": [u["number"] for u in units if u["word_count"] > DENSE_WORDS],
        "empty_units": [u["number"] for u in units if u["word_count"] < EMPTY_WORDS],
        "small_font_units": [
            u["number"] for u in units
            if u["min_font"] is not None and u["min_font"] < SMALL_FONT_PT
        ],
        "image_units": sum(1 for u in units if u["images"] > 0),
        "total_images": sum(u["images"] for u in units),
        "total_tables": sum(u["tables"] for u in units),
        "total_charts": sum(u["charts"] for u in units),
        "font_family_count": len(extracted["fonts_used"]),
        "repeated_lines": repeated[:8],
        "label": label,
    }


def local_density_score(stats):
    """A simple, explainable text-density score (0-100) based on word counts."""
    n = max(stats["unit_count"], 1)
    score = 100.0
    score -= 60 * (len(stats["dense_units"]) / n)       # many crowded slides
    if stats["avg_words"] > 80:
        score -= min(25, (stats["avg_words"] - 80) / 2)  # high average
    score -= 20 * (len(stats["empty_units"]) / n)        # many (nearly) empty slides
    return int(max(0, min(100, round(score))))


# ---------------------------------------------------------------------------
# 5. GEMINI
# ---------------------------------------------------------------------------
def fmt_units(label, numbers):
    return ", ".join(f"{label} {n}" for n in numbers) if numbers else "none"


def build_prompt(extracted, stats, topic, audience):
    label = extracted["unit"]
    facts = [
        f"File type: {extracted['file_type']}",
        f"Number of {label.lower()}s analysed: {stats['unit_count']} (file has {extracted['total_units']})",
        f"Total words: {stats['total_words']}, average per {label.lower()}: {stats['avg_words']}, "
        f"highest: {stats['max_words']}",
        f"Dense {label.lower()}s (more than {DENSE_WORDS} words): {fmt_units(label, stats['dense_units'])}",
        f"Almost empty {label.lower()}s (fewer than {EMPTY_WORDS} words): {fmt_units(label, stats['empty_units'])}",
        f"{label}s with small fonts (under {SMALL_FONT_PT} pt, approximate): "
        f"{fmt_units(label, stats['small_font_units'])}",
        f"{label}s containing images: {stats['image_units']} (total images: {stats['total_images']}, "
        f"tables: {stats['total_tables']}, charts: {stats['total_charts']})",
        f"Font families found: {stats['font_family_count']}",
    ]
    if stats["repeated_lines"]:
        facts.append("Lines that appear on more than one slide/page (could be footers or genuine repetition):")
        for line, nums in stats["repeated_lines"]:
            facts.append(f'  - "{line[:100]}" on {label.lower()}s {nums}')

    # The actual slide text
    parts = []
    for u in extracted["units"]:
        header = (f"[{label} {u['number']}] words={u['word_count']} images={u['images']} "
                  f"tables={u['tables']} charts={u['charts']} min_font={u['min_font']}")
        body = u["text"] if u["text"] else "(no text found)"
        if u["notes"]:
            body += f"\nSpeaker notes: {u['notes'][:300]}"
        parts.append(f"{header}\n{body}")
    slide_text = "\n\n".join(parts)
    truncated_note = ""
    if len(slide_text) > MAX_PROMPT_CHARS:
        slide_text = slide_text[:MAX_PROMPT_CHARS]
        truncated_note = "\n(NOTE: the text was cut because the presentation is very long.)"

    category_lines = "\n".join(f'  - "{key}": {name}' for key, name, _, _ in CATEGORIES)

    return f"""You are a strict but fair university presentation reviewer. Evaluate the presentation below.
You ONLY analyse. You must NOT rewrite the presentation or produce corrected slides.

USER INFORMATION
- Intended topic (may be empty): {topic or "not provided - infer it from the content"}
- Target audience (may be empty): {audience or "not provided"}

MEASURED FACTS (calculated by the program, they are reliable)
{chr(10).join(facts)}

DESIGN INSPECTION LIMITS
{extracted['design_note']}

RULES
1. Base every score and comment ONLY on the extracted text and measured facts. Do not invent slide text.
2. Scores are integers from 0 to 100. Use the full range honestly (90+ = excellent, 70-89 = good,
   50-69 = average, below 50 = weak). Do not default to round or identical scores.
3. Always mention the location as "{label} N" (for example "{label} 4") when an issue belongs to a specific
   {label.lower()}. Use "General" only for presentation-wide points.
4. For "design_formatting": judge ONLY from the measured facts above. If there is too little information,
   set its score to null. Never claim to have seen colours, layout, alignment or images.
5. "repeated_information" score: high score = little unnecessary repetition.
   "missing_information" score: high score = nothing important is missing.
   "text_density" score: high score = comfortable amount of text per {label.lower()}.
6. Suggestions must describe WHAT to improve and WHY in 1-2 sentences. Do not write replacement slides.
7. Keep every comment short (max 2 sentences). Write in English.
8. Return ONLY valid JSON, with no markdown fences and no extra text.

CATEGORY KEYS (all are required inside "category_scores"):
{category_lines}

JSON FORMAT
{{
  "detected_topic": "short topic of the presentation",
  "category_scores": {{ "<category key>": {{"score": 0-100 or null, "comment": "..."}} }},
  "strengths": [{{"title": "...", "detail": "...", "location": "{label} N or General"}}],
  "weaknesses": [{{"title": "...", "detail": "...", "location": "{label} N or General", "severity": "high|medium|low"}}],
  "grammar_issues": [{{"location": "{label} N", "issue": "the problem", "fix": "how to fix it"}}],
  "repeated_items": [{{"description": "what is repeated", "location": "{label} N, {label} M"}}],
  "missing_items": [{{"item": "important information that is missing"}}],
  "suggestions": [{{"suggestion": "...", "location": "{label} N or General", "priority": "high|medium|low"}}],
  "summary": "3-4 sentence overall analysis"
}}

PRESENTATION TEXT
{slide_text}{truncated_note}
"""


def friendly_api_error(exc):
    """Turn a raw API exception into a message a student can understand."""
    msg = str(exc)
    low = msg.lower()
    if "api key" in low or "api_key" in low or "permission_denied" in low or "401" in low or "403" in low:
        return "Gemini rejected the API key. Please check that GEMINI_API_KEY is correct and active."
    if "429" in low or "quota" in low or "resource_exhausted" in low or "rate" in low:
        return "The Gemini usage limit was reached. Please wait a minute and try again."
    if "503" in low or "unavailable" in low or "overloaded" in low:
        return "The Gemini service is busy right now. Please try again in a moment."
    if "404" in low or "not found" in low:
        return "The selected Gemini model was not found. Set a valid GEMINI_MODEL in your secrets."
    if "timeout" in low or "connect" in low or "network" in low:
        return "Could not reach the Gemini service. Please check your internet connection."
    return "The Gemini request failed: " + msg[:200]


def call_gemini(prompt, api_key, model):
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2, response_mime_type="application/json"
            ),
        )
        text = response.text
    except Exception as exc:
        raise AnalysisError(friendly_api_error(exc))
    if not text or not text.strip():
        raise InvalidResponseError("Gemini returned an empty answer (the content may have been blocked).")
    return text


def parse_json(text):
    """Convert Gemini's answer into a Python dict (tolerates ```json fences)."""
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.IGNORECASE)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start != -1 and end > start:
            try:
                return json.loads(cleaned[start:end + 1])
            except json.JSONDecodeError:
                pass
    raise InvalidResponseError("Gemini's answer was not valid JSON.")


# ---------------------------------------------------------------------------
# 6. CHECK + CLEAN THE AI ANSWER
# ---------------------------------------------------------------------------
def to_score(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        return max(0, min(100, int(round(float(value)))))
    except (TypeError, ValueError):
        return None


def to_text(value):
    if isinstance(value, list):
        return ", ".join(str(v) for v in value)
    return "" if value is None else str(value).strip()


def clean_items(value, keys, max_items=12):
    """Make sure a list from Gemini is a list of dicts with the expected keys."""
    items = []
    if isinstance(value, list):
        for entry in value[:max_items]:
            if isinstance(entry, dict):
                item = {k: to_text(entry.get(k)) for k in keys}
            elif isinstance(entry, str) and entry.strip():
                item = {k: "" for k in keys}
                item[keys[0]] = entry.strip()
            else:
                continue
            if item[keys[0]] or (len(keys) > 1 and item[keys[1]]):
                items.append(item)
    return items


def normalize_result(data):
    if not isinstance(data, dict):
        raise InvalidResponseError("Gemini's answer had an unexpected structure.")
    raw_scores = data.get("category_scores")
    if not isinstance(raw_scores, dict):
        raise InvalidResponseError("Gemini's answer did not contain category scores.")

    categories = {}
    for key, _, _, _ in CATEGORIES:
        entry = raw_scores.get(key)
        if isinstance(entry, dict):
            score, comment = to_score(entry.get("score")), to_text(entry.get("comment"))
        else:
            score, comment = to_score(entry), ""
        categories[key] = {"score": score, "comment": comment}

    if sum(1 for c in categories.values() if c["score"] is not None) < 6:
        raise InvalidResponseError("Gemini's answer was incomplete (too few category scores).")

    weaknesses = clean_items(data.get("weaknesses"), ["title", "detail", "location", "severity"])
    for w in weaknesses:
        w["severity"] = w["severity"].lower() if w["severity"].lower() in ("high", "medium", "low") else "medium"
    order = {"high": 0, "medium": 1, "low": 2}
    weaknesses.sort(key=lambda w: order[w["severity"]])

    suggestions = clean_items(data.get("suggestions"), ["suggestion", "location", "priority"])
    for s in suggestions:
        s["priority"] = s["priority"].lower() if s["priority"].lower() in ("high", "medium", "low") else "medium"
    suggestions.sort(key=lambda s: order[s["priority"]])

    return {
        "detected_topic": to_text(data.get("detected_topic")) or "Not detected",
        "categories": categories,
        "strengths": clean_items(data.get("strengths"), ["title", "detail", "location"]),
        "weaknesses": weaknesses,
        "grammar_issues": clean_items(data.get("grammar_issues"), ["location", "issue", "fix"], 15),
        "repeated_items": clean_items(data.get("repeated_items"), ["description", "location"]),
        "missing_items": clean_items(data.get("missing_items"), ["item"]),
        "suggestions": suggestions,
        "summary": to_text(data.get("summary")) or "No summary was returned.",
    }


def blend_text_density(result, stats):
    """Final text-density score = average of AI judgement and our word-count check."""
    local = local_density_score(stats)
    ai = result["categories"]["text_density"]["score"]
    final = local if ai is None else int(round((ai + local) / 2))
    note = (f"Final score = average of the AI judgement ({ai if ai is not None else 'n/a'}) and a "
            f"measured word-count check ({local}).")
    comment = result["categories"]["text_density"]["comment"]
    result["categories"]["text_density"] = {"score": final, "comment": (comment + " " + note).strip()}


def compute_overall(categories):
    """Weighted average of the categories that actually received a score."""
    total_weight = score_sum = 0
    for key, _, weight, _ in CATEGORIES:
        score = categories[key]["score"]
        if score is not None:
            total_weight += weight
            score_sum += weight * score
    return int(round(score_sum / total_weight)) if total_weight else 0


def grade_label(score):
    if score >= 85:
        return "Excellent"
    if score >= 70:
        return "Good"
    if score >= 55:
        return "Fair"
    if score >= 40:
        return "Needs Improvement"
    return "Weak"


def analyze_presentation(file_name, file_bytes, topic, audience, api_key, model):
    """The full pipeline: extract -> statistics -> Gemini -> validate -> score."""
    ext = os.path.splitext(file_name)[1].lower()
    if ext == ".pptx":
        extracted = extract_pptx(file_bytes)
    elif ext == ".pdf":
        extracted = extract_pdf(file_bytes)
    else:
        raise AnalysisError("Unsupported file type. Please upload a .pptx or .pdf file.")

    if not extracted["units"]:
        raise AnalysisError("The file contains no slides or pages.")
    stats = compute_stats(extracted)
    if stats["total_words"] < MIN_WORDS:
        raise AnalysisError(
            "Almost no text could be extracted. The file may contain only images or scanned pages. "
            "SlideCheck AI cannot read text inside images (no OCR)."
        )

    prompt = build_prompt(extracted, stats, topic, audience)
    last_error = None
    for _ in range(2):  # one retry if the JSON is invalid
        try:
            result = normalize_result(parse_json(call_gemini(prompt, api_key, model)))
            break
        except InvalidResponseError as exc:
            last_error = exc
    else:
        raise AnalysisError(f"{last_error} Please try again.")

    blend_text_density(result, stats)
    result["overall_score"] = compute_overall(result["categories"])
    result["grade"] = grade_label(result["overall_score"])
    result["design_note"] = extracted["design_note"]
    if result["categories"]["design_formatting"]["score"] is None:
        result["categories"]["design_formatting"]["comment"] = (
            result["categories"]["design_formatting"]["comment"] or "Not scored: too little design information."
        )

    meta = {
        "file_name": file_name, "file_type": extracted["file_type"], "unit": extracted["unit"],
        "units_analysed": stats["unit_count"], "total_units": extracted["total_units"],
        "total_words": stats["total_words"], "model": model,
        "date": datetime.date.today().strftime("%d %B %Y"), "topic_input": topic,
    }
    return result, meta, stats


# ---------------------------------------------------------------------------
# 7. PDF REPORT (ReportLab)
# ---------------------------------------------------------------------------
def pdf_text(value):
    """Escape text for ReportLab and replace characters the base font cannot draw."""
    text = "" if value is None else str(value)
    text = text.encode("cp1252", "replace").decode("cp1252")
    return html.escape(text)


def score_color(score):
    if score is None:
        return colors.HexColor("#b8aedd")
    if score >= 75:
        return colors.HexColor("#7c3aed")
    if score >= 50:
        return colors.HexColor("#60a5fa")
    return colors.HexColor("#f472b6")


def bar_drawing(score, width=150, height=9):
    d = Drawing(width, height)
    d.add(Rect(0, 0, width, height, rx=4, ry=4, fillColor=colors.HexColor("#ece6fb"), strokeColor=None))
    if score:
        d.add(Rect(0, 0, width * score / 100, height, rx=4, ry=4, fillColor=score_color(score), strokeColor=None))
    return d


def build_pdf_report(result, meta, stats):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=16 * mm, bottomMargin=18 * mm,
        title=f"{APP_NAME} Report", author=APP_NAME,
    )
    base = getSampleStyleSheet()
    purple, ink, soft = "#4c1d95", "#2d2150", "#6b5fa0"
    title = ParagraphStyle("t", parent=base["Title"], textColor=colors.HexColor(purple), fontSize=22, leading=26)
    sub = ParagraphStyle("s", parent=base["Normal"], textColor=colors.HexColor(soft), fontSize=10, alignment=1)
    h2 = ParagraphStyle("h2", parent=base["Heading2"], textColor=colors.HexColor(purple), fontSize=14,
                        spaceBefore=12, spaceAfter=6)
    body = ParagraphStyle("b", parent=base["Normal"], textColor=colors.HexColor(ink), fontSize=10, leading=14)
    item = ParagraphStyle("i", parent=body, leftIndent=12, bulletIndent=2, spaceAfter=5)
    small = ParagraphStyle("sm", parent=body, fontSize=8.5, leading=12, textColor=colors.HexColor(soft))

    story = [
        Paragraph(f"{APP_NAME} - {APP_SUBTITLE}", title),
        Paragraph(pdf_text(f"Presentation Analysis Report | {meta['date']}"), sub),
        Spacer(1, 10),
    ]

    # File information
    info = [
        ["File", pdf_text(meta["file_name"]), "Type", meta["file_type"]],
        [f"{meta['unit']}s analysed", f"{meta['units_analysed']} of {meta['total_units']}",
         "Total words", str(meta["total_words"])],
        ["Detected topic", Paragraph(pdf_text(result["detected_topic"]), body), "AI model", pdf_text(meta["model"])],
    ]
    info_table = Table(info, colWidths=[30 * mm, 62 * mm, 28 * mm, 54 * mm])
    info_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9), ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(ink)),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f4efff")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#d9ccf7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [info_table, Spacer(1, 12)]

    # Overall score
    overall = Table(
        [[Paragraph(f"<font size=34 color='{purple}'><b>{result['overall_score']}</b></font>"
                    f"<font size=12 color='{soft}'> / 100</font>",
                    ParagraphStyle("score", parent=body, leading=40)),
          Paragraph(f"<font size=15 color='{purple}'><b>{pdf_text(result['grade'])}</b></font><br/>"
                    f"<font size=9 color='{soft}'>Overall presentation score</font>", body)]],
        colWidths=[60 * mm, 114 * mm],
    )
    overall.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ece6fb")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#c4b5fd")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10), ("LEFTPADDING", (0, 0), (-1, -1), 14),
    ]))
    story += [overall, Paragraph("Overall Analysis Summary", h2), Paragraph(pdf_text(result["summary"]), body)]

    # Category scores
    story.append(Paragraph("Category-wise Scores", h2))
    rows = []
    for key, label, weight, _ in CATEGORIES:
        c = result["categories"][key]
        score_txt = "Not scored" if c["score"] is None else f"{c['score']}/100"
        rows.append([Paragraph(f"<b>{pdf_text(label)}</b>", body), bar_drawing(c["score"]),
                     Paragraph(score_txt, body)])
        if c["comment"]:
            rows.append([Paragraph(pdf_text(c["comment"]), small), "", ""])
    cat_table = Table(rows, colWidths=[78 * mm, 62 * mm, 34 * mm])
    style_cmds = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3),
                  ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    for i, row in enumerate(rows):
        if row[1] == "":
            style_cmds.append(("SPAN", (0, i), (2, i)))
            style_cmds.append(("LINEBELOW", (0, i), (2, i), 0.4, colors.HexColor("#ece6fb")))
    cat_table.setStyle(TableStyle(style_cmds))
    story.append(cat_table)

    def bullets(heading, entries, fmt):
        story.append(Paragraph(heading, h2))
        if not entries:
            story.append(Paragraph("Nothing to report.", body))
        for entry in entries:
            story.append(Paragraph(fmt(entry), item, bulletText="\u2022"))

    def loc_tag(loc):
        return f" <font color='#7c3aed'>[{pdf_text(loc)}]</font>" if loc else ""

    bullets("Strengths", result["strengths"],
            lambda s: f"<b>{pdf_text(s['title'])}</b>{loc_tag(s['location'])}<br/>{pdf_text(s['detail'])}")
    bullets("Weaknesses and Mistakes", result["weaknesses"],
            lambda w: f"<b>{pdf_text(w['title'])}</b>{loc_tag(w['location'])} "
                      f"<font color='#db2777'>({pdf_text(w['severity'])} severity)</font><br/>{pdf_text(w['detail'])}")
    bullets("Suggestions for Improvement", result["suggestions"],
            lambda s: f"<b>({pdf_text(s['priority'])} priority)</b>{loc_tag(s['location'])} {pdf_text(s['suggestion'])}")
    bullets("Grammar and Spelling Issues", result["grammar_issues"],
            lambda g: f"{loc_tag(g['location'])} {pdf_text(g['issue'])}"
                      + (f"<br/><i>Fix:</i> {pdf_text(g['fix'])}" if g["fix"] else ""))
    bullets("Repeated Information", result["repeated_items"],
            lambda r: f"{pdf_text(r['description'])}{loc_tag(r['location'])}")
    bullets("Missing Important Information", result["missing_items"], lambda m: pdf_text(m["item"]))

    # Measured facts and limitations
    story.append(Paragraph("Measured Facts (calculated by the program)", h2))
    facts = [
        f"Average words per {meta['unit'].lower()}: {stats['avg_words']} (highest: {stats['max_words']}).",
        f"Dense {meta['unit'].lower()}s (over {DENSE_WORDS} words): {fmt_units(meta['unit'], stats['dense_units'])}.",
        f"Almost empty {meta['unit'].lower()}s: {fmt_units(meta['unit'], stats['empty_units'])}.",
        f"{meta['unit']}s with small fonts (approximate): {fmt_units(meta['unit'], stats['small_font_units'])}.",
    ]
    for f in facts:
        story.append(Paragraph(pdf_text(f), item, bulletText="\u2022"))

    story.append(Paragraph("Limitations of this Analysis", h2))
    limits = [
        result["design_note"],
        "The analysis is produced by an AI model and may contain mistakes. Please treat it as guidance, not as a final grade.",
        "Text inside images, diagrams or scanned pages is not read (no OCR). Animations and videos are not analysed.",
        f"At most {MAX_UNITS} slides/pages are analysed. {APP_NAME} only analyses; it never edits the presentation.",
    ]
    for lim in limits:
        story.append(Paragraph(pdf_text(lim), item, bulletText="\u2022"))

    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor(soft))
        canvas.drawString(18 * mm, 10 * mm, f"{APP_NAME} - {APP_SUBTITLE}")
        canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# 8. USER INTERFACE (Streamlit + custom CSS)
# ---------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
html, body, .stApp, .stMarkdown, button, input, textarea { font-family: 'Poppins', sans-serif; }
.stApp { background: linear-gradient(135deg, #f4efff 0%, #e6f0ff 55%, #fff0f7 100%); color: #2d2150; }
[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 1180px; padding-top: 1.6rem; padding-bottom: 3rem; }
h1, h2, h3, h4 { color: #4c1d95; }
.hero { background: linear-gradient(120deg, #7c3aed 0%, #8b7cf6 45%, #60a5fa 100%); border-radius: 28px;
        padding: 2.2rem 2rem; color: #fff; text-align: center; box-shadow: 0 16px 40px rgba(124,58,237,.25);
        margin-bottom: 1.4rem; position: relative; overflow: hidden; }
.hero:after { content: ""; position: absolute; right: -40px; top: -40px; width: 180px; height: 180px;
              background: rgba(249,168,212,.45); border-radius: 50%; }
.hero h1 { color: #fff; font-size: 2.4rem; margin: 0; font-weight: 700; position: relative; }
.hero p { margin: .5rem 0 0; font-size: 1.05rem; opacity: .95; position: relative; }
.badge { display: inline-block; background: rgba(255,255,255,.22); border-radius: 999px; padding: .25rem .9rem;
         font-size: .8rem; margin-bottom: .7rem; position: relative; }
.card { background: #fff; border-radius: 22px; padding: 1.3rem 1.5rem; border: 1px solid #efe9fc;
        box-shadow: 0 10px 30px rgba(124,58,237,.10); margin-bottom: 1rem; }
.card h3 { margin-top: 0; font-size: 1.15rem; }
.check-list { list-style: none; padding: 0; margin: 0; }
.check-list li { padding: .28rem 0; font-size: .92rem; }
.check-list li:before { content: "✔"; color: #ec4899; margin-right: .55rem; font-weight: 700; }
.notice { background: #fff5fa; border: 1px solid #fbcfe8; border-radius: 16px; padding: .8rem 1rem;
          font-size: .88rem; color: #7a2556; }
.ring { width: 190px; height: 190px; border-radius: 50%; display: flex; align-items: center;
        justify-content: center; margin: 0 auto; }
.ring-inner { width: 146px; height: 146px; border-radius: 50%; background: #fff; display: flex;
              flex-direction: column; align-items: center; justify-content: center;
              box-shadow: inset 0 0 18px rgba(124,58,237,.12); }
.ring-score { font-size: 3rem; font-weight: 700; color: #6d28d9; line-height: 1; }
.ring-of { font-size: .75rem; color: #8a7bb8; }
.grade { text-align: center; margin-top: .7rem; font-weight: 600; font-size: 1.2rem; color: #4c1d95; }
.chip { display: inline-block; background: #f4efff; color: #5b3aa8; border-radius: 999px; padding: .2rem .8rem;
        font-size: .8rem; margin: .15rem .25rem .15rem 0; border: 1px solid #e3d8fb; }
.cat-head { display: flex; justify-content: space-between; align-items: center; font-weight: 600; color: #3b1f7a; }
.cat-score { color: #7c3aed; font-weight: 700; }
.bar { background: #ece6fb; border-radius: 999px; height: 11px; margin: .55rem 0; overflow: hidden; }
.bar-fill { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #8b5cf6, #60a5fa); }
.bar-fill.low { background: linear-gradient(90deg, #f472b6, #fb9ec9); }
.cat-comment { font-size: .85rem; color: #5d5188; }
.item { border-radius: 16px; padding: .8rem 1rem; margin-bottom: .6rem; border: 1px solid #efe9fc; background: #fff; }
.item-good { border-left: 6px solid #8b5cf6; }
.item-bad { border-left: 6px solid #f472b6; }
.item-tip { border-left: 6px solid #60a5fa; }
.item-title { font-weight: 600; color: #3b1f7a; font-size: .95rem; }
.item-detail { font-size: .87rem; color: #5d5188; margin-top: .25rem; }
.pill { display: inline-block; border-radius: 999px; padding: .08rem .6rem; font-size: .72rem; margin-left: .3rem; font-weight: 500; }
.pill-loc { background: #e8f0ff; color: #2c5aa8; }
.pill-high { background: #fde1ee; color: #b0225f; }
.pill-medium { background: #ece6fb; color: #5b3aa8; }
.pill-low { background: #e1efff; color: #2c5aa8; }
.stButton > button, .stDownloadButton > button { background: linear-gradient(90deg, #8b5cf6, #60a5fa); color: #fff;
        border: none; border-radius: 14px; padding: .65rem 1.4rem; font-weight: 600;
        box-shadow: 0 8px 20px rgba(124,58,237,.25); }
.stButton > button:hover, .stDownloadButton > button:hover { color: #fff; transform: translateY(-1px); }
[data-testid="stFileUploaderDropzone"] { background: #faf7ff; border: 2px dashed #a78bfa; border-radius: 20px; }
.stTextInput input { border-radius: 12px; }
.stTabs [data-baseweb="tab"] { border-radius: 12px 12px 0 0; font-weight: 500; }
@media (max-width: 768px) {
  .hero h1 { font-size: 1.7rem; } .ring { width: 150px; height: 150px; }
  .ring-inner { width: 114px; height: 114px; } .ring-score { font-size: 2.3rem; }
}
</style>
"""

e = html.escape  # shortcut: always escape AI text before putting it in HTML


def render_header():
    st.markdown(
        f"""<div class="hero"><div class="badge">AI-powered presentation analyzer</div>
        <h1>📊 {APP_NAME}</h1><p>{APP_SUBTITLE} - upload a presentation and get an honest, detailed evaluation</p></div>""",
        unsafe_allow_html=True,
    )


def ring_html(score):
    deg = int(score * 3.6)
    return (f'<div class="ring" style="background: conic-gradient(#7c3aed 0deg, #60a5fa {deg}deg, '
            f'#efe9fc {deg}deg 360deg);"><div class="ring-inner"><div class="ring-score">{score}</div>'
            f'<div class="ring-of">out of 100</div></div></div>')


def category_card(label, icon, weight, cat):
    score = cat["score"]
    if score is None:
        score_txt, width, low = "Not scored", 0, False
    else:
        score_txt, width, low = f"{score}/100", score, score < 50
    return (f'<div class="card"><div class="cat-head"><span>{icon} {e(label)}</span>'
            f'<span class="cat-score">{score_txt}</span></div>'
            f'<div class="bar"><div class="bar-fill {"low" if low else ""}" style="width:{width}%"></div></div>'
            f'<div class="cat-comment">{e(cat["comment"] or "No comment.")}</div>'
            f'<div class="cat-comment" style="margin-top:.35rem;opacity:.7">Weight in overall score: {weight}%</div></div>')


def item_card(kind, title, detail, location="", tag=""):
    loc = f'<span class="pill pill-loc">{e(location)}</span>' if location else ""
    tg = f'<span class="pill pill-{e(tag)}">{e(tag)}</span>' if tag else ""
    return (f'<div class="item item-{kind}"><div class="item-title">{e(title)}{loc}{tg}</div>'
            f'<div class="item-detail">{e(detail)}</div></div>')


def show_items(cards, empty_text):
    if cards:
        st.markdown("".join(cards), unsafe_allow_html=True)
    else:
        st.info(empty_text)


def render_results(result, meta, stats, pdf_bytes):
    chips = "".join(f'<span class="chip">{e(c)}</span>' for c in [
        f"📄 {meta['file_name']}", meta["file_type"],
        f"{meta['units_analysed']} {meta['unit'].lower()}s analysed", f"{meta['total_words']} words",
        f"Topic: {result['detected_topic']}",
    ])

    left, right = st.columns([1, 2])
    with left:
        st.markdown(f'<div class="card">{ring_html(result["overall_score"])}'
                    f'<div class="grade">{e(result["grade"])}</div></div>', unsafe_allow_html=True)
    with right:
        st.markdown(f'<div class="card"><h3>Overall Analysis</h3><div>{chips}</div>'
                    f'<p style="margin-top:.8rem">{e(result["summary"])}</p></div>', unsafe_allow_html=True)
        st.download_button("⬇️ Download Report as PDF", data=pdf_bytes,
                           file_name="SlideCheck_AI_Report.pdf", mime="application/pdf")

    tabs = st.tabs(["Category Scores", "Strengths & Weaknesses", "Suggestions", "Detailed Findings", "Limitations"])

    with tabs[0]:
        cols = st.columns(2)
        for i, (key, label, weight, icon) in enumerate(CATEGORIES):
            with cols[i % 2]:
                st.markdown(category_card(label, icon, weight, result["categories"][key]), unsafe_allow_html=True)

    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### ✅ Strengths")
            show_items([item_card("good", s["title"], s["detail"], s["location"]) for s in result["strengths"]],
                       "No clear strengths were reported.")
        with c2:
            st.markdown("### ⚠️ Weaknesses & Mistakes")
            show_items([item_card("bad", w["title"], w["detail"], w["location"], w["severity"])
                        for w in result["weaknesses"]], "No weaknesses were reported.")

    with tabs[2]:
        st.markdown("### 💡 How to Improve")
        st.caption("These are suggestions only. SlideCheck AI never edits your presentation.")
        show_items([item_card("tip", s["suggestion"], "", s["location"], s["priority"])
                    for s in result["suggestions"]], "No suggestions were returned.")

    with tabs[3]:
        d1, d2 = st.columns(2)
        with d1:
            st.markdown("### ✍️ Grammar & Spelling")
            show_items([item_card("bad", g["issue"], ("Fix: " + g["fix"]) if g["fix"] else "", g["location"])
                        for g in result["grammar_issues"]], "No grammar or spelling problems were reported.")
            st.markdown("### 🔁 Repeated Information")
            show_items([item_card("bad", r["description"], "", r["location"]) for r in result["repeated_items"]],
                       "No unnecessary repetition was reported.")
        with d2:
            st.markdown("### 🧩 Missing Information")
            show_items([item_card("tip", m["item"], "") for m in result["missing_items"]],
                       "No missing information was reported.")
            st.markdown("### 📏 Measured Facts")
            u = meta["unit"]
            st.markdown(
                f'<div class="card"><div class="cat-comment">'
                f'Average words per {u.lower()}: <b>{stats["avg_words"]}</b> (highest {stats["max_words"]})<br>'
                f'Dense {u.lower()}s (&gt;{DENSE_WORDS} words): <b>{e(fmt_units(u, stats["dense_units"]))}</b><br>'
                f'Almost empty: <b>{e(fmt_units(u, stats["empty_units"]))}</b><br>'
                f'Small fonts (approx.): <b>{e(fmt_units(u, stats["small_font_units"]))}</b></div></div>',
                unsafe_allow_html=True)

    with tabs[4]:
        st.markdown(
            f'<div class="notice"><b>What was NOT inspected</b><br>{e(result["design_note"])}</div><br>'
            '<div class="card"><ul class="check-list">'
            "<li>The score comes from an AI model and may contain mistakes - use it as guidance.</li>"
            "<li>Text inside pictures or scanned pages is not read (no OCR).</li>"
            "<li>Animations, transitions, audio and video are not analysed.</li>"
            f"<li>At most {MAX_UNITS} slides/pages are analysed per file.</li>"
            f"<li>{APP_NAME} only analyses. It does not create, edit or correct presentations.</li>"
            "</ul></div>", unsafe_allow_html=True)


def render_api_key_help():
    st.warning("⚠️ No Gemini API key was found, so analysis is disabled.")
    st.markdown(
        "**Local setup:** create `.streamlit/secrets.toml` and add `GEMINI_API_KEY = \"your-key\"` "
        "(or set the `GEMINI_API_KEY` environment variable).  \n"
        "**Streamlit Community Cloud:** open *App settings → Secrets* and paste the same line.  \n"
        "Get a free key at https://aistudio.google.com/apikey"
    )


def main():
    st.set_page_config(page_title=f"{APP_NAME} - {APP_SUBTITLE}", page_icon="📊", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    render_header()

    api_key = get_secret("GEMINI_API_KEY")
    model = get_secret("GEMINI_MODEL", DEFAULT_MODEL) or DEFAULT_MODEL

    left, right = st.columns([3, 2])
    with left:
        st.markdown('<div class="card"><h3>1. Upload your presentation</h3>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Drag and drop a .pptx or .pdf file here", type=["pptx", "pdf"],
                                    label_visibility="collapsed")
        topic = st.text_input("Presentation topic (optional)", placeholder="e.g. Introduction to Machine Learning")
        audience = st.text_input("Target audience (optional)", placeholder="e.g. 3rd-semester students")
        st.markdown("</div>", unsafe_allow_html=True)
        if not api_key:
            render_api_key_help()
        analyze = st.button("🔍 Analyze Presentation", disabled=not api_key)
    with right:
        st.markdown(
            '<div class="card"><h3>What SlideCheck AI checks</h3><ul class="check-list">'
            "<li>Content quality and topic relevance</li><li>Grammar and spelling</li>"
            "<li>Structure, introduction and conclusion</li><li>Readability and text density</li>"
            "<li>Repeated and missing information</li><li>Design facts (where they can be inspected)</li></ul></div>"
            '<div class="notice">SlideCheck AI <b>only analyses</b>. It never creates, edits or corrects your '
            "slides. Files are processed in memory and are not stored.</div>",
            unsafe_allow_html=True)

    if analyze:
        if uploaded is None:
            st.warning("Please upload a .pptx or .pdf file first.")
        else:
            data = uploaded.getvalue()
            if len(data) == 0:
                st.error("The uploaded file is empty.")
            elif len(data) > MAX_FILE_MB * 1024 * 1024:
                st.error(f"The file is larger than {MAX_FILE_MB} MB. Please upload a smaller file.")
            else:
                try:
                    with st.spinner("Reading your slides and asking Gemini for an evaluation..."):
                        result, meta, stats = analyze_presentation(uploaded.name, data, topic.strip(),
                                                                   audience.strip(), api_key, model)
                        pdf_bytes = build_pdf_report(result, meta, stats)
                    st.session_state["analysis"] = (result, meta, stats, pdf_bytes)
                except AnalysisError as exc:
                    st.session_state.pop("analysis", None)
                    st.error(str(exc))
                except Exception as exc:  # last safety net
                    st.session_state.pop("analysis", None)
                    st.error(f"Something unexpected went wrong: {str(exc)[:200]}")

    if "analysis" in st.session_state:
        st.markdown("---")
        render_results(*st.session_state["analysis"])

    st.markdown(f'<p style="text-align:center;color:#8a7bb8;font-size:.8rem;margin-top:2rem">'
                f"{APP_NAME} - {APP_SUBTITLE} | University Data Science Project</p>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()

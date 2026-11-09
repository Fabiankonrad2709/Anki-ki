import hashlib
import io
import json
import math
import os
import re

import anthropic
import genanki
import pandas as pd
import streamlit as st
from docx import Document
from pptx import Presentation
from pypdf import PdfReader

st.set_page_config(page_title="Anki-KI", page_icon="🦷", layout="centered")

TYPEN = ["Frage-Antwort", "Cloze", "Definition"]
LEVELS = {
    "Grundlagen": "Grundlagenwissen, klare Basisfakten",
    "Fortgeschritten": "vertiefte Zusammenhänge, Mechanismen, Differenzialdiagnosen",
    "Prüfungsniveau": "universitäres Prüfungsniveau (Staatsexamen/Klausur), präzise und anwendungsbezogen",
}
CHUNK = 12000
MAX_CHARS = 90000

SYSTEM = """Du bist Experte für Zahnmedizin und erstellst Anki-Karteikarten für Studierende.
Regeln:
- Nutze AUSSCHLIESSLICH Informationen aus dem gegebenen Lernmaterial. Erfinde nichts, ergänze kein Fremdwissen.
- Pro Karte genau eine konkrete Wissensfrage. Antworten kompakt (meist 1-2 Sätze oder Stichpunkte).
- Keine doppelten oder inhaltlich redundanten Karten.
- Berücksichtige klinische Relevanz und wichtige Zusammenhänge, soweit im Material belegt.
- Cloze-Karten: Lückentext im Format {{c1::Begriff}} im Feld "vorne", "hinten" ist optional (Zusatzinfo) oder leer.
- Definitionskarten: vorne = Begriff, hinten = knappe Definition.
- Frage-Antwort: vorne = Frage, hinten = Antwort.
- Tags: 2-4 sinnvolle Schlagwörter, ohne Leerzeichen (Unterstrich verwenden). Kategorie: Fachgebiet/Thema (kurz).
Antworte NUR mit einem JSON-Array, ohne Text drumherum. Jedes Element:
{"typ": "Frage-Antwort|Cloze|Definition", "vorne": "...", "hinten": "...", "kategorie": "...", "tags": ["..."]}"""


# ---------- Textextraktion ----------
def extract_text(f) -> str:
    name = f.name.lower()
    data = f.getvalue()
    if name.endswith(".pdf"):
        r = PdfReader(io.BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in r.pages)
    if name.endswith(".docx"):
        d = Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    if name.endswith(".pptx"):
        prs = Presentation(io.BytesIO(data))
        parts = []
        for i, s in enumerate(prs.slides, 1):
            parts.append(f"--- Folie {i} ---")
            for sh in s.shapes:
                if sh.has_text_frame:
                    parts.append(sh.text_frame.text)
                if getattr(sh, "has_table", False) and sh.has_table:
                    for row in sh.table.rows:
                        parts.append(" | ".join(c.text for c in row.cells))
            if s.has_notes_slide:
                parts.append(s.notes_slide.notes_text_frame.text)
        return "\n".join(parts)
    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")
    return ""


def chunk_text(text: str):
    text = re.sub(r"\n{3,}", "\n\n", text).strip()[:MAX_CHARS]
    chunks, cur = [], ""
    for para in text.split("\n"):
        if len(cur) + len(para) > CHUNK and cur:
            chunks.append(cur)
            cur = ""
        cur += para + "\n"
    if cur.strip():
        chunks.append(cur)
    return chunks


# ---------- Generierung ----------
def parse_json(s: str):
    a, b = s.find("["), s.rfind("]")
    if a == -1 or b == -1:
        return []
    try:
        return json.loads(s[a : b + 1])
    except json.JSONDecodeError:
        return []


def generate(api_key, model, text, n, level, typen, progress):
    client = anthropic.Anthropic(api_key=api_key)
    chunks = chunk_text(text)
    per = max(1, math.ceil(n * 1.2 / len(chunks)))  # etwas Puffer für Duplikate
    cards = []
    for i, ch in enumerate(chunks, 1):
        progress.progress(i / len(chunks), text=f"Abschnitt {i}/{len(chunks)} wird analysiert …")
        prompt = (
            f"Erstelle bis zu {per} Karteikarten.\n"
            f"Schwierigkeitsgrad: {LEVELS[level]}.\n"
            f"Erlaubte Kartentypen: {', '.join(typen)}.\n\n"
            f"LERNMATERIAL:\n{ch}"
        )
        msg = client.messages.create(
            model=model, max_tokens=8000, system=SYSTEM,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = "".join(b.text for b in msg.content if b.type == "text")
        cards += parse_json(raw)
    # bereinigen, Duplikate entfernen
    seen, out = set(), []
    for c in cards:
        if not isinstance(c, dict) or c.get("typ") not in typen:
            continue
        vorne = str(c.get("vorne", "")).strip()
        key = re.sub(r"\W+", "", vorne.lower())
        if not vorne or key in seen:
            continue
        seen.add(key)
        out.append({
            "Typ": c["typ"],
            "Vorderseite": vorne,
            "Rückseite": str(c.get("hinten", "")).strip(),
            "Kategorie": str(c.get("kategorie", "")).strip(),
            "Tags": ", ".join(c.get("tags", [])),
        })
    return out[:n]


# ---------- Export ----------
def _mid(s):  # stabile ID aus Text
    return int(hashlib.md5(s.encode()).hexdigest()[:8], 16)


CSS = ".card{font-family:Arial;font-size:20px;text-align:center;color:#222;background:#fff}"
BASIC = genanki.Model(
    1607392319, "Anki-KI Basic",
    fields=[{"name": "Front"}, {"name": "Back"}],
    templates=[{"name": "Karte 1", "qfmt": "{{Front}}",
                "afmt": "{{FrontSide}}<hr id=answer>{{Back}}"}], css=CSS)
CLOZE = genanki.Model(
    1607392320, "Anki-KI Cloze",
    fields=[{"name": "Text"}, {"name": "Extra"}],
    templates=[{"name": "Cloze", "qfmt": "{{cloze:Text}}",
                "afmt": "{{cloze:Text}}<br>{{Extra}}"}],
    model_type=genanki.Model.CLOZE, css=CSS)


def clean_tags(row):
    tags = [t.strip().replace(" ", "_") for t in str(row.get("Tags", "")).split(",") if t.strip()]
    kat = str(row.get("Kategorie", "")).strip().replace(" ", "_")
    return ([kat] if kat else []) + tags


def build_apkg(df: pd.DataFrame, deck_name: str) -> bytes:
    deck = genanki.Deck(_mid(deck_name), deck_name)
    for _, r in df.iterrows():
        v, h = str(r["Vorderseite"]), str(r["Rückseite"])
        if not v.strip() or v == "nan":
            continue
        h = "" if h == "nan" else h
        model = CLOZE if r["Typ"] == "Cloze" else BASIC
        deck.add_note(genanki.Note(model=model, fields=[v, h], tags=clean_tags(r)))
    buf = io.BytesIO()
    genanki.Package(deck).write_to_file(buf)
    return buf.getvalue()


# ---------- Oberfläche ----------
st.title("🦷 Anki-KI")
st.caption("Lernmaterial hochladen → KI erstellt Karteikarten → in Anki importieren")

with st.sidebar:
    st.header("Einstellungen")
    key = st.text_input("Claude API-Key", type="password",
                        value=os.environ.get("ANTHROPIC_API_KEY", ""))
    model = st.text_input("Modell", value="claude-sonnet-5-5")

files = st.file_uploader("Dateien (PDF, Word, PowerPoint, TXT)",
                         type=["pdf", "docx", "pptx", "txt"], accept_multiple_files=True)

c1, c2 = st.columns(2)
n = c1.number_input("Anzahl Karten", 5, 200, 30, step=5)
level = c2.selectbox("Schwierigkeitsgrad", list(LEVELS))
typen = st.multiselect("Kartentypen", TYPEN, default=TYPEN)
deck_name = st.text_input("Deckname", "Zahnmedizin::Neues Thema")

if st.button("Karten generieren", type="primary", use_container_width=True):
    if not key:
        st.error("Bitte API-Key in der Seitenleiste eingeben.")
    elif not files or not typen:
        st.error("Bitte Dateien hochladen und mindestens einen Kartentyp wählen.")
    else:
        text = "\n\n".join(extract_text(f) for f in files)
        if len(text.strip()) < 50:
            st.error("Aus den Dateien konnte kaum Text gelesen werden (evtl. gescannte PDF).")
        else:
            bar = st.progress(0.0)
            try:
                cards = generate(key, model, text, int(n), level, typen, bar)
                st.session_state["df"] = pd.DataFrame(cards)
                bar.empty()
            except Exception as e:
                bar.empty()
                st.error(f"Fehler: {e}")

if "df" in st.session_state and len(st.session_state["df"]):
    st.subheader("Vorschau & Bearbeitung")
    st.caption("Zellen anklicken zum Bearbeiten. Zeile links markieren und Entf/Papierkorb zum Löschen.")
    edited = st.data_editor(
        st.session_state["df"], num_rows="dynamic", use_container_width=True,
        column_config={"Typ": st.column_config.SelectboxColumn(options=TYPEN, required=True)},
    )
    st.write(f"**{len(edited)} Karten**")
    st.download_button("📦 Als .apkg exportieren", build_apkg(edited, deck_name),
                       file_name="karten.apkg", type="primary", use_container_width=True)

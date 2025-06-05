# -*- coding: utf-8 -*- 
"""
Created on Tue Feb 18 08:02:40 2025

@author: ays0
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Feb 17 12:02:33 2025

@author: ays0

Aanpassingen/uitbreidingen n.a.v. de wensen:
 - Direct incrementaal leren en (optioneel) herladen van het model in de huidige sessie
 - Instelbare AI-confidence en OCR-mintext via config.json (bv. "AI_CONFIDENCE" en "OCR_MINTEXT")
 - Elke nieuwe leverancier moet gekoppeld worden aan een leverancier nummer. 
   De gebruiker krijgt een keuzemenu waarin het nummer (bv. L00001) kan worden opgegeven. 
   Het is ook mogelijk meerdere leveranciers hetzelfde nummer te geven.
   Het leveranciersbeheer is nu in tabelvorm, waarin de mapping direct kan worden aangepast en meteen opgeslagen.
 - Alle correcties (zowel goede als foute) die tijdens de trainingsmodus worden gemaakt, worden permanent opgeslagen in een SQLite‑database (inclusief de volledige PDF‑tekst). Zo kan het model bij hertraining gebruikmaken van deze historische feedback.
 - De hertraining wordt nu uitsluitend gebaseerd op de database (en niet op een Excel‑bestand).
 - Spelfout “hertrein(en)” is aangepast naar “hertrain(en)” .
 - Na boeken in de verwerkingsmodus (boekingsmodus) wordt per veld direct een record met de volledige PDF‑tekst en de uiteindelijke waarde opgeslagen in de SQLite‑database (zoals bij de trainingsmodus correcties).
 - De leverancier mapping kan nu eenvoudig aangepast worden door een custom nummer in te voeren (moet voldoen aan het patroon: L gevolgd door 5 cijfers).
 - **Aparte modellen per type veld worden gebruikt.**  
   Dat wil zeggen: voor elk veld (bijv. factuurnummer, factuurdatum, bedrag, leverancier, projid, inkooporder‑nummer en geadresseerde) wordt een specifiek spaCy‑model getraind en opgeslagen in een map zoals:
       trained_invoice_model_FACTUUR_NUMMER  
       trained_invoice_model_FACTUUR_DATUM  
   enzovoort.
 - De overige functionaliteit blijft ongewijzigd.
 - **Nieuw:** Incrementaal leren gebeurt nu pas per batch van 5 facturen.
 - **Nieuw:** Voor het veld 'geadresseerde' wordt een lookup gedaan in de lijst bekende geadresseerden. Als er een match is, wordt deze altijd met 100% zekerheid ingevuld.
 - **Nieuw in dit script:**  
    1. Geadresseerde & Leverancier: Lookup in de bekende lijsten met exacte matching. Bij meerdere matches wordt de match met de meeste karakters gekozen.
    2. Datum: Alle datumformats worden herkend (inclusief punt- en slash‑gescheiden formaten). Enkel volledige, kwalificerende datums (met een 4-cijferig jaartal) worden teruggegeven. Zo voorkomt dit dat een incomplete datum (zoals "30/11/202") wordt gebruikt.
    3. Projid: Het document wordt doorzocht naar een string die begint met "P" of "G" gevolgd door 6 cijfers. Als deze wordt gevonden, wordt alles eromheen verwijderd. Enkel als er een projid-koppeling bestaat (lookup) wordt deze waarde gebruikt.
    4. Inkooporder Nummer: Herkenning vereist dat het nummer begint met "IK" of "IKGV".
    5. Bedrag: Er wordt gezocht naar het totaalbedrag (incl. BTW) met meerdere regex-synoniemen.
    6. Factuurnummer: Regex functies zijn uitgebreid met onder andere de synoniem "docnr:".
"""

import sys
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
import json
import pdfplumber
import spacy
from spacy.tokens import DocBin
from spacy.util import minibatch, compounding, filter_spans
from spacy.training import Example
import pandas as pd
from PIL import Image, ImageTk
import pytesseract
from pdf2image import convert_from_path
import re
import time
import threading
import random
import sqlite3  # Voor opslag van trainingscorrecties

# -------------------------------------
# Hulpfunctie: resource pad bepalen
# -------------------------------------
def get_resource_path(relative_path):
    """Zoekt het juiste pad, ook als we met PyInstaller exe zijn."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

##############################################################################
# Configuratie
##############################################################################
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# -------------------------------------
# Stel paden voor Tesseract en Poppler in
# -------------------------------------
TESSERACT_CMD = get_resource_path(os.path.join("Tesseract-OCR", "tesseract.exe"))
POPLER_BIN_PATH = get_resource_path(os.path.join("poppler-24.08.0", "Library", "bin"))

# -------------------------------------
# Configuratie laden
# -------------------------------------
if not os.path.exists(CONFIG_FILE):
    print(f"Let op: config.json niet gevonden op {CONFIG_FILE}. Maak hem aan of controleer het pad.")
    config_data = {
        "TESSERACT_CMD": TESSERACT_CMD,
        "POPLER_BIN_PATH": POPLER_BIN_PATH,
        "TRAINED_MODEL_DIR": "trained_invoice_model",
        "PROCESSED_PDF_DIR": "verwerkte_pdfs",
        "CORRECTED_DATA_FILE": "verwerkte_facturen.csv",
        "TRAINING_DATA_EXCEL": "training_data.xlsx",
        "BASE_SPACY_MODEL": "nl_core_news_sm",
        "INVOICE_FOLDER": None,
        "UI_STATE": {
            "main_window_geometry": "1200x800",
            "pdf_preview_geometry": "800x600",
            "pdf_zoom_factor": 0.8
        },
        "ENABLE_OCR_FALLBACK": True,
        "AI_CONFIDENCE": 0.95,
        "OCR_MINTEXT": 50,
        "INCREMENTAL_BATCH_SIZE": 1
    }
else:
    try:
        with open(CONFIG_FILE, "r") as f:
            config_data = json.load(f)
    except Exception as e:
        print(f"Fout bij laden van de configuratie: {e}")
        config_data = {}

TESSERACT_CMD       = config_data.get("TESSERACT_CMD", TESSERACT_CMD)
POPLER_BIN_PATH     = config_data.get("POPLER_BIN_PATH", POPLER_BIN_PATH)
TRAINED_MODEL_DIR   = config_data.get("TRAINED_MODEL_DIR", "trained_invoice_model")
PROCESSED_PDF_DIR   = config_data.get("PROCESSED_PDF_DIR", "verwerkte_pdfs")
CORRECTED_DATA_FILE = config_data.get("CORRECTED_DATA_FILE", "verwerkte_facturen.csv")
TRAINING_DATA_EXCEL = config_data.get("TRAINING_DATA_EXCEL", "training_data.xlsx")
BASE_SPACY_MODEL    = config_data.get("BASE_SPACY_MODEL", "nl_core_news_sm")
INVOICE_FOLDER      = config_data.get("INVOICE_FOLDER", None)
UI_STATE            = config_data.get("UI_STATE", {})
ENABLE_OCR_FALLBACK = config_data.get("ENABLE_OCR_FALLBACK", True)
AI_CONFIDENCE       = config_data.get("AI_CONFIDENCE", 0.95)
OCR_MINTEXT         = config_data.get("OCR_MINTEXT", 50)
INCREMENTAL_BATCH_SIZE = config_data.get("INCREMENTAL_BATCH_SIZE", 1)

pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

MAIN_WINDOW_GEOMETRY = UI_STATE.get("main_window_geometry", "1200x800")
PDF_PREVIEW_GEOMETRY = UI_STATE.get("pdf_preview_geometry", "800x600")
PDF_PREVIEW_ZOOM     = UI_STATE.get("pdf_zoom_factor", 0.8)

# Nieuwe globale variabelen
COUNTERS_FILE      = "counters.json"
EXTRA_DATA_EXCEL   = "alle_facturen.xlsx"
ANNOTATION_FILE    = "verwerkte_facturen_annotations.csv"  # Hier slaan we offsets op
DB_TRAINING        = "training_corrections.db"  # SQLite-database voor trainingscorrecties

# -------------------------------------
# Overige constante waarden
# -------------------------------------
LEARNED_OFFSETS_FILE   = "learned_offsets.json"
PROJID_MAPPING_FILE    = "projid_mappings.json"
GEADRESSEERDE_MAP_FILE = "geadresseerde_subfolder_map.json"
KNOWN_SUPPLIERS_FILE   = "known_suppliers.json"
KNOWN_GEADRESSEERDEN   = ["gv", "efs", "vg1", "vg2", "vs"]

# =============================================================================
# Nieuwe globale variabelen voor aparte modellen
# =============================================================================
FIELD_MAPPING = {
    "factuur_nummer": "FACTUUR_NUMMER",
    "factuur_datum": "FACTUUR_DATUM",
    "bedrag": "BEDRAG",
    "leverancier": "LEVERANCIER",
    "projid": "PROJID",
    "inkooporder_nummer": "INKOOPORDER_NUMMER",
    "geadresseerde": "GEADRESSEERDE",
}

def get_model_dir_for_field(field):
    """Bepaalt de modeldirectory voor een specifiek veld."""
    return f"{TRAINED_MODEL_DIR}_{field.upper()}"

def load_all_field_models():
    """Laadt voor elk veld het bijbehorende model (indien aanwezig)."""
    models = {}
    for field in FIELD_MAPPING.keys():
        model_dir = get_model_dir_for_field(field)
        if os.path.exists(model_dir):
            try:
                models[field] = spacy.load(model_dir)
                print(f"Model voor {field} geladen uit {model_dir}.")
            except Exception as e:
                print(f"Fout bij laden model voor {field}: {e}")
                models[field] = None
        else:
            models[field] = None
    return models

# -------------------------------------
# Hulpfuncties voor counters, mappings, offsets, leveranciers, pdf-extractie, candidates, heuristieken
# -------------------------------------
def load_counters():
    if os.path.exists(COUNTERS_FILE):
        try:
            with open(COUNTERS_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Fout bij laden van counters: {e}")
    return {}

def save_counters(counters):
    try:
        with open(COUNTERS_FILE, "w") as f:
            json.dump(counters, f, indent=4)
    except Exception as e:
        print(f"Fout bij opslaan van counters: {e}")

def load_projid_mappings():
    if os.path.exists(PROJID_MAPPING_FILE):
        try:
            with open(PROJID_MAPPING_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Fout bij laden van PROJID-koppelingen: {e}")
    return {}

def save_projid_mappings(mappings):
    try:
        with open(PROJID_MAPPING_FILE, "w") as f:
            json.dump(mappings, f, indent=4)
    except Exception as e:
        print(f"Fout bij opslaan van PROJID-koppelingen: {e}")

def lookup_proj_mapping(text):
    mappings = load_projid_mappings()
    text_lower = text.lower()
    for description, projid in mappings.items():
        if description.lower() in text_lower:
            return projid
    return None

def load_learned_offsets():
    if os.path.exists(LEARNED_OFFSETS_FILE):
        try:
            with open(LEARNED_OFFSETS_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Fout bij laden van geleerde offsets: {e}")
    return {}

def save_learned_offsets(offsets):
    try:
        with open(LEARNED_OFFSETS_FILE, "w") as f:
            json.dump(offsets, f, indent=4)
    except Exception as e:
        print(f"Fout bij opslaan van geleerde offsets: {e}")

LEARNED_OFFSETS = load_learned_offsets()

def find_corrected_value_location(text: str, corrected_value: str, threshold: float = 0.8) -> int:
    index = text.find(corrected_value)
    if index != -1:
        return index
    import difflib
    best_index = -1
    best_ratio = 0.0
    window_len = len(corrected_value)
    step = max(1, window_len // 10) if window_len > 10 else 1
    for i in range(0, len(text) - window_len + 1, step):
        candidate = text[i:i+window_len]
        ratio = difflib.SequenceMatcher(None, candidate, corrected_value).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_index = i
        if best_ratio >= threshold:
            break
    return best_index if best_ratio >= threshold else -1

def update_learned_offset(supplier, field, text, corrected_value):
    if not supplier or not corrected_value:
        return
    offset = text.find(corrected_value)
    if offset == -1:
        offset = find_corrected_value_location(text, corrected_value)
    if offset != -1:
        if supplier not in LEARNED_OFFSETS:
            LEARNED_OFFSETS[supplier] = {}
        LEARNED_OFFSETS[supplier][field] = {
            "offset": offset,
            "length": len(corrected_value),
            "value": corrected_value
        }
        save_learned_offsets(LEARNED_OFFSETS)
        print(f"Geleerde offset voor leverancier '{supplier}', veld '{field}' op positie {offset} bijgewerkt.")

def get_value_from_learned_offset(text, supplier, field):
    if supplier in LEARNED_OFFSETS and field in LEARNED_OFFSETS[supplier]:
        info = LEARNED_OFFSETS[supplier][field]
        offset = info.get("offset")
        length = info.get("length")
        if offset is not None and length:
            candidate = text[offset:offset+length]
            import difflib
            ratio = difflib.SequenceMatcher(None, candidate, info.get("value", "")).ratio()
            if ratio > 0.8:
                return candidate
    return None

def load_geadresseerde_map():
    if os.path.exists(GEADRESSEERDE_MAP_FILE):
        try:
            with open(GEADRESSEERDE_MAP_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Fout bij laden van geadresseerde-koppelingen: {e}")
    return {}

def save_geadresseerde_map(mapping: dict):
    try:
        with open(GEADRESSEERDE_MAP_FILE, "w") as f:
            json.dump(mapping, f, indent=4)
    except Exception as e:
        print(f"Fout bij opslaan van geadresseerde-koppelingen: {e}")

def load_known_suppliers():
    if os.path.exists(KNOWN_SUPPLIERS_FILE):
        try:
            with open(KNOWN_SUPPLIERS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
                elif isinstance(data, list):
                    return {name: f"L{idx+1:05d}" for idx, name in enumerate(data)}
        except Exception as e:
            print(f"Fout bij laden van bekende leveranciers: {e}")
    return {}

def save_known_suppliers(suppliers_dict):
    try:
        with open(KNOWN_SUPPLIERS_FILE, "w", encoding="utf-8") as f:
            json.dump(suppliers_dict, f, indent=4)
    except Exception as e:
        print(f"Fout bij opslaan van bekende leveranciers: {e}")

# Aangepaste lookup: altijd lookup in bekende lijsten met exacte matching.
def lookup_supplier_from_known_values(text):
    suppliers = load_known_suppliers()
    best_match = None
    best_length = 0
    for name in suppliers.keys():
        if re.search(r'\b' + re.escape(name) + r'\b', text, re.IGNORECASE):
            if len(name) > best_length:
                best_length = len(name)
                best_match = name
    return best_match

def lookup_geadresseerde_from_known_values(text):
    subfolder_map = load_geadresseerde_map()
    known_geo_keys = set(subfolder_map.keys())
    for g in KNOWN_GEADRESSEERDEN:
        known_geo_keys.add(g)
    best_match = None
    best_length = 0
    for geo in known_geo_keys:
        if re.search(r'\b' + re.escape(geo) + r'\b', text, re.IGNORECASE):
            if len(geo) > best_length:
                best_length = len(geo)
                best_match = geo
    return best_match

def get_pdf_path(original_pdf_path: str) -> str:
    if os.path.exists(original_pdf_path):
        return original_pdf_path
    else:
        candidate = os.path.join(PROCESSED_PDF_DIR, os.path.basename(original_pdf_path))
        if os.path.exists(candidate):
            return candidate
        else:
            return original_pdf_path

def show_pdf_scrollable_preview(pdf_path: str, container, dpi=100):
    for widget in container.winfo_children():
        widget.destroy()
    canvas = tk.Canvas(container, bg="#F0F0F0")
    v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=v_scroll.set)
    v_scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    def _on_mousewheel(event):
        canvas.yview_scroll(-1 * int(event.delta / 120), "units")
    canvas.bind("<MouseWheel>", _on_mousewheel)
    preview_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=preview_frame, anchor="nw")
    image_refs = []
    try:
        pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=POPLER_BIN_PATH)
        for i, page in enumerate(pages):
            img_tk = ImageTk.PhotoImage(page)
            label = ttk.Label(preview_frame, image=img_tk)
            label.image = img_tk
            label.pack(pady=5)
            image_refs.append(img_tk)
    except Exception as e:
        error_label = ttk.Label(preview_frame, text=f"Fout bij laden van PDF: {e}")
        error_label.pack()
    preview_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    container.image_refs = image_refs

def extract_text(pdf_path: str) -> str:
    text_pages = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                plumber_text = page.extract_text() or ""
                text_pages.append(plumber_text)
    except Exception as e:
        print(f"[extract_text] Fout bij pdfplumber: {e}")
    combined_plumber = "\n".join(text_pages)
    if ENABLE_OCR_FALLBACK:
        if len(combined_plumber.strip()) < OCR_MINTEXT:
            print("[extract_text] Te weinig tekst gevonden => OCR.")
            try:
                images = convert_from_path(pdf_path, poppler_path=POPLER_BIN_PATH)
                ocr_pages = []
                for im in images:
                    ocr_txt = pytesseract.image_to_string(im)
                    ocr_pages.append(ocr_txt)
                combined_ocr = "\n".join(ocr_pages)
                return combined_plumber + "\n" + combined_ocr
            except Exception as e:
                print(f"[extract_text] Fout bij OCR: {e}")
                return combined_plumber
        else:
            return combined_plumber
    else:
        return combined_plumber

def get_candidate_factuurnummers(text: str):
    pattern = re.compile(
        r'(?:factuurnummer|faktuurnummer|factuurnr|faktuurnr|inv[\.\s_-]?nr|nr\.?|invoice\s*number|nota|factuur|docnr)\s*[:#-]?\s*([\w/-]+)',
        flags=re.IGNORECASE
    )
    cands = pattern.findall(text)
    cands += re.findall(r'(?:INV|FCT)\d{2,}', text, flags=re.IGNORECASE)
    return list(set(cands))

def get_candidate_factuurdatums(text: str):
    pattern_with_label = re.compile(
        r'(?:factuurdatum|datum|date)\s*[:\-]?\s*([\d/.\-]+)', flags=re.IGNORECASE
    )
    cands = pattern_with_label.findall(text)
    cands += re.findall(r'\b\d{4}-\d{2}-\d{2}\b', text)
    cands += re.findall(r'\b\d{2}-\d{2}-\d{4}\b', text)
    cands += re.findall(r'\b\d{2}/\d{2}/\d{4}\b', text)
    cands += re.findall(r'\b\d{2}\.\d{2}\.\d{4}\b', text)
    cands += re.findall(r'\b\d{4}/\d{2}/\d{2}\b', text)
    return list(set(cands))

# --- NIEUWE FUNCTIE: Controleer of de datum effectief is (met 4-cijferig jaartal) ---
def is_effective_date(date_str):
    """
    Controleert of de meegegeven datum een geldig formaat heeft met een 4-cijferig jaartal.
    Ondersteunde formaten: DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY en YYYY-MM-DD.
    """
    patterns = [
        r'^\d{2}[/-]\d{2}[/-]\d{4}$',
        r'^\d{4}[/-]\d{2}[/-]\d{2}$',
        r'^\d{2}\.\d{2}\.\d{4}$'
    ]
    for pat in patterns:
        if re.match(pat, date_str):
            return True
    return False

def get_candidate_bedragen(text: str):
    pattern_with_label = re.compile(
        r'(?:totaal\s*incl\.?\s*btw|totaal incl btw|totaal btw|totaalbedrag|bedrag)\s*[:=]?\s*\€?\s*([\d\.,]+)',
        flags=re.IGNORECASE
    )
    cands = pattern_with_label.findall(text)
    cands += re.findall(r'\b\d{1,3}(?:\.\d{3})*,\d{2}\b', text)
    cands += re.findall(r'\b\d+\.\d{2}\b', text)
    return list(set(cands))

def get_candidate_leveranciers(text: str):
    cands = re.findall(r'\b[A-Z][A-Za-z]+(?: (?:BV|NV|GmbH|BVBA|LLC))\b', text)
    return list(set(cands))

def get_candidate_projid(text: str):
    pattern_uwref = re.compile(
        r'(?:uw\s*ref|projid)\s*[:#-]?\s*([GP]\d+)', flags=re.IGNORECASE
    )
    cands = pattern_uwref.findall(text)
    cands += re.findall(r'\b(?:PROJID|Project\s*ID)[:\s]+(\S+)', text, flags=re.IGNORECASE)
    return list(set(cands))

def get_candidate_inkoopordernummers(text: str):
    pattern_ik = re.compile(r'\b((?:IKGV|IK)\d+(?:-\d+)*)', re.IGNORECASE)
    cands = pattern_ik.findall(text)
    cands += re.findall(r'\b(?:INKOOPORDER(?:\s*NUMMER)?|Purchase\s*Order)[:\s]+(\S+)', text, flags=re.IGNORECASE)
    return list(set(cands))

def get_candidate_geadresseerde(text: str):
    cands = re.findall(r'(?:T\.a\.v\.|Ter attentie van)\s*(.*)', text, flags=re.IGNORECASE)
    cands += re.findall(r'Geadresseerde[:\s]+(.*)', text, flags=re.IGNORECASE)
    cands += re.findall(r'(?:Aan:|To:)\s*(.*)', text, flags=re.IGNORECASE)
    return list(set([c.strip() for c in cands if c.strip()]))

def heuristic_find_factuurnummer(text: str):
    pattern = re.compile(
        r'(?:factuurnummer|faktuurnummer|factuurnr|faktuurnr|inv[\.\s_-]?nr|nr\.?|invoice\s*number|nota|factuur|docnr)\s*[:#-]?\s*([\w/-]+)',
        flags=re.IGNORECASE
    )
    match = pattern.search(text)
    if match:
        return match.group(1)
    match2 = re.search(r'(INV|FCT)\d{2,}', text, re.IGNORECASE)
    if match2:
        return match2.group(0)
    return None

def heuristic_find_date(text: str):
    pattern_fact = re.compile(r'Datum\s+facturatie\s*[:\-]?\s*([\d/.\-]+)', re.IGNORECASE)
    match = pattern_fact.search(text)
    if match:
        return match.group(1)
    pattern_label = re.compile(r'(?:factuurdatum|datum)\s*[:\-]?\s*([\d/.\-]+)', re.IGNORECASE)
    match = pattern_label.search(text)
    if match:
        return match.group(1)
    match2 = re.search(r'\b(\d{4}-\d{2}-\d{2})\b', text)
    if match2:
        return match2.group(1)
    match3 = re.search(r'\b(\d{2}-\d{2}-\d{4})\b', text)
    if match3:
        return match3.group(1)
    match4 = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
    if match4:
        return match4.group(1)
    match5 = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
    if match5:
        return match5.group(1)
    match6 = re.search(r'\b(\d{4}/\d{2}/\d{2})\b', text)
    if match6:
        return match6.group(1)
    return None

def heuristic_find_bedrag(text: str):
    label_pat = re.findall(r'(?:totaal\s*incl\.?\s*btw|totaal incl btw|totaal btw|totaalbedrag|bedrag)\s*[:=]?\s*\€?\s*([\d\.,]+)', text, flags=re.IGNORECASE)
    matches = []
    if label_pat:
        matches.extend(label_pat)
    matches += re.findall(r'\b\d{1,3}(?:\.\d{3})*,\d{2}\b', text)
    matches += re.findall(r'\b\d+\.\d{2}\b', text)
    if matches:
        def parse_amount(s):
            s_clean = s.replace('.', '').replace(',', '.')
            try:
                return float(s_clean)
            except:
                return 0.0
        highest = max(matches, key=lambda x: parse_amount(x))
        return highest
    return None

def heuristic_find_projid(text: str):
    pattern_strict = re.compile(r'\b([GP]\d{6})(?!\d)', re.IGNORECASE)
    match_strict = pattern_strict.search(text)
    if match_strict:
        return match_strict.group(1)
    pattern = re.compile(r'(?:uw\s*ref|projid)\s*[:#-]?\s*([GP]\d+)', re.IGNORECASE)
    match = pattern.search(text)
    if match:
        return match.group(1)
    match2 = re.search(r'\b(?:PROJID|Project\s*ID)[:\s]+(\S+)', text, flags=re.IGNORECASE)
    if match2:
        return match2.group(1)
    return None

def heuristic_find_inkoopordernummer(text: str):
    pattern_ik = re.compile(r'\b((?:IKGV|IK)\d+(?:-\d+)*)', re.IGNORECASE)
    match = pattern_ik.search(text)
    if match:
        return match.group(1)
    match2 = re.search(r'\b(?:INKOOPORDER(?:\s*NUMMER)?|Purchase\s*Order)[:\s]+(\S+)', text, flags=re.IGNORECASE)
    if match2:
        return match2.group(1)
    return None

def update_training_data_excel(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path)
    if "text_value" not in df.columns:
        df["text_value"] = ""
    for idx, row in df.iterrows():
        if pd.isna(row.get("text_value", "")) or str(row.get("text_value", "")).strip() == "":
            pdf_path = get_pdf_path(row["pdf_path"])
            try:
                text = extract_text(pdf_path)
                df.at[idx, "text_value"] = text
            except Exception as e:
                print(f"Fout bij extractie van tekst uit {pdf_path}: {e}")
                continue
    df.to_excel(excel_path, index=False)
    return df

def create_spacy_training_data(df: pd.DataFrame, nlp) -> DocBin:
    db = DocBin()
    grouped = df.groupby("pdf_path")
    for pdf_path, group in grouped:
        actual_pdf_path = get_pdf_path(pdf_path)
        try:
            text = extract_text(actual_pdf_path)
        except Exception as e:
            print(f"Overgeslagen {actual_pdf_path} vanwege extractiefout: {e}")
            continue
        doc = nlp.make_doc(text)
        ents = []
        for idx, row in group.iterrows():
            label = row["label"]
            start = -1
            end = -1
            if ("start_offset" in row and "end_offset" in row and not pd.isna(row["start_offset"]) and not pd.isna(row["end_offset"])):
                start = int(row["start_offset"])
                end = int(row["end_offset"])
            else:
                val = str(row.get("text_value", "")).strip()
                if val:
                    tmp_start = text.find(val)
                    if tmp_start != -1:
                        start = tmp_start
                        end = start + len(val)
            if start >= 0 and end >= 0:
                span = doc.char_span(start, end, label=label)
                if span is not None:
                    ents.append(span)
        ents = filter_spans(ents)
        doc.ents = ents
        db.add(doc)
    return db

# =============================================================================
# Nieuwe functie: Training vanuit SQLite-database per veld
# =============================================================================
def train_spacy_model_field_from_db(field: str, output_dir: str, base_model: str, n_iter: int = 10):
    df = load_training_corrections_df()
    # Filter op correcties voor dit specifieke veld
    df = df[(df['field'] == field) & (df['corrected_value'].astype(str).str.strip() != "")]
    if df.empty:
        print(f"Geen trainingscorrecties gevonden voor veld {field}.")
        return
    try:
        nlp = spacy.load(base_model)
    except OSError:
        raise ValueError(f"Basismodel '{base_model}' niet gevonden of niet geïnstalleerd.")
    label = FIELD_MAPPING[field]
    if "ner" not in nlp.pipe_names:
        ner = nlp.add_pipe("ner", last=True)
    else:
        ner = nlp.get_pipe("ner")
    ner.add_label(label)
    db = DocBin()
    grouped = df.groupby("pdf_path")
    for pdf_path, group in grouped:
        pdf_text = group.iloc[0]['pdf_text']
        doc = nlp.make_doc(pdf_text)
        ents = []
        for _, row in group.iterrows():
            corrected_value = row['corrected_value']
            if not corrected_value:
                continue
            start = pdf_text.find(corrected_value)
            if start == -1:
                start = find_corrected_value_location(pdf_text, corrected_value)
            if start != -1:
                end = start + len(corrected_value)
                span = doc.char_span(start, end, label=label)
                if span is not None:
                    ents.append(span)
        ents = filter_spans(ents)
        doc.ents = ents
        db.add(doc)
    train_docs = list(db.get_docs(nlp.vocab))
    other_pipes = [p for p in nlp.pipe_names if p != "ner"]
    with nlp.disable_pipes(*other_pipes):
        optimizer = nlp.begin_training()
        for itn in range(n_iter):
            losses = {}
            random.shuffle(train_docs)
            batches = minibatch(train_docs, size=compounding(4., 32., 1.001))
            for batch in batches:
                examples = []
                for doc in batch:
                    entities = [(ent.start_char, ent.end_char, ent.label_) for ent in doc.ents]
                    example = Example.from_dict(doc, {"entities": entities})
                    examples.append(example)
                nlp.update(examples, drop=0.2, sgd=optimizer, losses=losses)
            print(f"Iteratie {itn} voor veld {field} | Verlies: {losses}")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    nlp.to_disk(output_dir)
    print(f"Model voor veld {field} opgeslagen in {output_dir}")

# ---------------------------------------------------------
# Aangepaste incremental_update_model_from_correction (veld-specifiek)
# ---------------------------------------------------------
def incremental_update_model_from_correction(pdf_path: str, corrections: dict, nlp, field: str, n_iter: int = 5, drop=0.1):
    try:
        text = extract_text(get_pdf_path(pdf_path))
    except Exception as e:
        print(f"Fout bij extractie tijdens online training voor {pdf_path}: {e}")
        return
    doc = nlp.make_doc(text)
    ents = []
    mapping = {
        "factuur_nummer": "FACTUUR_NUMMER",
        "factuur_datum": "FACTUUR_DATUM",
        "bedrag": "BEDRAG",
        "leverancier": "LEVERANCIER",
        "projid": "PROJID",
        "inkooporder_nummer": "INKOOPORDER_NUMMER",
        "geadresseerde": "GEADRESSEERDE",
    }
    for key, value in corrections.items():
        if key not in mapping:
            continue
        if value is not None:
            label = mapping.get(key)
            start = text.find(value) if value else -1
            if start == -1 and value:
                start = find_corrected_value_location(text, value)
            if start != -1:
                end = start + len(value)
                span = doc.char_span(start, end, label=label)
                if span is not None:
                    ents.append(span)
    ents = filter_spans(ents)
    doc.ents = ents
    gold_entities = [(ent.start_char, ent.end_char, ent.label_) for ent in ents]
    example = Example.from_dict(doc, {"entities": gold_entities})
    optimizer = nlp.resume_training() if hasattr(nlp, "resume_training") else nlp.begin_training()
    for i in range(n_iter):
        nlp.update([example], sgd=optimizer, drop=drop, losses={})
    nlp.to_disk(get_model_dir_for_field(field))
    print(f"Bijgewerkt (single) model voor veld {field} opgeslagen op: {get_model_dir_for_field(field)}")

# ---------------------------------------------------------
# Nieuwe functie: Batchgewijze incrementele training
# ---------------------------------------------------------
def incremental_update_models_from_batch(batch_corrections, nlp_models, n_iter: int = 5, drop=0.1):
    """
    batch_corrections: list van tuples (pdf_path, corrections_dict)
    nlp_models: dict van veld -> spaCy model
    """
    for field in FIELD_MAPPING.keys():
         examples = []
         model = nlp_models.get(field)
         if not model:
             continue
         for pdf_path, corrections in batch_corrections:
             if field in corrections and corrections[field]:
                 try:
                     text = extract_text(get_pdf_path(pdf_path))
                 except Exception as e:
                     print(f"Fout bij extractie tijdens batch training voor {pdf_path}: {e}")
                     continue
                 doc = model.make_doc(text)
                 label = FIELD_MAPPING[field]
                 value = corrections[field]
                 start = text.find(value) if value else -1
                 if start == -1:
                     start = find_corrected_value_location(text, value)
                 if start != -1:
                     end = start + len(value)
                     span = doc.char_span(start, end, label=label)
                     if span is not None:
                         doc.ents = [span]
                         example = Example.from_dict(doc, {"entities": [(start, end, label)]})
                         examples.append(example)
         if examples:
             optimizer = model.resume_training() if hasattr(model, "resume_training") else model.begin_training()
             for i in range(n_iter):
                 losses = {}
                 model.update(examples, sgd=optimizer, drop=drop, losses=losses)
             model.to_disk(get_model_dir_for_field(field))
             print(f"Bijgewerkt batch model voor veld {field} met {len(examples)} voorbeelden.")

# ---------------------------------------------------------
# Opslag van trainingscorrecties in SQLite
# ---------------------------------------------------------
def init_training_db():
    conn = sqlite3.connect(DB_TRAINING)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS training_corrections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pdf_path TEXT,
            pdf_text TEXT,
            field TEXT,
            original_value TEXT,
            corrected_value TEXT,
            feedback TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def store_training_correction(pdf_path, pdf_text, field, original_value, corrected_value, feedback):
    conn = sqlite3.connect(DB_TRAINING)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO training_corrections (pdf_path, pdf_text, field, original_value, corrected_value, feedback)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (pdf_path, pdf_text, field, original_value, corrected_value, feedback))
    conn.commit()
    conn.close()

def load_training_corrections_df():
    conn = sqlite3.connect(DB_TRAINING)
    df = pd.read_sql_query("SELECT * FROM training_corrections", conn)
    conn.close()
    return df

# -------------------------------------
# Leveranciersbeheer: leverancier nummer opvragen
# -------------------------------------
def open_supplier_number_popup():
    predefined_numbers = ["L00001", "L00002", "L00003", "L00004", "L00005"]
    root = tk.Tk()
    root.withdraw()
    nummer = simpledialog.askstring("Leverancier nummer", "Kies een leverancier nummer uit of voer een custom nummer in (bijv. L00001):\n" + ", ".join(predefined_numbers))
    root.destroy()
    if nummer is None:
        return None
    nummer = nummer.strip().upper()
    import re
    if re.match(r"^L\d{5}$", nummer):
        return nummer
    else:
        messagebox.showerror("Invoer fout", "Leverancier nummer moet beginnen met 'L' gevolgd door 5 cijfers (bijv. L00001).")
        return open_supplier_number_popup()

def prompt_for_geadresseerde_choice(geadresseerde):
    root = tk.Tk()
    root.withdraw()
    choice = simpledialog.askstring("Geadresseerde keuze", f"Voer de submap in voor geadresseerde '{geadresseerde}':")
    root.destroy()
    return choice

# -------------------------------------
# Functie om PDF te verplaatsen (inclusief leveranciersnummer)
# -------------------------------------
def move_processed_file(pdf_path: str, base_destination_folder: str, geadresseerde: str):
    if not os.path.exists(pdf_path):
        messagebox.showerror("Bestand niet gevonden", f"Het bestand {pdf_path} is niet gevonden.")
        return "", ""
    subfolder_map = load_geadresseerde_map()
    geo_lower = geadresseerde.strip().lower()
    if geo_lower in [g.lower() for g in KNOWN_GEADRESSEERDEN]:
        chosen_subfolder = None
        for sub in KNOWN_GEADRESSEERDEN:
            if sub.lower() == geo_lower:
                chosen_subfolder = sub
                break
    else:
        if geo_lower in subfolder_map:
            chosen_subfolder = subfolder_map[geo_lower]
        else:
            new_choice = prompt_for_geadresseerde_choice(geadresseerde)
            if not new_choice:
                new_choice = "GV"
            chosen_subfolder = new_choice
            if geo_lower:
                subfolder_map[geo_lower] = new_choice
                save_geadresseerde_map(subfolder_map)
    if not chosen_subfolder:
        chosen_subfolder = "GV"
    chosen_subfolder = chosen_subfolder.upper()
    dest_folder = os.path.join(base_destination_folder, chosen_subfolder)
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder, exist_ok=True)
    counters = load_counters()
    current_count = counters.get(chosen_subfolder, 300000)
    new_count = current_count + 1
    counters[chosen_subfolder] = new_count
    save_counters(counters)
    new_filename = f"{new_count}.pdf"
    dest_path = os.path.join(dest_folder, new_filename)
    try:
        shutil.move(pdf_path, dest_path)
    except Exception as e:
        messagebox.showerror("Verplaatsingsfout", f"Fout bij verplaatsen van {pdf_path} naar {dest_path}:\n{e}")
        return "", ""
    return dest_path, chosen_subfolder

def append_correction_to_csv(final_pdf_path: str, corrected_data: dict, csv_path: str):
    row = {
        "pdf_path": final_pdf_path,
        "geadresseerde": corrected_data.get("geadresseerde", ""),
        "leverancier": corrected_data.get("leverancier", ""),
        "factuur_nummer": corrected_data.get("factuur_nummer", ""),
        "factuur_datum": corrected_data.get("factuur_datum", ""),
        "projid": corrected_data.get("projid", ""),
        "inkooporder_nummer": corrected_data.get("inkooporder_nummer", ""),
        "bedrag": corrected_data.get("bedrag", ""),
        "dagboek": corrected_data.get("dagboek", ""),
    }
    new_df = pd.DataFrame([row])
    if not os.path.exists(csv_path):
        new_df.to_csv(csv_path, index=False)
    else:
        df = pd.read_csv(csv_path)
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_csv(csv_path, index=False)

def append_label_annotations(pdf_path: str, corrected_data: dict, annotation_csv: str):
    if not pdf_path or not os.path.exists(pdf_path):
        return
    try:
        text = extract_text(pdf_path)
    except Exception as e:
        print(f"append_label_annotations: fout bij extract_text: {e}")
        return
    mapping = {
        "factuur_nummer": ("FACTUUR_NUMMER", "factuur_nummer"),
        "factuur_datum": ("FACTUUR_DATUM", "factuur_datum"),
        "bedrag": ("BEDRAG", "bedrag"),
        "leverancier": ("LEVERANCIER", "leverancier"),
        "projid": ("PROJID", "projid"),
        "inkooporder_nummer": ("INKOOPORDER_NUMMER", "inkooporder_nummer"),
        "geadresseerde": ("GEADRESSEERDE", "geadresseerde"),
    }
    rows = []
    for key, val in corrected_data.items():
        if key not in mapping or not val:
            continue
        label_spacy, label_str = mapping[key]
        start_offset = text.find(val)
        if start_offset == -1:
            start_offset = find_corrected_value_location(text, val)
        end_offset = -1
        if start_offset != -1:
            end_offset = start_offset + len(val)
        row_data = {
            "pdf_path": pdf_path,
            "label": label_spacy,
            "text_value": val,
            "start_offset": start_offset,
            "end_offset": end_offset
        }
        rows.append(row_data)
    if not rows:
        return
    out_df = pd.DataFrame(rows)
    if os.path.exists(annotation_csv):
        old_df = pd.read_csv(annotation_csv)
        combined = pd.concat([old_df, out_df], ignore_index=True)
        combined.to_csv(annotation_csv, index=False)
    else:
        out_df.to_csv(annotation_csv, index=False)

def append_invoice_to_excel(final_pdf_path: str, corrected_data: dict, submap: str, excel_path: str):
    row = {
        "pdf_path": final_pdf_path,
        "geadresseerde": corrected_data.get("geadresseerde", ""),
        "leverancier": corrected_data.get("leverancier", ""),
        "factuur_nummer": corrected_data.get("factuur_nummer", ""),
        "factuur_datum": corrected_data.get("factuur_datum", ""),
        "projid": corrected_data.get("projid", ""),
        "inkooporder_nummer": corrected_data.get("inkooporder_nummer", ""),
        "bedrag": corrected_data.get("bedrag", ""),
        "booked": False,
        "new_filename": os.path.basename(final_pdf_path),
        "dagboek": corrected_data.get("dagboek", ""),
        "submap": submap
    }
    new_df = pd.DataFrame([row])
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            print(f"Fout bij laden van extra Excel: {e}")
            df = pd.DataFrame()
        df = pd.concat([df, new_df], ignore_index=True)
    else:
        df = new_df
    df.to_excel(excel_path, index=False)

def reset_model():
    if os.path.exists(TRAINED_MODEL_DIR):
        try:
            shutil.rmtree(TRAINED_MODEL_DIR)
            print(f"Modelmap '{TRAINED_MODEL_DIR}' is verwijderd.")
        except Exception as e:
            print(f"Fout bij verwijderen van de modelmap: {e}")
    else:
        print("Geen getraind model gevonden om te verwijderen.")
    try:
        nlp = spacy.load(BASE_SPACY_MODEL)
        os.makedirs(TRAINED_MODEL_DIR, exist_ok=True)
        nlp.to_disk(TRAINED_MODEL_DIR)
        print("Model is gereset: Basismodel opnieuw geladen en opgeslagen.")
        return nlp
    except Exception as e:
        print(f"Fout bij herladen van het basismodel: {e}")
        return None

# =============================================================================
# Aangepaste extract_invoice_data functie (met aparte modellen)
# =============================================================================
def extract_invoice_data(pdf_path: str, nlp_models: dict = None, use_regex_heuristieken: bool = True) -> dict:
    actual_pdf_path = get_pdf_path(pdf_path)
    try:
        text = extract_text(actual_pdf_path)
    except Exception as e:
        print(f"Fout bij extractie van tekst uit {actual_pdf_path}: {e}")
        text = ""
    results = { field: {'value': None, 'certainty': 0.0, 'source': ""} for field in FIELD_MAPPING.keys() }
    if nlp_models and text:
        for field, model in nlp_models.items():
            if model:
                doc = model(text)
                for ent in doc.ents:
                    if ent.label_ == FIELD_MAPPING[field]:
                        results[field] = {'value': ent.text, 'certainty': AI_CONFIDENCE, 'source': "AI"}
                        break
    if use_regex_heuristieken:
        if not results['factuur_nummer']['value']:
            fn = heuristic_find_factuurnummer(text)
            if fn:
                results['factuur_nummer'] = {'value': fn, 'certainty': 0.8, 'source': "Regex"}
        if not results['factuur_datum']['value']:
            dt = heuristic_find_date(text)
            # Alleen als de gevonden datum effectief is (volledig) overnemen
            if dt and is_effective_date(dt):
                results['factuur_datum'] = {'value': dt, 'certainty': 0.8, 'source': "Regex"}
        if not results['projid']['value']:
            projid_mapping = lookup_proj_mapping(text)
            if projid_mapping:
                results['projid'] = {'value': projid_mapping, 'certainty': 1.0, 'source': "Projid mapping"}
            else:
                pr = heuristic_find_projid(text)
                if pr:
                    results['projid'] = {'value': pr, 'certainty': 0.8, 'source': "Regex"}
        if not results['inkooporder_nummer']['value']:
            io = heuristic_find_inkoopordernummer(text)
            if io:
                results['inkooporder_nummer'] = {'value': io, 'certainty': 0.8, 'source': "Regex"}
        if not results['bedrag']['value']:
            bd = heuristic_find_bedrag(text)
            if bd:
                results['bedrag'] = {'value': bd, 'certainty': 0.8, 'source': "Regex"}
    # Voor leverancier: altijd lookup in bekende waarden
    supplier_known = lookup_supplier_from_known_values(text)
    if supplier_known:
         results['leverancier'] = {'value': supplier_known, 'certainty': 1.0, 'source': "Known lookup"}
    # Voor geadresseerde: lookup in known values, 100% zekerheid als match
    geo_known = lookup_geadresseerde_from_known_values(text)
    if geo_known:
        results['geadresseerde'] = {'value': geo_known, 'certainty': 1.0, 'source': "Known lookup"}
    # Als leverancier al is gevonden via AI of regex, update overige velden op basis van geleerde offsets.
    supplier_match = results['leverancier']['value']
    if supplier_match:
        for field in ['geadresseerde','factuur_nummer','factuur_datum','projid','inkooporder_nummer','bedrag']:
            learned = get_value_from_learned_offset(text, supplier_match, field)
            if learned:
                # Alleen bij factuur_datum: als de geleerde datum niet effectief is (bijv. "30/11/202"), negeren.
                if field == 'factuur_datum' and not is_effective_date(learned):
                    continue
                results[field] = {'value': learned, 'certainty': 0.95, 'source': "Geleerde locatie"}
    return results

# -------------------------------------
# HOOFD-GUI
# -------------------------------------
class InvoiceGUI:
    def __init__(self, master):
        self.master = master
        self.master.geometry(MAIN_WINDOW_GEOMETRY)
        self.master.title("Factuurverwerking - Geavanceerd")
        style = ttk.Style()
        style.theme_use("clam")
        background_main = "#F2F2F2"
        frame_bg = "#F8F8F8"
        text_fg = "#333333"
        button_bg = "#DDDDDD"
        button_fg = "#000000"
        style.configure(".", font=("Helvetica", 10))
        style.configure("TFrame", background=frame_bg)
        style.configure("TLabelFrame", background=frame_bg)
        style.configure("TLabel", background=frame_bg, foreground=text_fg)
        style.configure("TButton", background=button_bg, foreground=button_fg, relief="flat", padding=6, font=("Helvetica", 9, "bold"))
        style.map("TButton", foreground=[("active", "#ffffff")], background=[("active", "#888888")])
        self.master.configure(bg=background_main)
        self.frame_top = ttk.Frame(master, padding="5")
        self.frame_top.pack(side=tk.TOP, fill=tk.X)
        self.pdf_counter_label = ttk.Label(self.frame_top, text="Bestand 0/0")
        self.pdf_counter_label.pack(side=tk.LEFT, padx=5)
        self.btn_open_pdf = ttk.Button(self.frame_top, text="PDF openen", command=self.open_pdf)
        self.btn_open_pdf.pack(side=tk.LEFT, padx=5)
        self.btn_skip = ttk.Button(self.frame_top, text="Bestand overslaan", command=self.skip_file)
        self.btn_skip.pack(side=tk.LEFT, padx=5)
        self.btn_config = ttk.Button(self.frame_top, text="Configuratie", command=self.open_config_window)
        self.btn_config.pack(side=tk.LEFT, padx=5)
        self.btn_mapping_editor = ttk.Button(self.frame_top, text="ProjID-koppelingen", command=self.open_mapping_editor)
        self.btn_mapping_editor.pack(side=tk.LEFT, padx=5)
        self.btn_geadres_mapping = ttk.Button(self.frame_top, text="Geadresseerde-koppelingen", command=self.open_geadres_mapping_editor)
        self.btn_geadres_mapping.pack(side=tk.LEFT, padx=5)
        self.btn_suppliers = ttk.Button(self.frame_top, text="Leveranciers beheren", command=self.open_supplier_manager)
        self.btn_suppliers.pack(side=tk.LEFT, padx=5)
        self.btn_train_mode = ttk.Button(self.frame_top, text="Trainingsmodus", command=self.open_trainer_mode)
        self.btn_train_mode.pack(side=tk.LEFT, padx=5)
        self.btn_booking_mode = ttk.Button(self.frame_top, text="Boekingsmodus", command=self.open_booking_mode)
        self.btn_booking_mode.pack(side=tk.LEFT, padx=5)
        self.status_label = ttk.Label(self.master, text="Status: initialiseren...", background=background_main)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        self.frame_content = ttk.Frame(master, padding="5")
        self.frame_content.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
        self.frame_text = ttk.LabelFrame(self.frame_content, text="Geëxtraheerde PDF-tekst", padding="5")
        self.frame_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        self.pdf_text_widget = tk.Text(self.frame_text, wrap="word", width=60, height=25)
        self.pdf_text_scroll = ttk.Scrollbar(self.frame_text, orient="vertical", command=self.pdf_text_widget.yview)
        self.pdf_text_widget.configure(yscrollcommand=self.pdf_text_scroll.set)
        self.pdf_text_widget.tag_config("highlight", background="yellow")
        self.pdf_text_widget.pack(side="left", fill=tk.BOTH, expand=True)
        self.pdf_text_scroll.pack(side="right", fill=tk.Y)
        self.frame_right = ttk.LabelFrame(self.frame_content, text="Gedetecteerde gegevens", padding="10")
        self.frame_right.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        self.fields = {}
        self.field_sources = {}
        field_names = ['geadresseerde', 'leverancier', 'factuur_nummer', 'factuur_datum', 'projid', 'inkooporder_nummer', 'bedrag', 'dagboek']
        for idx, field_name in enumerate(field_names):
            lbl = ttk.Label(self.frame_right, text=f"{field_name}:")
            lbl.grid(row=idx, column=0, padx=5, pady=6, sticky="e")
            var = tk.StringVar()
            combobox = ttk.Combobox(self.frame_right, textvariable=var, width=35)
            if field_name == "leverancier":
                suppliers = load_known_suppliers()
                combobox['values'] = list(suppliers.keys())
                combobox.configure(state="normal")
            elif field_name == "dagboek":
                combobox['values'] = ["AKR", "AKC", "APF", "APC"]
                combobox.configure(state="normal")
            else:
                combobox.configure(state="normal")
            combobox.grid(row=idx, column=1, padx=5, pady=6, sticky="w")
            combobox.bind("<<ComboboxSelected>>", lambda e, f=field_name: self.on_field_change(f))
            combobox.bind("<KeyRelease>", lambda e, f=field_name: self.on_field_change(f))
            indicator = ttk.Label(self.frame_right, text="", width=20)
            indicator.grid(row=idx, column=2, padx=5, pady=6, sticky="w")
            self.fields[field_name] = {'label': lbl, 'combobox': combobox, 'var': var, 'indicator': indicator}
            self.field_sources[field_name] = ""
        self.btn_save = ttk.Button(self.frame_right, text="Opslaan (correcties)", command=self.save_data)
        self.btn_save.grid(row=len(field_names), column=0, columnspan=3, pady=15)
        self.current_pdf_path = None
        self.pdf_text = ""
        # Laden van veldspecifieke modellen
        try:
            self.nlp_models = load_all_field_models()
        except Exception as e:
            print(f"Geen bestaande modellen gevonden: {e}")
            self.nlp_models = {}
        self.train_batch_buffer = []  # Buffer voor batchgewijze online training (5 facturen per batch)
        self.pdf_list = []
        self.current_index = 0
        self.pdf_preview_window = tk.Toplevel(self.master)
        self.pdf_preview_window.title("PDF-preview")
        self.pdf_preview_window.geometry(PDF_PREVIEW_GEOMETRY)
        self.zoom_frame = ttk.Frame(self.pdf_preview_window, padding="5")
        self.zoom_frame.pack(side=tk.TOP, fill=tk.X)
        self.btn_zoom_in = ttk.Button(self.zoom_frame, text="Zoom in", command=self.zoom_in)
        self.btn_zoom_in.pack(side=tk.LEFT, padx=2)
        self.btn_zoom_out = ttk.Button(self.zoom_frame, text="Zoom uit", command=self.zoom_out)
        self.btn_zoom_out.pack(side=tk.LEFT, padx=2)
        self.btn_zoom_reset = ttk.Button(self.zoom_frame, text="Zoom resetten", command=self.reset_zoom)
        self.btn_zoom_reset.pack(side=tk.LEFT, padx=2)
        self.pdf_preview_container = ttk.Frame(self.pdf_preview_window)
        self.pdf_preview_container.pack(expand=True, fill=tk.BOTH)
        self.default_dpi = 100
        self.zoom_factor = PDF_PREVIEW_ZOOM
        self.preprocessed_data = {}
        self.master.bind("<Control-f>", self.open_find_dialog)
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        if INVOICE_FOLDER and os.path.isdir(INVOICE_FOLDER):
            self.pdf_list = [os.path.join(INVOICE_FOLDER, f) for f in os.listdir(INVOICE_FOLDER) if f.lower().endswith(".pdf")]
            self.pdf_list.sort()
            if self.pdf_list:
                self.preprocess_first_n_invoices(n=10)
                if len(self.pdf_list) > 10:
                    bg_thread = threading.Thread(target=self.preprocess_invoices_background, args=(self.pdf_list[10:],), daemon=True)
                    bg_thread.start()
                self.current_index = 0
                self.load_invoice(self.pdf_list[self.current_index])
            else:
                messagebox.showwarning("Geen PDF's", "Er zijn geen PDF-bestanden gevonden in de map: " + INVOICE_FOLDER)
                self.status_label.config(text="Status: Geen PDF-bestanden gevonden.")
        else:
            messagebox.showwarning("Map niet gevonden", "De geconfigureerde INVOICE_FOLDER bestaat niet.")
            self.status_label.config(text="Status: INVOICE_FOLDER niet gevonden.")

    def update_known_lists(self):
        self.fields["leverancier"]['combobox']['values'] = list(load_known_suppliers().keys())
        geo_map = load_geadresseerde_map()
        all_geos = list(set(KNOWN_GEADRESSEERDEN).union(set(geo_map.keys())))
        self.fields["geadresseerde"]['combobox']['values'] = all_geos

    def open_find_dialog(self, event=None):
        dialog = tk.Toplevel(self.master)
        dialog.title("Zoek in tekst")
        dialog.geometry("300x100")
        lbl = ttk.Label(dialog, text="Zoekterm:")
        lbl.pack(side=tk.TOP, padx=5, pady=5)
        find_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=find_var)
        entry.pack(side=tk.TOP, padx=5, pady=5)
        entry.focus()
        def do_search():
            term = find_var.get()
            if term:
                self.search_in_text(term)
        btn = ttk.Button(dialog, text="Zoeken", command=do_search)
        btn.pack(side=tk.TOP, padx=5, pady=5)

    def search_in_text(self, term):
        self.pdf_text_widget.tag_remove("search_highlight", '1.0', tk.END)
        start_pos = '1.0'
        first_idx = None
        while True:
            idx = self.pdf_text_widget.search(term, start_pos, nocase=1, stopindex=tk.END)
            if not idx:
                break
            if first_idx is None:
                first_idx = idx
            end_pos = f"{idx}+{len(term)}c"
            self.pdf_text_widget.tag_add('search_highlight', idx, end_pos)
            start_pos = end_pos
        self.pdf_text_widget.tag_config('search_highlight', background='yellow', foreground='black')
        if first_idx:
            self.pdf_text_widget.see(first_idx)
        else:
            messagebox.showinfo("Zoekresultaat", f"'{term}' niet gevonden.")

    def preprocess_first_n_invoices(self, n=10):
        start_time = time.time()
        limited_pdf_list = self.pdf_list[:n]
        total = len(limited_pdf_list)
        for i, pdf_path in enumerate(limited_pdf_list, start=1):
            filename_only = os.path.basename(pdf_path)
            self.status_label.config(text=f"Preprocess {i}/{total}: {filename_only} ...")
            self.master.update_idletasks()
            try:
                text = extract_text(pdf_path)
            except Exception as e:
                print(f"Fout bij het extraheren van {pdf_path}: {e}")
                text = ""
            invoice_data = extract_invoice_data(pdf_path, nlp_models=self.nlp_models, use_regex_heuristieken=True) if text else {}
            self.preprocessed_data[pdf_path] = (text, invoice_data)
        elapsed = time.time() - start_time
        self.status_label.config(text=f"Eerste {total} facturen verwerkt in {elapsed:.2f} sec. De rest gebeurt op de achtergrond.")
        self.master.update_idletasks()

    def preprocess_invoices_background(self, pdf_paths):
        start_time = time.time()
        total = len(pdf_paths)
        for i, pdf_path in enumerate(pdf_paths, start=1):
            filename_only = os.path.basename(pdf_path)
            print(f"[BG] Preprocess {i}/{total}: {filename_only} ...")
            try:
                text = extract_text(pdf_path)
            except Exception as e:
                print(f"[BG] Fout bij extractie {pdf_path}: {e}")
                text = ""
            invoice_data = extract_invoice_data(pdf_path, nlp_models=self.nlp_models, use_regex_heuristieken=True) if text else {}
            self.preprocessed_data[pdf_path] = (text, invoice_data)
        elapsed = time.time() - start_time
        print(f"[BG] Achtergrondverwerking van {total} PDF's klaar in {elapsed:.2f} sec.")

    def on_close(self):
        current_main_geo = self.master.winfo_geometry()
        try:
            if self.pdf_preview_window.winfo_exists():
                current_preview_geo = self.pdf_preview_window.winfo_geometry()
            else:
                current_preview_geo = PDF_PREVIEW_GEOMETRY
        except Exception as e:
            current_preview_geo = PDF_PREVIEW_GEOMETRY
        if "UI_STATE" not in config_data:
            config_data["UI_STATE"] = {}
        config_data["UI_STATE"]["main_window_geometry"] = current_main_geo
        config_data["UI_STATE"]["pdf_preview_geometry"] = current_preview_geo
        config_data["UI_STATE"]["pdf_zoom_factor"] = self.zoom_factor
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            print(f"Fout bij opslaan van de UI-status: {e}")
        self.master.destroy()

    def open_supplier_manager(self):
        sup_win = tk.Toplevel(self.master)
        sup_win.title("Leveranciers beheren")
        sup_win.geometry("600x400")
        columns = ("naam", "nummer")
        tree = ttk.Treeview(sup_win, columns=columns, show="headings", selectmode="browse")
        tree.heading("naam", text="Leverancier")
        tree.heading("nummer", text="Leverancier nummer")
        tree.column("naam", width=300)
        tree.column("nummer", width=150)
        tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        def refresh_tree():
            for row in tree.get_children():
                tree.delete(row)
            suppliers = load_known_suppliers()
            for naam, nummer in suppliers.items():
                tree.insert("", tk.END, values=(naam, nummer))
        refresh_tree()
        def add_or_update_supplier():
            naam = naam_var.get().strip()
            if not naam:
                messagebox.showwarning("Invoer fout", "Voer een leveranciersnaam in.")
                return
            nummer = open_supplier_number_popup()
            if nummer is None:
                return
            suppliers = load_known_suppliers()
            suppliers[naam] = nummer
            save_known_suppliers(suppliers)
            naam_var.set("")
            refresh_tree()
            self.fields["leverancier"]['combobox']['values'] = list(load_known_suppliers().keys())
        def delete_supplier():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Selectie", "Selecteer een leverancier om te verwijderen.")
                return
            item = tree.item(selected)
            naam = item["values"][0]
            suppliers = load_known_suppliers()
            if naam in suppliers:
                del suppliers[naam]
                save_known_suppliers(suppliers)
                refresh_tree()
                self.fields["leverancier"]['combobox']['values'] = list(load_known_suppliers().keys())
        frm_entries = ttk.Frame(sup_win)
        frm_entries.pack(padx=10, pady=5, fill=tk.X)
        ttk.Label(frm_entries, text="Leverancier naam:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        naam_var = tk.StringVar()
        ent_naam = ttk.Entry(frm_entries, textvariable=naam_var, width=40)
        ent_naam.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        btn_frame = ttk.Frame(sup_win)
        btn_frame.pack(padx=10, pady=5)
        btn_add = ttk.Button(btn_frame, text="Toevoegen/Bijwerken", command=add_or_update_supplier)
        btn_add.grid(row=0, column=0, padx=5)
        btn_delete = ttk.Button(btn_frame, text="Verwijderen", command=delete_supplier)
        btn_delete.grid(row=0, column=1, padx=5)

    def open_geadres_mapping_editor(self):
        geadres_win = tk.Toplevel(self.master)
        geadres_win.title("Geadresseerde-koppelingen")
        geadres_win.geometry("550x400")
        listbox = tk.Listbox(geadres_win, width=60)
        listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        def refresh_listbox():
            listbox.delete(0, tk.END)
            mapping = load_geadresseerde_map()
            for g_name, subf in mapping.items():
                listbox.insert(tk.END, f"{g_name}  -->  {subf}")
        refresh_listbox()
        frm_entries = ttk.Frame(geadres_win)
        frm_entries.pack(padx=10, pady=5, fill=tk.X)
        ttk.Label(frm_entries, text="Geadresseerde (naam):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        gea_var = tk.StringVar()
        ent_gea = ttk.Entry(frm_entries, textvariable=gea_var, width=40)
        ent_gea.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frm_entries, text="Submap:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        subf_var = tk.StringVar()
        combo_subf = ttk.Combobox(frm_entries, textvariable=subf_var, width=30, values=KNOWN_GEADRESSEERDEN, state="readonly")
        combo_subf.grid(row=1, column=1, padx=5, pady=5)
        def add_mapping():
            g_name = gea_var.get().strip().lower()
            chosen_sub = subf_var.get().strip()
            if g_name and chosen_sub:
                current_map = load_geadresseerde_map()
                current_map[g_name] = chosen_sub
                save_geadresseerde_map(current_map)
                gea_var.set("")
                subf_var.set("")
                refresh_listbox()
        def delete_mapping():
            selection = listbox.curselection()
            if selection:
                item = listbox.get(selection[0])
                parts = item.split("  -->  ")
                if len(parts) == 2:
                    g_name = parts[0].strip()
                    current_map = load_geadresseerde_map()
                    if g_name in current_map:
                        del current_map[g_name]
                        save_geadresseerde_map(current_map)
                        refresh_listbox()
        def load_selected_mapping(event):
            selection = listbox.curselection()
            if selection:
                item = listbox.get(selection[0])
                parts = item.split("  -->  ")
                if len(parts) == 2:
                    gea_var.set(parts[0].strip())
                    subf_var.set(parts[1].strip())
        listbox.bind("<<ListboxSelect>>", load_selected_mapping)
        btn_frame = ttk.Frame(geadres_win)
        btn_frame.pack(padx=10, pady=5)
        btn_add = ttk.Button(btn_frame, text="Toevoegen / Bijwerken", command=add_mapping)
        btn_add.grid(row=0, column=0, padx=5)
        btn_delete = ttk.Button(btn_frame, text="Verwijderen", command=delete_mapping)
        btn_delete.grid(row=0, column=1, padx=5)

    def update_pdf_preview(self):
        if self.current_pdf_path:
            dpi = int(self.default_dpi * self.zoom_factor)
            show_pdf_scrollable_preview(self.current_pdf_path, self.pdf_preview_container, dpi=dpi)

    def zoom_in(self):
        self.zoom_factor += 0.1
        self.update_pdf_preview()

    def zoom_out(self):
        if self.zoom_factor > 0.2:
            self.zoom_factor -= 0.1
            self.update_pdf_preview()

    def reset_zoom(self):
        self.zoom_factor = PDF_PREVIEW_ZOOM
        self.update_pdf_preview()

    def on_field_change(self, field_name):
        current_value = self.fields[field_name]['var'].get().strip()
        source = self.field_sources.get(field_name, "")
        if not current_value and field_name in ["projid", "inkooporder_nummer"]:
            self.fields[field_name]['indicator'].configure(text="   Geen waarde", foreground="#800080")
        else:
            found = False
            if self.pdf_text and current_value:
                if current_value.lower() in self.pdf_text.lower():
                    found = True
                else:
                    candidates = self.candidate_data.get(field_name, []) if hasattr(self, "candidate_data") else []
                    for cand in candidates:
                        if cand.strip().lower() == current_value.lower():
                            found = True
                            break
            found_text = "Gevonden" if found else "Niet gevonden"
            color = "green" if found else "red"
            self.fields[field_name]['indicator'].configure(text=f"   {found_text}  ({source})", foreground=color)

    def assign_to_field(self, field_name, selected_text):
        self.fields[field_name]['var'].set(selected_text)
        self.field_sources[field_name] = "Handmatig"
        self.on_field_change(field_name)
        messagebox.showinfo("Toegewezen", f"'{selected_text}' is toegewezen aan {field_name}.")

    def highlight_extracted_data(self, invoice_data):
        self.pdf_text_widget.tag_remove("highlight", "1.0", tk.END)
        for field, data in invoice_data.items():
            value = data.get('value')
            if value:
                start_idx = 0
                while True:
                    start_pos = self.pdf_text.find(value, start_idx)
                    if start_pos == -1:
                        break
                    end_pos = start_pos + len(value)
                    text_start_idx = f"1.0+{start_pos}c"
                    text_end_idx = f"1.0+{end_pos}c"
                    self.pdf_text_widget.tag_add("highlight", text_start_idx, text_end_idx)
                    start_idx = end_pos

    def open_config_window(self):
        config_win = tk.Toplevel(self.master)
        config_win.title("Configuratie")
        config_win.geometry("650x450")
        config_params = {
            "TESSERACT_CMD": TESSERACT_CMD,
            "POPLER_BIN_PATH": POPLER_BIN_PATH,
            "TRAINED_MODEL_DIR": TRAINED_MODEL_DIR,
            "PROCESSED_PDF_DIR": PROCESSED_PDF_DIR,
            "CORRECTED_DATA_FILE": CORRECTED_DATA_FILE,
            "TRAINING_DATA_EXCEL": TRAINING_DATA_EXCEL,
            "BASE_SPACY_MODEL": BASE_SPACY_MODEL,
            "INVOICE_FOLDER": INVOICE_FOLDER,
            "AI_CONFIDENCE": AI_CONFIDENCE,
            "OCR_MINTEXT": OCR_MINTEXT,
            "INCREMENTAL_BATCH_SIZE": INCREMENTAL_BATCH_SIZE
        }
        entries = {}
        frm_main = ttk.Frame(config_win, padding=10)
        frm_main.pack(fill="both", expand=True)
        row = 0
        for key, value in config_params.items():
            lbl = ttk.Label(frm_main, text=key)
            lbl.grid(row=row, column=0, padx=5, pady=5, sticky="e")
            var = tk.StringVar(value=str(value))
            ent = ttk.Entry(frm_main, textvariable=var, width=50)
            ent.grid(row=row, column=1, padx=5, pady=5, sticky="w")
            entries[key] = var
            row += 1
        ocr_var = tk.BooleanVar(value=ENABLE_OCR_FALLBACK)
        ckb_ocr = ttk.Checkbutton(frm_main, text="OCR Fallback inschakelen?", variable=ocr_var)
        ckb_ocr.grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        training_frame = ttk.LabelFrame(frm_main, text="Trainingsopties", padding=10)
        training_frame.grid(row=row, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        row += 1
        btn_train = ttk.Button(training_frame, text="Model trainen", command=self.train_model_action)
        btn_train.grid(row=0, column=0, padx=5, pady=5)
        btn_retrain = ttk.Button(training_frame, text="Model hertrain (met correcties)", command=self.retrain_with_corrections)
        btn_retrain.grid(row=0, column=1, padx=5, pady=5)
        btn_reset = ttk.Button(training_frame, text="Model resetten", command=self.reset_model_action)
        btn_reset.grid(row=0, column=2, padx=5, pady=5)
        def save_config():
            global TESSERACT_CMD, POPLER_BIN_PATH, TRAINED_MODEL_DIR
            global PROCESSED_PDF_DIR, CORRECTED_DATA_FILE, TRAINING_DATA_EXCEL
            global BASE_SPACY_MODEL, INVOICE_FOLDER, ENABLE_OCR_FALLBACK
            global AI_CONFIDENCE, OCR_MINTEXT, INCREMENTAL_BATCH_SIZE
            TESSERACT_CMD = entries["TESSERACT_CMD"].get()
            POPLER_BIN_PATH = entries["POPLER_BIN_PATH"].get()
            TRAINED_MODEL_DIR = entries["TRAINED_MODEL_DIR"].get()
            PROCESSED_PDF_DIR = entries["PROCESSED_PDF_DIR"].get()
            CORRECTED_DATA_FILE = entries["CORRECTED_DATA_FILE"].get()
            TRAINING_DATA_EXCEL = entries["TRAINING_DATA_EXCEL"].get()
            BASE_SPACY_MODEL = entries["BASE_SPACY_MODEL"].get()
            INVOICE_FOLDER = entries["INVOICE_FOLDER"].get()
            AI_CONFIDENCE = float(entries["AI_CONFIDENCE"].get() or 0.95)
            OCR_MINTEXT = int(entries["OCR_MINTEXT"].get() or 50)
            INCREMENTAL_BATCH_SIZE = int(entries["INCREMENTAL_BATCH_SIZE"].get() or 1)
            ENABLE_OCR_FALLBACK = ocr_var.get()
            config_data["TESSERACT_CMD"] = TESSERACT_CMD
            config_data["POPLER_BIN_PATH"] = POPLER_BIN_PATH
            config_data["TRAINED_MODEL_DIR"] = TRAINED_MODEL_DIR
            config_data["PROCESSED_PDF_DIR"] = PROCESSED_PDF_DIR
            config_data["CORRECTED_DATA_FILE"] = CORRECTED_DATA_FILE
            config_data["TRAINING_DATA_EXCEL"] = TRAINING_DATA_EXCEL
            config_data["BASE_SPACY_MODEL"] = BASE_SPACY_MODEL
            config_data["INVOICE_FOLDER"] = INVOICE_FOLDER
            config_data["AI_CONFIDENCE"] = AI_CONFIDENCE
            config_data["OCR_MINTEXT"] = OCR_MINTEXT
            config_data["INCREMENTAL_BATCH_SIZE"] = INCREMENTAL_BATCH_SIZE
            config_data["ENABLE_OCR_FALLBACK"] = ENABLE_OCR_FALLBACK
            try:
                with open(CONFIG_FILE, "w") as f:
                    json.dump(config_data, f, indent=4)
            except Exception as e:
                messagebox.showerror("Configuratie Opslaan", f"Fout bij opslaan: {e}")
            else:
                messagebox.showinfo("Configuratie", "Configuratie is opgeslagen.")
            config_win.destroy()
        save_button = ttk.Button(frm_main, text="Opslaan", command=save_config)
        save_button.grid(row=row, column=0, columnspan=2, pady=10)

    def reset_model_action(self):
        answer = messagebox.askyesno("Model resetten", "Weet u zeker dat u het model wilt resetten? Dit verwijdert alle veldspecifieke modellen en herlaadt het basismodel.")
        if answer:
            new_model = reset_model()
            if new_model:
                self.nlp_models = load_all_field_models()
                messagebox.showinfo("Reset geslaagd", "Alle veldspecifieke modellen zijn gereset en het basismodel is opnieuw geladen.")
            else:
                messagebox.showerror("Reset fout", "Fout bij resetten van de modellen.")
        else:
            messagebox.showinfo("Reset geannuleerd", "Reset is geannuleerd.")

    def open_pdf(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF-bestanden", "*.pdf")])
        if not pdf_path:
            return
        self.current_pdf_path = pdf_path
        try:
            self.pdf_text = extract_text(get_pdf_path(pdf_path))
        except Exception as e:
            messagebox.showerror("PDF-extractiefout", f"Fout bij het extraheren van tekst: {e}")
            return
        self.pdf_text_widget.delete("1.0", tk.END)
        self.pdf_text_widget.insert("1.0", self.pdf_text)
        show_pdf_scrollable_preview(pdf_path, self.pdf_preview_container, dpi=int(self.default_dpi * self.zoom_factor))
        self.update_known_lists()
        self.candidate_data = {
            'geadresseerde': get_candidate_geadresseerde(self.pdf_text),
            'leverancier': get_candidate_leveranciers(self.pdf_text),
            'factuur_nummer': get_candidate_factuurnummers(self.pdf_text),
            'factuur_datum': get_candidate_factuurdatums(self.pdf_text),
            'projid': get_candidate_projid(self.pdf_text),
            'inkooporder_nummer': get_candidate_inkoopordernummers(self.pdf_text),
            'bedrag': get_candidate_bedragen(self.pdf_text),
        }
        for field_name, field in self.fields.items():
            if field_name not in ["dagboek"]:
                field['combobox']['values'] = self.candidate_data.get(field_name, [])
        if not self.nlp_models:
            messagebox.showinfo("Model niet geladen", "Er zijn nog geen getrainde modellen. We gebruiken alleen regex/heuristieken.")
        result = extract_invoice_data(pdf_path, nlp_models=self.nlp_models, use_regex_heuristieken=True)
        self.fill_fields(result)
        self.highlight_extracted_data(result)
        self.pdf_counter_label.config(text="Losse PDF (niet in map)")

    def load_invoice(self, pdf_path):
        self.current_pdf_path = pdf_path
        if pdf_path in self.preprocessed_data:
            self.pdf_text, invoice_data = self.preprocessed_data[pdf_path]
        else:
            try:
                self.pdf_text = extract_text(get_pdf_path(pdf_path))
            except Exception as e:
                messagebox.showerror("PDF-extractiefout", f"Fout bij het extraheren van tekst: {e}")
                return
            invoice_data = extract_invoice_data(pdf_path, nlp_models=self.nlp_models, use_regex_heuristieken=True)
            self.preprocessed_data[pdf_path] = (self.pdf_text, invoice_data)
        self.pdf_text_widget.delete("1.0", tk.END)
        self.pdf_text_widget.insert("1.0", self.pdf_text)
        show_pdf_scrollable_preview(pdf_path, self.pdf_preview_container, dpi=int(self.default_dpi * self.zoom_factor))
        self.update_known_lists()
        self.candidate_data = {
            'geadresseerde': get_candidate_geadresseerde(self.pdf_text),
            'leverancier': get_candidate_leveranciers(self.pdf_text),
            'factuur_nummer': get_candidate_factuurnummers(self.pdf_text),
            'factuur_datum': get_candidate_factuurdatums(self.pdf_text),
            'projid': get_candidate_projid(self.pdf_text),
            'inkooporder_nummer': get_candidate_inkoopordernummers(self.pdf_text),
            'bedrag': get_candidate_bedragen(self.pdf_text),
        }
        for field_name, field in self.fields.items():
            if field_name != "dagboek":
                field['combobox']['values'] = self.candidate_data.get(field_name, [])
        self.fill_fields(invoice_data)
        self.highlight_extracted_data(invoice_data)
        total_count = len(self.pdf_list)
        current_number = self.current_index + 1
        self.pdf_counter_label.config(text=f"Bestand {current_number}/{total_count}")

    def fill_fields(self, result: dict):
        for field_name, data in result.items():
            val = data['value']
            self.fields[field_name]['var'].set(val if val else "")
            self.field_sources[field_name] = data.get('source', "")
            self.on_field_change(field_name)
        inkoopnr = self.fields["inkooporder_nummer"]['var'].get().strip()
        current_dagboek_val = self.fields["dagboek"]['var'].get().strip()
        if not current_dagboek_val:
            if inkoopnr:
                self.fields["dagboek"]['var'].set("AKR")
            else:
                self.fields["dagboek"]['var'].set("APF")
        self.on_field_change("dagboek")

    def save_data(self):
        if not self.current_pdf_path:
            messagebox.showwarning("Geen PDF", "Open eerst een PDF om te verwerken.")
            return
        self.status_label.config(text="Status: Opslaan en model bijwerken...")
        self.master.update_idletasks()
        corrected_data = {}
        for fn in self.fields.keys():
            corrected_data[fn] = self.fields[fn]['var'].get()
        final_path, chosen_submap = move_processed_file(self.current_pdf_path, PROCESSED_PDF_DIR, corrected_data.get("geadresseerde", "").strip())
        if not final_path:
            return
        append_correction_to_csv(final_path, corrected_data, CORRECTED_DATA_FILE)
        append_label_annotations(final_path, corrected_data, ANNOTATION_FILE)
        append_invoice_to_excel(final_path, corrected_data, chosen_submap, EXTRA_DATA_EXCEL)
        supplier_value = corrected_data.get("leverancier", "").strip()
        if supplier_value:
            if self.pdf_text:
                for field in ['geadresseerde','factuur_nummer','factuur_datum','projid','inkooporder_nummer','bedrag']:
                    field_value = corrected_data.get(field, "").strip()
                    if field_value:
                        update_learned_offset(supplier_value, field, self.pdf_text, field_value)
            suppliers = load_known_suppliers()
            if supplier_value not in suppliers:
                nummer = open_supplier_number_popup()
                if nummer is None:
                    nummer = "L00001"
                suppliers[supplier_value] = nummer
                save_known_suppliers(suppliers)
                print(f"Leverancier '{supplier_value}' toegevoegd aan known_suppliers.json")
        init_training_db()  # Initialiseer de DB als die nog niet bestaat.
        for field in FIELD_MAPPING.keys():
            value = corrected_data.get(field, "")
            store_training_correction(final_path, self.pdf_text, field, value, value, "geen correctie")
        self.train_batch_buffer.append((final_path, corrected_data))
        if len(self.train_batch_buffer) >= 20:
            if self.nlp_models:
                incremental_update_models_from_batch(self.train_batch_buffer, self.nlp_models, n_iter=5, drop=0.1)
            else:
                print("Geen spaCy-modellen geladen; kan niet online batch leren.")
            self.train_batch_buffer.clear()
        self.status_label.config(text="Status: Opslaan gereed. Volgende factuur wordt geladen...")
        self.master.update_idletasks()
        self.clear_pdf_display()
        self.load_next_invoice()

    def skip_file(self):
        if not self.current_pdf_path:
            messagebox.showwarning("Geen PDF", "Er is geen PDF om over te slaan.")
            return
        if not INVOICE_FOLDER:
            messagebox.showerror("Configuratiefout", "INVOICE_FOLDER is niet geconfigureerd.")
            return
        skip_folder = os.path.join(INVOICE_FOLDER, "Overgeslagen")
        if not os.path.exists(skip_folder):
            os.makedirs(skip_folder, exist_ok=True)
        fname = os.path.basename(self.current_pdf_path)
        try:
            shutil.move(self.current_pdf_path, os.path.join(skip_folder, fname))
            messagebox.showinfo("Bestand overgeslagen", f"PDF is overgeslagen en verplaatst naar: {skip_folder}")
        except Exception as e:
            messagebox.showerror("Fout", f"Fout bij het verplaatsen van de PDF: {e}")
            return
        self.status_label.config(text="Status: Bestand overgeslagen, volgende factuur wordt geladen...")
        self.master.update_idletasks()
        self.clear_pdf_display()
        self.load_next_invoice()

    def load_next_invoice(self):
        if self.pdf_list and self.current_index < len(self.pdf_list) - 1:
            self.current_index += 1
            next_pdf = self.pdf_list[self.current_index]
            self.load_invoice(next_pdf)
        else:
            messagebox.showinfo("Einde", "Er zijn geen verdere PDF's in de map.")
            self.status_label.config(text="Status: Alle documenten verwerkt.")
            self.pdf_counter_label.config(text=f"Bestand {len(self.pdf_list)}/{len(self.pdf_list)}")

    def clear_pdf_display(self):
        self.current_pdf_path = None
        self.pdf_text = ""
        for fn in self.fields.keys():
            self.fields[fn]['var'].set("")
            self.fields[fn]['indicator'].configure(text="")
            self.field_sources[fn] = ""
        self.pdf_text_widget.delete("1.0", tk.END)
        try:
            if self.pdf_preview_container.winfo_exists():
                for widget in self.pdf_preview_container.winfo_children():
                    widget.destroy()
        except Exception as e:
            print(f"Fout bij het leegmaken van de PDF-preview container: {e}")

    def train_model_action(self):
        excel_path = filedialog.askopenfilename(filetypes=[("Excel-bestanden", "*.xlsx *.xls")])
        if not excel_path:
            return
        messagebox.showinfo("Informatie", "Gebruik 'Model hertrain (met correcties)' voor veldspecifieke training.")
        
    def retrain_with_corrections(self):
        df = load_training_corrections_df()
        if df.empty:
            messagebox.showwarning("Geen correcties", "Er zijn geen correcties in de trainingsdatabase.")
            return
        for field in FIELD_MAPPING.keys():
            output_dir = get_model_dir_for_field(field)
            try:
                train_spacy_model_field_from_db(field, output_dir, BASE_SPACY_MODEL, n_iter=30)
            except Exception as e:
                messagebox.showerror("Hertrain fout", f"Fout bij hertraining voor veld {field}: {e}")
        messagebox.showinfo("Hertrain geslaagd", "Alle veldspecifieke modellen zijn hergetraind met correcties.")
        self.nlp_models = load_all_field_models()

    def open_mapping_editor(self):
        editor_win = tk.Toplevel(self.master)
        editor_win.title("ProjID-koppelingeditor")
        editor_win.geometry("500x400")
        listbox = tk.Listbox(editor_win, width=60)
        listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        def refresh_listbox():
            listbox.delete(0, tk.END)
            for desc, projid in load_projid_mappings().items():
                listbox.insert(tk.END, f"{desc}  -->  {projid}")
        refresh_listbox()
        frm_entries = ttk.Frame(editor_win)
        frm_entries.pack(padx=10, pady=5, fill=tk.X)
        ttk.Label(frm_entries, text="Omschrijving:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        desc_var = tk.StringVar()
        ent_desc = ttk.Entry(frm_entries, textvariable=desc_var, width=40)
        ent_desc.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frm_entries, text="Koppelde PROJID:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        projid_var = tk.StringVar()
        ent_projid = ttk.Entry(frm_entries, textvariable=projid_var, width=40)
        ent_projid.grid(row=1, column=1, padx=5, pady=5)
        def add_mapping():
            d = desc_var.get().strip()
            p = projid_var.get().strip()
            if d and p:
                current = load_projid_mappings()
                current[d] = p
                save_projid_mappings(current)
                refresh_listbox()
                desc_var.set("")
                projid_var.set("")
        def delete_mapping():
            selection = listbox.curselection()
            if selection:
                item = listbox.get(selection[0])
                parts = item.split("  -->  ")
                if len(parts) == 2:
                    d = parts[0].strip()
                    current = load_projid_mappings()
                    if d in current:
                        del current[d]
                        save_projid_mappings(current)
                        refresh_listbox()
        def load_selected_mapping(event):
            selection = listbox.curselection()
            if selection:
                item = listbox.get(selection[0])
                parts = item.split("  -->  ")
                if len(parts) == 2:
                    desc_var.set(parts[0].strip())
                    projid_var.set(parts[1].strip())
        listbox.bind("<<ListboxSelect>>", load_selected_mapping)
        btn_frame = ttk.Frame(editor_win)
        btn_frame.pack(padx=10, pady=5)
        btn_add = ttk.Button(btn_frame, text="Toevoegen / Bijwerken", command=add_mapping)
        btn_add.grid(row=0, column=0, padx=5)
        btn_delete = ttk.Button(btn_frame, text="Verwijderen", command=delete_mapping)
        btn_delete.grid(row=0, column=1, padx=5)

    def open_trainer_mode(self):
        trainer_win = tk.Toplevel(self.master)
        trainer_win.title("Trainingsmodus")
        trainer_win.geometry("950x750")
        MultiTrainerGUI(trainer_win, self.nlp_models)

    def open_booking_mode(self):
        booking_win = tk.Toplevel(self.master)
        booking_win.title("Boekingsmodus")
        booking_win.geometry("1200x800")
        BookingModeGUI(booking_win)

# -------------------------------------
# MultiTrainerGUI: Trainingsmodus
# -------------------------------------
class MultiTrainerGUI:
    def __init__(self, master, nlp_models):
        self.master = master
        self.master.configure(bg="#F2F2F2")
        self.nlp_models = nlp_models
        self.training_output_file = "training_labels.csv"
        self.pdf_list = []
        self.current_index = 0
        self.current_pdf_path = None
        self.pdf_text = ""
        self.candidate_values = []
        top_frame = ttk.Frame(self.master, padding="5")
        top_frame.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(top_frame, text="Selecteer veld:").pack(side=tk.LEFT, padx=5)
        self.fields_available = ['geadresseerde', 'leverancier', 'factuur_nummer', 'factuur_datum', 'projid', 'inkooporder_nummer', 'bedrag']
        self.selected_field = tk.StringVar(value=self.fields_available[0])
        self.field_dropdown = ttk.Combobox(top_frame, textvariable=self.selected_field, values=self.fields_available, state="readonly", width=20)
        self.field_dropdown.pack(side=tk.LEFT, padx=5)
        self.field_dropdown.bind("<<ComboboxSelected>>", self.reload_current_invoice)
        self.btn_select_folder = ttk.Button(top_frame, text="Selecteer trainingsmap", command=self.select_train_folder)
        self.btn_select_folder.pack(side=tk.LEFT, padx=5)
        main_frame = ttk.Frame(self.master, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        control_frame = ttk.Frame(main_frame, padding="5")
        control_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.label_field = ttk.Label(control_frame, text="Waarde:", font=("Helvetica", 14))
        self.label_field.pack(pady=5)
        self.value_var = tk.StringVar()
        self.value_display = ttk.Label(control_frame, textvariable=self.value_var, font=("Helvetica", 24, "bold"), foreground="blue")
        self.value_display.pack(pady=5)
        self.btn_good = tk.Button(control_frame, text="✔ Goed", command=self.mark_good, bg="#AAAAAA", fg="black", font=("Helvetica", 20, "bold"), width=10)
        self.btn_good.pack(pady=10, padx=5)
        self.btn_bad = tk.Button(control_frame, text="✖ Fout", command=self.mark_bad, bg="#CC5555", fg="white", font=("Helvetica", 20, "bold"), width=10)
        self.btn_bad.pack(pady=10, padx=5)
        self.btn_none = tk.Button(control_frame, text="Geen waarde", command=self.mark_none, bg="#8888CC", fg="white", font=("Helvetica", 20, "bold"), width=10)
        self.btn_none.pack(pady=10, padx=5)
        preview_frame = ttk.LabelFrame(main_frame, text="PDF-preview", padding="5")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.preview_container = ttk.Frame(preview_frame)
        self.preview_container.pack(fill=tk.BOTH, expand=True)
        self.btn_next = ttk.Button(self.master, text="Volgende", command=self.go_to_next, state="disabled")
        self.btn_next.pack(side=tk.BOTTOM, pady=5)
        self.status_label = ttk.Label(self.master, text="Status: Kies een map.")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        self.candidate_functions = {
            'factuur_datum': get_candidate_factuurdatums,
            'factuur_nummer': get_candidate_factuurnummers,
            'bedrag': get_candidate_bedragen,
            'leverancier': get_candidate_leveranciers,
            'projid': get_candidate_projid,
            'inkooporder_nummer': get_candidate_inkoopordernummers,
            'geadresseerde': get_candidate_geadresseerde
        }
        init_training_db()

    def select_train_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.pdf_list = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        self.pdf_list.sort()
        if not self.pdf_list:
            messagebox.showwarning("Leeg", "Geen PDF's gevonden in deze map.")
            return
        self.current_index = 0
        self.load_invoice(self.pdf_list[self.current_index])
        self.btn_next.config(state="normal")
        self.status_label.config(text=f"Map geselecteerd: {folder} ({len(self.pdf_list)} PDF's)")

    def load_invoice(self, pdf_path):
        self.current_pdf_path = pdf_path
        self.status_label.config(text=f"Bezig met: {os.path.basename(pdf_path)}")
        try:
            self.pdf_text = extract_text(pdf_path)
        except Exception as e:
            messagebox.showerror("Extractiefout", f"Fout bij extractie: {e}")
            self.pdf_text = ""
        result = extract_invoice_data(pdf_path, nlp_models=self.nlp_models, use_regex_heuristieken=True)
        field = self.selected_field.get()
        value = result.get(field, {}).get("value", "")
        self.value_var.set(value)
        candidate_func = self.candidate_functions.get(field, lambda x: [])
        self.candidate_values = candidate_func(self.pdf_text)
        show_pdf_scrollable_preview(pdf_path, self.preview_container, dpi=100)

    def reload_current_invoice(self, event=None):
        if self.current_pdf_path:
            self.load_invoice(self.current_pdf_path)

    def mark_good(self):
        corrected_value = self.value_var.get().strip()
        if corrected_value == "":
            return
        field = self.selected_field.get()
        store_training_correction(self.current_pdf_path, self.pdf_text, field, corrected_value, corrected_value, "goed")
        model_for_field = self.nlp_models.get(field)
        if model_for_field:
            incremental_update_model_from_correction(self.current_pdf_path, {field: corrected_value}, model_for_field, field, n_iter=5, drop=0.1)
        self.go_to_next()

    def mark_bad(self):
        field = self.selected_field.get()
        original_value = self.value_var.get().strip()
        alternatives = [c for c in self.candidate_values if c != original_value]
        if alternatives:
            new_value = alternatives[0]
            self.value_var.set(new_value)
            store_training_correction(self.current_pdf_path, self.pdf_text, field, original_value, new_value, "fout")
            model_for_field = self.nlp_models.get(field)
            if model_for_field:
                incremental_update_model_from_correction(self.current_pdf_path, {field: new_value}, model_for_field, field, n_iter=5, drop=0.1)
        self.go_to_next()

    def mark_none(self):
        field = self.selected_field.get()
        original_value = self.value_var.get().strip()
        corrected_value = ""
        store_training_correction(self.current_pdf_path, self.pdf_text, field, original_value, corrected_value, "geen waarde")
        model_for_field = self.nlp_models.get(field)
        if model_for_field:
            incremental_update_model_from_correction(self.current_pdf_path, {field: corrected_value}, model_for_field, field, n_iter=5, drop=0.1)
        self.go_to_next()

    def go_to_next(self):
        self.current_index += 1
        if self.current_index < len(self.pdf_list):
            self.load_invoice(self.pdf_list[self.current_index])
        else:
            self.btn_next.config(state="disabled")
            self.status_label.config(text="Geen verdere PDF's.")

# -------------------------------------
# Boekingsmodus
# -------------------------------------
class BookingModeGUI:
    def __init__(self, master):
        self.master = master
        self.master.configure(bg="#F2F2F2")
        self.invoice_data = None
        self.all_invoices_df = self.load_unbooked_invoices()
        self.invoices_df = self.all_invoices_df.copy()
        self.current_index = 0
        main_frame = ttk.Frame(self.master)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.left_frame = ttk.Frame(main_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.right_frame = ttk.Frame(main_frame, padding=5)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.btn_combi_filter = ttk.Button(self.right_frame, text="Combi-filter", command=self.open_combifilter)
        self.btn_combi_filter.pack(side=tk.TOP, padx=5, pady=5)
        self.details_frame = ttk.Frame(self.right_frame)
        self.details_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.details_text = tk.Text(self.details_frame, height=14, wrap="word", width=40)
        self.details_text.pack(side=tk.TOP, fill=tk.X, expand=False)
        ttk.Label(self.details_frame, text="Dagboek wijzigen:").pack(side=tk.TOP, pady=5)
        self.dagboek_booking_var = tk.StringVar()
        self.dagboek_booking_combo = ttk.Combobox(self.details_frame, textvariable=self.dagboek_booking_var, values=["AKR", "AKC", "APF", "APC"], state="readonly", width=8)
        self.dagboek_booking_combo.pack(side=tk.TOP, pady=5)
        self.btn_frame = ttk.Frame(self.right_frame, padding=5)
        self.btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.btn_skip = ttk.Button(self.btn_frame, text="Volgende overslaan", command=self.mark_skip)
        self.btn_skip.pack(side=tk.LEFT, padx=5, pady=5)
        self.btn_booked = ttk.Button(self.btn_frame, text="Volgende geboekt", command=self.mark_booked)
        self.btn_booked.pack(side=tk.LEFT, padx=5, pady=5)
        self.pdf_preview_container = ttk.Frame(self.left_frame)
        self.pdf_preview_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        zoom_frame = ttk.Frame(self.left_frame, padding=5)
        zoom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.btn_zoom_in = ttk.Button(zoom_frame, text="Zoom in", command=self.zoom_in)
        self.btn_zoom_in.pack(side=tk.LEFT, padx=2)
        self.btn_zoom_out = ttk.Button(zoom_frame, text="Zoom uit", command=self.zoom_out)
        self.btn_zoom_out.pack(side=tk.LEFT, padx=2)
        self.btn_zoom_reset = ttk.Button(zoom_frame, text="Zoom resetten", command=self.reset_zoom)
        self.btn_zoom_reset.pack(side=tk.LEFT, padx=2)
        self.default_dpi = 100
        self.zoom_factor = 0.8
        self.status_label = ttk.Label(self.master, text="Boekingsmodus gestart.", background="#F2F2F2")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        if self.invoices_df.empty:
            messagebox.showinfo("Boekingsmodus", "Er zijn geen ongeboekte facturen.")
            self.master.destroy()
            return
        self.load_current_invoice()

    def load_unbooked_invoices(self):
        if os.path.exists(EXTRA_DATA_EXCEL):
            try:
                df = pd.read_excel(EXTRA_DATA_EXCEL)
                if 'booked' in df.columns:
                    df = df[df['booked'] != True]
                return df.reset_index(drop=True)
            except Exception as e:
                messagebox.showerror("Boekingsmodus", f"Fout bij laden van Excel: {e}")
                return pd.DataFrame()
        else:
            return pd.DataFrame()

    def open_combifilter(self):
        combi_win = tk.Toplevel(self.master)
        combi_win.title("Submap+Dagboek-filter")
        combi_win.geometry("400x150")
        lbl = ttk.Label(combi_win, text="Kies combinatie (submap + dagboek):")
        lbl.pack(pady=10)
        group_df = (self.all_invoices_df[self.all_invoices_df['booked'] != True].groupby(['submap','dagboek']).size().reset_index(name='count'))
        combos = ["(ALLE)"]
        combos_map = {"(ALLE)": (None, None)}
        for idx, row in group_df.iterrows():
            submap_val = row['submap'] if pd.notna(row['submap']) else ""
            dagboek_val = row['dagboek'] if pd.notna(row['dagboek']) else ""
            c = row['count']
            combo_label = f"{dagboek_val or '(geen)'} + {submap_val or '(geen)'} ({c})"
            combos.append(combo_label)
            combos_map[combo_label] = (submap_val, dagboek_val)
        chosen_var = tk.StringVar(value="(ALLE)")
        combo = ttk.Combobox(combi_win, textvariable=chosen_var, values=combos, state="readonly", width=35)
        combo.pack(pady=5)
        def on_ok():
            selection = chosen_var.get()
            if selection in combos_map:
                sm, db = combos_map[selection]
                self.apply_combination_filter(sm, db)
            combi_win.destroy()
        def on_cancel():
            combi_win.destroy()
        btn_frm = ttk.Frame(combi_win)
        btn_frm.pack(pady=5)
        btn_ok = ttk.Button(combi_win, text="OK", command=on_ok)
        btn_ok.pack(side=tk.LEFT, padx=5)
        btn_cancel = ttk.Button(combi_win, text="Annuleren", command=on_cancel)
        btn_cancel.pack(side=tk.LEFT, padx=5)

    def apply_combination_filter(self, submap_filter, dagboek_filter):
        df = self.all_invoices_df.copy()
        df = df[df['booked'] != True]
        if submap_filter is not None and dagboek_filter is not None:
            df = df[(df['submap'].fillna('') == submap_filter.strip()) & (df['dagboek'].fillna('') == dagboek_filter.strip())]
        self.invoices_df = df.reset_index(drop=True)
        self.current_index = 0
        if self.invoices_df.empty:
            self.clear_preview()
            self.details_text.delete("1.0", tk.END)
            self.dagboek_booking_var.set("")
            messagebox.showinfo("Filter", "Geen facturen gevonden voor deze combinatie.")
        else:
            self.load_current_invoice()

    def clear_preview(self):
        for widget in self.pdf_preview_container.winfo_children():
            widget.destroy()

    def load_current_invoice(self):
        if self.current_index >= len(self.invoices_df):
            messagebox.showinfo("Boekingsmodus", "Geen verdere ongeboekte facturen.")
            self.master.destroy()
            return
        self.invoice_data = self.invoices_df.iloc[self.current_index]
        self.details_text.delete("1.0", tk.END)
        new_filename_raw = self.invoice_data.get('new_filename', '')
        filename_no_ext = os.path.splitext(str(new_filename_raw))[0]
        details = []
        details.append(f"Nieuwe bestandsnaam: {filename_no_ext}")
        details.append(f"Geadresseerde: {self.invoice_data.get('geadresseerde', '')}")
        suppliers = load_known_suppliers()
        lev = self.invoice_data.get('leverancier', '')
        if lev in suppliers:
            lev_display = suppliers[lev]
        else:
            lev_display = lev
        details.append(f"Leverancier: {lev_display}")
        details.append(f"Factuurnummer: {self.invoice_data.get('factuur_nummer', '')}")
        details.append(f"Factuurdatum: {self.invoice_data.get('factuur_datum', '')}")
        details.append(f"PROJID: {self.invoice_data.get('projid', '')}")
        details.append(f"Inkooporder nummer: {self.invoice_data.get('inkooporder_nummer', '')}")
        details.append(f"Bedrag: {self.invoice_data.get('bedrag', '')}")
        details.append(f"Submap: {self.invoice_data.get('submap', '')}")
        existing_dagboek = str(self.invoice_data.get('dagboek', '')).strip()
        details.append(f"Huidige dagboek: {existing_dagboek}")
        self.details_text.insert(tk.END, "\n".join(details))
        if existing_dagboek in ["AKR","AKC","APF","APC"]:
            self.dagboek_booking_var.set(existing_dagboek)
        else:
            self.dagboek_booking_var.set("AKR")
        pdf_path = self.invoice_data.get('pdf_path', '')
        if pdf_path and os.path.exists(pdf_path):
            self.clear_preview()
            show_pdf_scrollable_preview(pdf_path, self.pdf_preview_container, dpi=int(self.default_dpi*self.zoom_factor))
        else:
            self.clear_preview()
            lbl = ttk.Label(self.pdf_preview_container, text="PDF niet gevonden.")
            lbl.pack()

    def update_booking_status(self, booking_type):
        new_pdf_path = None
        if os.path.exists(EXTRA_DATA_EXCEL):
            try:
                df = pd.read_excel(EXTRA_DATA_EXCEL)
                mask = ((df['pdf_path'] == self.invoice_data['pdf_path']) & (df['new_filename'] == self.invoice_data['new_filename']))
                df.loc[mask, 'booked'] = True
                df.loc[mask, 'booking_type'] = booking_type
                chosen_dagboek = self.dagboek_booking_var.get().strip()
                if chosen_dagboek in ["AKR","AKC","APF","APC"]:
                    df.loc[mask, 'dagboek'] = chosen_dagboek
                new_pdf_path, chosen_submap = move_processed_file(self.invoice_data['pdf_path'], PROCESSED_PDF_DIR, self.invoice_data.get("geadresseerde", ""))
                if new_pdf_path:
                    df.loc[mask, 'pdf_path'] = new_pdf_path
                    df.loc[mask, 'new_filename'] = os.path.basename(new_pdf_path)
                df.to_excel(EXTRA_DATA_EXCEL, index=False)
            except Exception as e:
                messagebox.showerror("Boekingsmodus", f"Fout bij updaten van Excel: {e}")
        effective_pdf_path = new_pdf_path if new_pdf_path and os.path.exists(new_pdf_path) else self.invoice_data.get('pdf_path', '')
        if effective_pdf_path and os.path.exists(effective_pdf_path):
            pdf_text = extract_text(effective_pdf_path)
            for field in ['geadresseerde','leverancier','factuur_nummer','factuur_datum','projid','inkooporder_nummer','bedrag','dagboek']:
                value = self.invoice_data.get(field, "")
                store_training_correction(effective_pdf_path, pdf_text, field, "", value, "boeking")

    def mark_skip(self):
        self.update_booking_status("overslaan")
        self.next_invoice()

    def mark_booked(self):
        self.update_booking_status("geboekt")
        self.next_invoice()

    def next_invoice(self):
        self.current_index += 1
        if self.current_index < len(self.invoices_df):
            self.load_current_invoice()
        else:
            messagebox.showinfo("Boekingsmodus", "Geen verdere ongeboekte facturen.")
            self.master.destroy()

    def zoom_in(self):
        self.zoom_factor += 0.1
        self.reload_pdf_preview()

    def zoom_out(self):
        if self.zoom_factor > 0.2:
            self.zoom_factor -= 0.1
            self.reload_pdf_preview()

    def reset_zoom(self):
        self.zoom_factor = 0.8
        self.reload_pdf_preview()

    def reload_pdf_preview(self):
        if self.invoice_data is not None and not self.invoice_data.empty:
            pdf_path = self.invoice_data.get('pdf_path', '')
            if pdf_path and os.path.exists(pdf_path):
                self.clear_preview()
                show_pdf_scrollable_preview(pdf_path, self.pdf_preview_container,
                                            dpi=int(self.default_dpi*self.zoom_factor))

def main():
    root = tk.Tk()
    app = InvoiceGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

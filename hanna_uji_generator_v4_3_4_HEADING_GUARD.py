#!/usr/bin/env python3
"""
HANNA -> UJI PRODUCT IMPORTER V4.3.4 HEADING-GUARD

Prinsip utama:
1. Membaca data produk Hanna dari XLSX dan refresh URL produk secara default.
2. Memisahkan body Description dari dedicated source tab/panel berdasarkan provenance DOM; inline heading di Description tidak boleh membuat custom tab.
3. Menerjemahkan narrative evidence ke bahasa Indonesia sebelum product writer dengan proteksi angka, model, unit, standar, port, dan istilah teknis.
4. Menggunakan registry kategori UJI sebagai kategori canonical. Taxonomy Hanna tidak pernah menjadi fallback mentah.
5. Mengizinkan kategori baru yang valid, tetapi menandainya REVIEW_NEW_CATEGORY sampai disetujui.
6. Membuat nama produk ringkas dan meta description sebagai kalimat SEO utuh.
7. Membuat custom tab hanya dari dedicated source tab/panel yang benar-benar ada; Features/Benefits inline di Description tetap berada di Product Description.
8. FAQ bersifat opsional dan hanya dibuat bila evidence cukup.
9. Source ringkas menghasilkan output ringkas. Tidak ada minimum kata, FAQ, fitur, atau custom tab yang dipaksakan.
10. Memblokir referral, CTA, dan promosi yang mengarahkan customer ke hannainst.com.
11. Specification lock tetap dipertahankan agar data spesifikasi XLSX tidak ditimpa refresh website.
12. Merapikan Product Description secara deterministik setelah validasi: paragraf, heading inline, dan feature blocks ditata tanpa menambah atau mengubah fakta.
13. Heading inline hanya dikenali pada posisi heading yang eksplisit; kata seperti "aplikasi" di tengah kalimat tidak boleh memecah kalimat. Dangling fragment otomatis membatalkan pemisahan.
14. Product Description tidak lagi dipotong dengan hard word-limit pada fallback/formatter; panjang mengikuti overview yang sudah lolos validasi.
15. Menghasilkan XLSX 24 kolom sesuai template impor UJI beserta lembar REVIEW.

Kontak penjualan yang ditampilkan:
WhatsApp: +62 857-1711-2222
Email: quotations@jvm.co.id
"""

from __future__ import annotations

import argparse
from collections import Counter
from difflib import SequenceMatcher
import getpass
import hashlib
import html
import json
import logging
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urljoin, urlparse

try:
    import requests
    from bs4 import BeautifulSoup, Comment
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    print("Dependensi belum tersedia.")
    print("Jalankan: python -m pip install -r requirements_hanna_uji.txt")
    raise SystemExit(1) from exc


VERSION = "4.3.4-heading-guard"
DEFAULT_INPUT = "hannainst.com.xlsx"
DEFAULT_OUTPUT = "output/uji_import_hanna_v4_3_4_heading_guard.xlsx"
DEFAULT_SHEET = "25-11-2025"

IMPORT_COLUMNS = [
    "Name",
    "SKU",
    "Brand",
    "Slug",
    "SEO Title",
    "Product Description",
    "Product Short Description",
    "custom_tab_1_title",
    "custom_tab_1_content",
    "custom_tab_1_priority",
    "custom_tab_2_title",
    "custom_tab_2_content",
    "custom_tab_2_priority",
    "custom_tab_3_title",
    "custom_tab_3_content",
    "custom_tab_3_priority",
    "product categories",
    "product tags",
    "focus keyphrase",
    "meta description",
    "publication_date",
    "image_url",
    "Source URL",
    "SEO QA Status",
]

SOURCE_COLUMNS = [
    "url",
    "brand",
    "sku",
    "title",
    "taxonomy",
    "short_description",
    "description",
    "images",
    "specification",
]

DEFAULT_CONFIG = {
    "brand_output": "Hanna Instruments",
    "display_brand": "Hanna",
    "sku_prefix": "HANNA-",
    "whatsapp_display": "+62 857-1711-2222",
    "whatsapp_number": "6285717112222",
    "email": "quotations@jvm.co.id",
    "company_name": "UJI",
    "include_source_url_in_import": False,
    "download_images": False,
    "image_directory": "output/images",
    "request_delay_seconds": 1.2,
    "request_timeout_seconds": 30,
    "max_retries": 3,
    "user_agent": "Mozilla/5.0 (compatible; UJI-Product-Importer/4.1)",
    "use_deepseek_faq": True,
    "deepseek_base_url": "https://api.deepseek.com",
    "deepseek_model": "deepseek-v4-flash",
    "deepseek_api_key": "",
    "deepseek_timeout_seconds": 90,
    "deepseek_max_retries": 2,
    "deepseek_delay_seconds": 0.35,
    "deepseek_temperature": 0.1,
    "deepseek_max_tokens": 1000,
    "deepseek_cache_file": "output/deepseek_faq_polish_cache_v4.json",
    "deepseek_prompt_version": "uji-faq-polish-v4",
    "faq_count": 8,
    "use_deepseek_content": True,
    "deepseek_content_cache_file": "output/deepseek_product_content_cache_v4_3_3.json",
    "deepseek_content_prompt_version": "uji-product-provenance-tabs-v4.3.3",
    "deepseek_content_max_tokens": 3200,
    "deepseek_content_temperature": 0.1,
    "use_deepseek_translation": True,
    "deepseek_translation_cache_file": "output/deepseek_translation_cache_v4_3.json",
    "deepseek_translation_prompt_version": "uji-evidence-translation-v4.3",
    "deepseek_translation_max_tokens": 6000,
    "deepseek_translation_temperature": 0.0,
    "max_source_content_chars": 30000,
    "min_source_content_chars": 500,
    "min_description_words": 90,
}

CANONICAL_CATEGORIES: Tuple[str, ...] = (
    "Accelerometer", "Acidity Meter", "Adhesion Tester", "Air Purifier", "Air Quality Meter",
    "Aksesori & Suku Cadang", "Anemometer & Air Flow Meter", "ATP Hygiene Meter",
    "Automotive Diagnostic Tool", "Battery Tester", "Borescope & Inspection Camera",
    "Brightness & Whiteness Meter", "Cable Tester & Detector", "Calibrator", "Chlorine Meter",
    "Clamp Meter", "CO2 Meter", "Coating & Paint Testing", "Coating Thickness Gauge",
    "Concrete Tester", "Conductivity", "Data Logger", "Densitometer", "Dissolved Hydrogen Meter",
    "Dissolved Oxygen Meter", "Distance & Angle Meter", "Electrical Tester", "Elektrode & Probe",
    "EMF Meter", "Environmental Meter", "Fiber Optic Equipment", "Flaw Detector & NDT", "Flow Meter",
    "Food Quality Tester", "Force Gauge & Dynamometer", "Gas Detector & Analyzer",
    "Gloss & Appearance Meter", "Hardness Tester", "Haze", "High-Speed Camera",
    "Humidity Meter & Hygrometer", "Hydrometer & Specific Gravity Meter", "Impact Testing Machine",
    "Laboratory Consumables", "Laboratory Equipment", "Larutan Kalibrasi & Reagen", "Leak Detector",
    "Light, Lux & Luminance Meter", "Marine Survey Instrument", "Material Analyzer",
    "Material Testing Machine", "Metal Detector", "Microscope", "Moisture Meter", "Multimeter",
    "Multiparameter Water Quality Meter", "ORP Meter", "pH & ORP Meter", "pH Controller", "pH Meter",
    "Photometer & Colorimeter", "Polarimeter & Saccharimeter", "Power Analyzer & Meter",
    "Pressure Meter & Manometer", "Process Controller & Indicator", "Radiation Meter & Dosimeter",
    "Refractometer & Brix Meter", "Refrigerant Recovery Equipment", "Resistance & Earth Tester",
    "Retroreflectometer", "Safety Equipment", "Scale & Balance", "Software & Interface",
    "Solar Power Meter", "Sound Level Meter", "Spectrophotometer & Spectrometer",
    "Spring Testing Machine", "Stroboscope", "Surface Roughness Tester", "Tachometer",
    "TDS & Salinity Meter", "Tension Meter", "Thermal Imaging Camera", "Thermometer & Temperature Meter",
    "Torque Meter", "Torsion Testing Machine", "Transmitter & Sensor", "Turbidity Meter",
    "Ultrasonic Thickness Gauge", "Universal Testing Machine", "UV Meter", "Vacuum Pump",
    "Vibration Meter", "Viscometer", "Water Quality Test Kit", "WBGT Meter", "Weather & Rain Meter",
)

# Kategori baru diperbolehkan, tetapi tidak otomatis dianggap aman untuk publikasi.
# Produk dengan kategori ini mendapat REVIEW_NEW_CATEGORY sampai kategori tersebut
# disetujui dan dipindahkan ke CANONICAL_CATEGORIES/config registry.
VALID_NEW_CATEGORY_RULES: Sequence[Tuple[str, Sequence[str]]] = (
    ("Magnetic Stirrer", (r"\bmagnetic stirrers?\b", r"\bstirrers?\b")),
    ("Titrator & Burette", (r"\btitrators?\b", r"\btitration\b", r"\bburettes?\b", r"\bburets?\b")),
    ("Moisture Analyzer", (r"\bmoisture analy[sz]ers?\b",)),
)

SOURCE_SECTION_RULES: Sequence[Tuple[str, Sequence[str]]] = (
    ("WHATS_INCLUDED", (r"supplied", r"included", r"package contents?", r"what.s included", r"accessor")),
    ("CONNECTIVITY", (r"reporting", r"connection", r"connectivity", r"data transfer", r"usb", r"rs232", r"bluetooth", r"cloud", r"software")),
    ("INSTALLATION", (r"installation", r"install", r"mounting", r"setup")),
    ("MAINTENANCE", (r"maintenance", r"cleaning", r"storage", r"care")),
    ("COMPATIBILITY", (r"compatible", r"compatibility", r"for use with")),
    ("CALIBRATION", (r"calibr", r"standard", r"traceab", r"certificate", r"glp")),
    ("WORKING_PRINCIPLE", (r"principle", r"operation", r"working", r"technology", r"optical", r"method")),
    ("APPLICATION", (r"application", r"recommended use", r"significance", r"use cases?", r"water quality", r"wastewater")),
    ("FEATURES", (r"feature", r"at.a.glance", r"benefit", r"capabil", r"key features?")),
    ("OVERVIEW", (r"overview", r"description", r"ringkasan", r"detail produk lengkap", r"product")),
)

CATEGORY_TAXONOMY_BLOCKLIST = {
    "hanna best sellers", "best sellers", "aqua", "hydro", "p - z", "p-z", "featured products",
    "new products", "sale", "special offer", "water", "laboratory", "industrial",
}

PROMO_PATTERNS = [
    r"warranty information",
    r"subscribe\s*(?:and|&)\s*save",
    r"best sellers?",
    r"black friday(?: stock up)?",
    r"labor day deals?",
    r"buy\s*\d+\s*get\s*\d+",
    r"stock up",
    r"choose options?",
    r"add to cart",
    r"shop now",
    r"free shipping",
    r"limited time",
    r"special offer",
    r"sale price",
    r"regular price",
    r"save \d+%",
    r"delivery every",
    r"order now",
    r"visit (?:the )?(?:official )?website",
    r"visit hann(a|ah) instruments?",
    r"contact hann(a|ah) instruments?",
    r"according to hann(a|ah)",
    r"based on (?:the )?official specifications?",
    r"based on hann(a|ah)(?: instruments?)?",
    r"official hann(a|ah) specifications?",
    r"information (?:is )?sourced from hann(a|ah)",
    r"hannainst\.com",
]

PROMO_CATEGORY_TOKENS = {
    "subscribe and save",
    "best seller",
    "best sellers",
    "black friday",
    "black friday stock up",
    "labor day deal",
    "labor day deals",
    "stock up",
    "sale",
    "special offer",
    "featured products",
    "new products",
    "buy 10 get 1",
    "buy 5 get 1",
}

PARAMETER_PATTERNS: Dict[str, Sequence[str]] = {
    "pH": (r"(?<![a-z0-9])ph(?![a-z0-9])",),
    "ORP": (r"\borp\b", r"\bredox\b"),
    "EC": (r"\bec\b", r"\bconductivity\b", r"\belectrical conductivity\b"),
    "TDS": (r"\btds\b", r"\btotal dissolved solids\b"),
    "Resistivitas": (r"\bresistivity\b", r"\bresistivitas\b"),
    "Salinitas": (r"\bsalinity\b", r"\bsalinitas\b"),
    "Oksigen terlarut": (r"\bdissolved oxygen\b", r"\bdo range\b", r"\boxygen saturation\b"),
    "Suhu": (r"\btemperature\b", r"\btemp(?:erature)? range\b", r"\bsuhu\b"),
    "Kekeruhan": (r"\bturbidity\b", r"\bntu\b", r"\bfnu\b"),
    "Klorin bebas": (r"\bfree chlorine\b",),
    "Klorin total": (r"\btotal chlorine\b",),
    "Klorida": (r"\bchloride\b",),
    "Amonia": (r"\bammonia\b", r"\bammonium\b"),
    "Nitrat": (r"\bnitrate\b",),
    "Nitrit": (r"\bnitrite\b",),
    "Fosfat": (r"\bphosphate\b",),
    "Alkalinitas": (r"\balkalinity\b",),
    "Kalsium": (r"\bcalcium\b",),
    "Magnesium": (r"\bmagnesium\b",),
    "Brix": (r"\bbrix\b", r"\bsugar content\b"),
    "Kesadahan": (r"\bhardness\b",),
    "Fluorida": (r"\bfluoride\b",),
    "Besi": (r"\biron\b",),
    "Tembaga": (r"\bcopper\b",),
    "Silika": (r"\bsilica\b",),
    "Bromin": (r"\bbromine\b",),
    "Iodium": (r"\biodine\b",),
    "Tekanan": (r"\bpressure\b", r"\bpsi\b", r"\bbar\b"),
}


COMMON_TRANSLATIONS = {
    "waterproof": "tahan air",
    "portable": "portabel",
    "wireless": "nirkabel",
    "multiparameter": "multiparameter",
    "photometer": "fotometer",
    "colorimeter": "colorimeter",
    "meter": "alat ukur",
    "tester": "tester",
    "electrode": "elektrode",
    "probe": "probe",
    "solution": "larutan",
    "reagent": "reagen",
    "digital": "digital",
    "laboratory": "laboratorium",
    "general purpose": "penggunaan umum",
    "replacement": "pengganti",
    "glass body": "bodi kaca",
    "temperature": "suhu",
    "conductivity": "konduktivitas",
    "dissolved oxygen": "oksigen terlarut",
    "turbidity": "kekeruhan",
    "salinity": "salinitas",
}


SPEC_KEY_TRANSLATIONS: Sequence[Tuple[str, str]] = [
    (r"^description$", "Deskripsi"),
    (r"^range$", "Rentang"),
    (r"\brange\b", "Rentang"),
    (r"\baccuracy\b", "Akurasi"),
    (r"\bresolution\b", "Resolusi"),
    (r"\bcalibration\b", "Kalibrasi"),
    (r"\btemperature\b", "Suhu"),
    (r"\boperating environment\b", "Lingkungan operasi"),
    (r"\benvironment\b", "Lingkungan"),
    (r"\bdimensions?\b", "Dimensi"),
    (r"\bweight\b", "Berat"),
    (r"\bconnector\b", "Konektor"),
    (r"\bcable length\b", "Panjang kabel"),
    (r"\bcable\b", "Kabel"),
    (r"\bbattery life\b", "Daya tahan baterai"),
    (r"\bbattery type\b", "Jenis baterai"),
    (r"\bbattery\b", "Baterai"),
    (r"\bpower supply\b", "Catu daya"),
    (r"\brecommended use\b", "Penggunaan yang direkomendasikan"),
    (r"\bapplication\b", "Aplikasi"),
    (r"\bsensor type\b", "Jenis sensor"),
    (r"\breference\b", "Referensi"),
    (r"\bjunction\b", "Junction"),
    (r"\belectrolyte\b", "Elektrolit"),
    (r"\bwarranty\b", "Garansi"),
]

ENGLISH_MARKERS = {
    "the", "this", "with", "for", "and", "is", "are", "designed", "provides", "features",
    "ideal", "testing", "results", "accurate", "measurement", "measurements", "portable",
}
INDONESIAN_MARKERS = {
    "yang", "untuk", "dan", "dengan", "merupakan", "digunakan", "pengukuran", "produk",
    "alat", "larutan", "reagen", "elektrode", "sesuai", "kebutuhan",
}


SOURCE_NOISE_PATTERNS: Sequence[str] = (
    r"\bmsrp\b", r"\bwas:\b", r"\bnow:\b", r"\byou save\b", r"\bquantity\b",
    r"\badding to cart\b", r"\badd to cart\b", r"\bdecrease quantity\b",
    r"\bincrease quantity\b", r"\bcustomers also\b", r"\bback to top\b",
    r"\bgift certificates?\b", r"\blogin\b", r"\bregister\b", r"\bclose\s*[×x]?\b",
    r"\bfacebook\b", r"\blinkedin\b", r"\binstagram\b", r"\byoutube\b",
    r"\bprivacy policy\b", r"\bcookie policy\b", r"\bterms and conditions\b",
    r"\bretailer inquiries\b", r"\bshipping policy\b", r"\bneed help\b",
    r"\bmanual download\b", r"\bresources?\b",
)

RELEVANT_SECTION_PATTERNS: Sequence[str] = (
    r"feature", r"overview", r"description", r"benefit", r"performance",
    r"significance", r"application", r"recommended use", r"principle",
    r"operation", r"technology", r"measurement", r"calibration", r"standard",
    r"traceability", r"glp", r"data logging", r"data transfer", r"memory",
    r"display", r"supplied", r"included", r"package", r"design",
)

EXCLUDED_SECTION_PATTERNS: Sequence[str] = (
    r"^resources?$", r"^specifications?$", r"^manuals?$", r"customers also",
    r"related products?", r"reviews?", r"questions? and answers?",
)

HIGH_VALUE_SPEC_GROUPS: Sequence[Tuple[str, Sequence[str]]] = (
    ("Rentang pengukuran", (r"\brange\b",)),
    ("Akurasi", (r"\baccuracy\b",)),
    ("Resolusi", (r"\bresolution\b",)),
    ("Repeatabilitas", (r"\brepeatability\b",)),
    ("Metode pengukuran", (r"\bmethod\b", r"measurement principle", r"measuring principle")),
    ("Kalibrasi", (r"\bcalibration\b",)),
    ("Standar atau referensi", (r"\bstandard\b", r"\btraceab", r"\bcompliance\b")),
    ("Detektor atau sensor", (r"\bdetector\b", r"\bsensor\b", r"\bphotocell\b")),
    ("Sumber cahaya", (r"\blight source\b", r"\blamp\b")),
    ("Memori data", (r"\blog memory\b", r"\bmemory\b", r"\bdata logging\b", r"\brecords?\b")),
    ("Konektivitas", (r"\bserial interface\b", r"\busb\b", r"\brs232\b", r"\bbluetooth\b", r"\bwireless\b")),
    ("Layar", (r"\bdisplay\b", r"\blcd\b")),
    ("Catu daya", (r"\bpower supply\b", r"\bbattery\b")),
    ("Lingkungan operasi", (r"\benvironment\b", r"operating condition")),
    ("Dimensi", (r"\bdimensions?\b",)),
    ("Berat", (r"\bweight\b",)),
)


@dataclass
class Product:
    url: str = ""
    brand: str = ""
    sku: str = ""
    title: str = ""
    taxonomy: str = ""
    short_description: str = ""
    description: str = ""
    images: str = ""
    specification: str = ""
    # JSON map of dedicated source tabs/panels scraped from Hanna.
    # Inline headings inside Description are intentionally NOT stored here.
    source_tabs: str = ""


class ImporterError(RuntimeError):
    pass


def normalize_space(value: object) -> str:
    text = "" if value is None else str(value)
    text = html.unescape(text)
    text = text.replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def strip_html(value: object) -> str:
    if value is None:
        return ""
    soup = BeautifulSoup(str(value), "html.parser")
    for node in soup(["script", "style", "noscript", "svg"]):
        node.decompose()
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()
    return normalize_space(soup.get_text(" ", strip=True))


def remove_urls(text: str) -> str:
    text = re.sub(r"https?://\S+", " ", text, flags=re.I)
    text = re.sub(r"www\.\S+", " ", text, flags=re.I)
    return normalize_space(text)


def remove_promotional_content(value: object) -> str:
    text = strip_html(value)
    text = remove_urls(text)

    # Pecah kalimat agar kalimat promosi dapat dibuang tanpa menghapus konten teknis lain.
    sentences = re.split(r"(?<=[.!?])\s+|\s*\|\s*|\s*-->\s*", text)
    kept: List[str] = []
    for sentence in sentences:
        candidate = normalize_space(sentence)
        if not candidate:
            continue
        lower = candidate.lower()
        if any(re.search(pattern, lower, flags=re.I) for pattern in PROMO_PATTERNS):
            continue
        if re.search(r"\b(?:usd|\$)\s*\d", lower):
            continue
        kept.append(candidate)

    cleaned = " ".join(kept)
    for pattern in PROMO_PATTERNS:
        cleaned = re.sub(pattern, " ", cleaned, flags=re.I)
    return normalize_space(cleaned)


def clean_taxonomy(value: object) -> str:
    text = strip_html(value)
    parts = [normalize_space(part) for part in re.split(r"\s*>\s*|\s+/\s+", text)]
    cleaned: List[str] = []
    for part in parts:
        if not part:
            continue
        lower = part.lower()
        if lower in PROMO_CATEGORY_TOKENS:
            continue
        if any(re.fullmatch(pattern, lower, flags=re.I) for pattern in PROMO_PATTERNS):
            continue
        if part not in cleaned:
            cleaned.append(part)
    return " > ".join(cleaned)


def slugify(value: str) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    return value.strip("-")[:180]


def clean_sku(value: object) -> str:
    sku = normalize_space(value).upper()
    sku = re.sub(r"[^A-Z0-9._/-]+", "-", sku)
    return sku.strip("-")


def output_sku(source_sku: str, prefix: str) -> str:
    if not source_sku:
        return ""
    if source_sku.upper().startswith(prefix.upper()):
        return source_sku
    return f"{prefix}{source_sku}"


def clean_title(value: object, sku: str) -> str:
    title = remove_promotional_content(value)
    title = re.sub(r"\s*[-|]\s*Hanna Instruments\s*$", "", title, flags=re.I)
    # Hilangkan nama brand di awal judul sumber agar Name tidak menjadi
    # "Hanna Hanna ..." atau "Hanna Instruments Hanna ...".
    title = re.sub(r"^\s*Hanna(?:\s+Instruments)?\s*[-|:]?\s*", "", title, flags=re.I)
    title = re.sub(r"\bHanna Instruments\b", "", title, flags=re.I)
    title = normalize_space(title)
    if sku:
        sku_candidates = [sku]
        base_sku = re.split(r"[-/]", sku, maxsplit=1)[0]
        if base_sku and base_sku.lower() != sku.lower():
            sku_candidates.append(base_sku)
        for candidate in sku_candidates:
            if re.match(rf"^{re.escape(candidate)}(?:\s|[-|:])", title, flags=re.I):
                title = normalize_space(title[len(candidate):].lstrip(" -|:"))
                break
    return title


def build_name(title: str, sku: str, brand: str, category: str = "") -> str:
    """Buat nama produk singkat: brand + model + core product title.

    Range, isi paket, aksesori bundle, dan klausa spesifikasi dipindahkan ke
    description/specification agar nama produk tidak menjadi kalimat panjang.
    """
    core = normalize_space(title)
    core = re.split(r"\s*[;|]\s*", core, maxsplit=1)[0]
    core = re.sub(r"\s*[,;]\s*(?:range|accuracy|resolution|measurement range)\b.*$", "", core, flags=re.I)
    core = re.sub(
        r"\s+with\s+(?:an?\s+)?(?:HI[A-Z0-9./_-]+\s+)?(?:electrode|probe|sensor|cable|case|solution|reagent)\b.*$",
        "", core, flags=re.I,
    )
    core = re.sub(r"\s+including\s+.*$", "", core, flags=re.I)
    # Hapus parenthetical di ujung jika jelas berisi range, ukuran, volume, jumlah paket, atau bundle.
    core = re.sub(
        r"\s*\((?=[^)]*(?:\d|mL|L\b|ppm|mg/L|µS|μS|mS|cm|mm|sachets?|pack|range|with\b|w/))[^)]*\)\s*$",
        "", core, flags=re.I,
    )
    core = normalize_space(core).strip(" -,:;/")

    # Hindari mengulang SKU jika title sumber sudah mengandung model.
    if sku:
        core = re.sub(rf"^\s*{re.escape(sku)}\s*[-|:]?\s*", "", core, flags=re.I)
    if len(core) > 82:
        core = smart_truncate(core, 82)

    chunks = [brand]
    if sku:
        chunks.append(sku)
    if core:
        chunks.append(core)
    elif category:
        chunks.append(category)
    return normalize_space(" ".join(chunks))[:110]

def normalize_category_label(value: str) -> str:
    value = validation_text(value)
    value = value.replace(" dan ", " & ")
    value = re.sub(r"[^a-z0-9&]+", " ", value)
    return normalize_space(value)


def canonicalize_category(candidate: str, threshold: float = 0.88) -> Optional[str]:
    candidate = normalize_space(candidate)
    if not candidate:
        return None
    normalized = normalize_category_label(candidate)
    exact = {normalize_category_label(item): item for item in CANONICAL_CATEGORIES}
    if normalized in exact:
        return exact[normalized]

    best: Optional[str] = None
    best_score = 0.0
    candidate_tokens = set(normalized.split())
    for item in CANONICAL_CATEGORIES:
        item_norm = normalize_category_label(item)
        score = SequenceMatcher(None, normalized, item_norm).ratio()
        item_tokens = set(item_norm.split())
        if candidate_tokens and item_tokens:
            overlap = len(candidate_tokens & item_tokens) / max(1, min(len(candidate_tokens), len(item_tokens)))
            score = max(score, overlap * 0.96)
        if score > best_score:
            best_score = score
            best = item
    return best if best is not None and best_score >= threshold else None


def infer_category_with_status(title: str, taxonomy: str, description: str) -> Tuple[str, Optional[str]]:
    """Resolve kategori UJI tanpa memakai taxonomy Hanna sebagai fallback mentah."""
    title_text = normalize_space(title).lower()
    context = normalize_space(f"{title} {description}").lower()

    # 1) Identitas alat utama harus menang dari accessory yang hanya muncul setelah 'with'.
    is_meter_like = bool(re.search(r"\b(?:meter|tester|controller|analy[sz]er|photometer|colorimeter|spectrophotometer|refractometer|turbidimeter|thermometer)\b", title_text))
    has_ph = bool(re.search(r"(?<![a-z0-9])ph(?![a-z0-9])", title_text))
    has_orp = bool(re.search(r"\borp\b|\bredox\b", title_text))
    has_ec = bool(re.search(r"\bconductivity\b|(?<![a-z0-9])ec(?![a-z0-9])", title_text))
    has_tds = bool(re.search(r"\btds\b", title_text))
    has_salinity = bool(re.search(r"\bsalinity\b|\bsalinit", title_text))
    has_do = bool(re.search(r"\bdissolved oxygen\b|(?<![a-z0-9])do(?![a-z0-9])", title_text))
    measured_groups = sum((has_ph or has_orp, has_ec or has_tds or has_salinity, has_do))

    # Tentukan noun produk yang muncul lebih dulu pada title. Ini mencegah
    # "meter with electrode" menjadi Elektrode & Probe, sekaligus mencegah
    # "glass cuvettes for colorimeters" menjadi Photometer & Colorimeter.
    instrument_identity = re.search(
        r"\b(?:meter|tester|controller|analy[sz]er|photometer|colorimeter|spectrophotometer|spectrometer|"
        r"refractometer|turbidimeter|thermometer|titrator|burette|stirrer|grader|simulator|checker)\b",
        title_text, re.I,
    )
    support_patterns: Sequence[Tuple[str, str]] = (
        ("Larutan Kalibrasi & Reagen", r"\b(?:solutions?|reagents?|buffers?|standards?|electrolyte(?:\s+fill)?\s+solutions?|stabilizer\s+powder\s+packets?|powder\s+packets?)\b"),
        ("Elektrode & Probe", r"\b(?:electrodes?|probes?|sensors?)\b"),
        ("Laboratory Consumables", r"\b(?:cuvettes?|bottles?|caps?|tips?|tweezers?|syringes?|pipettes?|beakers?|vials?)\b"),
        ("Aksesori & Suku Cadang", r"\b(?:accessor(?:y|ies)|cables?|holders?|adapters?|cases?|connectors?|tubing|fittings?)\b"),
    )
    earliest_support: Optional[Tuple[int, str]] = None
    for support_category, pattern in support_patterns:
        match = re.search(pattern, title_text, re.I)
        if match and (earliest_support is None or match.start() < earliest_support[0]):
            earliest_support = (match.start(), support_category)
    if earliest_support and (instrument_identity is None or earliest_support[0] < instrument_identity.start()):
        return earliest_support[1], None

    if re.search(r"\bpH\s*controller\b|\bcontroller\b.*\bpH\b", title, re.I):
        return "pH Controller", None
    if is_meter_like and ("multiparameter" in title_text or measured_groups >= 2):
        return "Multiparameter Water Quality Meter", None
    if is_meter_like and has_ph and has_orp:
        return "pH & ORP Meter", None
    if is_meter_like and has_do:
        return "Dissolved Oxygen Meter", None
    if is_meter_like and has_orp:
        return "ORP Meter", None
    if is_meter_like and has_ph:
        return "pH Meter", None
    if is_meter_like and (has_tds or has_salinity) and not has_ec:
        return "TDS & Salinity Meter", None
    if is_meter_like and has_ec:
        return "Conductivity", None

    direct_rules: Sequence[Tuple[str, Sequence[str]]] = (
        ("Calibrator", (r"\bcalibrators?\b", r"\bsimulators?\b")),
        ("Photometer & Colorimeter", (r"\bphotometers?\b", r"\bcolorimeters?\b", r"\bchecker(?:®)?(?:\s+[a-z0-9-]+){0,2}\s+hc\b", r"\bchecker(?:®)?\s+hc\b")),
        ("Spectrophotometer & Spectrometer", (r"\bspectrophotometers?\b", r"\bspectrometers?\b")),
        ("Refractometer & Brix Meter", (r"\brefractometers?\b", r"\bbrix meters?\b")),
        ("Turbidity Meter", (r"\bturbidity meters?\b", r"\bturbidimeters?\b")),
        ("Thermometer & Temperature Meter", (r"\bthermometers?\b", r"\btemperature meters?\b")),
        ("Chlorine Meter", (r"\bchlorine meters?\b",)),
        ("Data Logger", (r"\bdata loggers?\b",)),
        ("Process Controller & Indicator", (r"\bcontrollers?\b", r"\bindicators?\b")),
        ("Scale & Balance", (r"\bbalances?\b", r"\bscales?\b")),
        ("Water Quality Test Kit", (r"\bwater quality test kits?\b", r"\bchemical test kits?\b")),
    )
    for category, patterns in direct_rules:
        if any(re.search(pattern, title_text, re.I) for pattern in patterns):
            return category, None

    # Jika title memakai nama dagang/generic seperti "grader", gunakan noun teknis
    # yang dinyatakan eksplisit di description. Supporting-product title sudah
    # diselesaikan di atas sehingga referensi ke meter pada accessory tidak menang.
    if not earliest_support:
        if re.search(r"\b(?:handheld\s+)?colorimeter\b|\bphotometer\b", context, re.I):
            return "Photometer & Colorimeter", None
        if re.search(r"\bspectrophotometer\b|\bspectrometer\b", context, re.I):
            return "Spectrophotometer & Spectrometer", None
        if re.search(r"\brefractometer\b", context, re.I):
            return "Refractometer & Brix Meter", None

    # 2) Supporting products. Specific solution rules precede generic reagent rules.
    if re.search(r"\b(?:(?:buffer|calibration|conductivity|orp|ph|storage|cleaning|filling|test)\s+(?:solution|standard)|calibration\s+buffer|electrolyte(?:\s+fill)?\s+solution)\b", title_text, re.I):
        return "Larutan Kalibrasi & Reagen", None
    if re.search(r"\breagents?\b|\bchecker reagent\b", title_text, re.I):
        return "Larutan Kalibrasi & Reagen", None
    if re.search(r"\b(?:electrodes?|probes?)\b", title_text, re.I):
        return "Elektrode & Probe", None
    if re.search(r"\b(?:cuvettes?|bottles?|caps?|tips?|tweezers?|syringes?|pipettes?|beakers?|vials?)\b", title_text, re.I):
        return "Laboratory Consumables", None
    if re.search(r"\b(?:accessor(?:y|ies)|replacement|spare|cables?|holders?|adapters?|cases?|connectors?|tubing|fittings?)\b", title_text, re.I):
        return "Aksesori & Suku Cadang", None

    # 3) Kategori baru yang memang valid tetapi belum ada di registry UJI.
    for proposed, patterns in VALID_NEW_CATEGORY_RULES:
        if any(re.search(pattern, title_text, re.I) for pattern in patterns):
            canonical = canonicalize_category(proposed)
            if canonical:
                return canonical, None
            return proposed, "REVIEW_NEW_CATEGORY"

    # 4) Taxonomy hanya boleh membantu bila nilainya cocok ke kategori canonical.
    for part in reversed(clean_taxonomy(taxonomy).split(" > ")):
        part = normalize_space(part)
        if not part or part.lower() in CATEGORY_TAXONOMY_BLOCKLIST or part.lower() in PROMO_CATEGORY_TOKENS:
            continue
        meaningful_tokens = [token for token in re.findall(r"[A-Za-z0-9]+", part) if len(token) > 1]
        if len(meaningful_tokens) < 2:
            continue
        canonical = canonicalize_category(part)
        if canonical:
            return canonical, None

    # 5) Fallback aman. Jangan pernah menampilkan taxonomy promo sebagai kategori.
    if re.search(r"\b(?:solution|reagent|standard|buffer)\b", title_text):
        return "Larutan Kalibrasi & Reagen", "REVIEW_CATEGORY"
    if re.search(r"\b(?:probe|electrode|sensor)\b", title_text):
        return "Elektrode & Probe", "REVIEW_CATEGORY"
    if re.search(r"\b(?:kit|bottle|cuvette|vial|cap|tip|tube|tubing)\b", title_text):
        return "Laboratory Consumables", "REVIEW_CATEGORY"
    if re.search(r"\b(?:meter|tester|analyzer|controller|instrument)\b", context):
        return "Laboratory Equipment", "REVIEW_CATEGORY"
    return "Laboratory Equipment", "REVIEW_CATEGORY"


def infer_category(title: str, taxonomy: str, description: str) -> str:
    return infer_category_with_status(title, taxonomy, description)[0]

def infer_parameters(
    title: str,
    taxonomy: str,
    spec_rows: Sequence[Tuple[str, str]],
    source_short: str = "",
) -> List[str]:
    # Parameter ukur ditentukan dari nama produk dan field spesifikasi pengukuran.
    # Field batas mekanis seperti "Max Pressure" pada elektrode tidak dianggap
    # sebagai parameter yang diukur.
    title_text = normalize_space(f"{title} {taxonomy}").lower()
    result: List[str] = []

    def add(label: str) -> None:
        if label not in result:
            result.append(label)

    for label, patterns in PARAMETER_PATTERNS.items():
        if any(re.search(pattern, title_text, flags=re.I) for pattern in patterns):
            add(label)

    measurement_qualifier = re.compile(r"\b(range|accuracy|resolution|calibration|measurement)\b", re.I)
    for key, _ in spec_rows:
        key_lower = normalize_space(key).lower()
        for label, patterns in PARAMETER_PATTERNS.items():
            if label == "Tekanan":
                continue
            if label == "Suhu":
                if re.search(r"\btemperature sensor\b|\btemperature range\b|\btemperature accuracy\b|\btemperature resolution\b", key_lower, re.I):
                    add(label)
                continue
            if measurement_qualifier.search(key_lower) and any(re.search(pattern, key_lower, flags=re.I) for pattern in patterns):
                add(label)

    # Tekanan hanya dianggap parameter jika tercantum pada nama produk atau field pengukuran tekanan.
    if re.search(r"\bpressure\b", title_text, re.I) or any(
        re.search(r"^\s*pressure\s+(range|accuracy|resolution|measurement)", key, re.I)
        for key, _ in spec_rows
    ):
        add("Tekanan")
    return result


def translate_product_type(title: str, category: str) -> str:
    category_map = {
        "Larutan Kalibrasi & Reagen": "larutan kalibrasi atau reagen pengujian",
        "Elektrode & Probe": "elektrode atau probe pengukuran",
        "Aksesori & Suku Cadang": "aksesori atau suku cadang instrumen",
        "Laboratory Consumables": "bahan habis pakai laboratorium",
        "Laboratory Equipment": "peralatan laboratorium",
        "Multiparameter Water Quality Meter": "alat ukur kualitas air multiparameter",
        "Photometer & Colorimeter": "fotometer atau colorimeter",
        "pH Meter": "alat ukur pH",
        "pH & ORP Meter": "alat ukur pH dan ORP",
        "Conductivity": "alat ukur konduktivitas",
        "TDS & Salinity Meter": "alat ukur TDS atau salinitas",
        "Dissolved Oxygen Meter": "alat ukur oksigen terlarut",
        "ORP Meter": "alat ukur ORP",
        "Turbidity Meter": "alat ukur kekeruhan",
        "Refractometer & Brix Meter": "refraktometer atau Brix meter",
        "Thermometer & Temperature Meter": "termometer atau alat ukur suhu",
        "Magnetic Stirrer": "pengaduk magnetik",
        "Titrator & Burette": "titrator atau buret",
        "Moisture Analyzer": "alat analisis kadar air",
        "Calibrator": "kalibrator",
    }
    return category_map.get(category, category.lower())

def translate_spec_key(key: str) -> str:
    result = normalize_space(key)
    for pattern, replacement in SPEC_KEY_TRANSLATIONS:
        if re.search(pattern, result, flags=re.I):
            result = re.sub(pattern, replacement, result, count=1, flags=re.I)
    result = re.sub(r"\s+", " ", result).strip(" :-")
    return result[:180]


def looks_english(text: str) -> bool:
    words = re.findall(r"[a-zA-Z]+", normalize_space(text).lower())
    if not words:
        return False
    en = sum(word in ENGLISH_MARKERS for word in words)
    ind = sum(word in INDONESIAN_MARKERS for word in words)
    return en >= 2 and en > ind


def smart_truncate(text: str, limit: int) -> str:
    text = normalize_space(text).strip(" -|:;,./")
    if len(text) <= limit:
        return text
    clipped = text[: limit + 1]
    if " " in clipped:
        clipped = clipped.rsplit(" ", 1)[0]
    return clipped.rstrip(" -|:;,./")


def build_seo_title(source_sku: str, title: str, category: str, limit: int = 60) -> str:
    suffix = " | UJI"
    preferred = normalize_space(f"Hanna {source_sku} {title}")
    if len(preferred) + len(suffix) <= limit:
        return preferred + suffix
    concise = normalize_space(f"Hanna {source_sku} {category}")
    if len(concise) + len(suffix) <= limit:
        return concise + suffix
    return smart_truncate(concise, limit - len(suffix)) + suffix


def parse_specification_table(value: object) -> List[Tuple[str, str]]:
    if value is None or not str(value).strip():
        return []
    soup = BeautifulSoup(str(value), "html.parser")
    rows: List[Tuple[str, str]] = []
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if len(cells) < 2:
            continue
        key = remove_promotional_content(cells[0].get_text(" ", strip=True))
        val = remove_promotional_content(" ".join(c.get_text(" ", strip=True) for c in cells[1:]))
        if not key or not val:
            continue
        if key.lower() in {"specification", "parameter"} and val.lower() in {"detail", "value", "nilai"}:
            continue
        pair = (key[:180], val[:1000])
        if pair not in rows:
            rows.append(pair)
    return rows


def escape_html(text: object) -> str:
    return html.escape(normalize_space(text), quote=True)


def source_line_is_noise(text: str) -> bool:
    candidate = normalize_space(text)
    if not candidate:
        return True
    lower = candidate.lower()
    if len(candidate) <= 2:
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?\s*(?:reviews?|stars?)?", lower):
        return True
    if re.search(r"(?:usd\s*)?\$\s*\d", candidate, flags=re.I):
        return True

    # Navigation/noise patterns seperti "close", "login", atau "resources"
    # hanya boleh membuang elemen pendek. Jangan membuang satu paragraf teknis
    # panjang hanya karena ada kalimat seperti "close the lid".
    if len(candidate) <= 220 and any(re.search(pattern, lower, flags=re.I) for pattern in SOURCE_NOISE_PATTERNS):
        return True
    if len(candidate) <= 220 and any(re.search(pattern, lower, flags=re.I) for pattern in PROMO_PATTERNS):
        return True
    return False


def clean_source_line(text: object) -> str:
    candidate = remove_urls(normalize_space(text))
    # XLSX Hanna sering menyimpan "Warranty Information -->" pada awal field
    # yang sama dengan deskripsi. Hapus hanya prefix noise, jangan seluruh isi.
    candidate = re.sub(r"^\s*warranty information\s*(?:-->)?\s*", "", candidate, flags=re.I)
    candidate = re.sub(r"^\s*(?:adding to cart|add to cart|quantity)\s*(?:-->)?\s*", "", candidate, flags=re.I)
    candidate = normalize_space(candidate)
    if source_line_is_noise(candidate):
        return ""
    return candidate


def html_to_structured_text(value: object, max_chars: int = 20000) -> str:
    if value is None or not str(value).strip():
        return ""
    raw = str(value)
    soup = BeautifulSoup(raw, "html.parser")
    for node in soup(["script", "style", "noscript", "svg", "form", "button", "input", "select", "option"]):
        node.decompose()
    for comment in soup.find_all(string=lambda item: isinstance(item, Comment)):
        comment.extract()

    lines: List[str] = []
    seen: set[str] = set()

    def add_line(line: str, prefix: str = "") -> None:
        cleaned = clean_source_line(line)
        if not cleaned:
            return
        fingerprint = re.sub(r"[^a-z0-9]+", "", cleaned.lower())
        if len(fingerprint) < 4 or fingerprint in seen:
            return
        seen.add(fingerprint)
        lines.append(prefix + cleaned)

    nodes = soup.find_all(["h1", "h2", "h3", "h4", "h5", "p", "li", "tr"])
    if not nodes:
        add_line(soup.get_text(" ", strip=True))
    else:
        for node in nodes:
            if node.name == "tr":
                cells = [clean_source_line(cell.get_text(" ", strip=True)) for cell in node.find_all(["th", "td"], recursive=False)]
                cells = [cell for cell in cells if cell]
                if len(cells) >= 2:
                    add_line(f"{cells[0]}: {' | '.join(cells[1:])}")
                continue
            content = node.get_text(" ", strip=True)
            if node.name and node.name.startswith("h"):
                add_line(content, "## ")
            elif node.name == "li":
                add_line(content, "- ")
            else:
                add_line(content)

    result = "\n".join(lines)
    if not result:
        result = clean_source_line(soup.get_text(" ", strip=True))
    if len(result) > max_chars:
        result = result[:max_chars].rsplit("\n", 1)[0]
    return result.strip()


def heading_level(node: object) -> int:
    name = getattr(node, "name", "") or ""
    match = re.fullmatch(r"h([1-6])", name, flags=re.I)
    return int(match.group(1)) if match else 7


def collect_heading_section_text(heading: object, max_chars: int = 5000) -> str:
    level = heading_level(heading)
    fragments: List[str] = []
    for sibling in heading.next_siblings:
        sibling_name = getattr(sibling, "name", "") or ""
        if re.fullmatch(r"h[1-6]", sibling_name, flags=re.I) and heading_level(sibling) <= level:
            break
        if sibling_name:
            fragments.append(str(sibling))
        elif str(sibling).strip():
            fragments.append(f"<p>{html.escape(str(sibling))}</p>")
        if sum(len(item) for item in fragments) >= max_chars * 2:
            break
    return html_to_structured_text("".join(fragments), max_chars=max_chars)


def extract_rich_product_content(soup: BeautifulSoup, json_description: str = "") -> str:
    """Ambil rich product content dari body halaman Hanna.

    Penting: fungsi ini TIDAK menentukan atau mengubah tabel spesifikasi. Tab
    Spesifikasi tetap diproses oleh parse_specification_table/build_spec_html.
    """
    for selector in (
        "header", "footer", "nav", "aside", ".modal", ".popup", ".cookie",
        ".productView-options", ".productView-price", ".productView-rating",
        ".productView-reviewLink", ".socialLinks", ".footer",
        "script", "style", "noscript", "form", "button",
    ):
        for node in soup.select(selector):
            node.decompose()

    sections: List[Tuple[str, str]] = []
    seen: set[str] = set()

    def add_section(title: str, body: object, max_chars: int = 9000) -> None:
        title_clean = clean_source_line(title) or "Informasi Produk"
        body_clean = html_to_structured_text(body, max_chars=max_chars)
        if not body_clean or len(body_clean) < 35:
            return
        fingerprint = re.sub(r"[^a-z0-9]+", "", body_clean.lower())
        if not fingerprint or fingerprint in seen:
            return
        # Blok identik atau hampir seluruhnya tercakup tidak dimasukkan dua kali.
        if any(len(fingerprint) > 180 and (fingerprint in old or old in fingerprint) for old in seen):
            return
        seen.add(fingerprint)
        sections.append((title_clean, body_clean))

    if json_description:
        add_section("Ringkasan Produk", json_description, max_chars=3500)

    # BigCommerce/Hanna dapat memindahkan rich description di antara beberapa
    # container. Pilih container dengan konten teknis paling kaya, bukan yang
    # pertama ditemukan.
    root_candidates: List[object] = []
    for selector in (
        ".productView-description",
        "[data-product-description]",
        "[itemprop='description']",
        ".product-description",
        ".productView-product",
        ".productView-info + div",
        "main .tabs-contents",
        "main",
    ):
        for node in soup.select(selector):
            if node not in root_candidates:
                root_candidates.append(node)

    def root_score(node: object) -> int:
        text = normalize_space(node.get_text(" ", strip=True))
        lower = text.lower()
        score = min(len(text) // 120, 120)
        for keyword in (
            "features at-a-glance", "features at a glance", "principle of operation",
            "significance of use", "calibration", "data logging", "data transfer",
            "good laboratory practice", "glp", "certificate of analysis",
            "application", "measurement", "standard", "traceable",
        ):
            if keyword in lower:
                score += 30
        # Hindari memilih seluruh halaman jika ada container produk yang lebih fokus.
        if len(text) > 60000:
            score -= 80
        return score

    rich_root = max(root_candidates, key=root_score) if root_candidates else None
    if rich_root is not None:
        # Simpan seluruh rich description. Heading di dalam blok tetap dipertahankan
        # oleh html_to_structured_text agar Features, Principle, Significance, dsb.
        # dapat dipakai kembali oleh generator/fallback.
        add_section("Detail Produk Lengkap", str(rich_root), max_chars=24000)

    # Tangkap bagian penting per-heading juga. Ini membuat evidence lebih stabil
    # jika struktur HTML Hanna berubah atau rich_root terlalu luas.
    heading_scope = rich_root if rich_root is not None else soup
    for heading in heading_scope.find_all(["h2", "h3", "h4", "h5", "h6"]):
        heading_text = clean_source_line(heading.get_text(" ", strip=True))
        if not heading_text:
            continue
        lower = heading_text.lower()
        if any(re.search(pattern, lower, flags=re.I) for pattern in EXCLUDED_SECTION_PATTERNS):
            continue
        if not any(re.search(pattern, lower, flags=re.I) for pattern in RELEVANT_SECTION_PATTERNS):
            continue
        body = collect_heading_section_text(heading, max_chars=8000)
        add_section(heading_text, body, max_chars=8000)

    # Beberapa halaman Hanna menaruh Certificate of Analysis, traceability,
    # supplied standards, atau package details tanpa heading eksplisit. Ambil
    # paragraf/list yang memuat kata kunci bernilai tinggi.
    keyword_patterns = (
        r"certificate of analysis", r"traceab", r"nist", r"supplied with",
        r"standard(?:s)?", r"data logging", r"data transfer", r"rs232", r"usb",
        r"memory", r"glp", r"fast tracker", r"software", r"included",
        r"package contents?", r"applications?", r"water quality", r"wastewater",
    )
    scope = rich_root if rich_root is not None else (soup.select_one("main") or soup.body or soup)
    extra_lines: List[str] = []
    extra_seen: set[str] = set()
    for node in scope.find_all(["p", "li"]):
        line = clean_source_line(node.get_text(" ", strip=True))
        if not line:
            continue
        if not any(re.search(pattern, line, flags=re.I) for pattern in keyword_patterns):
            continue
        marker = validation_text(line)
        if marker in extra_seen:
            continue
        extra_seen.add(marker)
        extra_lines.append(line)
        if len(extra_lines) >= 24:
            break
    if extra_lines:
        add_section("Kalibrasi, Data, dan Kelengkapan", "".join(f"<p>{html.escape(x)}</p>" for x in extra_lines), max_chars=9000)

    if not sections:
        main = soup.select_one("main") or soup.body or soup
        add_section("Informasi Produk", str(main), max_chars=18000)

    html_sections: List[str] = []
    for title, body in sections:
        body_lines = [line for line in body.splitlines() if line.strip()]
        body_html: List[str] = []
        list_items: List[str] = []

        def flush_list() -> None:
            nonlocal list_items
            if list_items:
                body_html.append("<ul>" + "".join(f"<li>{escape_html(item)}</li>" for item in list_items) + "</ul>")
                list_items = []

        for line in body_lines:
            if line.startswith("## "):
                flush_list()
                body_html.append(f"<h3>{escape_html(line[3:])}</h3>")
            elif line.startswith("- "):
                list_items.append(line[2:])
            else:
                flush_list()
                body_html.append(f"<p>{escape_html(line)}</p>")
        flush_list()
        html_sections.append(f"<section><h2>{escape_html(title)}</h2>{''.join(body_html)}</section>")
    return "".join(html_sections)

def build_source_evidence(
    name: str,
    source_sku: str,
    category: str,
    parameters: Sequence[str],
    source_short: str,
    source_description: str,
    spec_rows: Sequence[Tuple[str, str]],
    config: Dict[str, object],
    source_buckets: Optional[Dict[str, List[str]]] = None,
) -> str:
    """Susun evidence source-adaptive untuk AI dan validator.

    Hanya section yang benar-benar tersedia yang dicantumkan. Spesifikasi tetap
    dikunci dari XLSX/source dan tidak diubah oleh translation/product writer.
    """
    buckets = source_buckets if source_buckets is not None else classify_source_sections(source_short, source_description)
    profile = source_profile_from_buckets(buckets, spec_rows)
    available = [key for key, value in profile.items() if value and key != "FAQ_ELIGIBLE"]
    header = [
        f"NAMA PRODUK: {name}",
        f"SKU/MODEL: {source_sku}",
        f"KATEGORI: {category}",
        f"PARAMETER: {join_indonesian(parameters) or 'Tidak teridentifikasi'}",
        "AVAILABLE_SECTIONS: " + (", ".join(available) if available else "NONE"),
        f"FAQ_ELIGIBLE: {'YES' if profile.get('FAQ_ELIGIBLE') else 'NO'}",
    ]
    blocks: List[str] = ["\n".join(header)]

    if spec_rows:
        spec_text = "\n".join(f"- {key}: {value}" for key, value in spec_rows)
        blocks.append("SPECIFICATION (LOCKED FACTS):\n" + spec_text)

    labels = {
        "OVERVIEW": "OVERVIEW",
        "FEATURES": "FEATURES",
        "WORKING_PRINCIPLE": "WORKING PRINCIPLE",
        "APPLICATION": "APPLICATION",
        "CALIBRATION": "CALIBRATION / STANDARD / GLP",
        "CONNECTIVITY": "REPORTING / CONNECTION / DATA",
        "INSTALLATION": "INSTALLATION",
        "WHATS_INCLUDED": "WHAT'S INCLUDED",
        "MAINTENANCE": "MAINTENANCE",
        "COMPATIBILITY": "COMPATIBILITY",
    }
    for kind in labels:
        lines = buckets.get(kind, [])
        if not lines:
            continue
        blocks.append(labels[kind] + ":\n" + "\n".join(f"- {line}" for line in lines[:30]))

    limit = int(config.get("max_source_content_chars", 30000))
    selected: List[str] = []
    used = 0
    for block in blocks:
        if used + len(block) + 2 <= limit:
            selected.append(block)
            used += len(block) + 2
            continue
        remaining = limit - used
        if remaining > 500:
            clipped = block[:remaining].rsplit("\n", 1)[0].strip()
            if clipped:
                selected.append(clipped)
        break
    return "\n\n".join(selected).strip()

def supported_number_tokens(output: str, evidence: str) -> bool:
    return not unsupported_number_tokens(output, evidence)


def unsupported_number_tokens(output: str, evidence: str) -> List[str]:
    """Cari angka output yang benar-benar tidak didukung evidence."""
    output_tokens = set(normalized_number_tokens(output))
    evidence_tokens = set(normalized_number_tokens(evidence))
    word_numbers = {
        "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
        "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
        "ten": "10", "eleven": "11", "twelve": "12",
        "nol": "0", "satu": "1", "dua": "2", "tiga": "3", "empat": "4",
        "lima": "5", "enam": "6", "tujuh": "7", "delapan": "8", "sembilan": "9",
        "sepuluh": "10", "sebelas": "11", "dua belas": "12",
    }
    evidence_lower = validation_text(evidence)
    for word, number in word_numbers.items():
        if re.search(rf"\b{re.escape(word)}\b", evidence_lower):
            evidence_tokens.add(number)
    return sorted(output_tokens - evidence_tokens, key=lambda x: (len(x), x))

def extract_upper_technical_tokens(text: str) -> set[str]:
    return {
        token.upper()
        for token in re.findall(r"\b[A-Z][A-Z0-9./+_-]{1,}\b", text)
        if len(token) >= 2
    }


def clean_generated_text(value: object, max_words: int = 220) -> str:
    text = html.unescape(str(value or ""))
    text = re.sub(r"```(?:json|markdown|text)?", " ", text, flags=re.I)
    text = strip_html(text)
    text = normalize_space(text).strip('"').strip()
    words = text.split()
    if len(words) > max_words:
        text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
    return text


def high_value_spec_features(spec_rows: Sequence[Tuple[str, str]], limit: int = 10) -> List[Dict[str, str]]:
    features: List[Dict[str, str]] = []
    used_rows: set[Tuple[str, str]] = set()
    used_groups: Counter[str] = Counter()
    for group_title, patterns in HIGH_VALUE_SPEC_GROUPS:
        for key, value in spec_rows:
            marker = (normalize_space(key).lower(), normalize_space(value).lower())
            if marker in used_rows or not any(re.search(pattern, key, flags=re.I) for pattern in patterns):
                continue
            title = group_title if used_groups[group_title] == 0 else translate_spec_key(key)
            used_groups[group_title] += 1
            used_rows.add(marker)
            features.append({"title": title, "detail": normalize_space(value)})
            if len(features) >= limit:
                return features
    return features



def inject_inline_section_headings(text: str) -> str:
    """Pulihkan heading yang hilang pada source XLSX/plain text.

    Beberapa ekspor Hanna menyimpan heading seperti "Features at-a-glance"
    di tengah satu paragraf. Tanpa pemisahan ini, seluruh feature block dapat
    keliru dianggap sebagai Product Description.
    """
    marker_rules: Sequence[Tuple[str, str]] = (
        (r"\bFeatures?\s+at[- ]a[- ]glance\b", "Features at-a-glance"),
        (r"\bKey\s+Features\b", "Key Features"),
        (r"\bReporting\s+(?:and|&)\s+Connection\s+Options\b", "Reporting and Connection Options"),
        (r"\bWhat(?:'|’)?s\s+Included\b", "What's Included"),
        (r"\bWhat\s+is\s+Included\b", "What's Included"),
        (r"\bPrinciple\s+of\s+Operation\b", "Principle of Operation"),
        (r"\bWorking\s+Principle\b", "Working Principle"),
        (r"\bInstallation\s+(?:and|&)\s+(?:Setup|Mounting)\b", "Installation and Setup"),
        (r"\bData\s+Logging\s+(?:and|&)\s+Connectivity\b", "Data Logging and Connectivity"),
    )
    output_lines: List[str] = []
    for raw_line in str(text or "").splitlines() or [str(text or "")]:
        if raw_line.lstrip().startswith("## "):
            output_lines.append(raw_line)
            continue
        line = raw_line
        for pattern, heading in marker_rules:
            line = re.sub(pattern, f"\n## {heading}\n", line, flags=re.I)
        output_lines.extend(line.splitlines())
    return "\n".join(output_lines)


def split_feature_source_line(text: str) -> List[str]:
    """Pecah feature block plain-text menjadi item feature individual.

    Format umum Hanna adalah "Feature Title - explanation". Fungsi ini hanya
    memecah pola yang cukup kuat agar tanda minus/range teknis tidak ikut pecah.
    """
    value = normalize_space(text)
    if not value:
        return []
    pattern = re.compile(
        r"(?:^|(?<=[.!?])\s+)(?P<title>[A-Z][A-Za-z0-9/&+()°%?'.,\- ]{2,72}?)\s+[\-–—]\s+"
    )
    matches = list(pattern.finditer(value))
    if not matches:
        return [value]
    items: List[str] = []
    for index, match in enumerate(matches):
        title = normalize_space(match.group("title"))
        body_start = match.end()
        body_end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        body = normalize_space(value[body_start:body_end])
        if title and body:
            items.append(f"{title}: {body}")
    return items or [value]



def description_body_to_overview(source_short: str, source_description: str) -> Dict[str, List[str]]:
    """Treat the entire source Description body as Product Description evidence.

    Inline headings such as "Features at-a-glance" remain part of OVERVIEW.
    They do NOT authorize a UJI custom tab. This preserves source provenance.
    """
    raw = source_description.strip() or source_short.strip()
    if not raw:
        return {}
    text = html_to_structured_text(raw, max_chars=30000) if "<" in raw and ">" in raw else raw
    lines: List[str] = []
    seen: set[str] = set()
    current_heading = ""
    for raw_line in text.splitlines():
        line = normalize_space(raw_line)
        if not line:
            continue
        if line.startswith("## "):
            current_heading = normalize_space(line[3:])
            continue
        if line.startswith("- "):
            line = normalize_space(line[2:])
        if not line or source_line_is_noise(line):
            continue
        if current_heading:
            # Preserve inline heading context inside the description itself.
            candidate = f"{current_heading}: {line}"
        else:
            candidate = line
        marker = validation_text(candidate)
        if marker and marker not in seen:
            lines.append(candidate)
            seen.add(marker)
    return {"OVERVIEW": lines} if lines else {}


def parse_source_tabs_json(value: object) -> Dict[str, str]:
    if not value or not str(value).strip():
        return {}
    try:
        payload = json.loads(str(value))
    except (json.JSONDecodeError, TypeError, ValueError):
        return {}
    if not isinstance(payload, dict):
        return {}
    result: Dict[str, str] = {}
    for key, body in payload.items():
        title = normalize_space(key)
        content = str(body or "").strip()
        if title and content:
            result[title] = content
    return result


def classify_dedicated_source_tabs(source_tabs: object) -> Dict[str, List[str]]:
    """Classify only genuine source tabs/accordions/panels.

    An inline heading inside the Description body can never enter this function,
    so it can never create a UJI custom tab by keyword alone.
    """
    tabs = parse_source_tabs_json(source_tabs)
    buckets: Dict[str, List[str]] = {}
    for title, body in tabs.items():
        lower = validation_text(title)
        matched: Optional[str] = None
        for kind, patterns in SOURCE_SECTION_RULES:
            if kind == "OVERVIEW":
                continue
            if any(re.search(pattern, lower, re.I) for pattern in patterns):
                matched = kind
                break
        if not matched:
            continue
        text = html_to_structured_text(body, max_chars=12000) if "<" in body and ">" in body else body
        bucket = buckets.setdefault(matched, [])
        seen = {validation_text(item) for item in bucket}
        for raw_line in text.splitlines():
            line = normalize_space(raw_line)
            if not line or line.startswith("## "):
                continue
            if line.startswith("- "):
                line = normalize_space(line[2:])
            line = clean_source_line(line)
            if not line:
                continue
            candidates = split_feature_source_line(line) if matched == "FEATURES" else [line]
            for candidate in candidates:
                marker = validation_text(candidate)
                if marker and marker not in seen:
                    bucket.append(candidate)
                    seen.add(marker)
    return {kind: lines for kind, lines in buckets.items() if lines}


def merge_evidence_buckets(description_buckets: Dict[str, List[str]], tab_buckets: Dict[str, List[str]]) -> Dict[str, List[str]]:
    result: Dict[str, List[str]] = {}
    for source in (description_buckets, tab_buckets):
        for kind, lines in source.items():
            bucket = result.setdefault(kind, [])
            seen = {validation_text(item) for item in bucket}
            for line in lines:
                marker = validation_text(line)
                if marker and marker not in seen:
                    bucket.append(line)
                    seen.add(marker)
    return {kind: lines for kind, lines in result.items() if lines}

def parse_source_sections(source_short: str, source_description: str) -> Dict[str, List[str]]:
    """Ubah source Hanna menjadi section map tanpa mencampur short summary.

    Full description selalu diprioritaskan. Short description hanya menjadi
    fallback jika source benar-benar tidak mempunyai description. Ini penting
    agar Product Description UJI berasal dari body description source, bukan
    dari teaser singkat yang sering berisi feature labels.
    """
    has_full_description = bool(normalize_space(source_description))
    raw = source_description.strip() if has_full_description else source_short.strip()
    if not raw:
        return {}
    if "<" in raw and ">" in raw:
        text = html_to_structured_text(raw, max_chars=30000)
    else:
        text = raw
    text = inject_inline_section_headings(text)

    sections: Dict[str, List[str]] = {}
    current = "Deskripsi Produk"
    sections[current] = []
    for raw_line in text.splitlines():
        line = normalize_space(raw_line)
        if not line:
            continue
        if line.startswith("## "):
            current = normalize_space(line[3:]) or "Informasi Produk"
            sections.setdefault(current, [])
            continue
        if line.startswith("- "):
            line = normalize_space(line[2:])
        if not line or source_line_is_noise(line):
            continue
        bucket = sections.setdefault(current, [])
        marker = validation_text(line)
        if marker and marker not in {validation_text(item) for item in bucket}:
            bucket.append(line)
    return {title: lines for title, lines in sections.items() if lines}


def classify_source_sections(source_short: str, source_description: str) -> Dict[str, List[str]]:
    """Kelompokkan source ke evidence yang stabil dan jaga OVERVIEW tetap murni.

    OVERVIEW hanya berisi description/overview asli. Features, konektivitas,
    isi paket, dan section lain dipisahkan agar tidak ikut masuk ke deskripsi utama.
    """
    raw_sections = parse_source_sections(source_short, source_description)
    buckets: Dict[str, List[str]] = {}
    specific_found = False

    def add(kind: str, lines: Sequence[str]) -> None:
        bucket = buckets.setdefault(kind, [])
        seen = {validation_text(item) for item in bucket}
        candidates: List[str] = []
        for line in lines:
            if kind == "FEATURES":
                candidates.extend(split_feature_source_line(line))
            else:
                candidates.append(line)
        for line in candidates:
            cleaned = clean_source_line(line)
            marker = validation_text(cleaned)
            if cleaned and marker and marker not in seen:
                bucket.append(cleaned)
                seen.add(marker)

    for title, lines in raw_sections.items():
        lower = validation_text(title)
        matched: Optional[str] = None
        for kind, patterns in SOURCE_SECTION_RULES:
            if any(re.search(pattern, lower, re.I) for pattern in patterns):
                matched = kind
                break
        if matched:
            add(matched, lines)
            if matched != "OVERVIEW":
                specific_found = True
        else:
            # Unnamed opening text is genuine product description, not a reason
            # to invent a new custom tab.
            add("OVERVIEW", lines)

    if specific_found and len(buckets.get("OVERVIEW", [])) > 6:
        buckets["OVERVIEW"] = buckets["OVERVIEW"][:6]
    return {kind: lines for kind, lines in buckets.items() if lines}


def source_profile_from_buckets(buckets: Dict[str, List[str]], spec_rows: Sequence[Tuple[str, str]]) -> Dict[str, bool]:
    profile = {kind: bool(buckets.get(kind)) for kind, _ in SOURCE_SECTION_RULES}
    profile["SPECIFICATION"] = bool(spec_rows)
    profile["FAQ_ELIGIBLE"] = bool(buckets.get("FEATURES")) and sum(1 for key, value in profile.items() if key not in {"OVERVIEW", "SPECIFICATION", "FAQ_ELIGIBLE"} and value) >= 1
    return profile


def trim_source_buckets(
    buckets: Dict[str, List[str]],
    total_limit: int = 22000,
) -> Dict[str, List[str]]:
    """Batasi evidence translation tanpa membiarkan satu section memakan semua token."""
    per_section_limits = {
        "OVERVIEW": 5000,
        "FEATURES": 5500,
        "WORKING_PRINCIPLE": 3500,
        "APPLICATION": 3000,
        "CALIBRATION": 3000,
        "CONNECTIVITY": 3500,
        "WHATS_INCLUDED": 2500,
        "INSTALLATION": 2500,
        "MAINTENANCE": 2200,
        "COMPATIBILITY": 2200,
    }
    priority = (
        "OVERVIEW", "FEATURES", "WORKING_PRINCIPLE", "APPLICATION", "CALIBRATION",
        "CONNECTIVITY", "WHATS_INCLUDED", "INSTALLATION", "MAINTENANCE", "COMPATIBILITY",
    )
    result: Dict[str, List[str]] = {}
    total = 0
    for kind in priority:
        lines = buckets.get(kind, [])
        if not lines or total >= total_limit:
            continue
        section_limit = min(per_section_limits.get(kind, 2500), total_limit - total)
        section: List[str] = []
        used = 0
        for line in lines:
            line = normalize_space(line)
            if not line:
                continue
            remaining = section_limit - used
            if remaining <= 80:
                break
            if len(line) > remaining:
                clipped = smart_truncate(line, remaining)
                if clipped:
                    section.append(clipped)
                    used += len(clipped)
                break
            section.append(line)
            used += len(line) + 1
        if section:
            result[kind] = section
            total += used
    return result


def buckets_to_structured_html(buckets: Dict[str, List[str]]) -> str:
    labels = {
        "OVERVIEW": "Overview",
        "FEATURES": "Features",
        "WORKING_PRINCIPLE": "Principle of Operation",
        "APPLICATION": "Applications",
        "CALIBRATION": "Calibration and Data",
        "CONNECTIVITY": "Reporting and Connection Options",
        "INSTALLATION": "Installation",
        "WHATS_INCLUDED": "What's Included",
        "MAINTENANCE": "Maintenance",
        "COMPATIBILITY": "Compatibility",
    }
    parts: List[str] = []
    for kind, lines in buckets.items():
        if not lines:
            continue
        parts.append(f"<h2>{escape_html(labels.get(kind, kind.title()))}</h2>")
        parts.extend(f"<p>{escape_html(line)}</p>" for line in lines)
    return "".join(parts)


def english_prose_score(text: str) -> float:
    """Estimasi proporsi prose Inggris. Technical tokens tidak dihitung sebagai kesalahan."""
    words = re.findall(r"[A-Za-z]+", normalize_space(text).lower())
    if not words:
        return 0.0
    technical = {
        "ph", "orp", "ec", "tds", "glp", "usb", "rs", "din", "brix", "ntu", "fnu",
        "ppm", "ppt", "wifi", "bluetooth", "nist", "ansi", "iso", "lcd", "led",
    }
    # Hanya hitung kata prose Inggris. Istilah teknis seperti range, accuracy,
    # data logging, calibration, USB, dan model tidak dianggap kesalahan sendiri.
    en_common = {
        "the", "this", "that", "these", "those", "with", "without", "for", "from", "and", "or",
        "is", "are", "was", "were", "can", "will", "provides", "provide", "allows", "allow",
        "includes", "include", "designed", "used", "using", "available", "supplied", "into", "through",
        "when", "where", "which", "while", "also", "offers", "offer", "helps", "help",
    }
    meaningful = [word for word in words if word not in technical and len(word) > 1]
    if not meaningful:
        return 0.0
    english_hits = sum(word in en_common for word in meaningful)
    return english_hits / len(meaningful)


def contains_untranslated_english_prose(text: str) -> bool:
    sentences = re.split(r"(?<=[.!?])\s+|<br\s*/?>|</p>|</li>", strip_html(text), flags=re.I)
    for sentence in sentences:
        sentence = normalize_space(sentence)
        if len(sentence.split()) < 6:
            continue
        score = english_prose_score(sentence)
        if score >= 0.12:
            return True
        if looks_english(sentence) and score >= 0.07:
            return True
    return False


def source_section_lines(
    sections: Dict[str, List[str]],
    title_patterns: Sequence[str],
    limit: int = 8,
) -> List[str]:
    results: List[str] = []
    seen: set[str] = set()
    for title, lines in sections.items():
        if not any(re.search(pattern, title, flags=re.I) for pattern in title_patterns):
            continue
        for line in lines:
            marker = validation_text(line)
            if not marker or marker in seen:
                continue
            seen.add(marker)
            results.append(line)
            if len(results) >= limit:
                return results
    return results


def localize_spec_detail(title: str, detail: str) -> str:
    """Terjemahkan penghubung umum pada nilai spesifikasi untuk fallback saja.

    Fungsi ini tidak menyentuh tab Spesifikasi. Nilai angka, unit, model, dan
    istilah teknis tetap dipertahankan.
    """
    value = normalize_space(detail)
    if not value:
        return value
    replacements = (
        (r"(?i)\bone, two or three point calibration\b", "kalibrasi satu, dua, atau tiga titik"),
        (r"(?i)\btwo, three, or four-point calibration\b", "kalibrasi dua, tiga, atau empat titik"),
        (r"(?i)\btwo, three or four point calibration\b", "kalibrasi dua, tiga, atau empat titik"),
        (r"(?i)\bof reading plus\b", "dari pembacaan ditambah"),
        (r"(?i)\bof reading or\b", "dari pembacaan atau"),
        (r"(?i)\bwhichever is greater\b", "mana yang lebih besar"),
        (r"(?i)\bto\b", "hingga"),
        (r"(?i)\bor\b", "atau"),
        (r"(?i)\band\b", "dan"),
        (r"(?i)\brecords?\b", "data"),
        (r"(?i)\bautomatic\b", "otomatis"),
        (r"(?i)\bnon-condensing\b", "tanpa kondensasi"),
        (r"(?i)\bafter\b", "setelah"),
        (r"(?i)\bminutes? of non-use\b", "menit tidak digunakan"),
    )
    for pattern, replacement in replacements:
        value = re.sub(pattern, replacement, value)
    return normalize_space(value)


def localize_source_line(line: str) -> str:
    """Lokalisasi deterministik untuk fakta umum Hanna pada mode tanpa AI.

    Jika kalimat Inggris tidak dapat diparafrasekan dengan aman, kalimat dibuang
    dari fallback. Fakta mentah tetap berada di evidence dan dapat dipakai AI.
    """
    text = normalize_space(line).strip(" -")
    if not text:
        return ""
    lower = text.lower()
    source_english_markers = re.compile(
        r"\b(?:supplied|reference meter|traceable|includes?|allows?|provides?|using|used|available|"
        r"stored|recalled|downloaded|compatible|reading modes?|light detector|water quality|drinking water|"
        r"wastewater|certificate of analysis|designed for|measurements?|calibration conditions?)\b",
        re.I,
    )
    if not looks_english(text) and not source_english_markers.search(text):
        return text

    # Fitur yang sangat umum pada halaman instrumen Hanna.
    if "data logging" in lower or ("internal memory" in lower and "measure" in lower):
        nums = re.findall(r"\b\d+(?:[.,]\d+)?\b", text)
        amount = nums[0] if nums else ""
        return normalize_space(
            f"Pencatatan data: hingga {amount} hasil pengukuran dapat disimpan dalam memori internal dan dipanggil kembali saat diperlukan."
            if amount else "Pencatatan data: hasil pengukuran dapat disimpan dalam memori internal dan dipanggil kembali saat diperlukan."
        )
    if "data transfer" in lower or ("download" in lower and ("usb" in lower or "rs232" in lower)):
        interfaces = []
        for token in ("RS232", "USB", "Bluetooth", "Wi-Fi"):
            if token.lower() in lower:
                interfaces.append(token)
        software = re.search(r"\b(HI\d{4,}[A-Z0-9-]*)\b", text, flags=re.I)
        interface_text = " atau ".join(interfaces) if interfaces else "antarmuka yang tersedia"
        software_text = f" menggunakan software {software.group(1)}" if software else ""
        return f"Transfer data: data tersimpan dapat dipindahkan ke komputer melalui {interface_text}{software_text}."
    if "glp" in lower or "good laboratory practice" in lower:
        details = []
        if "calibration point" in lower:
            details.append("titik kalibrasi")
        if "date" in lower:
            details.append("tanggal")
        if "time" in lower:
            details.append("waktu")
        suffix = f", termasuk {join_indonesian(details)}" if details else ""
        return f"Fungsi GLP membantu ketertelusuran kondisi kalibrasi{suffix}."
    if "multiple reading modes" in lower or "continuous measurement" in lower or "signal averaging" in lower:
        modes = []
        if "normal measurement" in lower:
            modes.append("normal")
        if "continuous measurement" in lower:
            modes.append("kontinu")
        if "signal averaging" in lower:
            modes.append("signal averaging")
        return f"Mode pembacaan: mendukung {join_indonesian(modes)}." if modes else "Mendukung beberapa mode pembacaan sesuai kebutuhan pengukuran."
    if "backlit" in lower and ("display" in lower or "lcd" in lower):
        return "Layar dengan backlight membantu pembacaan data dan navigasi saat pengoperasian maupun kalibrasi."
    if "fast tracker" in lower or "tag identification system" in lower or "iButton".lower() in lower:
        return "Fast Tracker Tag Identification System membantu pencatatan waktu dan lokasi pengukuran pada titik sampling yang berbeda."
    if "automatic temperature compensation" in lower or "atc" in lower:
        return "Kompensasi suhu otomatis membantu menyesuaikan hasil pengukuran terhadap perubahan suhu sesuai fungsi instrumen."
    if "waterproof" in lower:
        return "Desain tahan air mendukung penggunaan pada kondisi kerja yang sesuai dengan rating perlindungan produk."
    if "calibration" in lower and ("point" in lower or "standard" in lower):
        # Pertahankan angka, satuan, dan nama standar, tetapi ringkas struktur bahasa.
        clean = re.sub(r"(?i)^calibration\s*[–—:-]?\s*", "", text)
        clean = re.sub(r"(?i)can be performed by using", "dapat dilakukan menggunakan", clean)
        clean = re.sub(r"(?i)can be performed", "dapat dilakukan", clean)
        clean = re.sub(r"(?i)using the supplied", "menggunakan standar yang disertakan", clean)
        clean = re.sub(r"(?i)standards?", "standar", clean)
        clean = re.sub(r"(?i)one, two or three", "satu, dua, atau tiga", clean)
        clean = re.sub(r"(?i)two, three, or four", "dua, tiga, atau empat", clean)
        if not looks_english(clean):
            return "Kalibrasi: " + clean.rstrip(".") + "."
    if "certificate of analysis" in lower:
        extra = ""
        if "nist" in lower and "trace" in lower:
            extra = " Reference meter memiliki ketertelusuran NIST."
        return "Produk dilengkapi Certificate of Analysis." + extra
    if "nist" in lower and "trace" in lower:
        return "Reference meter memiliki ketertelusuran NIST."
    if "specially designed for water quality measurements" in lower:
        return "Instrumen dirancang untuk pengukuran kualitas air, termasuk pengukuran pada rentang yang sesuai dengan spesifikasi produk."
    if "drinking water" in lower and "turbidity" in lower:
        return "Pengukuran kekeruhan digunakan untuk membantu penilaian kualitas air minum."
    if "wastewater" in lower and "turbidity" in lower:
        return "Pengukuran kekeruhan dapat digunakan untuk pemantauan proses dan hasil pengolahan air limbah."
    if "optical system" in lower and ("detector" in lower or "lamp" in lower):
        parts = []
        if "tungsten" in lower:
            parts.append("lampu filamen tungsten")
        if "90" in text:
            parts.append("detektor cahaya hambur 90°")
        if "180" in text:
            parts.append("detektor cahaya transmisi 180°")
        if parts:
            return "Sistem optik menggunakan " + join_indonesian(parts) + "."
    if "ratio" in lower and "ntu" in lower and "detector" in lower:
        return "Pada mode ratio, nilai NTU dihitung dari sinyal detektor untuk membantu mengompensasi gangguan warna dan perubahan intensitas lampu."

    # Terjemahan ringan hanya untuk kalimat yang menjadi bersih setelah penggantian.
    replacements = (
        (r"(?i)\bup to\b", "hingga"),
        (r"(?i)\bcan be stored\b", "dapat disimpan"),
        (r"(?i)\binternal memory\b", "memori internal"),
        (r"(?i)\bmeasurements\b", "hasil pengukuran"),
        (r"(?i)\bmeasurement\b", "pengukuran"),
        (r"(?i)\bcalibration\b", "kalibrasi"),
        (r"(?i)\bdata\b", "data"),
        (r"(?i)\bdisplay\b", "layar"),
        (r"(?i)\bportable\b", "portabel"),
        (r"(?i)\bdesigned for\b", "dirancang untuk"),
        (r"(?i)\bwith\b", "dengan"),
        (r"(?i)\band\b", "dan"),
        (r"(?i)\bfor\b", "untuk"),
    )
    candidate = text
    for pattern, replacement in replacements:
        candidate = re.sub(pattern, replacement, candidate)
    candidate = normalize_space(candidate)
    return candidate if candidate and not looks_english(candidate) else ""


def localize_unique_lines(lines: Sequence[str], limit: int) -> List[str]:
    result: List[str] = []
    seen: set[str] = set()
    for line in lines:
        localized = localize_source_line(line)
        marker = validation_text(localized)
        if not localized or not marker or marker in seen:
            continue
        seen.add(marker)
        result.append(localized)
        if len(result) >= limit:
            break
    return result


def source_fact_to_feature(line: str) -> Dict[str, str]:
    cleaned = normalize_space(line).strip(" -:;.")
    lower = cleaned.lower()
    fixed_titles = (
        ("fungsi glp", "Fungsi GLP"),
        ("mode pembacaan", "Mode pembacaan"),
        ("pencatatan data", "Pencatatan data"),
        ("transfer data", "Transfer data"),
        ("sistem optik", "Sistem optik"),
        ("kompensasi suhu otomatis", "Kompensasi suhu otomatis"),
        ("fast tracker", "Fast Tracker"),
    )
    for prefix, title in fixed_titles:
        if lower.startswith(prefix):
            detail = cleaned.split(":", 1)[1].strip() if ":" in cleaned else cleaned
            return {"title": title, "detail": detail}
    for separator in (" - ", " – ", ": "):
        if separator in cleaned:
            left, right = cleaned.split(separator, 1)
            if 1 <= len(left.split()) <= 8 and right.strip():
                return {"title": smart_truncate(left, 70), "detail": normalize_space(right)}
    sentence = re.split(r"(?<=[.!?])\s+", cleaned, maxsplit=1)[0]
    words = sentence.split()
    feature_verbs = {
        "is", "are", "include", "includes", "included", "allow", "allows",
        "provide", "provides", "store", "stores", "record", "records",
        "support", "supports", "use", "uses", "offer", "offers", "feature", "features",
    }
    verb_index = next((index for index, word in enumerate(words) if word.lower().strip(" ,;:") in feature_verbs), None)
    if verb_index is not None and 1 <= verb_index <= 6:
        title_words = words[:verb_index]
    else:
        title_words = words[:5]
    title = " ".join(title_words).rstrip(" ,;:")
    return {"title": smart_truncate(title or "Fitur produk", 70), "detail": cleaned}


def append_unique_faq(items: List[Dict[str, str]], question: str, answer: str) -> None:
    question = normalize_space(question)
    answer = normalize_space(answer)
    if not question or not answer:
        return
    marker = validation_text(question)
    if marker in {validation_text(item.get("question", "")) for item in items}:
        return
    items.append({"question": question, "answer": answer})

def build_fallback_product_content(
    name: str,
    source_sku: str,
    category: str,
    parameters: Sequence[str],
    spec_rows: Sequence[Tuple[str, str]],
    source_short: str = "",
    source_description: str = "",
    source_buckets: Optional[Dict[str, List[str]]] = None,
) -> Dict[str, object]:
    """Fallback deterministik yang mengikuti kedalaman source.

    Tidak ada fitur dari tabel spesifikasi, tidak ada aplikasi/FAQ filler, dan
    tidak ada minimum panjang. Jika source hanya satu paragraf, output juga ringkas.
    """
    function = product_function_sentence(parameters, category)
    product_type = translate_product_type(name, category)
    buckets = source_buckets if source_buckets is not None else classify_source_sections(source_short, source_description)

    overview_lines = [normalize_space(item) for item in buckets.get("OVERVIEW", []) if normalize_space(item)]
    # Product Description tidak boleh dibuat dari nama, kategori, atau spesifikasi.
    # Jika source tidak memiliki description/overview, biarkan kosong.
    if overview_lines:
        overview = " ".join(overview_lines[:6])
        short_description = first_meaningful_sentence(overview, max_chars=280) or overview
    else:
        overview = ""
        short_description = ""

    feature_lines = [normalize_space(item) for item in buckets.get("FEATURES", []) if normalize_space(item)]
    features: List[Dict[str, str]] = []
    for line in feature_lines[:10]:
        feature = source_fact_to_feature(line)
        if feature.get("title") and feature.get("detail"):
            features.append(feature)

    principle_lines = [normalize_space(item) for item in buckets.get("WORKING_PRINCIPLE", []) if normalize_space(item)]
    applications = [normalize_space(item) for item in buckets.get("APPLICATION", []) if normalize_space(item)][:7]
    calibration_data = [normalize_space(item) for item in buckets.get("CALIBRATION", []) if normalize_space(item)]
    calibration_data.extend(normalize_space(item) for item in buckets.get("CONNECTIVITY", []) if normalize_space(item))
    calibration_data = list(dict.fromkeys(calibration_data))[:8]

    considerations_lines: List[str] = []
    for kind in ("COMPATIBILITY", "MAINTENANCE"):
        considerations_lines.extend(normalize_space(item) for item in buckets.get(kind, []) if normalize_space(item))
    considerations = " ".join(list(dict.fromkeys(considerations_lines))[:3])

    # Keunggulan hanya boleh berasal dari section Features/Benefits sumber.
    advantages = feature_lines[:9]

    faq_items: List[Dict[str, str]] = []
    profile = source_profile_from_buckets(buckets, spec_rows)
    if profile.get("FAQ_ELIGIBLE"):
        append_unique_faq(
            faq_items,
            f"Apa fungsi {name}?",
            " ".join(overview_lines[:2]),
        )
        if feature_lines:
            append_unique_faq(faq_items, f"Apa fitur utama {source_sku}?", " ".join(feature_lines[:4]))
        if principle_lines:
            append_unique_faq(faq_items, f"Bagaimana prinsip kerja {source_sku}?", " ".join(principle_lines[:4]))
        if applications:
            append_unique_faq(faq_items, f"Untuk aplikasi apa {source_sku} digunakan?", " ".join(applications[:4]))
        if buckets.get("CONNECTIVITY"):
            append_unique_faq(faq_items, f"Bagaimana konektivitas atau pengelolaan data {source_sku}?", " ".join(buckets["CONNECTIVITY"][:4]))
        if buckets.get("WHATS_INCLUDED"):
            append_unique_faq(faq_items, f"Apa yang termasuk dalam paket {source_sku}?", " ".join(buckets["WHATS_INCLUDED"][:4]))
        if buckets.get("MAINTENANCE"):
            append_unique_faq(faq_items, f"Bagaimana perawatan {source_sku}?", " ".join(buckets["MAINTENANCE"][:4]))

    return {
        "short_description": clean_generated_text(short_description, max_words=70),
        "overview": clean_description_for_formatting(overview),
        "feature_items": features,
        "working_principle": clean_generated_text(" ".join(principle_lines[:5]), max_words=180),
        "application_items": applications,
        "calibration_data_items": calibration_data,
        "considerations": clean_generated_text(considerations, max_words=130),
        "advantages": advantages,
        "faq": faq_items[:8],
    }

DESCRIPTION_INLINE_SECTION_RULES: Sequence[Tuple[str, str]] = (
    ("Fitur Utama", r"\b(?:Fitur\s+(?:sekilas|utama)|Features?\s+at[- ]a[- ]glance|Key\s+Features)\s*(?:[:：]|[\-–—])?\s*"),
    ("Teori Operasi", r"\b(?:Teori\s+Operasi|Theory\s+of\s+Operation)\s*(?:[:：]|[\-–—])?\s*"),
    ("Prinsip Kerja", r"\b(?:Prinsip\s+Kerja|Principle\s+of\s+Operation|Working\s+Principle)\s*(?:[:：]|[\-–—])?\s*"),
    ("Tabel Konversi", r"\b(?:Tabel\s+Konversi(?:\s+untuk\s+[^:.;]{1,30})?|Conversion\s+Table(?:\s+for\s+[^:.;]{1,30})?)\s*(?:[:：]|[\-–—])?\s*"),
    ("Cara Penggunaan", r"\b(?:Cara\s+Penggunaan|Petunjuk\s+Penggunaan|Directions?\s+for\s+Use|How\s+to\s+Use)\s*(?:[:：]|[\-–—])?\s*"),
    # "Aplikasi" adalah kata yang sangat umum dalam prose. Karena itu ia jauh
    # lebih ketat daripada heading lain dan hanya boleh menjadi section ketika
    # tampil sebagai label eksplisit seperti "Aplikasi:" / "Applications:".
    ("Aplikasi", r"\b(?:Aplikasi(?:\s+Penggunaan)?|Applications?|Recommended\s+Use)\s*(?:[:：]|[\-–—])\s*"),
)

DESCRIPTION_HEADING_STRONG_TITLES = {
    "Fitur Utama",
    "Teori Operasi",
    "Prinsip Kerja",
    "Tabel Konversi",
    "Cara Penggunaan",
}

DESCRIPTION_DANGLING_END_WORDS = {
    "untuk", "pada", "dalam", "dengan", "sebagai", "ke", "dari", "oleh",
    "dan", "atau", "terhadap", "mengenai", "berupa", "melalui", "antara",
    "hingga", "sampai", "agar", "karena", "yang", "serta", "termasuk",
}


def clean_description_for_formatting(value: object) -> str:
    """Bersihkan description untuk formatter TANPA word-limit/truncation.

    Content writer/validator sudah menentukan isi faktual. Layer formatter tidak
    boleh lagi memotong jumlah kata. Ia hanya menghapus wrapper/markup residual
    lalu menata struktur HTML.
    """
    text = html.unescape(str(value or ""))
    text = re.sub(r"```(?:json|markdown|text)?", " ", text, flags=re.I)
    text = strip_html(text)
    return normalize_space(text).strip('"').strip()


def _previous_nonspace_char(value: str, position: int) -> str:
    left = value[:position].rstrip()
    return left[-1] if left else ""


def _previous_word(value: str, position: int) -> str:
    left = value[:position].rstrip()
    match = re.search(r"([A-Za-zÀ-ÿ]+)$", left)
    return match.group(1).lower() if match else ""


def is_safe_inline_heading_match(value: str, start: int, end: int, title: str) -> bool:
    """Validasi posisi heading agar keyword dalam kalimat tidak ikut terpotong.

    Heading diterima hanya bila berada di awal description atau sesudah akhir
    kalimat yang jelas. Untuk kata generik seperti Aplikasi, regex sendiri sudah
    mewajibkan delimiter eksplisit (: / dash). Jika konteks sebelum heading
    berakhir dengan preposisi/konektor, match ditolak sebagai fragmen prose.
    """
    if start < 0 or end <= start:
        return False
    if start == 0:
        return True

    prev_char = _previous_nonspace_char(value, start)
    prev_word = _previous_word(value, start)
    if prev_word in DESCRIPTION_DANGLING_END_WORDS:
        return False

    # Setelah normalize_space, line break source memang hilang. Titik, tanda
    # tanya, atau tanda seru adalah boundary paling aman untuk heading inline.
    if prev_char not in ".!?":
        return False

    matched = value[start:end].strip()
    has_explicit_delimiter = bool(re.search(r"[:：\-–—]\s*$", matched))
    if title == "Aplikasi":
        return has_explicit_delimiter

    # Strong labels boleh tanpa colon bila muncul tepat sesudah akhir kalimat,
    # contoh: "... larutan. Features at-a-glance Four-ring Probe - ..."
    return has_explicit_delimiter or title in DESCRIPTION_HEADING_STRONG_TITLES


def has_dangling_description_fragment(text: str) -> bool:
    """Deteksi fragmen yang tidak layak ditinggalkan sebelum sebuah heading."""
    raw = normalize_space(text).rstrip()
    if not raw:
        return False
    # Kalimat yang sudah ditutup jelas bukan dangling fragment. Pemeriksaan ini
    # dilakukan sebelum parsing kata agar angka di akhir seperti "10.01." tidak
    # membuat kata sebelumnya (mis. "dan") salah dianggap connector gantung.
    if raw[-1:] in ".!?":
        return False
    value = raw.rstrip(" ,;:-–—")
    words = re.findall(r"[A-Za-zÀ-ÿ]+", value.lower())
    if not words:
        return False
    if words[-1] in DESCRIPTION_DANGLING_END_WORDS:
        return True
    # Fragmen sangat pendek tanpa tanda akhir biasanya menandakan split palsu,
    # misalnya "HI9812-51 adalah pH/EC/TDS meter untuk".
    if len(words) <= 8:
        return True
    return False


def split_description_paragraphs(text: str, max_chars: int = 620, max_sentences: int = 3) -> List[str]:
    """Pecah prose panjang menjadi paragraf tanpa menulis ulang isi.

    Pemisahan hanya dilakukan di batas kalimat. Angka, unit, model, simbol, dan
    susunan kata di dalam kalimat tidak disentuh.
    """
    value = normalize_space(text)
    if not value:
        return []
    sentences = [normalize_space(item) for item in re.split(r"(?<=[.!?])\s+(?=[A-ZÀ-Ý0-9])", value) if normalize_space(item)]
    if len(sentences) <= 1:
        return [value]
    paragraphs: List[str] = []
    current: List[str] = []
    current_chars = 0
    for sentence in sentences:
        projected = current_chars + (1 if current else 0) + len(sentence)
        if current and (len(current) >= max_sentences or projected > max_chars):
            paragraphs.append(" ".join(current))
            current = []
            current_chars = 0
        current.append(sentence)
        current_chars += (1 if current_chars else 0) + len(sentence)
    if current:
        paragraphs.append(" ".join(current))
    return paragraphs or [value]


def split_inline_feature_items(text: str) -> List[Tuple[str, str]]:
    """Pisahkan feature inline berbentuk `Judul - penjelasan` atau `Judul: penjelasan`.

    Regex sengaja konservatif agar rentang teknis seperti -0.1 atau 0-50°C tidak
    dianggap sebagai pemisah feature.
    """
    value = normalize_space(text)
    if not value:
        return []
    pattern = re.compile(
        r"(?:^|(?<=[.!?])\s+)"
        r"(?P<title>[A-ZÀ-Ý][A-Za-zÀ-ÿ0-9/&+()°%?'., ]{2,62}?)"
        r"\s+(?:[\-–—]|:)\s+"
    )
    matches = list(pattern.finditer(value))
    if not matches:
        return []
    items: List[Tuple[str, str]] = []
    for index, match in enumerate(matches):
        title = normalize_space(match.group("title")).strip(" .,:;-–—")
        body_start = match.end()
        body_end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        body = normalize_space(value[body_start:body_end]).strip()
        # Reject suspiciously long titles or fragments that look like normal prose.
        if not title or not body or len(title.split()) > 8:
            continue
        items.append((title, body))
    return items


def split_known_description_sections(text: str) -> List[Tuple[Optional[str], str]]:
    """Temukan heading inline eksplisit tanpa memecah prose normal.

    V4.3.4 menambahkan positional guard. Sebuah keyword hanya boleh menjadi
    heading ketika tampil di posisi heading yang aman. Jika hasil split akan
    meninggalkan dangling fragment, seluruh pemisahan tersebut dibatalkan.
    """
    value = normalize_space(text)
    if not value:
        return []
    matches: List[Tuple[int, int, str]] = []
    occupied: List[Tuple[int, int]] = []
    for title, pattern in DESCRIPTION_INLINE_SECTION_RULES:
        for match in re.finditer(pattern, value, flags=re.I):
            start, end = match.span()
            if not is_safe_inline_heading_match(value, start, end, title):
                continue
            if any(not (end <= left or start >= right) for left, right in occupied):
                continue
            matches.append((start, end, title))
            occupied.append((start, end))
    matches.sort(key=lambda item: item[0])
    if not matches:
        return [(None, value)]

    # Regression guard: jangan pernah membuat heading jika potongan tepat di
    # depannya berubah menjadi frase gantung seperti "meter untuk".
    previous_boundary = 0
    for start, _, _ in matches:
        preceding = normalize_space(value[previous_boundary:start])
        if preceding and has_dangling_description_fragment(preceding):
            return [(None, value)]
        previous_boundary = start

    sections: List[Tuple[Optional[str], str]] = []
    if matches[0][0] > 0:
        intro = normalize_space(value[:matches[0][0]])
        if intro:
            sections.append((None, intro))
    for index, (start, end, title) in enumerate(matches):
        body_end = matches[index + 1][0] if index + 1 < len(matches) else len(value)
        body = normalize_space(value[end:body_end]).strip(" :-–—")
        if body:
            sections.append((title, body))
    return sections


def format_product_description_overview(overview: str) -> str:
    """Render overview menjadi HTML rapi tanpa mengubah fakta.

    Layer ini berjalan *setelah* AI/fallback lolos factual validator. Ia tidak
    memanggil AI, tidak menerjemahkan ulang, dan tidak menambah kalimat. Hanya
    mengubah struktur visual: paragraf, h2/h3, dan feature blocks.
    """
    value = clean_description_for_formatting(overview)
    if not value:
        return ""

    html_blocks: List[str] = []
    for section_title, body in split_known_description_sections(value):
        if section_title:
            html_blocks.append(f"<h2>{escape_html(section_title)}</h2>")

        if section_title == "Fitur Utama":
            feature_items = split_inline_feature_items(body)
            if feature_items:
                consumed = " ".join(f"{title} - {detail}" for title, detail in feature_items)
                # Feature body berasal dari teks yang sama. Tidak ada ringkasan baru.
                for title, detail in feature_items:
                    html_blocks.append(f"<h3>{escape_html(title)}</h3>")
                    for paragraph in split_description_paragraphs(detail, max_chars=560, max_sentences=3):
                        html_blocks.append(f"<p>{escape_html(paragraph)}</p>")
                continue

        for paragraph in split_description_paragraphs(body):
            html_blocks.append(f"<p>{escape_html(paragraph)}</p>")

    # Safety check: formatter tidak boleh menghapus angka atau token teknis.
    formatted_text = strip_html("".join(html_blocks))
    source_numbers = set(normalized_number_tokens(value))
    formatted_numbers = set(normalized_number_tokens(formatted_text))
    if source_numbers != formatted_numbers:
        logging.warning("Description Formatter fallback: token angka berubah saat formatting.")
        return "".join(f"<p>{escape_html(p)}</p>" for p in split_description_paragraphs(value))
    source_tokens = extract_upper_technical_tokens(value)
    formatted_tokens = extract_upper_technical_tokens(formatted_text)
    if not source_tokens.issubset(formatted_tokens):
        logging.warning("Description Formatter fallback: token teknis berubah saat formatting.")
        return "".join(f"<p>{escape_html(p)}</p>" for p in split_description_paragraphs(value))
    return "".join(html_blocks)


def description_html_has_suspicious_heading_split(description_html: str) -> bool:
    """Fail-safe QA untuk heading yang meninggalkan fragmen kalimat.

    Formatter seharusnya sudah mencegah kondisi ini. Pemeriksaan kedua tetap
    dilakukan sebelum status READY agar regression di masa depan tidak lolos
    diam-diam saat generate massal.
    """
    if not normalize_space(description_html):
        return False
    soup = BeautifulSoup(str(description_html), "html.parser")
    for heading in soup.find_all(["h2", "h3"]):
        title = normalize_space(heading.get_text(" ", strip=True))
        if not title or title == "Informasi Harga dan Pemesanan":
            continue
        previous = heading.find_previous_sibling()
        if previous is None or getattr(previous, "name", "") != "p":
            continue
        previous_text = normalize_space(previous.get_text(" ", strip=True))
        if not previous_text:
            continue
        # Heading setelah paragraf yang belum berakhir sebagai kalimat adalah
        # sinyal kuat false split, walaupun fragmennya panjang.
        if previous_text.rstrip()[-1:] not in ".!?":
            return True
        if has_dangling_description_fragment(previous_text):
            return True
    return False


def build_rich_description_html(
    name: str,
    source_sku: str,
    content: Dict[str, object],
    config: Dict[str, object],
) -> str:
    """Bangun Product Description hanya dari OVERVIEW/DESCRIPTION source.

    Konten yang sudah lolos validator dirapikan secara deterministik. Formatter
    hanya membuat paragraf/heading/subheading dan tidak mengubah fakta. Dedicated
    source tabs tetap ditangani oleh dynamic custom tabs. Jika source tidak punya
    overview/description, Product Description tetap kosong.
    """
    overview = clean_description_for_formatting(content.get("overview"))
    if not overview:
        return ""

    description_html = format_product_description_overview(overview)
    blocks: List[str] = [description_html]
    blocks.append(
        "<h2>Informasi Harga dan Pemesanan</h2>"
        f'<p>Untuk informasi harga, ketersediaan, dan konsultasi teknis {escape_html(source_sku)}, '
        f'hubungi Tim {escape_html(config["company_name"])} melalui WhatsApp '
        f'<a href="https://wa.me/{escape_html(config["whatsapp_number"])}">{escape_html(config["whatsapp_display"])}</a> '
        f'atau email <a href="mailto:{escape_html(config["email"])}">{escape_html(config["email"])}</a>.</p>'
    )
    return "".join(blocks)


def build_rich_advantages_html(content: Dict[str, object]) -> str:
    advantages = content.get("advantages") if isinstance(content.get("advantages"), list) else []
    cleaned = [clean_generated_text(item, max_words=70) for item in advantages]
    cleaned = [item for item in cleaned if item]
    return "<h2>Keunggulan Produk</h2><ul>" + "".join(f"<li>{escape_html(item)}</li>" for item in cleaned[:10]) + "</ul>"


def build_rich_faq_html(content: Dict[str, object]) -> str:
    faq = content.get("faq") if isinstance(content.get("faq"), list) else []
    items: List[Dict[str, str]] = []
    for item in faq[:12]:
        if not isinstance(item, dict):
            continue
        question = clean_generated_text(item.get("question"), max_words=35)
        answer = clean_generated_text(item.get("answer"), max_words=130)
        if question and answer:
            items.append({"question": question, "answer": answer})
    return faq_items_to_html(items)


def build_lines_tab_html(title: str, lines: Sequence[str]) -> str:
    cleaned = [clean_generated_text(item, max_words=90) for item in lines]
    cleaned = [item for item in cleaned if item]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return f"<h2>{escape_html(title)}</h2><p>{escape_html(cleaned[0])}</p>"
    return f"<h2>{escape_html(title)}</h2><ul>" + "".join(f"<li>{escape_html(item)}</li>" for item in cleaned[:12]) + "</ul>"


def build_dynamic_custom_tabs(
    source_sku: str,
    spec_rows: Sequence[Tuple[str, str]],
    source_buckets: Dict[str, List[str]],
    content: Dict[str, object],
) -> List[Tuple[str, str]]:
    """Pilih maksimal tiga custom tab berdasarkan evidence sumber.

    Jika section sumber lebih banyak daripada slot yang tersedia, section tambahan
    digabung ke satu tab informatif daripada dibuang atau diganti filler FAQ.
    """
    selected: List[Tuple[str, str]] = []

    if spec_rows:
        selected.append(("Spesifikasi", build_spec_html(source_sku, list(spec_rows))))

    if source_buckets.get("FEATURES") and len(selected) < 3:
        advantages_html = build_rich_advantages_html(content)
        if advantages_html.count("<li>") >= 1:
            selected.append(("Keunggulan Produk", advantages_html))

    evidence_tabs: Sequence[Tuple[str, int, str]] = (
        ("WHATS_INCLUDED", 92, "Isi Paket"),
        ("CONNECTIVITY", 90, "Koneksi & Pelaporan"),
        ("INSTALLATION", 88, "Instalasi & Penggunaan"),
        ("WORKING_PRINCIPLE", 86, "Prinsip Kerja"),
        ("APPLICATION", 85, "Aplikasi & Penggunaan"),
        ("CALIBRATION", 84, "Kalibrasi & Standar"),
        ("MAINTENANCE", 82, "Perawatan"),
        ("COMPATIBILITY", 80, "Kompatibilitas"),
    )
    secondary: List[Tuple[int, str, str, str]] = []
    for kind, priority, title in evidence_tabs:
        lines = source_buckets.get(kind, [])
        html_content = build_lines_tab_html(title, lines) if lines else ""
        if html_content:
            secondary.append((priority, kind, title, html_content))
    secondary.sort(key=lambda item: item[0], reverse=True)

    remaining = max(0, 3 - len(selected))
    if remaining:
        if len(secondary) <= remaining:
            selected.extend((title, html_content) for _, _, title, html_content in secondary)
        elif remaining >= 2:
            # Ambil section terpenting secara terpisah, lalu gabungkan sisanya.
            for _, _, title, html_content in secondary[: remaining - 1]:
                selected.append((title, html_content))
            remainder = secondary[remaining - 1 :]
            combined = "".join(html_content for _, _, _, html_content in remainder)
            if combined:
                selected.append(("Informasi Produk", combined))
        else:
            combined = "".join(html_content for _, _, _, html_content in secondary)
            if combined:
                selected.append(("Informasi Produk", combined))

    # FAQ hanya memakai slot tersisa setelah semua section sumber yang bernilai tinggi.
    if len(selected) < 3 and source_buckets.get("FEATURES"):
        faq_html = build_rich_faq_html(content)
        if faq_html.count("<h3>") >= 3:
            selected.append(("FAQ", faq_html))

    return selected[:3]

def build_spec_html(source_sku: str, spec_rows: List[Tuple[str, str]]) -> str:
    rows = list(spec_rows)
    if source_sku and not any(key.lower() in {"model", "sku", "product code", "kode produk"} for key, _ in rows):
        rows.insert(0, ("Model", source_sku))
    if not rows:
        return "<h2>Spesifikasi</h2><p>Data spesifikasi perlu dikonfirmasi sebelum publikasi.</p>"
    tr_html = "".join(
        f'<tr><th scope="row">{escape_html(translate_spec_key(key))}</th><td>{escape_html(val)}</td></tr>'
        for key, val in rows
    )
    return (
        '<h2>Spesifikasi</h2>'
        '<table class="uji-spec-table uji-spec-general">'
        '<thead><tr><th>Parameter</th><th>Nilai</th></tr></thead>'
        f'<tbody>{tr_html}</tbody></table>'
    )


def first_meaningful_sentence(text: str, max_chars: int = 220) -> str:
    cleaned = remove_promotional_content(text)
    if not cleaned:
        return ""
    pieces = re.split(r"(?<=[.!?])\s+", cleaned)
    result = pieces[0].strip()
    if len(result) > max_chars:
        result = result[:max_chars].rsplit(" ", 1)[0].rstrip(" ,;:") + "."
    return result


def product_function_sentence(parameters: Sequence[str], category: str) -> str:
    if parameters:
        if len(parameters) == 1:
            return f"mengukur atau menganalisis parameter {parameters[0]}"
        if len(parameters) == 2:
            joined = f"{parameters[0]} dan {parameters[1]}"
        else:
            joined = ", ".join(parameters[:-1]) + f", dan {parameters[-1]}"
        return f"mengukur atau menganalisis parameter {joined}"
    category_lower = category.lower()
    if "larutan" in category_lower:
        return "mendukung proses kalibrasi, perawatan, atau verifikasi instrumen"
    if "reagen" in category_lower:
        return "mendukung analisis sampel menggunakan metode pengujian yang sesuai"
    if "elektrode" in category_lower or "electrode" in category_lower or "probe" in category_lower:
        return "digunakan sebagai sensor atau probe pada instrumen yang kompatibel"
    if "aksesori" in category_lower or "suku cadang" in category_lower or "accessories" in category_lower or "spare" in category_lower:
        return "digunakan sebagai aksesori atau komponen pendukung instrumen"
    if "consumable" in category_lower:
        return "digunakan sebagai bahan habis pakai laboratorium sesuai fungsi produknya"
    return "mendukung kegiatan pengukuran, pengujian, atau analisis sampel"


def build_short_description(name: str, parameters: Sequence[str], category: str, source_short: str) -> str:
    function = product_function_sentence(parameters, category)
    # Ringkasan dibuat ulang, tidak menempelkan kalimat sumber secara mentah.
    return normalize_space(f"{name} merupakan produk {category.lower()} untuk {function}.")


def build_description(
    name: str,
    source_sku: str,
    category: str,
    parameters: Sequence[str],
    source_short: str,
    source_description: str,
    config: Dict[str, object],
) -> str:
    function = product_function_sentence(parameters, category)
    product_type = translate_product_type(name, category)

    intro = (
        f"<p><strong>{escape_html(name)}</strong> merupakan {escape_html(product_type)} "
        f"yang digunakan untuk {escape_html(function)}. Informasi teknis utama tersedia pada tab spesifikasi "
        "agar pemilihan produk dapat disesuaikan dengan metode, sampel, dan kondisi penggunaan.</p>"
    )

    parameter_items = ""
    if parameters:
        parameter_items = "".join(f"<li>{escape_html(item)}</li>" for item in parameters)
    else:
        parameter_items = "<li>Fungsi produk perlu disesuaikan dengan instrumen dan metode pengujian yang digunakan.</li>"

    application = (
        "<h2>Fungsi dan Penggunaan</h2>"
        f"<ul>{parameter_items}"
        "<li>Mendukung pemeriksaan yang konsisten ketika digunakan sesuai prosedur kerja.</li>"
        "<li>Pemilihan model perlu mempertimbangkan rentang, akurasi, konektor, sensor, dan kondisi sampel.</li>"
        "</ul>"
    )

    considerations = (
        "<h2>Pertimbangan Sebelum Memilih</h2>"
        "<p>Periksa spesifikasi, kompatibilitas dengan instrumen, kebutuhan kalibrasi, bahan habis pakai, "
        "serta isi paket sebelum pemesanan. Jangan menyimpulkan kompatibilitas hanya dari kemiripan nama atau konektor.</p>"
    )

    contact = (
        "<h2>Informasi Harga dan Pemesanan</h2>"
        f'<p>Untuk informasi harga, ketersediaan, dan konsultasi teknis {escape_html(source_sku)}, '
        f'hubungi Tim {escape_html(config["company_name"])} melalui WhatsApp '
        f'<a href="https://wa.me/{escape_html(config["whatsapp_number"])}">{escape_html(config["whatsapp_display"])}</a> '
        f'atau email <a href="mailto:{escape_html(config["email"])}">{escape_html(config["email"])}</a>.</p>'
    )
    return intro + application + considerations + contact


def build_advantages_html(name: str, parameters: Sequence[str], category: str, title: str) -> str:
    items: List[str] = []
    if parameters:
        items.append("Mendukung pengukuran atau analisis " + ", ".join(parameters) + ".")
    else:
        items.append(f"Dirancang untuk kebutuhan {category.lower()}.")
    lower = title.lower()
    if "waterproof" in lower:
        items.append("Desain tahan air untuk mendukung penggunaan pada kondisi kerja yang sesuai.")
    if "portable" in lower or "handheld" in lower:
        items.append("Bentuk portabel memudahkan penggunaan di lapangan.")
    if "wireless" in lower or "bluetooth" in lower:
        items.append("Mendukung konektivitas nirkabel pada perangkat yang kompatibel.")
    if "digital" in lower:
        items.append("Sistem digital membantu pembacaan dan pengelolaan hasil pengukuran.")
    items.append("Pemilihan model dapat dikonsultasikan dengan Tim UJI sebelum pemesanan.")
    li = "".join(f"<li>{escape_html(item)}</li>" for item in items[:6])
    return f"<h2>Keunggulan Produk</h2><ul>{li}</ul>"


def spec_value_by_patterns(spec_rows: Sequence[Tuple[str, str]], patterns: Sequence[str]) -> Optional[Tuple[str, str]]:
    for key, value in spec_rows:
        if any(re.search(pattern, key, flags=re.I) for pattern in patterns):
            return translate_spec_key(key), normalize_space(value)
    return None


def spec_values_by_patterns(
    spec_rows: Sequence[Tuple[str, str]],
    patterns: Sequence[str],
    limit: int = 4,
) -> List[Tuple[str, str]]:
    results: List[Tuple[str, str]] = []
    seen: set[Tuple[str, str]] = set()
    for key, value in spec_rows:
        if not any(re.search(pattern, key, flags=re.I) for pattern in patterns):
            continue
        pair = (translate_spec_key(key), normalize_space(value))
        marker = (pair[0].lower(), pair[1].lower())
        if not pair[1] or marker in seen:
            continue
        seen.add(marker)
        results.append(pair)
        if len(results) >= limit:
            break
    return results


def faq_product_family(category: str) -> str:
    lower = category.lower()
    if "larutan" in lower:
        return "solution"
    if "reagen" in lower:
        return "reagent"
    if "elektrode" in lower or "probe" in lower:
        return "electrode"
    if "aksesori" in lower or "suku cadang" in lower:
        return "accessory"
    return "instrument"


def join_indonesian(items: Sequence[str]) -> str:
    cleaned = [normalize_space(item) for item in items if normalize_space(item)]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return f"{cleaned[0]} dan {cleaned[1]}"
    return ", ".join(cleaned[:-1]) + f", dan {cleaned[-1]}"


def format_spec_pairs(pairs: Sequence[Tuple[str, str]]) -> str:
    if not pairs:
        return ""
    if len(pairs) == 1:
        key, value = pairs[0]
        return f"{key} tercantum {value}."
    details = "; ".join(f"{key}: {value}" for key, value in pairs)
    return f"Rincian yang tercantum meliputi {details}."


def build_function_faq_answer(
    name: str,
    title: str,
    parameters: Sequence[str],
    category: str,
) -> str:
    family = faq_product_family(category)
    lower = title.lower()
    parameter_text = join_indonesian(parameters)

    if family == "solution":
        if "storage" in lower or "penyimpanan" in lower:
            purpose = "menjaga kondisi elektrode selama tidak digunakan"
        elif "cleaning" in lower or "pembersih" in lower:
            purpose = "membantu membersihkan elektrode atau sensor sesuai jenis larutannya"
        elif "calibration" in lower or "buffer" in lower or "standard" in lower:
            purpose = "mendukung proses kalibrasi atau verifikasi instrumen"
        elif "filling" in lower or "electrolyte" in lower:
            purpose = "digunakan sebagai larutan pengisi atau elektrolit pada elektrode yang sesuai"
        else:
            purpose = "mendukung kalibrasi, perawatan, atau pengujian instrumen sesuai fungsi produknya"
        return f"{name} merupakan larutan yang digunakan untuk {purpose}. Penggunaannya perlu disesuaikan dengan jenis elektrode, instrumen, dan prosedur kerja yang diterapkan."

    if family == "reagent":
        target = f"parameter {parameter_text}" if parameter_text else "parameter pengujian yang sesuai"
        return f"{name} merupakan reagen untuk membantu analisis {target}. Gunakan reagen bersama metode dan instrumen yang kompatibel agar proses pengujian mengikuti prosedur yang ditetapkan."

    if family == "electrode":
        target = f"parameter {parameter_text}" if parameter_text else "parameter yang sesuai dengan tipe sensornya"
        return f"{name} merupakan elektrode, probe, atau sensor untuk mengukur {target}. Kesesuaian konektor, rentang kerja, material sensor, dan instrumen perlu diperiksa sebelum penggunaan."

    if family == "accessory":
        return f"{name} merupakan aksesori atau komponen pengganti untuk mendukung penggunaan instrumen yang sesuai. Periksa model, bentuk konektor, ukuran, dan kompatibilitas sebelum pemesanan."

    target = f"parameter {parameter_text}" if parameter_text else "parameter pengujian sesuai fungsi instrumennya"
    return f"{name} merupakan instrumen untuk mengukur atau menganalisis {target}. Pemilihan model perlu mempertimbangkan rentang, akurasi, resolusi, metode kalibrasi, dan kondisi sampel."


def append_spec_faq(
    items: List[Dict[str, str]],
    question: str,
    spec_rows: Sequence[Tuple[str, str]],
    patterns: Sequence[str],
    lead: str = "",
    limit: int = 4,
) -> None:
    pairs = spec_values_by_patterns(spec_rows, patterns, limit=limit)
    if not pairs:
        return
    answer = format_spec_pairs(pairs)
    if lead:
        answer = f"{lead} {answer}"
    items.append({"question": question, "answer": normalize_space(answer)})


def build_instrument_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi {name}?",
        "answer": build_function_faq_answer(name, title, parameters, category),
    }]
    if parameters:
        items.append({
            "question": f"Parameter apa saja yang dapat diukur oleh {source_sku}?",
            "answer": f"Parameter yang teridentifikasi pada produk ini meliputi {join_indonesian(parameters)}. Ketersediaan setiap parameter tetap perlu dicocokkan dengan probe, mode pengukuran, dan konfigurasi produk yang dipilih.",
        })
    append_spec_faq(items, f"Berapa rentang pengukuran {source_sku}?", spec_rows, (r"\brange\b",), limit=4)
    append_spec_faq(
        items,
        f"Bagaimana akurasi dan resolusi {source_sku}?",
        spec_rows,
        (r"\baccuracy\b", r"\bresolution\b"),
        limit=4,
    )
    append_spec_faq(items, f"Bagaimana sistem kalibrasi {source_sku}?", spec_rows, (r"\bcalibration\b",), limit=3)
    append_spec_faq(
        items,
        f"Untuk aplikasi apa {source_sku} direkomendasikan?",
        spec_rows,
        (r"\brecommended use\b", r"\bapplication\b", r"\buse environment\b"),
        limit=3,
    )
    return items


def build_electrode_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi {name}?",
        "answer": build_function_faq_answer(name, title, parameters, category),
    }]
    append_spec_faq(
        items,
        f"Berapa rentang kerja {source_sku}?",
        spec_rows,
        (r"\brange\b", r"\btemperature\b", r"\boperating environment\b"),
        limit=4,
    )
    append_spec_faq(
        items,
        f"Apa jenis konektor dan kabel {source_sku}?",
        spec_rows,
        (r"\bconnector\b", r"\bcable\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Bagaimana konstruksi sensor atau elektrode {source_sku}?",
        spec_rows,
        (r"\bbody material\b", r"\bbody\b", r"\bjunction\b", r"\breference\b", r"\belectrolyte\b", r"\bsensor type\b", r"\btip\b"),
        limit=4,
    )
    append_spec_faq(
        items,
        f"Untuk sampel atau aplikasi apa {source_sku} sesuai?",
        spec_rows,
        (r"\brecommended use\b", r"\bapplication\b", r"\bsample\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Apakah {source_sku} dapat diisi ulang atau memerlukan perawatan tertentu?",
        spec_rows,
        (r"\brefillable\b", r"\bfilling solution\b", r"\bmaintenance\b", r"\bstorage\b"),
        limit=3,
    )
    return items


def build_solution_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = [{
        "question": f"Apa kegunaan {name}?",
        "answer": build_function_faq_answer(name, title, parameters, category),
    }]
    append_spec_faq(
        items,
        f"Berapa volume atau ukuran kemasan {source_sku}?",
        spec_rows,
        (r"\bvolume\b", r"\bpackage\b", r"\bpackaging\b", r"\bsize\b", r"\bquantity\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Nilai atau konsentrasi apa yang dimiliki {source_sku}?",
        spec_rows,
        (r"\bvalue\b", r"\bconcentration\b", r"\bph value\b", r"\bconductivity value\b", r"\bstandard value\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Untuk instrumen atau penggunaan apa {source_sku} ditujukan?",
        spec_rows,
        (r"\brecommended use\b", r"\bapplication\b", r"\bcompatible\b", r"\bfor use with\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Bagaimana ketentuan penyimpanan {source_sku}?",
        spec_rows,
        (r"\bstorage\b", r"\bshelf life\b", r"\btemperature\b"),
        limit=3,
    )
    return items


def build_reagent_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi {name}?",
        "answer": build_function_faq_answer(name, title, parameters, category),
    }]
    if parameters:
        items.append({
            "question": f"Parameter apa yang diuji menggunakan {source_sku}?",
            "answer": f"Reagen ini ditujukan untuk pengujian {join_indonesian(parameters)}. Pastikan metode, rentang, dan instrumen yang digunakan sesuai dengan kebutuhan analisis.",
        })
    append_spec_faq(items, f"Metode pengujian apa yang digunakan {source_sku}?", spec_rows, (r"\bmethod\b",), limit=3)
    append_spec_faq(
        items,
        f"Berapa jumlah pengujian atau isi kemasan {source_sku}?",
        spec_rows,
        (r"\btests?\b", r"\bnumber of tests\b", r"\bpackage\b", r"\bquantity\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Instrumen apa yang kompatibel dengan {source_sku}?",
        spec_rows,
        (r"\bcompatible\b", r"\bfor use with\b", r"\binstrument\b", r"\bchecker\b", r"\bphotometer\b"),
        limit=3,
    )
    append_spec_faq(items, f"Berapa rentang pengujian {source_sku}?", spec_rows, (r"\brange\b",), limit=3)
    return items


def build_accessory_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi {name}?",
        "answer": build_function_faq_answer(name, title, parameters, category),
    }]
    append_spec_faq(
        items,
        f"Produk apa yang kompatibel dengan {source_sku}?",
        spec_rows,
        (r"\bcompatible\b", r"\bfor use with\b", r"\binstrument\b", r"\bmodel\b"),
        limit=4,
    )
    append_spec_faq(
        items,
        f"Apa ukuran atau material {source_sku}?",
        spec_rows,
        (r"\bdimensions?\b", r"\bsize\b", r"\bmaterial\b", r"\bweight\b", r"\blength\b"),
        limit=4,
    )
    append_spec_faq(
        items,
        f"Apa isi kemasan {source_sku}?",
        spec_rows,
        (r"\bpackage\b", r"\bquantity\b", r"\bincluded\b", r"\bcontents?\b"),
        limit=3,
    )
    append_spec_faq(
        items,
        f"Apa jenis konektor atau sambungan {source_sku}?",
        spec_rows,
        (r"\bconnector\b", r"\bcable\b", r"\bthread\b", r"\bconnection\b"),
        limit=3,
    )
    return items


def ensure_minimum_faq(
    items: List[Dict[str, str]],
    source_sku: str,
    category: str,
) -> List[Dict[str, str]]:
    if len(items) < 3:
        items.append({
            "question": f"Apa yang perlu diperiksa sebelum memilih {source_sku}?",
            "answer": f"Periksa fungsi produk, kategori {category.lower()}, rentang kerja, kompatibilitas, konektor, bahan habis pakai, dan isi kemasan yang tersedia. Konfirmasi teknis diperlukan ketika data produk belum mencantumkan rincian tersebut.",
        })
    if len(items) < 3:
        items.append({
            "question": f"Apakah {source_sku} harus disesuaikan dengan instrumen tertentu?",
            "answer": "Ya. Kesesuaian model, konektor, sensor, metode, dan aksesori perlu diperiksa sebelum penggunaan. Jangan menentukan kompatibilitas hanya dari kemiripan nama atau bentuk produk.",
        })
    return items[:5]


def build_template_faq_items(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
) -> List[Dict[str, str]]:
    family = faq_product_family(category)
    if family == "solution":
        items = build_solution_faq_items(name, source_sku, title, parameters, category, spec_rows)
    elif family == "reagent":
        items = build_reagent_faq_items(name, source_sku, title, parameters, category, spec_rows)
    elif family == "electrode":
        items = build_electrode_faq_items(name, source_sku, title, parameters, category, spec_rows)
    elif family == "accessory":
        items = build_accessory_faq_items(name, source_sku, title, parameters, category, spec_rows)
    else:
        items = build_instrument_faq_items(name, source_sku, title, parameters, category, spec_rows)
    return ensure_minimum_faq(items, source_sku, category)


def faq_items_to_html(items: Sequence[Dict[str, str]]) -> str:
    return "".join(
        f"<h3>{escape_html(item['question'])}</h3><p>{escape_html(item['answer'])}</p>"
        for item in items
        if item.get("question") and item.get("answer")
    )


def faq_polish_cache_key(
    name: str,
    source_sku: str,
    items: Sequence[Dict[str, str]],
    config: Dict[str, object],
) -> str:
    payload = {
        "prompt_version": config.get("deepseek_prompt_version"),
        "model": config.get("deepseek_model"),
        "name": name,
        "sku": source_sku,
        "faq": list(items),
    }
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def load_faq_cache(path: Path) -> Dict[str, object]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        logging.warning("Cache FAQ tidak dapat dibaca dan akan dibuat ulang: %s", path)
        return {}


def save_faq_cache(path: Path, cache: Dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def normalized_number_tokens(text: str) -> List[str]:
    tokens: List[str] = []
    for token in re.findall(r"(?<![A-Za-z])[-+]?\d+(?:[.,]\d+)?", text):
        normalized = token.replace(",", ".")
        try:
            number = float(normalized)
            tokens.append(f"{number:.12g}")
        except ValueError:
            tokens.append(normalized)
    return sorted(tokens)


def validation_text(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    return re.sub(r"\s+", " ", text).strip()


def extract_protected_terms(
    original: str,
    source_sku: str,
    parameters: Sequence[str],
) -> List[str]:
    terms: List[str] = []
    if source_sku and source_sku.lower() in original.lower():
        terms.append(source_sku)
    for parameter in parameters:
        if parameter and validation_text(parameter) in validation_text(original):
            terms.append(parameter)
    technical_patterns = (
        r"\bpH\b", r"\bORP\b", r"\bEC\b", r"\bTDS\b", r"\bNTU\b", r"\bFNU\b",
        r"\bppm\b", r"\bppt\b", r"\bPSU\b", r"\bmg/L\b", r"\bµg/L\b", r"\bμg/L\b",
        r"\bµS/cm\b", r"\bμS/cm\b", r"\bmS/cm\b", r"°C", r"°F", r"\bmV\b",
        r"\bmm\b", r"\bcm\b", r"\bkg\b", r"\bg\b", r"\bmL\b", r"\bL\b",
    )
    for pattern in technical_patterns:
        for match in re.findall(pattern, original, flags=re.I):
            value = match if isinstance(match, str) else "".join(match)
            if value:
                terms.append(value)
    unique: List[str] = []
    for term in terms:
        if term and validation_text(term) not in {validation_text(item) for item in unique}:
            unique.append(term)
    return unique


def clean_polished_answer(text: str) -> str:
    value = html.unescape(str(text or "")).strip()
    value = re.sub(r"^```(?:text|markdown)?\s*|\s*```$", "", value, flags=re.I | re.S)
    value = re.sub(r"^\s*(?:jawaban|answer)\s*[:\-]\s*", "", value, flags=re.I)
    value = value.strip().strip('"').strip()
    return normalize_space(value)


def validate_polished_answer(
    original: str,
    polished: str,
    source_sku: str,
    parameters: Sequence[str],
) -> str:
    polished = clean_polished_answer(polished)
    if not polished:
        raise ValueError("Hasil penyuntingan AI kosong.")
    word_count = len(polished.split())
    if word_count < 5 or word_count > 120:
        raise ValueError("Panjang hasil penyuntingan AI tidak wajar.")

    forbidden = (
        "hannainst.com", "website hanna", "sumber hanna", "spesifikasi resmi",
        "berdasarkan sumber", "harga terbaru", "stok tersedia", "garansi resmi",
        "pasti kompatibel", "tersedia sekarang", "menurut hanna",
    )
    lower = polished.lower()
    if any(term in lower for term in forbidden):
        raise ValueError("Hasil penyuntingan AI memuat klaim atau rujukan yang dilarang.")

    if normalized_number_tokens(polished) != normalized_number_tokens(original):
        raise ValueError("AI mengubah, menambah, atau menghapus angka teknis.")

    polished_validation = validation_text(polished)
    for term in extract_protected_terms(original, source_sku, parameters):
        if validation_text(term) not in polished_validation:
            raise ValueError(f"AI menghapus istilah teknis wajib: {term}")
    return polished


def parse_polished_blocks(content: str, expected_count: int) -> Dict[int, str]:
    results: Dict[int, str] = {}
    for number, answer in re.findall(r"<A(\d+)>(.*?)</A\1>", content, flags=re.I | re.S):
        index = int(number)
        if 1 <= index <= expected_count:
            results[index] = clean_polished_answer(answer)
    return results


class DeepSeekFAQPolisher:
    def __init__(self, config: Dict[str, object], api_key: str, enabled: bool) -> None:
        self.config = config
        self.api_key = api_key.strip()
        self.enabled = enabled and bool(self.api_key)
        self.session = requests.Session()
        self.cache_path = Path(str(config.get("deepseek_cache_file", "output/deepseek_faq_polish_cache_v4.json")))
        self.cache = load_faq_cache(self.cache_path)
        self.dirty_count = 0

    def flush(self) -> None:
        if self.dirty_count:
            save_faq_cache(self.cache_path, self.cache)
            self.dirty_count = 0

    def polish(
        self,
        name: str,
        source_sku: str,
        items: Sequence[Dict[str, str]],
        parameters: Sequence[str],
    ) -> Tuple[List[Dict[str, str]], str]:
        if not self.enabled:
            raise ImporterError("Penyuntingan FAQ melalui DeepSeek tidak aktif atau API key belum tersedia.")

        key = faq_polish_cache_key(name, source_sku, items, self.config)
        cached = self.cache.get(key)
        if isinstance(cached, dict) and isinstance(cached.get("faq"), list):
            cached_items = cached["faq"]
            if len(cached_items) == len(items):
                validated: List[Dict[str, str]] = []
                for original, cached_item in zip(items, cached_items):
                    answer = validate_polished_answer(
                        original["answer"],
                        str(cached_item.get("answer", "")),
                        source_sku,
                        parameters,
                    )
                    validated.append({"question": original["question"], "answer": answer})
                return validated, "AI_POLISH_CACHE"

        input_blocks = []
        output_example = []
        for index, item in enumerate(items, start=1):
            input_blocks.append(
                f"<FAQ{index}>\n<QUESTION>{item['question']}</QUESTION>\n"
                f"<ORIGINAL_ANSWER>{item['answer']}</ORIGINAL_ANSWER>\n</FAQ{index}>"
            )
            output_example.append(f"<A{index}>Jawaban yang sudah diperhalus tanpa fakta baru.</A{index}>")

        system_prompt = (
            "Anda adalah editor bahasa Indonesia untuk katalog alat ukur dan laboratorium. "
            "Tugas Anda hanya memperhalus susunan kalimat jawaban FAQ yang sudah faktual. "
            "Jangan menambah, menghapus, menyimpulkan, menerjemahkan, atau mengubah fakta. "
            "Pertahankan semua angka, tanda positif atau negatif, rentang, satuan, SKU, model, parameter, "
            "jenis konektor, material, metode, kompatibilitas, dan istilah teknis. "
            "Jangan menyebut sumber, website produsen, spesifikasi resmi, harga, stok, garansi, atau sertifikasi. "
            "Jangan mengubah pertanyaan. Keluarkan hanya blok jawaban sesuai format yang diminta, tanpa JSON dan tanpa markdown."
        )
        user_prompt = (
            "Perhalus setiap ORIGINAL_ANSWER agar lebih natural, jelas, ringkas, dan tetap profesional. "
            "Jangan menambah fakta baru. Gunakan format keluaran berikut secara persis:\n"
            + "\n".join(output_example)
            + "\n\nDATA FAQ:\n"
            + "\n".join(input_blocks)
        )

        endpoint = str(self.config.get("deepseek_base_url", "https://api.deepseek.com")).rstrip("/") + "/chat/completions"
        payload = {
            "model": str(self.config.get("deepseek_model", "deepseek-v4-flash")),
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "thinking": {"type": "disabled"},
            "temperature": float(self.config.get("deepseek_temperature", 0.1)),
            "max_tokens": int(self.config.get("deepseek_max_tokens", 1000)),
            "stream": False,
        }
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last_error: Optional[Exception] = None

        for attempt in range(1, int(self.config.get("deepseek_max_retries", 2)) + 1):
            try:
                attempt_payload = dict(payload)
                attempt_payload["messages"] = list(payload["messages"])
                if attempt > 1 and last_error is not None:
                    attempt_payload["messages"].append({
                        "role": "user",
                        "content": (
                            "Revisi jawaban sebelumnya. Validator menolak hasil karena: "
                            f"{last_error}. Hapus atau perbaiki hanya bagian yang tidak didukung evidence. "
                            "Jangan menambah fakta baru. Tetap keluarkan satu objek JSON valid tanpa markdown."
                        ),
                    })
                response = self.session.post(
                    endpoint,
                    headers=headers,
                    json=attempt_payload,
                    timeout=int(self.config.get("deepseek_timeout_seconds", 90)),
                )
                response.raise_for_status()
                body = response.json()
                content = str(body["choices"][0]["message"]["content"] or "").strip()
                if not content:
                    raise ValueError("DeepSeek mengembalikan konten kosong.")
                blocks = parse_polished_blocks(content, len(items))
                if not blocks:
                    raise ValueError("Format blok jawaban DeepSeek tidak terbaca.")

                final_items: List[Dict[str, str]] = []
                polished_count = 0
                for index, original in enumerate(items, start=1):
                    candidate = blocks.get(index, "")
                    if candidate:
                        try:
                            answer = validate_polished_answer(
                                original["answer"], candidate, source_sku, parameters
                            )
                            polished_count += 1
                        except ValueError as exc:
                            logging.warning(
                                "Jawaban FAQ %s nomor %s memakai template karena validasi gagal: %s",
                                source_sku, index, exc,
                            )
                            answer = original["answer"]
                    else:
                        answer = original["answer"]
                    final_items.append({"question": original["question"], "answer": answer})

                if polished_count == 0:
                    raise ValueError("Tidak ada jawaban AI yang lolos validasi.")

                source = "AI_POLISH" if polished_count == len(items) else "AI_POLISH_PARTIAL"
                self.cache[key] = {
                    "faq": final_items,
                    "model": body.get("model", payload["model"]),
                    "sku": source_sku,
                    "source": source,
                }
                self.dirty_count += 1
                if self.dirty_count >= 10:
                    self.flush()
                time.sleep(float(self.config.get("deepseek_delay_seconds", 0.35)))
                return final_items, source
            except (requests.RequestException, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                last_error = exc
                logging.warning(
                    "Penyuntingan FAQ DeepSeek gagal untuk %s, percobaan %s: %s",
                    source_sku, attempt, exc,
                )
                time.sleep(min(attempt * 2, 6))

        raise ImporterError(f"Penyuntingan FAQ DeepSeek gagal setelah beberapa percobaan: {last_error}")


def extract_json_object(content: str) -> Dict[str, object]:
    cleaned = html.unescape(str(content or "")).strip()
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.I | re.S).strip()
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("Objek JSON tidak ditemukan pada hasil AI.")
    payload = json.loads(cleaned[start:end + 1])
    if not isinstance(payload, dict):
        raise ValueError("Hasil AI bukan objek JSON.")
    return payload


def product_content_cache_key(
    name: str,
    source_sku: str,
    evidence: str,
    config: Dict[str, object],
) -> str:
    payload = {
        "prompt_version": config.get("deepseek_content_prompt_version"),
        "model": config.get("deepseek_model"),
        "name": name,
        "sku": source_sku,
        "evidence_hash": hashlib.sha256(evidence.encode("utf-8")).hexdigest(),
    }
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def validate_product_content(
    payload: Dict[str, object],
    evidence: str,
    source_sku: str,
    category: str,
) -> Dict[str, object]:
    """Validasi konten AI sesuai section yang benar-benar tersedia pada source."""
    match = re.search(r"AVAILABLE_SECTIONS:\s*([^\n]+)", evidence, re.I)
    available = set()
    if match:
        available = {normalize_space(item).upper() for item in match.group(1).split(",") if normalize_space(item) and normalize_space(item).upper() != "NONE"}
    faq_eligible = bool(re.search(r"FAQ_ELIGIBLE:\s*YES", evidence, re.I))

    short_description = clean_generated_text(payload.get("short_description"), max_words=90)
    overview = clean_description_for_formatting(payload.get("overview"))
    working_principle = clean_generated_text(payload.get("working_principle"), max_words=180)
    considerations = clean_generated_text(payload.get("considerations"), max_words=140)
    has_overview = "OVERVIEW" in available
    if has_overview and (not short_description or not overview):
        raise ValueError("Source memiliki OVERVIEW tetapi deskripsi AI kosong.")
    if not has_overview and (short_description or overview):
        raise ValueError("AI membuat deskripsi padahal source tidak memiliki OVERVIEW/DESCRIPTION.")

    # Jika source mempunyai description, angka/model teknis yang memang tertulis
    # di OVERVIEW tidak boleh hilang saat AI mengolah Product Description.
    if has_overview:
        overview_match = re.search(
            r"(?:^|\n\n)OVERVIEW:\n(?P<body>.*?)(?=\n\n(?:FEATURES|WORKING PRINCIPLE|APPLICATION|CALIBRATION / STANDARD / GLP|REPORTING / CONNECTION / DATA|INSTALLATION|WHAT'S INCLUDED|MAINTENANCE|COMPATIBILITY|SPECIFICATION \(LOCKED FACTS\)):\n|\Z)",
            evidence,
            flags=re.I | re.S,
        )
        if overview_match:
            source_overview = overview_match.group("body")
            missing_numbers = sorted(set(normalized_number_tokens(source_overview)) - set(normalized_number_tokens(overview)))
            if missing_numbers:
                raise ValueError("AI menghilangkan angka teknis dari source description: " + ", ".join(missing_numbers[:10]))
            source_tokens = extract_upper_technical_tokens(source_overview)
            output_tokens_overview = extract_upper_technical_tokens(overview)
            missing_tokens = sorted(token for token in source_tokens if token not in output_tokens_overview)
            if missing_tokens:
                raise ValueError("AI menghilangkan model/istilah teknis dari source description: " + ", ".join(missing_tokens[:8]))

    feature_items: List[Dict[str, str]] = []
    raw_features = payload.get("feature_items")
    if isinstance(raw_features, list):
        for item in raw_features:
            if not isinstance(item, dict):
                continue
            title = clean_generated_text(item.get("title"), max_words=18)
            detail = clean_generated_text(item.get("detail"), max_words=90)
            if title and detail:
                feature_items.append({"title": title, "detail": detail})

    def clean_list(key: str, max_words: int, max_items: int = 10) -> List[str]:
        raw = payload.get(key)
        if not isinstance(raw, list):
            return []
        result = [clean_generated_text(item, max_words=max_words) for item in raw]
        return [item for item in result if item][:max_items]

    application_items = clean_list("application_items", 65, 8)
    calibration_data_items = clean_list("calibration_data_items", 75, 8)
    advantages = clean_list("advantages", 80, 10)

    faq: List[Dict[str, str]] = []
    raw_faq = payload.get("faq")
    if isinstance(raw_faq, list):
        for item in raw_faq:
            if not isinstance(item, dict):
                continue
            question = clean_generated_text(item.get("question"), max_words=35)
            answer = clean_generated_text(item.get("answer"), max_words=135)
            if question and answer:
                faq.append({"question": question, "answer": answer})

    # Source-adaptive hard rules. Jangan membuat section hanya untuk mengisi template.
    if "FEATURES" not in available and (feature_items or advantages):
        raise ValueError("AI membuat fitur/keunggulan padahal source tidak memiliki section Features.")
    if "WORKING_PRINCIPLE" not in available and working_principle:
        raise ValueError("AI membuat prinsip kerja tanpa evidence khusus.")
    if "APPLICATION" not in available and application_items:
        raise ValueError("AI membuat aplikasi tanpa evidence khusus.")
    if not ({"CALIBRATION", "CONNECTIVITY"} & available) and calibration_data_items:
        raise ValueError("AI membuat kalibrasi/data tanpa evidence khusus.")
    if not faq_eligible and faq:
        raise ValueError("AI membuat FAQ untuk source yang tidak memenuhi syarat FAQ.")

    normalized = {
        "short_description": short_description,
        "overview": overview,
        "feature_items": feature_items[:10],
        "working_principle": working_principle,
        "application_items": application_items,
        "calibration_data_items": calibration_data_items,
        "considerations": considerations,
        "advantages": advantages,
        "faq": faq[:10],
    }
    output_text = json.dumps(normalized, ensure_ascii=False)
    forbidden = (
        "hannainst.com", "hanna website", "website hanna", "situs hanna", "kunjungi hanna",
        "menurut hanna", "menurut produsen", "berdasarkan sumber", "berdasarkan spesifikasi",
        "berdasarkan dokumentasi", "spesifikasi resmi", "sumber resmi", "website resmi",
        "situs resmi", "dokumentasi resmi", "manufacturer website", "official website",
        "harga terbaru", "stok tersedia", "diskon", "promo", "add to cart",
        "kunjungi website", "lihat website", "cek website", "shop now",
    )
    lower = output_text.lower()
    if any(term in lower for term in forbidden):
        raise ValueError("Hasil AI memuat rujukan sumber atau klaim promosi yang dilarang.")

    bad_numbers = unsupported_number_tokens(output_text, evidence)
    if bad_numbers:
        raise ValueError("AI menambahkan angka yang tidak tersedia pada bukti produk: " + ", ".join(bad_numbers[:12]))

    allowed_tokens = {"UJI", "FAQ", "PC"}
    output_tokens = extract_upper_technical_tokens(output_text) - allowed_tokens
    evidence_tokens = extract_upper_technical_tokens(evidence) | {source_sku.upper()}

    def token_supported(token: str) -> bool:
        if token in evidence_tokens:
            return True
        parts = [part for part in re.split(r"[/+&]", token) if part]
        return len(parts) > 1 and all(part in evidence_tokens for part in parts)

    unsupported = sorted(token for token in output_tokens if not token_supported(token))
    if unsupported:
        raise ValueError("AI menambahkan istilah teknis yang tidak didukung: " + ", ".join(unsupported[:8]))
    return normalized

class DeepSeekEvidenceTranslator:
    """Terjemahkan narrative evidence sebelum product writer.

    Angka, unit, model, dan token teknis uppercase divalidasi agar tidak berubah.
    Jika penerjemahan gagal, fallback hanya memakai kalimat yang dapat dilokalkan
    dengan aman; prose Inggris yang tersisa tidak diteruskan sebagai konten publik.
    """
    def __init__(self, config: Dict[str, object], api_key: str, enabled: bool) -> None:
        self.config = config
        self.api_key = api_key.strip()
        self.enabled = enabled and bool(self.api_key)
        self.session = requests.Session()
        self.cache_path = Path(str(config.get("deepseek_translation_cache_file", "output/deepseek_translation_cache_v4_3.json")))
        self.cache = load_faq_cache(self.cache_path)
        self.dirty_count = 0

    def flush(self) -> None:
        if self.dirty_count:
            save_faq_cache(self.cache_path, self.cache)
            self.dirty_count = 0

    def _cache_key(self, source_sku: str, buckets: Dict[str, List[str]]) -> str:
        payload = {
            "prompt_version": self.config.get("deepseek_translation_prompt_version"),
            "model": self.config.get("deepseek_model"),
            "sku": source_sku,
            "buckets": buckets,
        }
        return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()

    def _validate(self, original: Dict[str, List[str]], translated: Dict[str, object]) -> Dict[str, List[str]]:
        result: Dict[str, List[str]] = {}
        original_text = json.dumps(original, ensure_ascii=False)
        translated_text = json.dumps(translated, ensure_ascii=False)
        if normalized_number_tokens(original_text) != normalized_number_tokens(translated_text):
            raise ValueError("Translator mengubah angka teknis.")
        required_tokens = extract_upper_technical_tokens(original_text)
        translated_tokens = extract_upper_technical_tokens(translated_text)
        missing = sorted(token for token in required_tokens if token not in translated_tokens)
        if missing:
            raise ValueError("Translator menghapus token teknis: " + ", ".join(missing[:8]))

        for kind, source_lines in original.items():
            raw_lines = translated.get(kind, [])
            if not isinstance(raw_lines, list) or len(raw_lines) != len(source_lines):
                raise ValueError(f"Jumlah baris terjemahan tidak sesuai untuk {kind}.")
            cleaned_lines: List[str] = []
            for line in raw_lines:
                cleaned = normalize_space(line)
                if cleaned:
                    cleaned_lines.append(cleaned)
            if len(cleaned_lines) != len(source_lines):
                raise ValueError(f"Terjemahan kosong pada section {kind}.")
            result[kind] = cleaned_lines
        return result

    def translate(self, source_sku: str, buckets: Dict[str, List[str]]) -> Tuple[Dict[str, List[str]], str]:
        if not buckets:
            return {}, "NO_SOURCE_TEXT"
        if not self.enabled:
            return self.safe_fallback(buckets), "TRANSLATION_TEMPLATE"

        key = self._cache_key(source_sku, buckets)
        cached = self.cache.get(key)
        if isinstance(cached, dict) and isinstance(cached.get("buckets"), dict):
            try:
                return self._validate(buckets, cached["buckets"]), "TRANSLATION_CACHE"
            except ValueError:
                pass

        system_prompt = (
            "Anda adalah penerjemah teknis untuk katalog alat ukur dan laboratorium. "
            "Terjemahkan hanya prose bahasa Inggris ke bahasa Indonesia yang natural. "
            "Pertahankan model, SKU, angka, tanda plus/minus, rentang, satuan, simbol, chemical formula, "
            "nama metode, nama software, port, standar, singkatan teknis, dan proper technical terms persis seperti input. "
            "Istilah teknis Inggris boleh dipertahankan bila lebih lazim, tetapi struktur kalimat utama harus bahasa Indonesia. "
            "Jangan menambah atau menghapus fakta. Jangan menyebut website sumber. Keluarkan satu objek JSON valid tanpa markdown."
        )
        user_prompt = (
            "Terjemahkan nilai pada setiap array. Jangan mengubah nama key dan jangan mengubah jumlah item setiap array.\n"
            + json.dumps(buckets, ensure_ascii=False, indent=2)
        )
        endpoint = str(self.config.get("deepseek_base_url", "https://api.deepseek.com")).rstrip("/") + "/chat/completions"
        payload = {
            "model": str(self.config.get("deepseek_model", "deepseek-v4-flash")),
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}],
            "thinking": {"type": "disabled"},
            "temperature": float(self.config.get("deepseek_translation_temperature", 0.0)),
            "max_tokens": int(self.config.get("deepseek_translation_max_tokens", 6000)),
            "stream": False,
        }
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last_error: Optional[Exception] = None
        for attempt in range(1, int(self.config.get("deepseek_max_retries", 2)) + 1):
            try:
                response = self.session.post(endpoint, headers=headers, json=payload, timeout=int(self.config.get("deepseek_timeout_seconds", 90)))
                response.raise_for_status()
                body = response.json()
                parsed = extract_json_object(str(body["choices"][0]["message"]["content"] or ""))
                validated = self._validate(buckets, parsed)
                self.cache[key] = {"buckets": validated, "sku": source_sku, "source": "TRANSLATION_AI"}
                self.dirty_count += 1
                if self.dirty_count >= 8:
                    self.flush()
                time.sleep(float(self.config.get("deepseek_delay_seconds", 0.35)))
                return validated, "TRANSLATION_AI"
            except (requests.RequestException, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                last_error = exc
                logging.warning("Terjemahan evidence gagal untuk %s percobaan %s: %s", source_sku, attempt, exc)
                time.sleep(min(attempt * 2, 6))
        logging.warning("Terjemahan %s memakai fallback aman: %s", source_sku, last_error)
        return self.safe_fallback(buckets), "TRANSLATION_FALLBACK"

    @staticmethod
    def safe_fallback(buckets: Dict[str, List[str]]) -> Dict[str, List[str]]:
        result: Dict[str, List[str]] = {}
        for kind, lines in buckets.items():
            localized: List[str] = []
            for line in lines:
                candidate = localize_source_line(line)
                if not candidate:
                    continue
                # Jangan meneruskan prose Inggris rusak ke output. Technical labels singkat tetap boleh.
                if contains_untranslated_english_prose(candidate):
                    continue
                localized.append(candidate)
            if localized:
                result[kind] = localized
        return result


class DeepSeekProductWriter:
    def __init__(self, config: Dict[str, object], api_key: str, enabled: bool) -> None:
        self.config = config
        self.api_key = api_key.strip()
        self.enabled = enabled and bool(self.api_key)
        self.session = requests.Session()
        self.cache_path = Path(str(config.get("deepseek_content_cache_file", "output/deepseek_product_content_cache_v4_3.json")))
        self.cache = load_faq_cache(self.cache_path)
        self.dirty_count = 0

    def flush(self) -> None:
        if self.dirty_count:
            save_faq_cache(self.cache_path, self.cache)
            self.dirty_count = 0

    def generate(
        self,
        name: str,
        source_sku: str,
        category: str,
        parameters: Sequence[str],
        evidence: str,
    ) -> Tuple[Dict[str, object], str]:
        if not self.enabled:
            raise ImporterError("Generator konten DeepSeek tidak aktif atau API key belum tersedia.")

        key = product_content_cache_key(name, source_sku, evidence, self.config)
        cached = self.cache.get(key)
        if isinstance(cached, dict) and isinstance(cached.get("content"), dict):
            try:
                validated = validate_product_content(cached["content"], evidence, source_sku, category)
                return validated, "AI_RICH_CACHE"
            except ValueError as exc:
                logging.warning("Cache konten %s tidak lolos validator V4.3.2 dan akan dibuat ulang: %s", source_sku, exc)
                self.cache.pop(key, None)
                self.dirty_count += 1

        system_prompt = (
            "Anda adalah technical product copywriter bahasa Indonesia untuk katalog UJI. "
            "Gunakan hanya EVIDENCE. Jangan memakai pengetahuan luar, jangan mengasumsikan kompatibilitas, aplikasi, manfaat, atau spesifikasi. "
            "EVIDENCE sudah diterjemahkan ke bahasa Indonesia bila memungkinkan. Pertahankan model, angka, rentang, satuan, simbol, standar, "
            "nama metode, software, port, dan istilah teknis persis seperti bukti. Bahasa utama harus Indonesia; istilah teknis Inggris boleh dipertahankan. "
            "Jangan pernah menyebut hannainst.com, website Hanna, website sumber, produsen sebagai sumber, atau mengarahkan pembaca ke website brand. "
            "Jangan membuat konten untuk section yang tidak tercantum pada AVAILABLE_SECTIONS. Jika data source ringkas, hasil juga harus ringkas. "
            "Keluarkan satu objek JSON valid tanpa markdown."
        )
        user_prompt = f"""Susun konten katalog UJI secara SOURCE-ADAPTIVE.

ATURAN WAJIB:
1. short_description: HANYA jika AVAILABLE_SECTIONS memuat OVERVIEW. Ringkas isi overview menjadi 1-2 kalimat. Jika OVERVIEW tidak tersedia, isi string kosong.
2. overview: HANYA olah fakta dari section OVERVIEW/DESCRIPTION source. OVERVIEW mencakup seluruh body Description asli, termasuk heading inline seperti Features at-a-glance jika heading itu memang berada di dalam Description. Terjemahkan dan parafrase ke bahasa Indonesia yang natural, pertahankan fakta teknis penting, angka, model, rentang, sensor, dan konteks penggunaan yang memang ada. Jangan mengambil SPECIFICATION atau DEDICATED SOURCE TAB untuk menambah deskripsi. Jika OVERVIEW tidak tersedia, isi string kosong.
3. feature_items: HANYA jika AVAILABLE_SECTIONS memuat FEATURES. Jangan mengambil Range/Accuracy/Resolution dari SPECIFICATION sebagai fitur. Jika tidak ada FEATURES, isi [].
4. advantages: HANYA dari FEATURES/benefit source. Jika tidak ada FEATURES, isi [].
5. working_principle: hanya jika WORKING_PRINCIPLE tersedia. Jika tidak, isi string kosong.
6. application_items: hanya jika APPLICATION tersedia. Jika tidak, isi [].
7. calibration_data_items: hanya jika CALIBRATION atau CONNECTIVITY tersedia. Jika tidak, isi [].
8. considerations: hanya jika source mempunyai informasi kompatibilitas/perawatan/penggunaan yang relevan. Jika tidak, isi string kosong.
9. faq: hanya jika FAQ_ELIGIBLE: YES. Jika NO, isi []. Jika YES, buat 3-6 FAQ yang benar-benar didukung evidence.
10. Jangan mengubah atau mengonversi angka teknis. Jangan menambah fakta.
11. Jangan menulis filler seperti 'hasil lebih akurat', 'mudah digunakan', atau 'cocok untuk berbagai aplikasi' kecuali source benar-benar menyatakannya.
12. Jangan menulis frasa seperti 'menurut Hanna', 'berdasarkan sumber', 'berdasarkan spesifikasi', 'website resmi', atau CTA ke Hanna.
13. Product Description dan custom tab UJI harus terasa lengkap hanya sejauh source memang lengkap. Jangan memperpanjang Product Description dengan filler.
14. Jika source description hanya satu paragraf, olah paragraf itu menjadi satu paragraf Indonesia yang utuh. Jika source description panjang, tetap prioritaskan fakta yang ada tanpa mengarang tambahan.
15. Jangan menentukan custom tab dari kata/heading yang muncul di dalam OVERVIEW. Heading inline seperti Features at-a-glance tetap bagian Product Description. Custom tab hanya boleh berasal dari AVAILABLE_SECTIONS yang datang dari dedicated source tab/panel dan dari Specification Lock.

FORMAT JSON:
{{
  "short_description": "...",
  "overview": "...",
  "feature_items": [{{"title": "...", "detail": "..."}}],
  "working_principle": "...",
  "application_items": ["..."],
  "calibration_data_items": ["..."],
  "considerations": "...",
  "advantages": ["..."],
  "faq": [{{"question": "...", "answer": "..."}}]
}}

EVIDENCE:
{evidence}
"""
        endpoint = str(self.config.get("deepseek_base_url", "https://api.deepseek.com")).rstrip("/") + "/chat/completions"
        payload = {
            "model": str(self.config.get("deepseek_model", "deepseek-v4-flash")),
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "thinking": {"type": "disabled"},
            "temperature": float(self.config.get("deepseek_content_temperature", 0.1)),
            "max_tokens": int(self.config.get("deepseek_content_max_tokens", 3200)),
            "stream": False,
        }
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last_error: Optional[Exception] = None
        for attempt in range(1, int(self.config.get("deepseek_max_retries", 2)) + 1):
            try:
                response = self.session.post(
                    endpoint,
                    headers=headers,
                    json=payload,
                    timeout=int(self.config.get("deepseek_timeout_seconds", 90)),
                )
                response.raise_for_status()
                body = response.json()
                raw_content = str(body["choices"][0]["message"]["content"] or "").strip()
                parsed = extract_json_object(raw_content)
                validated = validate_product_content(parsed, evidence, source_sku, category)
                self.cache[key] = {
                    "content": validated,
                    "model": body.get("model", payload["model"]),
                    "sku": source_sku,
                    "source": "AI_RICH",
                }
                self.dirty_count += 1
                if self.dirty_count >= 5:
                    self.flush()
                time.sleep(float(self.config.get("deepseek_delay_seconds", 0.35)))
                return validated, "AI_RICH"
            except (requests.RequestException, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                last_error = exc
                logging.warning(
                    "Generator konten DeepSeek gagal untuk %s, percobaan %s: %s",
                    source_sku, attempt, exc,
                )
                time.sleep(min(attempt * 2, 6))
        raise ImporterError(f"Generator konten DeepSeek gagal setelah beberapa percobaan: {last_error}")

def build_faq_html(
    name: str,
    source_sku: str,
    title: str,
    parameters: Sequence[str],
    category: str,
    spec_rows: Sequence[Tuple[str, str]],
    source_short: str,
    source_description: str,
    faq_polisher: Optional[DeepSeekFAQPolisher],
    require_polish: bool,
) -> Tuple[str, str, Optional[str]]:
    del source_short, source_description  # FAQ dibangun dari data terstruktur, bukan kalimat sumber mentah.
    template_items = build_template_faq_items(name, source_sku, title, parameters, category, spec_rows)
    if faq_polisher and faq_polisher.enabled:
        try:
            items, source = faq_polisher.polish(name, source_sku, template_items, parameters)
            return faq_items_to_html(items), source, None
        except ImporterError as exc:
            logging.warning("FAQ %s tetap memakai template karena AI gagal: %s", source_sku, exc)
            return faq_items_to_html(template_items), "TEMPLATE_FALLBACK", None
    issue = "REVIEW_FAQ_POLISH" if require_polish else None
    return faq_items_to_html(template_items), "TEMPLATE", issue

def build_tags(source_sku: str, category: str, parameters: Sequence[str]) -> str:
    tags = ["Hanna", "Hanna Instruments", source_sku, category, "instrumen pengujian", "alat laboratorium"]
    tags.extend(parameters)
    unique: List[str] = []
    for tag in tags:
        tag = normalize_space(tag)
        if tag and tag.lower() not in {item.lower() for item in unique}:
            unique.append(tag)
    return ", ".join(unique[:15])


def truncate_meta(text: str, limit: int = 155) -> str:
    text = normalize_space(text)
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:") + "."


def build_meta_description(
    source_sku: str,
    category: str,
    parameters: Sequence[str],
    name: str = "",
) -> str:
    """Buat meta description sebagai kalimat utuh, bukan truncate short description."""
    parameter_phrase = ""
    if parameters:
        if len(parameters) == 1:
            parameter_phrase = parameters[0]
        elif len(parameters) == 2:
            parameter_phrase = f"{parameters[0]} dan {parameters[1]}"
        else:
            parameter_phrase = ", ".join(parameters[:2]) + f", dan {parameters[2]}"

    product_label = normalize_space(name)
    product_label = re.sub(rf"^\s*Hanna\s+{re.escape(source_sku)}\s*", "", product_label, flags=re.I).strip()
    if product_label:
        product_label = smart_truncate(product_label, 72)

    category_phrase = translate_product_type("", category)
    candidates: List[str] = []
    if parameter_phrase:
        candidates.extend([
            f"Hanna {source_sku} adalah {category_phrase} untuk pengukuran {parameter_phrase}. Lihat spesifikasi produk dan konsultasikan kebutuhan dengan Tim UJI.",
            f"Hanna {source_sku} untuk pengukuran {parameter_phrase}. Lihat spesifikasi, fitur yang tersedia, dan konsultasikan kebutuhan dengan Tim UJI.",
        ])
    if product_label:
        candidates.append(
            f"Hanna {source_sku} {product_label}. Lihat spesifikasi dan informasi produk, lalu konsultasikan kebutuhan dengan Tim UJI."
        )
    candidates.append(
        f"Hanna {source_sku} adalah {category_phrase}. Lihat spesifikasi dan informasi produk, lalu konsultasikan kebutuhan dengan Tim UJI."
    )

    for candidate in candidates:
        candidate = normalize_space(candidate)
        if 105 <= len(candidate) <= 155:
            return candidate
    for candidate in candidates:
        candidate = normalize_space(candidate)
        if len(candidate) <= 155:
            return candidate

    # Fallback terakhir memakai kalimat pendek utuh, bukan pemotongan klausa.
    return normalize_space(f"Hanna {source_sku} adalah {category_phrase}. Lihat spesifikasi produk dan konsultasikan kebutuhan dengan Tim UJI.")

def detect_issues(
    product: Product,
    spec_rows: Sequence[Tuple[str, str]],
    generated: Dict[str, str],
    additional_issue: Optional[str] = None,
) -> List[str]:
    """QA source-adaptive. Missing optional content is not an error."""
    issues: List[str] = []
    if not product.sku:
        issues.append("REVIEW_SKU")
    if not product.title:
        issues.append("REVIEW_TITLE")
    if not product.images:
        issues.append("REVIEW_IMAGE")
    if product.specification and strip_html(product.specification) and not spec_rows:
        issues.append("REVIEW_SPEC_PARSE")

    category = generated.get("product categories", "")
    if not category:
        issues.append("REVIEW_CATEGORY")
    elif category not in CANONICAL_CATEGORIES and category not in {item[0] for item in VALID_NEW_CATEGORY_RULES}:
        issues.append("REVIEW_CATEGORY")

    # Hard fail jika output customer-facing mengandung referral/promo sumber.
    public_fields = [column for column in IMPORT_COLUMNS if column != "Source URL"]
    scan_text = " ".join(str(generated.get(key, "")) for key in public_fields).lower()
    blocked_source_phrases = (
        "hannainst.com", "according to hanna", "menurut hanna", "menurut produsen",
        "berdasarkan sumber", "berdasarkan spesifikasi", "berdasarkan dokumentasi",
        "spesifikasi resmi", "sumber resmi", "website resmi hanna", "situs resmi hanna",
        "official hanna website", "manufacturer website", "kunjungi website hanna",
        "lihat website hanna", "cek website hanna", "shop now", "add to cart",
    )
    if any(phrase in scan_text for phrase in blocked_source_phrases):
        issues.append("BLOCKED_SOURCE_PROMO")

    product_description = generated.get("Product Description", "")
    if strip_html(product_description).strip() and "+62 857-1711-2222" not in product_description:
        issues.append("REVIEW_CTA")
    if description_html_has_suspicious_heading_split(product_description):
        issues.append("REVIEW_DESCRIPTION_FORMAT")

    if contains_untranslated_english_prose(generated.get("Product Description", "")) or contains_untranslated_english_prose(
        generated.get("Product Short Description", "")
    ):
        issues.append("REVIEW_LANGUAGE")

    name = normalize_space(generated.get("Name", ""))
    if len(name) > 110:
        issues.append("REVIEW_NAME_LENGTH")

    seo_title = generated.get("SEO Title", "")
    if len(seo_title) < 25 or len(seo_title) > 65 or re.search(r"[|/.,:-]\s*\| UJI$", seo_title):
        issues.append("REVIEW_SEO_TITLE")

    meta = normalize_space(generated.get("meta description", ""))
    if len(meta) < 95 or len(meta) > 160:
        issues.append("REVIEW_META")
    if re.search(r"\b(?:dan|atau|untuk|dengan|yang|merupakan|sebagai)\.?$", meta, re.I):
        issues.append("REVIEW_META")

    raw_sections = {item.strip().upper() for item in generated.get("_raw_source_sections", "").split(",") if item.strip()}
    has_source_overview = "OVERVIEW" in raw_sections
    has_public_description = bool(strip_html(generated.get("Product Description", "")).strip())
    if has_source_overview and not has_public_description:
        issues.append("REVIEW_DESCRIPTION")
    elif not has_source_overview and has_public_description:
        issues.append("REVIEW_DESCRIPTION_SOURCE")

    # Tab boleh kosong. Tetapi judul tanpa isi atau isi tanpa judul adalah error.
    for index in range(1, 4):
        title = normalize_space(generated.get(f"custom_tab_{index}_title", ""))
        content = normalize_space(generated.get(f"custom_tab_{index}_content", ""))
        if bool(title) != bool(content):
            issues.append("REVIEW_CUSTOM_TAB")

    if additional_issue:
        issues.append(additional_issue)
    return list(dict.fromkeys(issues))

def transform_product(
    product: Product,
    config: Dict[str, object],
    faq_polisher: Optional[DeepSeekFAQPolisher] = None,
    require_polish: bool = False,
    content_writer: Optional[DeepSeekProductWriter] = None,
    require_content: bool = False,
    translator: Optional[DeepSeekEvidenceTranslator] = None,
) -> Tuple[Dict[str, str], List[str]]:
    # FAQ polisher lama dipertahankan di signature agar CLI/config lama tetap kompatibel,
    # tetapi V4.3 hanya membuat FAQ jika source memang memenuhi syarat.
    del faq_polisher, require_polish

    source_sku = clean_sku(product.sku)
    title = clean_title(product.title, source_sku)
    taxonomy = clean_taxonomy(product.taxonomy)
    source_short = html_to_structured_text(product.short_description, max_chars=5000)
    source_description = html_to_structured_text(product.description, max_chars=30000)
    spec_rows = parse_specification_table(product.specification)

    brand = str(config["brand_output"])
    display_brand = normalize_space(config.get("display_brand", "Hanna")) or "Hanna"
    category, category_issue = infer_category_with_status(
        title, taxonomy, source_short + " " + source_description
    )
    parameters = infer_parameters(title, taxonomy, spec_rows, source_short)
    name = build_name(title, source_sku, display_brand, category)
    import_sku = output_sku(source_sku, str(config["sku_prefix"]))
    slug = slugify(name)

    core_title = re.sub(
        rf"^\s*{re.escape(display_brand)}\s+{re.escape(source_sku)}\s*",
        "", name, flags=re.I,
    ).strip()
    focus = smart_truncate(normalize_space(f"Hanna {source_sku} {core_title or category}"), 120)
    seo_title = build_seo_title(source_sku, core_title or category, category, 60)

    # 1) PROVENANCE-AWARE SOURCE MODEL.
    # Seluruh body Description tetap menjadi OVERVIEW, termasuk inline heading
    # seperti "Features at-a-glance". Hanya DOM tab/panel terpisah yang boleh
    # masuk FEATURES/CONNECTIVITY/WHATS_INCLUDED/dll dan membuat custom tab UJI.
    description_buckets = description_body_to_overview(source_short, source_description)
    dedicated_tab_buckets = classify_dedicated_source_tabs(product.source_tabs)
    raw_buckets = merge_evidence_buckets(description_buckets, dedicated_tab_buckets)
    raw_buckets = trim_source_buckets(raw_buckets, total_limit=24000)

    # 2) Terjemahkan narrative evidence sebelum masuk ke product writer.
    if translator is not None:
        translated_buckets, translation_source = translator.translate(source_sku, raw_buckets)
    else:
        translated_buckets = DeepSeekEvidenceTranslator.safe_fallback(raw_buckets)
        translation_source = "TRANSLATION_TEMPLATE"

    # Dedicated tab buckets harus diterjemahkan dari hasil gabungan berdasarkan kind,
    # tetapi hanya kind yang memang berasal dari source tab boleh membuat custom tab.
    translated_tab_buckets = {
        kind: list(translated_buckets.get(kind, []))
        for kind in dedicated_tab_buckets
        if translated_buckets.get(kind)
    }

    raw_line_count = sum(len(lines) for lines in raw_buckets.values())
    translated_line_count = sum(len(lines) for lines in translated_buckets.values())
    translation_issue: Optional[str] = None
    if raw_line_count and translated_line_count == 0:
        translation_issue = "REVIEW_TRANSLATION"
    elif raw_line_count >= 3 and translation_source in {"TRANSLATION_FALLBACK", "TRANSLATION_TEMPLATE"} and translated_line_count < max(1, raw_line_count // 2):
        translation_issue = "REVIEW_TRANSLATION"

    translated_html = buckets_to_structured_html(translated_buckets)
    evidence = build_source_evidence(
        name,
        source_sku,
        category,
        parameters,
        "",
        translated_html,
        spec_rows,
        config,
        source_buckets=translated_buckets,
    )

    # 3) Fallback mengikuti source depth, lalu AI boleh memperhalus tanpa menambah section.
    content = build_fallback_product_content(
        name,
        source_sku,
        category,
        parameters,
        spec_rows,
        "",
        translated_html,
        source_buckets=translated_buckets,
    )
    content_source = "SOURCE_ADAPTIVE_TEMPLATE"
    content_issue: Optional[str] = None

    if content_writer and content_writer.enabled:
        try:
            content, content_source = content_writer.generate(
                name, source_sku, category, parameters, evidence
            )
        except ImporterError as exc:
            logging.warning("Konten %s memakai fallback source-adaptive karena AI gagal: %s", source_sku, exc)
            content_source = "SOURCE_ADAPTIVE_FALLBACK"
            if require_content:
                content_issue = "REVIEW_CONTENT_AI"
    elif require_content:
        content_issue = "REVIEW_CONTENT_AI"

    short_desc = clean_generated_text(content.get("short_description"), max_words=80)
    full_desc = build_rich_description_html(name, source_sku, content, config)

    # 4) Custom tab benar-benar dinamis. Maksimal 3 tab dan tidak ada tab filler.
    tabs = build_dynamic_custom_tabs(source_sku, spec_rows, translated_tab_buckets, content)
    tab_values: Dict[str, str] = {}
    for index in range(1, 4):
        if index <= len(tabs):
            tab_title, tab_content = tabs[index - 1]
            tab_values[f"custom_tab_{index}_title"] = tab_title
            tab_values[f"custom_tab_{index}_content"] = tab_content
            tab_values[f"custom_tab_{index}_priority"] = str(index * 10)
        else:
            tab_values[f"custom_tab_{index}_title"] = ""
            tab_values[f"custom_tab_{index}_content"] = ""
            tab_values[f"custom_tab_{index}_priority"] = ""

    faq_source = content_source if any(title == "FAQ" for title, _ in tabs) else "NONE"
    meta = build_meta_description(source_sku, category, parameters, name=name)

    generated: Dict[str, str] = {
        "Name": name,
        "SKU": import_sku,
        "Brand": brand,
        "Slug": slug,
        "SEO Title": seo_title,
        "Product Description": full_desc,
        "Product Short Description": short_desc,
        **tab_values,
        "product categories": category,
        "product tags": build_tags(source_sku, category, parameters),
        "focus keyphrase": focus,
        "meta description": meta,
        "publication_date": "",
        "image_url": normalize_image_urls(product.images),
        "Source URL": product.url if bool(config.get("include_source_url_in_import")) else "",
        "SEO QA Status": "",
        "_faq_source": faq_source,
        "_content_source": content_source,
        "_translation_source": translation_source,
        "_source_chars": str(len(evidence)),
        "_source_sections": ", ".join(sorted(translated_buckets.keys())),
        "_raw_source_sections": ", ".join(sorted(raw_buckets.keys())),
        "_dedicated_source_tabs": ", ".join(sorted(dedicated_tab_buckets.keys())),
        "_parameters": ", ".join(parameters),
    }

    pre_issues = [issue for issue in (category_issue, translation_issue, content_issue) if issue]
    issues = detect_issues(product, spec_rows, generated, pre_issues[0] if pre_issues else None)
    for issue in pre_issues[1:]:
        if issue not in issues:
            issues.append(issue)
    generated["SEO QA Status"] = "READY" if not issues else "; ".join(issues)
    return generated, issues

def normalize_image_urls(value: object) -> str:
    text = normalize_space(value)
    if not text:
        return ""
    urls = re.split(r"\s*,\s*|\s*\|\s*|\s+", text)
    cleaned: List[str] = []
    for url in urls:
        url = url.strip(" ,;\"'")
        if not url.lower().startswith(("http://", "https://")):
            continue
        if url not in cleaned:
            cleaned.append(url)
    return ",".join(cleaned)


def load_config(path: Path) -> Dict[str, object]:
    config = dict(DEFAULT_CONFIG)
    if path.exists():
        with path.open("r", encoding="utf-8") as handle:
            supplied = json.load(handle)
        if not isinstance(supplied, dict):
            raise ImporterError("Isi konfigurasi harus berupa objek JSON.")
        config.update(supplied)
    return config


def identify_source_sheet(workbook, preferred_name: str):
    if preferred_name in workbook.sheetnames:
        return workbook[preferred_name]
    for worksheet in workbook.worksheets:
        headers = [normalize_space(cell.value).lower() for cell in worksheet[1]]
        if all(column in headers for column in SOURCE_COLUMNS):
            return worksheet
    return workbook.active


def read_products_from_xlsx(path: Path, sheet_name: str) -> Iterable[Product]:
    if not path.exists():
        raise ImporterError(f"File input tidak ditemukan: {path}")
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = identify_source_sheet(workbook, sheet_name)
    headers = [normalize_space(cell.value).lower() for cell in worksheet[1]]
    positions = {name: headers.index(name) for name in SOURCE_COLUMNS if name in headers}
    missing = [name for name in SOURCE_COLUMNS if name not in positions]
    if missing:
        raise ImporterError("Kolom sumber tidak lengkap: " + ", ".join(missing))

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if not any(value not in (None, "") for value in row):
            continue
        kwargs = {}
        for name, index in positions.items():
            value = row[index] if index < len(row) else None
            if name in {"description", "short_description", "specification", "images"}:
                kwargs[name] = "" if value is None else str(value)
            else:
                kwargs[name] = normalize_space(value)
        product = Product(**kwargs)
        if product.sku or product.title or product.url:
            yield product


def create_http_session(config: Dict[str, object]) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": str(config["user_agent"]),
        "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
    })
    return session


def request_page(session: requests.Session, url: str, config: Dict[str, object]) -> str:
    last_error: Optional[Exception] = None
    for attempt in range(1, int(config["max_retries"]) + 1):
        try:
            response = session.get(url, timeout=int(config["request_timeout_seconds"]))
            response.raise_for_status()
            return response.text
        except requests.RequestException as exc:
            last_error = exc
            logging.warning("Gagal mengambil %s, percobaan %s: %s", url, attempt, exc)
            time.sleep(attempt * 2)
    raise ImporterError(f"Gagal mengambil halaman {url}: {last_error}")


def json_ld_objects(soup: BeautifulSoup) -> Iterable[dict]:
    for script in soup.find_all("script", attrs={"type": re.compile(r"application/ld\+json", re.I)}):
        raw = script.string or script.get_text("", strip=True)
        if not raw:
            continue
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            continue
        values = payload if isinstance(payload, list) else [payload]
        for item in values:
            if isinstance(item, dict) and "@graph" in item and isinstance(item["@graph"], list):
                values.extend(node for node in item["@graph"] if isinstance(node, dict))
            if isinstance(item, dict):
                yield item


def score_specification_table(table: object) -> int:
    rows = [tr for tr in table.find_all("tr") if len(tr.find_all(["th", "td"])) >= 2]
    text_value = normalize_space(table.get_text(" ", strip=True)).lower()
    score = len(rows)
    for keyword in ("range", "accuracy", "resolution", "calibration", "sku", "power supply", "dimensions", "weight"):
        if keyword in text_value:
            score += 6
    if "rounded to nearest" in text_value or "report results" in text_value:
        score -= 12
    return score



def _tab_kind_from_title(title: str) -> Optional[str]:
    lower = validation_text(title)
    # Description and specification are handled separately.
    if re.search(r"\b(description|overview|product description|specifications?)\b", lower, re.I):
        return None
    for kind, patterns in SOURCE_SECTION_RULES:
        if kind == "OVERVIEW":
            continue
        if any(re.search(pattern, lower, re.I) for pattern in patterns):
            return kind
    return None


def extract_dedicated_source_tabs(soup: BeautifulSoup) -> Dict[str, str]:
    """Extract only DOM-backed tabs/accordions/panels from the Hanna page.

    A heading inside normal Description content is deliberately ignored. A section
    qualifies only when a tab/accordion control points to a separate content panel
    through href, aria-controls, data-target, data-tab-target, or a recognized
    accordion sibling structure.
    """
    found: Dict[str, str] = {}
    seen_panels: set[int] = set()

    def add(title: str, panel: object) -> None:
        title = clean_source_line(title)
        if not title or _tab_kind_from_title(title) is None or panel is None:
            return
        marker = id(panel)
        if marker in seen_panels:
            return
        body = html_to_structured_text(str(panel), max_chars=14000)
        if not body or len(strip_html(body)) < 20:
            return
        # Remove an echoed panel title when the panel starts with the same heading.
        body_lines = [normalize_space(x) for x in body.splitlines() if normalize_space(x)]
        if body_lines and validation_text(body_lines[0].replace("## ", "")) == validation_text(title):
            body_lines = body_lines[1:]
        body = "\n".join(body_lines).strip()
        if body:
            found[title] = body
            seen_panels.add(marker)

    # Strong provenance: a control explicitly targets a separate panel.
    for control in soup.find_all(["a", "button"]):
        title = normalize_space(control.get_text(" ", strip=True))
        if not title or _tab_kind_from_title(title) is None:
            continue
        # Do not treat ordinary jump links inside Description as tabs.
        parent_signature_parts: List[str] = []
        parent = getattr(control, "parent", None)
        for _ in range(4):
            if parent is None:
                break
            parent_signature_parts.extend(parent.get("class", []) or [])
            if parent.get("role"):
                parent_signature_parts.append(str(parent.get("role")))
            parent = getattr(parent, "parent", None)
        parent_signature = " ".join(parent_signature_parts).lower()
        control_signature = " ".join(control.get("class", []) or []).lower()
        is_tab_control = (
            str(control.get("role", "")).lower() == "tab"
            or str(control.get("data-toggle", "")).lower() in {"tab", "collapse"}
            or control.has_attr("data-tab")
            or control.has_attr("data-accordion-title")
            or any(token in (parent_signature + " " + control_signature) for token in (
                "tabs", "tab-title", "nav-tabs", "accordion", "accordion-title",
                "accordion-navigation", "panel-heading", "card-header"
            ))
        )
        if not is_tab_control:
            continue
        target = (
            control.get("aria-controls")
            or control.get("data-target")
            or control.get("data-tab-target")
            or control.get("href")
        )
        if not target:
            continue
        target = str(target).strip()
        target_id = ""
        if target.startswith("#"):
            target_id = target[1:]
        elif re.fullmatch(r"[A-Za-z][A-Za-z0-9_:\-.]*", target):
            target_id = target
        if not target_id:
            continue
        panel = soup.find(id=target_id)
        if panel is not None and panel is not control:
            add(title, panel)

    # Accordion pattern without aria-controls: heading/control followed by a sibling panel.
    accordion_control_selectors = (
        ".accordion-title", ".accordion-navigation", ".accordion-toggle",
        "[data-accordion-title]", ".card-header", ".panel-title",
    )
    for selector in accordion_control_selectors:
        for control in soup.select(selector):
            title = normalize_space(control.get_text(" ", strip=True))
            if not title or _tab_kind_from_title(title) is None:
                continue
            panel = control.find_next_sibling()
            if panel is None and getattr(control, "parent", None) is not None:
                panel = control.parent.find_next_sibling()
            if panel is not None:
                add(title, panel)

    return found


def extract_primary_description_content(soup: BeautifulSoup, json_description: str = "") -> str:
    """Extract Product Description only, excluding other dedicated panels.

    Selection order is intentional: a dedicated Description panel wins over a
    larger wrapper such as .productView-description that may contain all tabs.
    """
    selector_groups = (
        ("#tab-description", "#description", "[id*='description'][role='tabpanel']"),
        ("[data-product-description]", "[itemprop='description']", ".product-description"),
        (".productView-description",),
    )
    for selectors in selector_groups:
        candidates: List[object] = []
        for selector in selectors:
            for node in soup.select(selector):
                if node not in candidates:
                    candidates.append(node)
        if not candidates:
            continue
        # Prefer the richest candidate within the same specificity group.
        best = max(candidates, key=lambda node: len(normalize_space(node.get_text(" ", strip=True))))
        clone = BeautifulSoup(str(best), "html.parser")
        # If this is a wrapper, remove non-description tab panels before flattening.
        for panel in clone.select("[role='tabpanel'], .tab-content, .tabs-content, .accordion-content"):
            panel_id = validation_text(panel.get("id", ""))
            panel_classes = validation_text(" ".join(panel.get("class", []) or []))
            signature = panel_id + " " + panel_classes
            if "description" not in signature and panel is not clone:
                panel.decompose()
        structured = html_to_structured_text(str(clone), max_chars=30000)
        if structured:
            return structured
    if json_description:
        return html_to_structured_text(json_description, max_chars=30000)
    return ""

def scrape_product_page(url: str, session: requests.Session, config: Dict[str, object]) -> Product:
    page = request_page(session, url, config)
    soup = BeautifulSoup(page, "html.parser")
    product_data: Dict[str, object] = {}
    for item in json_ld_objects(soup):
        item_type = item.get("@type")
        types = item_type if isinstance(item_type, list) else [item_type]
        if "Product" in types:
            product_data = item
            break

    title = normalize_space(product_data.get("name"))
    sku = normalize_space(product_data.get("sku") or product_data.get("mpn"))
    json_description = str(product_data.get("description") or "")
    image_obj = product_data.get("image") or []
    images = [image_obj] if isinstance(image_obj, str) else list(image_obj) if isinstance(image_obj, list) else []
    brand_obj = product_data.get("brand")
    if isinstance(brand_obj, dict):
        brand = normalize_space(brand_obj.get("name"))
    else:
        brand = normalize_space(brand_obj)

    if not title:
        heading = soup.select_one("h1.productView-title, h1.page-heading, h1")
        title = normalize_space(heading.get_text(" ", strip=True) if heading else "")
    if not sku:
        sku_node = soup.select_one("[data-product-sku], .productView-info-value[data-product-sku], dd.productView-info-value")
        sku = normalize_space(
            sku_node.get("data-product-sku")
            if sku_node and sku_node.has_attr("data-product-sku")
            else sku_node.get_text(" ", strip=True) if sku_node else ""
        )

    short_description = json_description
    if not short_description:
        short_node = soup.select_one(".productView-description, [itemprop='description'], .product-description")
        short_description = html_to_structured_text(str(short_node), max_chars=1800) if short_node else ""

    # PROVENANCE LOCK: Description dan dedicated source tabs dipisahkan sejak DOM.
    # Inline heading di Description tidak pernah dianggap sebagai source tab.
    content_soup = BeautifulSoup(page, "html.parser")
    rich_description = extract_primary_description_content(content_soup, json_description)
    dedicated_tabs = extract_dedicated_source_tabs(BeautifulSoup(page, "html.parser"))

    # SPECIFICATION LOCK:
    # Gunakan logika V3 persis untuk fallback spesifikasi dari halaman Hanna.
    # Dalam alur normal, specification dari XLSX sumber tetap dipertahankan
    # dan tidak ditimpa oleh hasil refresh website.
    best_table = ""
    best_rows = 0
    for table in soup.find_all("table"):
        row_count = sum(1 for tr in table.find_all("tr") if len(tr.find_all(["th", "td"])) >= 2)
        if row_count > best_rows:
            best_rows = row_count
            best_table = str(table)

    breadcrumbs = [
        normalize_space(node.get_text(" ", strip=True))
        for node in soup.select(".breadcrumb a, nav[aria-label*='breadcrumb' i] a")
    ]
    taxonomy = " > ".join(
        item for item in breadcrumbs
        if item and item.lower() not in {"home", "hanna instruments"}
    )

    if not images:
        for img in soup.select(".productView-image img, [data-image-gallery-main] img, img[itemprop='image']"):
            src = img.get("data-zoom-image") or img.get("data-src") or img.get("src")
            if src:
                images.append(urljoin(url, src))

    return Product(
        url=url,
        brand=brand or "Hanna Instruments",
        sku=sku,
        title=title,
        taxonomy=taxonomy,
        short_description=short_description,
        description=rich_description,
        images=",".join(dict.fromkeys(images)),
        specification=best_table,
        source_tabs=json.dumps(dedicated_tabs, ensure_ascii=False),
    )


def optionally_refresh_product(product: Product, session: requests.Session, config: Dict[str, object], refresh: bool) -> Product:
    if not refresh or not product.url:
        return product
    try:
        scraped = scrape_product_page(product.url, session, config)
    except Exception as exc:  # Proses harus tetap berjalan untuk produk lain.
        logging.error("Refresh gagal untuk %s: %s", product.url, exc)
        return product

    # SPECIFICATION LOCK + SOURCE IDENTITY LOCK:
    # Refresh website hanya dipakai untuk memperkaya konten naratif dan gambar.
    # SKU, title, taxonomy, dan terutama specification yang sudah ada di XLSX
    # TIDAK BOLEH ditimpa. Ini menjaga tab Spesifikasi tetap sama seperti V3.
    return Product(
        url=product.url,
        brand=product.brand or scraped.brand,
        sku=product.sku or scraped.sku,
        title=product.title or scraped.title,
        taxonomy=product.taxonomy or scraped.taxonomy,
        short_description=scraped.short_description or product.short_description,
        description=scraped.description or product.description,
        images=scraped.images or product.images,
        specification=product.specification or scraped.specification,
        source_tabs=scraped.source_tabs or product.source_tabs,
    )


def style_worksheet(worksheet, widths: Dict[int, int]) -> None:
    header_fill = PatternFill("solid", fgColor="17365D")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for index, width in widths.items():
        worksheet.column_dimensions[get_column_letter(index)].width = width
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_output(
    output_path: Path,
    rows: List[Dict[str, str]],
    review_rows: List[Dict[str, str]],
    config: Dict[str, object],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    ws = workbook.active
    ws.title = "IMPORT_UJI"
    ws.append(IMPORT_COLUMNS)
    for item in rows:
        ws.append([item.get(column, "") for column in IMPORT_COLUMNS])

    style_worksheet(ws, {
        1: 52, 2: 20, 3: 20, 4: 52, 5: 58, 6: 95, 7: 75,
        8: 22, 9: 85, 10: 12, 11: 24, 12: 80, 13: 12, 14: 18,
        15: 80, 16: 12, 17: 38, 18: 65, 19: 55, 20: 75, 21: 18,
        22: 75, 23: 70, 24: 32,
    })

    review = workbook.create_sheet("REVIEW")
    review_headers = ["No", "SKU Asli", "Nama Produk", "Status", "Masalah", "Content Generator", "Translation Layer", "FAQ Generator", "Karakter Bukti", "Section Evidence", "Dedicated Source Tabs", "Parameter Terdeteksi", "URL Sumber", "Gambar", "Kategori"]
    review.append(review_headers)
    for number, item in enumerate(review_rows, start=1):
        review.append([
            number,
            item.get("source_sku", ""),
            item.get("name", ""),
            item.get("status", ""),
            item.get("issues", ""),
            item.get("content_source", ""),
            item.get("translation_source", ""),
            item.get("faq_source", ""),
            item.get("source_chars", ""),
            item.get("source_sections", ""),
            item.get("dedicated_source_tabs", ""),
            item.get("parameters", ""),
            item.get("source_url", ""),
            item.get("images", ""),
            item.get("category", ""),
        ])
    style_worksheet(review, {1: 8, 2: 22, 3: 60, 4: 34, 5: 55, 6: 24, 7: 24, 8: 18, 9: 16, 10: 38, 11: 35, 12: 75, 13: 75, 14: 38})

    info = workbook.create_sheet("PETUNJUK")
    info_rows = [
        ["HANNA -> UJI PRODUCT IMPORTER", f"Versi {VERSION}"],
        ["Jumlah produk", len(rows)],
        ["Produk READY", sum(1 for item in rows if item["SEO QA Status"] == "READY")],
        ["Produk perlu review", sum(1 for item in rows if item["SEO QA Status"] != "READY")],
        ["WhatsApp", config["whatsapp_display"]],
        ["Email", config["email"]],
        ["Catatan", "Kolom Source URL dikosongkan secara default agar tidak tampil di website. URL asli tersedia pada sheet REVIEW."],
        ["Catatan gambar", "URL gambar sumber masih digunakan untuk proses impor. Setelah impor, simpan gambar ke Media Library uji.co.id agar tidak melakukan hotlink permanen."],
        ["Konten", "V4.3.4 mempertahankan provenance DOM, Description Formatter deterministik, dan Heading Guard: seluruh body Description/Overview tetap menjadi Product Description, termasuk heading inline. Custom tab hanya dibuat dari dedicated source tab/panel yang benar-benar terpisah di Hanna. Jika source tidak punya description/overview, Product Description boleh kosong."],
        ["Scraping", "Mode refresh aktif secara default untuk mengambil rich content seperti Features, Principle of Operation, aplikasi, kalibrasi, GLP, data logging, isi paket, dan bagian teknis lain. Tab Spesifikasi tetap memakai data specification dari XLSX sumber dan tidak ditimpa oleh refresh website."],
        ["DeepSeek", "API key dibaca dari environment DEEPSEEK_API_KEY atau config. Translation layer berjalan sebelum product writer dan validator mengunci angka, model, unit, serta istilah teknis."],
        ["Kategori", "Daftar kategori UJI yang diberikan menjadi registry canonical. Taxonomy Hanna tidak pernah dipakai sebagai fallback mentah. Kategori baru yang valid diberi REVIEW_NEW_CATEGORY sebelum disetujui."],
        ["Custom tab", "Maksimal tiga tab dipilih hanya dari Specification Lock dan dedicated source tab/panel: Keunggulan Produk, Isi Paket, Koneksi & Pelaporan, Instalasi, Kalibrasi, Perawatan, Kompatibilitas, atau FAQ. Inline heading di Description tidak membuat custom tab."],
        ["Publikasi", "Publikasikan produk berstatus READY terlebih dahulu. Periksa semua kode REVIEW sebelum impor massal."],
    ]
    for row in info_rows:
        info.append(row)
    info["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    info["B1"].font = Font(bold=True, color="FFFFFF")
    info["A1"].fill = PatternFill("solid", fgColor="17365D")
    info["B1"].fill = PatternFill("solid", fgColor="17365D")
    info.column_dimensions["A"].width = 25
    info.column_dimensions["B"].width = 110
    for row in info.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    workbook.save(output_path)


def download_images(rows: List[Dict[str, str]], config: Dict[str, object]) -> None:
    if not bool(config.get("download_images")):
        return
    directory = Path(str(config["image_directory"]))
    directory.mkdir(parents=True, exist_ok=True)
    session = create_http_session(config)
    manifest: List[Tuple[str, str, str]] = []
    for item in rows:
        sku = item.get("SKU", "product")
        for index, url in enumerate(item.get("image_url", "").split(","), start=1):
            url = url.strip()
            if not url:
                continue
            suffix = Path(urlparse(url).path).suffix.lower()
            if suffix not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                suffix = ".jpg"
            filename = f"{slugify(sku)}-{index}{suffix}"
            filepath = directory / filename
            try:
                response = session.get(url, timeout=int(config["request_timeout_seconds"]))
                response.raise_for_status()
                filepath.write_bytes(response.content)
                manifest.append((sku, url, str(filepath)))
            except requests.RequestException as exc:
                logging.error("Gagal mengunduh gambar %s: %s", url, exc)
    manifest_path = directory / "manifest.csv"
    with manifest_path.open("w", encoding="utf-8-sig", newline="") as handle:
        handle.write("SKU,URL Sumber,File Lokal\n")
        for sku, url, path in manifest:
            handle.write(f'"{sku}","{url}","{path}"\n')


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Mengambil konten lengkap produk Hanna dan membuat XLSX siap impor ke uji.co.id."
    )
    parser.add_argument("--input", default=DEFAULT_INPUT, help=f"File XLSX sumber. Default: {DEFAULT_INPUT}")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help=f"File XLSX hasil. Default: {DEFAULT_OUTPUT}")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Nama sheet sumber. Default: {DEFAULT_SHEET}")
    parser.add_argument("--config", default="config_hanna_uji_v4_3.json", help="File konfigurasi JSON.")
    parser.add_argument("--limit", type=int, default=0, help="Batasi jumlah produk. Gunakan 20 untuk uji coba.")
    parser.add_argument(
        "--refresh",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Scrape ulang setiap URL. Aktif secara default; gunakan --no-refresh untuk memakai data XLSX apa adanya.",
    )
    parser.add_argument(
        "--content-mode",
        choices=("auto", "deepseek", "template"),
        default="auto",
        help="auto: gunakan AI jika API key tersedia; deepseek: wajib AI; template: tanpa AI.",
    )
    parser.add_argument(
        "--faq-mode",
        choices=("auto", "polish", "deepseek", "template"),
        default="auto",
        help="Kompatibilitas V4 lama. Pada V4.3 FAQ tetap opsional dan hanya dibuat jika source memenuhi syarat.",
    )
    parser.add_argument(
        "--prompt-api-key",
        action="store_true",
        help="Minta DeepSeek API key secara tersembunyi.",
    )
    parser.add_argument("--log", default="output/hanna_uji_v4_3_4.log", help="Lokasi file log.")
    parser.add_argument("--version", action="version", version=VERSION)
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    log_path = Path(args.log)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )

    faq_polisher: Optional[DeepSeekFAQPolisher] = None
    content_writer: Optional[DeepSeekProductWriter] = None
    translator: Optional[DeepSeekEvidenceTranslator] = None
    try:
        config = load_config(Path(args.config))
        api_key = normalize_space(os.environ.get("DEEPSEEK_API_KEY", "") or config.get("deepseek_api_key", ""))
        ai_requested = args.content_mode != "template"
        if args.prompt_api_key and not api_key and ai_requested:
            api_key = getpass.getpass("Masukkan DeepSeek API key: ").strip()

        translation_enabled = args.content_mode != "template" and bool(config.get("use_deepseek_translation", True))
        content_enabled = args.content_mode != "template" and bool(config.get("use_deepseek_content", True))
        require_content = args.content_mode == "deepseek"
        if require_content and not api_key:
            raise ImporterError(
                "Content mode deepseek dipilih, tetapi API key belum tersedia. Gunakan --prompt-api-key "
                "atau set environment DEEPSEEK_API_KEY."
            )

        translator = DeepSeekEvidenceTranslator(config, api_key, translation_enabled)
        content_writer = DeepSeekProductWriter(config, api_key, content_enabled)

        # --faq-mode dipertahankan hanya agar command lama tidak error. Pada V4.3
        # FAQ dibuat oleh source-adaptive product writer dan tidak dapat dipaksa.
        require_polish = False
        faq_polisher = None
        if args.faq_mode not in {"auto", "template"}:
            logging.info("FAQ mode %s diterima sebagai opsi legacy; V4.3 tetap memakai FAQ source-adaptive.", args.faq_mode)

        if translation_enabled and not api_key:
            logging.warning("DeepSeek API key tidak ditemukan. Translation layer memakai fallback aman dan dapat menandai REVIEW_TRANSLATION.")
        if content_enabled and not api_key:
            logging.warning("DeepSeek API key tidak ditemukan. Konten memakai source-adaptive template tanpa memaksa fitur/FAQ.")

        products = read_products_from_xlsx(Path(args.input), args.sheet)
        session = create_http_session(config)
        rows: List[Dict[str, str]] = []
        review_rows: List[Dict[str, str]] = []

        for number, raw_product in enumerate(products, start=1):
            if args.limit and number > args.limit:
                break
            product = optionally_refresh_product(raw_product, session, config, args.refresh)
            generated, issues = transform_product(
                product,
                config,
                faq_polisher,
                require_polish,
                content_writer,
                require_content,
                translator,
            )
            rows.append(generated)
            if issues:
                review_rows.append({
                    "source_sku": product.sku,
                    "name": generated["Name"],
                    "status": generated["SEO QA Status"],
                    "issues": ", ".join(issues),
                    "content_source": generated.get("_content_source", ""),
                    "translation_source": generated.get("_translation_source", ""),
                    "faq_source": generated.get("_faq_source", ""),
                    "source_chars": generated.get("_source_chars", ""),
                    "source_sections": generated.get("_source_sections", ""),
                    "dedicated_source_tabs": generated.get("_dedicated_source_tabs", ""),
                    "parameters": generated.get("_parameters", ""),
                    "source_url": product.url,
                    "images": product.images,
                    "category": generated["product categories"],
                })
            if number % 20 == 0:
                logging.info("Diproses: %s produk", number)
            if args.refresh:
                time.sleep(float(config["request_delay_seconds"]))

        if not rows:
            raise ImporterError("Tidak ada produk yang berhasil dibaca dari file sumber.")

        if translator:
            translator.flush()
        if content_writer:
            content_writer.flush()
        if faq_polisher:
            faq_polisher.flush()
        write_output(Path(args.output), rows, review_rows, config)
        download_images(rows, config)
        ready = sum(1 for item in rows if item["SEO QA Status"] == "READY")
        rich_ai = sum(1 for item in rows if item.get("_content_source") in {"AI_RICH", "AI_RICH_CACHE"})
        translated_ai = sum(1 for item in rows if item.get("_translation_source") in {"TRANSLATION_AI", "TRANSLATION_CACHE"})
        logging.info(
            "Selesai. Total: %s | READY: %s | Review: %s | Konten AI: %s | Terjemahan AI: %s",
            len(rows), ready, len(rows) - ready, rich_ai, translated_ai,
        )
        logging.info("File hasil: %s", Path(args.output).resolve())
        return 0
    except (ImporterError, OSError, ValueError, json.JSONDecodeError) as exc:
        if translator:
            translator.flush()
        if content_writer:
            content_writer.flush()
        if faq_polisher:
            faq_polisher.flush()
        logging.error("Proses dihentikan: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

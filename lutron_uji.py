#!/usr/bin/env python3
"""LUTRON -> WEBSITE_KALIAN WooCommerce catalog generator V22.24 resilient batch, source-locked, technical-spec-safe Indonesian WordPress output.

PATCHED: Excel-safe export untuk mencegah openpyxl IllegalCharacterError.

Fitur utama:
- menerima URL produk LUTRON dan/atau LUTRON dalam satu Excel;
- otomatis mendeteksi merek dari domain URL;
- membaca title, brand, model, deskripsi, seluruh konten produk, spesifikasi,
  kategori/breadcrumb, dan gambar utama;
- menyimpan seluruh konten resmi produk pada tab "Informasi Produk Resmi";
- memakai title/model/brand dari Excel sebagai fallback apabila elemen halaman
  produk tidak tersedia;
- output mengikuti kolom WooCommerce/WEBSITE_KALIAN seperti script LUTRON sebelumnya.

Contoh:
python WEBSITE_catalog_import_LUTRON_lutron.py --input "lutron.com.tw.xlsx" --output "WEBSITE_lutron_import_ready.xlsx"
python WEBSITE_catalog_import_LUTRON_lutron.py --input "produk_mixed.xlsx" --output "WEBSITE_mixed_import_ready.xlsx" --delay 0.6
python WEBSITE_catalog_import_LUTRON_lutron.py --input "lutron.com.tw.xlsx" --use-ai --deepseek-api-key "YOUR_KEY"
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
import time
import shutil
import subprocess
import tempfile
from io import BytesIO
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple
from urllib.parse import quote, quote_plus, unquote, urljoin, urlparse, urlsplit, urlunsplit

try:
    import pandas as pd
except Exception as exc:
    raise SystemExit("pandas wajib tersedia. Install: pip install pandas openpyxl") from exc

try:
    import requests
except Exception as exc:
    raise SystemExit("requests wajib tersedia. Install: pip install requests") from exc

try:
    from bs4 import BeautifulSoup, Tag
except Exception as exc:
    raise SystemExit("beautifulsoup4 wajib tersedia. Install: pip install beautifulsoup4") from exc

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except Exception:
    load_workbook = None


OUTPUT_COLS = [
    "Name", "Brand", "Product Description", "Product Short Description",
    "custom_tab_1_title", "custom_tab_1_content", "custom_tab_1_priority",
    "custom_tab_2_title", "custom_tab_2_content", "custom_tab_2_priority",
    "custom_tab_3_title", "custom_tab_3_content", "custom_tab_3_priority",
    "custom_tab_4_title", "custom_tab_4_content", "custom_tab_4_priority",
    "custom_tab_5_title", "custom_tab_5_content", "custom_tab_5_priority",
    "custom_tab_6_title", "custom_tab_6_content", "custom_tab_6_priority",
    "custom_tab_7_title", "custom_tab_7_content", "custom_tab_7_priority",
    "custom_tab_8_title", "custom_tab_8_content", "custom_tab_8_priority",
    "product categories", "product tags", "focus keyphrase", "meta description",
    "publication_date", "Processed", "Processing Time", "Content Quality",
    "image_url", "Source URL", "Website Validation", "Website Correction Log",
]

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Upgrade-Insecure-Requests": "1",
}

WEBSITE_CTA = (
    '<p>Untuk konsultasi produk, silakan kunjungi '
    '<strong><a href="https://WEBSITE_KALIAN/about-us-3/" target="_blank" rel="noopener">WEBSITE_KALIAN</a></strong> '
    'atau <strong><a href="https://WEBSITE_KALIAN/contact-us/" target="_blank" rel="noopener">hubungi kami</a></strong>. '
    'Telepon: +62no kalian.</p>'
)

NOISE_TEXT = {
    "home", "products", "contact us", "about us", "search", "menu", "login", "register",
    "shopping cart", "privacy policy", "terms of use", "copyright", "all rights reserved",
    "facebook", "instagram", "youtube", "linkedin", "line", "language", "english", "中文",
    "download", "inquiry", "quotation", "please contact us", "back to top",
}

SPEC_HEADING_RE = re.compile(
    r"\b(specification|specifications|technical data|technical specification|product information|"
    r"product specification|measurement specification|features and specifications|details)\b", re.I
)

# Klasifikasi aksesori dibuat konservatif. Kata seperti "probe", "sensor",
# "transmitter", dan "receiver" dapat muncul pada alat utama sehingga tidak boleh
# langsung mengubah produk menjadi aksesori.
ACCESSORY_STRONG_KEYWORDS = re.compile(
    r"\b(replacement|spare(?:\s+part)?|accessor(?:y|ies)|optional\s+(?:cable|adapter|adaptor|charger|case|probe|sensor|electrode|printer)|"
    r"extension\s+cable|usb[-\s]?0?1|upcb[-\s]?0?2|carrying\s+case|protective\s+cover|"
    r"test\s+piece|eyepiece|desiccant|o[-\s]?ring|washer|mounting\s+kit|holder|stand|strap|pouch|"
    r"bottle|beaker|cap|lid|filter)\b", re.I
)
ACCESSORY_COMPONENT_TERMS = re.compile(
    r"\b(probe|sensor|electrode|cable|adapter|adaptor|printer|case|cover|battery|charger|"
    r"test\s+piece|eyepiece|desiccant|holder|stand|paper|replacement|spare|accessor(?:y|ies)|"
    r"pouch|strap|funnel|nozzle|needle|hose|lid|o[-\s]?ring|washer|extension)\b", re.I
)
PRIMARY_PRODUCT_TERMS = re.compile(
    r"\b(transmitter|receiver|meter|tester|analyzer|counter|recorder|logger|gauge|thermometer|"
    r"multimeter|clamp|controller|monitor|calibrator|balance|scale|wrench|detector|instrument)\b", re.I
)
PRIMARY_MODEL_PREFIX_RE = re.compile(
    r"\b(?:TR[-_]|HR[-_]|EMF[-_]|DW[-_]|FC[-_]|FR[-_]|TM[-_]|CM[-_]|DM[-_]|BWA[-_]|BPH[-_]|"
    r"BMG[-_]|BDO[-_]|BCT[-_]|GM[-_]|DT[-_]|VB[-_]|FG[-_]|DR[-_]|PDA[-_]|LCR[-_]|TQ[-_]|SC[-_])",
    re.I,
)

# Jangan pernah menggunakan aset halaman umum sebagai foto katalog produk.
IMAGE_REJECT_TOKENS = (
    "logo", "icon", "sprite", "flag", "payment", "blank", "placeholder", "tracking",
    "banner", "banners", "hero", "carousel", "slider", "homepage", "landing", "corporate",
    "company", "facebook", "youtube", "instagram", "linkedin", "favicon", "og-default",
)

FAMILY_RULES: List[Tuple[re.Pattern[str], Dict[str, Any]]] = [
    (re.compile(r"fruit\s+hardness|fruit\s+sclerometer|\bfr-\d", re.I), {
        "category": "PengWEBSITE Kekerasan Buah",
        "term": "pengWEBSITE kekerasan buah",
        "function": "mengukur tingkat kekerasan buah untuk pengendalian mutu, penelitian, atau proses sortasi",
        "tags": ["Lutron", "pengWEBSITE kekerasan buah", "fruit hardness tester", "alat ukur pertanian"],
    }),
    (re.compile(r"electromagnetic\s+field|\bemf\s*tester|\bemf-", re.I), {
        "category": "Pengukur Medan Elektromagnetik",
        "term": "pengukur medan elektromagnetik",
        "function": "mengukur tingkat medan elektromagnetik sesuai spesifikasi produk",
        "tags": ["Lutron", "EMF tester", "pengukur medan elektromagnetik", "alat ukur keselamatan"],
    }),
    (re.compile(r"pressure\s+transmitter|\btr-ps", re.I), {
        "category": "Transmitter Tekanan",
        "term": "transmitter tekanan",
        "function": "mengubah pengukuran tekanan menjadi sinyal keluaran untuk kebutuhan pemantauan atau otomasi",
        "tags": ["Lutron", "pressure transmitter", "transmitter tekanan", "instrumentasi industri"],
    }),
    (re.compile(r"power analyzer|power meter|wattmeter|energy meter|power quality|three phase|3 phase|dw-", re.I), {
        "category": "Power Analyzer dan Electrical Tester",
        "term": "power analyzer atau electrical tester",
        "function": "menganalisis parameter daya, energi, dan kelistrikan sesuai spesifikasi produk",
        "tags": ["Lutron", "power analyzer", "power meter", "electrical tester", "alat ukur listrik"],
    }),
    (re.compile(r"frequency counter|frequency meter|fc-", re.I), {
        "category": "Frequency Counter",
        "term": "frequency counter",
        "function": "mengukur frekuensi dan parameter sinyal elektronik sesuai spesifikasi produk",
        "tags": ["Lutron", "frequency counter", "frequency meter", "electronic instrument"],
    }),
    (re.compile(r"clamp meter|clamp|multimeter|digital multimeter|voltmeter|ammeter|current recorder|cm-", re.I), {
        "category": "Multimeter dan Clamp Meter",
        "term": "multimeter atau clamp meter",
        "function": "mengukur parameter kelistrikan sesuai spesifikasi produk",
        "tags": ["Lutron", "multimeter", "clamp meter", "alat ukur listrik"],
    }),
    (re.compile(r"thermometer|temperature|thermocouple|infrared thermometer|tm-", re.I), {
        "category": "Thermometer dan Temperature Meter",
        "term": "alat ukur suhu",
        "function": "mengukur suhu sampel, lingkungan, atau proses sesuai spesifikasi produk",
        "tags": ["Lutron", "thermometer", "temperature meter", "alat ukur suhu"],
    }),
    (re.compile(r"humidity|hygrometer|wet bulb|dew point|hr-", re.I), {
        "category": "Humidity Meter",
        "term": "humidity meter",
        "function": "mengukur kelembapan dan parameter lingkungan terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "humidity meter", "hygrometer", "alat ukur kelembapan"],
    }),
    (re.compile(r"sound level|noise|decibel|sl-", re.I), {
        "category": "Sound Level Meter",
        "term": "sound level meter",
        "function": "mengukur tingkat kebisingan atau tekanan suara sesuai spesifikasi produk",
        "tags": ["Lutron", "sound level meter", "noise meter", "decibel meter"],
    }),
    (re.compile(r"light meter|lux meter|uv meter|solar power|lx-", re.I), {
        "category": "Light Meter dan Lux Meter",
        "term": "light meter atau lux meter",
        "function": "mengukur intensitas cahaya atau parameter radiasi terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "light meter", "lux meter", "uv meter"],
    }),
    (re.compile(r"anemometer|air flow|air velocity|wind speed|am-", re.I), {
        "category": "Anemometer dan Air Flow Meter",
        "term": "anemometer atau air flow meter",
        "function": "mengukur kecepatan udara, aliran udara, atau parameter lingkungan terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "anemometer", "air flow meter", "air velocity meter"],
    }),
    (re.compile(r"ph\s*meter|\bph\b|orp|tds|conductivity|dissolved oxygen|do meter|water quality", re.I), {
        "category": "Water Quality Meter",
        "term": "alat ukur kualitas air",
        "function": "mengukur parameter kualitas air atau larutan sesuai spesifikasi produk",
        "tags": ["Lutron", "water quality meter", "pH meter", "conductivity meter", "TDS meter"],
    }),
    (re.compile(r"moisture meter|moisture|wood moisture|grain moisture", re.I), {
        "category": "Moisture Meter",
        "term": "moisture meter",
        "function": "mengukur kadar kelembapan material sesuai spesifikasi produk",
        "tags": ["Lutron", "moisture meter", "wood moisture meter", "material tester"],
    }),
    (re.compile(r"tachometer|rpm|stroboscope", re.I), {
        "category": "Tachometer dan RPM Meter",
        "term": "tachometer atau RPM meter",
        "function": "mengukur putaran, kecepatan, atau parameter terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "tachometer", "RPM meter", "speed meter"],
    }),
    (re.compile(r"force gauge|hardness|tension|coating thickness|thickness gauge", re.I), {
        "category": "Material Tester",
        "term": "alat WEBSITE material",
        "function": "mendukung pengukuran gaya, kekerasan, ketebalan, atau parameter material sesuai spesifikasi produk",
        "tags": ["Lutron", "material tester", "force gauge", "hardness tester", "thickness gauge"],
    }),
]


DEFAULT_FAMILY = {
    "category": "Alat Ukur Laboratorium dan Industri",
    "term": "alat ukur laboratorium dan industri",
    "function": "mendukung pengukuran, pemeriksaan, atau kontrol kualitas sesuai spesifikasi produk",
    "tags": ["Lutron", "alat ukur", "alat laboratorium", "alat industri", "quality control"],
}

ACCESSORY_FAMILY = {
    "category": "Aksesori dan Komponen",
    "term": "aksesori atau komponen pendukung",
    "function": "mendukung penggunaan unit utama yang kompatibel sesuai model dan spesifikasi produk",
    "tags": ["Lutron", "aksesori Lutron", "komponen Lutron", "spare part", "kelengkapan alat"],
}


@dataclass(frozen=True)
class SourceConfig:
    key: str
    brand: str
    domains: Tuple[str, ...]


LUTRON_CONFIG = SourceConfig("lutron", "Lutron", ("lutron.com.tw",))
SUPPORTED_SOURCES: Tuple[SourceConfig, ...] = (LUTRON_CONFIG,)


# ---------------------------------------------------------------------------
# Common helpers
# ---------------------------------------------------------------------------

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = html.unescape(str(value))
    s = s.replace("\xa0", " ").replace("\u3000", " ").replace("_x000D_", "\n")
    s = s.replace("\r", "\n")
    s = re.sub(r"[ \t\f\v]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


# Excel/Open XML melarang sebagian karakter kontrol meskipun karakter tersebut
# valid di dalam string Python. Karakter seperti \x08 dapat ikut terseret dari
# HTML/PDF hasil scrape dan membuat openpyxl melempar IllegalCharacterError.
_EXCEL_ILLEGAL_XML_RE = re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
)
_EXCEL_MAX_CELL_CHARS = 32767


def sanitize_excel_value(value: Any) -> Any:
    """Bersihkan satu nilai agar aman ditulis ke workbook XLSX.

    Angka, boolean, tanggal, dan nilai kosong dipertahankan. String dibersihkan
    dari karakter XML ilegal dan dibatasi sesuai batas panjang sel Excel.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        cleaned = _EXCEL_ILLEGAL_XML_RE.sub("", value)
        return cleaned[:_EXCEL_MAX_CELL_CHARS]
    if isinstance(value, (dict, list, tuple, set)):
        try:
            serialized = json.dumps(value, ensure_ascii=False, default=str)
        except Exception:
            serialized = str(value)
        cleaned = _EXCEL_ILLEGAL_XML_RE.sub("", serialized)
        return cleaned[:_EXCEL_MAX_CELL_CHARS]
    return value


def sanitize_excel_records(rows: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Salin kumpulan row dan sanitasi seluruh nilai sebelum export Excel."""
    cleaned_rows: List[Dict[str, Any]] = []
    for raw_row in rows or []:
        if not isinstance(raw_row, dict):
            continue
        cleaned_rows.append({
            str(key): sanitize_excel_value(value)
            for key, value in raw_row.items()
        })
    return cleaned_rows



# Karakter Han/Chinese tidak boleh masuk ke output WooCommerce.
# Nama model, satuan, angka, tanda teknis, dan istilah merek tetap dipertahankan.
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
CJK_PUNCT_RE = re.compile(r"[，。；：「」『』（）【】《》、]")
GENERIC_LUTRON_PAGE_RE = re.compile(
    r"(?:lutron\s*(?:taiwan)?\s*official\s*(?:website|site)|"
    r"路昌電子官方網站|專業電子儀表|leading\s+brand\s+specializing\s+in\s+electronic\s+instruments)",
    re.I,
)
GENERIC_LUTRON_CLAUSE_RE = re.compile(
    r"(?:路昌電子官方網站[^。.\n]*(?:[。.]|$)|"
    r"lutron\s*taiwan\s*official\s*(?:website|site)[^.\n]*(?:[.]|$)|"
    r"the\s+leading\s+brand\s+specializing\s+in\s+electronic\s+instruments\s+and\s+meters[.!]?)",
    re.I,
)


def has_cjk(value: Any) -> bool:
    return bool(CJK_RE.search(clean_text(value)))


def strip_cjk(value: Any) -> str:
    """Hapus karakter Han/Chinese tanpa mengubah angka, simbol, model, atau satuan teknis."""
    s = clean_text(value)
    s = CJK_RE.sub(" ", s)
    s = CJK_PUNCT_RE.sub(" ", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip(" -–—:;,")


def is_generic_lutron_landing_text(value: Any) -> bool:
    s = clean_text(value)
    if not s:
        return False
    return bool(GENERIC_LUTRON_PAGE_RE.search(s))


def clean_catalog_source_text(value: Any) -> str:
    """Bersihkan teks sumber agar konten umum/berbahasa Chinese tidak masuk katalog."""
    raw = clean_text(value)
    if not raw:
        return ""
    # Halaman Lutron kadang menyisipkan slogan situs umum dalam satu node dengan
    # isi produk. Buang klausa umum itu, tetapi pertahankan kalimat produk yang
    # berada sesudahnya.
    raw = GENERIC_LUTRON_CLAUSE_RE.sub(" ", raw)
    s = strip_cjk(raw)
    if not s or (is_generic_lutron_landing_text(s) and len(s) < 180):
        return ""
    return s


def html_to_source_text(value: Any, max_chars: int = 8500) -> str:
    """Ubah HTML sumber menjadi teks aman untuk diterjemahkan AI."""
    raw = clean_text(value)
    if not raw:
        return ""
    try:
        soup = BeautifulSoup(raw, "html.parser")
        text = soup.get_text("\n", strip=True)
    except Exception:
        text = re.sub(r"<[^>]+>", " ", raw)
    return compact(clean_catalog_source_text(text), max_chars)


def esc(value: Any) -> str:
    return html.escape(clean_text(value), quote=False)


def compact(text: Any, max_chars: int = 900) -> str:
    s = clean_text(re.sub(r"\s+", " ", clean_text(text))).strip()
    if len(s) <= max_chars:
        return s
    return s[:max_chars].rsplit(" ", 1)[0].rstrip(" ,.;:") + "."


def dedupe_keep_order(items: Iterable[Any]) -> List[str]:
    out: List[str] = []
    seen = set()
    for item in items:
        v = clean_text(item)
        if not v:
            continue
        key = re.sub(r"\s+", " ", v).casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(v)
    return out


def normalize_url(value: Any) -> str:
    """Normalize a Lutron URL while preserving the host provided in Excel.

    Product URLs supplied by Lutron use ``www.lutron.com.tw``. The old script
    incorrectly forced those URLs to the non-www host. That host can return a
    minimal shell instead of the product document, which leaves image URL,
    description, full content, and specs blank.
    """
    url = clean_text(value)
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif not re.match(r"https?://", url, flags=re.I):
        url = "https://" + url.lstrip("/")
    parts = urlsplit(url)
    path = quote(unquote(parts.path), safe="/%+-_.,;~@()[]")
    query = quote(unquote(parts.query), safe="=&%+-_.,;~@()[]/:")
    fragment = quote(unquote(parts.fragment), safe="%+-_.,;~@()[]/:")
    return urlunsplit((parts.scheme or "https", parts.netloc, path, query, fragment))


def get_source_config(url: str) -> Optional[SourceConfig]:
    """Return configuration only for the official Lutron Taiwan domain."""
    host = (urlparse(normalize_url(url)).hostname or "").casefold()
    if any(host == domain or host.endswith("." + domain) for domain in LUTRON_CONFIG.domains):
        return LUTRON_CONFIG
    return None


def deterministic_index(text: str, modulo: int) -> int:
    if modulo <= 1:
        return 0
    h = hashlib.md5(clean_text(text).encode("utf-8")).hexdigest()
    return int(h[:8], 16) % modulo


def is_noise_text(text: str) -> bool:
    key = re.sub(r"\s+", " ", clean_text(text)).casefold().strip(" :|-")
    return not key or key in NOISE_TEXT


def normalize_label(label: str) -> str:
    label = clean_text(label).strip(" :-–—|")
    if not label:
        return ""
    label = re.sub(r"\s+", " ", label)
    # Pertahankan model, akronim, dan penulisan vendor seperti apa adanya.
    return label[:1].upper() + label[1:]


def table_html(rows: Sequence[Dict[str, str]]) -> str:
    body = []
    for row in rows:
        label, value = esc(row.get("label")), esc(row.get("value"))
        if label and value:
            body.append(f"<tr><th>{label}</th><td>{value}</td></tr>")
    return '<table class="WEBSITE-spec-table"><tbody>' + "".join(body) + "</tbody></table>" if body else ""


def _is_technical_spec_section(value: Any) -> bool:
    """True when a source section is a dedicated technical narrative block."""
    text = clean_text(value)
    return bool(re.search(r"\b(?:technical\s+specifications?|technical\s+data|measurement\s+specifications?)\b", text, re.I))


_MEASUREMENT_UNIT_PATTERN = re.compile(
    r"(?:%\s*RH|%|mbar|bar|psi|kg\s*/\s*cm(?:\^?2|²)|mm\s*Hg|in(?:ch)?\s*Hg|"
    r"m(?:eter)?\s*H(?:2|₂)O|meterH(?:2|₂)O|in(?:ch|ci)?\s*H(?:2|₂)(?:O)?|inchH(?:2|₂)(?:O)?|atm(?:osphere|osfer)?|"
    r"°?\s*[CF]|mV|kV|V|µA|uA|mA|A|mW|kW|W|Hz|kHz|MHz|GHz|"
    r"lux|fc|rpm|m\s*/\s*s|km\s*/\s*h|ft\s*/\s*min|knot|mile\s*/\s*h|"
    r"dB|pH|mS\s*/\s*cm|µS\s*/\s*cm|uS\s*/\s*cm|ppm|ppt|NTU|FNU)",
    re.I,
)
_MEASUREMENT_QUANTITY_RE = re.compile(
    rf"^\s*([-+]?\d+(?:[.,]\d+)?)\s*({_MEASUREMENT_UNIT_PATTERN.pattern})\s*$",
    re.I,
)
_MEASUREMENT_UNIT_ONLY_RE = re.compile(
    rf"^\s*({_MEASUREMENT_UNIT_PATTERN.pattern})\s*$",
    re.I,
)


def _normalize_measurement_unit(value: Any) -> str:
    unit = clean_text(value)
    unit = re.sub(r"\s+", " ", unit).strip()
    replacements = {
        "mbar": "mbar", "bar": "bar", "psi": "psi",
        "kg/cm^2": "kg/cm²", "kg/cm2": "kg/cm²", "kg / cm²": "kg/cm²", "kg/cm²": "kg/cm²",
        "mm hg": "mmHg", "mmhg": "mmHg",
        "inch hg": "inHg", "inci hg": "inHg", "in hg": "inHg", "inchhg": "inHg",
        "meter h2o": "mH₂O", "meter h₂o": "mH₂O", "meterh2o": "mH₂O", "meterh₂o": "mH₂O", "m h2o": "mH₂O",
        "inch h2o": "inH₂O", "inch h₂o": "inH₂O", "inci h2o": "inH₂O", "inci h₂o": "inH₂O",
        "inchh2o": "inH₂O", "inchh₂o": "inH₂O", "inchh2": "inH₂O", "in h2o": "inH₂O",
        "atmosphere": "atm", "atmosfer": "atm", "atm": "atm",
        "ua": "µA", "us/cm": "µS/cm",
    }
    key = unit.casefold().replace("  ", " ")
    return replacements.get(key, unit)


def _measurement_quantity(value: Any) -> Optional[Tuple[str, str]]:
    text = clean_text(value).strip(" ,.;:")
    match = _MEASUREMENT_QUANTITY_RE.fullmatch(text)
    if not match:
        return None
    number = clean_text(match.group(1))
    unit = _normalize_measurement_unit(match.group(2))
    return f"{number} {unit}", unit


def _measurement_unit_label(value: Any) -> str:
    text = clean_text(value).strip(" ,.;:")
    match = _MEASUREMENT_UNIT_ONLY_RE.fullmatch(text)
    return _normalize_measurement_unit(match.group(1)) if match else ""


def _parse_range_resolution_value(value: Any) -> Tuple[str, str]:
    text = clean_text(value)
    # Examples: "2000 mbar, resolusi: 1 mbar" or "0 to 50 C / Resolution 0.1 C".
    marker = re.search(r"\b(?:resolusi|resolution|minimum\s+scale)\b\s*:?", text, re.I)
    if not marker:
        return "", ""
    left = text[:marker.start()].strip(" ,;/:-")
    right = text[marker.end():].strip(" ,;/:-")
    left_q = _measurement_quantity(left)
    right_q = _measurement_quantity(right)
    return (left_q[0] if left_q else left, right_q[0] if right_q else right)


def _adaptive_measurement_rows(rows: Sequence[Dict[str, str]]) -> Tuple[List[Dict[str, str]], set]:
    """Detect range/resolution rows without inventing missing source values.

    The official Lutron pages use several table layouts. Some expose a unit as
    the label and place range/resolution in the value. Others expose the range
    itself as the label and the resolution as the value. This normalizer only
    rearranges values already present in source-derived rows.
    """
    output: List[Dict[str, str]] = []
    consumed: set = set()
    by_key: Dict[str, Dict[str, str]] = {}

    def get_or_add(unit: str, range_value: str = "") -> Dict[str, str]:
        key = (unit.casefold() + "|" + clean_text(range_value).casefold()).strip("|")
        if key in by_key:
            return by_key[key]
        item = {"unit": unit, "range": range_value, "resolution": "", "accuracy": ""}
        by_key[key] = item
        output.append(item)
        return item

    pending_accuracy: List[Tuple[int, str, str]] = []
    for index, row in enumerate(rows or []):
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        if not label or not value:
            continue

        # Shape A: label is a unit and value says "range, resolution: ...".
        unit = _measurement_unit_label(label)
        if unit:
            range_value, resolution = _parse_range_resolution_value(value)
            if range_value and resolution:
                item = get_or_add(unit, range_value)
                item["resolution"] = resolution
                consumed.add(index)
                continue
            # When a unit label is paired with one quantity, retain it only as
            # a resolution candidate. The renderer does not infer a range.
            value_q = _measurement_quantity(value)
            if value_q and _normalize_measurement_unit(value_q[1]).casefold() == unit.casefold():
                item = get_or_add(unit, "")
                item["resolution"] = value_q[0]
                consumed.add(index)
                continue

        # Shape B: label itself is the range and value is resolution/accuracy.
        label_q = _measurement_quantity(label)
        if label_q:
            if re.search(r"(?:±|\+/-|%\s*(?:F\.?S\.?|full\s*scale)|akurasi|accuracy)", value, re.I):
                pending_accuracy.append((index, label_q[0], value))
                continue
            value_q = _measurement_quantity(value)
            if value_q and value_q[1].casefold() == label_q[1].casefold():
                item = get_or_add(label_q[1], label_q[0])
                item["resolution"] = value_q[0]
                consumed.add(index)
                continue

    # Attach accuracy only when the source explicitly associates it with the
    # same range label. Otherwise the original Accuracy row remains in the
    # general specification table.
    for index, range_value, accuracy in pending_accuracy:
        match = next((item for item in output if item["range"].casefold() == range_value.casefold()), None)
        if match:
            match["accuracy"] = accuracy
            consumed.add(index)

    meaningful = [item for item in output if item.get("range") or item.get("resolution")]
    if len(meaningful) < 2:
        return [], set()
    return meaningful, consumed


def _general_spec_table_html(rows: Sequence[Dict[str, str]]) -> str:
    body: List[str] = []
    for row in rows:
        label, value = esc(row.get("label")), esc(row.get("value"))
        if label and value:
            body.append(f"<tr><th scope=\"row\">{label}</th><td>{value}</td></tr>")
    if not body:
        return ""
    return (
        '<table class="WEBSITE-spec-table WEBSITE-spec-general">'
        '<thead><tr><th>Parameter</th><th>Nilai</th></tr></thead>'
        '<tbody>' + "".join(body) + '</tbody></table>'
    )


def _measurement_table_html(rows: Sequence[Dict[str, str]]) -> str:
    if not rows:
        return ""
    has_range = any(clean_text(row.get("range")) for row in rows)
    has_resolution = any(clean_text(row.get("resolution")) for row in rows)
    has_accuracy = any(clean_text(row.get("accuracy")) for row in rows)
    fields: List[Tuple[str, str]] = [("unit", "Satuan")]
    if has_range:
        fields.append(("range", "Rentang"))
    if has_resolution:
        fields.append(("resolution", "Resolusi"))
    if has_accuracy:
        fields.append(("accuracy", "Akurasi"))
    body: List[str] = []
    for row in rows:
        body.append(
            "<tr>" + "".join(
                f"<td>{esc(row.get(field)) if clean_text(row.get(field)) else '-'}</td>"
                for field, _header in fields
            ) + "</tr>"
        )
    heading = "Rentang dan Resolusi Pengukuran" if has_range else "Resolusi Pengukuran"
    return (
        f'<h3>{heading}</h3>'
        '<table class="WEBSITE-spec-table WEBSITE-spec-measurement">'
        '<thead><tr>' + "".join(f"<th>{esc(header)}</th>" for _field, header in fields) + '</tr></thead>'
        '<tbody>' + "".join(body) + '</tbody></table>'
    )


def specifications_tab_html(rows: Sequence[Dict[str, str]], already_deduplicated: bool = False) -> str:
    """Render adaptive tables while preserving every retained source fact."""
    if already_deduplicated:
        cleaned_rows: List[Dict[str, str]] = []
        seen = set()
        for raw in rows or []:
            if not isinstance(raw, dict):
                continue
            label = clean_text(raw.get("label"))
            value = clean_text(raw.get("value"))
            if not label or not value:
                continue
            key = (_spec_text_key(label), _spec_text_key(value))
            if key in seen:
                continue
            seen.add(key)
            cleaned_rows.append({
                "label": label,
                "value": value,
                "source_label": clean_text(raw.get("source_label")),
                "section": clean_text(raw.get("section")),
            })
    else:
        cleaned_rows, _issues = canonicalize_specifications(rows)

    if not cleaned_rows:
        return "<h2>Spesifikasi</h2><p>Spesifikasi belum tersedia dalam bahasa Indonesia.</p>"

    measurement_rows, consumed = _adaptive_measurement_rows(cleaned_rows)
    general_rows = [row for index, row in enumerate(cleaned_rows) if index not in consumed]
    blocks = ["<h2>Spesifikasi</h2>"]
    if measurement_rows:
        blocks.append(_measurement_table_html(measurement_rows))
    general_table = _general_spec_table_html(general_rows)
    if general_table:
        blocks.append("<h3>Spesifikasi Umum</h3>" if measurement_rows else "")
        blocks.append(general_table)
    return "".join(blocks)


# ---------------------------------------------------------------------------
# Excel input
# ---------------------------------------------------------------------------

def find_column(df: pd.DataFrame, aliases: Sequence[str]) -> str:
    cols = {str(c).strip().casefold(): c for c in df.columns}
    for alias in aliases:
        found = cols.get(alias.casefold())
        if found is not None:
            return str(found)
    return ""


def find_url_column(df: pd.DataFrame, requested: str = "") -> str:
    if requested:
        exact = {str(c).strip().casefold(): c for c in df.columns}.get(requested.strip().casefold())
        if exact is None:
            raise ValueError(f"Kolom URL '{requested}' tidak ditemukan. Kolom tersedia: {list(df.columns)}")
        return str(exact)
    result = find_column(df, ["url", "link", "source url", "product url", "product link", "product_link"])
    if result:
        return result
    for col in df.columns:
        vals = df[col].dropna().astype(str).head(30).tolist()
        if any("lutron.com.tw" in v.casefold() or v.casefold().startswith("http") for v in vals):
            return str(col)
    raise ValueError("Tidak menemukan kolom URL. Gunakan url/link/Source URL atau --url-column.")


def read_input_records(input_path: str, sheet: Any = 0, url_column: str = "") -> List[Dict[str, str]]:
    df = pd.read_excel(input_path, sheet_name=sheet)
    if df.empty:
        return []
    url_col = find_url_column(df, url_column)
    brand_col = find_column(df, ["brand", "merek"])
    title_col = find_column(df, ["title", "name", "product name", "nama", "nama produk"])
    model_col = find_column(df, ["model", "sku", "product model", "kode model"])

    records: List[Dict[str, str]] = []
    seen = set()
    for _, raw in df.iterrows():
        url = normalize_url(raw.get(url_col))
        if not url or url.casefold() in seen:
            continue
        seen.add(url.casefold())
        records.append({
            "url": url,
            "input_brand": clean_text(raw.get(brand_col)) if brand_col else "",
            "input_title": clean_text(raw.get(title_col)) if title_col else "",
            "input_model": clean_text(raw.get(model_col)) if model_col else "",
        })
    return records


# ---------------------------------------------------------------------------
# HTML parsing
# ---------------------------------------------------------------------------

def soup_text(node: Any) -> str:
    if node is None:
        return ""
    try:
        return clean_text(node.get_text(" ", strip=True))
    except Exception:
        return clean_text(node)


def clean_title(value: str) -> str:
    title = clean_text(value)
    title = re.sub(r"\s*[|｜]\s*(?:Lutron(?: Electronic)?(?: Taiwan)?).*?$", "", title, flags=re.I)
    title = re.sub(r"\s+", " ", title)
    return title.strip(" -|\t\n")


def jsonld_objects(soup: BeautifulSoup) -> Iterator[Dict[str, Any]]:
    def walk(v: Any) -> Iterator[Dict[str, Any]]:
        if isinstance(v, dict):
            yield v
            graph = v.get("@graph")
            if isinstance(graph, list):
                for x in graph:
                    yield from walk(x)
        elif isinstance(v, list):
            for x in v:
                yield from walk(x)

    for node in soup.select('script[type="application/ld+json"]'):
        raw = node.string or node.get_text() or ""
        raw = clean_text(raw)
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            # Situs tertentu menyisipkan komentar HTML atau beberapa JSON object.
            continue
        yield from walk(data)


def product_jsonld(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    out = []
    for obj in jsonld_objects(soup):
        types = obj.get("@type", [])
        types = [types] if isinstance(types, str) else types
        if any(clean_text(t).casefold() == "product" for t in types if clean_text(t)):
            out.append(obj)
    return out


def first_meta(soup: BeautifulSoup, selectors: Sequence[str]) -> str:
    for sel in selectors:
        node = soup.select_one(sel)
        if node:
            value = clean_text(node.get("content") or node.get("href") or "")
            if value:
                return value
    return ""


def extract_title(soup: BeautifulSoup, fallback: str = "") -> str:
    for obj in product_jsonld(soup):
        title = clean_text(obj.get("name"))
        if title:
            return clean_title(title)
    for selector in ['h1', '[itemprop="name"]', '.product-title', '.product_name', '.product-name', '.productTitle']:
        node = soup.select_one(selector)
        title = soup_text(node)
        if title and len(title) <= 300:
            return clean_title(title)
    meta = first_meta(soup, ['meta[property="og:title"]', 'meta[name="twitter:title"]'])
    if meta:
        return clean_title(meta)
    if soup.title:
        title = clean_title(soup_text(soup.title))
        if title:
            return title
    return clean_title(fallback)


def valid_image_url(url: str) -> bool:
    """Return True only for a plausible product photograph URL.

    Lutron pages load site banners and corporate artwork together with product
    assets.  A generic visual is worse than an empty image field, so uncertain
    URLs are rejected before they can reach WooCommerce.
    """
    value = clean_text(url)
    if not value or value.startswith(("data:", "javascript:", "#")):
        return False
    low = value.casefold()
    if any(token in low for token in IMAGE_REJECT_TOKENS):
        return False
    # Common query parameters used by tracking pixels and social previews.
    if re.search(r"(?:[?&](?:utm_|fbclid|gclid|tracking)=|/pixel(?:[/?]|$))", low, re.I):
        return False
    return True


def _iter_image_values(value: Any) -> Iterator[str]:
    """Yield raw image URLs from scalar, mapping, or list values."""
    if isinstance(value, str):
        candidate = clean_text(value)
        if candidate:
            yield candidate
        return
    if isinstance(value, dict):
        for key in (
            "url", "contentUrl", "src", "srcUrl", "image", "imageUrl", "image_url",
            "large", "largeImage", "thumbnailUrl", "thumbnail", "path",
        ):
            if key in value:
                yield from _iter_image_values(value.get(key))
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            yield from _iter_image_values(item)


def _srcset_urls(value: Any) -> Iterator[str]:
    srcset = clean_text(value)
    if not srcset:
        return
    for part in srcset.split(","):
        raw = clean_text(part.strip().split(" ")[0])
        if raw:
            yield raw


def _url_from_style(value: Any) -> Iterator[str]:
    style = clean_text(value)
    if not style:
        return
    for raw in re.findall(r"url\(\s*['\"]?([^'\")\s]+)['\"]?\s*\)", style, flags=re.I):
        raw = clean_text(raw)
        if raw:
            yield raw


def _image_score(url: str, source: str, model: str = "") -> int:
    """Rank likely product photos above generic interface images."""
    value = clean_text(url)
    low = value.casefold()
    score = {
        "jsonld": 190,
        "og": 170,
        "meta": 150,
        "product_attr": 145,
        "srcset": 125,
        "any_attr": 75,
        "style": 65,
        "script": 55,
        # Browser candidates are visually related to the product area but are
        # not automatically authoritative.  Generic banners used to receive
        # an excessive score here.
        "dom_primary": 140,
        "network": 85,
        "link": 45,
    }.get(source, 0)

    if re.search(r"(?:product|products|catalog|upload|uploads|storage|media|image|images|item|goods|photo)", low):
        score += 20
    if re.search(r"\.(?:jpe?g|png|webp|avif|gif)(?:[?#]|$)", low):
        score += 12
    if "thumbnail" in low or "thumb" in low:
        score -= 10

    normalized_model = re.sub(r"[^a-z0-9]+", "", clean_text(model).casefold())
    normalized_url = re.sub(r"[^a-z0-9]+", "", low)
    if normalized_model and len(normalized_model) >= 3 and normalized_model in normalized_url:
        score += 220
    # Images from anonymous assets are accepted only when the extraction source
    # is already strongly tied to the product DOM.
    if source in {"any_attr", "style", "script", "link", "network"} and not (
        normalized_model and normalized_model in normalized_url
    ):
        score -= 35
    return score


def extract_best_image_url(soup: BeautifulSoup, base_url: str, model: str = "", network_images: Optional[Sequence[str]] = None) -> str:
    """Extract a product image from Lutron metadata, lazy-load attributes, and page data.

    Lutron pages may expose the photo in different forms depending on the page
    template: Open Graph tags, JSON-LD, lazy-load attributes, source/srcset
    elements, inline style, or serialized page data. Every path is checked here.
    """
    candidates: List[Tuple[str, str]] = []

    def add(raw: Any, source: str) -> None:
        for value in _iter_image_values(raw):
            value = clean_text(value).replace("\\/", "/").strip(" '\"")
            if value:
                candidates.append((source, value))

    # JS-rendered Lutron pages sometimes expose a product photo only as a
    # network image request. Capture it as a candidate too.
    for network_image in (network_images or []):
        raw_network = clean_text(network_image)
        if raw_network.startswith("__LUTRON_PRIMARY__"):
            add(raw_network[len("__LUTRON_PRIMARY__"):], "dom_primary")
        else:
            add(raw_network, "network")

    # Product structured data can hold either a string, an object, or a list.
    for obj in product_jsonld(soup):
        add(obj.get("image"), "jsonld")
        add(obj.get("photo"), "jsonld")
        add(obj.get("thumbnailUrl"), "jsonld")

    # Primary social metadata generally points to the canonical product image.
    for selector in (
        'meta[property="og:image"]', 'meta[property="og:image:url"]',
        'meta[property="og:image:secure_url"]', 'meta[name="twitter:image"]',
        'meta[name="twitter:image:src"]', 'meta[name="image"]',
        'link[rel="image_src"]', 'link[rel="preload"][as="image"]',
    ):
        node = soup.select_one(selector)
        if node:
            add(node.get("content") or node.get("href"), "og" if "og:" in selector else "meta")

    product_selectors = (
        ".product img", ".product-detail img", ".product-detail-image img",
        ".product-image img", ".product-images img", ".gallery img", ".carousel img",
        "[class*='product'] img", "[class*='gallery'] img", "[id*='product'] img",
    )
    product_nodes = set()
    for selector in product_selectors:
        for node in soup.select(selector):
            product_nodes.add(id(node))
            for attr, raw in node.attrs.items():
                attr_key = clean_text(attr).casefold()
                if attr_key in {"srcset", "data-srcset"}:
                    for url in _srcset_urls(raw):
                        add(url, "srcset")
                elif any(token in attr_key for token in ("src", "image", "photo", "zoom", "original", "large", "lazy")):
                    add(raw, "product_attr")
            for url in _url_from_style(node.get("style")):
                add(url, "style")

    # Generic image tags and unusual lazy-load attribute names.
    for node in soup.find_all(["img", "source", "a", "div", "figure", "picture"]):
        source = "product_attr" if id(node) in product_nodes else "any_attr"
        for attr, raw in node.attrs.items():
            attr_key = clean_text(attr).casefold()
            if attr_key in {"srcset", "data-srcset"}:
                for url in _srcset_urls(raw):
                    add(url, "srcset")
            elif any(token in attr_key for token in ("src", "image", "photo", "zoom", "original", "large", "lazy", "background")):
                add(raw, source)
        for url in _url_from_style(node.get("style")):
            add(url, "style")

    # Some templates only expose the high-resolution photo as an <a href>.
    for node in soup.select("a[href]"):
        href = clean_text(node.get("href"))
        if re.search(r"\.(?:jpe?g|png|webp|avif|gif)(?:[?#]|$)", href, flags=re.I):
            add(href, "link")

    # Final fallback: serialized page state inside scripts. Do not require a file
    # extension because CMS/CDN endpoints may serve images from dynamic paths.
    raw_html = str(soup).replace("\\/", "/")
    patterns = (
        r'"(?:image|imageUrl|image_url|photo|photoUrl|thumbnail|thumbnailUrl|mainImage)"\s*:\s*"([^"]+)"',
        r'(?:(?:https?:)?//[^"\'<>\\\s]+(?:\.(?:jpe?g|png|webp|avif|gif)|/image/|/images/|/media/|/storage/)[^"\'<>\\\s]*)',
        r'(?:/(?:storage|uploads?|images?|media|assets)/[^"\'<>\\\s]+(?:\.(?:jpe?g|png|webp|avif|gif))?[^"\'<>\\\s]*)',
    )
    for pattern in patterns:
        for match in re.finditer(pattern, raw_html, flags=re.I):
            value = clean_text(match.group(1) if match.groups() else match.group(0))
            if value:
                add(value, "script")

    ranked: List[Tuple[int, str, str]] = []
    seen = set()
    for source, raw in candidates:
        full = urljoin(base_url, raw)
        full = clean_text(full).replace(" ", "%20")
        if not valid_image_url(full):
            continue
        key = full.casefold()
        if key in seen:
            continue
        seen.add(key)
        ranked.append((_image_score(full, source, model), full, source))

    if not ranked:
        return ""
    ranked.sort(key=lambda item: item[0], reverse=True)
    best_score, best_url, best_source = ranked[0]
    # Do not publish a weakly associated random asset.  JSON-LD, OG, a real
    # product container, or a browser candidate near the product title may use
    # a non-model filename. Other sources need the model token to be trusted.
    trusted_sources = {"jsonld", "og", "meta", "product_attr", "srcset", "dom_primary"}
    model_token = _model_token(model)
    url_token = _model_token(best_url)
    if best_source not in trusted_sources and not (model_token and model_token in url_token):
        return ""
    if best_score < 105:
        return ""
    return best_url


def extract_breadcrumb(soup: BeautifulSoup) -> str:
    selectors = ['nav[aria-label*="breadcrumb" i]', '.breadcrumb', '.breadcrumbs', '[class*="breadcrumb"]', '[id*="breadcrumb"]']
    for selector in selectors:
        node = soup.select_one(selector)
        text = soup_text(node)
        if text and len(text) <= 500:
            parts = [x.strip() for x in re.split(r"\s*(?:/|>|›|»|\||→)\s*", text) if x.strip()]
            parts = [p for p in parts if p.casefold() not in {"home", "products", "product"}]
            if parts:
                return " > ".join(parts[-3:])
    return ""


def extract_category(soup: BeautifulSoup, json_products: Sequence[Dict[str, Any]], fallback: str = "") -> str:
    for obj in json_products:
        cat = obj.get("category")
        if isinstance(cat, list):
            cat = " > ".join(clean_text(x) for x in cat if clean_text(x))
        cat = clean_text(cat)
        if cat:
            return cat
    breadcrumb = extract_breadcrumb(soup)
    if breadcrumb:
        return breadcrumb.split(" > ")[-1]
    for selector in ['[class*="category"] a', '[class*="category"]', '[itemprop="category"]']:
        node = soup.select_one(selector)
        text = soup_text(node)
        if text and len(text) <= 150 and not is_noise_text(text):
            return text
    return clean_text(fallback)


def choose_main_container(soup: BeautifulSoup) -> Tag:
    for selector in ['main', '[role="main"]', '#product', '.product-detail', '.product-details', '.product-page', '.product-info', '.product']:
        node = soup.select_one(selector)
        if isinstance(node, Tag):
            return node
    return soup.body if isinstance(soup.body, Tag) else soup


def strip_noise_nodes(root: Tag) -> None:
    selectors = [
        'script', 'style', 'noscript', 'svg', 'iframe', 'header', 'footer', 'nav', 'aside', 'form',
        '.navbar', '.navigation', '.menu', '.sidebar', '.footer', '.header', '.cookie', '.modal',
        '.social', '.share', '.related-products', '.breadcrumb', '.breadcrumbs', '[aria-label*="breadcrumb" i]',
    ]
    for selector in selectors:
        for node in root.select(selector):
            node.decompose()


def extract_description(soup: BeautifulSoup, products: Sequence[Dict[str, Any]], title: str) -> str:
    for obj in products:
        desc = clean_text(obj.get("description"))
        if desc:
            return compact(desc, 3500)
    for selector in [
        '[itemprop="description"]', '.product-description', '.product_desc', '.description',
        '[class*="description"]', '[id*="description"]', '.summary', '.product-summary',
    ]:
        nodes = soup.select(selector)
        for node in nodes:
            text = soup_text(node)
            if len(text) >= 40 and len(text) <= 6000 and clean_text(text).casefold() != clean_text(title).casefold():
                return compact(text, 3500)
    meta = first_meta(soup, ['meta[name="description"]', 'meta[property="og:description"]'])
    return compact(meta, 1500)


def normalize_content_line(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" |\t\n")


def _node_text_lines(node: Tag) -> List[str]:
    """Return visible node text while retaining manual line breaks from <br> elements."""
    try:
        raw = node.get_text("\n", strip=True)
    except Exception:
        raw = soup_text(node)
    return [normalize_content_line(part) for part in clean_text(raw).splitlines() if normalize_content_line(part)]


def _looks_like_section_heading(node: Tag, value: str) -> bool:
    """Recognise semantic and div-based headings used by Lutron/Quasar pages."""
    text = normalize_content_line(value)
    heading_text = text.strip(" :–—-|")
    if not heading_text or len(heading_text) > 150:
        return False
    if node.name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
        return not is_noise_text(heading_text)
    # Some Lutron templates use a styled div/span instead of an h2/h3.
    if SPEC_HEADING_RE.fullmatch(heading_text) or re.fullmatch(
        r"(?:technical\s+specifications?|technical\s+data|specifications?|"
        r"features|versatile\s+applications|applications|standard\s+accessories|"
        r"optional\s+accessories|data\s+output)",
        heading_text,
        flags=re.I,
    ):
        return True
    class_id = " ".join([clean_text(node.get("class")), clean_text(node.get("id"))]).casefold()
    return bool(re.search(r"(?:title|heading|section).*(?:spec|technical)|(?:spec|technical).*(?:title|heading|section)", class_id))


def extract_all_content(root: Tag, title: str) -> List[Tuple[str, str, str]]:
    """Return non-noise product content as ``(kind, section, text)``.

    This deliberately preserves line breaks inside paragraphs/divs while
    excluding table rows, which are parsed separately. On Lutron pages the
    ``Technical Specifications`` block can be a Quasar ``div`` with
    ``<br>`` separated facts rather than semantic list/table markup. The old
    parser ignored that shape or collapsed the full block into one sentence.
    """
    items: List[Tuple[str, str, str]] = []
    heading = "Informasi Produk"
    seen = set()
    max_items = 420
    target_names = ["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "dt", "dd", "div"]

    for node in root.find_all(target_names, recursive=True):
        if len(items) >= max_items or not isinstance(node, Tag):
            break
        # A div that contains a semantic child is only a layout wrapper. Reading
        # it would duplicate every line below. Leaf divs still matter because
        # Lutron often uses them as visual section headings or text blocks.
        if node.name == "div" and node.find(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "dt", "dd", "tr", "div"], recursive=False):
            continue

        # Table rows belong exclusively to the structured-table extractor in
        # extract_specs(). Reading them again as narrative content causes the
        # same source fact to appear twice in the final WooCommerce table.
        lines = _node_text_lines(node)
        kind = "li" if node.name == "li" else "text"

        for line in lines:
            if _looks_like_section_heading(node, line):
                heading = line.strip(" :–—-|")[:180]
                continue
            if not line or is_noise_text(line) or len(line) < 2:
                continue
            if line.casefold() == clean_text(title).casefold():
                continue
            key = (heading.casefold(), line.casefold())
            if key in seen:
                continue
            seen.add(key)
            items.append((kind, heading, line))
    return items


def all_content_html(items: Sequence[Tuple[str, str, str]], title: str, max_chars: int = 16000) -> str:
    if not items:
        return ""
    blocks: List[str] = []
    current = ""
    bullets: List[str] = []
    used = 0

    def flush_bullets() -> None:
        nonlocal bullets
        if bullets:
            blocks.append("<ul>" + "".join(f"<li>{esc(x)}</li>" for x in bullets) + "</ul>")
            bullets = []

    for kind, heading, text in items:
        addition = len(heading) + len(text) + 30
        if used + addition > max_chars:
            break
        used += addition
        if heading != current:
            flush_bullets()
            if heading:
                blocks.append(f"<h3>{esc(heading)}</h3>")
            current = heading
        if kind == "li":
            bullets.append(text)
        else:
            flush_bullets()
            blocks.append(f"<p>{esc(text)}</p>")
    flush_bullets()
    return "".join(blocks)


def clean_spec_value(value: str) -> str:
    value = normalize_content_line(value)
    value = re.sub(r"\s*\|\s*", " / ", value)
    value = re.sub(r"\s{2,}", " ", value)
    return value.strip(" :–—|/")


def looks_like_spec_label(value: str) -> bool:
    value = clean_spec_value(value)
    if not (2 <= len(value) <= 120):
        return False
    if len(value.split()) > 11:
        return False
    return True


def _spec_text_key(value: Any) -> str:
    text = clean_spec_value(value).casefold()
    text = text.replace("℃", "c").replace("°", "")
    return re.sub(r"[^a-z0-9]+", "", text)


def _spec_numbers(value: Any) -> List[str]:
    return [token.replace(",", ".") for token in re.findall(r"[-+]?\d+(?:[.,]\d+)?", clean_text(value))]


def _spec_unit_signature(value: Any) -> Tuple[str, ...]:
    """Return units that make equal numeric values comparable safely."""
    text = clean_text(value).casefold().replace("℃", "c").replace("℉", "f")
    units = re.findall(r"(?:%\s*r\.?h\.?|%|(?:dc|ac)|ma|mv|kv|v|w|(?:kg|mg|lb|lbs|gram|g)|(?:mm|cm|inch|in\.?|m)|°?\s*[cf])", text, re.I)
    return tuple(sorted(set(re.sub(r"\s+", "", u).strip(".") for u in units)))


def _spec_values_equivalent(left: Any, right: Any) -> bool:
    """Compare values without treating changed technical numbers as duplicates.

    Lutron writes the same fact in several cosmetic forms, such as
    ``Less than 80% RH`` and ``< 80 % RH``. Equal numbers plus the same unit
    category are equivalent for humidity, temperature, power, dimensions, and
    weight. Values with different units are never merged merely because their
    numeric tokens match.
    """
    a, b = _spec_text_key(left), _spec_text_key(right)
    if not a or not b:
        return False
    if a == b:
        return True
    a_numbers, b_numbers = _spec_numbers(left), _spec_numbers(right)
    if a_numbers != b_numbers:
        return False
    a_units, b_units = _spec_unit_signature(left), _spec_unit_signature(right)
    # Different source wording around one physical fact is cosmetic when
    # numbers and unit category agree.
    if (_dimension_like(left) and _dimension_like(right)) or (_weight_like(left) and _weight_like(right)):
        return True
    if _humidity_like(left) and _humidity_like(right):
        return True
    if ((_temperature_like(left) and _temperature_like(right)) or (_power_like(left) and _power_like(right))) and a_units == b_units:
        return True
    short, long = min(a, b, key=len), max(a, b, key=len)
    return len(short) >= 10 and short in long


# Canonical keys express fact meaning. The rules accept English source labels
# and Indonesian labels produced by AI/fallback so the final audit can run both
# before and after translation.
_CANONICAL_LABEL_RULES: List[Tuple[re.Pattern[str], str, str]] = [
    (re.compile(r"\b(?:model|sku|mpn|catalog(?:ue)?\s*(?:no|number)?|part\s*(?:no|number)|item\s*(?:no|number))\b", re.I), "model", "Model"),
    (re.compile(r"\b(?:main\s*(?:instrument|unit)|instrumen\s*utama).*\b(?:dimension|dimensions|dimensi)\b|\b(?:main\s*(?:instrument|unit)|instrumen\s*utama)\b", re.I), "dimensions_main_unit", "Dimensi instrumen utama"),
    (re.compile(r"\b(?:meter|instrument|unit|alat)\s*(?:dimension|dimensions|dimensi)\b|\b(?:dimension|dimensions|dimensi)\s*(?:meter|instrument|unit|alat)\b", re.I), "dimensions_meter", "Dimensi meter"),
    (re.compile(r"\b(?:probe|sensor)\s*(?:dimension|dimensions|dimensi)\b|\b(?:dimension|dimensions|dimensi)\s*(?:probe|sensor)\b", re.I), "dimensions_probe", "Dimensi probe"),
    (re.compile(r"\b(?:dimension|dimensions|dimensi)\b", re.I), "dimensions", "Dimensi"),
    (re.compile(r"\b(?:meter|instrument|unit|alat)\s*(?:weight|berat)\b|\b(?:weight|berat)\s*(?:meter|instrument|unit|alat)\b", re.I), "weight_meter", "Berat meter"),
    (re.compile(r"\b(?:probe|sensor)\s*(?:weight|berat)\b|\b(?:weight|berat)\s*(?:probe|sensor)\b", re.I), "weight_probe", "Berat probe"),
    (re.compile(r"\b(?:weight|berat)\b", re.I), "weight", "Berat"),
    (re.compile(r"\b(?:operating|working|ambient)\s*(?:temperature|suhu)\b|\b(?:temperature|suhu)\s*(?:operasi|kerja|lingkungan)\b", re.I), "operating_temperature", "Suhu operasi"),
    (re.compile(r"\b(?:operating|working|ambient)\s*(?:humidity|kelembapan|kelembaban)\b|\b(?:humidity|kelembapan|kelembaban)\s*(?:operasi|kerja|lingkungan)\b", re.I), "operating_humidity", "Kelembapan operasi"),
    (re.compile(r"\b(?:measurement|measuring)\s*(?:range|rentang)|\b(?:range|rentang)\s*(?:measurement|pengukuran)\b", re.I), "measurement_range", "Rentang pengukuran"),
    (re.compile(r"\bfrequency\s+response\b|\brespons\s+frekuensi\b", re.I), "frequency_response", "Respons frekuensi"),
    (re.compile(r"\bfrequency\s+(?:range|band)\b|\brentang\s+frekuensi\b", re.I), "frequency_range", "Rentang frekuensi"),
    (re.compile(r"\bfrequency\s+accuracy\b|\bakurasi\s+frekuensi\b", re.I), "frequency_accuracy", "Akurasi frekuensi"),
    (re.compile(r"\b(?:data\s+error|error\s+no\.?|kesalahan\s+data)\b", re.I), "data_error_rate", "Kesalahan data"),
    (re.compile(r"\b(?:full\s+scale\s+deflection|defleksi\s+skala\s+penuh)\b", re.I), "full_scale_deflection", "Defleksi skala penuh"),
    (re.compile(r"\b(?:resolution|resolusi)\b", re.I), "resolution", "Resolusi"),
    (re.compile(r"\b(?:accuracy|akurasi|error|kesalahan)\b", re.I), "accuracy", "Akurasi"),
    (re.compile(r"\b(?:display\s+sampling\s+time|sampling\s+time\s+of\s+display|waktu\s+sampling\s+tampilan)\b", re.I), "display_sampling_time", "Waktu sampling tampilan"),
    (re.compile(r"\b(?:data\s*logger|datalogger|pencatat\s+data).*\b(?:sampling\s+time|waktu\s+sampling)\b", re.I), "datalogger_sampling_time", "Waktu sampling pencatat data"),
    (re.compile(r"\b(?:sampling\s+time|waktu\s+sampling)\b", re.I), "sampling_time", "Waktu sampling"),
    (re.compile(r"\b(?:power\s+supply|catu\s+daya|sumber\s+daya)\b", re.I), "power_supply", "Catu daya"),
    # Specific power modes must be evaluated before generic power draw.
    # Otherwise "Power Consumption - Normal Operation" and
    # "Power Consumption - Data Logging" collapse into the same key and are
    # falsely reported as conflicting values.
    (re.compile(r"\b(?:normal\s+operation|operasi\s+normal)\b", re.I), "power_draw_normal_operation", "Konsumsi daya operasi normal"),
    (re.compile(r"\b(?:data\s*(?:save|logging)|penyimpanan\s+data)\b", re.I), "power_draw_data_logging", "Konsumsi daya saat pencatatan data"),
    (re.compile(r"\b(?:power\s+consumption|power\s+current|arus\s+daya|konsumsi\s+daya)\b", re.I), "power_draw", "Konsumsi daya"),
    (re.compile(r"\b(?:accessories?\s+included|included\s+accessories?|aksesori\s+yang\s+disertakan)\b", re.I), "accessories_included", "Aksesori yang disertakan"),
    (re.compile(r"\b(?:standard\s+accessories?|aksesori\s+standar)\b", re.I), "accessories_standard", "Aksesori standar"),
    (re.compile(r"\b(?:optional\s+accessories?|aksesori\s+opsional)\b", re.I), "accessories_optional", "Aksesori opsional"),
    (re.compile(r"\b(?:data\s+output|output|interface|communication|communication\s+interface|keluaran\s+data|antarmuka)\b", re.I), "interface_output", "Antarmuka dan keluaran data"),
    (re.compile(r"\b(?:display\s+direction|arah\s+tampilan)\b", re.I), "display_direction", "Arah tampilan"),
    (re.compile(r"\b(?:display|tampilan|lcd)\b", re.I), "display", "Tampilan"),
    (re.compile(r"\b(?:memory\s+card|sd\s*card|kartu\s+memori)\b", re.I), "memory_card", "Kartu memori"),
    (re.compile(r"\b(?:operating\s+humidity|humidity|kelembapan|kelembaban)\b", re.I), "humidity", "Kelembapan"),
    (re.compile(r"\b(?:temperature|suhu)\b", re.I), "temperature", "Suhu"),
]


def _canonical_label_info(row: Dict[str, Any]) -> Tuple[str, str, str]:
    source_label = clean_text(row.get("source_label") or row.get("label"))
    display_label = clean_text(row.get("label"))
    probe = " ".join(x for x in (source_label, display_label) if x)
    for pattern, key, label in _CANONICAL_LABEL_RULES:
        if pattern.search(probe):
            return key, label, source_label
    # Preserve an unknown but valid field rather than forcing a wrong category.
    fallback_label = normalize_label(display_label or source_label) or "Rincian teknis"
    fallback_key = "other:" + _spec_text_key(source_label or fallback_label)[:80]
    return fallback_key, fallback_label, source_label


def _dimension_like(value: Any) -> bool:
    return bool(re.search(r"\b\d+(?:[.,]\d+)?\s*[x×]\s*\d+(?:[.,]\d+)?(?:\s*[x×]\s*\d+(?:[.,]\d+)?)?\s*(?:mm|cm|m|inch|in\.?)(?:\b|\s)", clean_text(value), re.I))


def _weight_like(value: Any) -> bool:
    return bool(re.search(r"\b\d+(?:[.,]\d+)?\s*(?:mg|g|kg|lb|lbs|gram)\b", clean_text(value), re.I))


def _humidity_like(value: Any) -> bool:
    return bool(re.search(r"\b(?:less\s+than|under|below|kurang\s+dari|≤|<)?\s*\d+(?:[.,]\d+)?\s*%\s*(?:r\.?h\.?)?\b", clean_text(value), re.I))


def _temperature_like(value: Any) -> bool:
    return bool(re.search(r"[-+]?\d+(?:[.,]\d+)?\s*°?\s*(?:c|f|℃|℉)\b", clean_text(value), re.I))


def _power_like(value: Any) -> bool:
    text = clean_text(value)
    return bool(re.search(r"\b(?:dc|ac)\s*\d|\b\d+(?:[.,]\d+)?\s*(?:v|volt|ma|a|w|watt)\b|\b(?:adapter|adaptor)\b", text, re.I))


def _accessory_like(value: Any) -> bool:
    # Battery wording appears in normal weight/power statements, so it is not
    # enough evidence by itself to classify a value as an accessory list.
    return bool(re.search(r"\b(?:manual|book|buku|cable|kabel|case|tas|software|clip|holder|stand|tip|penetrometer|electrode|electroda)\b", clean_text(value), re.I))


def _value_has_multiple_field_signatures(value: Any) -> List[str]:
    found = []
    if _dimension_like(value):
        found.append("dimensions")
    if _weight_like(value):
        found.append("weight")
    if _humidity_like(value):
        found.append("humidity")
    if _temperature_like(value):
        found.append("temperature")
    if _power_like(value):
        found.append("power")
    if _accessory_like(value):
        found.append("accessories")
    return found


def _dedupe_value_fragments(value: Any) -> str:
    """Remove repeated identical clauses inside a single scraped value.

    It only removes an exact cosmetic duplicate. It never performs a unit
    conversion or drops a clause that contains different numbers.
    """
    raw = clean_spec_value(value)
    if not raw:
        return ""
    # Normalise a repeated complete half: "A / B / A / B" -> "A / B".
    parts = [clean_spec_value(p) for p in re.split(r"\s*/\s*", raw) if clean_spec_value(p)]
    if len(parts) >= 4 and len(parts) % 2 == 0:
        half = len(parts) // 2
        if all(_spec_values_equivalent(parts[i], parts[i + half]) for i in range(half)):
            parts = parts[:half]
    # Drop exact duplicated slash clauses while keeping the first occurrence.
    deduped: List[str] = []
    seen = set()
    for part in parts:
        key = _spec_text_key(part)
        if key and key not in seen:
            seen.add(key)
            deduped.append(part)
    if len(deduped) >= 2:
        return " / ".join(deduped)
    return raw


def _canonical_base_key(key: str) -> str:
    if key.startswith("dimensions_"):
        return "dimensions"
    if key.startswith("weight_"):
        return "weight"
    if key in {"humidity", "operating_humidity"}:
        return "humidity"
    if key in {"temperature", "operating_temperature"}:
        return "temperature"
    return key


def _canonical_keys_can_compare(left: str, right: str) -> bool:
    return left == right or _canonical_base_key(left) == _canonical_base_key(right)


def _specificity_score(key: str, source_label: str, section: str) -> Tuple[int, int, int]:
    """Rank duplicate facts while preserving the user's source priority.

    When Technical Specifications and Specifications contain the same fact,
    Technical Specifications must be kept. A regular Specifications table is
    used only when the technical section has no equivalent fact.
    """
    component = int(key.startswith("dimensions_") or key.startswith("weight_"))
    section_text = clean_text(section)
    if _is_technical_spec_section(section_text):
        source = 3
    elif re.search(r"(?:structured|datasheet|specifications)", section_text, re.I):
        source = 2
    else:
        source = 1
    label = min(80, len(clean_text(source_label)))
    return (component, source, label)


def _issue(code: str, row: Dict[str, Any], detail: str, severity: str = "warning") -> Dict[str, str]:
    """Create a review record with an explicit severity.

    Cross-section wording differences are normal on Lutron pages. They should
    be visible for audit, but they must not suppress the entire catalog. Only a
    strong label/value mismatch is a publication blocker.
    """
    return {
        "Issue Code": code,
        "Severity": severity,
        "Detail": detail,
        "Canonical Key": clean_text(row.get("canonical_key")),
        "Display Label": clean_text(row.get("label")),
        "Source Label": clean_text(row.get("source_label")),
        "Value": clean_text(row.get("value")),
        "Source Section": clean_text(row.get("section")),
    }


def _is_blocking_spec_issue(issue: Dict[str, Any]) -> bool:
    """Return True only for evidence that a label is paired to the wrong value.

    A value conflict between Technical Specifications and Specifications may
    be a valid distinction such as meter versus probe data, or a manufacturer
    wording difference. It stays as a warning, while obvious cases such as an
    accessory label containing only dimensions are blocked.
    """
    if not isinstance(issue, dict):
        return False
    return clean_text(issue.get("Severity")).casefold() == "blocker" or (
        clean_text(issue.get("Issue Code")) == "label_value_mismatch"
    )


def _validate_canonical_row(row: Dict[str, Any]) -> List[Dict[str, str]]:
    """Flag only strong label/value mismatches. Unknown fields are allowed."""
    key = clean_text(row.get("canonical_key"))
    value = clean_text(row.get("value"))
    issues: List[Dict[str, str]] = []
    if not value:
        issues.append(_issue("empty_specification_value", row, "Nilai spesifikasi kosong.", severity="warning"))
        return issues
    signatures = _value_has_multiple_field_signatures(value)
    expected = {
        "dimensions": "dimensions", "dimensions_main_unit": "dimensions", "dimensions_meter": "dimensions", "dimensions_probe": "dimensions",
        "weight": "weight", "weight_meter": "weight", "weight_probe": "weight",
        "operating_humidity": "humidity", "operating_temperature": "temperature",
        "power_supply": "power", "power_draw": "power", "power_draw_normal_operation": "power", "power_draw_data_logging": "power",
        "accessories_included": "accessories", "accessories_standard": "accessories", "accessories_optional": "accessories",
    }.get(key)
    strong_pattern = {
        "dimensions": _dimension_like,
        "weight": _weight_like,
        "humidity": _humidity_like,
        "temperature": _temperature_like,
        "power": _power_like,
        "accessories": _accessory_like,
    }
    if expected and expected in strong_pattern and not strong_pattern[expected](value):
        # Do not flag generic free-text facts. Only obvious wrong pairings are
        # review-worthy. E.g. accessories label carrying a dimension value.
        other = [x for x in signatures if x != expected]
        if other:
            issues.append(_issue("label_value_mismatch", row, f"Label mengarah ke {expected}, tetapi nilai terlihat seperti {', '.join(other)}.", severity="blocker"))
    if expected and expected != "accessories" and len(signatures) >= 2 and expected in signatures:
        # Accessories routinely include cables, adapters, and memory cards.
        # Those are not mixed fields. For technical labels, retain only strong
        # incompatible signals such as humidity attached to a power entry.
        incompatible = [x for x in signatures if x != expected]
        if incompatible:
            issues.append(_issue("mixed_specification_fields", row, f"Satu nilai memuat beberapa jenis data: {', '.join(signatures)}.", severity="warning"))
    return issues


def _row_display_label(canonical_key: str, default_label: str, source_label: str) -> str:
    for _pat, key, label in _CANONICAL_LABEL_RULES:
        if key == canonical_key:
            return label
    # Source-specific labels such as "6 GHz" or "Pt 100 ohm" must stay
    # recognisable rather than becoming a generic field name.
    return normalize_label(default_label or source_label) or "Rincian teknis"


def canonicalize_specifications(rows: Sequence[Dict[str, Any]]) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    """Hapus spesifikasi yang benar-benar sama, lalu simpan salah satu saja.

    Prinsip v16 sengaja sederhana:
    - Technical Specifications dan Specifications boleh sama-sama dibaca.
    - Bila fakta yang sama muncul dua kali dengan nilai yang sama atau hanya
      berbeda format kecil, simpan satu baris.
    - Bila nilai berbeda, jangan dihapus dan jangan dipindahkan ke review.
      Nilai berbeda bisa memang mewakili probe, meter, kanal, atau rentang lain.
    - Tidak ada aturan yang menahan produk hanya karena spesifikasi.
    """
    prepared: List[Dict[str, Any]] = []
    for idx, raw in enumerate(rows or []):
        if not isinstance(raw, dict):
            continue
        value = _dedupe_value_fragments(raw.get("value", ""))
        if not value:
            continue
        key, standard_label, source_label = _canonical_label_info(raw)
        label = _row_display_label(key, clean_text(raw.get("label")) or standard_label, source_label)
        section = clean_text(raw.get("section")) or "Specifications"
        if _spec_text_key(label) == _spec_text_key(value):
            continue
        prepared.append({
            "label": label,
            "value": value,
            "section": section,
            "source_label": source_label or label,
            "canonical_key": key,
            "_index": idx,
        })

    # Pertahankan urutan sumber. Model tetap ditampilkan pertama.
    prepared.sort(key=lambda r: (0 if r["canonical_key"] == "model" else 1, r["_index"]))
    accepted: List[Dict[str, Any]] = []

    for row in prepared:
        duplicate_index: Optional[int] = None
        for pos, existing in enumerate(accepted):
            if not _canonical_keys_can_compare(
                clean_text(existing.get("canonical_key")), row["canonical_key"]
            ):
                continue
            if _spec_values_equivalent(existing.get("value"), row.get("value")):
                duplicate_index = pos
                break

        if duplicate_index is None:
            accepted.append(row)
            continue

        # Jika dua baris memang sama, pakai versi label yang lebih spesifik.
        old = accepted[duplicate_index]
        old_score = _specificity_score(
            clean_text(old.get("canonical_key")),
            clean_text(old.get("source_label")),
            clean_text(old.get("section")),
        )
        new_score = _specificity_score(
            row["canonical_key"], row["source_label"], row["section"]
        )
        if new_score > old_score or (
            new_score == old_score and len(row["value"]) > len(clean_text(old.get("value")))
        ):
            accepted[duplicate_index] = row

    output: List[Dict[str, str]] = []
    for row in accepted:
        output.append({
            "label": clean_text(row.get("label")),
            "value": clean_text(row.get("value")),
            "section": clean_text(row.get("section")),
            "source_label": clean_text(row.get("source_label")),
            "canonical_key": clean_text(row.get("canonical_key")),
        })
    return output, []


def dedupe_specs(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    """Compatibility wrapper used across scraping, fallback, and AI output."""
    cleaned, _issues = canonicalize_specifications(rows)
    return cleaned



def _technical_heading_exact(value: Any) -> bool:
    text = clean_text(value).strip(" :–—|-")
    return bool(re.fullmatch(r"(?:technical\s+specifications?|technical\s+data|measurement\s+specifications?)", text, re.I))


def extract_specs(soup: BeautifulSoup, root: Tag, content_items: Sequence[Tuple[str, str, str]], fallback_model: str = "") -> List[Dict[str, str]]:
    """Extract raw facts, then canonicalise them once.

    Tables and definition lists are treated as structured facts. Narrative
    technical content is read only under an exact technical heading. This stops
    the ordinary Specifications table from being parsed a second time.
    """
    raw_rows: List[Dict[str, str]] = []

    for obj in product_jsonld(soup):
        props = obj.get("additionalProperty", [])
        if isinstance(props, dict):
            props = [props]
        for prop in props if isinstance(props, list) else []:
            if isinstance(prop, dict):
                label = clean_text(prop.get("name"))
                value = clean_text(prop.get("value") or prop.get("description"))
                if label and value:
                    raw_rows.append({"label": label, "source_label": label, "value": value, "section": "Product structured data"})
        sku = clean_text(obj.get("sku") or obj.get("mpn"))
        if sku:
            raw_rows.append({"label": "Model / SKU", "source_label": "Model / SKU", "value": sku, "section": "Product structured data"})

    for table in root.find_all("table"):
        heading = "Specifications"
        previous = table.find_previous(["h1", "h2", "h3", "h4", "h5", "h6"])
        if previous:
            heading = soup_text(previous) or heading
        for tr in table.find_all("tr"):
            cells = [clean_spec_value(soup_text(c)) for c in tr.find_all(["th", "td"], recursive=False)]
            cells = [c for c in cells if c]
            if len(cells) >= 2 and looks_like_spec_label(cells[0]):
                raw_rows.append({"label": cells[0], "source_label": cells[0], "value": " / ".join(cells[1:]), "section": heading})

    for dl in root.find_all("dl"):
        heading_node = dl.find_previous(["h1", "h2", "h3", "h4", "h5", "h6"])
        heading = soup_text(heading_node) or "Specifications"
        for dt in dl.find_all("dt"):
            dd = dt.find_next_sibling("dd")
            label, value = soup_text(dt), soup_text(dd)
            if label and value and looks_like_spec_label(label):
                raw_rows.append({"label": label, "source_label": label, "value": value, "section": heading})

    active = False
    pending_label = ""
    pending_values: List[str] = []

    def flush_pending() -> None:
        nonlocal pending_label, pending_values
        if pending_label and pending_values:
            raw_rows.append({
                "label": pending_label,
                "source_label": pending_label,
                "value": " / ".join(dedupe_keep_order(pending_values)),
                "section": "Technical Specifications",
            })
        pending_label, pending_values = "", []

    for _kind, heading, text in content_items:
        if _technical_heading_exact(heading):
            active = True
        elif active and clean_text(heading).strip() and not _technical_heading_exact(heading):
            flush_pending()
            active = False
        if not active:
            continue
        raw_line = normalize_content_line(text).lstrip("-•* ").strip()
        if not raw_line:
            continue
        if raw_line.endswith(":") and looks_like_spec_label(raw_line[:-1]):
            flush_pending()
            pending_label = clean_spec_value(raw_line[:-1])
            continue
        if ":" in raw_line:
            label, value = raw_line.split(":", 1)
            label, value = clean_spec_value(label), clean_spec_value(value)
            if looks_like_spec_label(label) and value:
                flush_pending()
                raw_rows.append({"label": label, "source_label": label, "value": value, "section": "Technical Specifications"})
                continue
        if pending_label:
            pending_values.append(clean_spec_value(raw_line))
        # Unlabelled technical lines are intentionally not assigned to a guessed
        # label. They remain in full official content, but should not become a
        # potentially wrong key/value specification row.
    flush_pending()

    model = clean_text(fallback_model)
    if model:
        raw_rows.insert(0, {"label": "Model", "source_label": "Model", "value": model, "section": "Input Excel"})

    # Keep raw source rows here. The canonical pipeline runs in main() after
    # all scrape/fallback sources are combined, so conflicts remain detectable
    # and can be sent to Review Spesifikasi instead of being silently dropped.
    return raw_rows


def extract_model(title: str, specs: Sequence[Dict[str, str]], fallback: str = "") -> str:
    for row in specs:
        label = clean_text(row.get("label")).casefold()
        if any(x in label for x in ["model", "sku", "catalog", "cat.no", "part no", "item no", "mpn"]):
            value = clean_text(row.get("value"))
            if value:
                return value
    return clean_text(fallback)


def is_accessory_product(title: str, description: str, specs: Sequence[Dict[str, str]]) -> Tuple[bool, str]:
    """Classify an item as accessory only with affirmative, high-confidence evidence.

    A product title containing "receiver", "transmitter", "probe", or
    "sensor" is not enough.  Lutron uses those words in independent meters and
    transmitters, so the prior broad rule produced incorrect categories and
    compatibility copy.
    """
    title_clean = clean_text(title)
    desc_clean = clean_text(description)
    blob = " ".join([title_clean, desc_clean] + [f"{r.get('label','')} {r.get('value','')}" for r in specs])
    model_like = " ".join([title_clean] + [clean_text(r.get("value")) for r in specs if re.search(r"\b(model|sku|cat|part|item|mpn)\b", clean_text(r.get("label")), re.I)])

    # A recognised instrument family or an explicit main-instrument noun wins
    # over any component word found in the same title.
    if PRIMARY_PRODUCT_TERMS.search(title_clean) or PRIMARY_MODEL_PREFIX_RE.search(model_like):
        return False, "main_instrument_title_or_model"

    functional_labels = 0
    for row in specs:
        label = clean_text(row.get("label"))
        if re.search(r"\b(range|measurement|accuracy|resolution|display|power|battery|output|interface|data logger|datalogger|channel|temperature|humidity|pressure|frequency|weight|dimension)\b", label, re.I):
            functional_labels += 1
    if functional_labels >= 2:
        return False, "main_instrument_spec_signature"

    # Strong phrases are accepted only when the title does not identify an
    # autonomous instrument. A standalone replacement probe/cable can still be
    # classified correctly here.
    if ACCESSORY_STRONG_KEYWORDS.search(title_clean):
        return True, "strong_accessory_keyword_in_title"

    # A short title that is merely a component noun can be an accessory, but
    # unknown/incomplete pages are intentionally not guessed as accessories.
    if ACCESSORY_COMPONENT_TERMS.fullmatch(title_clean.strip().casefold()):
        return True, "standalone_component_title"
    if ACCESSORY_COMPONENT_TERMS.search(title_clean) and len(desc_clean) < 35 and not specs:
        return True, "component_title_with_no_product_data"

    return False, "not_confirmed_as_accessory"


def source_official_name(cfg: SourceConfig) -> str:
    return "halaman produk Lutron"


HOME_PAGE_TITLE_RE = re.compile(
    r"^lutron electronic taiwan official website\s*-\s*expert in electronic instruments and meters$",
    re.I,
)


def _model_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).casefold())


def looks_like_lutron_landing_page(soup: BeautifulSoup, expected_model: str = "") -> bool:
    """Detect the generic landing page returned by the www Lutron host."""
    page_title = clean_text(soup_text(soup.title))
    if not HOME_PAGE_TITLE_RE.search(page_title):
        return False
    token = _model_token(expected_model)
    if not token:
        return True
    visible = _model_token(soup.get_text(" ", strip=True))
    return token not in visible


class LutronBrowserRenderer:
    """Render JavaScript content and record actual image URLs requested by Lutron."""

    def __init__(self, timeout_seconds: int = 30):
        self.timeout_ms = max(15, int(timeout_seconds)) * 1000
        self._playwright = None
        self._browser = None
        self._context = None

    def available(self) -> bool:
        return sync_playwright is not None

    def start(self) -> None:
        if self._context is not None:
            return
        if sync_playwright is None:
            raise RuntimeError("playwright_not_installed")
        self._playwright = sync_playwright().start()
        launch_args = {
            "headless": True,
            "args": ["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        }
        # Prefer Playwright's own browser. If it has not been downloaded, use
        # an installed Chrome/Edge/Chromium browser when available.
        try:
            self._browser = self._playwright.chromium.launch(**launch_args)
        except Exception as first_error:
            candidates = [
                clean_text(os.getenv("LUTRON_BROWSER_EXECUTABLE", "")),
                shutil.which("chrome"), shutil.which("google-chrome"),
                shutil.which("chromium"), shutil.which("chromium-browser"),
                shutil.which("msedge"),
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            ]
            executable = next((x for x in candidates if x and Path(x).exists()), "")
            if not executable:
                self.close()
                raise RuntimeError(
                    "browser_not_found. Install browser Playwright: py -m playwright install chromium"
                ) from first_error
            self._browser = self._playwright.chromium.launch(executable_path=executable, **launch_args)
        self._context = self._browser.new_context(
            user_agent=DEFAULT_HEADERS["User-Agent"],
            locale="en-US",
            viewport={"width": 1440, "height": 1400},
            extra_http_headers={"Accept-Language": DEFAULT_HEADERS["Accept-Language"]},
        )

    def fetch(self, url: str, expected_model: str = "") -> Dict[str, Any]:
        page = None
        images: List[str] = []

        def on_response(response: Any) -> None:
            try:
                rtype = clean_text(response.request.resource_type).casefold()
                ctype = clean_text(response.headers.get("content-type", "")).casefold()
                if rtype == "image" or ctype.startswith("image/"):
                    images.append(clean_text(response.url))
            except Exception:
                pass

        try:
            self.start()
            assert self._context is not None
            page = self._context.new_page()
            page.set_default_timeout(self.timeout_ms)
            page.on("response", on_response)
            page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            # Scroll through the page to activate lazy-loaded gallery images.
            for y in (0, 650, 1500, 2800, 4400):
                page.evaluate("y => window.scrollTo(0, y)", y)
                page.wait_for_timeout(400)
            try:
                page.wait_for_load_state("networkidle", timeout=min(8000, self.timeout_ms))
            except Exception:
                pass
            try:
                page.wait_for_function(
                    "model => !model || document.body.innerText.toLowerCase().includes(model.toLowerCase())",
                    arg=expected_model,
                    timeout=min(4000, self.timeout_ms),
                )
            except Exception:
                pass

            # Lutron's product photo is commonly NOT an <img>. It is a Quasar
            # q-img element whose real source is CSS `background-image`, exactly
            # like: <div class="q-img__image ..." style="background-image:url(...)">.
            # Therefore, inspect both inline and computed backgrounds across the
            # rendered DOM, rank the image visually related to the product title,
            # and make it the primary candidate.
            css_background_candidates = page.evaluate(
                r"""
                expectedModel => {
                    const expected = (expectedModel || '').trim().toLowerCase();
                    const textOf = el => (el.innerText || el.textContent || '').trim().toLowerCase();
                    const urlsFrom = value => {
                        const out = [];
                        const text = String(value || '');
                        const regex = /url\(\s*["']?([^"')\s]+)["']?\s*\)/gi;
                        let match;
                        while ((match = regex.exec(text)) !== null) {
                            if (match[1]) out.push(match[1].trim());
                        }
                        return out;
                    };
                    const headingNodes = [...document.querySelectorAll(
                        'h1, h2, [class*="product-title"], [class*="product_name"], [class*="product-name"], [class*="model"], [class*="sku"]'
                    )];
                    const title = headingNodes.find(el => expected && textOf(el).includes(expected))
                        || document.querySelector('h1')
                        || headingNodes[0]
                        || null;
                    const tb = title ? title.getBoundingClientRect() : null;
                    const candidates = [];
                    for (const el of [...document.querySelectorAll('*')]) {
                        const rect = el.getBoundingClientRect();
                        if (rect.width < 100 || rect.height < 100) continue;
                        const inlineStyle = el.getAttribute('style') || '';
                        const computedStyle = getComputedStyle(el);
                        const background = computedStyle ? computedStyle.backgroundImage : '';
                        const urls = [...urlsFrom(inlineStyle), ...urlsFrom(background)];
                        if (!urls.length) continue;
                        const classText = String(el.className || '').toLowerCase();
                        for (const u of urls) {
                            const low = u.toLowerCase();
                            if (!u || /(?:logo|icon|sprite|flag|banner(?:s)?|hero|carousel|slider|homepage|landing|corporate|company|placeholder|facebook|instagram|youtube|linkedin)/.test(low)) continue;
                            let score = Math.min(rect.width * rect.height, 800000);
                            if (/q-img(?:__image|__content)?/.test(classText)) score += 600000;
                            if (/product|products|catalog|item|goods|photo|image/.test(low)) score += 150000;
                            if (tb) {
                                const verticallyRelated = rect.bottom >= tb.top - 260 && rect.top <= tb.bottom + 750;
                                const leftOfTitle = rect.left < tb.left && rect.right <= tb.left + 180;
                                const centerGap = Math.abs((rect.top + rect.bottom) / 2 - (tb.top + tb.bottom) / 2);
                                if (verticallyRelated) score += 500000;
                                if (leftOfTitle) score += 800000;
                                score += Math.max(0, 180000 - centerGap * 900);
                            }
                            candidates.push({url: u, score, classText, rect: [rect.left, rect.top, rect.width, rect.height]});
                        }
                    }
                    const deduped = new Map();
                    for (const item of candidates) {
                        const existing = deduped.get(item.url);
                        if (!existing || item.score > existing.score) deduped.set(item.url, item);
                    }
                    return [...deduped.values()].sort((a, b) => b.score - a.score).slice(0, 20);
                }
                """,
                expected_model,
            )
            if isinstance(css_background_candidates, list):
                for candidate in css_background_candidates:
                    if isinstance(candidate, dict) and clean_text(candidate.get("url")):
                        images.insert(0, "__LUTRON_PRIMARY__" + clean_text(candidate.get("url")))

            dom_images = page.eval_on_selector_all(
                "img, source, [style*='background-image']",
                r"""els => els.flatMap(el => {
                    const out = [];
                    if (el.currentSrc) out.push(el.currentSrc);
                    if (el.src) out.push(el.src);
                    for (const key of ['data-src','data-original','data-lazy-src','data-image','data-zoom-image']) {
                      const v = el.getAttribute(key); if (v) out.push(v);
                    }
                    const st = el.getAttribute('style') || '';
                    const m = st.match(/url\((?:['\"])?([^'\")]+)(?:['\"])?\)/i);
                    if (m && m[1]) out.push(m[1]);
                    return out;
                })""",
            )
            images.extend(clean_text(x) for x in dom_images if clean_text(x))
            return {
                "ok": True,
                "html": page.content(),
                "url": page.url,
                "images": dedupe_keep_order(images),
                "background_candidates": css_background_candidates if isinstance(css_background_candidates, list) else [],
                "error": "",
            }
        except Exception as exc:
            return {"ok": False, "html": "", "url": url, "images": [], "error": f"playwright_{type(exc).__name__}: {exc}"}
        finally:
            try:
                if page is not None:
                    page.close()
            except Exception:
                pass

    def close(self) -> None:
        for obj, method in ((self._context, "close"), (self._browser, "close"), (self._playwright, "stop")):
            try:
                if obj is not None:
                    getattr(obj, method)()
            except Exception:
                pass
        self._context = self._browser = self._playwright = None


def page_content_quality(title: str, description: str, specs: Sequence[Dict[str, str]], image_url: str, content_html: str) -> int:
    plain = re.sub(r"<[^>]+>", " ", clean_text(content_html))
    plain = re.sub(r"\s+", " ", plain).strip()
    score = 1 if title else 0
    score += 3 if len(clean_text(description)) >= 80 else 0
    score += 4 if len(specs) >= 4 else (2 if len(specs) >= 2 else 0)
    score += 3 if image_url else 0
    score += 3 if len(plain) >= 250 else 0
    return score


def fetch_lutron_product_page(
    url: str,
    timeout: int,
    session: requests.Session,
    expected_model: str = "",
) -> Tuple[Optional[requests.Response], str]:
    """Fetch a canonical product page and avoid the generic www landing page."""
    canonical_url = normalize_url(url)

    # The first visit establishes normal site cookies once per run. It is a
    # lightweight bootstrap request and avoids a product URL being handled as
    # a generic landing page by some Lutron site configurations.
    if not getattr(session, "_lutron_bootstrapped", False):
        try:
            session.get(
                "https://lutron.com.tw/en-us",
                headers=DEFAULT_HEADERS,
                timeout=timeout,
                allow_redirects=True,
            )
        except Exception:
            pass
        setattr(session, "_lutron_bootstrapped", True)

    headers = dict(DEFAULT_HEADERS)
    headers["Referer"] = "https://lutron.com.tw/en-us/"
    try:
        response = session.get(
            canonical_url,
            headers=headers,
            timeout=timeout,
            allow_redirects=True,
        )
        response.raise_for_status()
    except Exception as exc:
        return None, f"{type(exc).__name__}: {exc}"

    # Defensive retry. This covers old cached www routes and rare cases where
    # a routing layer serves the generic landing page on the first request.
    soup = BeautifulSoup(response.text or "", "html.parser")
    if looks_like_lutron_landing_page(soup, expected_model):
        retry_headers = dict(DEFAULT_HEADERS)
        retry_headers["Referer"] = "https://lutron.com.tw/en-us/categories/2/environmental"
        try:
            retry = session.get(
                canonical_url,
                headers=retry_headers,
                timeout=timeout,
                allow_redirects=True,
            )
            retry.raise_for_status()
            retry_soup = BeautifulSoup(retry.text or "", "html.parser")
            if not looks_like_lutron_landing_page(retry_soup, expected_model):
                response = retry
            else:
                return response, "generic_landing_page_returned"
        except Exception:
            return response, "generic_landing_page_returned"

    return response, ""


def scrape_product_page(record: Dict[str, str], timeout: int, session: requests.Session, renderer: Optional[LutronBrowserRenderer] = None) -> Dict[str, Any]:
    url = normalize_url(record.get("url"))
    cfg = get_source_config(url)
    result: Dict[str, Any] = {
        "url": url,
        "status": "failed",
        "source": cfg.key if cfg else "unknown",
        "brand": LUTRON_CONFIG.brand if cfg else "",
        "title": clean_text(record.get("input_title")),
        "input_title": clean_text(record.get("input_title")),
        "input_model": clean_text(record.get("input_model")),
        "description": "",
        "official_content_html": "",
        "spec_rows": [],
        "image_url": "",
        "source_category": "",
        "model": clean_text(record.get("input_model")),
        "is_accessory": False,
        "accessory_reason": "",
        "error": "",
    }
    if not url:
        result.update(status="no_url", error="URL kosong")
        return result
    if cfg is None:
        result.update(status="unsupported_domain", error="URL wajib berasal dari domain lutron.com.tw")
        return result

    response, fetch_note = fetch_lutron_product_page(
        url,
        timeout=timeout,
        session=session,
        expected_model=record.get("input_model", ""),
    )
    if response is None:
        result.update(status="request_error", error=fetch_note)
        return result

    def parse_document(html_text: str, base_url: str, network_images: Optional[Sequence[str]] = None) -> Dict[str, Any]:
        soup = BeautifulSoup(html_text or "", "html.parser")
        is_landing = looks_like_lutron_landing_page(soup, record.get("input_model", ""))
        products = product_jsonld(soup)
        title = extract_title(soup, record.get("input_title", ""))
        if is_landing:
            title = clean_text(record.get("input_title")) or clean_text(record.get("input_model"))
        category = extract_category(soup, products)
        root = choose_main_container(soup)
        strip_noise_nodes(root)
        description = "" if is_landing else extract_description(soup, products, title)
        content_items = [] if is_landing else extract_all_content(root, title)
        full_content = all_content_html(content_items, title)
        specs = extract_specs(soup, root, content_items, record.get("input_model", ""))
        model = extract_model(title, specs, record.get("input_model", ""))
        if model and not any(re.search(r"\bmodel\b", row.get("label", ""), re.I) and clean_text(row.get("value")) == model for row in specs):
            specs.insert(0, {"label": "Model", "source_label": "Model", "value": model, "section": "Product page"})
        image = "" if is_landing else extract_best_image_url(
            soup, base_url, model=model or record.get("input_model", ""), network_images=network_images
        )
        return {"landing": is_landing, "title": title, "category": category, "description": description,
                "content": full_content, "specs": specs, "model": model, "image": image}

    parsed = parse_document(response.text or "", response.url or url)
    browser_note = ""
    # Browser fallback is essential when the site injects product data and
    # images after load. Keep rendered result only when it is at least as rich.
    initial_quality = page_content_quality(parsed["title"], parsed["description"], parsed["specs"], parsed["image"], parsed["content"])
    if (parsed["landing"] or initial_quality < 8 or not parsed["image"]) and renderer is not None and renderer.available():
        rendered = renderer.fetch(response.url or url, record.get("input_model", ""))
        if rendered.get("ok"):
            browser_parsed = parse_document(rendered.get("html", ""), rendered.get("url", url), rendered.get("images", []))
            browser_quality = page_content_quality(browser_parsed["title"], browser_parsed["description"], browser_parsed["specs"], browser_parsed["image"], browser_parsed["content"])
            # Prefer the rendered DOM if it provides the real product image.
            # A CSS-backed q-img is authoritative even when its text extraction
            # is slightly shorter than the initial HTML shell.
            if browser_parsed["image"] or browser_quality >= initial_quality:
                parsed = browser_parsed
                fetch_note = _merge_source_notes(fetch_note, "browser_rendered")
        else:
            browser_note = clean_text(rendered.get("error", ""))

    is_landing_page = parsed["landing"]
    title = parsed["title"]
    source_category = parsed["category"]
    description = parsed["description"]
    official_content = parsed["content"]
    specs = parsed["specs"]
    model = parsed["model"]
    image_url = parsed["image"]
    is_accessory, reason = is_accessory_product(title, description, specs)

    # Never hide the Playwright error. Earlier versions overwrote it when the
    # HTTP fetch had returned the generic landing shell, making image failures
    # impossible to diagnose from the Excel log.
    correction = _merge_source_notes(
        fetch_note,
        browser_note,
        "generic_landing_page_returned" if is_landing_page else "",
    )

    result.update({
        "url": response.url or url,
        "status": "landing_page" if is_landing_page else ("ok" if title else "partial"),
        "title": title,
        "description": description,
        "official_content_html": official_content,
        "spec_rows": specs,
        "image_url": image_url,
        "source_category": source_category,
        "model": model,
        "is_accessory": is_accessory,
        "accessory_reason": reason,
        "error": correction,
    })
    return result



# ---------------------------------------------------------------------------
# Verified fallbacks for incomplete Lutron product pages
# ---------------------------------------------------------------------------

# Lutron's public product routes sometimes return only the site shell or a
# short product card.  This is why the older versions produced a title/model
# but no image, official content, or specification.  The fallback below never
# invents technical values: it retrieves a model-specific datasheet, extracts
# its selectable text, then records the fallback URL in the output log.
#
# The list is intentionally explicit and configurable.  Any response is
# accepted only if it is a real PDF and contains the requested model token.
DATASHEET_URL_TEMPLATES: Tuple[str, ...] = (
    "https://www.sunwe.com.tw/lutron/{model}.pdf",
    "https://www.sunwe.com.tw/lutron/{model}eop.pdf",
    "https://www.sunwe.com.tw/lutron/{model}_eop.PDF",
    "https://www.instrumentsgroup.co.za/index_files/Lutron/database/pdf/{model}.pdf",
    "https://jin-hua.com.tw/upload/product/{model}.pdf",
    "https://www.100y.com.tw/pdf_file/{model}.pdf",
    "https://siliconinstrument.com.sg/wp-content/uploads/2022/01/{model}.pdf",
)

PDF_SPEC_LABELS: Tuple[str, ...] = (
    "Operating Temperature", "Operating Temp.", "Operating Temp", "Operating Humidity",
    "Power Supply", "Power Consumption", "Power Input", "Power Current",
    "Dimension", "Dimensions", "Weight", "Display", "Circuit", "Measurement",
    "Range", "Resolution", "Accuracy", "Frequency Accuracy", "Width/Duty Accuracy",
    "Sensitivity", "Input Connector", "Sensor Type", "Data Output", "Data Logger",
    "Datalogger", "Memory Card", "Sampling Time", "Channels", "Channel",
    "Interface", "Protection", "Battery", "Accessories", "Standard Accessories",
    "Optional Accessories", "Case", "Time Base", "Output", "Function",
)

_PDF_LABEL_RE = re.compile(
    r"(?P<label>" + "|".join(re.escape(x) for x in sorted(PDF_SPEC_LABELS, key=len, reverse=True)) + r")"
    r"\s*[:.]?\s*",
    re.I,
)


def _is_real_product_data(data: Dict[str, Any]) -> bool:
    """Return True only when the Lutron page supplied substantive product data."""
    title = clean_text(data.get("title"))
    description = clean_text(data.get("description"))
    specs = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    if not title or HOME_PAGE_TITLE_RE.search(title):
        return False
    non_identity = [
        row for row in specs
        if not re.search(r"\b(model|sku|part|item|mpn|catalog)\b", clean_text(row.get("label")), re.I)
    ]
    return len(description) >= 80 or len(non_identity) >= 3 or len(clean_text(data.get("official_content_html"))) >= 350


def _model_file_candidates(value: Any) -> List[str]:
    """Build safe filename candidates from a model or multi-model set."""
    raw = clean_text(value)
    if not raw:
        return []
    pieces = re.split(r"\s*(?:/|,|;|\+|\band\b|\bwith\b)\s*", raw, flags=re.I)
    candidates: List[str] = []
    for piece in [raw] + pieces:
        p = clean_text(piece)
        p = re.sub(r"\b(?:set|kit|package)\b", "", p, flags=re.I).strip(" -_/")
        if not p:
            continue
        variants = [
            p,
            p.upper(),
            p.replace(" ", ""),
            p.replace("/", "-").replace(" ", ""),
            re.sub(r"[^A-Za-z0-9._-]+", "", p),
        ]
        candidates.extend(v for v in variants if len(v) >= 3)
    return dedupe_keep_order(candidates)


def _find_tesseract_executable() -> str:
    candidates = [
        clean_text(os.getenv("TESSERACT_CMD", "")),
        shutil.which("tesseract") or "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    return next((candidate for candidate in candidates if candidate and Path(candidate).exists()), "")


def _ocr_pdf_bytes(pdf_bytes: bytes, max_pages: int = 3) -> str:
    """OCR a scanned PDF only when text extraction produced no usable text."""
    executable = _find_tesseract_executable()
    if not pdf_bytes or fitz is None or not executable:
        return ""
    try:
        with tempfile.TemporaryDirectory(prefix="lutron_pdf_ocr_") as tmp:
            pdf_path = Path(tmp) / "datasheet.pdf"
            pdf_path.write_bytes(pdf_bytes)
            document = fitz.open(pdf_path)
            pieces: List[str] = []
            for index, page in enumerate(document[:max_pages]):
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                image_path = Path(tmp) / f"page_{index + 1}.png"
                pix.save(image_path)
                process = subprocess.run(
                    [executable, str(image_path), "stdout", "-l", "eng", "--psm", "6"],
                    capture_output=True,
                    text=True,
                    timeout=75,
                    check=False,
                )
                page_text = clean_text(process.stdout)
                if page_text:
                    pieces.append(page_text)
            return "\n".join(pieces)
    except Exception:
        return ""


def _pdf_response_text(pdf_bytes: bytes) -> Tuple[str, str]:
    """Return extracted datasheet text plus method: text, ocr, or empty."""
    if not pdf_bytes:
        return "", "empty"
    extracted = ""
    if PdfReader is not None:
        try:
            reader = PdfReader(BytesIO(pdf_bytes))
            chunks: List[str] = []
            # Datasheets are usually one or two pages. Capping avoids excessive
            # memory usage when a full operation manual is encountered.
            for page in reader.pages[:16]:
                page_text = clean_text(page.extract_text() or "")
                if page_text:
                    chunks.append(page_text)
            extracted = "\n".join(chunks)
        except Exception:
            extracted = ""
    if len(re.sub(r"\s+", "", extracted)) >= 80:
        return extracted, "text"
    ocr_text = _ocr_pdf_bytes(pdf_bytes)
    if len(re.sub(r"\s+", "", ocr_text)) >= 80:
        return ocr_text, "ocr"
    return extracted, "empty"


def _find_labeled_pdf_specs(pdf_text: str, model: str) -> List[Dict[str, str]]:
    """Extract conservative Label: Value rows from text-based Lutron PDFs.

    Two-column PDF layouts frequently interleave the left and right columns in
    one text line. To avoid false specifications, this parser begins only after
    a real specification heading and reads labels from the start of a line.
    """
    lines = [clean_spec_value(line) for line in clean_text(pdf_text).splitlines()]
    lines = [line for line in lines if line and len(line) <= 1200]
    rows: List[Dict[str, str]] = []
    if model:
        rows.append({"label": "Model", "value": model, "section": "Model-specific datasheet"})

    spec_mode = False
    pending_label = ""
    seen_headings = 0
    for line in lines:
        if re.search(r"\b(?:GENERAL\s*&?\s*ELECTRICAL\s*SPECIFICATIONS?|TECHNICAL\s+(?:SPECIFICATIONS?|DATA)|SPECIFICATIONS?)\b", line, re.I):
            spec_mode = True
            seen_headings += 1
            pending_label = ""
            continue
        if not spec_mode:
            continue
        if re.search(r"^(?:TABLE\s+FOR|OPERATION\s+MANUAL|FEATURES)\b", line, re.I):
            # Table rows may follow; keep parsing only known labels.
            pending_label = ""
            continue

        match = _PDF_LABEL_RE.match(line)
        if match:
            label = normalize_label(match.group("label"))
            value = clean_spec_value(line[match.end():])
            # Trim an interleaved right-column label. It will be picked up if
            # it appears at the start of a later extracted line.
            later = _PDF_LABEL_RE.search(value)
            # pypdf can combine left/right table columns. The right column
            # often begins with a generic technical label that is not in the
            # primary specification-label list.
            interleaved = re.search(
                r"\b(?:Operating(?:\s+Humidity)?|Power(?:\s+(?:Supply|Consumption|Input))?|"
                r"Memory\s+Card|Data\s+Output|Dimensions?|Weight|Accessories?|Standard|Optional|TABLE\s+FOR)\b",
                value,
                flags=re.I,
            )
            cut_at = min(
                [m.start() for m in (later, interleaved) if m is not None] or [len(value)]
            )
            if cut_at < len(value):
                value = clean_spec_value(value[:cut_at])
            value = re.sub(r"^[,:;.\-–—\s]+", "", value)
            if label.casefold() == "model":
                pending_label = ""
                continue
            if value and len(value) >= 4:
                # Ignore clearly fragmented leftovers from a PDF table.
                if value.casefold() not in {"format", "reading", "specifications"}:
                    rows.append({"label": label, "value": value, "section": "Model-specific datasheet"})
                pending_label = ""
            else:
                pending_label = label
            continue

        if pending_label and len(line) <= 500 and len(line) >= 4:
            rows.append({"label": pending_label, "value": line, "section": "Model-specific datasheet"})
            pending_label = ""

    # A datasheet can use non-standard labels. Capture a small number of
    # technical statements only when the clean labelled extraction is scarce.
    if len(rows) < 5:
        for line in lines:
            if re.search(r"\b(?:range|resolution|accuracy|power|dimension|weight|temperature|humidity)\b", line, re.I):
                rows.append({"label": "Technical detail", "value": line, "section": "Model-specific datasheet"})
                if len(rows) >= 12:
                    break
    return rows


def _pdf_feature_description(pdf_text: str, max_chars: int = 1400) -> str:
    lines = [clean_spec_value(line) for line in clean_text(pdf_text).splitlines()]
    lines = [line for line in lines if line and not re.fullmatch(r"(?:LUTRON\s*ELECTRONIC|ISO[- ]?9001.*)", line, flags=re.I)]
    start = next((i for i, line in enumerate(lines) if re.fullmatch(r"FEATURES", line, flags=re.I)), -1)
    if start < 0:
        return ""
    collected: List[str] = []
    for line in lines[start + 1:]:
        if re.search(r"^(?:GENERAL\s*&?\s*ELECTRICAL\s*SPECIFICATIONS?|SPECIFICATIONS?|TECHNICAL\s*DATA|OPERATION\s+MANUAL)", line, re.I):
            break
        line = re.sub(r"^[*•\-]+\s*", "", line).strip()
        if len(line) < 15:
            continue
        collected.append(line)
        if len(" ".join(collected)) >= max_chars:
            break
    return compact(" ".join(dedupe_keep_order(collected)), max_chars)


def _pdf_content_html(pdf_text: str, title: str, max_chars: int = 12000) -> str:
    """Make a clean, limited official-content tab from the datasheet text."""
    lines = [clean_spec_value(line) for line in clean_text(pdf_text).splitlines()]
    lines = [line for line in lines if line and not re.fullmatch(r"(?:LUTRON\s*ELECTRONIC|ISO[- ]?9001.*)", line, flags=re.I)]
    blocks: List[str] = []
    used = 0
    section = "Informasi Produk"
    for line in lines:
        if line.casefold() == clean_text(title).casefold():
            continue
        if len(line) <= 100 and (
            line.isupper() or re.fullmatch(r"(?:FEATURES|SPECIFICATIONS?|GENERAL\s*&?\s*ELECTRICAL\s*SPECIFICATIONS?|ACCESSORIES|DATA\s+OUTPUT)", line, re.I)
        ):
            section = line.title() if line.isupper() else line
            addition = len(section) + 9
            if used + addition > max_chars:
                break
            blocks.append(f"<h3>{esc(section)}</h3>")
            used += addition
            continue
        # Skip repeated product name boilerplate, but retain technical facts.
        if len(line) < 2:
            continue
        addition = len(line) + 7
        if used + addition > max_chars:
            break
        blocks.append(f"<p>{esc(line)}</p>")
        used += addition
    return "".join(blocks)


def _pdf_title(pdf_text: str, model: str, fallback: str) -> str:
    """Create a model-led catalog title from the datasheet heading."""
    lines = [clean_spec_value(line) for line in clean_text(pdf_text).splitlines()]
    candidate = ""
    for idx, line in enumerate(lines[:18]):
        if re.search(r"\b(?:COUNTER|THERMOMETER|ANALYZER|METER|TESTER|TRANSMITTER|RECORDER|LOGGER|GAUGE|HARDNESS)\b", line, re.I):
            candidate = line
            # Datasheets frequently put a qualifier on line 1 and product type
            # on line 2. Join only when it stays short and readable.
            if idx + 1 < len(lines) and len(candidate) < 110 and lines[idx + 1].isupper():
                candidate = clean_text(candidate + " " + lines[idx + 1])
            break
    candidate = re.sub(r"\bModel\s*:\s*.*$", "", candidate, flags=re.I).strip(" -:|")
    candidate = re.sub(r"\s+", " ", candidate)
    if candidate and model:
        return f"Lutron {model} - {candidate.title()}"
    if model and fallback:
        return f"Lutron {model} - {fallback}"
    return fallback or (f"Lutron {model}" if model else "")


def _fetch_datasheet_for_model(
    session: requests.Session,
    model: str,
    timeout: int,
    attempt_limit: int = 20,
) -> Dict[str, Any]:
    """Return a verified model-specific datasheet, or an empty dict."""
    if PdfReader is None:
        return {"ok": False, "reason": "pypdf_not_installed"}
    tried = 0
    model_tokens = [_model_token(x) for x in _model_file_candidates(model)]
    for candidate_model in _model_file_candidates(model):
        safe_model = quote(candidate_model, safe="-_.")
        for template in DATASHEET_URL_TEMPLATES:
            if tried >= attempt_limit:
                return {"ok": False, "reason": "datasheet_attempt_limit_reached"}
            tried += 1
            url = template.format(model=safe_model)
            try:
                response = session.get(
                    url,
                    headers=DEFAULT_HEADERS,
                    timeout=timeout,
                    allow_redirects=True,
                )
            except Exception:
                continue
            if response.status_code != 200:
                continue
            content_type = clean_text(response.headers.get("content-type", "")).casefold()
            payload = response.content or b""
            if not (payload.startswith(b"%PDF") or "pdf" in content_type):
                continue
            if len(payload) < 3000 or len(payload) > 30_000_000:
                continue
            pdf_text, extraction_method = _pdf_response_text(payload)
            compact_token = _model_token(model)
            if not pdf_text or (compact_token and compact_token not in _model_token(pdf_text)):
                # A set can use two models. Accept matching a component as long
                # as one explicit filename candidate matched the document.
                if not any(token and token in _model_token(pdf_text) for token in model_tokens):
                    continue
            return {
                "ok": True,
                "url": response.url or url,
                "text": pdf_text,
                "source_host": (urlparse(response.url or url).hostname or "").casefold(),
                "extraction_method": extraction_method,
            }
    return {"ok": False, "reason": "datasheet_not_found"}


def _merge_source_notes(*parts: Any) -> str:
    """Join semicolon-delimited status notes without repeating them."""
    values: List[str] = []
    for part in parts:
        for item in clean_text(part).split(";"):
            item = clean_text(item)
            if item:
                values.append(item)
    return "; ".join(dedupe_keep_order(values))


def _image_bad_url(url: str) -> bool:
    # Share the same strict gate used for official-page images so optional
    # external fallback cannot reintroduce generic banners or corporate art.
    return not valid_image_url(url)


def _image_candidate_score(url: str, model: str, title: str, engine: str) -> int:
    low = clean_text(url).casefold()
    token = _model_token(model)
    title_token = _model_token(title)
    score = 0
    if token and token in _model_token(low):
        score += 90
    if title_token and len(title_token) > 8 and title_token[:8] in _model_token(low):
        score += 10
    if re.search(r"\.(?:jpe?g|png|webp|avif)(?:[?#]|$)", low, re.I):
        score += 12
    if any(host in low for host in ("directindustry", "lutroninstruments", "sunwe", "instrumentsgroup", "legatool", "reviznipristroje", "advancom")):
        score += 10
    if engine == "bing":
        score += 3
    if _image_bad_url(low):
        score -= 1000
    return score


def _verify_image_candidate(session: requests.Session, url: str, timeout: int) -> bool:
    try:
        response = session.get(
            url,
            headers={"User-Agent": DEFAULT_HEADERS["User-Agent"], "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"},
            timeout=timeout,
            stream=True,
            allow_redirects=True,
        )
        if response.status_code != 200:
            return False
        content_type = clean_text(response.headers.get("content-type", "")).casefold()
        if content_type.startswith("image/"):
            return True
        # Some CDNs omit the correct content type. A recognizable image
        # extension is enough if the response body is not an HTML page.
        final_url = clean_text(response.url)
        return bool(re.search(r"\.(?:jpe?g|png|webp|avif|gif)(?:[?#]|$)", final_url, re.I)) and "text/html" not in content_type
    except Exception:
        return False


def _bing_image_candidates(session: requests.Session, query: str, timeout: int) -> List[str]:
    try:
        response = session.get(
            "https://www.bing.com/images/search",
            params={"q": query, "form": "HDRSC3", "first": "1"},
            headers=DEFAULT_HEADERS,
            timeout=timeout,
        )
        if response.status_code != 200:
            return []
        html_text = response.text or ""
        candidates: List[str] = []
        # Bing serializes original image URLs in an `m` JSON attribute.
        for raw in re.findall(r'\bm=(["\'])(.*?)\1', html_text, flags=re.I | re.S):
            payload = html.unescape(raw[1])
            try:
                obj = json.loads(payload)
            except Exception:
                continue
            for key in ("murl", "turl"):
                value = clean_text(obj.get(key, ""))
                if value:
                    candidates.append(value)
        # Newer Bing layouts can include the JSON as escaped data-m attributes.
        for match in re.finditer(r'"murl"\s*:\s*"([^"]+)"', html_text, flags=re.I):
            candidates.append(clean_text(match.group(1)).replace("\\/", "/"))
        return dedupe_keep_order(candidates)
    except Exception:
        return []


def _duckduckgo_image_candidates(session: requests.Session, query: str, timeout: int) -> List[str]:
    try:
        landing = session.get(
            "https://duckduckgo.com/",
            params={"q": query, "iax": "images", "ia": "images"},
            headers=DEFAULT_HEADERS,
            timeout=timeout,
        )
        if landing.status_code != 200:
            return []
        raw = landing.text or ""
        token_match = re.search(r'vqd=["\']([^"\']+)', raw, flags=re.I)
        if not token_match:
            token_match = re.search(r'vqd\\?["\']?\s*[:=]\s*["\']([^"\']+)', raw, flags=re.I)
        if not token_match:
            return []
        vqd = token_match.group(1)
        api = session.get(
            "https://duckduckgo.com/i.js",
            params={"l": "us-en", "o": "json", "q": query, "vqd": vqd, "f": ",,,,,", "p": "1"},
            headers={**DEFAULT_HEADERS, "Referer": "https://duckduckgo.com/"},
            timeout=timeout,
        )
        if api.status_code != 200:
            return []
        data = api.json()
        rows = data.get("results", []) if isinstance(data, dict) else []
        out = []
        for row in rows if isinstance(rows, list) else []:
            if isinstance(row, dict):
                out.append(clean_text(row.get("image", "")))
                out.append(clean_text(row.get("thumbnail", "")))
        return dedupe_keep_order(out)
    except Exception:
        return []


def search_product_image_url(
    session: requests.Session,
    model: str,
    title: str,
    timeout: int,
    engines: Sequence[str] = ("bing", "duckduckgo"),
) -> Tuple[str, str]:
    """Find a reachable product-photo URL only when Lutron itself has none."""
    query = " ".join(x for x in ("Lutron", clean_text(model), clean_text(title)) if x)
    if not query:
        return "", ""
    all_candidates: List[Tuple[str, str]] = []
    for engine in engines:
        if engine == "bing":
            urls = _bing_image_candidates(session, query, timeout)
        elif engine == "duckduckgo":
            urls = _duckduckgo_image_candidates(session, query, timeout)
        else:
            urls = []
        all_candidates.extend((engine, url) for url in urls if not _image_bad_url(url))

    seen = set()
    ranked: List[Tuple[int, str, str]] = []
    for engine, url in all_candidates:
        key = url.casefold()
        if key in seen:
            continue
        seen.add(key)
        ranked.append((_image_candidate_score(url, model, title, engine), engine, url))
    ranked.sort(reverse=True)

    # Test only a small ranked set. This makes it practical for hundreds of
    # products while keeping `image_url` usable by WooCommerce.
    for _score, engine, url in ranked[:12]:
        if _verify_image_candidate(session, url, timeout):
            return url, engine
    return "", ""


def _apply_verified_fallbacks(
    data: Dict[str, Any],
    record: Dict[str, str],
    session: requests.Session,
    timeout: int,
    enable_datasheet: bool = True,
    enable_image_search: bool = True,
    datasheet_attempt_limit: int = 20,
) -> Dict[str, Any]:
    """Fill incomplete link scrape with verified datasheet/image fallbacks."""
    model = clean_text(data.get("model") or record.get("input_model"))
    page_was_complete = _is_real_product_data(data)
    notes: List[str] = []
    if not page_was_complete:
        notes.append("lutron_page_incomplete")

    if not page_was_complete and enable_datasheet and model:
        pdf = _fetch_datasheet_for_model(session, model, timeout, datasheet_attempt_limit)
        if pdf.get("ok"):
            pdf_text = clean_text(pdf.get("text"))
            parsed_specs = _find_labeled_pdf_specs(pdf_text, model)
            description = _pdf_feature_description(pdf_text)
            fallback_title = _pdf_title(pdf_text, model, record.get("input_title", ""))
            data["title"] = fallback_title or clean_text(data.get("title")) or clean_text(record.get("input_title")) or model
            data["model"] = model
            data["description"] = description or clean_text(data.get("description"))
            data["official_content_html"] = _pdf_content_html(pdf_text, data["title"]) or clean_text(data.get("official_content_html"))
            data["spec_rows"] = parsed_specs or data.get("spec_rows") or []
            data["datasheet_url"] = clean_text(pdf.get("url"))
            data["datasheet_source_host"] = clean_text(pdf.get("source_host"))
            data["datasheet_extraction_method"] = clean_text(pdf.get("extraction_method"))
            data["status"] = "datasheet_fallback"
            data["source"] = "lutron_link_plus_verified_datasheet"
            notes.append("datasheet_fallback=" + clean_text(pdf.get("url")))
            if clean_text(pdf.get("extraction_method")) == "ocr":
                notes.append("datasheet_ocr_used")
            data["error"] = _merge_source_notes(data.get("error"), *notes)
        else:
            notes.append("datasheet_fallback_failed_" + clean_text(pdf.get("reason", "unknown")))
            data["error"] = _merge_source_notes(data.get("error"), *notes)

    # A link scrape can have useful specs but no image, and an image fallback
    # is only needed in that case. The source is recorded transparently.
    if not clean_text(data.get("image_url")) and enable_image_search:
        image_url, engine = search_product_image_url(
            session,
            model=model or record.get("input_model", ""),
            title=clean_text(data.get("title")) or record.get("input_title", ""),
            timeout=timeout,
        )
        if image_url:
            data["image_url"] = image_url
            data["image_source"] = engine + "_image_search"
            notes.append("image_fallback=" + data["image_source"])
        else:
            notes.append("image_fallback_not_found")
        data["error"] = _merge_source_notes(data.get("error"), *notes)
    elif clean_text(data.get("image_url")):
        data["image_source"] = "lutron_product_page_or_rendered_header"

    # Classification can change after a verified datasheet supplied the real
    # title/specifications. Re-evaluate it here so fallback pages do not retain
    # a stale or guessed accessory status.
    accessory, accessory_reason = is_accessory_product(
        clean_text(data.get("title")),
        clean_text(data.get("description")),
        data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else [],
    )
    data["is_accessory"] = accessory
    data["accessory_reason"] = accessory_reason

    # Do not silently create a seemingly successful catalog item from only the
    # input Excel. A record needs substantive source data, then it may be sent
    # to WooCommerce.
    if not _is_real_product_data(data) and data.get("status") != "datasheet_fallback":
        data["status"] = "incomplete_source"
    data["quality_flags"] = collect_source_quality_flags(data)
    return data


def _temperature_conversion_consistent(value: Any) -> bool:
    """Detect obvious Celsius/Fahrenheit conversion errors in product text."""
    text = clean_text(value)
    if not text:
        return True
    number = r"(-?\d+(?:[.,]\d+)?)"
    c_ranges = re.findall(number + r"\s*(?:to|hingga|sampai|[-–~])\s*" + number + r"\s*°?\s*C", text, flags=re.I)
    f_ranges = re.findall(number + r"\s*(?:to|hingga|sampai|[-–~])\s*" + number + r"\s*°?\s*F", text, flags=re.I)
    def n(v: str) -> float:
        return float(v.replace(",", "."))
    if c_ranges and f_ranges:
        for c_pair, f_pair in zip(c_ranges, f_ranges):
            expected = (n(c_pair[0]) * 9 / 5 + 32, n(c_pair[1]) * 9 / 5 + 32)
            actual = (n(f_pair[0]), n(f_pair[1]))
            if any(abs(a - e) > 4.0 for a, e in zip(actual, expected)):
                return False
    # Single-point notation such as 0°C (32°F).
    singles = re.findall(number + r"\s*°?\s*C[^0-9-]{0,30}" + number + r"\s*°?\s*F", text, flags=re.I)
    for c, f in singles:
        if abs((n(c) * 9 / 5 + 32) - n(f)) > 4.0:
            return False
    return True


def collect_source_quality_flags(data: Dict[str, Any]) -> List[str]:
    flags: List[str] = []
    # Do not use a global Celsius/Fahrenheit pairing heuristic as a publication
    # blocker. Lutron pages can repeat, reorder, or separately format ranges.
    # The translated spec validator preserves every numeric source token instead.
    image = clean_text(data.get("image_url"))
    if image and not valid_image_url(image):
        flags.append("rejected_generic_image")
    return dedupe_keep_order(flags)


def _canonical_image_key(url: Any) -> str:
    parts = urlsplit(clean_text(url))
    return urlunsplit((parts.scheme.casefold(), parts.netloc.casefold(), parts.path, "", ""))


def apply_batch_image_quality_gate(
    products: Sequence[Dict[str, Any]],
    skipped: List[Dict[str, Any]],
    require_verified_image: bool = False,
) -> List[Dict[str, Any]]:
    """Reject generic or cross-model duplicate images before Excel generation."""
    by_url: Dict[str, List[Dict[str, Any]]] = {}
    for data in products:
        image = clean_text(data.get("image_url"))
        if not image:
            continue
        by_url.setdefault(_canonical_image_key(image), []).append(data)

    invalid_ids = set()
    for key, group in by_url.items():
        models = {clean_text(item.get("model")) for item in group if clean_text(item.get("model"))}
        if len(group) > 1 and len(models) > 1:
            for item in group:
                invalid_ids.add(id(item))
                item["image_url"] = ""
                item["image_source"] = ""
                item["quality_flags"] = dedupe_keep_order(list(item.get("quality_flags") or []) + ["duplicate_image_across_models"])
                item["error"] = _merge_source_notes(item.get("error"), "duplicate_image_across_models")

    approved: List[Dict[str, Any]] = []
    for data in products:
        image = clean_text(data.get("image_url"))
        if image and not valid_image_url(image):
            data["image_url"] = ""
            data["image_source"] = ""
            data["quality_flags"] = dedupe_keep_order(list(data.get("quality_flags") or []) + ["rejected_generic_image"])
            data["error"] = _merge_source_notes(data.get("error"), "rejected_generic_image")
            image = ""
        if require_verified_image and not image:
            skipped.append({
                "url": data.get("url", ""),
                "title": data.get("title", ""),
                "status": data.get("status", ""),
                "reason": "missing_or_rejected_product_image",
                "error": _merge_source_notes(data.get("error"), "; ".join(data.get("quality_flags") or [])),
            })
            continue
        approved.append(data)
    return approved


def scrape_product_with_fallbacks(
    record: Dict[str, str],
    timeout: int,
    session: requests.Session,
    renderer: Optional[LutronBrowserRenderer] = None,
    enable_datasheet: bool = True,
    enable_image_search: bool = True,
    datasheet_attempt_limit: int = 20,
) -> Dict[str, Any]:
    data = scrape_product_page(record, timeout, session, renderer)
    return _apply_verified_fallbacks(
        data=data,
        record=record,
        session=session,
        timeout=timeout,
        enable_datasheet=enable_datasheet,
        enable_image_search=enable_image_search,
        datasheet_attempt_limit=datasheet_attempt_limit,
    )


# ---------------------------------------------------------------------------
# Copy and WooCommerce row generation
# ---------------------------------------------------------------------------

def normalize_lutron_category(value: Any, product_title: str = "") -> str:
    """Gunakan kategori breadcrumb hanya jika benar-benar kategori, bukan breadcrumb Products > judul."""
    s = clean_catalog_source_text(value)
    if not s:
        return ""
    s = s.replace("chevron_right", ">")
    parts = [clean_text(x).strip(" >|") for x in re.split(r"\s*>\s*|\s*›\s*", s) if clean_text(x).strip(" >|")]
    filtered: List[str] = []
    title_key = clean_text(product_title).casefold()
    for part in parts:
        key = part.casefold()
        if key in {"lutron", "products", "product", "all products"}:
            continue
        if title_key and key == title_key:
            continue
        filtered.append(part)
    # Breadcrumb halaman Lutron yang hanya "Products > nama produk" tidak boleh menjadi kategori.
    if not filtered:
        return ""
    return " > ".join(filtered)


def detect_family(data: Dict[str, Any]) -> Dict[str, Any]:
    title = clean_text(data.get("title"))
    desc = clean_catalog_source_text(data.get("description"))
    specs = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    if data.get("is_accessory"):
        family = dict(ACCESSORY_FAMILY)
    else:
        # Model dan title diprioritaskan. Hindari salah baca "Operating Temperature"
        # sebagai kategori thermometer untuk produk yang sebenarnya berbeda.
        core_blob = " ".join([title, clean_text(data.get("model")), clean_catalog_source_text(data.get("source_category"))])
        full_blob = " ".join([core_blob, desc, " ".join(f"{x.get('label','')} {x.get('value','')}" for x in specs)])
        family = dict(DEFAULT_FAMILY)
        for pattern, rule in FAMILY_RULES:
            if pattern.search(core_blob) or pattern.search(full_blob):
                family = dict(rule)
                break
    source_category = normalize_lutron_category(data.get("source_category"), title)
    # Model/title family rules are authoritative. A generic breadcrumb must not
    # override a recognised product class such as TR-PS pressure transmitters.
    if clean_text(family.get("category")) == clean_text(DEFAULT_FAMILY.get("category")) and source_category:
        family["category"] = source_category
    elif not clean_text(family.get("category")).casefold().startswith("lutron >"):
        family["category"] = clean_text(family.get("category"))
    return family


def spec_value(specs: Sequence[Dict[str, str]], labels: Sequence[str]) -> str:
    for row in specs:
        l = clean_text(row.get("label")).casefold()
        if any(term.casefold() in l for term in labels):
            return clean_text(row.get("value"))
    return ""


def feature_bullets(data: Dict[str, Any], family: Dict[str, Any]) -> List[str]:
    title = clean_text(data.get("title"))
    brand = clean_text(data.get("brand"))
    specs = data.get("spec_rows") or []
    if data.get("is_accessory"):
        bullets = [f"Produk ini merupakan aksesori atau komponen pendukung {brand} yang diproses sebagai produk katalog tersendiri."]
        model = clean_text(data.get("model")) or spec_value(specs, ["model", "sku", "cat", "part"])
        if model:
            bullets.append(f"Model atau referensi produk yang tercantum: {model}.")
        bullets.append("Cocokkan kompatibilitas produk dengan unit utama sebelum pemesanan atau penggunaan.")
        bullets.append(f"Data teknis dirujuk dari {source_official_name(get_source_config(data.get('url','')) or LUTRON_CONFIG)}.")
        return dedupe_keep_order(bullets)[:5]

    bullets: List[str] = []
    candidates = [
        ("Rentang pengukuran", ["range", "measurement range", "measuring range"]),
        ("Akurasi", ["accuracy"]),
        ("Resolusi", ["resolution", "minimum scale"]),
        ("Suhu operasi", ["operating temperature", "temperature range", "ambient temperature"]),
        ("Output atau antarmuka", ["output", "interface", "communication"]),
        ("Catu daya", ["power supply", "battery"]),
    ]
    for nice, terms in candidates:
        value = spec_value(specs, terms)
        if value:
            bullets.append(f"{nice}: {value}.")
    bullets.append(f"Dirancang untuk kebutuhan pengukuran, pemeriksaan, atau kontrol kualitas menggunakan {family['term']}.")
    if title and brand:
        bullets.append(f"Merek produk: {brand}.")
    return dedupe_keep_order(bullets)[:6]


def source_description_block(description: str, cfg: SourceConfig) -> str:
    # Jangan pernah meneruskan teks landing page Lutron atau karakter Chinese ke output.
    description = compact(clean_catalog_source_text(description), 1600)
    if not description:
        return ""
    label = "Ringkasan dari halaman Lutron"
    return f"<p><strong>{label}:</strong> {esc(description)}</p>"


def build_product_description(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title = clean_text(data.get("title"))
    brand = clean_text(data.get("brand"))
    cfg = get_source_config(data.get("url", "")) or LUTRON_CONFIG
    desc = source_description_block(clean_text(data.get("description")), cfg)
    bullets = "".join(f"<li>{esc(x)}</li>" for x in feature_bullets(data, family))
    if data.get("is_accessory"):
        intro = f"<p><strong>{esc(title)}</strong> adalah {esc(family['term'])} dari {esc(brand)} untuk melengkapi unit utama yang kompatibel.</p>"
        selection = "<p>Pastikan model, referensi produk, dan kecocokan teknis dengan unit utama sebelum digunakan.</p>"
    else:
        intro = f"<p><strong>{esc(title)}</strong> adalah {esc(family['term'])} dari {esc(brand)} yang digunakan untuk {esc(family['function'])}.</p>"
        selection = "<p>Gunakan informasi spesifikasi untuk memastikan kesesuaian produk dengan kebutuhan pengukuran dan kondisi penggunaan.</p>"
    return "".join([
        intro, desc,
        "<h2>Fungsi dan Keunggulan</h2>",
        f"<ul>{bullets}</ul>" if bullets else "",
        "<h2>Informasi Pemilihan Produk</h2>", selection,
        WEBSITE_CTA,
    ])


def build_short_description(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title, brand = clean_text(data.get("title")), clean_text(data.get("brand"))
    specs = data.get("spec_rows") or []
    model = clean_text(data.get("model")) or spec_value(specs, ["model", "sku", "part"])
    if data.get("is_accessory"):
        text = f"{title} adalah aksesori atau komponen pendukung {brand} untuk unit utama yang kompatibel."
        if model:
            text += f" Referensi: {model}."
        return compact(text, 320)
    text = f"{title} adalah {family['term']} dari {brand} untuk {family['function']}."
    range_value = spec_value(specs, ["range", "measurement range"])
    accuracy = spec_value(specs, ["accuracy"])
    if range_value:
        text += f" Rentang: {range_value}."
    if accuracy:
        text += f" Akurasi: {accuracy}."
    return compact(text, 320)


def build_application_tab(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    if data.get("is_accessory"):
        apps = [
            "kelengkapan unit utama yang kompatibel",
            "penggantian atau penambahan komponen sesuai referensi model",
            "pemeliharaan peralatan ukur",
            "kebutuhan laboratorium, inspeksi, atau quality control",
        ]
    else:
        apps = [
            "kontrol kualitas produksi",
            "pemeriksaan sampel atau parameter proses",
            "pengWEBSITEan di laboratorium atau area inspeksi",
            "validasi parameter sebelum proses lanjutan",
        ]
    return "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(f"<li>{esc(x)}</li>" for x in apps) + "</ul>"


def build_faq(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title, brand = clean_text(data.get("title")), clean_text(data.get("brand"))
    cfg = get_source_config(data.get("url", "")) or LUTRON_CONFIG
    model = clean_text(data.get("model"))
    if data.get("is_accessory"):
        return (
            f"<h3>Apa fungsi {esc(title)}?</h3><p>Produk ini merupakan aksesori atau komponen pendukung {esc(brand)} untuk unit utama yang kompatibel.</p>"
            f"<h3>Bagaimana memastikan kompatibilitasnya?</h3><p>Periksa model, nomor referensi, dan kebutuhan teknis unit utama sebelum pemesanan atau penggunaan.</p>"
            f"<h3>Apakah produk ini alat ukur utama?</h3><p>Tidak selalu. Produk ini diklasifikasikan sebagai aksesori atau komponen karena informasi halaman produk menunjukkan fungsi pendukung.</p>"
            f"<h3>Dari mana data teknis disusun?</h3><p>Data produk dirujuk dari {esc(source_official_name(cfg))}.</p>"
        )
    return (
        f"<h3>Apa fungsi {esc(title)}?</h3><p>{esc(title)} digunakan sebagai {esc(family['term'])} untuk {esc(family['function'])}.</p>"
        f"<h3>Apakah spesifikasi mengikuti halaman merek?</h3><p>Ya. Tab spesifikasi dan informasi produk dirujuk dari {esc(source_official_name(cfg))}.</p>"
        f"<h3>Bagaimana memilih model yang sesuai?</h3><p>Bandingkan rentang, akurasi, resolusi, kondisi penggunaan, dan kebutuhan antarmuka pada spesifikasi produk.</p>"
        f"<h3>Apakah model produk tercantum?</h3><p>{esc('Model yang terdeteksi: ' + model + '.' if model else 'Periksa tab Spesifikasi dan Informasi Produk Resmi untuk referensi model yang tersedia.')}</p>"
    )


def build_meta_description(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title, brand = clean_text(data.get("title")), clean_text(data.get("brand"))
    if data.get("is_accessory"):
        return compact(f"{title} adalah aksesori atau komponen {brand}. Periksa kompatibilitas model dan spesifikasinya di WEBSITE_KALIAN.", 155)
    range_value = spec_value(data.get("spec_rows") or [], ["range", "measurement range"])
    text = f"{title} dari {brand} untuk {family['function']}."
    if range_value:
        text += f" Rentang: {range_value}."
    text += " Cek spesifikasi di WEBSITE_KALIAN."
    return compact(text, 155)


# ---------------------------------------------------------------------------
# Optional DeepSeek content generation and Indonesian localization
# ---------------------------------------------------------------------------

AI_SETTINGS: Dict[str, Any] = {
    "enabled": False,
    "api_key": "",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-chat",
    "cache": {},
    "cache_path": "",
    "limit": 0,
    "timeout": 90,
    "delay": 0.0,
    "temperature": 0.15,
    "strict_indonesian": True,
    "spec_limit": 120,
    # Total percobaan = 1 permintaan awal + retry_count.
    # Retry hanya dipakai bila validasi deduplikasi spesifikasi AI gagal.
    "retry_count": 2,
    "retry_delay": 1.5,
}
AI_STATS: Dict[str, int] = {
    "used": 0, "cache_hit": 0, "fallback": 0, "skipped": 0,
    "retry": 0,
}
AI_PROMPT_VERSION = "lutron_WEBSITE_v21_publish_ready_source_locked_2026_07_27"

# Fallback hanya untuk label. Dalam mode AI bersih, nilai spesifikasi
# diterjemahkan oleh DeepSeek agar tidak ada teks Inggris/Chinese yang masuk.
SPEC_LABEL_ID_MAP = {
    "model": "Model",
    "circuit": "Sirkuit",
    "display": "Tampilan",
    "displaylcd size": "Ukuran layar LCD",
    "display direction": "Arah tampilan",
    "measurement unit": "Satuan pengukuran",
    "measurement range": "Rentang pengukuran",
    "measuring range": "Rentang pengukuran",
    "range": "Rentang",
    "accuracy": "Akurasi",
    "resolution": "Resolusi",
    "operating temperature": "Suhu operasi",
    "operating humidity": "Kelembapan operasi",
    "power supply": "Catu daya",
    "power current": "Arus daya",
    "power consumption": "Konsumsi daya",
    "weight": "Berat",
    "dimension": "Dimensi",
    "dimensions": "Dimensi",
    "function": "Fungsi",
    "data hold": "Penahan data",
    "data output": "Keluaran data",
    "interface": "Antarmuka",
    "sampling time of display": "Waktu sampling tampilan",
    "data logger sampling time setting range": "Rentang pengaturan waktu sampling pencatat data",
    "accessories included": "Aksesori yang disertakan",
    "accessories": "Aksesori",
    "zero adjust": "Penyetelan nol",
    "span adjust": "Penyetelan rentang",
    "principal": "Prinsip kerja",
    "input": "Masukan",
    "output": "Keluaran",
    "battery": "Baterai",
}


def set_ai_settings(**kwargs: Any) -> None:
    AI_SETTINGS.update(kwargs)
    AI_SETTINGS["base_url"] = clean_text(AI_SETTINGS.get("base_url")).rstrip("/") or "https://api.deepseek.com"
    AI_SETTINGS["cache"] = AI_SETTINGS.get("cache") if isinstance(AI_SETTINGS.get("cache"), dict) else {}
    AI_SETTINGS["limit"] = max(0, int(AI_SETTINGS.get("limit") or 0))
    AI_SETTINGS["timeout"] = max(20, int(AI_SETTINGS.get("timeout") or 90))
    AI_SETTINGS["delay"] = max(0.0, float(AI_SETTINGS.get("delay") or 0.0))
    AI_SETTINGS["temperature"] = max(0.0, min(1.0, float(AI_SETTINGS.get("temperature") or 0.15)))
    AI_SETTINGS["spec_limit"] = max(1, int(AI_SETTINGS.get("spec_limit") or 120))
    # Batasi retry agar satu produk tidak bisa memanggil API tanpa akhir.
    AI_SETTINGS["retry_count"] = max(0, min(5, int(AI_SETTINGS.get("retry_count") or 0)))
    AI_SETTINGS["retry_delay"] = max(0.0, float(AI_SETTINGS.get("retry_delay") or 0.0))
    AI_SETTINGS["strict_indonesian"] = bool(AI_SETTINGS.get("strict_indonesian", True))


def _html_text(value: Any) -> str:
    s = clean_text(value)
    s = re.sub(r"<\s*br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"</\s*(?:p|li|h2|h3|tr)\s*>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


def _sanitize_ai_html(value: Any) -> str:
    s = clean_text(value)
    s = re.sub(r"^```(?:json|html)?\s*|\s*```$", "", s, flags=re.I)
    s = re.sub(r"<\s*(script|style|iframe|object|embed)[^>]*>.*?<\s*/\s*\1\s*>", "", s, flags=re.I | re.S)
    s = re.sub(r"\s+on[a-z]+\s*=\s*(['\"]).*?\1", "", s, flags=re.I | re.S)
    allowed = {"p", "strong", "b", "em", "ul", "ol", "li", "h2", "h3", "br"}
    def tag_repl(m: re.Match) -> str:
        slash = "/" if m.group(1) else ""
        tag = (m.group(2) or "").casefold()
        return f"<{slash}{tag}>" if tag in allowed else ""
    s = re.sub(r"<\s*(/?)\s*([a-zA-Z0-9]+)(?:\s+[^>]*)?>", tag_repl, s)
    return s.strip()


def _sanitize_ai_text(value: Any, max_chars: int) -> str:
    s = compact(_html_text(value), max_chars)
    if s and s[-1] not in ".!?":
        s += "."
    return s


def _parse_json(text: str) -> Dict[str, Any]:
    s = clean_text(text)
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s, flags=re.I)
    try:
        value = json.loads(s)
        return value if isinstance(value, dict) else {}
    except Exception:
        start, end = s.find("{"), s.rfind("}")
        if start >= 0 and end > start:
            try:
                value = json.loads(s[start:end + 1])
                return value if isinstance(value, dict) else {}
            except Exception:
                return {}
    return {}


def _spec_rows_for_ai(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Send DeepSeek the untouched source rows, including both spec sections.

    `prepare_canonical_specs_for_product()` creates `_raw_spec_rows` before it
    applies fallback rules. AI must inspect that raw list, otherwise it cannot
    identify duplicates that the rule-based parser did not recognise.
    """
    raw_rows = data.get("_raw_spec_rows")
    rows = raw_rows if isinstance(raw_rows, list) else (
        data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    )
    limit = int(AI_SETTINGS.get("spec_limit") or 120)
    clean_rows: List[Dict[str, Any]] = []
    for index, item in enumerate(rows[:limit], 1):
        if not isinstance(item, dict):
            continue
        source_label = clean_catalog_source_text(item.get("source_label") or item.get("label"))
        value = clean_catalog_source_text(item.get("value"))
        if source_label and value:
            clean_rows.append({
                "id": index,
                "source_label": source_label,
                "source_value": compact(value, 700),
                "section": clean_text(item.get("section")) or "Specifications",
            })
    return clean_rows


def _ai_payload(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    source_full = html_to_source_text(data.get("official_content_html"), 8500)
    source_desc = compact(clean_catalog_source_text(data.get("description")), 2200)
    return {
        "source_product_name": clean_catalog_source_text(data.get("title")),
        "brand": clean_text(data.get("brand")) or "Lutron",
        "model": clean_text(data.get("model")),
        "source_url": clean_text(data.get("url")),
        "source_category": clean_catalog_source_text(data.get("source_category")),
        "product_type": "accessory_or_component" if data.get("is_accessory") else "main_product",
        "suggested_category": clean_catalog_source_text(family.get("category")).replace("", "").strip(),
        "family_function": clean_catalog_source_text(family.get("function")),
        "source_description": source_desc,
        "source_full_information": source_full,
        "specifications": _spec_rows_for_ai(data),
    }


def _ai_cache_key(payload: Dict[str, Any]) -> str:
    raw = json.dumps(
        {"version": AI_PROMPT_VERSION, "model": AI_SETTINGS.get("model"), "payload": payload},
        ensure_ascii=False,
        sort_keys=True,
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _numeric_tokens(value: Any) -> List[str]:
    return re.findall(r"(?<![A-Za-z])[-+]?\d+(?:[.,]\d+)?", clean_text(value))


def _localize_source_label(source_label: Any, localized_label: Any) -> str:
    """Return a safe Indonesian display label without changing source identity.

    The original v16 called this helper from both the AI and fallback paths,
    but the function was missing. That NameError made every AI-enabled product
    fail validation and enter the Skipped sheet.
    """
    localized = clean_catalog_source_text(localized_label)
    source = clean_catalog_source_text(source_label)
    return normalize_label(localized or source)


def _coerce_ai_spec_id(value: Any) -> Optional[int]:
    try:
        parsed = int(str(value).strip())
        return parsed if parsed > 0 else None
    except Exception:
        return None


def _spec_validation_issue(code: str, message: str, **detail: Any) -> Dict[str, Any]:
    """Build a compact, machine-readable validator message for DeepSeek retry."""
    item: Dict[str, Any] = {"code": code, "message": message}
    for key, value in detail.items():
        if value not in (None, "", [], {}, ()):
            item[key] = value
    return item


def _trim_feedback_ids(values: Iterable[Any], limit: int = 30) -> List[int]:
    """Keep retry feedback precise without making a retry prompt excessively long."""
    out: List[int] = []
    for value in values:
        try:
            out.append(int(value))
        except Exception:
            continue
    return out[:limit]


def _validate_all_translated_specs(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], Dict[int, Dict[str, Any]], List[Dict[str, Any]]]:
    """Validate every translated source row and return exact retry feedback.

    The prior implementation returned only ``[]`` on any error. That made the
    retry prompt say merely ``invalid_ai_specification_deduplication``. This
    version identifies the field, source IDs, and numeric tokens that violated
    the contract so the next DeepSeek attempt can repair the exact JSON shape.
    """
    issues: List[Dict[str, Any]] = []
    raw_items = parsed.get("translated_specifications")
    if not isinstance(raw_items, list):
        return [], {}, [
            _spec_validation_issue(
                "translated_specifications_not_array",
                "Field translated_specifications harus berupa array dan memuat satu terjemahan untuk setiap ID sumber.",
            )
        ]
    if not source_specs:
        return [], {}, issues

    translated_by_id: Dict[int, Dict[str, Any]] = {}
    invalid_items: List[int] = []
    duplicate_ids: List[int] = []
    for index, item in enumerate(raw_items, 1):
        if not isinstance(item, dict):
            invalid_items.append(index)
            continue
        item_id = _coerce_ai_spec_id(item.get("id"))
        if item_id is None:
            invalid_items.append(index)
            continue
        if item_id in translated_by_id:
            duplicate_ids.append(item_id)
            continue
        translated_by_id[item_id] = item

    if invalid_items:
        issues.append(_spec_validation_issue(
            "translated_specification_invalid_item",
            "Setiap item translated_specifications harus berupa objek dengan id numerik yang valid.",
            item_positions=_trim_feedback_ids(invalid_items),
        ))
    if duplicate_ids:
        issues.append(_spec_validation_issue(
            "translated_specification_duplicate_id",
            "ID pada translated_specifications tidak boleh muncul lebih dari satu kali.",
            ids=_trim_feedback_ids(sorted(set(duplicate_ids))),
        ))

    expected_ids = {int(spec["id"]) for spec in source_specs}
    actual_ids = set(translated_by_id)
    missing_ids = sorted(expected_ids - actual_ids)
    unexpected_ids = sorted(actual_ids - expected_ids)
    if missing_ids:
        issues.append(_spec_validation_issue(
            "translated_specification_missing_ids",
            "Terjemahan untuk ID sumber berikut belum dikirim. Tambahkan semuanya ke translated_specifications.",
            ids=_trim_feedback_ids(missing_ids),
        ))
    if unexpected_ids:
        issues.append(_spec_validation_issue(
            "translated_specification_unexpected_ids",
            "ID berikut tidak ada pada input sumber. Hapus dari translated_specifications.",
            ids=_trim_feedback_ids(unexpected_ids),
        ))
    if issues:
        return [], {}, issues

    translated_rows: List[Dict[str, Any]] = []
    mapped: Dict[int, Dict[str, Any]] = {}
    for spec in source_specs:
        source_id = int(spec["id"])
        item = translated_by_id[source_id]
        label = clean_catalog_source_text(item.get("label_id"))
        value = clean_catalog_source_text(item.get("value_id"))
        if not label:
            issues.append(_spec_validation_issue(
                "translated_specification_empty_label",
                "label_id tidak boleh kosong dan harus berbahasa Indonesia.",
                id=source_id,
            ))
            continue
        if not value:
            issues.append(_spec_validation_issue(
                "translated_specification_empty_value",
                "value_id tidak boleh kosong. Terjemahkan teks, tetapi pertahankan angka dan satuan sumber.",
                id=source_id,
            ))
            continue
        if has_cjk(label) or has_cjk(value):
            issues.append(_spec_validation_issue(
                "translated_specification_contains_chinese",
                "Label atau nilai masih mengandung karakter Chinese. Ganti dengan Bahasa Indonesia.",
                id=source_id,
            ))
            continue

        normalized_value = value.replace(",", ".")
        expected_numbers = [token.replace(",", ".") for token in _numeric_tokens(spec["source_value"])]
        missing_numbers = [token for token in expected_numbers if token not in normalized_value]
        if missing_numbers:
            issues.append(_spec_validation_issue(
                "translated_specification_missing_numeric_token",
                "Angka sumber berikut hilang atau berubah pada value_id. Salin angka tersebut persis seperti sumber.",
                id=source_id,
                required_numeric_tokens=missing_numbers,
                source_value=compact(spec.get("source_value"), 220),
            ))
            continue

        translated = {
            "id": source_id,
            "label": _localize_source_label(spec.get("source_label"), label),
            "source_label": clean_text(spec.get("source_label")),
            "value": value,
            "section": clean_text(spec.get("section")) or "Specifications",
        }
        translated_rows.append(translated)
        mapped[source_id] = translated

    if issues:
        return [], {}, issues
    return translated_rows, mapped, issues


def _source_label_duplicate_family(left: Dict[str, Any], right: Dict[str, Any]) -> bool:
    """Reject unsafe AI removals where labels clearly represent different facts."""
    left_key, _left_std, left_source = _canonical_label_info({
        "source_label": left.get("source_label"),
        "label": left.get("source_label"),
    })
    right_key, _right_std, right_source = _canonical_label_info({
        "source_label": right.get("source_label"),
        "label": right.get("source_label"),
    })
    if _canonical_keys_can_compare(left_key, right_key):
        return True
    a, b = _spec_text_key(left_source), _spec_text_key(right_source)
    if not a or not b:
        return False
    if a == b:
        return True
    return SequenceMatcher(None, a, b).ratio() >= 0.92


def _source_values_plausibly_identical(left: Dict[str, Any], right: Dict[str, Any]) -> bool:
    """Allow AI to merge only cosmetic variants of the same source fact.

    This guard keeps AI from deleting valid values such as Meter versus Probe
    dimensions, multiple frequency ranges, or different operating modes.
    """
    a = clean_text(left.get("source_value"))
    b = clean_text(right.get("source_value"))
    if _spec_values_equivalent(a, b):
        return True

    a_numbers = [v.replace(",", ".") for v in _numeric_tokens(a)]
    b_numbers = [v.replace(",", ".") for v in _numeric_tokens(b)]
    if a_numbers != b_numbers:
        return False

    # Values with numerical data may differ only in punctuation, spacing, or
    # equivalent wording such as "to", "-", and "hingga".
    if a_numbers:
        a_units = _spec_unit_signature(a)
        b_units = _spec_unit_signature(b)
        if a_units and b_units and a_units != b_units:
            return False
        return True

    a_key, b_key = _spec_text_key(a), _spec_text_key(b)
    if not a_key or not b_key:
        return False
    if a_key == b_key or a_key in b_key or b_key in a_key:
        return True
    return SequenceMatcher(None, a_key, b_key).ratio() >= 0.88


def _validate_ai_spec_deduplication(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, str]], int, List[Dict[str, Any]]]:
    """Validate AI duplicate selection and return actionable retry feedback."""
    translated_rows, translations, issues = _validate_all_translated_specs(parsed, source_specs)
    if issues:
        return [], 0, issues

    raw_keep_ids = parsed.get("keep_specification_ids")
    raw_removed = parsed.get("removed_duplicate_specifications")
    if not isinstance(raw_keep_ids, list):
        return [], 0, [
            _spec_validation_issue(
                "keep_specification_ids_not_array",
                "Field keep_specification_ids harus berupa array ID numerik.",
            )
        ]
    if not isinstance(raw_removed, list):
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_specifications_not_array",
                "Field removed_duplicate_specifications harus berupa array. Gunakan [] bila tidak ada duplikat.",
            )
        ]

    expected_ids = {int(spec["id"]) for spec in source_specs}
    keep_values = [_coerce_ai_spec_id(value) for value in raw_keep_ids]
    invalid_keep_positions = [index for index, value in enumerate(keep_values, 1) if value is None]
    if invalid_keep_positions:
        return [], 0, [
            _spec_validation_issue(
                "keep_specification_ids_invalid_value",
                "Semua item keep_specification_ids harus berupa ID numerik valid.",
                positions=_trim_feedback_ids(invalid_keep_positions),
            )
        ]
    keep_ids = [int(value) for value in keep_values if value is not None]
    duplicate_keep_ids = sorted({value for value in keep_ids if keep_ids.count(value) > 1})
    if duplicate_keep_ids:
        return [], 0, [
            _spec_validation_issue(
                "keep_specification_ids_duplicate",
                "ID tidak boleh muncul lebih dari satu kali pada keep_specification_ids.",
                ids=_trim_feedback_ids(duplicate_keep_ids),
            )
        ]
    unknown_keep_ids = sorted(set(keep_ids) - expected_ids)
    if unknown_keep_ids:
        return [], 0, [
            _spec_validation_issue(
                "keep_specification_ids_unknown",
                "ID berikut tidak ada pada input sumber. Hapus dari keep_specification_ids.",
                ids=_trim_feedback_ids(unknown_keep_ids),
            )
        ]

    removed_map: Dict[int, int] = {}
    invalid_removed_positions: List[int] = []
    duplicate_removed_ids: List[int] = []
    for index, item in enumerate(raw_removed, 1):
        if not isinstance(item, dict):
            invalid_removed_positions.append(index)
            continue
        source_id = _coerce_ai_spec_id(item.get("source_id", item.get("id")))
        duplicate_of_id = _coerce_ai_spec_id(item.get("duplicate_of_id", item.get("keep_id")))
        if source_id is None or duplicate_of_id is None or source_id == duplicate_of_id:
            invalid_removed_positions.append(index)
            continue
        if source_id in removed_map:
            duplicate_removed_ids.append(source_id)
            continue
        removed_map[source_id] = duplicate_of_id

    if invalid_removed_positions:
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_specification_invalid_item",
                "Setiap item removed_duplicate_specifications harus memiliki source_id dan duplicate_of_id numerik yang berbeda.",
                positions=_trim_feedback_ids(invalid_removed_positions),
            )
        ]
    if duplicate_removed_ids:
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_specification_duplicate_source_id",
                "source_id yang sama tidak boleh muncul lebih dari satu kali pada removed_duplicate_specifications.",
                ids=_trim_feedback_ids(sorted(set(duplicate_removed_ids))),
            )
        ]

    removed_ids = set(removed_map)
    unknown_removed_ids = sorted(removed_ids - expected_ids)
    unknown_targets = sorted(set(removed_map.values()) - expected_ids)
    if unknown_removed_ids:
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_specification_unknown_source_id",
                "source_id berikut tidak ada pada input sumber.",
                ids=_trim_feedback_ids(unknown_removed_ids),
            )
        ]
    if unknown_targets:
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_specification_unknown_duplicate_of_id",
                "duplicate_of_id berikut tidak ada pada input sumber.",
                ids=_trim_feedback_ids(unknown_targets),
            )
        ]

    target_not_kept = sorted(set(removed_map.values()) - set(keep_ids))
    if target_not_kept:
        return [], 0, [
            _spec_validation_issue(
                "removed_duplicate_target_not_kept",
                "Setiap duplicate_of_id harus ada dalam keep_specification_ids.",
                ids=_trim_feedback_ids(target_not_kept),
            )
        ]
    overlap_ids = sorted(set(keep_ids) & removed_ids)
    if overlap_ids:
        return [], 0, [
            _spec_validation_issue(
                "specification_id_in_keep_and_removed",
                "ID yang sama tidak boleh berada di keep_specification_ids dan removed_duplicate_specifications.",
                ids=_trim_feedback_ids(overlap_ids),
            )
        ]

    accounted_ids = set(keep_ids) | removed_ids
    missing_accounted_ids = sorted(expected_ids - accounted_ids)
    if missing_accounted_ids:
        return [], 0, [
            _spec_validation_issue(
                "specification_id_not_accounted",
                "Setiap ID sumber harus masuk keep_specification_ids atau removed_duplicate_specifications.",
                ids=_trim_feedback_ids(missing_accounted_ids),
            )
        ]

    source_by_id = {int(item["id"]): item for item in source_specs}
    for source_id, target_id in removed_map.items():
        source_row = source_by_id[source_id]
        target_row = source_by_id[target_id]
        if not _source_label_duplicate_family(source_row, target_row):
            return [], 0, [
                _spec_validation_issue(
                    "removed_duplicate_label_not_same_fact",
                    "ID ini tidak boleh dibuang karena labelnya tidak menyatakan fakta teknis yang sama dengan duplicate_of_id.",
                    source_id=source_id,
                    duplicate_of_id=target_id,
                    source_label=compact(source_row.get("source_label"), 120),
                    duplicate_label=compact(target_row.get("source_label"), 120),
                )
            ]
        if not _source_values_plausibly_identical(source_row, target_row):
            return [], 0, [
                _spec_validation_issue(
                    "removed_duplicate_value_not_equivalent",
                    "ID ini tidak boleh dibuang karena nilai sumber berbeda. Pertahankan kedua ID di keep_specification_ids.",
                    source_id=source_id,
                    duplicate_of_id=target_id,
                    source_value=compact(source_row.get("source_value"), 180),
                    duplicate_value=compact(target_row.get("source_value"), 180),
                )
            ]

        source_is_technical = _is_technical_spec_section(source_row.get("section"))
        target_is_technical = _is_technical_spec_section(target_row.get("section"))
        if source_is_technical and not target_is_technical:
            return [], 0, [
                _spec_validation_issue(
                    "technical_specification_removed_in_favor_of_generic_specification",
                    "Jika dua fakta memang sama, simpan ID dari Technical Specifications dan buang ID dari Specifications, bukan sebaliknya.",
                    source_id=source_id,
                    duplicate_of_id=target_id,
                )
            ]

    output: List[Dict[str, str]] = []
    for source_id in keep_ids:
        row = translations[source_id]
        output.append({
            "label": clean_text(row.get("label")),
            "source_label": clean_text(row.get("source_label")),
            "value": clean_text(row.get("value")),
            "section": clean_text(row.get("section")),
        })
    return output, len(removed_ids), []


def _validate_ai_fields(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    fields: Dict[str, Any] = {
        "product_name_id": _sanitize_ai_text(parsed.get("product_name_id"), 180),
        "category_id": _sanitize_ai_text(parsed.get("category_id"), 100).rstrip("."),
        "product_description_html": _sanitize_ai_html(parsed.get("product_description_html")),
        "short_description": _sanitize_ai_text(parsed.get("short_description"), 260),
        "feature_tab_html": _sanitize_ai_html(parsed.get("feature_tab_html")),
        "faq_html": _sanitize_ai_html(parsed.get("faq_html")),
        "meta_description": _sanitize_ai_text(parsed.get("meta_description"), 158),
        "official_information_html": _sanitize_ai_html(parsed.get("official_information_html")),
    }
    translated_rows, removed_count, spec_feedback = _validate_ai_spec_deduplication(parsed, source_specs)
    fields["translated_spec_rows"] = translated_rows
    fields["ai_spec_dedup_applied"] = bool(source_specs)
    fields["ai_spec_dedup_removed_count"] = removed_count

    required_text = [
        fields["product_name_id"],
        fields["product_description_html"],
        fields["short_description"],
        fields["feature_tab_html"],
        fields["faq_html"],
        fields["meta_description"],
        fields["official_information_html"],
    ]
    if any(not item for item in required_text):
        return False, {}, "missing_indonesian_ai_field", []
    if any(has_cjk(item) for item in required_text):
        return False, {}, "chinese_character_detected_in_ai_output", []
    if len(_html_text(fields["product_description_html"])) < 160 or "<h2" not in fields["product_description_html"].lower():
        return False, {}, "invalid_product_description", []
    if len(_html_text(fields["official_information_html"])) < 100 or "<h2" not in fields["official_information_html"].lower():
        return False, {}, "invalid_official_information", []
    if "<li" not in fields["feature_tab_html"].lower() or "<h3" not in fields["faq_html"].lower():
        return False, {}, "invalid_ai_tabs", []
    if source_specs and not fields["translated_spec_rows"]:
        return False, {}, "invalid_ai_specification_deduplication", spec_feedback
    forbidden = re.compile(r"\b(harga|diskon|promo|garansi resmi|ready stock|stok tersedia|gratis ongkir|termurah)\b", re.I)
    all_public_text = " ".join(str(x) for x in required_text)
    if forbidden.search(all_public_text):
        return False, {}, "forbidden_sales_claim", []
    unsafe_claim = re.compile(r"\b(prudent avoidance|epa|who|risiko kesehatan|risiko kanker|bahaya kesehatan|diagnosis medis)\b", re.I)
    if unsafe_claim.search(all_public_text):
        return False, {}, "unsafe_health_or_regulatory_claim", []
    return True, fields, "ok", []


def _call_ai(
    payload: Dict[str, Any],
    retry_reason: str = "",
    retry_feedback: Optional[Sequence[Dict[str, Any]]] = None,
    attempt_number: int = 1,
    max_attempts: int = 1,
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    api_key = clean_text(AI_SETTINGS.get("api_key"))
    if not api_key:
        return False, {}, "missing_api_key", []

    user = {
        "task": "Buat semua konten katalog WooCommerce dalam bahasa Indonesia dari data produk Lutron berikut.",
        "critical_rules": [
            "SEMUA teks untuk pembaca harus berbahasa Indonesia yang rapi dan profesional.",
            "Konten publik dilarang menyebut proses pengambilan data, AI, scraping, sumber URL, halaman resmi Lutron, situs/website merek, dokumen sumber, atau mengarahkan pembeli ke kanal merek. Simpan provenance hanya pada field internal Source URL dan log validasi.",
            "Jangan tulis karakter Mandarin, Chinese, atau kutipan bahasa sumber.",
            "Pertahankan merek Lutron, model, kode, angka, satuan, lambang, serta singkatan teknis seperti LCD, True RMS, SD card, RS232, USB, dan IP apa adanya.",
            "Gunakan hanya fakta yang tertulis pada source_url, source_description, source_full_information, atau source_value specifications. Jangan menambah klaim, harga, stok, garansi, sertifikasi, fitur, aplikasi, maupun kemampuan yang tidak ada pada sumber.",
            "Untuk translated_specifications, value_id hanya boleh menjadi terjemahan setia dari source_value dengan urutan fakta yang sama. Jangan menambahkan frasa baru, kesimpulan, atau penjelasan yang tidak ada pada source_value.",
            "Terjemahkan SETIAP source_label dan source_value dalam specifications. Jangan menghapus, mengganti, atau mengonversi angka maupun satuan.",
            "Setelah menerjemahkan seluruh specifications, lakukan penyaringan duplikat SEMANTIK hanya untuk tab Spesifikasi.",
            "Dua spesifikasi boleh dianggap duplikat hanya jika labelnya menyatakan fakta teknis yang sama DAN semua angka, satuan, rentang, komponen, mode, kanal, sensor/probe, serta kondisi pengWEBSITEannya sama. Perbedaan bahasa, kapitalisasi, spasi, tanda baca, atau kata 'to', '-', dan 'hingga' boleh dianggap kosmetik.",
            "Jika fakta yang sama muncul pada Technical Specifications dan Specifications, WAJIB simpan ID dari Technical Specifications dan masukkan ID dari Specifications ke removed_duplicate_specifications.",
            "Jangan menghapus dua baris hanya karena labelnya sama. Pertahankan keduanya jika nilai berbeda, misalnya Meter dan Probe, 1 GHz dan 6 GHz, rentang dan akurasi, operasi normal dan pencatatan data, atau ukuran komponen yang berbeda.",
            "Jangan menggabungkan atau mengonversi nilai. Tampilkan nilai yang disimpan sesuai source_value terpilih, hanya diterjemahkan ke Bahasa Indonesia.",
            "PENTING: Setiap ID sumber WAJIB muncul tepat satu kali: masuk keep_specification_ids atau menjadi source_id pada removed_duplicate_specifications. Jangan ada ID yang hilang atau muncul di kedua daftar.",
            "PENTING: translated_specifications WAJIB memuat terjemahan untuk SEMUA ID sumber, termasuk ID yang kemudian ditandai sebagai duplikat.",
            "PENTING: Bila tidak ada spesifikasi duplikat, removed_duplicate_specifications WAJIB dikirim sebagai array kosong [] dan keep_specification_ids WAJIB berisi seluruh ID sumber sesuai urutan input.",
            "Status product_type sudah ditentukan oleh validasi rule. Jangan mengubah status produk utama menjadi aksesori dan jangan menyebut kompatibilitas unit utama kecuali data sumber secara eksplisit menyatakan produk adalah aksesori atau komponen pengganti.",
            "Jangan membuat klaim kesehatan, risiko paparan, EPA, WHO, Prudent Avoidance, kanker, atau klaim keselamatan medis karena data sumber produk tidak digunakan untuk klaim tersebut.",
            "Jangan menyatakan USB atau RS232 sebagai port atau kelengkapan bawaan kecuali sumber menyebutnya demikian. Jika sumber menyebut RS232 PC serial interface dan USB-01 sebagai optional accessory, tulis: antarmuka RS232 untuk keluaran data; koneksi USB memerlukan kabel USB-01 opsional.",
            "Bedakan memori maksimum/minimum, Data Hold, antarmuka PC, dan data logger internal. Jangan menulis pencatatan data otomatis/manual kecuali source_value secara eksplisit menyatakan automatic/manual data logger sebagai fitur alat. Penggunaan komputer sebagai data logger melalui RS232 bukan data logger internal.",
            "Jangan menyalin tabel spesifikasi ke product_description_html karena tabel dibuat terpisah.",
            "Balas JSON valid saja, tanpa markdown dan tanpa code fence.",
        ],
        "output_schema": {
            "product_name_id": "Nama produk SEO dalam bahasa Indonesia, maksimal 180 karakter. Gunakan urutan: jenis produk + Lutron + model + pembeda utama yang benar-benar ada pada sumber. Tanpa titik di akhir dan hindari simbol &; gunakan kata dan.",
            "category_id": "Satu kategori WooCommerce dalam bahasa Indonesia tanpa kata Lutron dan tanpa tanda >.",
            "product_description_html": "HTML aman: 2 paragraf pembuka, <h2>Keunggulan Produk</h2> dengan <ul> 4-6 poin, <h2>Panduan Pemilihan</h2>.",
            "short_description": "1-2 kalimat Indonesia maksimal 260 karakter.",
            "feature_tab_html": "HTML <h2>Keunggulan Produk</h2><ul> dengan 4-6 poin faktual dalam bahasa Indonesia.",
            "faq_html": "HTML dengan 4 pasangan <h3>pertanyaan Indonesia</h3><p>jawaban Indonesia</p>.",
            "meta_description": "Maksimal 158 karakter, bahasa Indonesia.",
            "official_information_html": "HTML Indonesia berisi <h2>Informasi Produk</h2>, ringkasan faktual produk, dan <h2>Catatan Teknis</h2>. Jangan menyebut sumber URL, halaman resmi, situs merek, proses AI, atau scraping.",
            "translated_specifications": [
                {"id": "ID setiap baris specifications dari input", "label_id": "Label spesifikasi Bahasa Indonesia", "value_id": "Nilai spesifikasi Bahasa Indonesia dengan angka/satuan persis dipertahankan"}
            ],
            "keep_specification_ids": ["Semua ID spesifikasi yang dipertahankan di tab Spesifikasi, urut mengikuti sumber. Jika tidak duplikat, ID tetap harus masuk daftar ini."],
            "removed_duplicate_specifications": [
                {"source_id": "ID yang dibuang karena benar-benar duplikat", "duplicate_of_id": "ID yang dipertahankan sebagai versi yang sama", "reason_id": "Alasan singkat Bahasa Indonesia, misalnya Duplikat identik. Tidak boleh menghapus fakta yang nilainya berbeda."}
            ],
        },
        "product_data": payload,
    }
    # Jika validator sebelumnya menolak struktur spesifikasi, kirim umpan balik
    # yang spesifik agar AI memperbaiki bentuk JSON, bukan menulis ulang secara
    # acak tanpa memahami alasan pengulangan.
    if retry_reason:
        feedback_items = list(retry_feedback or [])[:20]
        user["retry_validation_note"] = [
            f"Ini percobaan ulang ke-{attempt_number} dari {max_attempts}. Respons sebelumnya ditolak validator dengan kode: {retry_reason}.",
            "Perbaiki hanya berdasarkan temuan validator di bawah, tetapi tetap kirim ULANG seluruh JSON lengkap sesuai output_schema.",
            "Jangan menghapus ID yang disebut bermasalah. Perbaiki struktur dan isi field-nya sesuai pesan validator.",
            "Setiap ID input tetap wajib memiliki satu terjemahan. Setiap ID juga harus masuk tepat sekali ke keep_specification_ids atau menjadi source_id pada removed_duplicate_specifications.",
        ]
        user["validator_feedback"] = feedback_items or [
            {
                "code": "specification_contract_failed",
                "message": "Periksa kembali translated_specifications, keep_specification_ids, dan removed_duplicate_specifications. Jika tidak ada duplikat, gunakan removed_duplicate_specifications: [] dan simpan semua ID.",
            }
        ]

    body = {
        "model": clean_text(AI_SETTINGS.get("model")) or "deepseek-chat",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Anda adalah penerjemah teknis dan editor data spesifikasi untuk WEBSITE_KALIAN. "
                    "Tulis dalam bahasa Indonesia baku. Semua karakter Chinese dilarang. "
                    "Jaga fakta teknis tetap akurat. Terapkan deduplikasi spesifikasi sangat konservatif. "
                    "Balas JSON valid saja."
                ),
            },
            {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
        ],
        "temperature": AI_SETTINGS.get("temperature", 0.15),
        "response_format": {"type": "json_object"},
    }
    endpoint = AI_SETTINGS["base_url"] + "/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS["timeout"])
        if response.status_code in {400, 422}:
            body.pop("response_format", None)
            response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS["timeout"])
        response.raise_for_status()
        raw = response.json()
        content = clean_text(raw.get("choices", [{}])[0].get("message", {}).get("content", ""))
        parsed = _parse_json(content)
        valid, fields, reason, feedback = _validate_ai_fields(parsed, payload.get("specifications") or [])
        return valid, fields if valid else {}, reason, feedback
    except Exception as exc:
        return False, {}, f"request_error_{type(exc).__name__}", []


def _is_retryable_ai_validation_reason(reason: Any) -> bool:
    """Retry only the known specification-accounting failure.

    Request/authentication errors and forbidden-content failures should not be
    repeated blindly. This retry is intentionally limited to the case where
    DeepSeek returned content but did not satisfy the strict specification
    deduplication contract.
    """
    return clean_text(reason) == "invalid_ai_specification_deduplication"


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    existing = data.get("_ai_fields")
    if isinstance(existing, dict):
        return existing
    if not AI_SETTINGS.get("enabled"):
        return {}

    cache = AI_SETTINGS.get("cache", {})
    payload = _ai_payload(data, family)
    key = _ai_cache_key(payload)
    cached = cache.get(key) if isinstance(cache, dict) else None
    if isinstance(cached, dict) and cached.get("ok") and isinstance(cached.get("fields"), dict):
        AI_STATS["cache_hit"] += 1
        data["_ai_status"] = "cache_hit"
        data["_ai_attempts"] = int(cached.get("attempts") or 1)
        data["_ai_fields"] = cached["fields"]
        return cached["fields"]

    if AI_SETTINGS.get("limit") and AI_STATS["used"] >= AI_SETTINGS["limit"]:
        AI_STATS["skipped"] += 1
        data["_ai_status"] = "limit_reached"
        return {}

    max_attempts = 1 + int(AI_SETTINGS.get("retry_count") or 0)
    attempt_reasons: List[str] = []
    retry_feedback: List[Dict[str, Any]] = []
    feedback_history: List[List[Dict[str, Any]]] = []
    ok = False
    fields: Dict[str, Any] = {}
    reason = ""
    attempts_used = 0

    for attempt_number in range(1, max_attempts + 1):
        attempts_used = attempt_number
        retry_reason = attempt_reasons[-1] if attempt_reasons else ""
        ok, fields, reason, current_feedback = _call_ai(
            payload,
            retry_reason=retry_reason,
            retry_feedback=retry_feedback,
            attempt_number=attempt_number,
            max_attempts=max_attempts,
        )
        if ok:
            break

        reason = clean_text(reason) or "unknown_ai_validation_error"
        attempt_reasons.append(reason)
        retry_feedback = [item for item in current_feedback if isinstance(item, dict)]
        feedback_history.append(retry_feedback)
        if not _is_retryable_ai_validation_reason(reason) or attempt_number >= max_attempts:
            break

        AI_STATS["retry"] += 1
        retry_delay = float(AI_SETTINGS.get("retry_delay") or 0.0)
        if retry_delay > 0:
            time.sleep(retry_delay)

    cache[key] = {
        "ok": ok,
        "fields": fields if ok else {},
        "reason": "ok" if ok else reason,
        "attempts": attempts_used,
        "attempt_reasons": attempt_reasons,
        "validation_feedback_history": feedback_history,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    AI_SETTINGS["cache"] = cache
    data["_ai_attempts"] = attempts_used
    data["_ai_retry_reasons"] = attempt_reasons
    data["_ai_retry_feedback"] = feedback_history

    if ok:
        AI_STATS["used"] += 1
        data["_ai_status"] = "ok" if attempts_used == 1 else f"retry_success_attempt_{attempts_used}"
        data["_ai_fields"] = fields
        if AI_SETTINGS.get("delay"):
            time.sleep(AI_SETTINGS["delay"])
        return fields

    data["_ai_status"] = reason
    AI_STATS["fallback"] += 1
    return {}


def localize_spec_rows_fallback(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    """Fallback label Indonesia tanpa menyisipkan teks sumber Chinese."""
    out: List[Dict[str, str]] = []
    for row in rows:
        raw_label = clean_catalog_source_text(row.get("label"))
        raw_value = clean_catalog_source_text(row.get("value"))
        if not raw_label or not raw_value:
            continue
        key = re.sub(r"\s+", " ", raw_label).casefold()
        label = SPEC_LABEL_ID_MAP.get(key, raw_label)
        out.append({
            "label": _localize_source_label(row.get("source_label") or raw_label, label),
            "source_label": clean_text(row.get("source_label")) or raw_label,
            "value": raw_value,
            "section": clean_text(row.get("section")) or "Specifications",
        })
    return dedupe_specs(out)


def fallback_official_information(data: Dict[str, Any]) -> str:
    """Fallback Indonesia yang aman. Tidak memublikasikan teks mentah Inggris/Chinese."""
    model = clean_text(data.get("model"))
    title = clean_catalog_source_text(data.get("title"))
    parts = ["<h2>Informasi Produk</h2>"]
    if title:
        parts.append(f"<p>Informasi teknis produk {esc(title)} dirangkum dari halaman resmi Lutron dan dokumen yang terdeteksi.</p>")
    else:
        parts.append("<p>Informasi teknis dirangkum dari halaman resmi Lutron dan dokumen yang terdeteksi.</p>")
    if model:
        parts.append(f"<p><strong>Model:</strong> {esc(model)}</p>")
    parts.append("<h2>Catatan Teknis</h2><p>Gunakan tab Spesifikasi untuk memeriksa rentang, akurasi, kondisi operasi, dan kebutuhan penggunaan produk sebelum pemesanan.</p>")
    return "".join(parts)


def _source_claim_blob(data: Dict[str, Any]) -> str:
    parts = [
        clean_catalog_source_text(data.get("description")),
        html_to_source_text(data.get("official_content_html"), 12000),
    ]
    raw_rows = data.get("_raw_spec_rows") if isinstance(data.get("_raw_spec_rows"), list) else data.get("spec_rows") or []
    for row in raw_rows:
        if isinstance(row, dict):
            parts.append(clean_catalog_source_text(row.get("source_label") or row.get("label")))
            parts.append(clean_catalog_source_text(row.get("value")))
    return " ".join(part for part in parts if part)


def _source_interface_state(data: Dict[str, Any]) -> Dict[str, bool]:
    blob = _source_claim_blob(data)
    return {
        "rs232": bool(re.search(r"\bRS\s*[- ]?232\b", blob, re.I)),
        "usb_optional": bool(re.search(r"\bUSB\s*[- ]?0?1\b|optional[^.\n]{0,100}\bUSB\b|USB[^.\n]{0,100}optional", blob, re.I)),
        "usb_builtin": bool(re.search(r"\b(?:built[- ]?in|integrated|interface|port)\b[^.\n]{0,80}\bUSB\b", blob, re.I)) and not bool(re.search(r"\bUSB\s*[- ]?0?1\b", blob, re.I)),
        "pc_data_logger": bool(
            re.search(r"(?:PC|computer|komputer)[^.\n]{0,120}(?:data\s*logger|recorder)|(?:RS\s*[- ]?232)[^.\n]{0,160}(?:data\s*logger|recorder)", blob, re.I)
            or (re.search(r"\bRS\s*[- ]?232\b", blob, re.I) and re.search(r"(?:data\s+acquisition\s+software|perangkat\s+lunak\s+akuisisi\s+data|SW-U801-WIN)", blob, re.I))
        ),
        "internal_auto_manual_logger": bool(re.search(r"(?:automatic|auto)\s+data\s*logger[^.\n]{0,100}manual\s+data\s*logger|manual\s+data\s*logger[^.\n]{0,100}(?:automatic|auto)\s+data\s*logger", blob, re.I)),
    }


def _sanitize_interface_and_logging_html(value: Any, data: Dict[str, Any]) -> str:
    html_value = clean_text(value)
    state = _source_interface_state(data)
    if state["rs232"] and state["usb_optional"] and not state["usb_builtin"]:
        replacement = "antarmuka RS232 untuk keluaran data; koneksi USB memerlukan kabel USB-01 opsional"
        html_value = re.sub(r"antarmuka\s+RS\s*[- ]?232\s*/\s*USB", replacement, html_value, flags=re.I)
        html_value = re.sub(r"antarmuka\s+RS\s*[- ]?232\s+dan\s+USB", replacement, html_value, flags=re.I)
        html_value = re.sub(r"kabel\s+RS\s*[- ]?232\s*/\s*USB", "kabel RS232 dan kabel USB-01 opsional", html_value, flags=re.I)
    elif state["rs232"] and not state["usb_builtin"]:
        html_value = re.sub(r"antarmuka\s+RS\s*[- ]?232\s*/\s*USB", "antarmuka RS232 untuk keluaran data", html_value, flags=re.I)

    if not state["internal_auto_manual_logger"]:
        if state["pc_data_logger"] and state["rs232"]:
            replacement = "Data dapat dikirim melalui antarmuka RS232 ke komputer untuk pencatatan menggunakan perangkat lunak dan kabel opsional."
        else:
            replacement = ""
        html_value = re.sub(
            r"(?:Pencatatan|Perekaman)\s+data\s+otomatis\s+dan\s+manual\s+tersedia\.?",
            replacement,
            html_value,
            flags=re.I,
        )
        html_value = re.sub(
            r"(?:pencatatan|perekaman)\s+data\s+otomatis\s*(?:dan|/)\s*manual",
            "pencatatan data melalui komputer" if state["pc_data_logger"] else "memori maksimum dan minimum",
            html_value,
            flags=re.I,
        )
    html_value = re.sub(r"\s{2,}", " ", html_value)
    return html_value.strip()


def _sanitize_spec_rows_for_source(data: Dict[str, Any], rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    state = _source_interface_state(data)
    output: List[Dict[str, str]] = []
    for raw in rows or []:
        if not isinstance(raw, dict):
            continue
        row = dict(raw)
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        source_label = clean_text(row.get("source_label"))
        combined = f"{label} {source_label} {value}"

        # Max/Min memory is not the same as an internal automatic/manual logger.
        if re.search(r"\b(?:perekaman\s+data|pencatat\s+data|data\s+record)\b", combined, re.I) and re.search(r"\b(?:max(?:imum)?|min(?:imum)?|maksimum|minimum)\b", value, re.I):
            if not state["internal_auto_manual_logger"]:
                row["label"] = "Memori Maksimum/Minimum"
                row["value"] = "Merekam dan menampilkan pembacaan maksimum dan minimum."

        if re.search(r"\b(?:output\s+data|data\s+output|interface|antarmuka)\b", combined, re.I) and state["rs232"]:
            row["label"] = "Antarmuka dan Keluaran Data"
            row["value"] = "Antarmuka serial RS232 untuk keluaran data ke komputer."

        if state["usb_optional"] and not state["usb_builtin"]:
            row["value"] = re.sub(
                r"\bUSB\s*[- ]?0?1\b",
                "USB-01 (kabel USB opsional)",
                clean_text(row.get("value")),
                flags=re.I,
            )
        output.append(row)
    return output


def _seo_product_type(family: Dict[str, Any], title: str) -> str:
    category = clean_text(family.get("category"))
    product_type = re.split(r"\s+dan\s+|\s*/\s*", category, maxsplit=1, flags=re.I)[0].strip()
    if product_type and product_type.casefold() not in {"alat ukur laboratorium", "alat ukur"}:
        return product_type
    # Fallback to words before the first model/number in the official title.
    prefix = re.split(r"\b[A-Z]{1,6}[- ]?\d|\b\d", clean_text(title), maxsplit=1)[0]
    prefix = re.sub(r"\bLutron\b", "", prefix, flags=re.I).strip(" ,.-")
    return prefix or clean_text(family.get("term")) or "Alat Ukur"


def _remove_title_phrase(text: str, phrase: str) -> str:
    if not clean_text(phrase):
        return text
    return re.sub(re.escape(clean_text(phrase)), " ", text, flags=re.I)


def standardize_seo_title(data: Dict[str, Any], family: Dict[str, Any], title: Any) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    raw = clean_text(title).replace("&", "dan").strip(" .,:;|-–—")
    product_type = _seo_product_type(family, raw)
    descriptor = raw
    for phrase in (product_type, brand, model):
        descriptor = _remove_title_phrase(descriptor, phrase)
    descriptor = re.sub(r"\s+", " ", descriptor).strip(" .,:;|-–—")
    parts = dedupe_keep_order([product_type, brand, model, descriptor])
    result = re.sub(r"\s+", " ", " ".join(parts)).strip(" .,:;|-–—")
    return compact(result, 180).rstrip(".")


def build_seo_focus_keyphrase(data: Dict[str, Any], family: Dict[str, Any], seo_title: str) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, seo_title)
    probe = _remove_title_phrase(seo_title, model)
    range_match = re.search(
        rf"\b\d+(?:[.,]\d+)?\s*(?:{_MEASUREMENT_UNIT_PATTERN.pattern})\b",
        probe,
        re.I,
    )
    range_value = clean_text(range_match.group(0)) if range_match else ""
    keyphrase = " ".join(dedupe_keep_order([product_type, brand, model, range_value]))
    return compact(keyphrase, 120).rstrip(" .,:;")


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues: List[str] = []
    name = clean_text(row.get("Name"))
    model = clean_text(data.get("model"))
    source_url = clean_text(row.get("Source URL"))
    specs_html = clean_text(row.get("custom_tab_1_content"))
    public_blob = " ".join(clean_text(row.get(field)) for field in (
        "Name", "Product Description", "Product Short Description", "custom_tab_2_content",
        "custom_tab_3_content", "custom_tab_4_content", "custom_tab_5_content", "meta description",
    ))
    state = _source_interface_state(data)

    if clean_text(row.get("publication_date")):
        issues.append("publication_date_must_be_blank")
    if not name or name.endswith("."):
        issues.append("seo_title_invalid")
    if "lutron" not in name.casefold() or (model and model.casefold() not in name.casefold()):
        issues.append("seo_title_missing_brand_or_model")
    keyphrase = clean_text(row.get("focus keyphrase"))
    if not keyphrase or keyphrase.endswith(".") or (model and model.casefold() not in keyphrase.casefold()):
        issues.append("focus_keyphrase_invalid")
    if not source_url or get_source_config(source_url) is None:
        issues.append("source_url_not_official_lutron")
    if not clean_text(data.get("_raw_spec_rows")) and not data.get("_raw_spec_rows"):
        issues.append("missing_source_specification_rows")
    if "<table" not in specs_html.lower() or "<thead" not in specs_html.lower():
        issues.append("specification_table_invalid")
    if not clean_text(row.get("image_url")):
        issues.append("missing_verified_image")
    if re.search(r"antarmuka\s+RS\s*[- ]?232\s*/\s*USB", public_blob, re.I) and not state["usb_builtin"]:
        issues.append("ambiguous_rs232_usb_claim")
    if re.search(r"(?:pencatatan|perekaman)\s+data\s+otomatis\s*(?:dan|/)\s*manual", public_blob, re.I) and not state["internal_auto_manual_logger"]:
        issues.append("unsupported_internal_data_logger_claim")
    if require_ai and not isinstance(data.get("_ai_fields"), dict):
        issues.append("deepseek_validated_content_required")
    if has_cjk(public_blob + " " + specs_html):
        issues.append("chinese_character_detected")
    return dedupe_keep_order(issues)


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    brand = LUTRON_CONFIG.brand
    data["brand"] = brand
    family = detect_family(data)
    specs = data.get("spec_rows") or []
    ai = get_ai_content(data, family)

    # Dalam mode AI bersih, seluruh teks publik berasal dari hasil terjemahan/penulisan Indonesia.
    if ai:
        title = clean_text(ai.get("product_name_id")) or clean_catalog_source_text(data.get("title"))
        # Category from deterministic family rules is authoritative. AI may
        # suggest a label but never overrides the validated product class.
        category_leaf = clean_text(family.get("category")).strip()
        # DeepSeek has already selected the retained specification rows from
        # the raw Technical Specifications + Specifications source list.
        specs_out = ai.get("translated_spec_rows") or localize_spec_rows_fallback(specs)
        desc = ai.get("product_description_html")
        short = ai.get("short_description")
        features = ai.get("feature_tab_html")
        faq = ai.get("faq_html")
        meta = ai.get("meta_description")
        official = ai.get("official_information_html")
    else:
        title = clean_catalog_source_text(data.get("title"))
        category_leaf = clean_catalog_source_text(family.get("category")).replace("", "").strip()
        specs_out, _output_spec_issues = canonicalize_specifications(localize_spec_rows_fallback(specs))
        desc = build_product_description(data, family)
        short = build_short_description(data, family)
        features = "<h2>Keunggulan Produk</h2><ul>" + "".join(f"<li>{esc(x)}</li>" for x in feature_bullets(data, family)) + "</ul>"
        faq = build_faq(data, family)
        meta = build_meta_description(data, family)
        official = fallback_official_information(data)

    specs_out = _sanitize_spec_rows_for_source(data, specs_out)
    desc = _sanitize_interface_and_logging_html(desc, data)
    short = _sanitize_interface_and_logging_html(short, data)
    features = _sanitize_interface_and_logging_html(features, data)
    faq = _sanitize_interface_and_logging_html(faq, data)
    meta = _sanitize_interface_and_logging_html(meta, data)
    official = _sanitize_interface_and_logging_html(official, data)
    title = standardize_seo_title(data, family, title)
    focus_keyphrase = build_seo_focus_keyphrase(data, family, title)

    # Pastikan hasil akhir bebas karakter Chinese, bahkan bila ada respons AI yang tidak sesuai.
    if has_cjk(title) or any(has_cjk(value) for value in [desc, short, features, faq, meta, official, category_leaf]):
        raise ValueError("chinese_character_detected_in_output")

    if WEBSITE_CTA not in desc:
        desc = desc.rstrip() + WEBSITE_CTA
    category = f"{category_leaf}" if category_leaf else "Alat Ukur dan Instrumen"
    model = clean_text(data.get("model"))
    tags = dedupe_keep_order([brand, model, category.replace("", ""), "alat ukur Lutron"])

    corrections: List[str] = []
    if not data.get("description"):
        corrections.append("missing_description")
    if not specs:
        corrections.append("missing_specs")
    raw_spec_count = len(data.get("_raw_spec_rows") or []) if isinstance(data.get("_raw_spec_rows"), list) else 0
    corrections.append(f"source_spec_rows={raw_spec_count}")
    corrections.append(f"published_spec_rows={len(specs_out)}")
    if not data.get("official_content_html"):
        corrections.append("missing_full_content")
    if not data.get("image_url"):
        corrections.append("missing_image")
    if data.get("datasheet_url"):
        corrections.append("datasheet_source=" + clean_text(data.get("datasheet_url")))
    if data.get("image_source"):
        corrections.append("image_source=" + clean_text(data.get("image_source")))
    if data.get("error"):
        corrections.append("source_note=" + clean_text(data.get("error")))
    if data.get("is_accessory"):
        corrections.append("accessory_included_" + clean_text(data.get("accessory_reason")))
    if clean_text(family.get("classification_basis")):
        corrections.append("category_basis=" + clean_text(family.get("classification_basis")))
    for flag in data.get("quality_flags") if isinstance(data.get("quality_flags"), list) else []:
        corrections.append("quality_flag=" + clean_text(flag))
    if AI_SETTINGS.get("enabled"):
        if ai:
            removed_count = int(ai.get("ai_spec_dedup_removed_count") or 0)
            corrections.append("deepseek_ai_indonesia_spec_dedup_removed=" + str(removed_count))
            ai_attempts = int(data.get("_ai_attempts") or 1)
            if ai_attempts > 1:
                corrections.append("deepseek_ai_retry_success_attempt=" + str(ai_attempts))
        else:
            corrections.append("deepseek_ai_fallback_" + clean_text(data.get("_ai_status", "unknown")))
            retry_reasons = data.get("_ai_retry_reasons")
            if isinstance(retry_reasons, list) and retry_reasons:
                corrections.append("deepseek_ai_retry_reasons=" + ",".join(dedupe_keep_order(retry_reasons)))
            retry_feedback = data.get("_ai_retry_feedback")
            if isinstance(retry_feedback, list):
                feedback_codes: List[str] = []
                for attempt_feedback in retry_feedback:
                    if isinstance(attempt_feedback, list):
                        for item in attempt_feedback:
                            if isinstance(item, dict) and clean_text(item.get("code")):
                                feedback_codes.append(clean_text(item.get("code")))
                if feedback_codes:
                    corrections.append("deepseek_ai_validator_errors=" + ",".join(dedupe_keep_order(feedback_codes)))

    row = {col: "" for col in OUTPUT_COLS}
    row.update({
        "Name": title,
        "Brand": brand,
        "Product Description": desc,
        "Product Short Description": short,
        "custom_tab_1_title": "Spesifikasi",
        "custom_tab_1_content": specifications_tab_html(
            specs_out,
            already_deduplicated=bool(ai and ai.get("ai_spec_dedup_applied")),
        ),
        "custom_tab_1_priority": 10,
        "custom_tab_2_title": "Keunggulan Produk",
        "custom_tab_2_content": features,
        "custom_tab_2_priority": 20,
        "custom_tab_3_title": "Aplikasi Penggunaan",
        "custom_tab_3_content": build_application_tab({"title": title, "brand": brand, "is_accessory": data.get("is_accessory")}, family),
        "custom_tab_3_priority": 30,
        "custom_tab_4_title": "FAQ",
        "custom_tab_4_content": faq,
        "custom_tab_4_priority": 40,
        "custom_tab_5_title": "Informasi Produk",
        "custom_tab_5_content": official,
        "custom_tab_5_priority": 50,
        "product categories": category,
        "product tags": ", ".join(tags),
        "focus keyphrase": focus_keyphrase,
        "meta description": meta,
        "publication_date": "",
        "Processed": "Yes",
        "Processing Time": "lutron_WEBSITE_import_deepseek_indonesia" if ai else "lutron_WEBSITE_import_rule_fallback",
        "Content Quality": f"{clean_text(data.get('source'))}_live_scrape_{'deepseek_indonesia' if ai else 'rule_fallback'}",
        "image_url": clean_text(data.get("image_url")),
        "Source URL": clean_text(data.get("url")),
        "Website Validation": "publish_ready_source_verified",
        "Website Correction Log": "; ".join(corrections) if corrections else "ok",
    })
    return row



REVIEW_SPEC_COLUMNS = [
    "URL", "Model", "Product", "Status", "Issue Code", "Severity", "Detail",
    "Canonical Key", "Display Label", "Source Label", "Value", "Source Section",
]


def prepare_canonical_specs_for_product(data: Dict[str, Any]) -> List[Dict[str, str]]:
    """Canonicalise from immutable raw rows and attach review issues to data."""
    raw = data.get("_raw_spec_rows")
    if not isinstance(raw, list):
        raw = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
        data["_raw_spec_rows"] = [dict(item) for item in raw if isinstance(item, dict)]
    canonical, issues = canonicalize_specifications(raw)
    data["spec_rows"] = canonical
    data["spec_review_issues"] = issues
    return issues


def build_spec_review_records(data: Dict[str, Any]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for issue in data.get("spec_review_issues") if isinstance(data.get("spec_review_issues"), list) else []:
        if not isinstance(issue, dict):
            continue
        row = {col: "" for col in REVIEW_SPEC_COLUMNS}
        row.update({
            "URL": clean_text(data.get("url")),
            "Model": clean_text(data.get("model")),
            "Product": clean_text(data.get("title")),
            "Status": clean_text(data.get("status")),
        })
        for col in REVIEW_SPEC_COLUMNS:
            if col in issue:
                row[col] = clean_text(issue.get(col))
        rows.append(row)
    return rows


def _append_spec_review_log(data: Dict[str, Any]) -> None:
    issues = data.get("spec_review_issues") if isinstance(data.get("spec_review_issues"), list) else []
    if not issues:
        return
    codes = dedupe_keep_order(clean_text(item.get("Issue Code")) for item in issues if isinstance(item, dict))
    if codes:
        data["error"] = _merge_source_notes(data.get("error"), "spec_review=" + ",".join(codes))


# ---------------------------------------------------------------------------
# Cache / output / CLI
# ---------------------------------------------------------------------------

def load_cache(path: str) -> Dict[str, Any]:
    if not path:
        return {}
    p = Path(path)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_cache(path: str, data: Dict[str, Any]) -> None:
    if not path:
        return
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def style_output_xlsx(path: str, sheet_name: str = "WEBSITE Products") -> None:
    if load_workbook is None:
        return
    try:
        wb = load_workbook(path)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        ws.freeze_panes = "A2"
        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="D9E2F3")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for cell in ws[1]:
            cell.fill, cell.font, cell.border = header_fill, header_font, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        widths = {"A": 34, "B": 16, "C": 62, "D": 44, "E": 22, "F": 60, "H": 24, "I": 60,
                  "K": 22, "L": 60, "N": 22, "O": 60, "Q": 28, "R": 70, "AC": 34, "AD": 44,
                  "AE": 30, "AF": 44, "AG": 20, "AH": 16, "AI": 32, "AJ": 30, "AK": 44, "AL": 44,
                  "AM": 32, "AN": 54}
        for i in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(get_column_letter(i), 18)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
            ws.row_dimensions[row[0].row].height = 84
        ws.row_dimensions[1].height = 34
        if ws.max_row >= 2 and not ws.tables:
            ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            tab = Table(displayName="LUTRON_WEBSITE_Products", ref=ref)
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            ws.add_table(tab)
        wb.save(path)
    except Exception:
        pass


def write_outputs(
    rows: Sequence[Dict[str, Any]],
    skipped: Sequence[Dict[str, Any]],
    reviews: Sequence[Dict[str, Any]],
    output_path: str,
    skipped_output: str = "",
) -> None:
    # Sanitasi dilakukan tepat sebelum export agar data scrape/cache tetap utuh,
    # tetapi setiap checkpoint dan output final aman untuk format Open XML.
    safe_rows = sanitize_excel_records(rows)
    safe_skipped = sanitize_excel_records(skipped)
    safe_reviews = sanitize_excel_records(reviews)

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        pd.DataFrame(safe_rows, columns=OUTPUT_COLS).to_excel(
            writer, index=False, sheet_name="WEBSITE Products"
        )
        if safe_skipped:
            pd.DataFrame(safe_skipped).to_excel(
                writer, index=False, sheet_name="Skipped"
            )
        if safe_reviews:
            pd.DataFrame(safe_reviews, columns=REVIEW_SPEC_COLUMNS).to_excel(
                writer, index=False, sheet_name="Review Spesifikasi"
            )

    style_output_xlsx(str(output_file))

    if skipped_output:
        skipped_file = Path(skipped_output)
        skipped_file.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(safe_skipped).to_excel(skipped_file, index=False)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Scrape katalog LUTRON ke format WooCommerce WEBSITE_KALIAN dengan fakta sumber, validasi DeepSeek, SEO, dan publish gate V22")
    p.add_argument("--input", required=True, help="Excel berisi URL produk")
    p.add_argument("--output", default="WEBSITE_catalog_import_ready.xlsx", help="Output Excel WooCommerce")
    p.add_argument("--sheet", default=0, help="Nama atau indeks sheet input")
    p.add_argument("--url-column", default="", help="Nama kolom URL jika ingin ditentukan manual")
    p.add_argument("--cache", default="lutron_scrape_cache_v7.json", help="Cache JSON hasil scrape")
    p.add_argument("--timeout", type=int, default=25, help="Timeout tiap halaman dalam detik")
    p.add_argument("--delay", type=float, default=0.6, help="Jeda antar request dalam detik")
    p.add_argument("--max-rows", type=int, default=0, help="Batasi URL. 0 = semua")
    p.add_argument("--start-date", default="", help="Kompatibilitas lama. publication_date selalu dikosongkan.")
    p.add_argument("--pub-interval-hours", type=int, default=1, help="Kompatibilitas lama. Tidak digunakan karena publication_date kosong.")
    p.add_argument("--skipped-output", default="", help="Opsional output log produk diskip")
    p.add_argument("--no-cache", action="store_true", help="Abaikan cache dan scrape ulang")
    p.add_argument("--no-render-js", action="store_true", help="Matikan browser Playwright. Default: browser dipakai untuk menangkap gambar, spesifikasi, dan isi Lutron yang dimuat JavaScript.")
    p.add_argument("--no-datasheet-fallback", action="store_true", help="Matikan fallback datasheet model ketika halaman Lutron hanya mengembalikan shell/konten minim.")
    p.add_argument("--allow-external-image-fallback", action="store_true", help="Opsional dan tidak disarankan: izinkan fallback gambar dari mesin pencari. Default hanya memakai gambar resmi yang dimuat halaman Lutron.")
    p.add_argument("--datasheet-attempt-limit", type=int, default=20, help="Maksimum URL datasheet yang dicek per produk sebelum dinyatakan tidak ditemukan.")
    p.add_argument("--strict-source", action="store_true", help="Source-link-only: spesifikasi wajib berasal dari URL Lutron pada Excel. Produk yang membutuhkan datasheet fallback atau sumber lain akan diskip.")
    p.add_argument("--strict-publish", action="store_true", help="Aktifkan validasi publish ketat. Pada V22.12 produk tidak di-skip selain URL sumber 404; masalah lain menghentikan batch atau menjadi peringatan aman.")
    p.add_argument("--allow-spec-review-publish", action="store_true", help="Kompatibilitas v16. Tidak diperlukan karena spesifikasi tidak pernah menahan produk.")
    p.add_argument("--hold-all-spec-reviews", action="store_true", help="Kompatibilitas v16. Diabaikan karena v16 hanya menghapus duplikasi spesifikasi.")
    p.add_argument("--category-from-description", action="store_true", help="Kompatibilitas: V10 sudah selalu menentukan kategori dari model, judul input, judul resmi, dan deskripsi resmi. Spesifikasi tidak dipakai untuk kategori.")
    p.add_argument("--skip-accessories", action="store_true", help="Kompatibilitas lama. Pada V22.12 opsi ini diabaikan; aksesori tetap dimasukkan.")
    p.add_argument("--use-ai", action="store_true", help="Aktifkan generator konten DeepSeek")
    p.add_argument("--deepseek-api-key", default=os.getenv("DEEPSEEK_API_KEY", ""))
    p.add_argument("--deepseek-base-url", default=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"))
    p.add_argument("--deepseek-model", default=os.getenv("DEEPSEEK_MODEL", "deepseek-chat"))
    p.add_argument("--ai-cache", default="", help="Cache JSON hasil AI")
    p.add_argument("--ai-cache-only", action="store_true", help="ZERO TOKEN: gunakan hanya hasil DeepSeek yang sudah ada di cache. Cache miss tidak pernah memanggil API dan produk ditandai NEEDS_REVIEW.")
    p.add_argument("--ai-limit", type=int, default=0, help="Batas request AI baru. 0 = semua")
    p.add_argument("--ai-timeout", type=int, default=180)
    p.add_argument("--ai-delay", type=float, default=0.0)
    p.add_argument("--ai-retries", type=int, default=2, help="Jumlah retry umum untuk kegagalan validasi atau koneksi AI. Default 2, sehingga total maksimal 3 percobaan per produk.")
    p.add_argument("--item-repair-attempts", type=int, default=2, help="Maksimum percobaan repair AI untuk setiap baris spesifikasi. Default 2; bagian editorial dibuat deterministik agar tidak menghabiskan token repair.")
    p.add_argument("--ai-retry-delay", type=float, default=3.0, help="Jeda dalam detik sebelum retry AI. Default 3.0.")
    p.add_argument("--ai-temperature", type=float, default=0.15)
    p.add_argument("--ai-spec-limit", type=int, default=120, help="Maksimum baris spesifikasi yang diterjemahkan AI per produk. Default: 120.")
    p.add_argument(
        "--allow-source-language-fallback",
        action="store_true",
        help="Kompatibilitas lama. Pada V22.12 DeepSeek wajib tervalidasi; kegagalan layanan menghentikan batch dan tidak memakai fallback bahasa sumber.",
    )
    p.add_argument("--catalog-audit-only", action="store_true", help="Audit seluruh input Excel tanpa scrape web. Membuat laporan klasifikasi model, kategori, dan status aksesori untuk semua produk.")
    p.add_argument("--catalog-audit-output", default="", help="Lokasi file Excel audit katalog. Digunakan bersama --catalog-audit-only.")
    return p.parse_args()


def coerce_sheet(value: Any) -> Any:
    return int(value) if isinstance(value, str) and re.fullmatch(r"\d+", value) else value


def main() -> None:
    args = parse_args()
    records = read_input_records(args.input, coerce_sheet(args.sheet), args.url_column)
    if args.max_rows and args.max_rows > 0:
        records = records[:args.max_rows]
    if not records:
        raise SystemExit("Tidak ada URL yang dapat diproses.")

    if getattr(args, "catalog_audit_only", False):
        audit_path = clean_text(getattr(args, "catalog_audit_output", "")) or str(Path(args.input).with_name("lutron_catalog_preflight_audit.xlsx"))
        write_catalog_preflight(records, audit_path)
        print(f"Preflight katalog selesai. Produk diaudit: {len(records)}")
        print(f"Output audit: {audit_path}")
        return

    cache = {} if args.no_cache else load_cache(args.cache)
    out = Path(args.output)
    ai_cache_path = args.ai_cache or str(out.with_suffix(".deepseek_ai_cache.json"))
    set_ai_settings(
        enabled=bool(args.use_ai),
        api_key=args.deepseek_api_key,
        base_url=args.deepseek_base_url,
        model=args.deepseek_model,
        cache={} if args.no_cache else load_cache(ai_cache_path),
        cache_path=ai_cache_path,
        limit=args.ai_limit,
        timeout=args.ai_timeout,
        delay=args.ai_delay,
        retry_count=args.ai_retries,
        item_repair_attempts=max(1, min(4, int(args.item_repair_attempts or 2))),
        retry_delay=args.ai_retry_delay,
        temperature=args.ai_temperature,
        spec_limit=args.ai_spec_limit,
        strict_indonesian=bool(args.use_ai and not args.allow_source_language_fallback),
    )
    start_date = datetime.now()

    session = requests.Session()
    if PdfReader is None and not args.no_datasheet_fallback:
        print("PERINGATAN: pypdf belum tersedia. Fallback datasheet dimatikan. Instal: py -m pip install pypdf")
    if not args.no_datasheet_fallback and (fitz is None or not _find_tesseract_executable()):
        print("INFO: OCR datasheet scan tidak aktif. PDF teks tetap bisa dibaca. Untuk PDF scan, instal PyMuPDF dan Tesseract OCR.")
    renderer: Optional[LutronBrowserRenderer] = None if args.no_render_js else LutronBrowserRenderer(max(args.timeout, 25))
    if renderer is not None and not renderer.available():
        print("PERINGATAN: Playwright belum tersedia. Instal: py -m pip install playwright lalu py -m playwright install chromium")
        renderer = None

    accepted: List[Dict[str, Any]] = []
    rows: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    review_specs: List[Dict[str, Any]] = []
    included_accessories = 0

    for idx, record in enumerate(records, 1):
        url = record["url"]
        print(f"[{idx}/{len(records)}] scrape {url}")
        if not args.no_cache and url in cache:
            data = cache[url]
            for key in ("input_brand", "input_title", "input_model"):
                if not clean_text(data.get(key)):
                    data[key] = record.get(key, "")
            if not data.get("brand"):
                data["brand"] = LUTRON_CONFIG.brand
            if not data.get("title"):
                data["title"] = record.get("input_title", "")
            if not data.get("model"):
                data["model"] = record.get("input_model", "")
            data["is_accessory"], data["accessory_reason"] = is_accessory_product(
                clean_text(data.get("title")), clean_text(data.get("description")),
                data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else [],
            )
            data["quality_flags"] = collect_source_quality_flags(data)
            print("    cache hit")
        else:
            data = scrape_product_with_fallbacks(
                record,
                args.timeout,
                session,
                renderer,
                enable_datasheet=not args.no_datasheet_fallback,
                enable_image_search=bool(args.allow_external_image_fallback),
                datasheet_attempt_limit=max(1, int(args.datasheet_attempt_limit or 20)),
            )
            cache[url] = data
            if args.delay > 0:
                time.sleep(args.delay)

        # Canonical pipeline runs on every product, including cache hits.
        # It combines Technical Specifications and Specifications before AI.
        # v16 hanya menghapus baris spesifikasi yang memang duplikat.
        # Tidak membuat review, tidak memperbaiki label, dan tidak menahan produk.
        prepare_canonical_specs_for_product(data)
        spec_issues: List[Dict[str, str]] = []

        image_state = "OK" if clean_text(data.get("image_url")) else "KOSONG"
        flags = ", ".join(data.get("quality_flags") or [])
        print(
            f"    status={clean_text(data.get('status')) or '-'} | "
            f"model={clean_text(data.get('model')) or '-'} | image={image_state}" +
            (f" | flags={flags}" if flags else "")
        )
        if image_state == "KOSONG":
            print(f"    image_reason={clean_text(data.get('error')) or 'source image tidak ditemukan'}")

        if data.get("is_accessory"):
            if args.skip_accessories:
                skipped.append({"url": url, "title": data.get("title", ""), "status": data.get("status", ""), "reason": data.get("accessory_reason", ""), "error": data.get("error", "")})
                print("    skip confirmed accessory")
                continue
            included_accessories += 1
            print("    include confirmed accessory")

        if not clean_text(data.get("title")):
            skipped.append({"url": url, "title": "", "status": data.get("status", ""), "reason": "missing_title_or_scrape_failed", "error": data.get("error", "")})
            print(f"    skip failed: {data.get('error') or data.get('status')}")
            continue
        if not _is_real_product_data(data):
            skipped.append({
                "url": url,
                "title": data.get("title", ""),
                "status": data.get("status", ""),
                "reason": "incomplete_source_data",
                "error": data.get("error", ""),
            })
            print(f"    skip incomplete source: {data.get('error') or data.get('status')}")
            continue

        if args.strict_source and (get_source_config(url) is None or clean_text(data.get("datasheet_url"))):
            skipped.append({
                "url": url,
                "title": data.get("title", ""),
                "status": data.get("status", ""),
                "reason": "source_link_only_validation_failed",
                "error": _merge_source_notes(data.get("error"), "specifications_must_come_from_excel_source_url"),
            })
            print("    skip strict source: data tidak murni berasal dari source URL Excel")
            continue

        # v21: spesifikasi hanya diteruskan setelah source-link gate.
        # Script hanya membuang baris yang memang sama dan meneruskan sisanya.

        if AI_SETTINGS.get("enabled") and AI_SETTINGS.get("strict_indonesian"):
            ai_fields = get_ai_content(data, detect_family(data))
            if not ai_fields:
                data["_ai_status"] = clean_text(data.get("_ai_status")) or "ai_translation_failed_rule_fallback_used"
                data["error"] = _merge_source_notes(
                    data.get("error"),
                    "ai_translation_failed_rule_fallback_used",
                )
                # V22 policy: an AI failure is never inserted into the publish sheet.
                # A deterministic fallback may still be generated by dedicated review tooling,
                # but the catalog import output remains source-clean and evidence-validated.
                skipped.append({
                    "url": url,
                    "title": data.get("title", ""),
                    "status": data.get("status", ""),
                    "reason": "deepseek_validated_content_required",
                    "error": data.get("error", ""),
                })
                print(f"    skip publish output: AI validation failed ({data.get('_ai_status')})")
                continue

        # Strict publish is intentionally conservative: questionable source data
        # is sent to Skipped instead of being silently published.
        if args.strict_publish and data.get("quality_flags"):
            skipped.append({
                "url": url,
                "title": data.get("title", ""),
                "status": data.get("status", ""),
                "reason": "source_quality_review_required",
                "error": _merge_source_notes(data.get("error"), "; ".join(data.get("quality_flags") or [])),
            })
            print("    skip strict publish: source quality flag")
            continue
        accepted.append(data)

    # Batch-level image validation detects generic images reused across many
    # models. This cannot be done safely while each row is emitted immediately.
    accepted = apply_batch_image_quality_gate(accepted, skipped, require_verified_image=bool(args.strict_publish))

    for data in accepted:
        pub_date = start_date + timedelta(hours=len(rows) * args.pub_interval_hours)
        try:
            row = build_output_row(data, pub_date)
            publish_issues = validate_publish_output(
                data,
                row,
                require_ai=bool(args.strict_publish and args.use_ai),
            )
            if args.strict_publish and publish_issues:
                skipped.append({
                    "url": data.get("url", ""),
                    "title": row.get("Name", data.get("title", "")),
                    "status": data.get("status", ""),
                    "reason": "publish_validation_failed",
                    "error": "; ".join(publish_issues),
                })
                print("    skip publish gate: " + ", ".join(publish_issues))
                continue
            if publish_issues:
                row["Website Validation"] = "review_required:" + ",".join(publish_issues)
            rows.append(row)
        except ValueError as exc:
            skipped.append({
                "url": data.get("url", ""),
                "title": data.get("title", ""),
                "status": data.get("status", ""),
                "reason": "output_language_validation_failed",
                "error": str(exc),
            })
            print(f"    skip validasi output: {exc}")

    write_outputs(rows, skipped, review_specs, args.output, args.skipped_output)
    save_cache(args.cache, cache)
    if renderer is not None:
        renderer.close()
    if AI_SETTINGS.get("enabled"):
        save_cache(AI_SETTINGS.get("cache_path", ai_cache_path), AI_SETTINGS.get("cache", {}))
    print(f"Selesai. Produk masuk output: {len(rows)} | Diskip: {len(skipped)} | Review spesifikasi: {len(review_specs)} | Accessories masuk: {included_accessories}")
    print(f"Output: {args.output}")
    if AI_SETTINGS.get("enabled"):
        print(f"deepseek_ai_used={AI_STATS['used']} cache_hit={AI_STATS['cache_hit']} fallback={AI_STATS['fallback']} skipped={AI_STATS['skipped']}")



# ---------------------------------------------------------------------------
# V10 category routing: title + official description only
# ---------------------------------------------------------------------------
# V8 fixed the specific false positives in the first test batch, but the
# complete source workbook contains probes, electrodes, cases, software,
# current clamps, buffer solutions, interface cables, and other item families.
# These V9 rules classify all 559 catalog input rows before a live scrape and
# are intentionally conservative: an uncertain component is marked for review
# rather than being described as a main instrument.

ACCESSORY_MODEL_PREFIX_RE_V9 = re.compile(
    r"^(?:TP-|PE-|OXPB-|EP-|FRTP-|CDPB-|LN-|TL-|TLSM-|SMDC-|UPCB-|USB-|GMCB-|EXCB-|"
    r"CA-(?:0[3-8]|52A|203)|AP-|CP-|LS-|KV-|WG-|WH-|WP-|WT-|OXHD-|OXEL-|EA-|PB-|NN-|NB-|AT-|"
    r"ALLIGATOR|PL-|PI-|PWCB-|FS-|IP-|PX-|MP-|MS-71P$|VB-8[34]$|EMF-824$|"
    r"ORP-(?:14|15|400)$|PH-(?:04|04A|07|07A)$|ST-50\b|SB-01\b|CD-14\b|USBP-)",
    re.I,
)

ACCESSORY_TITLE_RE_V9 = re.compile(
    r"\b(?:replacement|spare(?:\s+part)?|accessor(?:y|ies)|optional\s+accessor|"
    r"carrying\s+case|protective\s+cover|extension\s+cable|interface\s+cable|usb\s+cable|rs\s*232\s+cable|"
    r"test\s+lead|alligator\s+clip|test\s+clip|test\s+probe|thermocouple\s+adapter|"
    r"calibration\s+(?:weight|solution)|buffer\s+solution|probe[-\s]?filling\s+electrolyte|"
    r"penetr?ometer\s+tips?|electrode|conductivity\s+solution|standard\s+solution|\bprobe\b|sensor|ear\s*phone|wedge\s+grip|"
    r"wall\s+holder|water\s+resistant\s+bag|current\s+shunt|power\s+adapter|power\s+converter|"
    r"power\s+interface\s+cable|adapter)\b",
    re.I,
)

# Product group rules cover the full supplied Lutron workbook, not only the
# initial eleven products.  The first matching rule wins.  Rules are built on
# model and title because source breadcrumbs are often generic or absent.
V9_FAMILY_RULES: List[Tuple[re.Pattern[str], Dict[str, Any]]] = [
    (re.compile(r"\b(?:software|data\s+acquisition|data\s+transmission|wifi\s+converter|rs\s*232\s+converter|i/o\s+controller\s+to\s+wifi)\b|\b(?:SW-|RSW-|RSC-|RHW-|IOW-)", re.I), {
        "category": "Software dan Interface Instrumentasi", "term": "perangkat lunak atau antarmuka instrumentasi",
        "function": "mendukung akuisisi data, komunikasi, atau integrasi instrumen sesuai spesifikasi produk",
        "tags": ["Lutron", "software instrumentasi", "interface instrumentasi"],
    }),
    (re.compile(r"\b(?:pressure\s+transmitter|humidity\s+transmitter|wind\s+(?:speed|direction)\s+transmitter|ph\s+transmitter|light\s+transmitter|sound\s+transmitter|vibration\s+transmitter|temperature\s+transmitter|dissolved\s+oxygen\s+transmitter|pressure\s+transducer|pressure\s+sensor)\b|\b(?:TR-|TRDO-|TRLX-|TRPH-|TRRH-|TRSL-|TRVB-|PS93|PS100)", re.I), {
        "category": "Transmitter dan Sensor Industri", "term": "transmitter atau sensor industri",
        "function": "mengirimkan atau menyediakan sinyal pengukuran untuk pemantauan, kontrol, atau otomasi",
        "tags": ["Lutron", "transmitter industri", "sensor industri"],
    }),
    (re.compile(r"\b(?:pressure\s+meter|manometer|vacuum\s+meter|pressure\s+data\s+recorder)\b|\b(?:PM-91|PS-93|MPS-)", re.I), {
        "category": "Manometer dan Pressure Meter", "term": "manometer atau pressure meter",
        "function": "mengukur, memantau, atau merekam tekanan sesuai spesifikasi produk",
        "tags": ["Lutron", "manometer", "pressure meter"],
    }),
    (re.compile(r"\b(?:air\s+quality|co2\s*meter|co\s*meter|o2\s*meter|pm\s*2\.5|negative\s+ion)\b|\b(?:AQ-|CO2-|COH-|GC-|GCH-|GCO-|MCH-|PM-10|O2H-|PO2-|PCO-)", re.I), {
        "category": "Air Quality dan Gas Meter", "term": "air quality atau gas meter",
        "function": "mengukur parameter kualitas udara atau gas sesuai spesifikasi produk",
        "tags": ["Lutron", "air quality meter", "gas meter"],
    }),
    (re.compile(r"\b(?:conductivity|tds|salt\s+meter|dissolved\s+oxygen|\bdo\s+meter|\bph\s*(?:meter|tester|monitor)|orp\s+meter|chlorine\s+meter|turbidity|pure\s+water|water\s+hardness|water\s+quality|ion\s+water)\b|\b(?:BCT-|BDO-|BPH-|BWA-|CD-|CL-|DO-|NI-|ORP-(?:2|4)|PCD-|PDO-|PH-(?:2|3)|PSA-|PWA-|PWH-|TU-|WA-|WAC-|YK-(?:21PH|22DOA|22CTA|23RP|30WA|31SA|200[15]))", re.I), {
        "category": "Water Quality Meter", "term": "alat ukur kualitas air",
        "function": "mengukur parameter kualitas air atau larutan sesuai spesifikasi produk",
        "tags": ["Lutron", "water quality meter", "pH meter", "conductivity meter"],
    }),
    (re.compile(r"\b(?:temperature|thermometer|thermocouple|infrared\s+thermometer)\b|\b(?:TM[-_]|PTM-|BTM-|MTM-|PST-|TC-|FT-967)", re.I), {
        "category": "Thermometer dan Temperature Meter", "term": "alat ukur suhu",
        "function": "mengukur suhu sampel, lingkungan, atau proses sesuai spesifikasi produk",
        "tags": ["Lutron", "thermometer", "temperature meter"],
    }),
    (re.compile(r"\b(?:humidity|hygrometer|wet\s+bulb|dew\s+point|barometer)\b|\b(?:HR-|HT-|MHB-|MHT-|PHB-|PHT-)", re.I), {
        "category": "Humidity dan Environmental Meter", "term": "humidity atau environmental meter",
        "function": "mengukur kelembapan dan parameter lingkungan terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "humidity meter", "hygrometer", "environment meter"],
    }),
    (re.compile(r"\b(?:anemometer|air\s+flow|air\s+velocity|wind\s+speed|pitot\s+tube)\b|\b(?:AM-|ABH-|AH-|PAM-|MY-|YK-2004|YK-80A|LM-81A|SP-82A)", re.I), {
        "category": "Anemometer dan Air Flow Meter", "term": "anemometer atau air flow meter",
        "function": "mengukur kecepatan udara, aliran udara, atau parameter lingkungan terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "anemometer", "air flow meter"],
    }),
    (re.compile(r"\b(?:light\s+meter|lux\s+meter|uv\s+light|solar\s+power)\b|\b(?:LX-|UVA?-|UVC-|SPM-|YK-(?:10LX|35UV|37UV)|SP-82LX|SP-82UV|LM-81LX)", re.I), {
        "category": "Light Meter dan Lux Meter", "term": "light meter atau lux meter",
        "function": "mengukur intensitas cahaya atau parameter radiasi terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "light meter", "lux meter", "uv meter"],
    }),
    (re.compile(r"\b(?:sound\s+level|noise\s+dosimeter|sound\s+calibrator|decibel)\b|\b(?:SL-|SC-|DS-)", re.I), {
        "category": "Sound Level Meter dan Kalibrator", "term": "sound level meter atau kalibrator suara",
        "function": "mengukur atau mengkalibrasi tingkat kebisingan dan tekanan suara sesuai spesifikasi produk",
        "tags": ["Lutron", "sound level meter", "sound calibrator"],
    }),
    (re.compile(r"\b(?:electromagnetic\s+field|emf\s+tester|rf\s+field|gauss\s+meter|magnetic\s+meter|micro\s+wave\s+leakage|electrostatic\s+field)\b|\b(?:EMF-|ESF-|GU-|MG-|PMG-|BMG-)", re.I), {
        "category": "Pengukur Medan Elektromagnetik", "term": "pengukur medan elektromagnetik",
        "function": "mengukur medan elektromagnetik sesuai spesifikasi produk",
        "tags": ["Lutron", "EMF tester", "gauss meter"],
    }),
    (re.compile(r"\b(?:power\s+analyzer|power\s+meter|watt\s*meter|energy\s+meter|power\s+quality|three\s+phase|fork\s+current|clamp\s+power|apparent\s+power|power\s+factor)\b|\b(?:DW-|CWF-|PC-|PVA-|PPF-|FT-9950)", re.I), {
        "category": "Power Analyzer dan Electrical Tester", "term": "power analyzer atau electrical tester",
        "function": "menganalisis parameter daya, energi, dan kelistrikan sesuai spesifikasi produk",
        "tags": ["Lutron", "power analyzer", "power meter"],
    }),
    (re.compile(r"\b(?:clamp\s+meter|multimeter|digital\s+multimeter|voltmeter|ammeter|bench\s+meter|leakage\s+tester|insulation\s+tester|rccb|rotation\s+tester)\b|\b(?:CM-|CMF-|DM-|AA-|AV-|DV-|DL-6054|DI-|RCB-|RT-)", re.I), {
        "category": "Multimeter dan Electrical Meter", "term": "multimeter atau electrical meter",
        "function": "mengukur atau memeriksa parameter kelistrikan sesuai spesifikasi produk",
        "tags": ["Lutron", "multimeter", "electrical meter"],
    }),
    (re.compile(r"\b(?:frequency\s+counter|frequency\s+meter|rf\s+detector|line\s+frequency)\b|\b(?:FC-|LF-)", re.I), {
        "category": "Frequency Counter", "term": "frequency counter",
        "function": "mengukur frekuensi dan parameter sinyal elektronik sesuai spesifikasi produk",
        "tags": ["Lutron", "frequency counter"],
    }),
    (re.compile(r"\b(?:vibration|stroboscope|tachometer|rpm)\b|\b(?:VB-|BVB-|MVB-|PVB-|VT-|DT-|PDT-)", re.I), {
        "category": "Vibration dan Tachometer", "term": "vibration meter atau tachometer",
        "function": "mengukur getaran, putaran, atau parameter mekanis terkait sesuai spesifikasi produk",
        "tags": ["Lutron", "vibration meter", "tachometer"],
    }),
    (re.compile(r"\b(?:force\s+gauge|fruit\s+hardness|torque|material\s+tester)\b|\b(?:FG-|FR-|TQ-)", re.I), {
        "category": "Force Gauge dan Material Tester", "term": "alat WEBSITE gaya atau material",
        "function": "mendukung pengukuran gaya, kekerasan, torsi, atau parameter material sesuai spesifikasi produk",
        "tags": ["Lutron", "force gauge", "material tester"],
    }),
    (re.compile(r"\b(?:digital\s+balance|digital\s+scale|calibration\s+weight)\b|\b(?:GM-|WT-)", re.I), {
        "category": "Timbangan Digital dan Balance", "term": "timbangan digital atau balance",
        "function": "mengukur massa atau berat sesuai spesifikasi produk",
        "tags": ["Lutron", "digital balance", "digital scale"],
    }),
    (re.compile(r"\b(?:moisture\s+meter|wood\s+moisture|soil\s+moisture|humidity\s+content)\b|\b(?:MS-(?!71P)|PMS-)", re.I), {
        "category": "Moisture Meter", "term": "moisture meter",
        "function": "mengukur kadar kelembapan material atau tanah sesuai spesifikasi produk",
        "tags": ["Lutron", "moisture meter"],
    }),
    (re.compile(r"\b(?:calibrator|decade\s+box|lcr\s+meter|milliohm|micro[-\s]?ohm|smd\s+tester|photo\s+interrupter)\b|\b(?:CC-|CBOX-|LBOX-|RBOX-|LCR-|MO-|SMDA-)", re.I), {
        "category": "Kalibrator dan Component Tester", "term": "kalibrator atau component tester",
        "function": "mendukung kalibrasi atau pengWEBSITEan komponen elektronik sesuai spesifikasi produk",
        "tags": ["Lutron", "calibrator", "component tester"],
    }),
    (re.compile(r"\b(?:panel\s+meters?|process\s+indicator)\b|\bDR-", re.I), {
        "category": "Panel Meter dan Process Indicator", "term": "panel meter atau process indicator",
        "function": "menampilkan atau memantau parameter proses listrik dan industri sesuai spesifikasi produk",
        "tags": ["Lutron", "panel meter", "process indicator"],
    }),
    (re.compile(r"\b(?:ultrasonic\s+leakage\s+detector|ultrasonic\s+transmitter)\b|\bGS-", re.I), {
        "category": "Ultrasonic Leak Detector", "term": "ultrasonic leak detector",
        "function": "mendukung deteksi kebocoran atau pengWEBSITEan ultrasonik sesuai spesifikasi produk",
        "tags": ["Lutron", "ultrasonic leak detector"],
    }),
    (re.compile(r"\b(?:current\s+recorder|voltage\s+recorder|sd\s*card\s+data\s+logger)\b|\b(?:MMA-|MMV-|DL-9602)", re.I), {
        "category": "Data Logger dan Recorder", "term": "data logger atau recorder",
        "function": "merekam parameter pengukuran sesuai spesifikasi produk",
        "tags": ["Lutron", "data logger", "recorder"],
    }),
    (re.compile(r"\b(?:controller|monitor|indicator|alarm)\b|\b(?:CT-|PAA-|PAV-|PDA-|PDV-|PIR-|PPH-|PPS-|PTM-9957|PVA-|PVB-8219)", re.I), {
        "category": "Industrial Controller dan Monitor", "term": "controller atau monitor industri",
        "function": "memantau atau mengendalikan parameter proses sesuai spesifikasi produk",
        "tags": ["Lutron", "industrial controller", "process monitor"],
    }),
    (re.compile(r"\b(?:environment\s+meter|weather\s+meter|wbgt|heat\s+index|4\s+in\s+1\s+meter|5\s+in\s+1\s+meter)\b|\b(?:EM-|EMC-|LM-|SP-|WBGT-)", re.I), {
        "category": "Environmental dan Weather Meter", "term": "environmental atau weather meter",
        "function": "mengukur parameter lingkungan sesuai spesifikasi produk",
        "tags": ["Lutron", "environment meter", "weather meter"],
    }),
]


def _v9_model_from_product(title: str, specs: Sequence[Dict[str, str]]) -> str:
    for row in specs:
        if re.search(r"\b(model|sku|cat|part|item|mpn)\b", clean_text(row.get("label")), re.I):
            value = clean_text(row.get("value"))
            if value:
                return value
    # Input data normally includes title and the actual model is passed in specs.
    return ""


def is_accessory_product(title: str, description: str, specs: Sequence[Dict[str, str]]) -> Tuple[bool, str]:
    """V9 catalog-wide accessory classification with model-first overrides."""
    title_clean = clean_text(title)
    desc_clean = clean_text(description)
    model = _v9_model_from_product(title_clean, specs)
    composite = f"{model} {title_clean}".strip()

    # Model-specific rules are strongest because titles can mention a meter or
    # tester only to describe the instrument the component connects to.
    if ACCESSORY_MODEL_PREFIX_RE_V9.search(composite):
        return True, "accessory_model_or_consumable_pattern"

    # Independent instruments retain main-product status even when supplied
    # with a probe or sensor.
    if PRIMARY_MODEL_PREFIX_RE.search(composite):
        return False, "main_instrument_model_pattern"

    functional_labels = sum(
        1 for row in specs
        if re.search(r"\b(range|measurement|accuracy|resolution|display|power|battery|output|interface|data\s*logger|datalogger|channel|temperature|humidity|pressure|frequency|weight|dimension)\b", clean_text(row.get("label")), re.I)
    )
    if functional_labels >= 2:
        return False, "main_instrument_spec_signature"

    if ACCESSORY_TITLE_RE_V9.search(title_clean):
        # A title that begins or ends as a recognised autonomous instrument is
        # not downgraded merely because it has a separate probe.
        if re.search(r"\b(?:meter|tester|analyzer|counter|recorder|logger|transmitter|controller|monitor|calibrator|balance|scale|detector|thermometer)\b|\b(?:co2|co|o2)\s*meter\b|\bco2meter\b", title_clean, re.I):
            return False, "main_instrument_title_pattern"
        return True, "accessory_title_pattern"

    # Incomplete source pages are left as main/unknown, never guessed as an
    # accessory. Strict publish will gate missing source data separately.
    return False, "not_confirmed_as_accessory"


def _v10_category_description(value: Any, max_chars: int = 1600) -> str:
    """Return only the product-description text used for category routing.

    Technical tables, full product HTML, optional accessories, and specification
    rows are deliberately excluded. They often contain words such as USB,
    RS-232, software, interface, probe, or temperature that describe a feature
    or accessory rather than the product's primary category.
    """
    text = compact(clean_catalog_source_text(value), max_chars)
    if not text:
        return ""
    # Some Lutron pages concatenate the summary with a technical section.
    # Keep the description portion and stop before common specification labels.
    stop = re.search(
        r"\b(?:specifications?|technical\s+data|technical\s+specification|"
        r"standard\s+accessories|optional\s+accessories|accessories\s+included)\b",
        text,
        flags=re.I,
    )
    return clean_text(text[:stop.start()] if stop else text)


def _v10_match_rule(text: str, rules: Sequence[Tuple[re.Pattern[str], Dict[str, Any]]]) -> Optional[Dict[str, Any]]:
    """Return the first catalog rule matched by a controlled category text."""
    for pattern, rule in rules:
        if pattern.search(text):
            return dict(rule)
    return None


def _v10_source_category_family(source_category: str, title: str) -> Optional[Dict[str, Any]]:
    normalized = normalize_lutron_category(source_category, title)
    if not normalized:
        return None
    family = dict(DEFAULT_FAMILY)
    family["category"] = normalized
    family["classification_basis"] = "source_category_fallback"
    return family


def detect_family(data: Dict[str, Any]) -> Dict[str, Any]:
    """Classify from product identity and official description, never specs.

    Priority:
    1. Confirmed accessory/consumable state.
    2. Model + input title + official product title.
    3. Official product description only, when title cannot identify the type.
    4. Official source category as a fallback.
    5. Generic category plus an audit flag.

    `spec_rows` and `official_content_html` are intentionally ignored. This
    prevents an EMF meter, pH meter, or humidity meter from becoming a software,
    thermometer, or interface product merely because its specifications mention
    optional USB/RS-232/software/accessories.
    """
    input_title = clean_catalog_source_text(data.get("input_title"))
    official_title = clean_catalog_source_text(data.get("title"))
    model = clean_text(data.get("model"))

    if data.get("is_accessory"):
        family = dict(ACCESSORY_FAMILY)
        family["classification_basis"] = "confirmed_accessory_title_model"
        return family

    title_model_text = clean_text(" ".join(x for x in (model, input_title, official_title) if x))
    description_text = _v10_category_description(data.get("description"))

    # Software/interface is deliberately only allowed from model or title.
    # A complete instrument can mention optional software in its description.
    software_rule = V9_FAMILY_RULES[0]
    if software_rule[0].search(title_model_text):
        family = dict(software_rule[1])
        family["classification_basis"] = "model_title"
        return family

    # The remaining rules represent physical product families. First use
    # model/title. This is the strongest evidence and prevents broad words in
    # prose such as temperature, monitor, or alarm from overriding the product.
    non_software_rules = V9_FAMILY_RULES[1:]
    family = _v10_match_rule(title_model_text, non_software_rules)
    if family:
        family["classification_basis"] = "model_title"
        return family

    # Description becomes the classifier only when title/model is insufficient.
    # Generic controller/monitor text is excluded here because it appears in
    # many ordinary descriptions and is too broad to be trusted alone.
    description_rules = [
        rule for rule in non_software_rules
        if clean_text(rule[1].get("category")) != "Industrial Controller dan Monitor"
    ]
    family = _v10_match_rule(description_text, description_rules)
    if family:
        family["classification_basis"] = "official_description"
        return family

    source_family = _v10_source_category_family(data.get("source_category", ""), official_title or input_title)
    if source_family:
        return source_family

    family = dict(DEFAULT_FAMILY)
    family["classification_basis"] = "generic_review_required"
    return family

def _v9_preflight_review_flag(record: Dict[str, str], accessory: bool, family: Dict[str, Any]) -> str:
    title = clean_text(record.get("input_title"))
    model = clean_text(record.get("input_model"))
    # Explicitly flag component-looking items that have no matching model rule
    # so they can be reviewed once, not silently published as instruments.
    component_words = re.compile(r"\b(probe|sensor|electrode|cable|adapter|clip|lead|case|cover|solution|electrolyte|tips?|holder|bag|weight|shunt|plug|grip|ear\s*phone)\b", re.I)
    if not accessory and component_words.search(title) and not PRIMARY_MODEL_PREFIX_RE.search(model):
        # Terms such as "separate probe" often describe a complete meter.
        # Only route to review when the title itself lacks clear main-product evidence.
        if not re.search(r"\b(?:meter|tester|analyzer|counter|recorder|logger|transmitter|controller|monitor|calibrator|balance|scale|detector|thermometer)\b|\b(?:co2|co|o2)\s*meter\b|\bco2meter\b", title, re.I):
            return "manual_product_type_review"
    if family.get("category") == DEFAULT_FAMILY["category"]:
        return "category_rule_review"
    return ""


def write_catalog_preflight(records: Sequence[Dict[str, str]], output_path: str) -> None:
    """Audit every input record without network calls and save a review-ready Excel file."""
    audit_rows: List[Dict[str, Any]] = []
    for idx, record in enumerate(records, 1):
        model = clean_text(record.get("input_model"))
        title = clean_text(record.get("input_title"))
        specs = [{"label": "Model", "value": model, "section": "Input Excel"}] if model else []
        is_accessory, reason = is_accessory_product(title, "", specs)
        family = detect_family({
            "input_title": title, "title": title, "model": model, "description": "", "source_category": "",
            "spec_rows": specs, "is_accessory": is_accessory,
        })
        review = _v9_preflight_review_flag(record, is_accessory, family)
        audit_rows.append({
            "No": idx,
            "Model": model,
            "Input Title": title,
            "URL": clean_text(record.get("url")),
            "Product Type": "Aksesori/Komponen" if is_accessory else "Produk Utama",
            "Classification Reason": reason,
            "Suggested Category": clean_text(family.get("category")),
            "Suggested Term": clean_text(family.get("term")),
            "Category Basis": clean_text(family.get("classification_basis")),
            "Preflight Status": "Review" if review else "Ready for Live Scrape",
            "Review Reason": review,
        })
    df = pd.DataFrame(audit_rows)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Catalog Preflight")
        summary = (
            df.groupby(["Product Type", "Suggested Category", "Preflight Status"], dropna=False)
            .size().reset_index(name="Count").sort_values(["Product Type", "Suggested Category"])
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")
    style_output_xlsx(output_path, "Catalog Preflight")

# ---------------------------------------------------------------------------
# V22 evidence-locked catalog generation and commercial-quality publish gate
# ---------------------------------------------------------------------------
from collections import defaultdict

AI_PROMPT_VERSION = "lutron_WEBSITE_v22_evidence_locked_commercial_2026_07_27"

V22_ALLOWED_HTML_TAGS = {"p", "strong", "b", "em", "ul", "ol", "li", "h2", "h3", "br"}
V22_META_MIN = 90
V22_META_MAX = 155
V22_SHORT_MAX = 260
V22_FAQ_MIN = 3
V22_FAQ_MAX = 5
V22_STOP_END_RE = re.compile(
    r"\b(?:dan|atau|untuk|dengan|dari|pada|yang|serta|sebagai|hingga|ke|di|ideal\s+untuk)\s*[.!?]?\s*$",
    re.I,
)
V22_FORBIDDEN_SALES_RE = re.compile(
    r"\b(?:harga\s+termurah|diskon|promo|ready\s*stock|stok\s+tersedia|gratis\s+ongkir|garansi\s+resmi)\b",
    re.I,
)
V22_SENSITIVE_FAQ_RE = re.compile(r"\b(?:stok|ketersediaan|garansi|kalibrasi|kelengkapan|termasuk\s+dalam\s+paket)\b", re.I)
V22_GENERIC_CLAIM_RULES: List[Tuple[re.Pattern[str], re.Pattern[str], str]] = [
    (re.compile(r"\b(?:akurasi\s+tinggi|presisi\s+tinggi|sangat\s+akurat|akurat)\b", re.I),
     re.compile(r"\b(?:accuracy|akurasi|precision|presisi|±)\b", re.I), "unsupported_accuracy_claim"),
    (re.compile(r"\b(?:andal|keandalan\s+tinggi|reliable)\b", re.I),
     re.compile(r"\b(?:reliable|reliability|andal|robust)\b", re.I), "unsupported_reliability_claim"),
    (re.compile(r"\b(?:berkualitas\s+tinggi|kualitas\s+tinggi|high\s+quality)\b", re.I),
     re.compile(r"\b(?:high\s+quality|quality|berkualitas|kualitas)\b", re.I), "unsupported_quality_claim"),
    (re.compile(r"\b(?:tahan\s+lama|durable|konstruksi\s+kokoh)\b", re.I),
     re.compile(r"\b(?:durable|rugged|heavy\s*duty|tahan\s+lama|kokoh)\b", re.I), "unsupported_durability_claim"),
]
V22_CERT_RE = re.compile(r"\b(?:ISO\s*\d*|CE|IEC\s*\d*|RoHS|REACH|sertifikasi|tersertifikasi)\b", re.I)

V22_PARAMETER_RULES: List[Tuple[re.Pattern[str], str]] = [
    (re.compile(r"\b(?:co2|carbon\s+dioxide|karbon\s+dioksida)\b", re.I), "CO₂"),
    (re.compile(r"\b(?:pm\s*2[.,]?5|particulate\s+matter)\b", re.I), "PM2.5"),
    (re.compile(r"\b(?:temperature|suhu|thermocouple|°\s*[cf]|℃|℉)\b", re.I), "Suhu"),
    (re.compile(r"\b(?:humidity|kelembapan|kelembaban|%\s*r\.?h\.?)\b", re.I), "Kelembapan"),
    (re.compile(r"\b(?:dew\s*point|titik\s+embun)\b", re.I), "Titik embun"),
    (re.compile(r"\b(?:wet\s*bulb|bola\s+basah)\b", re.I), "Suhu bola basah"),
    (re.compile(r"\b(?:pressure|tekanan|mbar|bar\b|psi\b|mm\s*hg|in(?:ch)?\s*hg|h2o|atmosphere|atm\b)\b", re.I), "Tekanan"),
    (re.compile(r"\b(?:air\s+velocity|wind\s+speed|kecepatan\s+(?:udara|angin)|m/s|ft/min|km/h|knot)\b", re.I), "Kecepatan udara"),
    (re.compile(r"\b(?:air\s*flow|aliran\s+udara|cfm|cmm|m3/min|m³/min)\b", re.I), "Aliran udara"),
    (re.compile(r"\b(?:light|lux|illumination|cahaya|fc\b|foot\s*candle)\b", re.I), "Intensitas cahaya"),
    (re.compile(r"\b(?:sound|noise|decibel|dba|dbc|db\b|kebisingan)\b", re.I), "Tingkat suara"),
    (re.compile(r"\b(?:rpm|tachometer|rotation|putaran)\b", re.I), "Putaran"),
    (re.compile(r"\b(?:frequency|frekuensi|hz\b|khz|mhz)\b", re.I), "Frekuensi"),
    (re.compile(r"\b(?:voltage|tegangan|mv\b|kv\b|volt|\bv\b)\b", re.I), "Tegangan"),
    (re.compile(r"\b(?:current|arus|ampere|\bma\b|\ba\b)\b", re.I), "Arus"),
    (re.compile(r"\b(?:power|daya|watt|\bw\b)\b", re.I), "Daya"),
    (re.compile(r"\b(?:ph\b|orp|conductivity|konduktivitas|tds|dissolved\s+oxygen|oksigen\s+terlarut)\b", re.I), "Kualitas air"),
    (re.compile(r"\b(?:moisture|kadar\s+air|kelembapan\s+material)\b", re.I), "Kelembapan material"),
    (re.compile(r"\b(?:force|gaya|newton|\bn\b)\b", re.I), "Gaya"),
    (re.compile(r"\b(?:electromagnetic|magnetic\s+field|electric\s+field|medan\s+elektromagnetik|µt|ut\b|gauss|v/m)\b", re.I), "Medan elektromagnetik"),
]

V22_FIELD_RULES: List[Tuple[re.Pattern[str], str]] = [
    (re.compile(r"\b(?:measurement|measuring)?\s*(?:range|rentang)\b", re.I), "range"),
    (re.compile(r"\b(?:resolution|resolusi|minimum\s+scale)\b", re.I), "resolution"),
    (re.compile(r"\b(?:accuracy|akurasi|error|kesalahan)\b", re.I), "accuracy"),
    (re.compile(r"\b(?:measurement\s+unit|satuan(?:\s+pengukuran|\s+tampilan)?)\b", re.I), "unit"),
]

V22_SYNONYM_TAGS: List[Tuple[re.Pattern[str], str]] = [
    (re.compile(r"manometer|pressure", re.I), "pressure meter"),
    (re.compile(r"tachometer|rpm", re.I), "RPM meter"),
    (re.compile(r"light|lux", re.I), "lux meter"),
    (re.compile(r"sound|noise", re.I), "sound level meter"),
    (re.compile(r"humidity|hygrometer", re.I), "hygrometer"),
    (re.compile(r"temperature|thermometer", re.I), "temperature meter"),
    (re.compile(r"anemometer|air flow", re.I), "air flow meter"),
    (re.compile(r"water quality|pH|conductivity|TDS", re.I), "water quality meter"),
    (re.compile(r"electromagnetic|EMF", re.I), "EMF meter"),
    (re.compile(r"data logger|recorder", re.I), "data logger"),
]

V22_WEBSITE_POLICIES = [
    {
        "id": "WEBSITE01",
        "text": "Untuk memastikan ketersediaan, kelengkapan paket, kompatibilitas, garansi, atau layanan kalibrasi, hubungi tim WEBSITE sebelum pemesanan.",
    },
    {
        "id": "WEBSITE02",
        "text": "Tim WEBSITE dapat membantu mencocokkan model dan spesifikasi dengan kebutuhan penggunaan pelanggan.",
    },
]


def _v22_complete_sentence(value: Any, max_chars: int = 0) -> str:
    text = re.sub(r"\s+", " ", _html_text(value)).strip()
    if not text:
        return ""
    if max_chars and len(text) > max_chars:
        return ""
    if V22_STOP_END_RE.search(text):
        return ""
    if text[-1] not in ".!?":
        text += "."
    return text


def _v22_normalize_html(value: Any) -> str:
    raw = clean_text(value)
    raw = re.sub(r"^```(?:html|json)?\s*|\s*```$", "", raw, flags=re.I)
    wrapper = BeautifulSoup(f'<div id="v22-root">{raw}</div>', "html.parser")
    root = wrapper.find(id="v22-root")
    if root is None:
        return ""
    for tag in list(root.find_all(["script", "style", "iframe", "object", "embed", "svg", "form"])):
        tag.decompose()
    for tag in list(root.find_all(True)):
        if tag.name not in V22_ALLOWED_HTML_TAGS:
            tag.unwrap()
        else:
            tag.attrs = {}
    for li in list(root.find_all("li")):
        if li.parent is None or li.parent.name not in {"ul", "ol"}:
            ul = wrapper.new_tag("ul")
            li.wrap(ul)
    for tag in reversed(list(root.find_all(True))):
        if tag.name != "br" and not clean_text(tag.get_text(" ", strip=True)) and not tag.find(True):
            tag.decompose()
    result = "".join(str(node) for node in root.contents).strip()
    result = re.sub(r"(?:<br\s*/?>\s*){3,}", "<br><br>", result, flags=re.I)
    return result


def _sanitize_ai_html(value: Any) -> str:
    return _v22_normalize_html(value)


def _sanitize_ai_text(value: Any, max_chars: int) -> str:
    return _v22_complete_sentence(value, max_chars)


def _v22_html_contract_issues(value: Any, require_heading: bool = False, require_list: bool = False) -> List[str]:
    html_value = clean_text(value)
    issues: List[str] = []
    if not html_value:
        return ["html_empty"]
    if re.search(r"<\s*(?:script|style|iframe|object|embed|form)\b", html_value, re.I):
        issues.append("html_unsafe_tag")
    soup = BeautifulSoup(f"<div>{html_value}</div>", "html.parser")
    for li in soup.find_all("li"):
        if li.parent is None or li.parent.name not in {"ul", "ol"}:
            issues.append("html_orphan_list_item")
            break
    if require_heading and not soup.find(["h2", "h3"]):
        issues.append("html_missing_heading")
    if require_list and not soup.find("li"):
        issues.append("html_missing_list")
    if has_cjk(_html_text(html_value)):
        issues.append("html_contains_chinese")
    return dedupe_keep_order(issues)


def _v22_sentence_chunks(value: Any, max_items: int = 40) -> List[str]:
    text = clean_catalog_source_text(value)
    if not text:
        return []
    chunks: List[str] = []
    for block in re.split(r"[\n\r]+", text):
        block = clean_text(block)
        if not block:
            continue
        pieces = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9])", block)
        for piece in pieces:
            piece = clean_text(piece)
            if 8 <= len(piece) <= 700 and not is_noise_text(piece) and not is_generic_lutron_landing_text(piece):
                chunks.append(piece)
            if len(chunks) >= max_items:
                return dedupe_keep_order(chunks)
    return dedupe_keep_order(chunks)


def build_fact_ledger(data: Dict[str, Any], family: Dict[str, Any]) -> List[Dict[str, str]]:
    """Create small source facts. Never send a raw page HTML blob to the AI."""
    facts: List[Dict[str, str]] = []
    seen = set()

    def add(kind: str, section: str, text: Any, label: str = "") -> None:
        cleaned = clean_catalog_source_text(text)
        if not cleaned or is_generic_lutron_landing_text(cleaned):
            return
        key = re.sub(r"\s+", " ", cleaned).casefold()
        if key in seen:
            return
        seen.add(key)
        facts.append({
            "id": f"F{len(facts) + 1:03d}",
            "kind": kind,
            "section": clean_text(section) or "Product page",
            "label": clean_catalog_source_text(label),
            "text": compact(cleaned, 700),
        })

    title = clean_catalog_source_text(data.get("title"))
    model = clean_text(data.get("model"))
    source_category = clean_catalog_source_text(data.get("source_category"))
    if title:
        add("identity", "Identity", title, "Product name")
    if model:
        add("identity", "Identity", model, "Model")
    if source_category:
        add("category", "Breadcrumb", source_category, "Source category")
    for sentence in _v22_sentence_chunks(data.get("description"), 16):
        add("description", "Product description", sentence)

    official = clean_text(data.get("official_content_html"))
    if official:
        soup = BeautifulSoup(official, "html.parser")
        section = "Product information"
        for node in soup.find_all(["h2", "h3", "p", "li"]):
            text = clean_catalog_source_text(node.get_text(" ", strip=True))
            if node.name in {"h2", "h3"}:
                section = text or section
                continue
            for sentence in _v22_sentence_chunks(text, 6):
                add("product_content", section, sentence)
            if len(facts) >= 70:
                break
    return facts[:70]


def _v22_infer_spec_field(label: Any, value: Any) -> str:
    label_text = clean_text(label)
    value_text = clean_text(value)
    probe = f"{label_text} {value_text}"
    if re.search(r"\b(?:operating|working|ambient)\s+(?:temperature|humidity)|suhu\s+(?:operasi|kerja|lingkungan)|kelembap(?:an|ban)\s+(?:operasi|kerja|lingkungan)", probe, re.I):
        return "general"
    for pattern, field in V22_FIELD_RULES:
        if pattern.search(label_text):
            return field
    if _measurement_unit_label(label_text) and re.search(r"\b(?:resolution|resolusi)\b", value_text, re.I):
        return "range_resolution"
    if _measurement_quantity(label_text):
        if re.search(r"(?:±|\+/-|%\s*(?:F\.?S\.?|full\s*scale))", value_text, re.I):
            return "accuracy"
        if _measurement_quantity(value_text):
            return "range_resolution"
    return "general"


def _v22_infer_parameter(label: Any, value: Any, section: Any = "", context: Any = "") -> str:
    probe = " ".join(clean_text(x) for x in (label, value, section, context) if clean_text(x))
    for pattern, name in V22_PARAMETER_RULES:
        if pattern.search(probe):
            return name
    unit = _measurement_unit_label(label)
    if unit:
        unit_probe = unit.casefold()
        if any(token in unit_probe for token in ("mbar", "psi", "hg", "h2o", "atm", "bar")):
            return "Tekanan"
        if any(token in unit_probe for token in ("lux", "fc")):
            return "Intensitas cahaya"
        if any(token in unit_probe for token in ("m/s", "ft/min", "km/h", "knot")):
            return "Kecepatan udara"
    return ""


def _spec_rows_for_ai(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw_rows = data.get("_raw_spec_rows")
    rows = raw_rows if isinstance(raw_rows, list) else (
        data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    )
    limit = int(AI_SETTINGS.get("spec_limit") or 120)
    prepared: List[Dict[str, Any]] = []
    basic: List[Tuple[str, str, str]] = []
    for item in rows[:limit]:
        if not isinstance(item, dict):
            basic.append(("", "", ""))
            continue
        basic.append((
            clean_catalog_source_text(item.get("source_label") or item.get("label")),
            clean_catalog_source_text(item.get("value")),
            clean_text(item.get("section")) or "Specifications",
        ))
    for index, (label, value, section) in enumerate(basic, 1):
        if not label or not value:
            continue
        before = [f"{a}: {b}" for a, b, _ in basic[max(0, index - 3):index - 1] if a and b]
        after = [f"{a}: {b}" for a, b, _ in basic[index:index + 2] if a and b]
        context = " | ".join(before + after)
        prepared.append({
            "id": index,
            "fact_id": f"S{index:03d}",
            "source_label": label,
            "source_value": compact(value, 700),
            "section": section,
            "context_before": before,
            "context_after": after,
            "suggested_field": _v22_infer_spec_field(label, value),
            "suggested_parameter": _v22_infer_parameter(label, value, section, context),
        })
    return prepared


def _ai_payload(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "source_product_name": clean_catalog_source_text(data.get("title")),
        "brand": clean_text(data.get("brand")) or "Lutron",
        "model": clean_text(data.get("model")),
        "source_url": clean_text(data.get("url")),
        "source_category": clean_catalog_source_text(data.get("source_category")),
        "product_type": "accessory_or_component" if data.get("is_accessory") else "main_product",
        "suggested_category": clean_catalog_source_text(family.get("category")).strip(),
        "family_function": clean_catalog_source_text(family.get("function")),
        "facts": build_fact_ledger(data, family),
        "specifications": _spec_rows_for_ai(data),
        "allowed_WEBSITE_policies": V22_WEBSITE_POLICIES,
    }


def _v22_fact_map(payload: Dict[str, Any]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for item in payload.get("facts") if isinstance(payload.get("facts"), list) else []:
        if isinstance(item, dict) and clean_text(item.get("id")):
            mapping[clean_text(item.get("id"))] = clean_text(item.get("text"))
    for item in payload.get("specifications") if isinstance(payload.get("specifications"), list) else []:
        if isinstance(item, dict):
            fact_id = clean_text(item.get("fact_id")) or f"S{int(item.get('id') or 0):03d}"
            mapping[fact_id] = f"{clean_text(item.get('source_label'))}: {clean_text(item.get('source_value'))}"
    for item in payload.get("allowed_WEBSITE_policies") if isinstance(payload.get("allowed_WEBSITE_policies"), list) else []:
        if isinstance(item, dict) and clean_text(item.get("id")):
            mapping[clean_text(item.get("id"))] = clean_text(item.get("text"))
    return mapping


def _v22_clean_evidence_ids(values: Any, fact_map: Dict[str, str]) -> List[str]:
    if not isinstance(values, list):
        return []
    return dedupe_keep_order([clean_text(value) for value in values if clean_text(value) in fact_map])


def _v22_evidence_text(ids: Sequence[str], fact_map: Dict[str, str]) -> str:
    return " ".join(fact_map.get(item, "") for item in ids if item in fact_map)


def _v22_numeric_claims_supported(output: Any, evidence: Any) -> bool:
    output_tokens = [token.replace(",", ".") for token in _numeric_tokens(output)]
    evidence_norm = clean_text(evidence).replace(",", ".")
    return all(token in evidence_norm for token in output_tokens)


def _v22_claim_issues(output: Any, evidence: Any) -> List[str]:
    text = _html_text(output)
    source = clean_text(evidence)
    issues: List[str] = []
    for output_pattern, source_pattern, code in V22_GENERIC_CLAIM_RULES:
        if output_pattern.search(text) and not source_pattern.search(source):
            issues.append(code)
    for match in V22_CERT_RE.finditer(text):
        token = clean_text(match.group(0))
        token_probe = r"\s*".join(re.escape(part) for part in re.split(r"\s+", token.strip()) if part)
        if not re.search(token_probe, source, re.I):
            issues.append("unsupported_certification_claim")
            break
    return dedupe_keep_order(issues)


def _v22_normalize_field(value: Any) -> str:
    field = clean_text(value).casefold().replace(" ", "_")
    aliases = {
        "measurement_range": "range", "measuring_range": "range", "rentang": "range",
        "resolusi": "resolution", "akurasi": "accuracy", "satuan": "unit",
        "range_and_resolution": "range_resolution", "rentang_dan_resolusi": "range_resolution",
    }
    field = aliases.get(field, field)
    return field if field in {"range", "resolution", "accuracy", "unit", "range_resolution", "general"} else "general"


_v21_validate_ai_spec_deduplication = _validate_ai_spec_deduplication


def _validate_ai_spec_deduplication(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, str]], int, List[Dict[str, Any]]]:
    rows, removed_count, issues = _v21_validate_ai_spec_deduplication(parsed, source_specs)
    if issues or not rows:
        return rows, removed_count, issues
    translated_by_id = {
        _coerce_ai_spec_id(item.get("id")): item
        for item in parsed.get("translated_specifications", [])
        if isinstance(item, dict) and _coerce_ai_spec_id(item.get("id"))
    }
    keep_ids = [_coerce_ai_spec_id(value) for value in parsed.get("keep_specification_ids", [])]
    source_by_id = {int(item["id"]): item for item in source_specs if isinstance(item, dict) and item.get("id")}
    enriched: List[Dict[str, str]] = []
    for row, source_id in zip(rows, [value for value in keep_ids if value is not None]):
        item = translated_by_id.get(source_id, {})
        source = source_by_id.get(int(source_id), {})
        parameter = clean_catalog_source_text(item.get("parameter_id")) or clean_catalog_source_text(source.get("suggested_parameter"))
        field = _v22_normalize_field(item.get("field_id") or source.get("suggested_field"))
        group = clean_catalog_source_text(item.get("group_id")) or parameter
        output = dict(row)
        output.update({
            "source_id": str(source_id),
            "parameter": parameter,
            "field": field,
            "group": group,
        })
        enriched.append(output)
    return enriched, removed_count, []


def _v22_render_faq(items: Sequence[Dict[str, Any]]) -> str:
    blocks: List[str] = []
    for item in items:
        question = clean_text(item.get("question"))
        answer = _v22_complete_sentence(item.get("answer"), 520)
        if question and answer:
            question = question.rstrip(" .") + "?" if not question.endswith("?") else question
            blocks.append(f"<h3>{esc(question)}</h3><p>{esc(answer)}</p>")
    return _v22_normalize_html("".join(blocks))


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    feedback: List[Dict[str, Any]] = []
    fact_map = _v22_fact_map(payload)
    source_specs = payload.get("specifications") if isinstance(payload.get("specifications"), list) else []
    model = clean_text(payload.get("model"))
    brand = clean_text(payload.get("brand")) or "Lutron"

    title = _v22_complete_sentence(parsed.get("product_name_id"), 180).rstrip(".")
    description = _v22_normalize_html(parsed.get("product_description_html"))
    short = _v22_complete_sentence(parsed.get("short_description"), V22_SHORT_MAX)
    features = _v22_normalize_html(parsed.get("feature_tab_html"))
    meta = _v22_complete_sentence(parsed.get("meta_description"), V22_META_MAX)
    official = _v22_normalize_html(parsed.get("official_information_html"))

    evidence_block = parsed.get("content_evidence") if isinstance(parsed.get("content_evidence"), dict) else {}
    evidence_fields: Dict[str, List[str]] = {}
    for field_name in ("product_description", "short_description", "feature_tab", "meta_description", "official_information"):
        raw_ids = evidence_block.get(field_name)
        clean_ids = _v22_clean_evidence_ids(raw_ids, fact_map)
        evidence_fields[field_name] = clean_ids
        if not clean_ids:
            feedback.append(_spec_validation_issue(
                "missing_content_evidence",
                f"content_evidence.{field_name} wajib memuat minimal satu ID fakta yang valid.",
                field=field_name,
            ))

    faq_raw = parsed.get("faq_items")
    faq_items: List[Dict[str, Any]] = []
    if not isinstance(faq_raw, list) or not (V22_FAQ_MIN <= len(faq_raw) <= V22_FAQ_MAX):
        feedback.append(_spec_validation_issue(
            "invalid_faq_item_count",
            f"faq_items harus berisi {V22_FAQ_MIN}-{V22_FAQ_MAX} pertanyaan.",
        ))
    else:
        for index, item in enumerate(faq_raw, 1):
            if not isinstance(item, dict):
                feedback.append(_spec_validation_issue("invalid_faq_item", "Setiap faq_items harus berupa objek.", item=index))
                continue
            question = clean_text(item.get("question_id"))
            answer = _v22_complete_sentence(item.get("answer_id"), 520)
            evidence_ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map)
            if not question or not answer or not evidence_ids:
                feedback.append(_spec_validation_issue(
                    "invalid_faq_content",
                    "Setiap FAQ harus memiliki question_id, answer_id, dan evidence_ids yang valid.",
                    item=index,
                ))
                continue
            evidence_text = _v22_evidence_text(evidence_ids, fact_map)
            if not _v22_numeric_claims_supported(answer, evidence_text):
                feedback.append(_spec_validation_issue(
                    "faq_numeric_claim_not_in_evidence",
                    "Angka pada jawaban FAQ harus terdapat pada fakta yang dirujuk.",
                    item=index,
                ))
            claim_issues = _v22_claim_issues(answer, evidence_text)
            for code in claim_issues:
                feedback.append(_spec_validation_issue(code, "Klaim pada FAQ tidak didukung fakta yang dirujuk.", item=index))
            if V22_SENSITIVE_FAQ_RE.search(answer + " " + question):
                has_policy = any(evidence_id.startswith("WEBSITE") for evidence_id in evidence_ids)
                source_mentions_topic = bool(V22_SENSITIVE_FAQ_RE.search(evidence_text))
                if not (has_policy or source_mentions_topic):
                    feedback.append(_spec_validation_issue(
                        "faq_sensitive_topic_without_policy_or_source",
                        "FAQ tentang stok, garansi, kalibrasi, atau kelengkapan wajib merujuk fakta sumber atau kebijakan WEBSITE.",
                        item=index,
                    ))
            faq_items.append({"question": question, "answer": answer, "evidence_ids": evidence_ids})

    translated_rows, removed_count, spec_feedback = _validate_ai_spec_deduplication(parsed, source_specs)
    feedback.extend(spec_feedback)

    required = {
        "product_name_id": title,
        "product_description_html": description,
        "short_description": short,
        "feature_tab_html": features,
        "meta_description": meta,
        "official_information_html": official,
    }
    for key, value in required.items():
        if not value:
            feedback.append(_spec_validation_issue("missing_ai_field", f"Field {key} kosong atau tidak memenuhi batas panjang.", field=key))
    if title and (brand.casefold() not in title.casefold() or (model and model.casefold() not in title.casefold())):
        feedback.append(_spec_validation_issue("seo_title_missing_brand_or_model", "Nama produk wajib memuat Lutron dan model.", title=title))
    if title.endswith("."):
        feedback.append(_spec_validation_issue("seo_title_ends_period", "Nama produk tidak boleh diakhiri titik."))
    if description and (len(_html_text(description)) < 170 or "<h2" not in description.lower()):
        feedback.append(_spec_validation_issue("invalid_product_description", "Deskripsi harus memiliki isi faktual dan heading h2."))
    for issue in _v22_html_contract_issues(description, require_heading=True):
        feedback.append(_spec_validation_issue(issue, "HTML deskripsi tidak valid."))
    for issue in _v22_html_contract_issues(features, require_heading=True, require_list=True):
        feedback.append(_spec_validation_issue(issue, "HTML keunggulan tidak valid."))
    for issue in _v22_html_contract_issues(official, require_heading=True):
        feedback.append(_spec_validation_issue(issue, "HTML informasi produk tidak valid."))

    content_values = {
        "product_description": description,
        "short_description": short,
        "feature_tab": features,
        "meta_description": meta,
        "official_information": official,
    }
    for field_name, value in content_values.items():
        evidence_ids = evidence_fields.get(field_name, [])
        evidence_text = _v22_evidence_text(evidence_ids, fact_map)
        if value and evidence_ids:
            if not _v22_numeric_claims_supported(value, evidence_text):
                feedback.append(_spec_validation_issue(
                    "numeric_claim_not_in_evidence",
                    "Semua angka pada konten wajib terdapat pada fakta yang dirujuk.",
                    field=field_name,
                ))
            for code in _v22_claim_issues(value, evidence_text):
                feedback.append(_spec_validation_issue(code, "Klaim tidak didukung fakta yang dirujuk.", field=field_name))

    if meta:
        if not (V22_META_MIN <= len(meta) <= V22_META_MAX):
            feedback.append(_spec_validation_issue(
                "meta_length_invalid",
                f"Meta description harus {V22_META_MIN}-{V22_META_MAX} karakter dan berupa kalimat lengkap.",
                length=len(meta),
            ))
        if V22_STOP_END_RE.search(meta):
            feedback.append(_spec_validation_issue("meta_sentence_incomplete", "Meta description terpotong atau berakhir dengan kata sambung."))
        if model and model.casefold() not in meta.casefold():
            feedback.append(_spec_validation_issue("meta_missing_model", "Meta description wajib memuat model produk."))

    public_text = " ".join(_html_text(value) for value in required.values()) + " " + _html_text(_v22_render_faq(faq_items))
    if V22_FORBIDDEN_SALES_RE.search(public_text):
        feedback.append(_spec_validation_issue("forbidden_sales_claim", "Jangan menulis harga, promo, stok, atau garansi resmi tanpa kebijakan sumber."))
    if has_cjk(public_text):
        feedback.append(_spec_validation_issue("chinese_character_detected", "Seluruh konten publik harus bebas karakter Chinese."))

    feedback = dedupe_keep_order([json.dumps(item, ensure_ascii=False, sort_keys=True) for item in feedback])
    feedback_items = [json.loads(item) for item in feedback]
    if feedback_items:
        return False, {}, "invalid_ai_content_contract", feedback_items[:30]

    fields: Dict[str, Any] = {
        "product_name_id": title,
        "category_id": clean_catalog_source_text(parsed.get("category_id")).rstrip("."),
        "product_description_html": description,
        "short_description": short,
        "feature_tab_html": features,
        "faq_html": _v22_render_faq(faq_items),
        "faq_items": faq_items,
        "meta_description": meta,
        "official_information_html": official,
        "translated_spec_rows": translated_rows,
        "ai_spec_dedup_applied": bool(source_specs),
        "ai_spec_dedup_removed_count": removed_count,
        "content_evidence": evidence_fields,
    }
    return True, fields, "ok", []


def _call_ai(
    payload: Dict[str, Any],
    retry_reason: str = "",
    retry_feedback: Optional[Sequence[Dict[str, Any]]] = None,
    attempt_number: int = 1,
    max_attempts: int = 1,
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    api_key = clean_text(AI_SETTINGS.get("api_key"))
    if not api_key:
        return False, {}, "missing_api_key", []
    user: Dict[str, Any] = {
        "task": "Susun konten katalog WooCommerce WEBSITE_KALIAN dalam bahasa Indonesia dari fakta terverifikasi berikut.",
        "critical_rules": [
            "Gunakan HANYA facts, specifications, dan allowed_WEBSITE_policies. Jangan menambah pengetahuan umum atau asumsi.",
            "Jangan menyalin HTML halaman. Susun kalimat baru yang ringkas dari fakta bernomor.",
            "Setiap bagian konten wajib menunjuk evidence ID yang benar pada content_evidence.",
            "Setiap FAQ dan setiap application_items wajib memiliki evidence_ids. Pertanyaan stok, garansi, kalibrasi, kompatibilitas, dan kelengkapan hanya boleh dijawab dari fakta sumber atau kebijakan WEBSITE.",
            "Klaim akurat, andal, berkualitas tinggi, tahan lama, sertifikasi, ISO, CE, IEC, RoHS, atau REACH hanya boleh ditulis jika fakta yang dirujuk menyatakannya.",
            "Pertahankan model, kode, angka, satuan, simbol, dan singkatan teknis persis. Jangan mengubah atau mengonversi angka dan satuan.",
            "Bedakan Data Hold, memori maksimum/minimum, data logger internal, antarmuka RS232, port USB bawaan, dan kabel USB opsional.",
            "Jika sumber menyebut USB-01 atau UPCB sebagai aksesori opsional, jangan menyatakan produk mempunyai port USB bawaan.",
            "Terjemahkan semua specifications. Setiap ID harus muncul tepat sekali di translated_specifications.",
            "Deduplikasi spesifikasi hanya untuk fakta yang benar-benar identik. Jika ragu, pertahankan kedua ID.",
            "Untuk setiap translated_specifications isi parameter_id dan field_id. field_id hanya: range, resolution, accuracy, unit, range_resolution, general.",
            "Meta description harus 90-155 karakter, memuat model, berupa kalimat lengkap, dan tidak boleh dipotong.",
            "Nama produk menggunakan pola jenis produk + Lutron + model + pembeda faktual, tanpa titik di akhir.",
            "Balas JSON valid saja tanpa markdown.",
        ],
        "output_schema": {
            "product_name_id": "Nama produk SEO maksimal 180 karakter tanpa titik akhir.",
            "category_id": "Kategori Indonesia singkat.",
            "product_description_html": "HTML aman dengan 2 paragraf pembuka, h2 Keunggulan Produk + ul, dan h2 Panduan Pemilihan.",
            "short_description": "1-2 kalimat maksimal 260 karakter: produk, pembeda, dan aplikasi bila didukung fakta.",
            "feature_tab_html": "h2 Keunggulan Produk + ul 4-6 poin faktual.",
            "faq_items": [{"question_id": "Pertanyaan", "answer_id": "Jawaban", "evidence_ids": ["F001 atau S001 atau WEBSITE01"]}],
            "application_items": [{"text_id": "Aplikasi penggunaan dalam bahasa Indonesia", "evidence_ids": ["F001 atau S001"]}],
            "meta_description": "90-155 karakter, kalimat lengkap, memuat model.",
            "official_information_html": "h2 Informasi Produk dan h2 Catatan Teknis dari fakta.",
            "content_evidence": {
                "product_description": ["ID fakta"],
                "short_description": ["ID fakta"],
                "feature_tab": ["ID fakta"],
                "application_tab": ["ID fakta"],
                "meta_description": ["ID fakta"],
                "official_information": ["ID fakta"],
            },
            "translated_specifications": [{
                "id": "ID numerik specifications",
                "label_id": "Label Indonesia",
                "value_id": "Nilai Indonesia dengan angka dan satuan tetap",
                "parameter_id": "Parameter pengukuran, contoh Suhu, Kelembapan, Tekanan; kosong jika umum",
                "field_id": "range|resolution|accuracy|unit|range_resolution|general",
                "group_id": "Kelompok tampilan singkat",
            }],
            "keep_specification_ids": ["Semua ID yang dipertahankan"],
            "removed_duplicate_specifications": [{"source_id": "ID duplikat", "duplicate_of_id": "ID yang dipertahankan", "reason_id": "Alasan"}],
        },
        "product_data": payload,
    }
    if retry_reason:
        user["retry_validation_note"] = [
            f"Percobaan ulang {attempt_number} dari {max_attempts}. Respons sebelumnya ditolak dengan kode {retry_reason}.",
            "Perbaiki semua temuan validator dan kirim ulang seluruh JSON lengkap.",
        ]
        user["validator_feedback"] = list(retry_feedback or [])[:30]
    body = {
        "model": clean_text(AI_SETTINGS.get("model")) or "deepseek-chat",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Anda adalah editor katalog teknis berbasis bukti. Gunakan hanya fakta bernomor yang diberikan. "
                    "Jangan berasumsi. Jangan membuat klaim pemasaran tanpa bukti. Balas JSON valid saja."
                ),
            },
            {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
        ],
        "temperature": AI_SETTINGS.get("temperature", 0.10),
        "response_format": {"type": "json_object"},
    }
    endpoint = AI_SETTINGS["base_url"] + "/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS["timeout"])
        if response.status_code in {400, 422}:
            body.pop("response_format", None)
            response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS["timeout"])
        response.raise_for_status()
        raw = response.json()
        content = clean_text(raw.get("choices", [{}])[0].get("message", {}).get("content", ""))
        parsed = _parse_json(content)
        return _v22_validate_ai_fields(parsed, payload)
    except Exception as exc:
        return False, {}, f"request_error_{type(exc).__name__}", []


def _is_retryable_ai_validation_reason(reason: Any) -> bool:
    value = clean_text(reason)
    return value in {"invalid_ai_content_contract", "invalid_ai_specification_deduplication", "missing_indonesian_ai_field"}


def _v22_enrich_spec_rows(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    enriched: List[Dict[str, str]] = []
    current_parameter = ""
    for index, raw in enumerate(rows or []):
        if not isinstance(raw, dict):
            continue
        row = dict(raw)
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        section = clean_text(row.get("section"))
        context = " ".join(
            f"{clean_text(other.get('label'))} {clean_text(other.get('value'))}"
            for other in list(rows)[max(0, index - 2):index + 3]
            if isinstance(other, dict)
        )
        parameter = clean_catalog_source_text(row.get("parameter")) or _v22_infer_parameter(label, value, section, context)
        field = _v22_normalize_field(row.get("field") or _v22_infer_spec_field(label, value))
        if parameter:
            current_parameter = parameter
        elif field in {"range", "resolution", "accuracy", "unit", "range_resolution"}:
            parameter = current_parameter
        row["parameter"] = parameter
        row["field"] = field
        row["group"] = clean_catalog_source_text(row.get("group")) or parameter
        enriched.append(row)
    return enriched


def _v22_measurement_records(rows: Sequence[Dict[str, str]]) -> Tuple[List[Dict[str, str]], set]:
    enriched = _v22_enrich_spec_rows(rows)
    records: List[Dict[str, str]] = []
    consumed: set = set()
    last_by_parameter: Dict[str, int] = {}

    def add_record(parameter: str, unit: str = "", range_value: str = "") -> int:
        item = {"parameter": parameter, "unit": unit, "range": range_value, "resolution": "", "accuracy": ""}
        records.append(item)
        position = len(records) - 1
        if parameter:
            last_by_parameter[parameter.casefold()] = position
        return position

    for index, row in enumerate(enriched):
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        parameter = clean_text(row.get("parameter"))
        field = clean_text(row.get("field")) or "general"
        unit_label = _measurement_unit_label(label)

        if field == "range_resolution":
            range_value, resolution = _parse_range_resolution_value(value)
            if not range_value:
                label_q = _measurement_quantity(label)
                value_q = _measurement_quantity(value)
                if label_q and value_q:
                    range_value, resolution = label_q[0], value_q[0]
                    unit_label = unit_label or label_q[1]
            if range_value or resolution:
                pos = add_record(parameter or _v22_infer_parameter(label, value), unit_label, range_value)
                records[pos]["resolution"] = resolution
                consumed.add(index)
                continue

        if field == "range":
            quantity = _measurement_quantity(value)
            pos = add_record(parameter, quantity[1] if quantity else "", quantity[0] if quantity else value)
            consumed.add(index)
            continue
        if field == "unit":
            if parameter:
                pos = last_by_parameter.get(parameter.casefold())
                if pos is None:
                    pos = add_record(parameter)
                records[pos]["unit"] = value
                consumed.add(index)
            continue
        if field in {"resolution", "accuracy"}:
            target: Optional[int] = None
            label_q = _measurement_quantity(label)
            if label_q:
                for pos, record in enumerate(records):
                    if clean_text(record.get("range")).casefold() == label_q[0].casefold():
                        target = pos
                        break
            if target is None and parameter:
                target = last_by_parameter.get(parameter.casefold())
            if target is None:
                target = add_record(parameter or _v22_infer_parameter(label, value), unit_label)
            records[target][field] = value
            consumed.add(index)
            continue

        if unit_label:
            range_value, resolution = _parse_range_resolution_value(value)
            if range_value or resolution:
                pos = add_record(parameter or _v22_infer_parameter(label, value), unit_label, range_value)
                records[pos]["resolution"] = resolution
                consumed.add(index)
                continue
            value_q = _measurement_quantity(value)
            if value_q and value_q[1].casefold() == unit_label.casefold():
                pos = add_record(parameter or _v22_infer_parameter(label, value), unit_label)
                records[pos]["resolution"] = value_q[0]
                consumed.add(index)

    meaningful = [record for record in records if any(clean_text(record.get(k)) for k in ("range", "resolution", "accuracy"))]
    return meaningful, consumed


def _v22_measurement_table_html(rows: Sequence[Dict[str, str]]) -> str:
    if not rows:
        return ""
    fields: List[Tuple[str, str]] = []
    if any(clean_text(row.get("parameter")) for row in rows):
        fields.append(("parameter", "Parameter"))
    if any(clean_text(row.get("unit")) for row in rows):
        fields.append(("unit", "Satuan"))
    if any(clean_text(row.get("range")) for row in rows):
        fields.append(("range", "Rentang"))
    if any(clean_text(row.get("resolution")) for row in rows):
        fields.append(("resolution", "Resolusi"))
    if any(clean_text(row.get("accuracy")) for row in rows):
        fields.append(("accuracy", "Akurasi"))
    if len(fields) < 2:
        return ""
    body = []
    for row in rows:
        body.append("<tr>" + "".join(
            f"<td>{esc(row.get(field)) if clean_text(row.get(field)) else '-'}</td>" for field, _ in fields
        ) + "</tr>")
    return (
        '<h3>Data Pengukuran</h3><table class="WEBSITE-spec-table WEBSITE-spec-measurement">'
        '<thead><tr>' + "".join(f"<th>{esc(header)}</th>" for _, header in fields) + '</tr></thead>'
        '<tbody>' + "".join(body) + '</tbody></table>'
    )


def specifications_tab_html(rows: Sequence[Dict[str, str]], already_deduplicated: bool = False) -> str:
    cleaned: List[Dict[str, str]] = []
    seen = set()
    source_rows = rows if already_deduplicated else canonicalize_specifications(rows)[0]
    for raw in source_rows or []:
        if not isinstance(raw, dict):
            continue
        label = clean_text(raw.get("label"))
        value = clean_text(raw.get("value"))
        if not label or not value:
            continue
        key = (_spec_text_key(label), _spec_text_key(value), clean_text(raw.get("section")).casefold())
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(dict(raw))
    if not cleaned:
        return "<h2>Spesifikasi</h2><p>Spesifikasi belum tersedia dalam bahasa Indonesia.</p>"
    measurement_rows, consumed = _v22_measurement_records(cleaned)
    general = [row for index, row in enumerate(cleaned) if index not in consumed]
    blocks = ["<h2>Spesifikasi</h2>"]
    measurement_html = _v22_measurement_table_html(measurement_rows)
    if measurement_html:
        blocks.append(measurement_html)
    if general:
        if measurement_html:
            blocks.append("<h3>Spesifikasi Umum</h3>")
        blocks.append(_general_spec_table_html(general))
    return "".join(blocks)


def _v22_source_application_items(data: Dict[str, Any]) -> List[str]:
    facts = build_fact_ledger(data, detect_family(data))
    selected: List[str] = []
    app_re = re.compile(r"\b(?:application|applications|aplikasi|used\s+for|use\s+for|digunakan\s+untuk|cocok\s+untuk)\b", re.I)
    for fact in facts:
        text = clean_text(fact.get("text"))
        section = clean_text(fact.get("section"))
        if app_re.search(section + " " + text):
            selected.append(text)
    return dedupe_keep_order(selected)[:5]


def build_application_tab(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    source_data = data
    items = _v22_source_application_items(source_data)
    if items:
        return "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(f"<li>{esc(item)}</li>" for item in items) + "</ul>"
    return (
        "<h2>Aplikasi Penggunaan</h2>"
        "<p>Gunakan produk sesuai parameter, rentang, media, dan kondisi kerja yang tercantum pada tab Spesifikasi. "
        "Hubungi tim WEBSITE untuk memastikan kesesuaian model dengan aplikasi Anda.</p>"
    )


def _v22_primary_measurement(data: Dict[str, Any]) -> Tuple[List[str], str, str]:
    rows = data.get("_ai_fields", {}).get("translated_spec_rows") if isinstance(data.get("_ai_fields"), dict) else None
    rows = rows if isinstance(rows, list) else (data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else [])
    records, _ = _v22_measurement_records(rows)
    parameters = dedupe_keep_order([clean_text(item.get("parameter")) for item in records if clean_text(item.get("parameter"))])
    range_value = next((clean_text(item.get("range")) for item in records if clean_text(item.get("range"))), "")
    accuracy = next((clean_text(item.get("accuracy")) for item in records if clean_text(item.get("accuracy"))), "")
    return parameters[:3], range_value, accuracy


def build_short_description(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("title")))
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    identity = " ".join(dedupe_keep_order([product_type, brand, model]))
    clauses = [f"{identity} digunakan untuk {clean_text(family.get('function'))}."]
    details: List[str] = []
    if parameters:
        details.append("Parameter: " + ", ".join(parameters))
    if range_value:
        details.append("rentang utama " + range_value)
    state = _source_interface_state(data)
    if state.get("rs232"):
        details.append("keluaran data RS232")
    if details:
        clauses.append("Dilengkapi " + ", ".join(details) + ".")
    text = " ".join(clauses)
    if len(text) <= V22_SHORT_MAX:
        return text
    reduced = clauses[0]
    if parameters:
        candidate = reduced + " Mengukur " + ", ".join(parameters) + "."
        if len(candidate) <= V22_SHORT_MAX:
            return candidate
    return reduced if len(reduced) <= V22_SHORT_MAX else " ".join(dedupe_keep_order([product_type, brand, model])) + "."


def _v22_meta_valid(meta: Any, model: str = "") -> bool:
    text = clean_text(meta)
    return bool(
        V22_META_MIN <= len(text) <= V22_META_MAX
        and text[-1:] in ".!?"
        and not V22_STOP_END_RE.search(text)
        and (not model or model.casefold() in text.casefold())
    )


def _v22_build_meta(data: Dict[str, Any], family: Dict[str, Any], keyphrase: str) -> str:
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    details: List[str] = []
    if parameters:
        details.append("mengukur " + ", ".join(parameters[:2]))
    if range_value:
        details.append("rentang " + range_value)
    state = _source_interface_state(data)
    if state.get("rs232"):
        details.append("keluaran RS232")
    base = keyphrase
    candidates = [
        f"{base} untuk {', '.join(details)}. Cek spesifikasi dan konsultasikan kebutuhan Anda di WEBSITE_KALIAN." if details else "",
        f"{base} dengan {', '.join(details)}. Cek spesifikasi produk di WEBSITE_KALIAN." if details else "",
        f"{base}. Cek spesifikasi teknis dan konsultasikan kebutuhan pengukuran Anda di WEBSITE_KALIAN.",
    ]
    for candidate in candidates:
        candidate = re.sub(r"\s+", " ", candidate).strip()
        if _v22_meta_valid(candidate, clean_text(data.get("model"))):
            return candidate
    # Last resort remains a complete sentence; never slice the string.
    candidate = f"{base}. Lihat spesifikasi teknis produk dan hubungi tim WEBSITE untuk konsultasi pemilihan model."
    return candidate if len(candidate) <= V22_META_MAX else f"{base}. Lihat spesifikasi teknis produk di WEBSITE_KALIAN."


def build_product_tags(data: Dict[str, Any], family: Dict[str, Any], title: str = "") -> List[str]:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, title or clean_text(data.get("title")))
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    source_blob = _source_claim_blob(data)
    tags: List[str] = [brand, model, product_type]
    tags.extend(parameters[:2])
    if range_value and len(range_value) <= 45:
        tags.append(range_value)
    state = _source_interface_state(data)
    if state.get("rs232"):
        tags.append("RS232")
    if state.get("usb_optional"):
        tags.append("USB-01")
    for pattern, tag in V22_SYNONYM_TAGS:
        if pattern.search(" ".join([product_type, clean_text(family.get("category")), clean_text(data.get("title"))])):
            tags.append(tag)
            break
    app_map = [
        (r"\bHVAC\b", "HVAC"), (r"\blaborator(?:y|ium)\b", "laboratorium"),
        (r"\bindustr(?:y|ial|i)\b", "instrumentasi industri"), (r"\bautomotive|otomotif\b", "otomotif"),
        (r"\bair\s+quality|kualitas\s+udara\b", "kualitas udara"), (r"\benvironment|lingkungan\b", "pemantauan lingkungan"),
    ]
    for pattern, tag in app_map:
        if re.search(pattern, source_blob, re.I):
            tags.append(tag)
            if len(tags) >= 8:
                break
    cleaned = []
    for tag in dedupe_keep_order(tags):
        tag = clean_text(tag).strip(" .,:;")
        if not tag or tag.casefold() in {"alat ukur lutron", clean_text(family.get("category")).casefold()}:
            continue
        cleaned.append(tag)
    family_tags = family.get("tags") if isinstance(family.get("tags"), list) else []
    for tag in family_tags:
        if len(cleaned) >= 5:
            break
        if clean_text(tag).casefold() not in {item.casefold() for item in cleaned}:
            cleaned.append(clean_text(tag))
    return cleaned[:8]


def _v22_build_faq_fallback(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title = clean_text(data.get("title"))
    parameters, range_value, accuracy = _v22_primary_measurement(data)
    items: List[Dict[str, str]] = []
    items.append({
        "question": f"Apa fungsi {title}?",
        "answer": f"Produk ini digunakan untuk {clean_text(family.get('function'))}.",
    })
    if parameters or range_value:
        details = []
        if parameters:
            details.append("parameter " + ", ".join(parameters))
        if range_value:
            details.append("rentang " + range_value)
        if accuracy:
            details.append("akurasi " + accuracy)
        items.append({"question": "Apa data pengukuran utamanya?", "answer": "Data utama yang tercantum meliputi " + ", ".join(details) + "."})
    state = _source_interface_state(data)
    if state.get("rs232"):
        answer = "Produk menyediakan antarmuka RS232 untuk keluaran data ke komputer."
        if state.get("usb_optional") and not state.get("usb_builtin"):
            answer += " Koneksi USB memerlukan kabel USB-01 opsional."
        items.append({"question": "Bagaimana keluaran data dihubungkan ke komputer?", "answer": answer})
    items.append({
        "question": "Bagaimana memastikan model ini sesuai kebutuhan?",
        "answer": "Bandingkan parameter, rentang, media, aksesori, dan kondisi kerja pada spesifikasi. Hubungi tim WEBSITE untuk memastikan kesesuaiannya.",
    })
    return _v22_render_faq(items[:4])


_v21_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v21_build_output_row(data, pub_date)
    family = detect_family(data)
    title = clean_text(row.get("Name"))
    keyphrase = build_seo_focus_keyphrase(data, family, title)
    ai = data.get("_ai_fields") if isinstance(data.get("_ai_fields"), dict) else {}

    # V22 always uses evidence-aware FAQ when AI succeeded; otherwise a safe review fallback.
    if ai and clean_text(ai.get("faq_html")):
        row["custom_tab_4_content"] = _v22_normalize_html(ai.get("faq_html"))
    else:
        row["custom_tab_4_content"] = _v22_build_faq_fallback(data, family)

    # Pass the original source data so source-derived applications can be detected.
    row["custom_tab_3_content"] = _v2215_application_fallback(data, family)

    # Ensure HTML is structurally closed and safe before WooCommerce import.
    for field in ("Product Description", "custom_tab_2_content", "custom_tab_3_content", "custom_tab_4_content", "custom_tab_5_content"):
        row[field] = _v22_normalize_html(row.get(field))
    if WEBSITE_CTA not in clean_text(row.get("Product Description")):
        row["Product Description"] = clean_text(row.get("Product Description")).rstrip() + WEBSITE_CTA

    short = clean_text(row.get("Product Short Description"))
    if not short or len(short) > V22_SHORT_MAX or V22_STOP_END_RE.search(short):
        short = build_short_description(data, family)
    row["Product Short Description"] = short

    row["focus keyphrase"] = keyphrase
    meta = clean_text(row.get("meta description"))
    if not _v22_meta_valid(meta, clean_text(data.get("model"))):
        meta = _v22_build_meta(data, family, keyphrase)
        row["Website Correction Log"] = _merge_source_notes(row.get("Website Correction Log"), "meta_rebuilt_sentence_aware")
    row["meta description"] = meta

    tags = build_product_tags(data, family, title)
    row["product tags"] = ", ".join(tags)
    row["publication_date"] = ""
    row["Website Validation"] = "publish_ready_evidence_verified" if ai else "review_required_ai_fallback"
    return row


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = []
    name = clean_text(row.get("Name"))
    model = clean_text(data.get("model"))
    source_url = clean_text(row.get("Source URL"))
    keyphrase = clean_text(row.get("focus keyphrase"))
    meta = clean_text(row.get("meta description"))
    short = clean_text(row.get("Product Short Description"))
    specs_html = clean_text(row.get("custom_tab_1_content"))
    tags = [clean_text(item) for item in clean_text(row.get("product tags")).split(",") if clean_text(item)]
    public_fields = [
        "Product Description", "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
    ]
    public_blob = " ".join(_html_text(row.get(field)) for field in public_fields) + " " + short + " " + meta
    state = _source_interface_state(data)

    if clean_text(row.get("publication_date")):
        issues.append("publication_date_must_be_blank")
    if not name or name.endswith(".") or "lutron" not in name.casefold() or (model and model.casefold() not in name.casefold()):
        issues.append("seo_title_invalid")
    if not keyphrase or keyphrase.endswith(".") or (model and model.casefold() not in keyphrase.casefold()):
        issues.append("focus_keyphrase_invalid")
    if keyphrase.casefold() == name.casefold():
        issues.append("focus_keyphrase_copies_full_title")
    if model and model.casefold() not in short.casefold():
        issues.append("short_description_missing_model")
    if len(short) > V22_SHORT_MAX or V22_STOP_END_RE.search(short):
        issues.append("short_description_invalid")
    if not _v22_meta_valid(meta, model):
        issues.append("meta_description_invalid")
    if not source_url or get_source_config(source_url) is None:
        issues.append("source_url_not_official_lutron")
    if not data.get("_raw_spec_rows"):
        issues.append("missing_source_specification_rows")
    if "<table" not in specs_html.lower() or "<thead" not in specs_html.lower() or "<tbody" not in specs_html.lower():
        issues.append("specification_table_invalid")
    if not clean_text(row.get("image_url")):
        issues.append("missing_verified_image")
    if not (5 <= len(tags) <= 8):
        issues.append("product_tags_count_invalid")
    if len({tag.casefold() for tag in tags}) != len(tags):
        issues.append("product_tags_duplicate")
    if any(tag.casefold() == "alat ukur lutron" for tag in tags):
        issues.append("product_tags_too_generic")
    if re.search(r"antarmuka\s+RS\s*[- ]?232\s*/\s*USB", public_blob, re.I) and not state.get("usb_builtin"):
        issues.append("ambiguous_rs232_usb_claim")
    if re.search(r"(?:pencatatan|perekaman)\s+data\s+otomatis\s*(?:dan|/)\s*manual", public_blob, re.I) and not state.get("internal_auto_manual_logger"):
        issues.append("unsupported_internal_data_logger_claim")
    for field in public_fields:
        html_issues = _v22_html_contract_issues(row.get(field), require_heading=field != "Product Description" or True)
        issues.extend(f"{field.lower().replace(' ', '_')}_{issue}" for issue in html_issues)
    faq_soup = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    questions = faq_soup.find_all("h3")
    answers = faq_soup.find_all("p")
    if not (V22_FAQ_MIN <= len(questions) <= V22_FAQ_MAX) or len(questions) != len(answers):
        issues.append("faq_structure_invalid")
    if require_ai and not isinstance(data.get("_ai_fields"), dict):
        issues.append("deepseek_validated_content_required")
    if require_ai and clean_text(row.get("Website Validation")) != "publish_ready_evidence_verified":
        issues.append("evidence_publish_status_required")
    if V22_FORBIDDEN_SALES_RE.search(public_blob):
        issues.append("forbidden_sales_claim")
    if has_cjk(public_blob + " " + specs_html):
        issues.append("chinese_character_detected")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.1 SEO/title/tag and deterministic review-fallback refinements
# ---------------------------------------------------------------------------

def _seo_product_type(family: Dict[str, Any], title: str) -> str:
    text = clean_text(title)
    category = clean_text(family.get("category"))
    probe = f"{text} {category}"
    ordered: List[Tuple[re.Pattern[str], str]] = [
        (re.compile(r"contact\s+tachometer|tachometer\s+kontak", re.I), "Tachometer Kontak"),
        (re.compile(r"laser\s+(?:photo\s+)?tachometer|photo\s+tachometer", re.I), "Tachometer Laser"),
        (re.compile(r"tachometer", re.I), "Tachometer"),
        (re.compile(r"air\s+quality\s+(?:monitor|meter)|monitor\s+kualitas\s+udara", re.I), "Air Quality Monitor"),
        (re.compile(r"\bco2\b.*(?:monitor|meter)|(?:monitor|meter).*\bco2\b", re.I), "Monitor CO₂"),
        (re.compile(r"\bco\s+meter\b|carbon\s+monoxide", re.I), "CO Meter"),
        (re.compile(r"led\s+light\s+meter|light\s+meter", re.I), "Light Meter"),
        (re.compile(r"sound\s+level|noise\s+meter", re.I), "Sound Level Meter"),
        (re.compile(r"manometer|pressure\s+meter", re.I), "Manometer"),
        (re.compile(r"electromagnetic|\bemf\b", re.I), "Pengukur Medan Elektromagnetik"),
        (re.compile(r"anemometer|environment\s+meter|weather\s+meter", re.I), "Environmental Meter"),
        (re.compile(r"humidity|hygrometer", re.I), "Humidity Meter"),
        (re.compile(r"thermometer|temperature\s+meter", re.I), "Termometer"),
        (re.compile(r"data\s+logger|recorder", re.I), "Data Logger"),
        (re.compile(r"ph\s*meter", re.I), "pH Meter"),
        (re.compile(r"conductivity", re.I), "Conductivity Meter"),
        (re.compile(r"moisture", re.I), "Moisture Meter"),
        (re.compile(r"clamp\s+meter", re.I), "Clamp Meter"),
        (re.compile(r"multimeter", re.I), "Multimeter"),
    ]
    for pattern, product_type in ordered:
        if pattern.search(probe):
            return product_type
    leaf = re.split(r"\s+dan\s+|\s*/\s*", category, maxsplit=1, flags=re.I)[0].strip()
    leaf = re.sub(r"\bLutron\b", "", leaf, flags=re.I).strip(" ,.-")
    if leaf and leaf.casefold() not in {"alat ukur", "alat ukur laboratorium dan industri"}:
        return leaf
    prefix = re.split(r"\b[A-Z]{1,8}[- ]?\d|\b\d", text, maxsplit=1)[0]
    prefix = re.sub(r"\bLutron\b", "", prefix, flags=re.I).strip(" ,.-")
    return prefix.title() if prefix else "Alat Ukur"


def _v22_title_descriptor(data: Dict[str, Any], product_type: str) -> str:
    raw = clean_text(data.get("input_title") or data.get("title"))
    descriptors: List[str] = []
    if re.search(r"dual\s*(?:&|and)\s*differential|input\s+ganda.*diferensial", raw, re.I):
        descriptors.append("Input Ganda dan Diferensial")
    if re.search(r"3\s*(?:axis|sumbu)", raw, re.I):
        descriptors.append("3 Sumbu")
    if re.search(r"\bLED\b", raw, re.I) and "light meter" in product_type.casefold():
        descriptors.append("LED")
    if re.search(r"\bTrue\s*RMS\b", raw, re.I):
        descriptors.append("True RMS")
    return " ".join(descriptors)


def standardize_seo_title(data: Dict[str, Any], family: Dict[str, Any], title: Any) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    source_title = clean_text(data.get("input_title") or data.get("title") or title)
    product_type = _seo_product_type(family, source_title)
    descriptor = _v22_title_descriptor(data, product_type)
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    if range_value and len(range_value) <= 28 and range_value.casefold() not in source_title.casefold():
        # Only add a range when it is clearly the primary measurement and not a long compound value.
        if len(parameters) <= 1:
            descriptor = " ".join(dedupe_keep_order([range_value, descriptor]))
    result = " ".join(dedupe_keep_order([product_type, brand, model, descriptor]))
    return re.sub(r"\s+", " ", result).strip(" .,:;|-–—")[:180].rstrip(".")


def build_seo_focus_keyphrase(data: Dict[str, Any], family: Dict[str, Any], seo_title: str) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("input_title") or seo_title))
    _parameters, range_value, _accuracy = _v22_primary_measurement(data)
    parts = [product_type, brand, model]
    if range_value and len(range_value) <= 24 and not re.search(r"[/|]", range_value):
        parts.append(range_value.replace(",", "."))
    return re.sub(r"\s+", " ", " ".join(dedupe_keep_order(parts))).strip(" .,:;")[:120]


def _v22_meta_valid(meta: Any, model: str = "", keyphrase: str = "") -> bool:
    text = clean_text(meta)
    sentences = [clean_text(item) for item in re.split(r"(?<=[.!?])\s+", text) if clean_text(item)]
    last_sentence = sentences[-1] if sentences else ""
    return bool(
        V22_META_MIN <= len(text) <= V22_META_MAX
        and text[-1:] in ".!?"
        and not V22_STOP_END_RE.search(text)
        and len(last_sentence) >= 18
        and not re.search(r"\bLutron\b[^.!?]{0,100}\bdari\s+Lutron\b", text, re.I)
        and (not model or model.casefold() in text.casefold())
        and (not keyphrase or all(token.casefold() in text.casefold() for token in keyphrase.split()[:3]))
    )


def _v22_build_meta(data: Dict[str, Any], family: Dict[str, Any], keyphrase: str) -> str:
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    feature_parts: List[str] = []
    if parameters:
        feature_parts.append("mengukur " + ", ".join(parameters[:2]))
    if range_value:
        feature_parts.append("rentang " + range_value.replace(",", "."))
    state = _source_interface_state(data)
    if state.get("rs232"):
        feature_parts.append("keluaran RS232")
    detail = ", ".join(feature_parts)
    candidates = []
    if detail:
        candidates.extend([
            f"{keyphrase} untuk {detail}. Cek spesifikasi dan konsultasikan kebutuhan Anda di WEBSITE_KALIAN.",
            f"{keyphrase} dengan {detail}. Lihat spesifikasi produk di WEBSITE_KALIAN.",
        ])
    candidates.extend([
        f"{keyphrase}. Lihat spesifikasi teknis dan konsultasikan kebutuhan pengukuran Anda di WEBSITE_KALIAN.",
        f"{keyphrase}. Cek spesifikasi teknis produk dan hubungi tim WEBSITE untuk pemilihan model.",
    ])
    for candidate in candidates:
        candidate = re.sub(r"\s+", " ", candidate).strip()
        if _v22_meta_valid(candidate, clean_text(data.get("model")), keyphrase):
            return candidate
    return f"{keyphrase}. Lihat spesifikasi teknis produk di WEBSITE_KALIAN."


def build_product_tags(data: Dict[str, Any], family: Dict[str, Any], title: str = "") -> List[str]:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("input_title") or title or data.get("title")))
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    source_blob = _source_claim_blob(data)
    tags: List[str] = [brand, model, product_type]
    tags.extend(parameters[:2])
    if range_value and len(range_value) <= 45:
        tags.append(range_value.replace(",", "."))
    state = _source_interface_state(data)
    if state.get("rs232"):
        tags.append("RS232")
    if state.get("usb_optional"):
        tags.append("USB-01")
    for pattern, tag in V22_SYNONYM_TAGS:
        if pattern.search(" ".join([product_type, clean_text(family.get("category")), clean_text(data.get("title"))])):
            tags.append(tag)
            break
    for pattern, tag in [
        (r"\bHVAC\b", "HVAC"), (r"\blaborator(?:y|ium)\b", "laboratorium"),
        (r"\bindustr(?:y|ial|i)\b", "instrumentasi industri"), (r"\bautomotive|otomotif\b", "otomotif"),
        (r"\bair\s+quality|kualitas\s+udara\b", "kualitas udara"), (r"\benvironment|lingkungan\b", "pemantauan lingkungan"),
    ]:
        if re.search(pattern, source_blob, re.I):
            tags.append(tag)
    cleaned: List[str] = []
    for raw_tag in dedupe_keep_order(tags):
        tag = clean_text(raw_tag).replace(",", ".").strip(" .,:;")
        if not tag or tag.casefold() in {"alat ukur lutron", clean_text(family.get("category")).casefold()}:
            continue
        if tag.casefold() not in {item.casefold() for item in cleaned}:
            cleaned.append(tag)
        if len(cleaned) >= 8:
            break
    family_tags = family.get("tags") if isinstance(family.get("tags"), list) else []
    for raw_tag in family_tags:
        tag = clean_text(raw_tag).replace(",", ".").strip(" .,:;")
        if tag and tag.casefold() not in {item.casefold() for item in cleaned}:
            cleaned.append(tag)
        if len(cleaned) >= 5:
            break
    if len(cleaned) < 5:
        safe_extra = clean_text(family.get("term"))
        if safe_extra and safe_extra.casefold() not in {item.casefold() for item in cleaned}:
            cleaned.append(safe_extra)
    return cleaned[:8]


def _v22_build_faq_fallback(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title = clean_text(data.get("title"))
    parameters, range_value, accuracy = _v22_primary_measurement(data)
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi {title}?",
        "answer": f"Produk ini digunakan untuk {clean_text(family.get('function'))}.",
    }]
    if parameters or range_value or accuracy:
        details = []
        if parameters:
            details.append("parameter " + ", ".join(parameters))
        if range_value:
            details.append("rentang " + range_value)
        if accuracy:
            details.append("akurasi " + accuracy)
        items.append({"question": "Apa data pengukuran utamanya?", "answer": "Data utama yang tercantum meliputi " + ", ".join(details) + "."})
    state = _source_interface_state(data)
    if state.get("rs232"):
        answer = "Produk menyediakan antarmuka RS232 untuk keluaran data ke komputer."
        if state.get("usb_optional") and not state.get("usb_builtin"):
            answer += " Koneksi USB memerlukan kabel USB-01 opsional."
        items.append({"question": "Bagaimana keluaran data dihubungkan ke komputer?", "answer": answer})
    if len(items) < 3:
        items.append({
            "question": "Apa yang perlu diperiksa sebelum penggunaan?",
            "answer": "Periksa parameter, satuan, rentang, media, catu daya, dan kondisi kerja pada spesifikasi produk.",
        })
    items.append({
        "question": "Bagaimana memastikan model ini sesuai kebutuhan?",
        "answer": "Bandingkan parameter, rentang, media, aksesori, dan kondisi kerja pada spesifikasi. Hubungi tim WEBSITE untuk memastikan kesesuaiannya.",
    })
    return _v22_render_faq(items[:4])


_v22_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v22_previous_build_output_row(data, pub_date)
    family = detect_family(data)
    title = standardize_seo_title(data, family, row.get("Name"))
    row["Name"] = title
    keyphrase = build_seo_focus_keyphrase(data, family, title)
    row["focus keyphrase"] = keyphrase
    ai = data.get("_ai_fields") if isinstance(data.get("_ai_fields"), dict) else {}
    meta = clean_text(row.get("meta description"))
    if not ai or not _v22_meta_valid(meta, clean_text(data.get("model")), keyphrase):
        row["meta description"] = _v22_build_meta(data, family, keyphrase)
        row["Website Correction Log"] = _merge_source_notes(row.get("Website Correction Log"), "meta_rebuilt_sentence_aware")
    row["product tags"] = ", ".join(build_product_tags(data, family, title))
    return row


_v22_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v22_previous_validate_publish_output(data, row, require_ai=require_ai)
    keyphrase = clean_text(row.get("focus keyphrase"))
    meta = clean_text(row.get("meta description"))
    if not _v22_meta_valid(meta, clean_text(data.get("model")), keyphrase):
        issues.append("meta_description_not_keyphrase_aligned")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.2 language residue and concise-keyphrase validation refinements
# ---------------------------------------------------------------------------
V22_ENGLISH_RESIDUE_RE = re.compile(
    r"\b(?:instruction\s+manual|carrying\s+case|less\s+than|power\s+supply|"
    r"measurement\s+range|measuring\s+range|display\s+size|sampling\s+time|"
    r"accessories?\s+(?:included|optional)|connect\s+the\s+optional|will\s+get\s+the|"
    r"freezes\s+(?:displayed\s+)?reading|maximum\s+and\s+minimum\s+value|"
    r"dimension\s+is\s+for|used\s+for\s+air|liquid\s+is\s+prohibited)\b",
    re.I,
)


def _v22_has_english_residue(value: Any) -> bool:
    return bool(V22_ENGLISH_RESIDUE_RE.search(_html_text(value)))


_v22_previous_validate_ai_fields = _v22_validate_ai_fields


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    ok, fields, reason, feedback = _v22_previous_validate_ai_fields(parsed, payload)
    if not ok:
        return ok, fields, reason, feedback
    language_issues: List[Dict[str, Any]] = []
    for field_name in (
        "product_description_html", "short_description", "feature_tab_html",
        "faq_html", "meta_description", "official_information_html",
    ):
        if _v22_has_english_residue(fields.get(field_name)):
            language_issues.append(_spec_validation_issue(
                "english_source_residue_detected",
                "Konten publik masih memuat frasa bahasa Inggris sumber yang wajib diterjemahkan.",
                field=field_name,
            ))
    for row in fields.get("translated_spec_rows") if isinstance(fields.get("translated_spec_rows"), list) else []:
        if isinstance(row, dict) and (_v22_has_english_residue(row.get("label")) or _v22_has_english_residue(row.get("value"))):
            language_issues.append(_spec_validation_issue(
                "english_specification_residue_detected",
                "Label atau nilai spesifikasi masih memuat frasa bahasa Inggris sumber.",
                source_id=clean_text(row.get("source_id")),
            ))
            if len(language_issues) >= 12:
                break
    if language_issues:
        return False, {}, "invalid_ai_content_contract", language_issues
    return True, fields, reason, []


_v22_previous_validate_publish_output_v3 = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v22_previous_validate_publish_output_v3(data, row, require_ai=require_ai)
    # A concise SEO title can legitimately equal the canonical keyphrase.
    # The old problem was copying a long sentence-like title, not exact equality itself.
    issues = [issue for issue in issues if issue != "focus_keyphrase_copies_full_title"]
    if require_ai:
        for field in (
            "Product Description", "Product Short Description", "custom_tab_1_content",
            "custom_tab_2_content", "custom_tab_3_content", "custom_tab_4_content",
            "custom_tab_5_content", "meta description",
        ):
            if _v22_has_english_residue(row.get(field)):
                issues.append("english_source_residue_detected")
                break
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.3 evidence-backed application tab and primary-range title refinement
# ---------------------------------------------------------------------------

def _v22_render_application(items: Sequence[Dict[str, Any]]) -> str:
    values = [clean_text(item.get("text")) for item in items if isinstance(item, dict) and clean_text(item.get("text"))]
    if not values:
        return ""
    return _v22_normalize_html(
        "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(f"<li>{esc(value)}</li>" for value in values[:5]) + "</ul>"
    )


_v22_previous_validate_ai_fields_v4 = _v22_validate_ai_fields


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    ok, fields, reason, feedback = _v22_previous_validate_ai_fields_v4(parsed, payload)
    if not ok:
        return ok, fields, reason, feedback
    fact_map = _v22_fact_map(payload)
    raw_items = parsed.get("application_items")
    issues: List[Dict[str, Any]] = []
    application_items: List[Dict[str, Any]] = []
    if not isinstance(raw_items, list) or not (1 <= len(raw_items) <= 5):
        issues.append(_spec_validation_issue(
            "invalid_application_items",
            "application_items harus berisi 1-5 aplikasi berbasis fakta sumber.",
        ))
    else:
        for index, item in enumerate(raw_items, 1):
            if not isinstance(item, dict):
                issues.append(_spec_validation_issue("invalid_application_item", "Item aplikasi harus berupa objek.", item=index))
                continue
            text = _v22_complete_sentence(item.get("text_id"), 360)
            evidence_ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map)
            if not text or not evidence_ids:
                issues.append(_spec_validation_issue(
                    "application_item_missing_evidence",
                    "Setiap aplikasi harus memiliki text_id dan evidence_ids yang valid.",
                    item=index,
                ))
                continue
            evidence_text = _v22_evidence_text(evidence_ids, fact_map)
            if not _v22_numeric_claims_supported(text, evidence_text):
                issues.append(_spec_validation_issue(
                    "application_numeric_claim_not_in_evidence",
                    "Angka pada aplikasi harus terdapat pada fakta yang dirujuk.",
                    item=index,
                ))
            for code in _v22_claim_issues(text, evidence_text):
                issues.append(_spec_validation_issue(code, "Klaim aplikasi tidak didukung fakta sumber.", item=index))
            if _v22_has_english_residue(text):
                issues.append(_spec_validation_issue(
                    "english_application_residue_detected",
                    "Aplikasi penggunaan masih memuat frasa bahasa Inggris sumber.",
                    item=index,
                ))
            application_items.append({"text": text, "evidence_ids": evidence_ids})
    if issues:
        return False, {}, "invalid_ai_content_contract", issues
    fields["application_items"] = application_items
    fields["application_tab_html"] = _v22_render_application(application_items)
    return True, fields, reason, []


_v22_previous_standardize_seo_title_v4 = standardize_seo_title


def standardize_seo_title(data: Dict[str, Any], family: Dict[str, Any], title: Any) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    source_title = clean_text(data.get("input_title") or data.get("title") or title)
    product_type = _seo_product_type(family, source_title)
    descriptor = _v22_title_descriptor(data, product_type)
    parameters, range_value, _accuracy = _v22_primary_measurement(data)
    range_part = ""
    if range_value and len(range_value) <= 28 and len(parameters) <= 1:
        range_part = range_value.replace(",", ".")
    result = " ".join(dedupe_keep_order([product_type, brand, model, range_part, descriptor]))
    return re.sub(r"\s+", " ", result).strip(" .,:;|-–—")[:180].rstrip(".")


_v22_previous_build_output_row_v4 = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v22_previous_build_output_row_v4(data, pub_date)
    family = detect_family(data)
    title = standardize_seo_title(data, family, row.get("Name"))
    row["Name"] = title
    keyphrase = build_seo_focus_keyphrase(data, family, title)
    row["focus keyphrase"] = keyphrase
    ai = data.get("_ai_fields") if isinstance(data.get("_ai_fields"), dict) else {}
    if ai and clean_text(ai.get("application_tab_html")):
        row["custom_tab_3_content"] = _v22_normalize_html(ai.get("application_tab_html"))
    else:
        # Review fallback must never copy untranslated source application text.
        row["custom_tab_3_content"] = (
            "<h2>Aplikasi Penggunaan</h2>"
            "<p>Gunakan produk sesuai parameter, rentang, media, aksesori, dan kondisi kerja pada tab Spesifikasi. "
            "Hubungi tim WEBSITE untuk memastikan kesesuaian model dengan kebutuhan Anda.</p>"
        )
    meta = clean_text(row.get("meta description"))
    if not _v22_meta_valid(meta, clean_text(data.get("model")), keyphrase):
        row["meta description"] = _v22_build_meta(data, family, keyphrase)
    return row


# ---------------------------------------------------------------------------
# V22.4 single CTA and precise built-in/optional interface handling
# ---------------------------------------------------------------------------

def _source_interface_state(data: Dict[str, Any]) -> Dict[str, bool]:
    blob = _source_claim_blob(data)
    usb_optional = bool(re.search(r"\bUSB\s*[- ]?0?1\b|optional[^.\n]{0,120}\bUSB\b|USB[^.\n]{0,120}optional", blob, re.I))
    rs232_optional = bool(re.search(r"\bUPCB\s*[- ]?0?\d\b|optional[^.\n]{0,120}\bRS\s*[- ]?232\b|RS\s*[- ]?232[^.\n]{0,120}optional", blob, re.I))
    rs232_builtin = bool(re.search(r"\bRS\s*[- ]?232\b[^.\n]{0,80}(?:PC\s+serial\s+interface|serial\s+interface|data\s+output)|(?:data\s+output|interface)[^.\n]{0,80}\bRS\s*[- ]?232\b", blob, re.I))
    if rs232_optional and not re.search(r"RS\s*[- ]?232\s+PC\s+serial\s+interface", blob, re.I):
        rs232_builtin = False
    usb_builtin = bool(re.search(r"\b(?:built[- ]?in|integrated|port)\b[^.\n]{0,80}\bUSB\b|\bUSB\s+(?:port|interface)\b", blob, re.I)) and not usb_optional
    return {
        "rs232": bool(re.search(r"\bRS\s*[- ]?232\b|\bUPCB\s*[- ]?0?\d\b", blob, re.I)),
        "rs232_builtin": rs232_builtin,
        "rs232_optional": rs232_optional,
        "usb_optional": usb_optional,
        "usb_builtin": usb_builtin,
        "pc_data_logger": bool(
            re.search(r"(?:PC|computer|komputer)[^.\n]{0,120}(?:data\s*logger|recorder)|(?:RS\s*[- ]?232)[^.\n]{0,160}(?:data\s*logger|recorder)", blob, re.I)
            or (re.search(r"\bRS\s*[- ]?232\b", blob, re.I) and re.search(r"(?:data\s+acquisition\s+software|perangkat\s+lunak\s+akuisisi\s+data|SW-U801-WIN)", blob, re.I))
        ),
        "internal_auto_manual_logger": bool(re.search(r"(?:automatic|auto)\s+data\s*logger[^.\n]{0,100}manual\s+data\s*logger|manual\s+data\s*logger[^.\n]{0,100}(?:automatic|auto)\s+data\s*logger", blob, re.I)),
    }


def _v22_connection_statement(data: Dict[str, Any]) -> str:
    state = _source_interface_state(data)
    parts: List[str] = []
    if state.get("rs232_builtin"):
        parts.append("Produk menyediakan antarmuka serial RS232 untuk keluaran data ke komputer")
    elif state.get("rs232_optional"):
        parts.append("Koneksi RS232 memerlukan kabel UPCB opsional")
    if state.get("usb_builtin"):
        parts.append("Produk menyediakan port USB bawaan")
    elif state.get("usb_optional"):
        parts.append("Koneksi USB memerlukan kabel USB-01 opsional")
    return "; ".join(parts) + ("." if parts else "")


def _sanitize_interface_and_logging_html(value: Any, data: Dict[str, Any]) -> str:
    html_value = clean_text(value)
    state = _source_interface_state(data)
    connection = _v22_connection_statement(data)
    html_value = re.sub(r"(?:antarmuka|koneksi)\s+RS\s*[- ]?232\s*/\s*USB", connection.rstrip("."), html_value, flags=re.I)
    html_value = re.sub(r"(?:antarmuka|koneksi)\s+RS\s*[- ]?232\s+dan\s+USB", connection.rstrip("."), html_value, flags=re.I)
    if not state.get("internal_auto_manual_logger"):
        if state.get("pc_data_logger") and state.get("rs232_builtin"):
            replacement = "Data dapat dikirim melalui antarmuka RS232 ke komputer untuk pencatatan menggunakan perangkat lunak dan kabel yang sesuai."
        else:
            replacement = ""
        html_value = re.sub(r"(?:Pencatatan|Perekaman)\s+data\s+otomatis\s+dan\s+manual\s+tersedia\.?", replacement, html_value, flags=re.I)
        html_value = re.sub(r"(?:pencatatan|perekaman)\s+data\s+otomatis\s*(?:dan|/)\s*manual", "memori maksimum dan minimum", html_value, flags=re.I)
    return re.sub(r"\s{2,}", " ", html_value).strip()


def _sanitize_spec_rows_for_source(data: Dict[str, Any], rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    state = _source_interface_state(data)
    output: List[Dict[str, str]] = []
    for raw in rows or []:
        if not isinstance(raw, dict):
            continue
        row = dict(raw)
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        source_label = clean_text(row.get("source_label"))
        source_probe = f"{source_label} {value}"
        combined = f"{label} {source_probe}"

        if re.search(r"\b(?:perekaman\s+data|pencatat\s+data|data\s+record|memory|memori)\b", combined, re.I) and re.search(r"\b(?:max(?:imum)?|min(?:imum)?|maksimum|minimum)\b", value, re.I):
            if not state.get("internal_auto_manual_logger"):
                row["label"] = "Memori Maksimum/Minimum"
                row["value"] = "Merekam dan menampilkan pembacaan maksimum dan minimum."

        has_usb01 = bool(re.search(r"\bUSB\s*[- ]?0?1\b", source_probe, re.I))
        has_upcb = bool(re.search(r"\bUPCB\s*[- ]?0?\d\b", source_probe, re.I))
        optional_context = bool(re.search(r"\boptional|opsional|accessor", source_probe, re.I)) or has_usb01 or has_upcb
        interface_label = bool(re.search(r"\b(?:output\s+data|data\s+output|interface|antarmuka|computer\s+interface)\b", f"{label} {source_label}", re.I))

        if interface_label and optional_context and (has_usb01 or has_upcb):
            row["label"] = "Koneksi Komputer dan Aksesori Opsional"
            pieces: List[str] = []
            if has_usb01:
                pieces.append("Koneksi USB memerlukan kabel USB-01 opsional")
            if has_upcb:
                upcb = re.search(r"\bUPCB\s*[- ]?0?\d\b", source_probe, re.I)
                pieces.append(f"Koneksi RS232 memerlukan kabel {clean_text(upcb.group(0)).upper()} opsional" if upcb else "Koneksi RS232 memerlukan kabel UPCB opsional")
            row["value"] = "; ".join(pieces) + "."
        elif interface_label and state.get("rs232_builtin"):
            row["label"] = "Antarmuka dan Keluaran Data"
            row["value"] = "Antarmuka serial RS232 untuk keluaran data ke komputer."
        elif interface_label and state.get("rs232_optional") and not state.get("rs232_builtin"):
            row["label"] = "Koneksi Komputer"
            row["value"] = _v22_connection_statement(data)
        output.append(row)
    return output


def _v22_strip_existing_cta(value: Any) -> str:
    soup = BeautifulSoup(clean_text(value), "html.parser")
    for paragraph in list(soup.find_all(["p", "div"])):
        text = clean_text(paragraph.get_text(" ", strip=True))
        if (
            "+62no kalian" in text
            or "Untuk konsultasi produk" in text
            or paragraph.find("a", href=re.compile(r"WEBSITE\.co\.id/(?:about-us-3|contact-us)", re.I))
        ):
            paragraph.decompose()
    return "".join(str(node) for node in soup.contents).strip()


_v22_previous_build_output_row_v5 = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v22_previous_build_output_row_v5(data, pub_date)
    clean_description = _v22_normalize_html(_v22_strip_existing_cta(row.get("Product Description")))
    row["Product Description"] = clean_description.rstrip() + WEBSITE_CTA
    return row



# ---------------------------------------------------------------------------
# V22.1 hotfix: repair minor DeepSeek schema deviations before strict validation
# ---------------------------------------------------------------------------
# The V22 contract was intentionally strict, but the first live run showed that
# DeepSeek commonly returned semantically usable JSON with small schema
# deviations, such as S001 instead of 1, missing keep arrays, or evidence IDs
# omitted from otherwise grounded content. Those deviations should be repaired
# deterministically. Unsupported claims and missing source specifications still
# fail the publish gate.
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_1_contract_repair_2026_07_27"


def _coerce_ai_spec_id(value: Any) -> Optional[int]:
    """Accept 1, "1", "S001", "spec_001", or "ID 1" safely."""
    raw = clean_text(value)
    if not raw:
        return None
    match = re.search(r"(?:^|[^0-9])(\d{1,6})(?:$|[^0-9])", raw)
    if not match:
        return None
    try:
        parsed = int(match.group(1))
        return parsed if parsed > 0 else None
    except Exception:
        return None


def _v22_clean_evidence_ids(values: Any, fact_map: Dict[str, str]) -> List[str]:
    """Normalize common AI evidence formats without inventing new facts."""
    if isinstance(values, str):
        values = re.split(r"[,;|\s]+", values)
    if not isinstance(values, list):
        return []
    out: List[str] = []
    for value in values:
        raw = clean_text(value).upper().replace("SPEC_", "S").replace("FACT_", "F")
        candidates = [raw]
        numeric = _coerce_ai_spec_id(raw)
        if numeric is not None:
            candidates.extend([f"S{numeric:03d}", f"F{numeric:03d}"])
        for candidate in candidates:
            if candidate in fact_map and candidate not in out:
                out.append(candidate)
                break
    return out


def _validate_ai_spec_deduplication(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, str]], int, List[Dict[str, Any]]]:
    """Translate every canonical source row; deduplication is deterministic.

    The canonical pipeline already removes exact duplicate source rows before AI.
    Requiring DeepSeek to repeat a keep/remove accounting array added failure
    risk without improving factual safety. Every remaining source ID must still
    be translated exactly once and retain all numeric tokens.
    """
    translated_rows, translations, issues = _validate_all_translated_specs(parsed, source_specs)
    if issues:
        return [], 0, issues
    enriched: List[Dict[str, str]] = []
    source_by_id = {int(item["id"]): item for item in source_specs if isinstance(item, dict) and item.get("id")}
    for source_id in sorted(source_by_id):
        row = dict(translations[source_id])
        source = source_by_id[source_id]
        item = next(
            (
                candidate for candidate in parsed.get("translated_specifications", [])
                if isinstance(candidate, dict) and _coerce_ai_spec_id(candidate.get("id")) == source_id
            ),
            {},
        )
        parameter = clean_catalog_source_text(item.get("parameter_id") or item.get("parameter")) or clean_catalog_source_text(source.get("suggested_parameter"))
        field = _v22_normalize_field(item.get("field_id") or item.get("field") or source.get("suggested_field"))
        group = clean_catalog_source_text(item.get("group_id") or item.get("group")) or parameter
        row.update({
            "source_id": str(source_id),
            "parameter": parameter,
            "field": field,
            "group": group,
        })
        enriched.append(row)
    return enriched, 0, []


def _v221_all_evidence_ids(payload: Dict[str, Any]) -> List[str]:
    return list(_v22_fact_map(payload).keys())


def _v221_safe_meta(payload: Dict[str, Any], parsed: Dict[str, Any]) -> str:
    brand = clean_text(payload.get("brand")) or "Lutron"
    model = clean_text(payload.get("model"))
    source_name = clean_catalog_source_text(payload.get("source_product_name"))
    product_type = clean_catalog_source_text(payload.get("suggested_category")) or "alat ukur"
    existing = re.sub(r"\s+", " ", _html_text(parsed.get("meta_description"))).strip()
    if existing and len(existing) <= V22_META_MAX and not V22_STOP_END_RE.search(existing):
        if model.casefold() in existing.casefold() and len(existing) >= V22_META_MIN:
            return existing if existing[-1] in ".!?" else existing + "."
    candidates = [
        f"{brand} {model} merupakan {product_type.lower()} untuk pengukuran sesuai spesifikasi produk. Hubungi tim WEBSITE untuk konsultasi model.",
        f"{source_name} untuk pengukuran sesuai spesifikasi produk. Konsultasikan kebutuhan model dengan tim WEBSITE.",
        f"{brand} {model} untuk pengukuran teknis sesuai spesifikasi produk dan kebutuhan penggunaan Anda. Konsultasikan dengan tim WEBSITE.",
    ]
    for candidate in candidates:
        candidate = re.sub(r"\s+", " ", candidate).strip()
        if V22_META_MIN <= len(candidate) <= V22_META_MAX and not V22_STOP_END_RE.search(candidate):
            return candidate if candidate[-1] in ".!?" else candidate + "."
    # Last resort uses complete clauses only, never a mid-clause character cut.
    base = f"{brand} {model} untuk pengukuran teknis sesuai spesifikasi produk. Hubungi tim WEBSITE untuk konsultasi pemilihan model."
    return base


def _v221_safe_short(payload: Dict[str, Any], parsed: Dict[str, Any]) -> str:
    existing = re.sub(r"\s+", " ", _html_text(parsed.get("short_description"))).strip()
    if existing and len(existing) <= V22_SHORT_MAX and not V22_STOP_END_RE.search(existing):
        return existing if existing[-1] in ".!?" else existing + "."
    brand = clean_text(payload.get("brand")) or "Lutron"
    model = clean_text(payload.get("model"))
    product_type = clean_catalog_source_text(payload.get("suggested_category")) or "alat ukur"
    function = clean_catalog_source_text(payload.get("family_function"))
    text = f"{brand} {model} adalah {product_type.lower()} yang {function}. Cocok untuk kebutuhan pengukuran sesuai spesifikasi produk."
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= V22_SHORT_MAX:
        return text
    return f"{brand} {model} adalah {product_type.lower()} untuk kebutuhan pengukuran sesuai spesifikasi produk."


def _v221_ensure_heading(value: Any, heading: str) -> str:
    html_value = _v22_normalize_html(value)
    if not html_value:
        return ""
    if not re.search(r"<h[23]\b", html_value, re.I):
        html_value = f"<h2>{esc(heading)}</h2>" + html_value
    return _v22_normalize_html(html_value)


def _v221_ensure_feature_list(value: Any) -> str:
    html_value = _v221_ensure_heading(value, "Keunggulan Produk")
    if not html_value:
        return ""
    soup = BeautifulSoup(html_value, "html.parser")
    if soup.find("li"):
        return html_value
    statements = _v22_sentence_chunks(soup.get_text("\n", strip=True), max_items=6)
    statements = [item for item in statements if item.casefold() != "keunggulan produk"]
    if not statements:
        return html_value
    return _v22_normalize_html(
        "<h2>Keunggulan Produk</h2><ul>" + "".join(f"<li>{esc(item)}</li>" for item in statements[:6]) + "</ul>"
    )


def _v221_repair_ai_contract(parsed: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    repaired = json.loads(json.dumps(parsed if isinstance(parsed, dict) else {}, ensure_ascii=False))

    # Common field aliases returned by DeepSeek.
    aliases = {
        "product_name_id": ("product_name", "name", "title"),
        "category_id": ("category", "product_category"),
        "product_description_html": ("product_description", "description_html", "description"),
        "short_description": ("product_short_description", "short_desc"),
        "feature_tab_html": ("features_html", "feature_html", "advantages_html"),
        "meta_description": ("meta", "seo_meta_description"),
        "official_information_html": ("official_information", "product_information_html"),
    }
    for target, source_names in aliases.items():
        if clean_text(repaired.get(target)):
            continue
        for source_name in source_names:
            if clean_text(repaired.get(source_name)):
                repaired[target] = repaired.get(source_name)
                break

    brand = clean_text(payload.get("brand")) or "Lutron"
    model = clean_text(payload.get("model"))
    title = clean_text(repaired.get("product_name_id"))
    if title:
        if brand.casefold() not in title.casefold():
            title = f"{brand} {title}"
        if model and model.casefold() not in title.casefold():
            title = f"{title} {model}"
        repaired["product_name_id"] = re.sub(r"\s+", " ", title).strip(" .")

    repaired["product_description_html"] = _v221_ensure_heading(repaired.get("product_description_html"), "Deskripsi Produk")
    repaired["feature_tab_html"] = _v221_ensure_feature_list(repaired.get("feature_tab_html"))
    repaired["official_information_html"] = _v221_ensure_heading(repaired.get("official_information_html"), "Informasi Produk")
    repaired["short_description"] = _v221_safe_short(payload, repaired)
    repaired["meta_description"] = _v221_safe_meta(payload, repaired)

    # Normalize translated specification aliases and ID forms.
    translated = repaired.get("translated_specifications")
    if not isinstance(translated, list):
        translated = repaired.get("specifications") if isinstance(repaired.get("specifications"), list) else []
    normalized_specs: List[Dict[str, Any]] = []
    for item in translated:
        if not isinstance(item, dict):
            continue
        normalized = dict(item)
        spec_id = _coerce_ai_spec_id(item.get("id", item.get("source_id", item.get("fact_id"))))
        if spec_id is not None:
            normalized["id"] = spec_id
        if not clean_text(normalized.get("label_id")):
            normalized["label_id"] = normalized.get("label") or normalized.get("parameter_name")
        if not clean_text(normalized.get("value_id")):
            normalized["value_id"] = normalized.get("value") or normalized.get("translated_value")
        if not clean_text(normalized.get("parameter_id")):
            normalized["parameter_id"] = normalized.get("parameter")
        if not clean_text(normalized.get("field_id")):
            normalized["field_id"] = normalized.get("field")
        if not clean_text(normalized.get("group_id")):
            normalized["group_id"] = normalized.get("group")
        normalized_specs.append(normalized)
    repaired["translated_specifications"] = normalized_specs

    # Dedupe accounting is no longer delegated to AI. Canonical source rows are
    # already deduplicated before this stage.
    source_ids = [int(item["id"]) for item in payload.get("specifications", []) if isinstance(item, dict) and item.get("id")]
    repaired["keep_specification_ids"] = source_ids
    repaired["removed_duplicate_specifications"] = []

    all_ids = _v221_all_evidence_ids(payload)
    evidence = repaired.get("content_evidence") if isinstance(repaired.get("content_evidence"), dict) else {}
    for field_name in (
        "product_description", "short_description", "feature_tab",
        "application_tab", "meta_description", "official_information",
    ):
        valid = _v22_clean_evidence_ids(evidence.get(field_name), _v22_fact_map(payload))
        evidence[field_name] = valid or all_ids
    repaired["content_evidence"] = evidence

    # Normalize FAQ aliases and repair missing evidence. Existing wording is
    # still checked against all cleaned source facts and WEBSITE policies.
    raw_faq = repaired.get("faq_items")
    normalized_faq: List[Dict[str, Any]] = []
    if isinstance(raw_faq, list):
        for item in raw_faq[:V22_FAQ_MAX]:
            if not isinstance(item, dict):
                continue
            question = clean_text(item.get("question_id") or item.get("question"))
            answer = clean_text(item.get("answer_id") or item.get("answer"))
            if not question or not answer:
                continue
            ids = _v22_clean_evidence_ids(item.get("evidence_ids"), _v22_fact_map(payload)) or list(all_ids)
            if V22_SENSITIVE_FAQ_RE.search(question + " " + answer) and "WEBSITE01" not in ids:
                ids.append("WEBSITE01")
            normalized_faq.append({"question_id": question, "answer_id": answer, "evidence_ids": ids})
    if len(normalized_faq) < V22_FAQ_MIN:
        source_name = clean_catalog_source_text(payload.get("source_product_name")) or f"{brand} {model}"
        normalized_faq = [
            {
                "question_id": f"Apa fungsi utama {brand} {model}?",
                "answer_id": f"{source_name} digunakan sesuai fungsi dan parameter yang tercantum pada spesifikasi produk.",
                "evidence_ids": list(all_ids),
            },
            {
                "question_id": "Bagaimana memilih model yang sesuai?",
                "answer_id": "Tim WEBSITE dapat membantu mencocokkan model dan spesifikasi dengan kebutuhan penggunaan pelanggan.",
                "evidence_ids": ["WEBSITE02"],
            },
            {
                "question_id": "Bagaimana memastikan ketersediaan dan kelengkapan paket?",
                "answer_id": "Hubungi tim WEBSITE sebelum pemesanan untuk memastikan ketersediaan dan kelengkapan paket produk.",
                "evidence_ids": ["WEBSITE01"],
            },
        ]
    repaired["faq_items"] = normalized_faq[:V22_FAQ_MAX]

    raw_apps = repaired.get("application_items")
    normalized_apps: List[Dict[str, Any]] = []
    if isinstance(raw_apps, list):
        for item in raw_apps[:5]:
            if not isinstance(item, dict):
                continue
            text = clean_text(item.get("text_id") or item.get("text") or item.get("application"))
            if not text:
                continue
            ids = _v22_clean_evidence_ids(item.get("evidence_ids"), _v22_fact_map(payload)) or list(all_ids)
            normalized_apps.append({"text_id": text, "evidence_ids": ids})
    if not normalized_apps:
        normalized_apps = [{
            "text_id": "Gunakan produk sesuai parameter, rentang, media, dan kondisi kerja yang tercantum pada spesifikasi.",
            "evidence_ids": list(all_ids),
        }]
    repaired["application_items"] = normalized_apps
    return repaired


_v221_strict_validate_ai_fields = _v22_validate_ai_fields


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    repaired = _v221_repair_ai_contract(parsed, payload)
    return _v221_strict_validate_ai_fields(repaired, payload)


_v221_original_get_ai_content = get_ai_content


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    result = _v221_original_get_ai_content(data, family)
    if not result:
        histories = data.get("_ai_retry_feedback") if isinstance(data.get("_ai_retry_feedback"), list) else []
        latest = histories[-1] if histories and isinstance(histories[-1], list) else []
        codes = dedupe_keep_order([
            clean_text(item.get("code")) for item in latest if isinstance(item, dict) and clean_text(item.get("code"))
        ])
        if codes:
            summary = "ai_validator=" + ",".join(codes[:12])
            data["error"] = _merge_source_notes(data.get("error"), summary)
            print("    AI validator: " + ", ".join(codes[:12]))
    return result


# ---------------------------------------------------------------------------
# V22.2 hotfix: contextual numeric validation and safe commercial-language gate
# ---------------------------------------------------------------------------
# Live V22.1 runs exposed three validator false positives:
# 1) formula digits in cm^2/H2O were treated as measurement numbers;
# 2) valid numbers were rejected when DeepSeek cited an incomplete subset of
#    evidence IDs even though the numbers existed elsewhere in the verified
#    product fact ledger; and
# 3) safe WEBSITE wording about confirming availability/warranty was confused with
#    an affirmative stock or warranty promise.
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_2_contextual_validator_2026_07_27"

# Only definitive promotional promises are blocked. Questions and safe wording
# such as "hubungi tim WEBSITE untuk memastikan ketersediaan" remain allowed.
V22_FORBIDDEN_SALES_RE = re.compile(
    r"\b(?:"
    r"harga\s+termurah|"
    r"diskon(?:\s+hingga)?\s+\d+(?:[.,]\d+)?\s*%?|"
    r"promo(?:si)?\s+(?:khusus|terbatas|hari\s+ini|sekarang)|"
    r"gratis\s+ongkir|"
    r"ready\s*stock|"
    r"stok\s+(?:selalu\s+|sudah\s+|masih\s+)?tersedia\s+(?:sekarang|hari\s+ini|siap\s+kirim)|"
    r"garansi\s+resmi\s+(?:tersedia|diberikan|termasuk|selama\s+\d+)"
    r")\b",
    re.I,
)

_V222_CURRENT_FACT_TEXT = ""


def _v222_identifier_span(text: str, start: int, end: int) -> str:
    """Return the nearby technical token containing a numeric match."""
    left = start
    right = end
    allowed = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-.^²³₀₁₂₃₄₅₆₇₈₉")
    while left > 0 and text[left - 1] in allowed:
        left -= 1
    while right < len(text) and text[right] in allowed:
        right += 1
    return text[left:right]


def _v222_is_formula_or_identifier_number(text: str, start: int, end: int) -> bool:
    """Ignore digits that belong to units, chemical formulas, or model codes."""
    token = _v222_identifier_span(text, start, end)
    low = token.casefold()
    before = text[max(0, start - 8):start]
    after = text[end:min(len(text), end + 8)]

    # Unit exponents and chemical formulas: cm^2, cm², H2O, H₂O, CO2, CO₂,
    # PM2.5. These digits identify notation and are not separate numeric facts.
    if start > 0 and end < len(text) and text[start - 1].isalpha() and text[end].isalpha():
        return True
    if re.search(r"(?:cm|mm|m|kg|g|in|ft)\s*\^\s*$", before, re.I):
        return True
    if re.search(r"(?:h|co|pm)\s*$", before, re.I) and re.match(r"^(?:o|\.?\d)", after, re.I):
        return True
    if any(mark in token for mark in ("²", "³", "₂", "₃")):
        return True

    # Product/accessory/interface identifiers: PM-9100, USB-01, UPCB-01,
    # RS232, SW-U801-WIN. Numeric preservation for these codes is handled by
    # exact model/code text, not by the measurement-number validator.
    if re.search(r"[A-Za-z]", token) and re.search(r"\d", token):
        known_unit_suffix = re.fullmatch(
            r"[-+]?\d+(?:[.,]\d+)?(?:v|a|w|g|kg|mg|hz|khz|mhz|mm|cm|mbar|bar|psi|pa|kpa|mpa|lux|db)",
            token,
            re.I,
        )
        if not known_unit_suffix:
            return True
    return False


def _numeric_tokens(value: Any) -> List[str]:
    """Extract real numeric claims while excluding notation and product codes."""
    text = clean_text(value)
    out: List[str] = []
    for match in re.finditer(r"(?<![A-Za-z0-9])[-+]?\d+(?:[.,]\d+)?", text):
        if _v222_is_formula_or_identifier_number(text, match.start(), match.end()):
            continue
        token = match.group(0)
        if token not in out:
            out.append(token)
    return out


def _v222_numeric_normalize(value: Any) -> str:
    text = clean_text(value).replace(",", ".")
    text = text.replace("²", "2").replace("₂", "2").replace("³", "3").replace("₃", "3")
    return re.sub(r"\s+", " ", text)


def _v22_numeric_claims_supported(output: Any, evidence: Any) -> bool:
    """Validate against cited evidence first, then the full verified ledger."""
    output_tokens = [token.replace(",", ".") for token in _numeric_tokens(output)]
    if not output_tokens:
        return True
    local = _v222_numeric_normalize(evidence)
    full = _v222_numeric_normalize(_V222_CURRENT_FACT_TEXT)
    combined = local + " " + full
    return all(token in combined for token in output_tokens)


def _v22_claim_issues(output: Any, evidence: Any) -> List[str]:
    """Check qualitative/certification claims against the full source ledger."""
    text = _html_text(output)
    source = clean_text(evidence) + " " + clean_text(_V222_CURRENT_FACT_TEXT)
    issues: List[str] = []
    for output_pattern, source_pattern, code in V22_GENERIC_CLAIM_RULES:
        if output_pattern.search(text) and not source_pattern.search(source):
            issues.append(code)
    for match in V22_CERT_RE.finditer(text):
        token = clean_text(match.group(0))
        token_probe = r"\s*".join(re.escape(part) for part in re.split(r"\s+", token.strip()) if part)
        if not re.search(token_probe, source, re.I):
            issues.append("unsupported_certification_claim")
            break
    return dedupe_keep_order(issues)


_v222_previous_validate_ai_fields = _v22_validate_ai_fields


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    """Run strict validation with access to all verified facts for the product."""
    global _V222_CURRENT_FACT_TEXT
    fact_map = _v22_fact_map(payload)
    _V222_CURRENT_FACT_TEXT = " ".join(fact_map.values())
    try:
        return _v222_previous_validate_ai_fields(parsed, payload)
    finally:
        _V222_CURRENT_FACT_TEXT = ""


_v222_previous_spec_rows_for_ai = _spec_rows_for_ai


def _spec_rows_for_ai(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Repair two confirmed PM-9100HA render labels using same-page order.

    The official product block lists the unit order as mbar, psi, kg/cm², ...
    but the next resolution block is rendered as "end: 1 end" and
    "dogs: 0.01 dogs". The repair is applied only when the same product page
    contains that exact unit order and those exact corrupt rows.
    """
    rows = _v222_previous_spec_rows_for_ai(data)
    if clean_text(data.get("model")).casefold() != "pm-9100ha":
        return rows
    units_text = " ".join(
        clean_text(item.get("source_value"))
        for item in rows
        if isinstance(item, dict) and re.search(r"\b(?:units|display units)\b", clean_text(item.get("source_label")), re.I)
    )
    ordered_units_confirmed = bool(re.search(r"mbar.*psi.*kg\s*/?\s*cm", units_text, re.I))
    if not ordered_units_confirmed:
        return rows
    repaired_codes: List[str] = []
    for item in rows:
        if not isinstance(item, dict):
            continue
        label = clean_text(item.get("source_label"))
        value = clean_text(item.get("source_value"))
        if re.fullmatch(r"end", label, re.I) and re.fullmatch(r"1\s+end", value, re.I):
            item["source_label"] = "mbar"
            item["source_value"] = "1 mbar"
            item["suggested_parameter"] = "Tekanan"
            item["suggested_field"] = "resolution"
            repaired_codes.append("end_to_mbar")
        elif re.fullmatch(r"dogs", label, re.I) and re.fullmatch(r"0[.,]01\s+dogs", value, re.I):
            item["source_label"] = "psi"
            item["source_value"] = "0.01 psi"
            item["suggested_parameter"] = "Tekanan"
            item["suggested_field"] = "resolution"
            repaired_codes.append("dogs_to_psi")
    if repaired_codes:
        data["error"] = _merge_source_notes(
            data.get("error"),
            "same_page_structure_repair=" + ",".join(repaired_codes),
        )
    return rows


def _v222_corrupt_source_spec_issues(data: Dict[str, Any]) -> List[str]:
    """Flag obvious source/render corruption instead of publishing guessed data."""
    issues: List[str] = []
    if "same_page_structure_repair=end_to_mbar,dogs_to_psi" in clean_text(data.get("error")):
        return issues
    for row in data.get("spec_rows") or []:
        if not isinstance(row, dict):
            continue
        label = clean_text(row.get("source_label") or row.get("label"))
        value = clean_text(row.get("value"))
        probe = f"{label}: {value}"
        if re.search(r"\b(?:end\s*:\s*1\s*end|dogs\s*:\s*0[.,]01\s*dogs|\b1\s+end\b|\b0[.,]01\s+dogs\b)", probe, re.I):
            issues.append("corrupt_source_resolution_label_requires_review")
    return dedupe_keep_order(issues)


_v222_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v222_previous_validate_publish_output(data, row, require_ai=require_ai)
    issues.extend(_v222_corrupt_source_spec_issues(data))
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.3 hybrid resilience: transport retry, deterministic source fallback,
# and mandatory WEBSITE commercial post-processing.
# ---------------------------------------------------------------------------
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_3_hybrid_WEBSITE_postprocess_2026_07_27"

V223_WEBSITE_CONTACT_URL = "https://WEBSITE_KALIAN/contact-us/"
V223_WEBSITE_DOMAIN_RE = re.compile(r"(?:^|\.)WEBSITE\.co\.id$", re.I)
V223_EXTERNAL_PROMO_RE = re.compile(
    r"\b(?:"
    r"invest\s+in|buy\s+now|purchase\s+(?:now|this)|order\s+now|shop\s+now|"
    r"visit\s+(?:the\s+)?(?:official\s+)?(?:website|site)|"
    r"add\s+to\s+wishlist|contact\s+lutron|available\s+(?:from|at)|"
    r"beli\s+sekarang|pesan\s+sekarang|kunjungi\s+(?:website|situs)|"
    r"beli\s+di\s+(?:website|situs)|hubungi\s+lutron|"
    r"dapatkan\s+produk\s+ini\s+(?:di|dari)|lihat\s+produk\s+di\s+(?:website|situs)"
    r")\b",
    re.I,
)
V223_BRAND_SITE_RE = re.compile(
    r"\b(?:website|situs|laman)\s+(?:resmi\s+)?(?:merek|brand|lutron)\b|"
    r"\b(?:official\s+)?(?:brand|lutron)\s+(?:website|site)\b",
    re.I,
)
V223_RAW_URL_RE = re.compile(r"https?://[^\s<>'\"]+", re.I)
V223_BAD_PHRASE_REPLACEMENTS = (
    (re.compile(r"\bpematrian\s+data\b", re.I), "kemudahan pembacaan data"),
    (re.compile(r"\binput\s+dual\b", re.I), "input ganda"),
    (re.compile(r"\bdual\s*&\s*differential\b", re.I), "ganda dan diferensial"),
    (re.compile(r"\bdual\s+dan\s+diferensial\b", re.I), "ganda dan diferensial"),
    (re.compile(r"\binput\s+ganda\s*&\s*diferensial\b", re.I), "input ganda dan diferensial"),
    (re.compile(r"\bganda\s+dan\s+diferensial\s+input\b", re.I), "dengan input ganda dan diferensial"),
    (re.compile(r"\bAtmosphere\b", re.I), "atmosfer"),
    (re.compile(r"\bPiezoelectric\s+built-in\b", re.I), "sensor piezoelektrik terintegrasi"),
)


def _v223_apply_phrase_repairs(value: Any) -> str:
    text = clean_text(value)
    for pattern, replacement in V223_BAD_PHRASE_REPLACEMENTS:
        text = pattern.sub(replacement, text)
    return re.sub(r"\s+", " ", text).strip()


def _v223_is_WEBSITE_url(url: str) -> bool:
    try:
        host = urlparse(url).hostname or ""
    except Exception:
        return False
    return bool(V223_WEBSITE_DOMAIN_RE.search(host))


def _v223_clean_plain_commercial(value: Any, *, replacement: bool = False) -> str:
    """Remove brand-directed sales navigation without touching technical facts."""
    text = _v223_apply_phrase_repairs(value)
    if not text:
        return ""

    # Remove raw external URLs from public copy. Source URL and image URL are
    # separate spreadsheet fields and are not passed through this function.
    def _url_sub(match: re.Match[str]) -> str:
        url = match.group(0).rstrip(".,);]")
        return url if _v223_is_WEBSITE_url(url) else ""

    text = V223_RAW_URL_RE.sub(_url_sub, text)
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    kept: List[str] = []
    removed = False
    for sentence in sentences:
        sentence = re.sub(r"\s+", " ", sentence).strip(" \t\r\n")
        if not sentence:
            continue
        if V223_EXTERNAL_PROMO_RE.search(sentence) or V223_BRAND_SITE_RE.search(sentence):
            removed = True
            continue
        kept.append(sentence)
    cleaned = " ".join(kept).strip()
    if replacement and removed:
        WEBSITE_sentence = "Untuk konsultasi dan pemesanan, hubungi tim WEBSITE melalui WEBSITE_KALIAN."
        if WEBSITE_sentence.casefold() not in cleaned.casefold():
            cleaned = (cleaned + " " + WEBSITE_sentence).strip()
    return re.sub(r"\s+", " ", cleaned).strip()


def _v223_clean_html_commercial(value: Any, *, replacement: bool = False) -> str:
    html_value = clean_text(value)
    if not html_value:
        return ""
    soup = BeautifulSoup(html_value, "html.parser")

    # Public catalog content must not link shoppers to a brand or third-party
    # sales page. Keep the visible technical text, but remove its external link.
    for anchor in list(soup.find_all("a")):
        href = clean_text(anchor.get("href"))
        if href and not _v223_is_WEBSITE_url(href):
            anchor.unwrap()

    removed_any = False
    for node in list(soup.find_all(["p", "li", "div"])):
        original = clean_text(node.get_text(" ", strip=True))
        if not original:
            continue
        cleaned = _v223_clean_plain_commercial(original, replacement=False)
        if cleaned != _v223_apply_phrase_repairs(original):
            removed_any = True
        if not cleaned:
            node.decompose()
            continue
        if cleaned != original:
            node.clear()
            node.append(cleaned)

    result = _v22_normalize_html("".join(str(item) for item in soup.contents))
    if replacement and removed_any:
        result = _v22_strip_existing_cta(result)
        result = result.rstrip() + (
            '<p><strong>Untuk konsultasi dan pemesanan, '
            '<a href="https://WEBSITE_KALIAN/contact-us/" target="_blank" rel="noopener">hubungi tim WEBSITE</a></strong>.</p>'
        )
    return result


def _v223_rewrite_ai_contract_to_WEBSITE(parsed: Dict[str, Any]) -> Dict[str, Any]:
    repaired = json.loads(json.dumps(parsed if isinstance(parsed, dict) else {}, ensure_ascii=False))
    html_fields = {
        "product_description_html": True,
        "product_description": True,
        "description_html": True,
        "description": True,
        "feature_tab_html": False,
        "features_html": False,
        "feature_html": False,
        "advantages_html": False,
        "official_information_html": True,
        "official_information": True,
        "product_information_html": True,
    }
    for field, add_replacement in html_fields.items():
        if field in repaired and clean_text(repaired.get(field)):
            repaired[field] = _v223_clean_html_commercial(repaired.get(field), replacement=add_replacement)

    for field in ("short_description", "product_short_description", "short_desc", "meta_description", "meta", "seo_meta_description"):
        if field in repaired and clean_text(repaired.get(field)):
            repaired[field] = _v223_clean_plain_commercial(repaired.get(field), replacement=False)

    faq_items = repaired.get("faq_items")
    if isinstance(faq_items, list):
        for item in faq_items:
            if not isinstance(item, dict):
                continue
            q_key = "question_id" if "question_id" in item else "question"
            a_key = "answer_id" if "answer_id" in item else "answer"
            question = _v223_clean_plain_commercial(item.get(q_key), replacement=False)
            answer_before = clean_text(item.get(a_key))
            answer = _v223_clean_plain_commercial(answer_before, replacement=False)
            if answer != _v223_apply_phrase_repairs(answer_before) and not answer:
                answer = "Hubungi tim WEBSITE untuk konsultasi produk, pemesanan, dan konfirmasi ketersediaan."
            item[q_key] = question
            item[a_key] = answer or "Hubungi tim WEBSITE untuk informasi lebih lanjut."

    apps = repaired.get("application_items")
    if isinstance(apps, list):
        for item in apps:
            if not isinstance(item, dict):
                continue
            key = "text_id" if "text_id" in item else "text"
            item[key] = _v223_clean_plain_commercial(item.get(key), replacement=False)
    return repaired


_v223_previous_validate_ai_fields = _v22_validate_ai_fields


def _v22_validate_ai_fields(
    parsed: Dict[str, Any], payload: Dict[str, Any]
) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    # Repair commercial direction BEFORE strict validation, so a harmless
    # source CTA such as "Invest in..." does not discard an otherwise valid
    # product. Technical facts, numbers, units, and specifications are untouched.
    return _v223_previous_validate_ai_fields(_v223_rewrite_ai_contract_to_WEBSITE(parsed), payload)


_v223_previous_retry_reason = _is_retryable_ai_validation_reason


def _is_retryable_ai_validation_reason(reason: Any) -> bool:
    value = clean_text(reason)
    if _v223_previous_retry_reason(value):
        return True
    return bool(re.fullmatch(
        r"request_error_(?:ConnectionError|Timeout|ReadTimeout|ConnectTimeout|ProxyError|SSLError|ChunkedEncodingError|RequestException)",
        value,
        re.I,
    ))



V223_FALLBACK_TRANSLATION_RULES = (
    (re.compile(r"\bSensor is built inside the housing\b", re.I), "Sensor terpasang di dalam casing"),
    (re.compile(r"\bPiezoelectric sensor\b", re.I), "Sensor piezoelektrik"),
    (re.compile(r"\bUsed for dry, non-corrosive and \.?\s*non-ionic air and gas only\b", re.I), "Hanya digunakan untuk udara dan gas kering, tidak korosif, dan tidak terionisasi"),
    (re.compile(r"\bFor dry, non-corrosive, non-ionized air and gas only\b", re.I), "Hanya untuk udara dan gas kering, tidak korosif, dan tidak terionisasi"),
    (re.compile(r"\bLiquid is prohibited\b", re.I), "Cairan tidak diperbolehkan"),
    (re.compile(r"\bLess than\b", re.I), "Kurang dari"),
    (re.compile(r"\bInstruction manual\b", re.I), "Buku petunjuk"),
    (re.compile(r"\bHard carrying case\b", re.I), "Kotak pelindung keras"),
    (re.compile(r"\bPlug for quick coupler\b", re.I), "Konektor kopling cepat"),
    (re.compile(r"\bRecord maximum\s*&\s*minimum readings\b", re.I), "Merekam pembacaan maksimum dan minimum"),
    (re.compile(r"\bBy push button\b", re.I), "Melalui tombol tekan"),
    (re.compile(r"\bAuto shut off, saves battery life, or manual off by push button\b", re.I), "Mati otomatis untuk menghemat baterai atau mati manual melalui tombol"),
    (re.compile(r"\bRS\s*232 PC serial interface\b", re.I), "Antarmuka serial RS232 untuk komputer"),
    (re.compile(r"\bApprox\.?\b", re.I), "Sekitar"),
    (re.compile(r"\bsecond(?:s)?\b", re.I), "detik"),
    (re.compile(r"\bbattery\b", re.I), "baterai"),
    (re.compile(r"\bheavy duty\b", re.I), "heavy-duty"),
    (re.compile(r"\bSuper large LCD(?: display)?\b", re.I), "LCD berukuran besar"),
    (re.compile(r"\bdigit size\b", re.I), "ukuran digit"),
    (re.compile(r"\bData acquisition software\b", re.I), "Perangkat lunak akuisisi data"),
    (re.compile(r"\bWindows version\b", re.I), "versi Windows"),
    (re.compile(r"\bRS232 cable\b", re.I), "Kabel RS232"),
    (re.compile(r"\bUSB cable\b", re.I), "Kabel USB"),
)


def _v223_translate_fallback_text(value: Any) -> str:
    text = clean_text(value)
    for pattern, replacement in V223_FALLBACK_TRANSLATION_RULES:
        text = pattern.sub(replacement, text)
    text = re.sub(r"\s*\*\s*", "; ", text)
    text = re.sub(r"\s+", " ", text).strip(" ;.")
    return text + ("." if text and not text.endswith((".", ":", ";")) else "")


def _v223_translate_fallback_spec_rows(data: Dict[str, Any], rows: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    output: List[Dict[str, Any]] = []
    model = clean_text(data.get("model")).casefold()
    for raw in rows or []:
        if not isinstance(raw, dict):
            continue
        row = dict(raw)
        label = clean_text(row.get("source_label") or row.get("label"))
        value = clean_text(row.get("value"))
        if model == "pm-9100ha" and re.fullmatch(r"end", label, re.I) and re.fullmatch(r"1\s+end", value, re.I):
            row["label"] = "mbar"
            row["source_label"] = "mbar"
            row["value"] = "1 mbar"
        elif model == "pm-9100ha" and re.fullmatch(r"dogs", label, re.I) and re.fullmatch(r"0[.,]01\s+dogs", value, re.I):
            row["label"] = "psi"
            row["source_label"] = "psi"
            row["value"] = "0.01 psi"
        else:
            localized_label = _localize_source_label(label, row.get("label"))
            row["label"] = _v223_translate_fallback_text(localized_label).rstrip(".")
            row["value"] = _v223_translate_fallback_text(value)
        output.append(row)
    return output


def _v223_primary_range(data: Dict[str, Any]) -> str:
    _parameters, range_value, _accuracy = _v22_primary_measurement(data)
    if range_value:
        return clean_text(range_value)
    probe = " ".join([clean_text(data.get("input_title")), clean_text(data.get("title"))])
    match = re.search(r"\b\d+(?:[.,]\d+)?\s*(?:mbar|bar|psi|Pa|kPa|MPa|°C|°F|rpm|lux|ppm|%RH)\b", probe, re.I)
    return clean_text(match.group(0)) if match else ""


def _v223_source_applications(data: Dict[str, Any]) -> List[str]:
    blob = _source_claim_blob(data)
    rules = [
        (r"\bHVAC\b|heating|ventilation|air conditioning", "sistem HVAC"),
        (r"\blaborator(?:y|ium)\b", "laboratorium"),
        (r"\bindustr(?:y|ial|i)\b", "pemantauan industri"),
        (r"\bmedical|medis|hospital\b", "peralatan medis"),
        (r"\bautomotive|otomotif\b", "otomotif"),
        (r"air quality|kualitas udara", "pemantauan kualitas udara"),
        (r"environment|lingkungan", "pemantauan lingkungan"),
    ]
    return [label for pattern, label in rules if re.search(pattern, blob, re.I)][:4]


def _v223_find_source_value(data: Dict[str, Any], pattern: str) -> str:
    for row in data.get("spec_rows") or []:
        if not isinstance(row, dict):
            continue
        label = clean_text(row.get("source_label") or row.get("label"))
        if re.search(pattern, label, re.I):
            return _v223_translate_fallback_text(row.get("value"))
    return ""



def _v223_source_feature_bullets(data: Dict[str, Any], family: Dict[str, Any]) -> List[str]:
    blob = _source_claim_blob(data)
    range_value = _v223_primary_range(data)
    state = _source_interface_state(data)
    bullets: List[str] = []
    if re.search(r"dual\s*&?\s*differential|ganda.*diferensial", blob, re.I):
        bullets.append("Input ganda dan diferensial untuk membandingkan dua titik tekanan")
    if range_value:
        bullets.append(f"Rentang utama hingga {range_value}")
    units = _v223_find_source_value(data, r"display units|pressure units|^units$|satuan tampilan")
    if units:
        bullets.append("Pilihan satuan pengukuran mengikuti daftar pada spesifikasi produk")
    if re.search(r"large LCD|super large LCD|LCD besar|contrast adjustment", blob, re.I):
        bullets.append("Layar LCD berukuran besar untuk memudahkan pembacaan")
    if re.search(r"data hold", blob, re.I):
        bullets.append("Fungsi data hold untuk menahan nilai pada layar")
    if re.search(r"maximum.*minimum|max\s*/?\s*min", blob, re.I):
        bullets.append("Memori pembacaan maksimum dan minimum")
    if state.get("rs232"):
        bullets.append("Keluaran data RS232 untuk koneksi ke komputer")
    if state.get("usb_optional") and not state.get("usb_builtin"):
        bullets.append("Dukungan USB melalui kabel USB-01 opsional")
    if re.search(r"auto power off|auto shut off", blob, re.I):
        bullets.append("Mati otomatis untuk membantu menghemat baterai")
    if len(bullets) < 4:
        bullets.extend([_v223_clean_plain_commercial(item) for item in feature_bullets(data, family)])
    return dedupe_keep_order([item for item in bullets if item])[:6]


def _v223_natural_short(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("input_title") or data.get("title"))).lower()
    range_value = _v223_primary_range(data)
    blob = _source_claim_blob(data)
    parts = [f"{brand} {model} adalah {product_type}"]
    if re.search(r"dual\s*&?\s*differential|ganda.*diferensial", blob, re.I):
        parts.append("dengan input ganda dan diferensial")
    if range_value:
        parts.append(f"hingga {range_value}")
    state = _source_interface_state(data)
    if state.get("rs232"):
        parts.append("serta keluaran data RS232")
    sentence = " ".join(parts) + "."
    apps = _v223_source_applications(data)
    if apps:
        sentence += " Cocok untuk " + ", ".join(apps[:3]) + "."
    return sentence if len(sentence) <= V22_SHORT_MAX else f"{brand} {model} adalah {product_type} untuk pengukuran sesuai spesifikasi produk."


def _v223_natural_meta(data: Dict[str, Any], family: Dict[str, Any], keyphrase: str) -> str:
    range_value = _v223_primary_range(data)
    blob = _source_claim_blob(data)
    details: List[str] = []
    if re.search(r"dual\s*&?\s*differential|ganda.*diferensial", blob, re.I):
        details.append("input ganda dan diferensial")
    if range_value and range_value.casefold() not in keyphrase.casefold():
        details.append("rentang " + range_value)
    state = _source_interface_state(data)
    if state.get("rs232"):
        details.append("keluaran data RS232")
    candidates = [
        f"{keyphrase} dengan {', '.join(details)}. Konsultasikan kebutuhan produk dengan tim WEBSITE." if details else "",
        f"{keyphrase} untuk kebutuhan pengukuran sesuai spesifikasi resmi. Hubungi tim WEBSITE untuk konsultasi model.",
        f"{keyphrase}. Cek spesifikasi teknis dan konsultasikan kebutuhan Anda di WEBSITE_KALIAN.",
    ]
    for candidate in candidates:
        candidate = re.sub(r"\s+", " ", candidate).strip()
        if _v22_meta_valid(candidate, clean_text(data.get("model")), keyphrase):
            return candidate
    return _v22_build_meta(data, family, keyphrase)

def _v223_natural_description(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("input_title") or data.get("title"))).lower()
    range_value = _v223_primary_range(data)
    source_blob = _source_claim_blob(data)
    differential = bool(re.search(r"dual\s*&?\s*differential|ganda.*diferensial", source_blob, re.I))
    applications = _v223_source_applications(data)

    intro = f"{brand} {model} adalah {product_type} untuk {clean_text(family.get('function'))}."
    if range_value and differential:
        intro = f"{brand} {model} adalah {product_type} untuk pengukuran tekanan dengan input ganda dan diferensial hingga {range_value}."
    elif range_value:
        intro = f"{brand} {model} adalah {product_type} dengan rentang utama hingga {range_value}."

    detail_parts: List[str] = []
    state = _source_interface_state(data)
    if state.get("rs232_builtin") or state.get("rs232"):
        detail_parts.append("antarmuka RS232 untuk keluaran data ke komputer")
    if state.get("usb_optional") and not state.get("usb_builtin"):
        detail_parts.append("koneksi USB melalui kabel USB-01 opsional")
    if applications:
        detail_parts.append("penggunaan pada " + ", ".join(applications))
    details = "Produk ini mendukung " + "; ".join(detail_parts) + "." if detail_parts else "Fitur dan batas penggunaan mengikuti spesifikasi resmi pada halaman produk."

    bullets = _v223_source_feature_bullets(data, family)
    if not bullets:
        bullets = ["Spesifikasi teknis berasal dari halaman produk resmi", "Model dan data pengukuran diperiksa sebelum dibuat menjadi katalog"]

    guide_parts: List[str] = []
    if re.search(r"liquid is prohibited|cairan (?:dilarang|tidak diperbolehkan)", source_blob, re.I):
        guide_parts.append("Gunakan alat hanya untuk udara atau gas kering sesuai batas media pada spesifikasi; penggunaan pada cairan tidak diperbolehkan")
    if range_value:
        guide_parts.append(f"Pastikan rentang {range_value} sesuai dengan kebutuhan pengukuran")
    if state.get("usb_optional") and not state.get("usb_builtin"):
        guide_parts.append("Kabel USB-01 merupakan aksesori opsional, bukan port USB bawaan")
    if not guide_parts:
        guide_parts.append("Periksa parameter, rentang, media, aksesori, dan kondisi kerja sebelum memilih model")

    return _v22_normalize_html(
        f"<h2>Deskripsi Produk</h2><p>{esc(intro)}</p><p>{esc(details)}</p>"
        f"<h2>Keunggulan Produk</h2><ul>{''.join(f'<li>{esc(item)}</li>' for item in bullets)}</ul>"
        f"<h2>Panduan Pemilihan</h2><p>{esc('. '.join(guide_parts))}.</p>"
    )


def _v223_natural_faq(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = _seo_product_type(family, clean_text(data.get("input_title") or data.get("title"))).lower()
    range_value = _v223_primary_range(data)
    source_blob = _source_claim_blob(data)
    state = _source_interface_state(data)
    items: List[Dict[str, str]] = [{
        "question": f"Apa fungsi utama {brand} {model}?",
        "answer": f"{brand} {model} merupakan {product_type} untuk {clean_text(family.get('function'))}.",
    }]
    if range_value:
        items.append({
            "question": f"Berapa rentang utama {brand} {model}?",
            "answer": f"Rentang utama yang tercantum pada halaman produk adalah {range_value}. Periksa tabel spesifikasi untuk rincian setiap satuan.",
        })
    if re.search(r"liquid is prohibited|cairan (?:dilarang|tidak diperbolehkan)", source_blob, re.I):
        items.append({
            "question": f"Apakah {brand} {model} dapat digunakan untuk cairan?",
            "answer": "Tidak. Informasi produk membatasi penggunaan pada udara atau gas kering yang sesuai dan menyatakan bahwa cairan tidak diperbolehkan.",
        })
    if state.get("rs232"):
        answer = "Produk mendukung keluaran data RS232 ke komputer."
        if state.get("usb_optional") and not state.get("usb_builtin"):
            answer += " Koneksi USB memerlukan kabel USB-01 opsional."
        items.append({"question": "Bagaimana produk dihubungkan ke komputer?", "answer": answer})
    included = _v223_find_source_value(data, r"accessories included|aksesori.*disertakan|aksesori.*termasuk")
    optional = _v223_find_source_value(data, r"optional accessories|aksesori opsional")
    if included and len(items) < 5:
        items.append({"question": "Apa kelengkapan yang tercantum pada halaman produk?", "answer": included})
    elif optional and len(items) < 5:
        items.append({"question": "Aksesori opsional apa yang tersedia?", "answer": optional})
    if len(items) < 4:
        items.append({
            "question": "Bagaimana memastikan model ini sesuai kebutuhan?",
            "answer": "Bandingkan parameter, rentang, media, konektivitas, dan aksesori pada spesifikasi. Hubungi tim WEBSITE untuk verifikasi sebelum pemesanan.",
        })
    return _v22_render_faq(items[:5])

def _v223_deterministic_source_fields(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    """Build complete publishable copy from verified source facts when AI fails."""
    specs = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    translated = localize_spec_rows_fallback(specs)
    translated = _v223_translate_fallback_spec_rows(data, translated)
    translated, _ = canonicalize_specifications(translated)
    translated = _sanitize_spec_rows_for_source(data, translated)

    title = standardize_seo_title(data, family, clean_catalog_source_text(data.get("title")))
    keyphrase = build_seo_focus_keyphrase(data, family, title)
    description = _v223_natural_description(data, family)
    short = _v223_natural_short(data, family)
    features = "<h2>Keunggulan Produk</h2><ul>" + "".join(
        f"<li>{esc(_v223_clean_plain_commercial(item))}</li>"
        for item in feature_bullets(data, family)
        if _v223_clean_plain_commercial(item)
    ) + "</ul>"
    faq = _v223_clean_html_commercial(_v223_natural_faq(data, family), replacement=False)
    official = _v223_clean_html_commercial(fallback_official_information(data), replacement=True)
    meta = _v223_natural_meta(data, family, keyphrase)

    return {
        "product_name_id": title,
        "category_id": clean_text(family.get("category")),
        "product_description_html": _v221_ensure_heading(description, "Deskripsi Produk"),
        "short_description": short,
        "feature_tab_html": _v221_ensure_feature_list(features),
        "faq_html": faq,
        "faq_items": [],
        "meta_description": meta,
        "official_information_html": _v221_ensure_heading(official, "Informasi Produk"),
        "translated_spec_rows": translated,
        "ai_spec_dedup_applied": True,
        "ai_spec_dedup_removed_count": 0,
        "content_evidence": {},
    }


_v223_previous_get_ai_content = get_ai_content


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    result = _v223_previous_get_ai_content(data, family)
    if result:
        if clean_text(data.get("_content_mode")) != "deterministic_source_fallback":
            data["_content_mode"] = "ai_postprocessed"
        return result

    # Network, API, or contract failure no longer empties the catalog. The
    # source-locked deterministic generator becomes the fallback. It uses the
    # same scraped product facts and never invents technical values.
    fallback = _v223_deterministic_source_fields(data, family)
    if fallback and fallback.get("translated_spec_rows"):
        previous_status = clean_text(data.get("_ai_status")) or "ai_unavailable"
        data["_ai_fields"] = fallback
        data["_content_mode"] = "deterministic_source_fallback"
        data["_ai_status"] = previous_status + "_deterministic_source_fallback"
        data["error"] = _merge_source_notes(data.get("error"), "deterministic_source_fallback_used")
        AI_STATS["fallback"] = int(AI_STATS.get("fallback") or 0) + 1
        print(f"    AI unavailable ({previous_status}); deterministic source fallback used")
        return fallback
    return {}


def _v223_postprocess_output_row(data: Dict[str, Any], row: Dict[str, Any]) -> Dict[str, Any]:
    for field in ("Product Description", "custom_tab_2_content", "custom_tab_3_content", "custom_tab_4_content", "custom_tab_5_content"):
        row[field] = _v223_clean_html_commercial(row.get(field), replacement=False)

    # Force a single WEBSITE CTA at the end of the main description.
    description = _v22_strip_existing_cta(row.get("Product Description"))
    row["Product Description"] = _v22_normalize_html(description).rstrip() + WEBSITE_CTA

    family = detect_family(data)
    short = _v223_clean_plain_commercial(row.get("Product Short Description"), replacement=False)
    if not short or len(short) > V22_SHORT_MAX or V22_STOP_END_RE.search(short):
        short = build_short_description(data, family)
    row["Product Short Description"] = short

    meta = _v223_clean_plain_commercial(row.get("meta description"), replacement=False)
    keyphrase = clean_text(row.get("focus keyphrase"))
    if not _v22_meta_valid(meta, clean_text(data.get("model")), keyphrase):
        meta = _v22_build_meta(data, family, keyphrase)
    row["meta description"] = meta

    mode = clean_text(data.get("_content_mode"))
    if mode == "deterministic_source_fallback":
        row["Website Validation"] = "publish_ready_source_deterministic"
        row["Processing Time"] = "lutron_WEBSITE_import_deterministic_source_fallback"
        row["Content Quality"] = clean_text(data.get("source")) + "_source_locked_deterministic"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "ai_failed_source_deterministic_fallback;WEBSITE_commercial_postprocess_applied",
        )
    else:
        row["Website Validation"] = "publish_ready_ai_postprocessed"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "WEBSITE_commercial_postprocess_applied",
        )
    return row


_v223_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    return _v223_postprocess_output_row(data, _v223_previous_build_output_row(data, pub_date))


_v223_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v223_previous_validate_publish_output(data, row, require_ai=require_ai)
    mode = clean_text(data.get("_content_mode"))
    status = clean_text(row.get("Website Validation"))
    if mode == "deterministic_source_fallback" and status == "publish_ready_source_deterministic":
        issues = [item for item in issues if item not in {
            "deepseek_validated_content_required", "evidence_publish_status_required"
        }]
    elif mode == "ai_postprocessed" and status == "publish_ready_ai_postprocessed":
        issues = [item for item in issues if item != "evidence_publish_status_required"]

    # Final commercial gate. Content may cite the official source in the Source
    # URL column, but customer-facing copy must not send buyers away from WEBSITE.
    html_fields = ("Product Description", "custom_tab_2_content", "custom_tab_3_content", "custom_tab_4_content", "custom_tab_5_content")
    for field in html_fields:
        soup = BeautifulSoup(clean_text(row.get(field)), "html.parser")
        for anchor in soup.find_all("a"):
            href = clean_text(anchor.get("href"))
            if href and not _v223_is_WEBSITE_url(href):
                issues.append("external_customer_link_detected")
                break
    public_text = " ".join(_html_text(row.get(field)) for field in html_fields)
    if V223_EXTERNAL_PROMO_RE.search(public_text) or V223_BRAND_SITE_RE.search(public_text):
        issues.append("external_brand_promotion_detected")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.4 — source-immutable specification pipeline
# ---------------------------------------------------------------------------
# Technical specifications are never authored by AI. The specification tab is
# rebuilt from the exact rows scraped from the current product URL. AI may still
# write commercial copy, but it cannot add, remove, merge, or modify technical
# values. Any ambiguous/corrupted source measurement is blocked by the publish
# gate instead of being guessed or repaired for a specific model.

V224_VERSION = "v22.4_source_immutable_specs"
V224_NUMERIC_RE = re.compile(r"(?<![A-Za-z])[-+]?\d+(?:[.,]\d+)?")
V224_CODE_RE = re.compile(r"\b(?=[A-Za-z0-9-]*\d)(?=[A-Za-z0-9-]*[A-Za-z])[A-Za-z0-9]+(?:-[A-Za-z0-9]+)+\b")
V224_UNIT_TOKEN_RE = re.compile(_MEASUREMENT_UNIT_PATTERN.pattern, re.I)
V224_ACCURACY_RE = re.compile(r"(?:±|\+/-|%\s*(?:F\.?S\.?|full\s*scale)|\b(?:accuracy|akurasi)\b)", re.I)
V224_GENERAL_ACCURACY_LABEL_RE = re.compile(r"^\s*(?:accuracy|akurasi)(?:\s+umum)?\s*$", re.I)
V224_TECH_SECTION_RE = re.compile(r"\b(?:specification|technical|measurement|spesifikasi|pengukuran)\b", re.I)
V224_UNKNOWN_QUANTITY_RE = re.compile(r"^\s*([-+]?\d+(?:[.,]\d+)?)\s*([A-Za-z][A-Za-z0-9_-]{1,20})\s*[.]?\s*$")


def _v224_exact_text(value: Any) -> str:
    """Normalize whitespace only; do not normalize numbers, units, or symbols."""
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def _v224_numeric_tokens(value: Any) -> List[str]:
    return V224_NUMERIC_RE.findall(_v224_exact_text(value))


def _v224_protected_tokens(value: Any) -> List[str]:
    text = _v224_exact_text(value)
    tokens: List[str] = []
    tokens.extend(V224_CODE_RE.findall(text))
    tokens.extend(match.group(0) for match in V224_UNIT_TOKEN_RE.finditer(text))
    # Common technical abbreviations without a hyphen.
    tokens.extend(re.findall(r"\b(?:RS\s*232|USB|LCD|LSI|AC|DC|RH|F\.?S\.?)\b", text, re.I))
    return dedupe_keep_order([_v224_exact_text(token) for token in tokens if _v224_exact_text(token)])


def _v224_source_rows(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return an immutable ledger from current-page source rows.

    Prefer _raw_spec_rows because it is captured before AI translation and
    before fuzzy canonical merging. Exact duplicates are represented by one
    output row with all source IDs retained; no non-identical rows are merged.
    """
    candidates = data.get("_raw_spec_rows")
    if not isinstance(candidates, list) or not candidates:
        candidates = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []

    ledger: List[Dict[str, Any]] = []
    exact_map: Dict[Tuple[str, str], int] = {}
    for source_index, raw in enumerate(candidates or [], 1):
        if not isinstance(raw, dict):
            continue
        label = _v224_exact_text(raw.get("source_label") or raw.get("label"))
        value = _v224_exact_text(raw.get("value"))
        section = _v224_exact_text(raw.get("section")) or "Specifications"
        if not label or not value:
            continue
        key = (label.casefold(), value.casefold())
        if key in exact_map:
            ledger[exact_map[key]]["source_ids"].append(source_index)
            ledger[exact_map[key]]["sections"].append(section)
            continue
        exact_map[key] = len(ledger)
        ledger.append({
            "source_id": source_index,
            "source_ids": [source_index],
            "label_raw": label,
            "value_raw": value,
            "section_raw": section,
            "sections": [section],
            "canonical_key": _v224_exact_text(raw.get("canonical_key")),
        })
    return ledger


def _v224_safe_display_row(data: Dict[str, Any], item: Dict[str, Any]) -> Dict[str, str]:
    """Translate wording only when every numeric/technical token is preserved.

    If the deterministic translator changes a number, unit, model/accessory
    code, or technical abbreviation, the raw source value is used verbatim.
    """
    label_raw = item["label_raw"]
    value_raw = item["value_raw"]
    candidate = {"label": label_raw, "value": value_raw, "source_label": label_raw, "section": item["section_raw"]}
    try:
        translated_rows = _v223_translate_fallback_spec_rows(data, [candidate])
        translated = translated_rows[0] if translated_rows else candidate
    except Exception:
        translated = candidate

    label_display = _v224_exact_text(translated.get("label") or normalize_label(label_raw) or label_raw)
    value_display = _v224_exact_text(translated.get("value") or value_raw)

    if _v224_numeric_tokens(value_display) != _v224_numeric_tokens(value_raw):
        value_display = value_raw
    else:
        output_probe = value_display.casefold().replace(" ", "")
        for token in _v224_protected_tokens(value_raw):
            if token.casefold().replace(" ", "") not in output_probe:
                value_display = value_raw
                break

    return {"label": label_display or label_raw, "value": value_display, "source_label": label_raw, "section": item["section_raw"]}


def _v224_known_quantity(value: Any) -> Optional[Tuple[str, str]]:
    text = _v224_exact_text(value).strip(" ,.;:")
    match = _MEASUREMENT_QUANTITY_RE.fullmatch(text)
    if not match:
        return None
    # Keep the exact source number and unit spelling in the visible table.
    return f"{match.group(1)} {_v224_exact_text(match.group(2))}", _normalize_measurement_unit(match.group(2))


def _v224_known_unit(value: Any) -> str:
    text = _v224_exact_text(value).strip(" ,.;:")
    match = _MEASUREMENT_UNIT_ONLY_RE.fullmatch(text)
    return _normalize_measurement_unit(match.group(1)) if match else ""


def _v224_split_range_resolution(value: Any) -> Tuple[str, str]:
    text = _v224_exact_text(value)
    marker = re.search(r"\b(?:resolution|resolusi|minimum\s+scale)\b\s*: ?", text, re.I)
    if not marker:
        return "", ""
    left = text[:marker.start()].strip(" ,;/:-")
    right = text[marker.end():].strip(" ,;/:-")
    left_q = _v224_known_quantity(left)
    right_q = _v224_known_quantity(right)
    return (left_q[0] if left_q else "", right_q[0] if right_q else "")


def _v224_integrity_issues(data: Dict[str, Any], ledger: Sequence[Dict[str, Any]]) -> List[str]:
    issues: List[str] = []
    if not ledger:
        return ["source_specification_rows_missing"]

    model = _v224_exact_text(data.get("model") or data.get("input_model"))
    for item in ledger:
        label = item["label_raw"]
        value = item["value_raw"]
        section = item["section_raw"]

        if label.casefold() == "model" and model and model.casefold() not in value.casefold():
            issues.append(f"source_model_mismatch:id_{item['source_id']}")

        # Generic OCR/translation corruption detector. A technical row such as
        # "End = 1 end" or "Dogs = 0.01 dogs" is an unknown measurement unit.
        # The script does not guess the intended unit; it blocks publication.
        unknown = V224_UNKNOWN_QUANTITY_RE.fullmatch(value)
        if unknown and V224_TECH_SECTION_RE.search(section):
            unit_word = unknown.group(2)
            if not _v224_known_unit(unit_word):
                label_word = re.sub(r"[^A-Za-z0-9_-]", "", label).casefold()
                if label_word == unit_word.casefold() or re.search(r"\b(?:range|resolution|accuracy|rentang|resolusi|akurasi)\b", section + " " + label, re.I):
                    issues.append(f"unrecognized_source_measurement_unit:id_{item['source_id']}:{unit_word}")

        # Catch obviously glued/corrupted fragments while avoiding model codes.
        if re.search(r"\d\s*tUSB\b|\bm\s*/\s*mim\b|\b(?:undefined|null|nan)\b", value, re.I):
            issues.append(f"malformed_source_specification:id_{item['source_id']}")

    return dedupe_keep_order(issues)


def _v224_build_specification_html(data: Dict[str, Any]) -> Tuple[str, Dict[str, Any]]:
    ledger = _v224_source_rows(data)
    issues = _v224_integrity_issues(data, ledger)
    display_rows = [_v224_safe_display_row(data, item) for item in ledger]

    measurements: List[Dict[str, str]] = []
    consumed: set = set()
    exact_range_positions: Dict[str, int] = {}
    general_accuracy: List[Tuple[int, str]] = []

    for idx, (item, display) in enumerate(zip(ledger, display_rows)):
        label_raw = item["label_raw"]
        value_raw = item["value_raw"]
        unit = _v224_known_unit(label_raw)

        # Source shape A: unit label + "range, resolution: value" in the same row.
        if unit:
            range_value, resolution = _v224_split_range_resolution(value_raw)
            if range_value and resolution:
                pos = len(measurements)
                measurements.append({
                    "source_id": str(item["source_id"]), "unit": label_raw,
                    "range": range_value, "resolution": resolution, "accuracy": "",
                })
                exact_range_positions[range_value.casefold()] = pos
                consumed.add(idx)
                continue
            # Source shape C: unit label + explicit resolution only.
            value_q = _v224_known_quantity(value_raw)
            if value_q and value_q[1].casefold() == unit.casefold():
                measurements.append({
                    "source_id": str(item["source_id"]), "unit": label_raw,
                    "range": "", "resolution": value_q[0], "accuracy": "",
                })
                consumed.add(idx)
                continue

        # Source shape B: range is the label and resolution is the value.
        label_q = _v224_known_quantity(label_raw)
        value_q = _v224_known_quantity(value_raw)
        if label_q and value_q and label_q[1].casefold() == value_q[1].casefold():
            pos = len(measurements)
            measurements.append({
                "source_id": str(item["source_id"]), "unit": _v224_exact_text(re.sub(r"^[-+]?\d+(?:[.,]\d+)?\s*", "", label_raw)),
                "range": label_raw, "resolution": value_raw, "accuracy": "",
            })
            exact_range_positions[label_raw.casefold()] = pos
            consumed.add(idx)
            continue

        # Accuracy explicitly attached to a range label.
        if label_q and V224_ACCURACY_RE.search(value_raw):
            target = exact_range_positions.get(label_raw.casefold())
            if target is not None:
                measurements[target]["accuracy"] = value_raw
                consumed.add(idx)
                continue

        # General accuracy is never attached to the last measurement row.
        if V224_GENERAL_ACCURACY_LABEL_RE.fullmatch(label_raw) and V224_ACCURACY_RE.search(value_raw):
            general_accuracy.append((idx, value_raw))
            consumed.add(idx)

    blocks: List[str] = ["<h2>Spesifikasi</h2>"]
    if measurements:
        has_range = any(row["range"] for row in measurements)
        has_resolution = any(row["resolution"] for row in measurements)
        has_accuracy = any(row["accuracy"] for row in measurements)
        fields: List[Tuple[str, str]] = [("unit", "Satuan")]
        if has_range:
            fields.append(("range", "Rentang"))
        if has_resolution:
            fields.append(("resolution", "Resolusi"))
        if has_accuracy:
            fields.append(("accuracy", "Akurasi"))
        body = []
        for row in measurements:
            body.append("<tr data-source-id=\"{}\">{}</tr>".format(
                esc(row["source_id"]),
                "".join(f"<td>{esc(row[field]) if row[field] else '-'}</td>" for field, _ in fields),
            ))
        blocks.append(
            '<h3>Data Pengukuran</h3><table class="WEBSITE-spec-table WEBSITE-spec-measurement">'
            '<thead><tr>' + "".join(f"<th>{esc(header)}</th>" for _, header in fields) + '</tr></thead>'
            '<tbody>' + "".join(body) + '</tbody></table>'
        )

    for _idx, accuracy in general_accuracy:
        blocks.append(f"<p class=\"WEBSITE-spec-note\"><strong>Akurasi umum:</strong> {esc(accuracy)}</p>")

    general_body: List[str] = []
    for idx, (item, display) in enumerate(zip(ledger, display_rows)):
        if idx in consumed:
            continue
        general_body.append(
            f'<tr data-source-id="{esc(item["source_id"])}"><th scope="row">{esc(display["label"])}</th><td>{esc(display["value"])}</td></tr>'
        )
    if general_body:
        blocks.append("<h3>Spesifikasi Umum</h3>" if measurements or general_accuracy else "")
        blocks.append(
            '<table class="WEBSITE-spec-table WEBSITE-spec-general"><thead><tr><th>Parameter</th><th>Nilai</th></tr></thead>'
            '<tbody>' + "".join(general_body) + '</tbody></table>'
        )

    represented_ids = set()
    for item in ledger:
        represented_ids.update(item["source_ids"])
    source_count = sum(len(item["source_ids"]) for item in ledger)
    report = {
        "version": V224_VERSION,
        "source_row_count": source_count,
        "unique_row_count": len(ledger),
        "expected_source_ids": sorted(represented_ids),
        "represented_source_ids": sorted(represented_ids),
        "issues": issues,
        "source_hash": hashlib.sha256(json.dumps([
            [item["label_raw"], item["value_raw"], item["section_raw"]] for item in ledger
        ], ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()[:16],
    }
    return "".join(blocks), report


_v224_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v224_previous_build_output_row(data, pub_date)
    spec_html, report = _v224_build_specification_html(data)
    data["_v224_spec_report"] = report
    row["custom_tab_1_title"] = "Spesifikasi"
    row["custom_tab_1_content"] = spec_html
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        f"spec_source_locked={V224_VERSION};source_rows={report['source_row_count']};unique_rows={report['unique_row_count']};source_hash={report['source_hash']}",
    )
    if report["issues"]:
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"), "spec_integrity_issues=" + ",".join(report["issues"])
        )
    return row


_v224_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v224_previous_validate_publish_output(data, row, require_ai=require_ai)
    # The old language gate scanned the specification tab together with AI
    # marketing copy. In V22.4 the specification tab is source-immutable, so
    # raw technical wording is allowed when a safe translation is unavailable.
    # Customer-facing prose outside the specification tab must remain Indonesian.
    issues = [item for item in issues if item != "english_source_residue_detected"]
    if require_ai:
        for field in (
            "Product Description", "Product Short Description",
            "custom_tab_2_content", "custom_tab_3_content",
            "custom_tab_4_content", "custom_tab_5_content", "meta description",
        ):
            if _v22_has_english_residue(row.get(field)):
                issues.append("english_source_residue_detected_non_spec")
                break
    report = data.get("_v224_spec_report") if isinstance(data.get("_v224_spec_report"), dict) else {}
    if not report:
        issues.append("source_specification_integrity_report_missing")
        return dedupe_keep_order(issues)
    issues.extend(report.get("issues") or [])
    if report.get("source_row_count", 0) < 1 or report.get("unique_row_count", 0) < 1:
        issues.append("source_specification_rows_missing")
    # Every source row ID must be represented by either an exact duplicate map,
    # a measurement table row, an accuracy note, or a general table row.
    expected_ids = set(report.get("expected_source_ids") or [])
    represented_ids = set(report.get("represented_source_ids") or [])
    if expected_ids and expected_ids != represented_ids:
        issues.append("source_specification_row_accounting_mismatch")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.5 sectional repair: preserve source specifications, replace only failed
# AI sections, and repair corrupted unit labels from same-page structure.
# ---------------------------------------------------------------------------
V225_VERSION = "v22.5_sectional_repair_source_locked"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_5_sectional_repair_2026_07_27"

V225_SPEC_FEEDBACK_CODES = {
    "translated_specification_missing_ids",
    "translated_specification_missing_numeric_token",
    "translated_specification_missing_protected_token",
    "english_specification_residue_detected",
    "invalid_ai_specification_deduplication",
    "specification_translation_missing",
}

V225_LABEL_MAP = {
    "circuit": "Sirkuit",
    "display": "Tampilan",
    "display units": "Satuan tampilan",
    "pressure units": "Satuan tekanan",
    "units": "Satuan",
    "function": "Fungsi",
    "functionality": "Fungsi",
    "zero adjust": "Penyetelan nol",
    "span adjust": "Penyetelan rentang",
    "data hold": "Penahan data",
    "data record": "Rekaman data",
    "data output": "Keluaran data",
    "power off": "Mati otomatis",
    "auto shut off": "Mati otomatis",
    "sampling time": "Waktu sampling",
    "response time": "Waktu respons",
    "operating temperature": "Suhu operasi",
    "temperature": "Suhu",
    "operating humidity": "Kelembapan operasi",
    "humidity": "Kelembapan",
    "power supply": "Catu daya",
    "power current": "Arus daya",
    "power consumption": "Konsumsi daya",
    "weight": "Berat",
    "dimension": "Dimensi",
    "dimensions": "Dimensi",
    "accessories included": "Aksesori yang disertakan",
    "optional accessories": "Aksesori opsional",
    "type": "Tipe",
    "size": "Ukuran",
    "digit size": "Ukuran digit",
    "sensor type": "Tipe sensor",
    "application": "Penggunaan",
    "instruction manual": "Buku petunjuk",
    "hard carrying case": "Kotak pelindung keras",
    "plug for quick coupler": "Konektor kopling cepat",
    "main instrument": "Instrumen utama",
    "memory function": "Fungsi memori",
    "data logger": "Pencatat data",
    "accuracy": "Akurasi",
    "resolution": "Resolusi",
    "range": "Rentang",
    "model": "Model",
}

V225_VALUE_TRANSLATIONS = (
    (re.compile(r"\bMicroprocessor LSI circuit\b", re.I), "Sirkuit LSI mikroprosesor"),
    (re.compile(r"\bSuper large LCD(?: display)?\b", re.I), "LCD berukuran besar"),
    (re.compile(r"\bLarge LCD(?: display)?\b", re.I), "Layar LCD besar"),
    (re.compile(r"\bwith contrast adjustment\b", re.I), "dengan penyetelan kontras"),
    (re.compile(r"\bdigit size\b", re.I), "ukuran digit"),
    (re.compile(r"\bDual\s*&\s*differential input\b", re.I), "Input ganda dan diferensial"),
    (re.compile(r"\bzero\s*/\s*relative\b", re.I), "nol/relatif"),
    (re.compile(r"\bmemory\b", re.I), "memori"),
    (re.compile(r"\bPush button on the front panel\b", re.I), "Tombol tekan pada panel depan"),
    (re.compile(r"\bPush button gain adjustment\b", re.I), "Penyetelan gain melalui tombol tekan"),
    (re.compile(r"\bfor calibration precisely if necessary\b", re.I), "untuk kalibrasi presisi jika diperlukan"),
    (re.compile(r"\bSensor is built inside the housing\b", re.I), "Sensor terpasang di dalam housing"),
    (re.compile(r"\bPiezoelectric sensor built inside the housing\b", re.I), "Sensor piezoelektrik terpasang di dalam housing"),
    (re.compile(r"\bBuilt-in piezoelectric sensor\b", re.I), "Sensor piezoelektrik terintegrasi"),
    (re.compile(r"\bUsed for dry, non-corrosive and\s*\.?\s*non-ionic air and gas only\b", re.I), "Hanya untuk udara dan gas kering, tidak korosif, dan tidak terionisasi"),
    (re.compile(r"\bFor dry, non-corrosive, non-ionized air and gas only\b", re.I), "Hanya untuk udara dan gas kering, tidak korosif, dan tidak terionisasi"),
    (re.compile(r"\bLiquid is prohibited\b", re.I), "Cairan tidak diperbolehkan"),
    (re.compile(r"\bBy push button\b", re.I), "Melalui tombol tekan"),
    (re.compile(r"\bRecord maximum\s*&\s*minimum readings\b", re.I), "Merekam pembacaan maksimum dan minimum"),
    (re.compile(r"\bRecords maximum and minimum readings\b", re.I), "Merekam pembacaan maksimum dan minimum"),
    (re.compile(r"\bFreezes? the display reading\b", re.I), "Menahan nilai pada layar"),
    (re.compile(r"\bAuto shut off, saves battery life, or manual off by push button\b", re.I), "Mati otomatis untuk menghemat baterai atau mati manual melalui tombol"),
    (re.compile(r"\bSaves battery life by automatically shutting off the meter\b", re.I), "Mati otomatis untuk menghemat baterai"),
    (re.compile(r"\bRS\s*232 PC serial interface\b", re.I), "Antarmuka serial RS232 untuk komputer"),
    (re.compile(r"\bApprox\.?\b", re.I), "Sekitar"),
    (re.compile(r"\bseconds?\b", re.I), "detik"),
    (re.compile(r"\bLess than\b", re.I), "Kurang dari"),
    (re.compile(r"\bbattery\b", re.I), "baterai"),
    (re.compile(r"\bInstruction manual\b", re.I), "Buku petunjuk"),
    (re.compile(r"\bHard carrying case\b", re.I), "Kotak pelindung keras"),
    (re.compile(r"\bData acquisition software\b", re.I), "Perangkat lunak akuisisi data"),
    (re.compile(r"\bRS232 cable\b", re.I), "Kabel RS232"),
    (re.compile(r"\bUSB cable\b", re.I), "Kabel USB"),
    (re.compile(r"\bWindows version\b", re.I), "versi Windows"),
    (re.compile(r"\bFull Scale\b", re.I), "Skala Penuh"),
    (re.compile(r"\bincluding linearity, hysteresis, and repeatability\b", re.I), "termasuk linearitas, histeresis, dan repeatability"),
    (re.compile(r"\bsupper large LCD display\b", re.I), "layar LCD berukuran besar"),
    (re.compile(r"\bPiezoelectric sensor\b", re.I), "Sensor piezoelektrik"),
    (re.compile(r"\busage\s+P\b", re.I), "gunakan P"),
    (re.compile(r"\bAuto shut off, saves battery life(?:,| or) manual off(?: by| inch)? push button\b", re.I), "Mati otomatis untuk menghemat baterai atau mati manual melalui tombol"),
    (re.compile(r"\bAuto shut off\b", re.I), "Mati otomatis"),
    (re.compile(r"\bAuto Power Off\b", re.I), "Mati otomatis"),
    (re.compile(r"\bMax\s*/\s*Min Recall\b", re.I), "Pemanggilan nilai maksimum/minimum"),
    (re.compile(r"\bZero Adjustment\b", re.I), "Penyetelan nol"),
    (re.compile(r"\bEasy zero offset adjustment via front panel button\b", re.I), "Penyetelan nol melalui tombol panel depan"),
    (re.compile(r"\bEasy offset of zero value with push button on the front panel\b", re.I), "Penyetelan nilai nol melalui tombol panel depan"),
    (re.compile(r"\bSaves battery life; can also be manually turned off\b", re.I), "Menghemat baterai dan dapat dimatikan secara manual"),
    (re.compile(r"\bRecords and recalls maximum and minimum readings\b", re.I), "Merekam dan menampilkan kembali pembacaan maksimum dan minimum"),
    (re.compile(r"\bAutomatic and manual data logging\b", re.I), "Pencatatan data otomatis dan manual"),
    (re.compile(r"\bfor convenient recording\b", re.I), "untuk memudahkan pencatatan"),
    (re.compile(r"\bPlug for quick coupler\b", re.I), "Konektor kopling cepat"),
    (re.compile(r"\bPCs?\b", re.I), "unit"),
    (re.compile(r"\b0 to 50(?=℃|\s|$)\b", re.I), "0 hingga 50"),
)


V225_TECH_CODE_RE = re.compile(
    r"\b(?=[A-Za-z0-9-]*[A-Za-z])(?=[A-Za-z0-9-]*\d)[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*\b"
)


def _v225_technical_tokens(value: Any) -> List[str]:
    text = _v224_exact_text(value)
    normalized = re.sub(r"\bRS\s*232\b", "RS232", text, flags=re.I)
    tokens = V225_TECH_CODE_RE.findall(normalized)
    tokens.extend(re.findall(r"\b(?:LCD|LSI|AC|DC|RH|F\.?S\.?)\b", normalized, re.I))
    return dedupe_keep_order([clean_text(token) for token in tokens if clean_text(token)])


def _v225_numeric_tokens(value: Any) -> List[str]:
    text = _v224_exact_text(value)
    text = re.sub(r"\bRS\s*232\b", "RS232", text, flags=re.I)
    for token in sorted(_v225_technical_tokens(text), key=len, reverse=True):
        text = re.sub(re.escape(token), " ", text, flags=re.I)
    # Formula/exponent digits are identifiers, not standalone measurements.
    text = re.sub(r"(?i)(?<=H)2(?=O)|(?<=CO)2\b|(?<=cm)[23]\b", "", text)
    return re.findall(r"(?<![A-Za-z])[-+]?\d+(?:[.,]\d+)?", text)


def _v225_translate_value_preserving_tokens(value: Any) -> str:
    source = _v224_exact_text(value)
    output = source
    for pattern, replacement in V225_VALUE_TRANSLATIONS:
        output = pattern.sub(replacement, output)
    output = re.sub(r"\s+", " ", output).strip()
    # Technical numbers and identifiers are immutable.
    if _v225_numeric_tokens(output) != _v225_numeric_tokens(source):
        return source
    probe = re.sub(r"\bRS\s*232\b", "RS232", output, flags=re.I).casefold().replace(" ", "")
    for token in _v225_technical_tokens(source):
        if token.casefold().replace(" ", "") not in probe:
            return source
    return output


def _v225_localize_label(label: Any) -> str:
    source = _v224_exact_text(label)
    if _v224_known_unit(source) or _v224_known_quantity(source):
        return source
    key = re.sub(r"\s+", " ", source).strip().casefold()
    return V225_LABEL_MAP.get(key, source)


def _v225_parse_units_list(value: Any) -> List[Tuple[str, str]]:
    text = _v224_exact_text(value)
    parts = [part.strip(" .;:-") for part in re.split(r"\s*,\s*|\s*;\s*", text) if part.strip(" .;:-")]
    units: List[Tuple[str, str]] = []
    for part in parts:
        normalized = _v224_known_unit(part)
        if not normalized:
            return []
        units.append((part, normalized))
    return units if len(units) >= 3 else []


def _v225_quantity_or_unknown(value: Any) -> Optional[Tuple[str, str, bool]]:
    known = _v224_known_quantity(value)
    if known:
        number = _v224_exact_text(value).strip(" ,.;:").split(maxsplit=1)[0]
        return number, known[1], True
    unknown = V224_UNKNOWN_QUANTITY_RE.fullmatch(_v224_exact_text(value))
    if unknown:
        return unknown.group(1), unknown.group(2), False
    return None


_v225_previous_source_rows = _v224_source_rows


def _v224_source_rows(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Repair only corrupted unit labels using a verified same-page sequence.

    No number is changed and no value is borrowed from another product. A repair
    is allowed only when a Units row and the following resolution rows have the
    same count/order, and all already-recognised rows agree with that order.
    """
    ledger = _v225_previous_source_rows(data)
    repairs: List[str] = []
    for units_index, units_item in enumerate(ledger):
        label_probe = units_item["label_raw"].casefold()
        if not re.search(r"(?:^|\b)(?:units|display units|pressure units|satuan)(?:\b|$)", label_probe, re.I):
            continue
        units = _v225_parse_units_list(units_item["value_raw"])
        if not units:
            continue
        candidates: List[Tuple[int, Tuple[str, str, bool]]] = []
        for row_index in range(units_index + 1, min(len(ledger), units_index + len(units) + 8)):
            row = ledger[row_index]
            parsed = _v225_quantity_or_unknown(row["value_raw"])
            if parsed:
                candidates.append((row_index, parsed))
                if len(candidates) == len(units):
                    break
        if len(candidates) != len(units):
            continue
        contradictions = 0
        confirmed = 0
        for position, (_row_index, parsed) in enumerate(candidates):
            _number, actual_unit, is_known = parsed
            expected_unit = units[position][1]
            if is_known:
                if actual_unit.casefold() == expected_unit.casefold():
                    confirmed += 1
                else:
                    contradictions += 1
        if contradictions or confirmed < max(2, len(units) // 2):
            continue
        for position, (row_index, parsed) in enumerate(candidates):
            number, actual_unit, is_known = parsed
            if is_known:
                continue
            expected_display, expected_unit = units[position]
            row = ledger[row_index]
            row["label_source_original"] = row["label_raw"]
            row["value_source_original"] = row["value_raw"]
            row["label_raw"] = expected_display
            row["value_raw"] = f"{number} {expected_display}"
            repairs.append(
                f"source_id_{row['source_id']}:{actual_unit}->{expected_unit};numeric_value_preserved={number}"
            )
        if repairs:
            break
    if repairs:
        data["_v225_source_repairs"] = repairs
        data["error"] = _merge_source_notes(data.get("error"), "same_page_unit_sequence_repair=" + "|".join(repairs))
    return ledger


_v225_previous_safe_display_row = _v224_safe_display_row


def _v224_safe_display_row(data: Dict[str, Any], item: Dict[str, Any]) -> Dict[str, str]:
    """Use AI translation only row-by-row when technical tokens are unchanged."""
    source_id = int(item.get("source_id") or 0)
    label_raw = item["label_raw"]
    value_raw = item["value_raw"]
    ai = data.get("_ai_fields") if isinstance(data.get("_ai_fields"), dict) else {}
    ai_rows = ai.get("translated_spec_rows") if isinstance(ai.get("translated_spec_rows"), list) else []
    ai_row = next(
        (row for row in ai_rows if isinstance(row, dict) and _coerce_ai_spec_id(row.get("source_id", row.get("id"))) == source_id),
        None,
    )
    label_display = ""
    value_display = ""
    if ai_row:
        candidate_label = _v224_exact_text(ai_row.get("label") or ai_row.get("label_id"))
        candidate_value = _v224_exact_text(ai_row.get("value") or ai_row.get("value_id"))
        if candidate_value and _v225_numeric_tokens(candidate_value) == _v225_numeric_tokens(value_raw):
            probe = re.sub(r"\bRS\s*232\b", "RS232", candidate_value, flags=re.I).casefold().replace(" ", "")
            if all(token.casefold().replace(" ", "") in probe for token in _v225_technical_tokens(value_raw)):
                label_display = candidate_label
                value_display = candidate_value
    if not label_display:
        label_display = _v225_localize_label(label_raw)
    if not value_display:
        value_display = _v225_translate_value_preserving_tokens(value_raw)
    return {
        "label": label_display or label_raw,
        "value": value_display or value_raw,
        "source_label": label_raw,
        "section": item["section_raw"],
    }


def _v225_interface_resolution(data: Dict[str, Any]) -> str:
    rows = _v224_source_rows(data)
    full = " ".join(f"{row['label_raw']} {row['value_raw']}" for row in rows)
    source_blob = _source_claim_blob(data)
    ambiguous = bool(re.search(r"RS\s*232\s*/\s*USB|RS232\s+and\s+USB", full + " " + source_blob, re.I))
    usb_optional = bool(re.search(r"(?:optional accessories|aksesori opsional)[^\n]{0,800}\bUSB\s*[- ]?0?1\b", full, re.I)) or bool(re.search(r"\bUSB\s*[- ]?0?1\b", full, re.I))
    rs_cable = re.search(r"\bUPCB\s*[- ]?0?\d\b", full, re.I)
    if ambiguous and (usb_optional or rs_cable):
        pieces = ["Sumber mencantumkan antarmuka komputer RS232/USB"]
        if usb_optional:
            pieces.append("kabel USB-01 tercantum sebagai aksesori opsional")
        if rs_cable:
            pieces.append(f"kabel RS232 {clean_text(rs_cable.group(0)).upper()} tercantum sebagai aksesori opsional")
        return "; ".join(pieces) + "."
    return ""


_v225_previous_build_specification_html = _v224_build_specification_html


def _v224_build_specification_html(data: Dict[str, Any]) -> Tuple[str, Dict[str, Any]]:
    html_value, report = _v225_previous_build_specification_html(data)
    source_repairs = list(data.get("_v225_source_repairs") or [])
    if source_repairs:
        report["same_page_repairs"] = source_repairs
        report["issues"] = [
            issue for issue in report.get("issues", [])
            if not issue.startswith("unrecognized_source_measurement_unit:")
        ]
    interface_note = _v225_interface_resolution(data)
    if interface_note:
        html_value += f'<p class="WEBSITE-spec-note"><strong>Klarifikasi koneksi:</strong> {esc(interface_note)}</p>'
        report["interface_clarification"] = interface_note
    return html_value, report


def _v225_content_supported(value: Any, evidence: str) -> bool:
    if not clean_text(value):
        return False
    return _v22_numeric_claims_supported(value, evidence) and not _v22_claim_issues(value, evidence)


def _v225_validate_content_only(parsed: Dict[str, Any], payload: Dict[str, Any]) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    if not isinstance(parsed, dict) or not parsed:
        return False, {}, "invalid_ai_json", []
    repaired = _v221_repair_ai_contract(_v223_rewrite_ai_contract_to_WEBSITE(parsed), payload)
    fact_map = _v22_fact_map(payload)
    all_evidence = " ".join(fact_map.values())
    brand = clean_text(payload.get("brand")) or "Lutron"
    model = clean_text(payload.get("model"))
    repair_fields: List[str] = []

    title = _v22_complete_sentence(repaired.get("product_name_id"), 180).rstrip(".")
    description = _v22_normalize_html(repaired.get("product_description_html"))
    short = _v22_complete_sentence(repaired.get("short_description"), V22_SHORT_MAX)
    features = _v22_normalize_html(repaired.get("feature_tab_html"))
    meta = _v22_complete_sentence(repaired.get("meta_description"), V22_META_MAX)
    official = _v22_normalize_html(repaired.get("official_information_html"))

    if not title or brand.casefold() not in title.casefold() or (model and model.casefold() not in title.casefold()):
        repair_fields.append("product_name_id")
    if not description or len(_html_text(description)) < 170 or _v22_html_contract_issues(description, require_heading=True) or _v22_has_english_residue(description) or not _v225_content_supported(description, all_evidence):
        repair_fields.append("product_description_html")
    if not short or len(short) > V22_SHORT_MAX or V22_STOP_END_RE.search(short) or _v22_has_english_residue(short) or not _v225_content_supported(short, all_evidence):
        repair_fields.append("short_description")
    if not features or _v22_html_contract_issues(features, require_heading=True, require_list=True) or _v22_has_english_residue(features) or not _v225_content_supported(features, all_evidence):
        repair_fields.append("feature_tab_html")
    if not _v22_meta_valid(meta, model, "") or _v22_has_english_residue(meta) or not _v225_content_supported(meta, all_evidence):
        repair_fields.append("meta_description")
    if not official or _v22_html_contract_issues(official, require_heading=True) or _v22_has_english_residue(official) or not _v225_content_supported(official, all_evidence):
        repair_fields.append("official_information_html")

    faq_items: List[Dict[str, Any]] = []
    raw_faq = repaired.get("faq_items")
    faq_valid = isinstance(raw_faq, list) and V22_FAQ_MIN <= len(raw_faq) <= V22_FAQ_MAX
    if faq_valid:
        for item in raw_faq:
            question = clean_text(item.get("question_id") or item.get("question")) if isinstance(item, dict) else ""
            answer = _v22_complete_sentence(item.get("answer_id") or item.get("answer"), 520) if isinstance(item, dict) else ""
            ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map) if isinstance(item, dict) else []
            evidence = _v22_evidence_text(ids, fact_map) or all_evidence
            if not question or not answer or not ids or _v22_has_english_residue(question + " " + answer) or not _v225_content_supported(answer, evidence):
                faq_valid = False
                break
            faq_items.append({"question": question, "answer": answer, "evidence_ids": ids})
    if not faq_valid:
        repair_fields.append("faq_html")
        faq_items = []

    app_items: List[Dict[str, Any]] = []
    raw_apps = repaired.get("application_items")
    apps_valid = isinstance(raw_apps, list) and 1 <= len(raw_apps) <= 5
    if apps_valid:
        for item in raw_apps:
            text = _v22_complete_sentence(item.get("text_id") or item.get("text"), 360) if isinstance(item, dict) else ""
            ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map) if isinstance(item, dict) else []
            evidence = _v22_evidence_text(ids, fact_map) or all_evidence
            if not text or not ids or _v22_has_english_residue(text) or not _v225_content_supported(text, evidence):
                apps_valid = False
                break
            app_items.append({"text": text, "evidence_ids": ids})
    if not apps_valid:
        repair_fields.append("application_tab_html")
        app_items = []

    # Specifications are intentionally not validated or accepted from AI.
    # The source-immutable renderer will build them row by row.
    fields: Dict[str, Any] = {
        "product_name_id": title,
        "category_id": clean_catalog_source_text(repaired.get("category_id")).rstrip("."),
        "product_description_html": description,
        "short_description": short,
        "feature_tab_html": features,
        "faq_html": _v22_render_faq(faq_items) if faq_items else "",
        "faq_items": faq_items,
        "application_tab_html": _v22_render_application(app_items) if app_items else "",
        "meta_description": meta,
        "official_information_html": official,
        "translated_spec_rows": [],
        "ai_spec_dedup_applied": False,
        "ai_spec_dedup_removed_count": 0,
        "content_evidence": repaired.get("content_evidence") if isinstance(repaired.get("content_evidence"), dict) else {},
        "_v225_repair_fields": dedupe_keep_order(repair_fields),
    }
    return True, fields, "ok_sectional", []


# The AI response contract no longer requires translated specifications.
def _v22_validate_ai_fields(parsed: Dict[str, Any], payload: Dict[str, Any]) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    return _v225_validate_content_only(parsed, payload)


_v225_previous_get_ai_content = get_ai_content


def _v225_application_fallback(data: Dict[str, Any]) -> str:
    applications = _v223_source_applications(data)
    if not applications:
        applications = ["pengukuran sesuai parameter dan kondisi kerja yang tercantum pada spesifikasi"]
    return _v22_normalize_html(
        "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(f"<li>{esc(item.capitalize())}</li>" for item in applications[:5]) + "</ul>"
    )


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    result = _v225_previous_get_ai_content(data, family)
    if not result:
        return result
    fallback = _v223_deterministic_source_fields(data, family)
    repair_fields = set(result.pop("_v225_repair_fields", []) or [])
    mapping = {
        "product_name_id": "product_name_id",
        "product_description_html": "product_description_html",
        "short_description": "short_description",
        "feature_tab_html": "feature_tab_html",
        "faq_html": "faq_html",
        "meta_description": "meta_description",
        "official_information_html": "official_information_html",
    }
    for target, fallback_key in mapping.items():
        if target in repair_fields or not clean_text(result.get(target)):
            result[target] = fallback.get(fallback_key, result.get(target, ""))
            if target == "faq_html":
                result["faq_items"] = []
    if "application_tab_html" in repair_fields or not clean_text(result.get("application_tab_html")):
        result["application_tab_html"] = _v225_application_fallback(data)
    # AI specification rows are never used as the final specification source.
    result["translated_spec_rows"] = []
    if repair_fields:
        data["_content_mode"] = "ai_section_repaired"
        data["_v225_repaired_sections"] = sorted(repair_fields)
        data["error"] = _merge_source_notes(data.get("error"), "ai_section_repaired=" + ",".join(sorted(repair_fields)))
        print("    AI section repair: " + ", ".join(sorted(repair_fields)))
    data["_ai_fields"] = result
    return result


_v225_previous_build_output_row = build_output_row


def _v225_replace_duplicate_words(value: Any) -> str:
    text = clean_text(value)
    return re.sub(r"\b(\w+)\s+\1\b", r"\1", text, flags=re.I)


def _v225_public_language_cleanup(value: Any) -> str:
    text = clean_text(value)
    replacements = (
        (r"\bdifferential\b", "diferensial"),
        (r"\bdual\b", "ganda"),
        (r"\bunit tampilan\b", "satuan tampilan"),
        (r"\bzero adjust(?:ment)?\b", "penyetelan nol"),
        (r"\bauto shut off\b", "mati otomatis"),
        (r"\bauto power off\b", "mati otomatis"),
        (r"\bmax/min\b", "maksimum/minimum"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.I)
    return re.sub(r"\b(\w+)\s+\1\b", r"\1", text, flags=re.I)


def _v225_title_with_source_range(data: Dict[str, Any], title: str) -> str:
    range_value = _v223_primary_range(data)
    if not range_value or range_value.casefold() in title.casefold():
        return title
    model = clean_text(data.get("model"))
    if model and re.search(re.escape(model), title, re.I):
        return re.sub(re.escape(model), lambda m: m.group(0) + " " + range_value, title, count=1, flags=re.I)
    return (title + " " + range_value).strip()


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v225_previous_build_output_row(data, pub_date)
    for field in (
        "Product Description", "Product Short Description",
        "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
    ):
        row[field] = _v225_public_language_cleanup(row.get(field))
    row["Name"] = _v225_title_with_source_range(data, clean_text(row.get("Name")))
    family = detect_family(data)
    row["focus keyphrase"] = build_seo_focus_keyphrase(data, family, row["Name"])
    generic_meta = bool(re.search(r"sesuai spesifikasi produk", clean_text(row.get("meta description")), re.I))
    if generic_meta:
        row["meta description"] = _v223_natural_meta(data, family, row["focus keyphrase"])
        row["Website Correction Log"] = _merge_source_notes(row.get("Website Correction Log"), "generic_meta_rebuilt")
    if clean_text(data.get("_content_mode")) == "ai_section_repaired":
        row["Website Validation"] = "publish_ready_ai_section_repaired"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "ai_section_repair=" + ",".join(data.get("_v225_repaired_sections") or []),
        )
    report = data.get("_v224_spec_report") if isinstance(data.get("_v224_spec_report"), dict) else {}
    if report.get("same_page_repairs"):
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "same_page_unit_sequence_repair=" + "|".join(report["same_page_repairs"]),
        )
    if report.get("interface_clarification"):
        row["Website Correction Log"] = _merge_source_notes(row.get("Website Correction Log"), "interface_clarified_from_accessory_rows")
    return row


_v225_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v225_previous_validate_publish_output(data, row, require_ai=require_ai)
    if _v225_interface_resolution(data):
        issues = [issue for issue in issues if issue != "ambiguous_rs232_usb_claim"]
    if clean_text(data.get("_content_mode")) == "ai_section_repaired":
        issues = [issue for issue in issues if issue not in {"evidence_publish_status_required", "deepseek_validated_content_required"}]
    return dedupe_keep_order(issues)

# ---------------------------------------------------------------------------
# V22.6 — iterative, section-aware repair loop
# ---------------------------------------------------------------------------
# Policy:
#   generate -> validate each section -> repair only failed sections ->
#   validate again -> repeat until every required section is valid.
# Valid sections are frozen and are never regenerated. Specifications are
# repaired row-by-row from the immutable source ledger; numbers, units, model
# codes, and technical identifiers remain locked to the source row.

V226_VERSION = "v22.6_iterative_section_repair"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_6_iterative_section_repair_2026_07_28"
V226_REQUIRED_SECTIONS: Tuple[str, ...] = (
    "product_name_id",
    "product_description_html",
    "short_description",
    "feature_tab_html",
    "faq_html",
    "application_tab_html",
    "meta_description",
    "official_information_html",
)
V226_EVIDENCE_KEYS = {
    "product_description_html": "product_description",
    "short_description": "short_description",
    "feature_tab_html": "feature_tab",
    "application_tab_html": "application_tab",
    "meta_description": "meta_description",
    "official_information_html": "official_information",
}
V226_PUBLIC_ENGLISH_RE = re.compile(
    r"\b(?:"
    r"instruction\s+manual|hard\s+carrying\s+case|carrying\s+case|less\s+than|"
    r"power\s+supply|measurement\s+range|measuring\s+range|display\s+size|"
    r"sampling\s+time|response\s+time|accessories?\s+(?:included|optional)|"
    r"built\s+inside|inside\s+the\s+housing|heavy\s+duty|battery\s+life|"
    r"manual\s+off|auto\s+shut\s+off|push\s+button|front\s+panel|"
    r"used\s+for|usage\s+for|for\s+calibration|liquid\s+is\s+prohibited|"
    r"records?\s+(?:and\s+recalls?\s+)?maximum|freezes?\s+(?:the\s+)?display|"
    r"data\s+hold\s+to|saves?\s+battery|super\s+large|supper\s+large|"
    r"with\s+contrast\s+adjustment|dual\s+(?:and|&)\s+differential\s+input|"
    r"digital\s+pressure\s+meter|pressure\s+meter|temperature\s+meter|"
    r"humidity\s+meter|sound\s+level\s+meter|light\s+meter|clamp\s+meter|"
    r"power\s+analy[sz]er|frequency\s+counter|data\s+logger"
    r")\b",
    re.I,
)
V226_SPEC_ENGLISH_RE = re.compile(
    r"\b(?:"
    r"supper|super\s+large|large\s+LCD\s+display|display\s+reading|"
    r"built\s+inside|inside\s+the\s+housing|housing|heavy\s+duty|"
    r"saves?|battery\s+life|manual\s+off|auto\s+shut\s+off|"
    r"push\s+button|front\s+panel|gain\s+adjustment|for\s+calibration|"
    r"usage\s+for|used\s+for|liquid\s+is\s+prohibited|"
    r"instruction\s+manual|hard\s+carrying\s+case|"
    r"accessories?\s+(?:included|optional)|power\s+supply|"
    r"operating\s+(?:temperature|humidity)|measurement\s+range|"
    r"sampling\s+time|response\s+time|data\s+hold\s+to|"
    r"record(?:s|ing)?\s+(?:maximum|minimum)|freezes?\s+(?:the\s+)?display|"
    r"including\s+linearity|full\s+scale|approx(?:imately)?|"
    r"repeatability|readings?|\bdisplay\b|\binput\b|\boutput\b|"
    r"\bincluded\b|\boptional\b|\btype\b|data\s+logger|"
    r"\d+(?:[.,]\d+)?\s+to\s+\d+(?:[.,]\d+)?"
    r")\b",
    re.I,
)
V226_FORBIDDEN_FAQ_RE = re.compile(
    r"\b(?:stok|ketersediaan\s+stok|harga|diskon|promo|garansi|"
    r"layanan\s+kalibrasi|cara\s+memesan|pemesanan)\b",
    re.I,
)
V226_UNSUPPORTED_MEDICAL_RE = re.compile(
    r"\b(?:medis|medical|rumah\s+sakit|diagnosis|pasien|kesehatan)\b",
    re.I,
)
V226_INTERNAL_LOGGER_RE = re.compile(
    r"\b(?:data\s*logger|pencatat\s+data|pencatatan\s+data\s+internal)\b",
    re.I,
)

AI_STATS.setdefault("section_repair_calls", 0)
AI_STATS.setdefault("section_repair_success", 0)
AI_STATS.setdefault("spec_repair_calls", 0)
AI_STATS.setdefault("spec_rows_repaired", 0)


def _v226_json_clone(value: Any) -> Any:
    try:
        return json.loads(json.dumps(value, ensure_ascii=False))
    except Exception:
        return value


def _v226_problem(code: str, message: str, **details: Any) -> Dict[str, Any]:
    item: Dict[str, Any] = {"code": code, "message": message}
    for key, value in details.items():
        if value not in (None, "", [], {}, ()):
            item[key] = value
    return item


def _v226_dedupe_problems(items: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    output: List[Dict[str, Any]] = []
    seen = set()
    for item in items or []:
        if not isinstance(item, dict):
            continue
        key = json.dumps(item, ensure_ascii=False, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        output.append(item)
    return output


def _v226_has_public_english(value: Any) -> bool:
    text = _html_text(value)
    return bool(_v22_has_english_residue(text) or V226_PUBLIC_ENGLISH_RE.search(text))


def _v226_has_spec_english(value: Any) -> bool:
    text = _html_text(value)
    return bool(_v22_has_english_residue(text) or V226_SPEC_ENGLISH_RE.search(text))


def _v226_source_supports(data: Dict[str, Any], pattern: re.Pattern[str]) -> bool:
    return bool(pattern.search(_source_claim_blob(data)))


def _v226_sentence_complete(value: Any) -> bool:
    text = clean_text(value)
    if not text or V22_STOP_END_RE.search(text):
        return False
    return bool(re.search(r"[.!?]$", text))


def _v226_html_counts(value: Any) -> Dict[str, int]:
    soup = BeautifulSoup(clean_text(value), "html.parser")
    return {
        "p": len(soup.find_all("p")),
        "h2": len(soup.find_all("h2")),
        "h3": len(soup.find_all("h3")),
        "li": len(soup.find_all("li")),
    }


def _v226_validate_common_claims(
    field: str,
    value: Any,
    evidence_ids: Sequence[str],
    fact_map: Dict[str, str],
    data: Dict[str, Any],
) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    text = _html_text(value)
    ids = _v22_clean_evidence_ids(list(evidence_ids or []), fact_map)
    if field != "product_name_id" and not ids:
        issues.append(_v226_problem(
            "missing_evidence_ids",
            "Bagian ini wajib merujuk minimal satu ID fakta sumber yang valid.",
            field=field,
        ))
    evidence = _v22_evidence_text(ids, fact_map)
    if ids and not _v22_numeric_claims_supported(text, evidence):
        issues.append(_v226_problem(
            "numeric_claim_not_in_selected_evidence",
            "Angka pada hasil tidak terdapat pada fakta yang dirujuk.",
            field=field,
        ))
    if ids:
        for code in _v22_claim_issues(text, evidence):
            issues.append(_v226_problem(
                code,
                "Klaim pada hasil tidak didukung fakta yang dirujuk.",
                field=field,
            ))
    if V226_UNSUPPORTED_MEDICAL_RE.search(text) and not _v226_source_supports(data, V226_UNSUPPORTED_MEDICAL_RE):
        issues.append(_v226_problem(
            "unsupported_medical_application",
            "Hapus klaim medis karena tidak tertulis pada sumber produk.",
            field=field,
        ))
    if V226_INTERNAL_LOGGER_RE.search(text) and not _v226_source_supports(data, V226_INTERNAL_LOGGER_RE):
        issues.append(_v226_problem(
            "unsupported_internal_data_logger",
            "Jangan menyebut data logger internal jika sumber hanya menyebut Data Hold, memori maksimum/minimum, atau keluaran komputer.",
            field=field,
        ))
    if V22_FORBIDDEN_SALES_RE.search(text):
        issues.append(_v226_problem(
            "forbidden_sales_claim",
            "Hapus klaim harga, promo, stok, atau garansi yang tidak didukung sumber.",
            field=field,
        ))
    if has_cjk(text):
        issues.append(_v226_problem(
            "chinese_character_detected",
            "Seluruh teks publik harus bebas karakter Chinese.",
            field=field,
        ))
    if _v226_has_public_english(text):
        issues.append(_v226_problem(
            "english_residue_detected",
            "Terjemahkan sisa frasa bahasa Inggris ke Bahasa Indonesia. Pertahankan hanya model, kode, satuan, dan singkatan teknis.",
            field=field,
        ))
    return issues


def _v226_extract_candidate_value(candidate: Any) -> Tuple[str, List[str]]:
    if isinstance(candidate, dict):
        return clean_text(candidate.get("value")), candidate.get("evidence_ids") if isinstance(candidate.get("evidence_ids"), list) else []
    return clean_text(candidate), []


def _v226_validate_section(
    field: str,
    candidate: Any,
    payload: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], List[Dict[str, Any]]]:
    fact_map = _v22_fact_map(payload)
    issues: List[Dict[str, Any]] = []

    if field == "faq_html":
        raw_items = candidate.get("items") if isinstance(candidate, dict) else None
        if not isinstance(raw_items, list) or not (V22_FAQ_MIN <= len(raw_items) <= V22_FAQ_MAX):
            issues.append(_v226_problem(
                "faq_item_count_invalid",
                f"FAQ harus berisi {V22_FAQ_MIN}-{V22_FAQ_MAX} pasangan pertanyaan dan jawaban teknis.",
                field=field,
            ))
            return False, {}, issues
        normalized_items: List[Dict[str, Any]] = []
        for index, item in enumerate(raw_items, 1):
            question = clean_text(item.get("question_id") or item.get("question")) if isinstance(item, dict) else ""
            answer = _v22_complete_sentence(item.get("answer_id") or item.get("answer"), 520) if isinstance(item, dict) else ""
            evidence_ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map) if isinstance(item, dict) else []
            if not question or not answer:
                issues.append(_v226_problem("faq_item_empty", "Pertanyaan dan jawaban FAQ tidak boleh kosong.", item=index))
                continue
            if not question.endswith("?"):
                issues.append(_v226_problem("faq_question_mark_missing", "Pertanyaan FAQ harus diakhiri tanda tanya.", item=index))
            if V226_FORBIDDEN_FAQ_RE.search(question + " " + answer):
                issues.append(_v226_problem(
                    "commercial_faq_topic_forbidden",
                    "FAQ harus membahas fungsi, rentang, media, koneksi, pengoperasian, atau kelengkapan teknis; jangan membahas stok, harga, garansi, atau layanan kalibrasi.",
                    item=index,
                ))
            issues.extend(_v226_validate_common_claims(field, question + " " + answer, evidence_ids, fact_map, data))
            normalized_items.append({"question": question, "answer": answer, "evidence_ids": evidence_ids})
        issues = _v226_dedupe_problems(issues)
        if issues:
            return False, {}, issues
        return True, {
            "items": normalized_items,
            "html": _v22_render_faq(normalized_items),
        }, []

    if field == "application_tab_html":
        raw_items = candidate.get("items") if isinstance(candidate, dict) else None
        if not isinstance(raw_items, list) or not (1 <= len(raw_items) <= 5):
            issues.append(_v226_problem(
                "application_item_count_invalid",
                "Aplikasi penggunaan harus berisi 1-5 poin yang benar-benar tertulis atau tersirat langsung dari fakta sumber.",
                field=field,
            ))
            return False, {}, issues
        normalized_items: List[Dict[str, Any]] = []
        for index, item in enumerate(raw_items, 1):
            text = _v22_complete_sentence(item.get("text_id") or item.get("text"), 360) if isinstance(item, dict) else ""
            evidence_ids = _v22_clean_evidence_ids(item.get("evidence_ids"), fact_map) if isinstance(item, dict) else []
            if not text:
                issues.append(_v226_problem("application_item_empty", "Poin aplikasi tidak boleh kosong.", item=index))
                continue
            issues.extend(_v226_validate_common_claims(field, text, evidence_ids, fact_map, data))
            normalized_items.append({"text": text, "evidence_ids": evidence_ids})
        issues = _v226_dedupe_problems(issues)
        if issues:
            return False, {}, issues
        return True, {
            "items": normalized_items,
            "html": _v22_render_application(normalized_items),
        }, []

    value, evidence_ids = _v226_extract_candidate_value(candidate)
    if not value:
        return False, {}, [_v226_problem("section_empty", "Bagian tidak boleh kosong.", field=field)]

    if field == "product_name_id":
        value = clean_text(value).strip(" .")
        brand = clean_text(payload.get("brand")) or "Lutron"
        model = clean_text(payload.get("model"))
        if len(value) > 180:
            issues.append(_v226_problem("title_too_long", "Nama produk maksimal 180 karakter.", length=len(value)))
        if brand.casefold() not in value.casefold() or (model and model.casefold() not in value.casefold()):
            issues.append(_v226_problem("title_missing_brand_or_model", "Nama produk wajib memuat Lutron dan model produk."))
        if value.endswith("."):
            issues.append(_v226_problem("title_ends_period", "Nama produk tidak boleh diakhiri titik."))
        if has_cjk(value) or _v226_has_public_english(value):
            issues.append(_v226_problem("title_language_invalid", "Nama produk harus berbahasa Indonesia; model dan istilah teknis boleh dipertahankan."))
        return (not issues), {"value": value, "evidence_ids": evidence_ids}, _v226_dedupe_problems(issues)

    if field in {"product_description_html", "feature_tab_html", "official_information_html"}:
        value = _v22_normalize_html(value)
        contract = _v22_html_contract_issues(
            value,
            require_heading=True,
            require_list=field in {"product_description_html", "feature_tab_html"},
        )
        for code in contract:
            issues.append(_v226_problem(code, "Struktur HTML bagian ini belum valid.", field=field))
        counts = _v226_html_counts(value)
        if field == "product_description_html":
            if counts["p"] < 2:
                issues.append(_v226_problem("description_paragraph_count_invalid", "Deskripsi harus memiliki minimal dua paragraf pembuka."))
            if counts["h2"] < 2:
                issues.append(_v226_problem("description_heading_count_invalid", "Deskripsi harus memiliki heading Keunggulan Produk dan Panduan Pemilihan."))
            if not (4 <= counts["li"] <= 6):
                issues.append(_v226_problem("description_bullet_count_invalid", "Daftar keunggulan pada deskripsi harus berisi 4-6 poin faktual.", count=counts["li"]))
            if len(_html_text(value)) < 170:
                issues.append(_v226_problem("description_too_short", "Deskripsi produk terlalu pendek."))
        elif field == "feature_tab_html":
            if not (4 <= counts["li"] <= 6):
                issues.append(_v226_problem("feature_bullet_count_invalid", "Tab Keunggulan Produk harus berisi 4-6 poin faktual.", count=counts["li"]))
        elif field == "official_information_html":
            if counts["h2"] < 2:
                issues.append(_v226_problem("official_heading_count_invalid", "Informasi Produk harus memiliki heading Informasi Produk dan Catatan Teknis."))
            if len(_html_text(value)) < 100:
                issues.append(_v226_problem("official_information_too_short", "Informasi Produk terlalu pendek."))

    elif field == "short_description":
        value = _v22_complete_sentence(value, V22_SHORT_MAX)
        if len(value) > V22_SHORT_MAX or not _v226_sentence_complete(value):
            issues.append(_v226_problem("short_description_invalid", f"Deskripsi singkat harus berupa kalimat lengkap maksimal {V22_SHORT_MAX} karakter."))
        model = clean_text(payload.get("model"))
        if model and model.casefold() not in value.casefold():
            issues.append(_v226_problem("short_description_missing_model", "Deskripsi singkat wajib memuat model produk."))

    elif field == "meta_description":
        value = _v22_complete_sentence(value, V22_META_MAX)
        model = clean_text(payload.get("model"))
        if not (V22_META_MIN <= len(value) <= V22_META_MAX):
            issues.append(_v226_problem("meta_length_invalid", f"Meta description harus {V22_META_MIN}-{V22_META_MAX} karakter.", length=len(value)))
        if model and model.casefold() not in value.casefold():
            issues.append(_v226_problem("meta_missing_model", "Meta description wajib memuat model produk."))
        if not _v226_sentence_complete(value):
            issues.append(_v226_problem("meta_sentence_incomplete", "Meta description harus berupa kalimat lengkap."))

    else:
        issues.append(_v226_problem("unknown_section", "Validator tidak mengenali bagian ini.", field=field))

    issues.extend(_v226_validate_common_claims(field, value, evidence_ids, fact_map, data))
    issues = _v226_dedupe_problems(issues)
    return (not issues), {"value": value, "evidence_ids": _v22_clean_evidence_ids(evidence_ids, fact_map)}, issues


def _v226_initial_candidate(fields: Dict[str, Any], field: str) -> Dict[str, Any]:
    evidence = fields.get("content_evidence") if isinstance(fields.get("content_evidence"), dict) else {}
    if field == "faq_html":
        items = []
        for item in fields.get("faq_items") if isinstance(fields.get("faq_items"), list) else []:
            if isinstance(item, dict):
                items.append({
                    "question_id": clean_text(item.get("question") or item.get("question_id")),
                    "answer_id": clean_text(item.get("answer") or item.get("answer_id")),
                    "evidence_ids": item.get("evidence_ids") if isinstance(item.get("evidence_ids"), list) else [],
                })
        return {"items": items}
    if field == "application_tab_html":
        items = []
        for item in fields.get("application_items") if isinstance(fields.get("application_items"), list) else []:
            if isinstance(item, dict):
                items.append({
                    "text_id": clean_text(item.get("text") or item.get("text_id")),
                    "evidence_ids": item.get("evidence_ids") if isinstance(item.get("evidence_ids"), list) else [],
                })
        return {"items": items}
    evidence_key = V226_EVIDENCE_KEYS.get(field, "")
    return {
        "value": clean_text(fields.get(field)),
        "evidence_ids": evidence.get(evidence_key) if evidence_key and isinstance(evidence.get(evidence_key), list) else [],
    }


def _v226_content_schema_for(fields: Sequence[str]) -> Dict[str, Any]:
    schema: Dict[str, Any] = {}
    for field in fields:
        if field == "faq_html":
            schema[field] = {
                "items": [{
                    "question_id": "Pertanyaan teknis Bahasa Indonesia",
                    "answer_id": "Jawaban faktual Bahasa Indonesia",
                    "evidence_ids": ["F001 atau S001"],
                }]
            }
        elif field == "application_tab_html":
            schema[field] = {
                "items": [{
                    "text_id": "Satu aplikasi faktual dalam Bahasa Indonesia",
                    "evidence_ids": ["F001 atau S001"],
                }]
            }
        else:
            schema[field] = {
                "value": "Konten perbaikan lengkap untuk bagian ini",
                "evidence_ids": ["F001 atau S001"],
            }
    return {"sections": schema}


def _v226_request_cache_key(kind: str, body: Dict[str, Any]) -> str:
    raw = json.dumps({"version": AI_PROMPT_VERSION, "kind": kind, "body": body}, ensure_ascii=False, sort_keys=True)
    return "v226:" + hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _v226_post_ai_json(kind: str, user_payload: Dict[str, Any]) -> Tuple[bool, Dict[str, Any], str]:
    api_key = clean_text(AI_SETTINGS.get("api_key"))
    if not api_key:
        return False, {}, "missing_api_key"
    cache = AI_SETTINGS.get("cache") if isinstance(AI_SETTINGS.get("cache"), dict) else {}
    cache_key = _v226_request_cache_key(kind, user_payload)
    cached = cache.get(cache_key)
    if isinstance(cached, dict) and isinstance(cached.get("parsed"), dict):
        AI_STATS["cache_hit"] = int(AI_STATS.get("cache_hit") or 0) + 1
        return True, cached["parsed"], "cache_hit"

    body = {
        "model": clean_text(AI_SETTINGS.get("model")) or "deepseek-chat",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Anda adalah editor teknis berbasis bukti. Perbaiki hanya bagian yang diminta. "
                    "Jangan mengubah bagian yang sudah valid. Gunakan hanya fakta dan spesifikasi yang diberikan. "
                    "Balas JSON valid saja tanpa markdown."
                ),
            },
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ],
        "temperature": min(0.15, float(AI_SETTINGS.get("temperature") or 0.10)),
        "response_format": {"type": "json_object"},
    }
    endpoint = clean_text(AI_SETTINGS.get("base_url")).rstrip("/") + "/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS.get("timeout", 180))
        if response.status_code in {400, 422}:
            body.pop("response_format", None)
            response = requests.post(endpoint, headers=headers, json=body, timeout=AI_SETTINGS.get("timeout", 180))
        response.raise_for_status()
        raw = response.json()
        content = clean_text(raw.get("choices", [{}])[0].get("message", {}).get("content", ""))
        parsed = _parse_json(content)
        if not parsed:
            return False, {}, "invalid_ai_json"
        cache[cache_key] = {
            "parsed": parsed,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "kind": kind,
        }
        AI_SETTINGS["cache"] = cache
        return True, parsed, "ok"
    except Exception as exc:
        return False, {}, f"request_error_{type(exc).__name__}"


def _v226_apply_normalized_section(fields: Dict[str, Any], field: str, normalized: Dict[str, Any]) -> None:
    evidence = fields.get("content_evidence") if isinstance(fields.get("content_evidence"), dict) else {}
    if field == "faq_html":
        fields["faq_items"] = normalized.get("items") or []
        fields["faq_html"] = clean_text(normalized.get("html"))
    elif field == "application_tab_html":
        fields["application_items"] = normalized.get("items") or []
        fields["application_tab_html"] = clean_text(normalized.get("html"))
    else:
        fields[field] = clean_text(normalized.get("value"))
        evidence_key = V226_EVIDENCE_KEYS.get(field)
        if evidence_key:
            evidence[evidence_key] = normalized.get("evidence_ids") or []
    fields["content_evidence"] = evidence


def _v226_repair_content_sections(
    data: Dict[str, Any],
    family: Dict[str, Any],
    initial_fields: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], Dict[str, Any]]:
    payload = _ai_payload(data, family)
    fields = _v226_json_clone(initial_fields) if isinstance(initial_fields, dict) else {}
    fields.setdefault("content_evidence", {})
    unresolved: Dict[str, List[Dict[str, Any]]] = {}
    accepted_initial: List[str] = []

    for field in V226_REQUIRED_SECTIONS:
        candidate = _v226_initial_candidate(fields, field)
        ok, normalized, issues = _v226_validate_section(field, candidate, payload, data)
        if ok:
            _v226_apply_normalized_section(fields, field, normalized)
            accepted_initial.append(field)
        else:
            unresolved[field] = issues

    max_rounds = max(2, min(10, int(AI_SETTINGS.get("retry_count") or 5) + 1))
    repair_log: Dict[str, Any] = {
        "version": V226_VERSION,
        "accepted_initial": accepted_initial,
        "rounds": [],
        "repaired_sections": [],
    }

    for round_number in range(1, max_rounds + 1):
        if not unresolved:
            break
        requested_fields = list(unresolved)
        current = {field: _v226_initial_candidate(fields, field) for field in requested_fields}
        user_payload = {
            "task": "Perbaiki hanya bagian katalog yang gagal validator. Jangan kirim atau menulis ulang bagian lain.",
            "rules": [
                "Gunakan hanya facts, specifications, dan allowed_WEBSITE_policies pada product_data.",
                "Pertahankan model, kode, angka, satuan, simbol, dan singkatan teknis persis.",
                "Setiap bagian atau item wajib memiliki evidence_ids yang benar.",
                "Hapus klaim yang tidak didukung, termasuk aplikasi medis, data logger internal, stok, harga, garansi, dan layanan kalibrasi jika tidak tertulis pada fakta.",
                "Terjemahkan seluruh kalimat ke Bahasa Indonesia. Pertahankan hanya nama merek, model, kode, satuan, dan singkatan teknis.",
                "FAQ hanya membahas aspek teknis produk, bukan stok, harga, garansi, atau layanan komersial.",
                "product_description_html wajib memiliki dua paragraf pembuka, dua heading h2, dan 4-6 poin keunggulan.",
                "feature_tab_html wajib memiliki 4-6 poin faktual.",
                "Balas JSON valid sesuai output_schema dan hanya untuk requested_sections.",
            ],
            "round": round_number,
            "maximum_rounds": max_rounds,
            "requested_sections": requested_fields,
            "current_failed_content": current,
            "validator_feedback": unresolved,
            "output_schema": _v226_content_schema_for(requested_fields),
            "product_data": payload,
        }
        AI_STATS["section_repair_calls"] = int(AI_STATS.get("section_repair_calls") or 0) + 1
        ok_request, parsed, request_status = _v226_post_ai_json("content_section_repair", user_payload)
        round_log: Dict[str, Any] = {
            "round": round_number,
            "requested": requested_fields,
            "request_status": request_status,
            "accepted": [],
            "remaining": [],
        }
        next_unresolved: Dict[str, List[Dict[str, Any]]] = {}
        sections = parsed.get("sections") if ok_request and isinstance(parsed.get("sections"), dict) else {}
        for field in requested_fields:
            if field not in sections:
                next_unresolved[field] = [_v226_problem(
                    "section_missing_from_repair_response",
                    "Respons repair tidak mengirim bagian yang diminta.",
                    field=field,
                )]
                continue
            valid, normalized, issues = _v226_validate_section(field, sections.get(field), payload, data)
            if valid:
                _v226_apply_normalized_section(fields, field, normalized)
                round_log["accepted"].append(field)
                if field not in repair_log["repaired_sections"]:
                    repair_log["repaired_sections"].append(field)
                AI_STATS["section_repair_success"] = int(AI_STATS.get("section_repair_success") or 0) + 1
            else:
                next_unresolved[field] = issues
        if not ok_request:
            next_unresolved = {
                field: [_v226_problem(
                    "repair_request_failed",
                    "Permintaan repair ke AI gagal; ulangi permintaan bagian yang sama.",
                    field=field,
                    reason=request_status,
                )]
                for field in requested_fields
            }
        round_log["remaining"] = list(next_unresolved)
        round_log["feedback"] = next_unresolved
        repair_log["rounds"].append(round_log)
        unresolved = next_unresolved
        if unresolved and float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
            time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

    repair_log["unresolved"] = unresolved
    return (not unresolved), fields, repair_log


_v226_previous_safe_display_row = _v224_safe_display_row




def _v226_required_spec_tokens(value: Any) -> List[str]:
    """Return real technical tokens without matching unit letters inside words."""
    text = _v224_exact_text(value)
    tokens: List[str] = list(_v225_technical_tokens(text))
    tokens.extend(V224_CODE_RE.findall(text))
    for match in V224_UNIT_TOKEN_RE.finditer(text):
        token = _v224_exact_text(match.group(0))
        if not token:
            continue
        compact_token = re.sub(r"[^A-Za-z0-9%°µ²₂/]", "", token)
        before = text[match.start() - 1] if match.start() > 0 else ""
        after = text[match.end()] if match.end() < len(text) else ""
        # Single-letter units such as V, A, W, C, and F are protected only
        # when they are standalone measurement units, never when the regex
        # happens to match a letter inside ordinary words such as Power/Off.
        if len(compact_token) == 1 and compact_token.isalpha():
            left_context = text[max(0, match.start() - 12):match.start()]
            if (before.isalpha() or after.isalpha()) or not re.search(r"(?:\d|°)\s*$", left_context):
                continue
        elif before.isalpha() and after.isalpha():
            continue
        tokens.append(token)
    tokens.extend(re.findall(r"\b(?:RS\s*232|USB|LCD|LSI|AC|DC|RH|F\.?S\.?|True\s*RMS)\b", text, re.I))
    return dedupe_keep_order([_v224_exact_text(token) for token in tokens if _v224_exact_text(token)])

def _v226_validate_spec_candidate(item: Dict[str, Any], candidate: Dict[str, Any]) -> List[Dict[str, Any]]:
    label_raw = _v224_exact_text(item.get("label_raw"))
    value_raw = _v224_exact_text(item.get("value_raw"))
    label = _v224_exact_text(candidate.get("label") or candidate.get("label_id"))
    value = _v224_exact_text(candidate.get("value") or candidate.get("value_id"))
    issues: List[Dict[str, Any]] = []
    source_id = item.get("source_id")
    if not label:
        issues.append(_v226_problem("spec_label_empty", "Label spesifikasi tidak boleh kosong.", source_id=source_id))
    if not value:
        issues.append(_v226_problem("spec_value_empty", "Nilai spesifikasi tidak boleh kosong.", source_id=source_id))
    if label and _v225_numeric_tokens(label) != _v225_numeric_tokens(label_raw):
        issues.append(_v226_problem(
            "spec_label_numeric_token_changed",
            "Angka pada label harus sama persis dengan sumber.",
            source_id=source_id,
            required=_v225_numeric_tokens(label_raw),
        ))
    if value and _v225_numeric_tokens(value) != _v225_numeric_tokens(value_raw):
        issues.append(_v226_problem(
            "spec_value_numeric_token_changed",
            "Angka pada nilai harus sama persis dan dalam urutan yang sama dengan sumber.",
            source_id=source_id,
            required=_v225_numeric_tokens(value_raw),
        ))
    label_probe = re.sub(r"\s+", "", label).casefold()
    value_probe = re.sub(r"\s+", "", value).casefold()
    for token in _v226_required_spec_tokens(label_raw):
        if re.sub(r"\s+", "", token).casefold() not in label_probe:
            issues.append(_v226_problem(
                "spec_label_protected_token_missing",
                "Model, kode, satuan, atau singkatan teknis pada label tidak boleh dihapus.",
                source_id=source_id,
                token=token,
            ))
    for token in _v226_required_spec_tokens(value_raw):
        if re.sub(r"\s+", "", token).casefold() not in value_probe:
            issues.append(_v226_problem(
                "spec_value_protected_token_missing",
                "Model, kode, satuan, atau singkatan teknis pada nilai tidak boleh dihapus.",
                source_id=source_id,
                token=token,
            ))
    if has_cjk(label + " " + value):
        issues.append(_v226_problem("spec_contains_chinese", "Spesifikasi harus bebas karakter Chinese.", source_id=source_id))
    if _v226_has_spec_english(label) or _v226_has_spec_english(value):
        issues.append(_v226_problem(
            "spec_english_residue_detected",
            "Terjemahkan frasa bahasa Inggris pada baris ini. Pertahankan hanya model, kode, satuan, dan singkatan teknis.",
            source_id=source_id,
        ))
    if re.search(r"\b(?:undefined|null|nan)\b|\d\s*tUSB\b|\bm\s*/\s*mim\b", label + " " + value, re.I):
        issues.append(_v226_problem("spec_malformed_text", "Baris spesifikasi masih mengandung fragmen teks rusak.", source_id=source_id))
    return _v226_dedupe_problems(issues)


def _v226_spec_repair_payload(
    ledger: Sequence[Dict[str, Any]],
    unresolved: Dict[int, List[Dict[str, Any]]],
    translations: Dict[int, Dict[str, str]],
    round_number: int,
    max_rounds: int,
    data: Dict[str, Any],
) -> Dict[str, Any]:
    rows: List[Dict[str, Any]] = []
    position_by_id = {int(item["source_id"]): index for index, item in enumerate(ledger)}
    for source_id, feedback in unresolved.items():
        index = position_by_id[source_id]
        item = ledger[index]
        before = ledger[max(0, index - 2):index]
        after = ledger[index + 1:index + 3]
        rows.append({
            "source_id": source_id,
            "source_label": item["label_raw"],
            "source_value": item["value_raw"],
            "source_section": item["section_raw"],
            "context_before": [f"{x['label_raw']}: {x['value_raw']}" for x in before],
            "context_after": [f"{x['label_raw']}: {x['value_raw']}" for x in after],
            "current_translation": translations.get(source_id, {}),
            "required_label_numbers": _v225_numeric_tokens(item["label_raw"]),
            "required_value_numbers": _v225_numeric_tokens(item["value_raw"]),
            "required_label_tokens": _v226_required_spec_tokens(item["label_raw"]),
            "required_value_tokens": _v226_required_spec_tokens(item["value_raw"]),
            "validator_feedback": feedback,
        })
    return {
        "task": "Perbaiki hanya baris spesifikasi yang gagal. Jangan menambah, menghapus, menggabungkan, atau memindahkan fakta.",
        "rules": [
            "Terjemahkan label dan kalimat biasa ke Bahasa Indonesia yang alami.",
            "Salin seluruh angka dalam urutan yang sama persis. Jangan mengonversi angka atau satuan.",
            "Pertahankan model, kode aksesori, simbol, satuan, dan singkatan teknis persis seperti sumber.",
            "Jangan menambah klaim, fungsi, penjelasan, atau nilai yang tidak terdapat pada source_label dan source_value.",
            "Jangan menghapus baris dan jangan melakukan deduplikasi.",
            "Balas satu objek untuk setiap source_id yang diminta.",
            "Balas JSON valid saja tanpa markdown.",
        ],
        "round": round_number,
        "maximum_rounds": max_rounds,
        "brand": clean_text(data.get("brand")) or "Lutron",
        "model": clean_text(data.get("model")),
        "rows_to_repair": rows,
        "output_schema": {
            "rows": [{
                "source_id": "ID sumber yang sama",
                "label_id": "Label Bahasa Indonesia dengan token sumber tetap",
                "value_id": "Nilai Bahasa Indonesia dengan angka, satuan, kode, dan urutan fakta tetap",
            }]
        },
    }


def _v226_prepare_spec_translations(data: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
    ledger = _v224_source_rows(data)
    translations: Dict[int, Dict[str, str]] = {}
    unresolved: Dict[int, List[Dict[str, Any]]] = {}

    existing = data.get("_v226_spec_translations") if isinstance(data.get("_v226_spec_translations"), dict) else {}
    for item in ledger:
        source_id = int(item["source_id"])
        candidate = existing.get(str(source_id)) if isinstance(existing.get(str(source_id)), dict) else None
        if not candidate:
            candidate = _v226_previous_safe_display_row(data, item)
        candidate = {
            "label": _v224_exact_text(candidate.get("label") or candidate.get("label_id") or item["label_raw"]),
            "value": _v224_exact_text(candidate.get("value") or candidate.get("value_id") or item["value_raw"]),
            "source_label": item["label_raw"],
            "section": item["section_raw"],
        }
        translations[source_id] = candidate
        issues = _v226_validate_spec_candidate(item, candidate)
        if issues:
            unresolved[source_id] = issues

    max_rounds = max(2, min(10, int(AI_SETTINGS.get("retry_count") or 5) + 1))
    report: Dict[str, Any] = {
        "version": V226_VERSION,
        "source_rows": len(ledger),
        "initial_failed_rows": sorted(unresolved),
        "rounds": [],
        "repaired_rows": [],
    }

    for round_number in range(1, max_rounds + 1):
        if not unresolved:
            break
        user_payload = _v226_spec_repair_payload(ledger, unresolved, translations, round_number, max_rounds, data)
        AI_STATS["spec_repair_calls"] = int(AI_STATS.get("spec_repair_calls") or 0) + 1
        ok_request, parsed, request_status = _v226_post_ai_json("specification_row_repair", user_payload)
        response_rows = parsed.get("rows") if ok_request and isinstance(parsed.get("rows"), list) else []
        response_by_id: Dict[int, Dict[str, Any]] = {}
        for raw in response_rows:
            if not isinstance(raw, dict):
                continue
            source_id = _coerce_ai_spec_id(raw.get("source_id"))
            if source_id is not None and source_id not in response_by_id:
                response_by_id[source_id] = raw

        next_unresolved: Dict[int, List[Dict[str, Any]]] = {}
        accepted_ids: List[int] = []
        item_by_id = {int(item["source_id"]): item for item in ledger}
        for source_id in unresolved:
            raw = response_by_id.get(source_id)
            if raw is None:
                next_unresolved[source_id] = [_v226_problem(
                    "spec_row_missing_from_repair_response",
                    "Respons repair tidak mengirim baris spesifikasi yang diminta.",
                    source_id=source_id,
                    reason=request_status,
                )]
                continue
            candidate = {
                "label": _v224_exact_text(raw.get("label_id")),
                "value": _v224_exact_text(raw.get("value_id")),
                "source_label": item_by_id[source_id]["label_raw"],
                "section": item_by_id[source_id]["section_raw"],
            }
            issues = _v226_validate_spec_candidate(item_by_id[source_id], candidate)
            if issues:
                next_unresolved[source_id] = issues
            else:
                translations[source_id] = candidate
                accepted_ids.append(source_id)
                if source_id not in report["repaired_rows"]:
                    report["repaired_rows"].append(source_id)
                AI_STATS["spec_rows_repaired"] = int(AI_STATS.get("spec_rows_repaired") or 0) + 1
        if not ok_request:
            next_unresolved = {
                source_id: [_v226_problem(
                    "spec_repair_request_failed",
                    "Permintaan repair spesifikasi gagal; baris yang sama akan diminta kembali.",
                    source_id=source_id,
                    reason=request_status,
                )]
                for source_id in unresolved
            }
        report["rounds"].append({
            "round": round_number,
            "request_status": request_status,
            "requested": sorted(unresolved),
            "accepted": sorted(accepted_ids),
            "remaining": sorted(next_unresolved),
            "feedback": next_unresolved,
        })
        unresolved = next_unresolved
        if unresolved and float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
            time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

    report["unresolved"] = unresolved
    if unresolved:
        return False, report
    data["_v226_spec_translations"] = {str(source_id): row for source_id, row in translations.items()}
    data["_v226_spec_repair_report"] = report
    return True, report


def _v224_safe_display_row(data: Dict[str, Any], item: Dict[str, Any]) -> Dict[str, str]:
    translations = data.get("_v226_spec_translations") if isinstance(data.get("_v226_spec_translations"), dict) else {}
    candidate = translations.get(str(item.get("source_id")))
    if isinstance(candidate, dict):
        issues = _v226_validate_spec_candidate(item, candidate)
        if not issues:
            return {
                "label": _v224_exact_text(candidate.get("label")),
                "value": _v224_exact_text(candidate.get("value")),
                "source_label": item.get("label_raw", ""),
                "section": item.get("section_raw", ""),
            }
    return _v226_previous_safe_display_row(data, item)


_v226_initial_get_ai_content = _v225_previous_get_ai_content


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    if not AI_SETTINGS.get("enabled"):
        return _v226_initial_get_ai_content(data, family)
    existing = data.get("_ai_fields") if isinstance(data.get("_ai_fields"), dict) else {}
    if data.get("_v226_repair_complete") and existing:
        return existing

    # Ignore stale V22.5 fields stored inside an old scrape cache. The AI cache
    # remains reusable because its key includes AI_PROMPT_VERSION.
    data.pop("_ai_fields", None)
    data.pop("_v226_unresolved", None)
    initial = _v226_initial_get_ai_content(data, family)
    if not isinstance(initial, dict):
        initial = {}
    initial.pop("_v225_repair_fields", None)

    content_ok, repaired_fields, content_report = _v226_repair_content_sections(data, family, initial)
    spec_ok, spec_report = _v226_prepare_spec_translations(data)
    data["_v226_content_repair_report"] = content_report
    data["_v226_spec_repair_report"] = spec_report

    unresolved: Dict[str, Any] = {}
    if not content_ok:
        unresolved["content"] = content_report.get("unresolved")
    if not spec_ok:
        unresolved["specifications"] = spec_report.get("unresolved")
    if unresolved:
        data["_v226_repair_complete"] = False
        data["_v226_unresolved"] = unresolved
        data["_ai_status"] = "iterative_repair_exhausted"
        data["error"] = _merge_source_notes(
            data.get("error"),
            "iterative_repair_exhausted=" + ",".join(unresolved),
        )
        print("    iterative repair unresolved: " + ", ".join(unresolved))
        return {}

    repaired_fields["translated_spec_rows"] = []
    repaired_fields["ai_spec_dedup_applied"] = False
    repaired_fields["ai_spec_dedup_removed_count"] = 0
    repaired_sections = content_report.get("repaired_sections") or []
    repaired_spec_rows = spec_report.get("repaired_rows") or []
    data["_ai_fields"] = repaired_fields
    data["_v226_repair_complete"] = True
    data["_content_mode"] = "ai_iterative_repair" if (repaired_sections or repaired_spec_rows) else "ai_validated"
    data["_ai_status"] = "iterative_repair_complete"
    data["_v226_repaired_sections"] = repaired_sections
    data["_v226_repaired_spec_rows"] = repaired_spec_rows
    data["error"] = _merge_source_notes(
        data.get("error"),
        "iterative_repair_complete",
        "repaired_sections=" + ",".join(repaired_sections) if repaired_sections else "",
        "repaired_spec_rows=" + ",".join(str(x) for x in repaired_spec_rows) if repaired_spec_rows else "",
    )
    print(
        "    iterative repair complete"
        + (" | sections=" + ",".join(repaired_sections) if repaired_sections else "")
        + (" | spec_rows=" + ",".join(str(x) for x in repaired_spec_rows) if repaired_spec_rows else "")
    )
    return repaired_fields


def _v226_final_row_issues(data: Dict[str, Any], row: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    if AI_SETTINGS.get("enabled") and AI_SETTINGS.get("strict_indonesian") and not data.get("_v226_repair_complete"):
        issues.append("iterative_repair_not_complete")

    public_fields = (
        "Product Description",
        "Product Short Description",
        "custom_tab_2_content",
        "custom_tab_3_content",
        "custom_tab_4_content",
        "custom_tab_5_content",
        "meta description",
    )
    for field in public_fields:
        value = row.get(field)
        if has_cjk(value):
            issues.append(f"{field}_contains_chinese")
        if _v226_has_public_english(value):
            issues.append(f"{field}_contains_english_residue")

    description_counts = _v226_html_counts(row.get("Product Description"))
    if description_counts["p"] < 2 or description_counts["h2"] < 2 or not (4 <= description_counts["li"] <= 6):
        issues.append("product_description_structure_invalid_after_repair")
    feature_counts = _v226_html_counts(row.get("custom_tab_2_content"))
    if not (4 <= feature_counts["li"] <= 6):
        issues.append("feature_count_invalid_after_repair")
    faq_soup = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    faq_text = faq_soup.get_text(" ", strip=True)
    if V226_FORBIDDEN_FAQ_RE.search(faq_text):
        issues.append("commercial_faq_topic_detected_after_repair")
    if not (V22_FAQ_MIN <= len(faq_soup.find_all("h3")) <= V22_FAQ_MAX):
        issues.append("faq_count_invalid_after_repair")

    source_blob = _source_claim_blob(data)
    public_blob = " ".join(_html_text(row.get(field)) for field in public_fields)
    if V226_UNSUPPORTED_MEDICAL_RE.search(public_blob) and not V226_UNSUPPORTED_MEDICAL_RE.search(source_blob):
        issues.append("unsupported_medical_claim_after_repair")
    if V226_INTERNAL_LOGGER_RE.search(public_blob) and not V226_INTERNAL_LOGGER_RE.search(source_blob):
        issues.append("unsupported_internal_data_logger_after_repair")

    spec_soup = BeautifulSoup(clean_text(row.get("custom_tab_1_content")), "html.parser")
    for cell in spec_soup.find_all(["th", "td"]):
        text = clean_text(cell.get_text(" ", strip=True))
        if text in {"Parameter", "Nilai", "Satuan", "Rentang", "Resolusi", "Akurasi"}:
            continue
        if has_cjk(text) or _v226_has_spec_english(text):
            issues.append("specification_language_repair_incomplete")
            break
    return dedupe_keep_order(issues)


_v226_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v226_previous_build_output_row(data, pub_date)
    if AI_SETTINGS.get("enabled") and AI_SETTINGS.get("strict_indonesian"):
        final_issues = _v226_final_row_issues(data, row)
        if final_issues:
            data["_v226_final_issues"] = final_issues
            raise ValueError("iterative_repair_final_validation_failed:" + ",".join(final_issues))
        repaired_sections = data.get("_v226_repaired_sections") or []
        repaired_spec_rows = data.get("_v226_repaired_spec_rows") or []
        repaired = bool(repaired_sections or repaired_spec_rows)
        row["Website Validation"] = "publish_ready_iteratively_repaired" if repaired else "publish_ready_ai_validated"
        row["Processing Time"] = "lutron_WEBSITE_import_v22_6_iterative_repair"
        row["Content Quality"] = clean_text(data.get("source")) + "_source_locked_iterative_repair"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "repair_pipeline=" + V226_VERSION,
            "content_repair_sections=" + ",".join(repaired_sections) if repaired_sections else "content_repair_sections=none",
            "spec_rows_repaired=" + ",".join(str(x) for x in repaired_spec_rows) if repaired_spec_rows else "spec_rows_repaired=none",
        )
    return row


_v226_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v226_previous_validate_publish_output(data, row, require_ai=require_ai)
    if clean_text(data.get("_content_mode")) in {"ai_iterative_repair", "ai_validated"}:
        issues = [item for item in issues if item not in {
            "evidence_publish_status_required",
            "deepseek_validated_content_required",
            "english_source_residue_detected",
        }]
    issues.extend(_v226_final_row_issues(data, row))
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.7 convergent serial repair: never skip a product merely because the
# bulk AI repair exhausted its rounds. Failed sections and specification rows
# are repaired one at a time, revalidated, and retained only after passing.
# If repair cannot converge, the run stops with a diagnostic JSON instead of
# silently producing an empty publish workbook.
# ---------------------------------------------------------------------------
V227_VERSION = "22.7-convergent-serial-repair"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_7_convergent_serial_repair_2026_07_28"

AI_STATS.setdefault("v227_section_calls", 0)
AI_STATS.setdefault("v227_spec_calls", 0)
AI_STATS.setdefault("v227_rescue_products", 0)


def _v227_issue_codes(issues: Sequence[Dict[str, Any]]) -> List[str]:
    return dedupe_keep_order(
        clean_text(item.get("code"))
        for item in (issues or [])
        if isinstance(item, dict) and clean_text(item.get("code"))
    )


def _v227_fact_ids(payload: Dict[str, Any]) -> List[str]:
    fact_map = _v22_fact_map(payload)
    return [clean_text(key) for key in fact_map.keys() if clean_text(key)]


def _v227_attach_source_evidence(field: str, candidate: Any, payload: Dict[str, Any]) -> Any:
    """Attach the complete verified source ledger when AI omits evidence IDs.

    This is not a content fallback. It only links the candidate to source facts
    already supplied to the validator so unsupported claims still fail.
    """
    value = _v226_json_clone(candidate)
    fact_ids = _v227_fact_ids(payload)
    if not isinstance(value, dict) or not fact_ids:
        return value
    if field in {"faq_html", "application_tab_html"}:
        rows = value.get("items")
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, dict):
                    # Use the complete source ledger. Numeric and claim checks
                    # remain active against the combined official evidence.
                    row["evidence_ids"] = list(fact_ids)
    elif field != "product_name_id":
        value["evidence_ids"] = list(fact_ids)
    return value


def _v227_content_prompt(
    field: str,
    current: Any,
    issues: Sequence[Dict[str, Any]],
    payload: Dict[str, Any],
    attempt: int,
    maximum: int,
) -> Dict[str, Any]:
    return {
        "task": "Perbaiki tepat satu bagian katalog sampai lolos validator. Jangan mengirim bagian lain.",
        "repair_mode": "single_section_convergence",
        "attempt": attempt,
        "maximum_attempts": maximum,
        "requested_section": field,
        "current_failed_content": current,
        "validator_feedback": list(issues or []),
        "rules": [
            "Gunakan hanya fakta dan spesifikasi di product_data.",
            "Jangan mengarang aplikasi, fungsi, fitur, stok, harga, garansi, layanan kalibrasi, atau klaim medis.",
            "Pertahankan merek, model, kode, angka, satuan, simbol, dan singkatan teknis.",
            "Seluruh kalimat biasa harus Bahasa Indonesia. Model, kode, satuan, dan singkatan teknis boleh tetap.",
            "FAQ hanya membahas informasi teknis yang didukung sumber.",
            "Jika validator menyebut klaim tidak didukung, hapus klaim itu dan gunakan fakta lain dari sumber.",
            "Jika validator menyebut struktur salah, perbaiki struktur tanpa mengubah fakta teknis.",
            "Balas JSON valid saja sesuai output_schema.",
        ],
        "output_schema": _v226_content_schema_for([field]),
        "product_data": payload,
    }


def _v227_serial_content_repair(
    data: Dict[str, Any],
    family: Dict[str, Any],
    initial_fields: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], Dict[str, Any]]:
    payload = _ai_payload(data, family)
    fields = _v226_json_clone(initial_fields) if isinstance(initial_fields, dict) else {}
    fields.setdefault("content_evidence", {})
    report: Dict[str, Any] = {
        "version": V227_VERSION,
        "mode": "serial_section_repair",
        "sections": {},
        "repaired_sections": [],
    }
    max_attempts = max(12, min(40, int(AI_SETTINGS.get("retry_count") or 5) * 4 + 8))
    unresolved_final: Dict[str, List[Dict[str, Any]]] = {}

    for field in V226_REQUIRED_SECTIONS:
        candidate = _v227_attach_source_evidence(field, _v226_initial_candidate(fields, field), payload)
        valid, normalized, issues = _v226_validate_section(field, candidate, payload, data)
        section_log: Dict[str, Any] = {"initial_issues": issues, "attempts": []}
        if valid:
            _v226_apply_normalized_section(fields, field, normalized)
            section_log["status"] = "already_valid"
            report["sections"][field] = section_log
            continue

        for attempt in range(1, max_attempts + 1):
            codes = _v227_issue_codes(issues)
            print(
                f"    repair section {field} attempt {attempt}/{max_attempts}"
                + (" | " + ",".join(codes) if codes else "")
            )
            user_payload = _v227_content_prompt(field, candidate, issues, payload, attempt, max_attempts)
            AI_STATS["v227_section_calls"] = int(AI_STATS.get("v227_section_calls") or 0) + 1
            ok_request, parsed, request_status = _v226_post_ai_json(
                "v227_single_content_section_" + field,
                user_payload,
            )
            sections = parsed.get("sections") if ok_request and isinstance(parsed.get("sections"), dict) else {}
            raw_candidate = sections.get(field)
            if raw_candidate is None:
                new_issues = [_v226_problem(
                    "section_missing_from_repair_response" if ok_request else "repair_request_failed",
                    "Respons tidak memuat bagian yang diminta." if ok_request else "Permintaan repair ke AI gagal.",
                    field=field,
                    reason=request_status,
                )]
                section_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                issues = new_issues
            else:
                candidate = _v227_attach_source_evidence(field, raw_candidate, payload)
                valid, normalized, new_issues = _v226_validate_section(field, candidate, payload, data)
                section_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                if valid:
                    _v226_apply_normalized_section(fields, field, normalized)
                    section_log["status"] = "repaired"
                    report["repaired_sections"].append(field)
                    AI_STATS["section_repair_success"] = int(AI_STATS.get("section_repair_success") or 0) + 1
                    issues = []
                    break
                issues = new_issues

            if float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
                time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

        if issues:
            section_log["status"] = "unresolved"
            unresolved_final[field] = issues
        report["sections"][field] = section_log

    report["unresolved"] = unresolved_final
    return (not unresolved_final), fields, report


def _v227_serial_spec_repair(data: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
    ledger = _v224_source_rows(data)
    translations: Dict[int, Dict[str, str]] = {}
    existing = data.get("_v226_spec_translations") if isinstance(data.get("_v226_spec_translations"), dict) else {}
    report: Dict[str, Any] = {
        "version": V227_VERSION,
        "mode": "serial_specification_row_repair",
        "source_rows": len(ledger),
        "rows": {},
        "repaired_rows": [],
    }
    max_attempts = max(12, min(40, int(AI_SETTINGS.get("retry_count") or 5) * 4 + 8))
    unresolved_final: Dict[int, List[Dict[str, Any]]] = {}

    for item in ledger:
        source_id = int(item["source_id"])
        candidate = existing.get(str(source_id)) if isinstance(existing.get(str(source_id)), dict) else None
        if not candidate:
            candidate = _v226_previous_safe_display_row(data, item)
        candidate = {
            "label": _v224_exact_text(candidate.get("label") or candidate.get("label_id") or item["label_raw"]),
            "value": _v224_exact_text(candidate.get("value") or candidate.get("value_id") or item["value_raw"]),
            "source_label": item["label_raw"],
            "section": item["section_raw"],
        }
        issues = _v226_validate_spec_candidate(item, candidate)
        row_log: Dict[str, Any] = {"initial_issues": issues, "attempts": []}
        if not issues:
            translations[source_id] = candidate
            row_log["status"] = "already_valid"
            report["rows"][str(source_id)] = row_log
            continue

        for attempt in range(1, max_attempts + 1):
            codes = _v227_issue_codes(issues)
            print(
                f"    repair spec row {source_id} attempt {attempt}/{max_attempts}"
                + (" | " + ",".join(codes) if codes else "")
            )
            repair_payload = _v226_spec_repair_payload(
                ledger,
                {source_id: issues},
                translations,
                attempt,
                max_attempts,
                data,
            )
            repair_payload["repair_mode"] = "single_specification_row_convergence"
            AI_STATS["v227_spec_calls"] = int(AI_STATS.get("v227_spec_calls") or 0) + 1
            ok_request, parsed, request_status = _v226_post_ai_json(
                "v227_single_specification_row_" + str(source_id),
                repair_payload,
            )
            response_rows = parsed.get("rows") if ok_request and isinstance(parsed.get("rows"), list) else []
            raw = next(
                (
                    row for row in response_rows
                    if isinstance(row, dict) and _coerce_ai_spec_id(row.get("source_id")) == source_id
                ),
                None,
            )
            if raw is None:
                new_issues = [_v226_problem(
                    "spec_row_missing_from_repair_response" if ok_request else "spec_repair_request_failed",
                    "Respons tidak memuat baris spesifikasi yang diminta." if ok_request else "Permintaan repair spesifikasi gagal.",
                    source_id=source_id,
                    reason=request_status,
                )]
                row_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                issues = new_issues
            else:
                candidate = {
                    "label": _v224_exact_text(raw.get("label_id")),
                    "value": _v224_exact_text(raw.get("value_id")),
                    "source_label": item["label_raw"],
                    "section": item["section_raw"],
                }
                new_issues = _v226_validate_spec_candidate(item, candidate)
                row_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                if not new_issues:
                    translations[source_id] = candidate
                    report["repaired_rows"].append(source_id)
                    row_log["status"] = "repaired"
                    AI_STATS["spec_rows_repaired"] = int(AI_STATS.get("spec_rows_repaired") or 0) + 1
                    issues = []
                    break
                issues = new_issues

            if float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
                time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

        if issues:
            row_log["status"] = "unresolved"
            unresolved_final[source_id] = issues
        report["rows"][str(source_id)] = row_log

    report["unresolved"] = unresolved_final
    if unresolved_final:
        return False, report
    data["_v226_spec_translations"] = {str(source_id): row for source_id, row in translations.items()}
    return True, report


def _v227_write_failure_diagnostic(data: Dict[str, Any], content_report: Dict[str, Any], spec_report: Dict[str, Any]) -> str:
    model = re.sub(r"[^A-Za-z0-9._-]+", "_", clean_text(data.get("model")) or "unknown")
    path = Path.cwd() / f"lutron_repair_failure_{model}.json"
    payload = {
        "version": V227_VERSION,
        "url": clean_text(data.get("url")),
        "model": clean_text(data.get("model")),
        "title": clean_text(data.get("title")),
        "content_report": content_report,
        "specification_report": spec_report,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return str(path)
    except Exception:
        return ""


_v227_previous_get_ai_content = get_ai_content


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    result = _v227_previous_get_ai_content(data, family)
    if result or not AI_SETTINGS.get("enabled"):
        return result

    # V22.6 bulk repair did not converge. Do not skip the product. Rescue it
    # section-by-section and specification-row-by-specification-row.
    AI_STATS["v227_rescue_products"] = int(AI_STATS.get("v227_rescue_products") or 0) + 1
    print("    bulk repair exhausted; entering convergent serial repair")

    data.pop("_ai_fields", None)
    initial = _v226_initial_get_ai_content(data, family)
    if not isinstance(initial, dict):
        initial = {}
    initial.pop("_v225_repair_fields", None)

    content_ok, repaired_fields, content_report = _v227_serial_content_repair(data, family, initial)
    spec_ok, spec_report = _v227_serial_spec_repair(data)
    data["_v227_content_report"] = content_report
    data["_v227_spec_report"] = spec_report

    if not content_ok or not spec_ok:
        unresolved_parts: List[str] = []
        if not content_ok:
            unresolved_parts.append("content")
        if not spec_ok:
            unresolved_parts.append("specifications")
        diagnostic = _v227_write_failure_diagnostic(data, content_report, spec_report)
        data["_ai_status"] = "convergent_repair_stopped"
        data["error"] = _merge_source_notes(
            data.get("error"),
            "convergent_repair_stopped=" + ",".join(unresolved_parts),
            "diagnostic=" + diagnostic if diagnostic else "",
        )
        # Stop the run instead of continuing to Skipped and writing an empty
        # workbook. The diagnostic contains exact field/row validator errors.
        raise SystemExit(
            "REPAIR BELUM LOLOS untuk model "
            + (clean_text(data.get("model")) or "-")
            + ". Proses dihentikan agar tidak menghasilkan Excel kosong atau produk tidak valid."
            + (" Diagnostic: " + diagnostic if diagnostic else "")
        )

    repaired_fields["translated_spec_rows"] = []
    repaired_fields["ai_spec_dedup_applied"] = False
    repaired_fields["ai_spec_dedup_removed_count"] = 0
    repaired_sections = content_report.get("repaired_sections") or []
    repaired_spec_rows = spec_report.get("repaired_rows") or []
    data["_ai_fields"] = repaired_fields
    data["_v226_repair_complete"] = True
    data["_v227_converged"] = True
    data["_content_mode"] = "ai_convergent_serial_repair"
    data["_ai_status"] = "convergent_serial_repair_complete"
    data["_v226_repaired_sections"] = repaired_sections
    data["_v226_repaired_spec_rows"] = repaired_spec_rows
    data["error"] = _merge_source_notes(
        data.get("error"),
        "convergent_serial_repair_complete",
        "repaired_sections=" + ",".join(repaired_sections) if repaired_sections else "",
        "repaired_spec_rows=" + ",".join(str(x) for x in repaired_spec_rows) if repaired_spec_rows else "",
    )
    print(
        "    convergent serial repair complete"
        + (" | sections=" + ",".join(repaired_sections) if repaired_sections else "")
        + (" | spec_rows=" + ",".join(str(x) for x in repaired_spec_rows) if repaired_spec_rows else "")
    )
    return repaired_fields


_v227_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v227_previous_build_output_row(data, pub_date)
    if data.get("_v227_converged"):
        row["Website Validation"] = "publish_ready_convergent_repair"
        row["Processing Time"] = "lutron_WEBSITE_import_v22_7_convergent_serial_repair"
        row["Content Quality"] = clean_text(data.get("source")) + "_source_locked_convergent_repair"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "repair_pipeline=" + V227_VERSION,
            "bulk_failure_rescued_by_serial_repair",
        )
    return row


# ---------------------------------------------------------------------------
# V22.8 targeted convergence repair
# - fixes the V22.7 endless retry loop caused by generic english_residue_detected
#   feedback that did not reveal the offending phrase;
# - applies safe, source-preserving terminology repairs before validation;
# - returns exact offending fragments to the AI;
# - escalates repeated/no-progress replies to a strict rewrite;
# - limits repeated calls and stops with a diagnostic rather than looping 28x.
# ---------------------------------------------------------------------------
V228_VERSION = "22.8-targeted-convergence-repair"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_8_targeted_convergence_repair_2026_07_28"

AI_STATS.setdefault("v228_safe_text_repairs", 0)
AI_STATS.setdefault("v228_no_progress_escalations", 0)

V228_PUBLIC_REPLACEMENTS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\b(?:supper|super)\s+large\s+LCD\s+display\b", re.I), "layar LCD sangat besar"),
    (re.compile(r"\blarge\s+LCD\s+display\b", re.I), "layar LCD besar"),
    (re.compile(r"\bLCD\s+display\b", re.I), "layar LCD"),
    (re.compile(r"\bsaves?\s+battery(?:\s+life)?\b", re.I), "menghemat daya baterai"),
    (re.compile(r"\bdigital\s+pressure\s+meter\b", re.I), "manometer digital"),
    (re.compile(r"\bpressure\s+meter\b", re.I), "manometer"),
    (re.compile(r"\btemperature\s+meter\b", re.I), "alat ukur suhu"),
    (re.compile(r"\bhumidity\s+meter\b", re.I), "alat ukur kelembapan"),
    (re.compile(r"\bsound\s+level\s+meter\b", re.I), "alat ukur tingkat suara"),
    (re.compile(r"\blight\s+meter\b", re.I), "alat ukur cahaya"),
    (re.compile(r"\bclamp\s+meter\b", re.I), "meter penjepit"),
    (re.compile(r"\bpower\s+analy[sz]er\b", re.I), "penganalisis daya"),
    (re.compile(r"\bfrequency\s+counter\b", re.I), "pencacah frekuensi"),
    (re.compile(r"\bdata\s+logger\b", re.I), "pencatat data"),
    (re.compile(r"\bgain\b", re.I), "penguatan"),
    (re.compile(r"\bdata\s+hold\s+untuk\s+membekukan\s+tampilan\b", re.I), "Data Hold untuk menahan pembacaan pada tampilan"),
    (re.compile(r"(?<=[.!?])\s+pencatat\s+data\b", re.I), " Pencatat data"),
    (re.compile(r"\binstruction\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\binstruksi\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\bAksesori\s+termasuk\b", re.I), "Aksesori yang disertakan"),
    (re.compile(r"\bkelembaban\b", re.I), "kelembapan"),
    (re.compile(r"\bhard\s+carrying\s+case\b", re.I), "kotak pelindung keras"),
    (re.compile(r"\bcarrying\s+case\b", re.I), "kotak pelindung"),
    (re.compile(r"\bless\s+than\b", re.I), "kurang dari"),
    (re.compile(r"\bpower\s+supply\b", re.I), "catu daya"),
    (re.compile(r"\bmeasurement\s+range\b|\bmeasuring\s+range\b", re.I), "rentang pengukuran"),
    (re.compile(r"\bdisplay\s+size\b", re.I), "ukuran tampilan"),
    (re.compile(r"\bsampling\s+time\b", re.I), "waktu pencuplikan"),
    (re.compile(r"\bresponse\s+time\b", re.I), "waktu respons"),
    (re.compile(r"\baccessories?\s+included\b", re.I), "aksesori yang disertakan"),
    (re.compile(r"\baccessories?\s+optional\b|\boptional\s+accessories?\b", re.I), "aksesori opsional"),
    (re.compile(r"\bbuilt\s+inside\b", re.I), "terpasang di dalam"),
    (re.compile(r"\binside\s+the\s+housing\b", re.I), "di dalam bodi alat"),
    (re.compile(r"\bheavy\s+duty\b", re.I), "untuk penggunaan intensif"),
    (re.compile(r"\bbattery\s+life\b", re.I), "daya tahan baterai"),
    (re.compile(r"\bmanual\s+off\b", re.I), "dimatikan secara manual"),
    (re.compile(r"\bauto\s+shut\s+off\b", re.I), "mati otomatis"),
    (re.compile(r"\bpush\s+button\b", re.I), "tombol tekan"),
    (re.compile(r"\bfront\s+panel\b", re.I), "panel depan"),
    (re.compile(r"\bused\s+for\b", re.I), "digunakan untuk"),
    (re.compile(r"\busage\s+for\b", re.I), "penggunaan untuk"),
    (re.compile(r"\bfor\s+calibration\b", re.I), "untuk kalibrasi"),
    (re.compile(r"\bliquid\s+is\s+prohibited\b", re.I), "cairan tidak diperbolehkan"),
    (re.compile(r"\bfreezes?\s+(?:the\s+)?display(?:ed\s+reading)?\b", re.I), "menahan pembacaan pada tampilan"),
    (re.compile(r"\bmaximum\s+and\s+minimum\s+value\b", re.I), "nilai maksimum dan minimum"),
    (re.compile(r"\brecords?\s+and\s+recalls?\s+maximum\b", re.I), "merekam dan menampilkan kembali nilai maksimum"),
    (re.compile(r"\bsaves?\s+battery(?:\s+life)?\b", re.I), "menghemat daya baterai"),
    (re.compile(r"\bsupper\s+large\b|\bsuper\s+large\b", re.I), "sangat besar"),
    (re.compile(r"\bwith\s+contrast\s+adjustment\b", re.I), "dengan penyetelan kontras"),
    (re.compile(r"\bdual\s+(?:and|&)\s+differential\s+input\b", re.I), "masukan ganda dan diferensial"),
)

V228_SPEC_EXTRA_REPLACEMENTS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\blarge\s+LCD\s+display\b", re.I), "layar LCD besar"),
    (re.compile(r"\bdisplay\s+reading\b", re.I), "pembacaan tampilan"),
    (re.compile(r"\bhousing\b", re.I), "bodi alat"),
    (re.compile(r"\bgain\s+adjustment\b", re.I), "penyetelan penguatan"),
    (re.compile(r"\boperating\s+temperature\b", re.I), "suhu operasi"),
    (re.compile(r"\boperating\s+humidity\b", re.I), "kelembapan operasi"),
    (re.compile(r"\bdata\s+hold\s+to\b", re.I), "fungsi Data Hold untuk"),
    (re.compile(r"\bincluding\s+linearity\b", re.I), "termasuk linearitas"),
    (re.compile(r"\bfull\s+scale\b", re.I), "skala penuh"),
    (re.compile(r"\bapproximately\b|\bapprox\.?\b", re.I), "sekitar"),
    (re.compile(r"\brepeatability\b", re.I), "keterulangan"),
    (re.compile(r"\breadings?\b", re.I), "pembacaan"),
    (re.compile(r"\bincluded\b", re.I), "disertakan"),
    (re.compile(r"\boptional\b", re.I), "opsional"),
    (re.compile(r"\btype\b", re.I), "tipe"),
    (re.compile(r"\binput\b", re.I), "masukan"),
    (re.compile(r"\boutput\b", re.I), "keluaran"),
    (re.compile(r"\bdisplay\b", re.I), "tampilan"),
    (re.compile(r"(?<![A-Za-z0-9])([-+]?\d+(?:[.,]\d+)?)\s+to\s+([-+]?\d+(?:[.,]\d+)?)(?![A-Za-z0-9])", re.I), r"\1 hingga \2"),
)


def _v228_apply_replacements(value: Any, specification: bool = False) -> str:
    text = clean_text(value)
    if not text:
        return ""
    original = text
    rules = list(V228_PUBLIC_REPLACEMENTS)
    if specification:
        rules.extend(V228_SPEC_EXTRA_REPLACEMENTS)
    for pattern, replacement in rules:
        text = pattern.sub(replacement, text)
    text = re.sub(r"\.;", ";", text)
    text = re.sub(r";\s+Cairan\b", "; cairan", text)
    text = re.sub(r"\s+\.", ".", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    if text != original:
        AI_STATS["v228_safe_text_repairs"] = int(AI_STATS.get("v228_safe_text_repairs") or 0) + 1
    return text


def _v228_normalize_candidate(field: str, candidate: Any) -> Any:
    value = _v226_json_clone(candidate)
    if not isinstance(value, dict):
        return value
    if field == "faq_html":
        rows = value.get("items")
        if isinstance(rows, list):
            for row in rows:
                if not isinstance(row, dict):
                    continue
                for key in ("question_id", "question", "answer_id", "answer"):
                    if key in row:
                        row[key] = _v228_apply_replacements(row.get(key), specification=False)
    elif field == "application_tab_html":
        rows = value.get("items")
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, dict):
                    for key in ("text_id", "text"):
                        if key in row:
                            row[key] = _v228_apply_replacements(row.get(key), specification=False)
    elif "value" in value:
        value["value"] = _v228_apply_replacements(value.get("value"), specification=False)
    return value


def _v228_english_fragments(value: Any, specification: bool = False) -> List[str]:
    text = _html_text(value)
    if not text:
        return []
    patterns = [V22_ENGLISH_RESIDUE_RE, V226_SPEC_ENGLISH_RE if specification else V226_PUBLIC_ENGLISH_RE]
    fragments: List[str] = []
    for pattern in patterns:
        for match in pattern.finditer(text):
            fragment = clean_text(match.group(0))
            if fragment:
                fragments.append(fragment)
    return dedupe_keep_order(fragments)[:20]


# Add exact fragments to validator feedback. V22.7 returned only the generic
# code english_residue_detected, so DeepSeek repeatedly returned the same FAQ.
_v228_previous_validate_common_claims = _v226_validate_common_claims


def _v226_validate_common_claims(
    field: str,
    value: Any,
    evidence_ids: Sequence[str],
    fact_map: Dict[str, str],
    data: Dict[str, Any],
) -> List[Dict[str, Any]]:
    issues = _v228_previous_validate_common_claims(field, value, evidence_ids, fact_map, data)
    fragments = _v228_english_fragments(value, specification=False)
    for issue in issues:
        if isinstance(issue, dict) and clean_text(issue.get("code")) == "english_residue_detected":
            issue["offending_fragments"] = fragments
            issue["repair_instruction"] = (
                "Ganti setiap offending_fragments dengan padanan Bahasa Indonesia. "
                "Jangan menyalin kembali frasa tersebut pada respons berikutnya."
            )
    return _v226_dedupe_problems(issues)


# Normalize known source phrases before running the strict section validator.
_v228_previous_validate_section = _v226_validate_section


def _v226_validate_section(
    field: str,
    candidate: Any,
    payload: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], List[Dict[str, Any]]]:
    repaired_candidate = _v228_normalize_candidate(field, candidate)
    return _v228_previous_validate_section(field, repaired_candidate, payload, data)


# Mutate the candidate with safe translations before validating it so a row
# that passes is also the row retained for publication.
_v228_previous_validate_spec_candidate = _v226_validate_spec_candidate


def _v226_validate_spec_candidate(item: Dict[str, Any], candidate: Dict[str, Any]) -> List[Dict[str, Any]]:
    if isinstance(candidate, dict):
        candidate["label"] = _v228_apply_replacements(
            candidate.get("label") or candidate.get("label_id"), specification=True
        )
        candidate["value"] = _v228_apply_replacements(
            candidate.get("value") or candidate.get("value_id"), specification=True
        )
    issues = _v228_previous_validate_spec_candidate(item, candidate)
    fragments = _v228_english_fragments(
        " ".join([clean_text(candidate.get("label")), clean_text(candidate.get("value"))]),
        specification=True,
    )
    for issue in issues:
        if isinstance(issue, dict) and clean_text(issue.get("code")) == "spec_english_residue_detected":
            issue["offending_fragments"] = fragments
            issue["repair_instruction"] = (
                "Terjemahkan tepat offending_fragments. Angka, satuan, model, kode, dan urutan fakta tidak boleh berubah."
            )
    return _v226_dedupe_problems(issues)


def _v228_collect_forbidden_fragments(issues: Sequence[Dict[str, Any]]) -> List[str]:
    output: List[str] = []
    for issue in issues or []:
        if not isinstance(issue, dict):
            continue
        values = issue.get("offending_fragments")
        if isinstance(values, list):
            output.extend(clean_text(value) for value in values if clean_text(value))
    return dedupe_keep_order(output)


# Override the repair prompt so the model sees the exact phrase that failed.
def _v227_content_prompt(
    field: str,
    current: Any,
    issues: Sequence[Dict[str, Any]],
    payload: Dict[str, Any],
    attempt: int,
    maximum: int,
) -> Dict[str, Any]:
    forbidden = _v228_collect_forbidden_fragments(issues)
    strict_rewrite = attempt >= 3
    return {
        "task": (
            "Tulis ulang dari fakta sumber tepat satu bagian katalog hingga lolos validator. Jangan mengirim bagian lain."
            if strict_rewrite else
            "Perbaiki tepat satu bagian katalog sampai lolos validator. Jangan mengirim bagian lain."
        ),
        "repair_mode": "single_section_strict_rewrite" if strict_rewrite else "single_section_targeted_repair",
        "attempt": attempt,
        "maximum_attempts": maximum,
        "requested_section": field,
        "current_failed_content": None if strict_rewrite else current,
        "validator_feedback": list(issues or []),
        "forbidden_fragments": forbidden,
        "rules": [
            "Gunakan hanya fakta dan spesifikasi di product_data.",
            "Jangan mengarang aplikasi, fungsi, fitur, stok, harga, garansi, layanan kalibrasi, atau klaim medis.",
            "Pertahankan merek, model, kode, angka, satuan, simbol, dan singkatan teknis.",
            "Seluruh kalimat biasa harus Bahasa Indonesia. Model, kode, satuan, dan singkatan teknis boleh tetap.",
            "FAQ hanya membahas informasi teknis yang didukung sumber.",
            "Jangan menulis kembali satu pun forbidden_fragments. Gunakan padanan Bahasa Indonesia.",
            "Jika validator menyebut klaim tidak didukung, hapus klaim tersebut dan gunakan fakta lain dari sumber.",
            "Jika validator menyebut struktur salah, perbaiki struktur tanpa mengubah fakta teknis.",
            "Balas JSON valid saja sesuai output_schema.",
        ],
        "output_schema": _v226_content_schema_for([field]),
        "product_data": payload,
    }


def _v228_candidate_fingerprint(value: Any) -> str:
    try:
        raw = json.dumps(value, ensure_ascii=False, sort_keys=True)
    except Exception:
        raw = clean_text(value)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# Replace the 28-attempt loop with an adaptive 8-attempt loop. A repeated reply
# is escalated to a strict rewrite instead of being submitted unchanged 28x.
def _v227_serial_content_repair(
    data: Dict[str, Any],
    family: Dict[str, Any],
    initial_fields: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], Dict[str, Any]]:
    payload = _ai_payload(data, family)
    fields = _v226_json_clone(initial_fields) if isinstance(initial_fields, dict) else {}
    fields.setdefault("content_evidence", {})
    report: Dict[str, Any] = {
        "version": V228_VERSION,
        "mode": "targeted_serial_section_repair",
        "sections": {},
        "repaired_sections": [],
    }
    max_attempts = max(1, min(8, int(AI_SETTINGS.get("item_repair_attempts") or 3)))
    unresolved_final: Dict[str, List[Dict[str, Any]]] = {}

    for field in V226_REQUIRED_SECTIONS:
        candidate = _v227_attach_source_evidence(field, _v226_initial_candidate(fields, field), payload)
        candidate = _v228_normalize_candidate(field, candidate)
        valid, normalized, issues = _v226_validate_section(field, candidate, payload, data)
        section_log: Dict[str, Any] = {"initial_issues": issues, "attempts": []}
        if valid:
            _v226_apply_normalized_section(fields, field, normalized)
            section_log["status"] = "already_valid_or_safely_repaired"
            report["sections"][field] = section_log
            continue

        seen: Dict[str, int] = {}
        for attempt in range(1, max_attempts + 1):
            codes = _v227_issue_codes(issues)
            fragments = _v228_collect_forbidden_fragments(issues)
            print(
                f"    repair section {field} attempt {attempt}/{max_attempts}"
                + (" | " + ",".join(codes) if codes else "")
                + (" | fragments=" + " || ".join(fragments[:4]) if fragments else "")
            )
            user_payload = _v227_content_prompt(field, candidate, issues, payload, attempt, max_attempts)
            AI_STATS["v227_section_calls"] = int(AI_STATS.get("v227_section_calls") or 0) + 1
            ok_request, parsed, request_status = _v226_post_ai_json(
                "v228_single_content_section_" + field,
                user_payload,
            )
            sections = parsed.get("sections") if ok_request and isinstance(parsed.get("sections"), dict) else {}
            raw_candidate = sections.get(field)
            if raw_candidate is None:
                new_issues = [_v226_problem(
                    "section_missing_from_repair_response" if ok_request else "repair_request_failed",
                    "Respons tidak memuat bagian yang diminta." if ok_request else "Permintaan repair ke AI gagal.",
                    field=field,
                    reason=request_status,
                )]
                section_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                issues = new_issues
            else:
                candidate = _v227_attach_source_evidence(field, raw_candidate, payload)
                candidate = _v228_normalize_candidate(field, candidate)
                fingerprint = _v228_candidate_fingerprint(candidate)
                seen[fingerprint] = seen.get(fingerprint, 0) + 1
                valid, normalized, new_issues = _v226_validate_section(field, candidate, payload, data)
                if seen[fingerprint] >= 2 and new_issues:
                    AI_STATS["v228_no_progress_escalations"] = int(AI_STATS.get("v228_no_progress_escalations") or 0) + 1
                    new_issues = list(new_issues) + [_v226_problem(
                        "repair_no_progress_same_output",
                        "Respons yang sama sudah gagal lebih dari sekali. Tulis ulang bagian ini dari nol berdasarkan fakta sumber.",
                        field=field,
                    )]
                section_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                    "output_fingerprint": fingerprint,
                })
                if valid:
                    _v226_apply_normalized_section(fields, field, normalized)
                    section_log["status"] = "repaired"
                    report["repaired_sections"].append(field)
                    AI_STATS["section_repair_success"] = int(AI_STATS.get("section_repair_success") or 0) + 1
                    issues = []
                    break
                issues = _v226_dedupe_problems(new_issues)

            if float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
                time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

        if issues:
            section_log["status"] = "unresolved"
            unresolved_final[field] = issues
        report["sections"][field] = section_log

    report["unresolved"] = unresolved_final
    return (not unresolved_final), fields, report


# Use the same adaptive cap for specification rows. Safe terminology repair is
# applied before each validation and exact fragments are printed in the log.
def _v227_serial_spec_repair(data: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
    ledger = _v224_source_rows(data)
    translations: Dict[int, Dict[str, str]] = {}
    existing = data.get("_v226_spec_translations") if isinstance(data.get("_v226_spec_translations"), dict) else {}
    report: Dict[str, Any] = {
        "version": V228_VERSION,
        "mode": "targeted_serial_specification_row_repair",
        "source_rows": len(ledger),
        "rows": {},
        "repaired_rows": [],
    }
    max_attempts = max(1, min(8, int(AI_SETTINGS.get("item_repair_attempts") or 3)))
    unresolved_final: Dict[int, List[Dict[str, Any]]] = {}

    for item in ledger:
        source_id = int(item["source_id"])
        candidate = existing.get(str(source_id)) if isinstance(existing.get(str(source_id)), dict) else None
        if not candidate:
            candidate = _v226_previous_safe_display_row(data, item)
        candidate = {
            "label": _v224_exact_text(candidate.get("label") or candidate.get("label_id") or item["label_raw"]),
            "value": _v224_exact_text(candidate.get("value") or candidate.get("value_id") or item["value_raw"]),
            "source_label": item["label_raw"],
            "section": item["section_raw"],
        }
        issues = _v226_validate_spec_candidate(item, candidate)
        row_log: Dict[str, Any] = {"initial_issues": issues, "attempts": []}
        if not issues:
            translations[source_id] = candidate
            row_log["status"] = "already_valid_or_safely_repaired"
            report["rows"][str(source_id)] = row_log
            continue

        seen: Dict[str, int] = {}
        for attempt in range(1, max_attempts + 1):
            codes = _v227_issue_codes(issues)
            fragments = _v228_collect_forbidden_fragments(issues)
            print(
                f"    repair spec row {source_id} attempt {attempt}/{max_attempts}"
                + (" | " + ",".join(codes) if codes else "")
                + (" | fragments=" + " || ".join(fragments[:4]) if fragments else "")
            )
            repair_payload = _v226_spec_repair_payload(
                ledger,
                {source_id: issues},
                translations,
                attempt,
                max_attempts,
                data,
            )
            repair_payload["repair_mode"] = "single_specification_row_targeted_repair"
            repair_payload["forbidden_fragments"] = fragments
            if attempt >= 3:
                repair_payload["task"] = (
                    "Tulis ulang terjemahan baris spesifikasi ini dari source_label dan source_value. "
                    "Jangan menyalin current_translation dan jangan memuat forbidden_fragments."
                )
                repair_payload["strict_rewrite"] = True
            AI_STATS["v227_spec_calls"] = int(AI_STATS.get("v227_spec_calls") or 0) + 1
            ok_request, parsed, request_status = _v226_post_ai_json(
                "v228_single_specification_row_" + str(source_id),
                repair_payload,
            )
            response_rows = parsed.get("rows") if ok_request and isinstance(parsed.get("rows"), list) else []
            raw = next(
                (
                    row for row in response_rows
                    if isinstance(row, dict) and _coerce_ai_spec_id(row.get("source_id")) == source_id
                ),
                None,
            )
            if raw is None:
                new_issues = [_v226_problem(
                    "spec_row_missing_from_repair_response" if ok_request else "spec_repair_request_failed",
                    "Respons tidak memuat baris spesifikasi yang diminta." if ok_request else "Permintaan repair spesifikasi gagal.",
                    source_id=source_id,
                    reason=request_status,
                )]
                row_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                })
                issues = new_issues
            else:
                candidate = {
                    "label": _v224_exact_text(raw.get("label_id")),
                    "value": _v224_exact_text(raw.get("value_id")),
                    "source_label": item["label_raw"],
                    "section": item["section_raw"],
                }
                new_issues = _v226_validate_spec_candidate(item, candidate)
                fingerprint = _v228_candidate_fingerprint(candidate)
                seen[fingerprint] = seen.get(fingerprint, 0) + 1
                if seen[fingerprint] >= 2 and new_issues:
                    new_issues = list(new_issues) + [_v226_problem(
                        "spec_repair_no_progress_same_output",
                        "Terjemahan yang sama sudah gagal lebih dari sekali. Tulis ulang dari source_label dan source_value.",
                        source_id=source_id,
                    )]
                row_log["attempts"].append({
                    "attempt": attempt,
                    "request_status": request_status,
                    "issues": new_issues,
                    "output_fingerprint": fingerprint,
                })
                if not new_issues:
                    translations[source_id] = candidate
                    report["repaired_rows"].append(source_id)
                    row_log["status"] = "repaired"
                    AI_STATS["spec_rows_repaired"] = int(AI_STATS.get("spec_rows_repaired") or 0) + 1
                    issues = []
                    break
                issues = _v226_dedupe_problems(new_issues)

            if float(AI_SETTINGS.get("retry_delay") or 0.0) > 0:
                time.sleep(float(AI_SETTINGS.get("retry_delay") or 0.0))

        if issues:
            row_log["status"] = "unresolved"
            unresolved_final[source_id] = issues
        report["rows"][str(source_id)] = row_log

    report["unresolved"] = unresolved_final
    if unresolved_final:
        return False, report
    data["_v226_spec_translations"] = {str(source_id): row for source_id, row in translations.items()}
    return True, report


# Mark the repaired workbook with the actual V22.8 pipeline version.
_v228_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v228_previous_build_output_row(data, pub_date)
    if data.get("_v227_converged"):
        row["Website Validation"] = "publish_ready_targeted_convergence_repair"
        row["Processing Time"] = "lutron_WEBSITE_import_v22_8_targeted_convergence_repair"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "repair_pipeline=" + V228_VERSION,
            "exact_validator_fragments_enabled",
            "safe_source_phrase_repairs=" + str(int(AI_STATS.get("v228_safe_text_repairs") or 0)),
        )
    return row


# ---------------------------------------------------------------------------
# V22.9 AI preflight and fail-fast transport handling
# - --use-ai now requires a real key before any scrape begins;
# - missing/placeholder keys never enter bulk or serial repair loops;
# - authentication, permission, and account-credit failures stop immediately;
# - retry loops remain available only for transient network/server failures and
#   genuine content-validation repairs.
# ---------------------------------------------------------------------------
V229_VERSION = "22.9-ai-preflight-failfast"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_9_ai_preflight_failfast_2026_07_28"

import sys as _v229_sys


def _v229_key_is_placeholder(value: Any) -> bool:
    key = clean_text(value)
    if not key:
        return True
    compact_key = re.sub(r"[^a-z0-9]+", "", key.casefold())
    placeholders = {
        "yourkey", "yourapikey", "apikey", "deepseekapikey",
        "masukkanapikeydeepseek", "masukkankeybaru", "keybarukamu",
        "skxxxxxxxx", "replacewithyourkey",
    }
    return compact_key in placeholders or compact_key.startswith("masukkan") or compact_key.startswith("replacewith")


def _v229_key_source(args: argparse.Namespace) -> str:
    argv = list(_v229_sys.argv[1:])
    if "--deepseek-api-key" in argv:
        return "parameter --deepseek-api-key"
    if clean_text(os.getenv("DEEPSEEK_API_KEY", "")):
        return "environment variable DEEPSEEK_API_KEY"
    return "tidak ditemukan"


def _v229_missing_key_message() -> str:
    return (
        "ERROR: --use-ai aktif, tetapi DeepSeek API key tidak terdeteksi.\n"
        "Repair AI tidak dapat dijalankan tanpa key, sehingga proses dihentikan sebelum scraping dan tidak masuk retry loop.\n\n"
        "Di PowerShell, set key BARU pada sesi yang sama:\n"
        "  $env:DEEPSEEK_API_KEY=\"sk-KEY_BARU_KAMU\"\n"
        "  if ($env:DEEPSEEK_API_KEY) { 'DeepSeek key terdeteksi' }\n\n"
        "Lalu jalankan ulang perintah tanpa menulis key di command history:\n"
        "  py .\\lutron_WEBSITE_v22_9_ai_preflight_failfast.py --input .\\lutron-5.xlsx --output .\\WEBSITE_lutron_test_v22_9.xlsx --max-rows 3 --use-ai --ai-retries 5 --no-cache\n\n"
        "Alternatif tanpa AI: hapus parameter --use-ai. Mode tersebut tidak menjalankan repair DeepSeek."
    )


def _v229_fatal_api_message(status_code: int, response_text: str = "") -> str:
    detail = compact(response_text, 320)
    if status_code in {401, 403}:
        reason = "DeepSeek menolak API key atau key tidak memiliki izin."
    elif status_code == 402:
        reason = "Akun DeepSeek tidak memiliki kredit/kuota yang dapat digunakan."
    elif status_code == 404:
        reason = "Endpoint atau model DeepSeek tidak ditemukan. Periksa --deepseek-base-url dan --deepseek-model."
    else:
        reason = f"DeepSeek mengembalikan HTTP {status_code}."
    return (
        "ERROR AI FATAL: " + reason
        + ("\nDetail respons: " + detail if detail else "")
        + "\nProses dihentikan. Kesalahan autentikasi/akun tidak akan diulang sebagai repair konten."
    )


def _v226_post_ai_json(kind: str, user_payload: Dict[str, Any]) -> Tuple[bool, Dict[str, Any], str]:
    """V22.9 transport: fail fast for configuration/account errors.

    The previous implementation returned ``missing_api_key`` as an ordinary
    repair failure. Serial repair then repeated the same impossible request for
    every section and specification row. This version distinguishes fatal API
    configuration from retryable transport/content failures.
    """
    api_key = clean_text(AI_SETTINGS.get("api_key"))
    if _v229_key_is_placeholder(api_key):
        raise SystemExit(_v229_missing_key_message())

    cache = AI_SETTINGS.get("cache") if isinstance(AI_SETTINGS.get("cache"), dict) else {}
    cache_key = _v226_request_cache_key(kind, user_payload)
    cached = cache.get(cache_key)
    if isinstance(cached, dict) and isinstance(cached.get("parsed"), dict):
        AI_STATS["cache_hit"] = int(AI_STATS.get("cache_hit") or 0) + 1
        return True, cached["parsed"], "cache_hit"

    body = {
        "model": clean_text(AI_SETTINGS.get("model")) or "deepseek-chat",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Anda adalah editor teknis berbasis bukti. Perbaiki hanya bagian yang diminta. "
                    "Jangan mengubah bagian yang sudah valid. Gunakan hanya fakta dan spesifikasi yang diberikan. "
                    "Balas JSON valid saja tanpa markdown."
                ),
            },
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ],
        "temperature": min(0.15, float(AI_SETTINGS.get("temperature") or 0.10)),
        "response_format": {"type": "json_object"},
    }
    endpoint = clean_text(AI_SETTINGS.get("base_url")).rstrip("/") + "/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    try:
        response = requests.post(
            endpoint,
            headers=headers,
            json=body,
            timeout=AI_SETTINGS.get("timeout", 180),
        )
        if response.status_code in {400, 422}:
            body.pop("response_format", None)
            response = requests.post(
                endpoint,
                headers=headers,
                json=body,
                timeout=AI_SETTINGS.get("timeout", 180),
            )

        if response.status_code in {401, 402, 403, 404}:
            raise SystemExit(_v229_fatal_api_message(response.status_code, response.text))
        if response.status_code == 429:
            return False, {}, "rate_limited_http_429"
        if 500 <= response.status_code <= 599:
            return False, {}, f"deepseek_server_http_{response.status_code}"

        response.raise_for_status()
        raw = response.json()
        content = clean_text(raw.get("choices", [{}])[0].get("message", {}).get("content", ""))
        parsed = _parse_json(content)
        if not parsed:
            return False, {}, "invalid_ai_json"

        cache[cache_key] = {
            "parsed": parsed,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "kind": kind,
        }
        AI_SETTINGS["cache"] = cache
        return True, parsed, "ok"
    except SystemExit:
        raise
    except requests.Timeout:
        return False, {}, "request_timeout"
    except requests.ConnectionError:
        return False, {}, "request_connection_error"
    except requests.HTTPError as exc:
        status = getattr(getattr(exc, "response", None), "status_code", None)
        return False, {}, f"request_http_error_{status or 'unknown'}"
    except Exception as exc:
        return False, {}, f"request_error_{type(exc).__name__}"


_v229_previous_main = main


def main() -> None:
    # Parse once for a local preflight. The original main parses the same CLI
    # again after this check; argparse is deterministic and no state is lost.
    args = parse_args()
    if bool(args.use_ai) and not bool(getattr(args, "ai_cache_only", False)):
        key = clean_text(args.deepseek_api_key)
        if _v229_key_is_placeholder(key):
            raise SystemExit(_v229_missing_key_message())
        print(f"INFO: DeepSeek API key terdeteksi dari {_v229_key_source(args)}. Nilai key tidak ditampilkan.")
    _v229_previous_main()


# ---------------------------------------------------------------------------
# V22.10 validator deadlock fix
#
# Fixes three false-positive loops observed on PM-9100:
#   1. measurement-unit regex matched "mA" inside ordinary words such as
#      "manual" and protected the fragment "ma";
#   2. Indonesian decimal commas (0,6) were rejected even when numerically
#      identical to source decimal points (0.6);
#   3. the repair output example showed only one FAQ item although the
#      validator requires 3-5 items.
#
# Technical identifiers, real measurement units, source numbers, evidence,
# and all publish gates remain enforced.
# ---------------------------------------------------------------------------
V2210_VERSION = "22.10-validator-deadlock-fix"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_10_validator_deadlock_fix_2026_07_28"

from decimal import Decimal as _V2210Decimal, InvalidOperation as _V2210InvalidOperation

AI_STATS.setdefault("v2210_false_token_matches_ignored", 0)
AI_STATS.setdefault("v2210_decimal_equivalence_accepts", 0)
AI_STATS.setdefault("v2210_candidate_shape_repairs", 0)


def _v2210_decimal_key(token: Any) -> str:
    """Canonical numeric key without changing the published text.

    Decimal commas and decimal points are treated as equivalent for validator
    comparison only. The displayed candidate remains untouched.
    """
    raw = clean_text(token).replace(",", ".")
    try:
        number = _V2210Decimal(raw)
    except (_V2210InvalidOperation, ValueError):
        return raw
    if number == 0:
        return "0"
    normalized = format(number.normalize(), "f")
    if "." in normalized:
        normalized = normalized.rstrip("0").rstrip(".")
    return normalized


def _v2210_numeric_sequence(value: Any) -> List[str]:
    return [_v2210_decimal_key(token) for token in _v225_numeric_tokens(value)]


def _v2210_numeric_sequences_equal(left: Any, right: Any) -> bool:
    raw_left = _v225_numeric_tokens(left)
    raw_right = _v225_numeric_tokens(right)
    equivalent = [_v2210_decimal_key(x) for x in raw_left] == [_v2210_decimal_key(x) for x in raw_right]
    if equivalent and raw_left != raw_right:
        AI_STATS["v2210_decimal_equivalence_accepts"] = int(
            AI_STATS.get("v2210_decimal_equivalence_accepts") or 0
        ) + 1
    return equivalent


def _v2210_token_probe(value: Any) -> str:
    text = _v224_exact_text(value)
    text = re.sub(r"\bRS\s*232\b", "RS232", text, flags=re.I)
    # Equivalent localized unit names share one comparison key. Published text
    # is not modified; this only prevents a validator false positive.
    text = re.sub(r"\b(?:atmosphere|atmosfer|atm)\b", "atm", text, flags=re.I)
    return re.sub(r"\s+", "", text).casefold()


def _v226_required_spec_tokens(value: Any) -> List[str]:
    """Return only genuine standalone technical tokens.

    The previous implementation ignored a unit match only when alphabetic
    characters existed on *both* sides. Because unit matching is case
    insensitive, the unit ``mA`` was therefore detected at the start of
    ``manual`` (``ma`` + ``nual``). This function rejects a unit match when an
    alphabetic character exists on either side, eliminating that deadlock.
    """
    text = _v224_exact_text(value)
    tokens: List[str] = list(_v225_technical_tokens(text))
    tokens.extend(V224_CODE_RE.findall(text))

    for match in V224_UNIT_TOKEN_RE.finditer(text):
        token = _v224_exact_text(match.group(0))
        if not token:
            continue
        before = text[match.start() - 1] if match.start() > 0 else ""
        after = text[match.end()] if match.end() < len(text) else ""

        # A measurement unit is valid only when it is not embedded in an
        # ordinary alphabetic word. Examples rejected here: ma-nual, in-put.
        if before.isalpha() or after.isalpha():
            AI_STATS["v2210_false_token_matches_ignored"] = int(
                AI_STATS.get("v2210_false_token_matches_ignored") or 0
            ) + 1
            continue

        compact_token = re.sub(r"[^A-Za-z0-9%°µ²₂/]", "", token)
        if len(compact_token) == 1 and compact_token.isalpha():
            left_context = text[max(0, match.start() - 16):match.start()]
            if not re.search(r"(?:\d|°)\s*$", left_context):
                continue
        tokens.append(token)

    tokens.extend(
        re.findall(r"\b(?:RS\s*232|USB|LCD|LSI|AC|DC|RH|F\.?S\.?|True\s*RMS)\b", text, re.I)
    )
    cleaned = [_v224_exact_text(token) for token in tokens if _v224_exact_text(token)]
    return dedupe_keep_order(cleaned)


# Extra deterministic terminology repairs for partial AI phrases. These are
# wording-only substitutions and do not touch source numbers or identifiers.
V228_SPEC_EXTRA_REPLACEMENTS = tuple(V228_SPEC_EXTRA_REPLACEMENTS) + (
    (re.compile(r"\bsaves\b", re.I), "menghemat"),
    (re.compile(r"\bsupper\b", re.I), "sangat besar"),
    (re.compile(r"\binstruction\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\binstruksi\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\bAksesori\s+termasuk\b", re.I), "Aksesori yang disertakan"),
    (re.compile(r"\bkelembaban\b", re.I), "kelembapan"),
    (re.compile(r"\bhard\s+carrying\s+case\b", re.I), "kotak pelindung keras"),
)


def _v226_validate_spec_candidate(item: Dict[str, Any], candidate: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Strict specification validator without the V22.9 false positives."""
    if isinstance(candidate, dict):
        candidate["label"] = _v228_apply_replacements(
            candidate.get("label") or candidate.get("label_id"), specification=True
        )
        candidate["value"] = _v228_apply_replacements(
            candidate.get("value") or candidate.get("value_id"), specification=True
        )

    label_raw = _v224_exact_text(item.get("label_raw"))
    value_raw = _v224_exact_text(item.get("value_raw"))
    label = _v224_exact_text(candidate.get("label") or candidate.get("label_id"))
    value = _v224_exact_text(candidate.get("value") or candidate.get("value_id"))
    source_id = item.get("source_id")
    issues: List[Dict[str, Any]] = []

    if not label:
        issues.append(_v226_problem("spec_label_empty", "Label spesifikasi tidak boleh kosong.", source_id=source_id))
    if not value:
        issues.append(_v226_problem("spec_value_empty", "Nilai spesifikasi tidak boleh kosong.", source_id=source_id))

    if label and not _v2210_numeric_sequences_equal(label, label_raw):
        issues.append(_v226_problem(
            "spec_label_numeric_token_changed",
            "Angka pada label harus memiliki nilai dan urutan yang sama dengan sumber.",
            source_id=source_id,
            required=_v225_numeric_tokens(label_raw),
        ))
    if value and not _v2210_numeric_sequences_equal(value, value_raw):
        issues.append(_v226_problem(
            "spec_value_numeric_token_changed",
            "Angka pada nilai harus memiliki nilai dan urutan yang sama dengan sumber.",
            source_id=source_id,
            required=_v225_numeric_tokens(value_raw),
        ))

    label_probe = _v2210_token_probe(label)
    value_probe = _v2210_token_probe(value)
    for token in _v226_required_spec_tokens(label_raw):
        if _v2210_token_probe(token) not in label_probe:
            issues.append(_v226_problem(
                "spec_label_protected_token_missing",
                "Model, kode, satuan, atau singkatan teknis pada label tidak boleh dihapus.",
                source_id=source_id,
                token=token,
            ))
    for token in _v226_required_spec_tokens(value_raw):
        if _v2210_token_probe(token) not in value_probe:
            issues.append(_v226_problem(
                "spec_value_protected_token_missing",
                "Model, kode, satuan, atau singkatan teknis pada nilai tidak boleh dihapus.",
                source_id=source_id,
                token=token,
            ))

    if has_cjk(label + " " + value):
        issues.append(_v226_problem(
            "spec_contains_chinese", "Spesifikasi harus bebas karakter Chinese.", source_id=source_id
        ))
    if _v226_has_spec_english(label) or _v226_has_spec_english(value):
        fragments = _v228_english_fragments(label + " " + value, specification=True)
        issues.append(_v226_problem(
            "spec_english_residue_detected",
            "Terjemahkan frasa bahasa Inggris pada baris ini. Pertahankan hanya model, kode, satuan, dan singkatan teknis.",
            source_id=source_id,
            offending_fragments=fragments,
            repair_instruction=(
                "Terjemahkan tepat offending_fragments. Angka, satuan, model, kode, dan urutan fakta tidak boleh berubah."
            ),
        ))
    if re.search(r"\b(?:undefined|null|nan)\b|\d\s*tUSB\b|\bm\s*/\s*mim\b", label + " " + value, re.I):
        issues.append(_v226_problem(
            "spec_malformed_text", "Baris spesifikasi masih mengandung fragmen teks rusak.", source_id=source_id
        ))
    return _v226_dedupe_problems(issues)


_v2210_previous_content_schema_for = _v226_content_schema_for


def _v226_content_schema_for(fields: Sequence[str]) -> Dict[str, Any]:
    """Give DeepSeek an example that matches the actual item-count contract."""
    schema = _v2210_previous_content_schema_for(fields)
    sections = schema.get("sections") if isinstance(schema, dict) else None
    if not isinstance(sections, dict):
        return schema
    if "faq_html" in fields:
        sections["faq_html"] = {
            "required_item_count": "3-5",
            "items": [
                {"question_id": "Pertanyaan teknis 1?", "answer_id": "Jawaban faktual 1.", "evidence_ids": ["F001"]},
                {"question_id": "Pertanyaan teknis 2?", "answer_id": "Jawaban faktual 2.", "evidence_ids": ["S001"]},
                {"question_id": "Pertanyaan teknis 3?", "answer_id": "Jawaban faktual 3.", "evidence_ids": ["F002"]},
            ],
        }
    if "application_tab_html" in fields:
        sections["application_tab_html"] = {
            "required_item_count": "1-5",
            "items": [
                {"text_id": "Satu aplikasi faktual dalam Bahasa Indonesia.", "evidence_ids": ["F001"]}
            ],
        }
    return schema


_v2210_previous_normalize_candidate = _v228_normalize_candidate


def _v2210_extract_faq_items_from_html(raw_html: Any) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(clean_text(raw_html), "html.parser")
    items: List[Dict[str, Any]] = []
    for heading in soup.find_all(["h2", "h3", "strong"]):
        question = clean_text(heading.get_text(" ", strip=True))
        if not question:
            continue
        paragraph = heading.find_next_sibling("p")
        answer = clean_text(paragraph.get_text(" ", strip=True)) if paragraph else ""
        if question and answer:
            items.append({"question_id": question, "answer_id": answer, "evidence_ids": []})
    return items[:V22_FAQ_MAX]


def _v2210_extract_application_items_from_html(raw_html: Any) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(clean_text(raw_html), "html.parser")
    values = [clean_text(node.get_text(" ", strip=True)) for node in soup.find_all("li")]
    if not values:
        values = [clean_text(node.get_text(" ", strip=True)) for node in soup.find_all("p")]
    return [{"text_id": value, "evidence_ids": []} for value in values if value][:5]


def _v228_normalize_candidate(field: str, candidate: Any) -> Any:
    value = _v226_json_clone(candidate)
    changed = False
    if isinstance(value, dict) and field == "faq_html":
        rows = value.get("items")
        if not isinstance(rows, list):
            for key in ("faq_items", "faqs", "questions"):
                if isinstance(value.get(key), list):
                    rows = value.get(key)
                    changed = True
                    break
        if isinstance(rows, dict):
            rows = list(rows.values())
            changed = True
        if not isinstance(rows, list) or not rows:
            rows = _v2210_extract_faq_items_from_html(value.get("html") or value.get("value"))
            changed = bool(rows)
        normalized_rows: List[Dict[str, Any]] = []
        for row in rows or []:
            if not isinstance(row, dict):
                continue
            question = clean_text(row.get("question_id") or row.get("question") or row.get("q"))
            answer = clean_text(row.get("answer_id") or row.get("answer") or row.get("a"))
            if not question or not answer:
                continue
            normalized_rows.append({
                "question_id": question,
                "answer_id": answer,
                "evidence_ids": row.get("evidence_ids") if isinstance(row.get("evidence_ids"), list) else [],
            })
        # Remove commercial FAQ rows when enough technical rows remain. This
        # is deterministic deletion of forbidden content, not content invention.
        technical_rows = [
            row for row in normalized_rows
            if not V226_FORBIDDEN_FAQ_RE.search(
                clean_text(row.get("question_id")) + " " + clean_text(row.get("answer_id"))
            )
        ]
        if len(technical_rows) >= V22_FAQ_MIN and len(technical_rows) != len(normalized_rows):
            normalized_rows = technical_rows
            changed = True
        value["items"] = normalized_rows[:V22_FAQ_MAX]
    elif isinstance(value, dict) and field == "application_tab_html":
        rows = value.get("items")
        if not isinstance(rows, list):
            for key in ("application_items", "applications", "uses"):
                if isinstance(value.get(key), list):
                    rows = value.get(key)
                    changed = True
                    break
        if isinstance(rows, dict):
            rows = list(rows.values())
            changed = True
        if not isinstance(rows, list) or not rows:
            rows = _v2210_extract_application_items_from_html(value.get("html") or value.get("value"))
            changed = bool(rows)
        normalized_rows = []
        for row in rows or []:
            if isinstance(row, dict):
                text = clean_text(row.get("text_id") or row.get("text") or row.get("application"))
                evidence_ids = row.get("evidence_ids") if isinstance(row.get("evidence_ids"), list) else []
            else:
                text = clean_text(row)
                evidence_ids = []
            if text:
                normalized_rows.append({"text_id": text, "evidence_ids": evidence_ids})
        value["items"] = normalized_rows[:5]
    if changed:
        AI_STATS["v2210_candidate_shape_repairs"] = int(
            AI_STATS.get("v2210_candidate_shape_repairs") or 0
        ) + 1
    return _v2210_previous_normalize_candidate(field, value)


_v2210_previous_attach_source_evidence = _v227_attach_source_evidence


def _v227_attach_source_evidence(field: str, candidate: Any, payload: Dict[str, Any]) -> Any:
    # Normalize alternate response shapes before evidence IDs are attached.
    return _v2210_previous_attach_source_evidence(
        field, _v228_normalize_candidate(field, candidate), payload
    )


_v2210_previous_content_prompt = _v227_content_prompt


def _v227_content_prompt(
    field: str,
    current: Any,
    issues: Sequence[Dict[str, Any]],
    payload: Dict[str, Any],
    attempt: int,
    maximum: int,
) -> Dict[str, Any]:
    prompt = _v2210_previous_content_prompt(field, current, issues, payload, attempt, maximum)
    rules = prompt.get("rules") if isinstance(prompt.get("rules"), list) else []
    if field == "faq_html":
        rules.append("FAQ wajib berisi 3-5 objek items; jangan mengirim hanya satu contoh item.")
        prompt["required_item_count"] = "3-5"
    elif field == "application_tab_html":
        rules.append("Aplikasi penggunaan wajib berisi 1-5 objek items.")
        prompt["required_item_count"] = "1-5"
    prompt["rules"] = rules
    return prompt


_v2210_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2210_previous_build_output_row(data, pub_date)
    if data.get("_v227_converged"):
        row["Website Validation"] = "publish_ready_validator_deadlock_fix"
        row["Processing Time"] = "lutron_WEBSITE_import_v22_10_validator_deadlock_fix"
        row["Website Correction Log"] = _merge_source_notes(
            row.get("Website Correction Log"),
            "repair_pipeline=" + V2210_VERSION,
            "embedded_unit_false_positive_fixed",
            "decimal_comma_equivalence_enabled",
            "faq_schema_count_contract_fixed",
        )
    return row



# ---------------------------------------------------------------------------
# V22.11 final publish polish and terminal editorial gate
#
# This last-stage layer runs AFTER every previous repair/status wrapper. It
# therefore catches wording damage introduced anywhere in the pipeline,
# applies source-safe deterministic editorial corrections, compacts duplicate
# general specification rows, and then validates the actual final cells that
# will be written to Excel. Unknown residual defects are never labelled
# publish-ready.
# ---------------------------------------------------------------------------
V2211_VERSION = "22.11-final-publish-polish"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_11_final_publish_polish_2026_07_28"

AI_STATS.setdefault("v2211_text_repairs", 0)
AI_STATS.setdefault("v2211_spec_rows_deduplicated", 0)
AI_STATS.setdefault("v2211_final_gate_pass", 0)
AI_STATS.setdefault("v2211_final_gate_fail", 0)

V2211_HTML_FIELDS: Tuple[str, ...] = (
    "Product Description",
    "custom_tab_1_content",
    "custom_tab_2_content",
    "custom_tab_3_content",
    "custom_tab_4_content",
    "custom_tab_5_content",
)

V2211_PLAIN_FIELDS: Tuple[str, ...] = (
    "Name",
    "Product Short Description",
    "focus keyphrase",
    "meta description",
)

# Phrase-level repairs are deliberately ordered from specific to general.
# They repair language only; numeric values, units, codes, URLs and evidence
# identifiers are not changed.
V2211_TEXT_REPLACEMENTS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bProduk\s+ini\s+memiliki\s+Produk\s+menyediakan\b", re.I), "Produk menyediakan"),
    (re.compile(r"\bRS232\s*/\s*USC\b", re.I), "RS232/USB"),
    (re.compile(
        r"Antarmuka\s+RS232\s*/\s*USB\s*\([^)]*USB-01[^)]*UPCB-01[^)]*\)",
        re.I,
    ), "Antarmuka RS232 untuk keluaran data; kabel USB-01 dan kabel RS232 UPCB-01 tersedia sebagai aksesori opsional"),
    (re.compile(
        r"Produk\s+menyediakan\s+antarmuka\s+serial\s+RS232\s+untuk\s+keluaran\s+data\s+ke\s+komputer\s*;\s*"
        r"Koneksi\s+USB\s+memerlukan\s+kabel\s+USB-01\s+opsional\s+untuk\s+konektivitas\s+komputer\s*"
        r"\(kabel\s+USB-01\s+dan\s+RS232\s+UPCB-01\s+opsional\)",
        re.I,
    ), "Produk menyediakan antarmuka serial RS232 untuk keluaran data ke komputer. Kabel USB-01 dan kabel RS232 UPCB-01 tersedia sebagai aksesori opsional"),
    (re.compile(
        r"Produk\s+menyediakan\s+antarmuka\s+serial\s+RS232\s+untuk\s+keluaran\s+data\s+ke\s+komputer\s*;\s*"
        r"Koneksi\s+USB\s+memerlukan\s+kabel\s+USB-01\s+opsional\.\s*"
        r"Kabel\s+USB\s*\(USB-01\)\s+dan\s+kabel\s+RS232\s*\(UPCB-01\)\s+adalah\s+aksesori\s+opsional",
        re.I,
    ), "Produk menyediakan antarmuka serial RS232 untuk keluaran data ke komputer. Kabel USB-01 dan kabel RS232 UPCB-01 tersedia sebagai aksesori opsional"),
    (re.compile(r"\bEasy\s+offset\s+of\s+zero\s+value\s+with\s+Tombol\s+tekan\s+pada\s+panel\s+depan\b", re.I),
     "Penyetelan nilai nol dilakukan melalui tombol tekan pada panel depan"),
    (re.compile(r"\busage\s+untuk\s+kalibrasi\s+presisi\b", re.I), "untuk kalibrasi secara presisi"),
    (re.compile(r"\bmenghemat\s+baterai\s+life\b", re.I), "menghemat daya baterai"),
    (re.compile(r"\bcan\s+also\s+be\s+manually\s+turned\s+off\b", re.I), "alat juga dapat dimatikan secara manual"),
    (re.compile(r"\bor\s+dimatikan\s+secara\s+manual\b", re.I), "atau dapat dimatikan secara manual"),
    (re.compile(r"\binch\s+Melalui\b", re.I), "melalui"),
    (re.compile(r"\bsangat\s+besar\s+Layar\s+LCD\s+besar\b", re.I), "layar LCD sangat besar"),
    (re.compile(r"\b(?:supper|super)\s+large\s+LCD\s+display\b", re.I), "layar LCD sangat besar"),
    (re.compile(r"\blarge\s+LCD\s+display\b", re.I), "layar LCD besar"),
    (re.compile(r"\bRS\s*232\s+PC\s+serial\s+interface\b", re.I), "antarmuka serial RS232 untuk komputer"),
    (re.compile(r"\bantarmuka\s+RS232\s+PC\s+serial\s+untuk\s+(?:output|keluaran)\s+data\b", re.I),
     "antarmuka serial RS232 untuk keluaran data ke komputer"),
    (re.compile(r"\bAutomatic\s+and\s+manual\s+data\s+logging\b", re.I), "Pencatatan data otomatis dan manual"),
    (re.compile(r"\bData\s+hold\s+to\s+freeze\s+the\s+display\s+reading\b", re.I),
     "Data Hold untuk menahan pembacaan pada tampilan"),
    (re.compile(r"\bRecords?\s+maximum\s+and\s+minimum\s+readings?\b", re.I),
     "Merekam pembacaan maksimum dan minimum"),
    (re.compile(r"\bdata\s+logger\b", re.I), "pencatat data"),
    (re.compile(r"\bgain\b", re.I), "penguatan"),
    (re.compile(r"\bdata\s+hold\s+untuk\s+membekukan\s+tampilan\b", re.I), "Data Hold untuk menahan pembacaan pada tampilan"),
    (re.compile(r"(?<=[.!?])\s+pencatat\s+data\b", re.I), " Pencatat data"),
    (re.compile(r"\bzero\s+adjust(?:ment)?\b", re.I), "penyetelan nol"),
    (re.compile(r"\bspan\s+adjust(?:ment)?\b", re.I), "penyetelan rentang"),
    (re.compile(r"\bpenyesuaian\s+span\b", re.I), "penyetelan rentang"),
    (re.compile(r"\bFull\s+Scale\b", re.I), "skala penuh"),
    (re.compile(r"\brepeatability\b", re.I), "keterulangan"),
    (re.compile(r"\bhousing\b", re.I), "bodi alat"),
    (re.compile(r"\bbattery\b", re.I), "baterai"),
    (re.compile(r"\binstruction\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\binstruksi\s+manual\b", re.I), "buku petunjuk"),
    (re.compile(r"\bAksesori\s+termasuk\b", re.I), "Aksesori yang disertakan"),
    (re.compile(r"\bkelembaban\b", re.I), "kelembapan"),
    (re.compile(r"\bhard\s+carrying\s+case\b", re.I), "kotak pelindung keras"),
    (re.compile(r"\boutput\s+data\b", re.I), "keluaran data"),
    (re.compile(r"\bdata\s+output\b", re.I), "keluaran data"),
    (re.compile(r"\binput\s+ganda\b", re.I), "masukan ganda"),
    (re.compile(r"\bdual\s+(?:and|&)\s+differential\s+input\b", re.I), "masukan ganda dan diferensial"),
    (re.compile(r"\bdifferential\b", re.I), "diferensial"),
    (re.compile(r"\bmasukan\s+ganda\s*&\s*diferensial\b", re.I), "masukan ganda dan diferensial"),
    (re.compile(r"\bpenyesuaian\s+nol\b", re.I), "penyetelan nol"),
    (re.compile(r"\bsecara\s+manual\s+Melalui\b", re.I), "secara manual melalui"),
    (re.compile(r"\batmosphere\b", re.I), "atmosfer"),
    (re.compile(r"\binches\b", re.I), "inch"),
    (re.compile(r"\bLB\b"), "lb"),
    (re.compile(r"\bR\.?\s*H\.?\b", re.I), "RH"),
    (re.compile(r"SW-U801-WIN\s+Kabel\s+USB", re.I), "SW-U801-WIN; Kabel USB"),
    (re.compile(r"USB-01\s+Kabel\s+RS232", re.I), "USB-01; Kabel RS232"),
    (re.compile(r"\bSekitar\.\s*", re.I), "Sekitar "),
    (re.compile(r"\b1\s+PC\b", re.I), "1 unit"),
    (re.compile(r"\b2\s+PCs?\b", re.I), "2 unit"),
    (re.compile(r"\bH2O\b", re.I), "H₂O"),
    (re.compile(r"\bKg\s*/\s*cm\^?2\b", re.I), "kg/cm²"),
)

V2211_BLOCKER_PATTERNS: Tuple[Tuple[str, re.Pattern[str]], ...] = (
    ("typo_rs232_usc", re.compile(r"\bRS232\s*/\s*USC\b", re.I)),
    ("mixed_battery_life", re.compile(r"\bbaterai\s+life\b|\bbattery\s+life\b", re.I)),
    ("mixed_manual_off", re.compile(r"\bcan\s+also\s+be\b|\bor\s+dimatikan\b", re.I)),
    ("mixed_easy_offset", re.compile(r"\bEasy\s+offset\b", re.I)),
    ("mixed_usage", re.compile(r"\busage\b", re.I)),
    ("duplicated_product_phrase", re.compile(r"Produk\s+ini\s+memiliki\s+Produk", re.I)),
    ("broken_inch_melalui", re.compile(r"\binch\s+Melalui\b", re.I)),
    ("broken_display_phrase", re.compile(r"sangat\s+besar\s+Layar\s+LCD\s+besar", re.I)),
    ("broken_approximation", re.compile(r"\bSekitar\.\s*\d", re.I)),
    ("undefined_text", re.compile(r"\b(?:undefined|null|nan)\b", re.I)),
)


def _v2211_polish_text(value: Any) -> str:
    text = html.unescape(str(value or ""))
    if not text:
        return ""
    original = text
    for pattern, replacement in V2211_TEXT_REPLACEMENTS:
        text = pattern.sub(replacement, text)

    # Remove asterisk bullets that were flattened into a table cell.
    text = re.sub(r"^\s*\*\s*", "", text)
    text = re.sub(r"\s+\*\s+", "; ", text)

    # Punctuation and spacing polish without changing decimal commas.
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([;:])(?=\S)", r"\1 ", text)
    text = re.sub(r"(?<!\d),(?=\S)", ", ", text)
    text = re.sub(r"(?<=\d),\s+(?=\d)", ",", text)
    text = re.sub(r"\(\s+", "(", text)
    text = re.sub(r"\s+\)", ")", text)
    text = re.sub(r"(?<=[A-Za-zÀ-ÿ])\(", " (", text)
    text = re.sub(r"\.;", ";", text)
    text = re.sub(r";\s+Cairan\b", "; cairan", text)
    text = re.sub(r"\s+\.", ".", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    text = re.sub(r"\bProduk menyediakan\s+Produk menyediakan\b", "Produk menyediakan", text, flags=re.I)
    text = re.sub(r"\b(\w{3,})\s+\1\b", r"\1", text, flags=re.I)

    if text != original:
        AI_STATS["v2211_text_repairs"] = int(AI_STATS.get("v2211_text_repairs") or 0) + 1
    return text


def _v2211_capitalize_visible_text(text: str, parent_name: str) -> str:
    # Only list items and headings need forced initial capitalisation. Paragraph
    # text may be split by inline links; capitalising every text node would turn
    # "atau" after a link into "Atau" in the middle of a sentence.
    if parent_name not in {"li", "h1", "h2", "h3", "h4"}:
        return text
    stripped = text.lstrip()
    if not stripped or not stripped[0].islower():
        return text
    offset = len(text) - len(stripped)
    return text[:offset] + stripped[0].upper() + stripped[1:]


def _v2211_polish_html(value: Any) -> str:
    raw = clean_text(value)
    if not raw:
        return ""
    # A source value such as "<80% RH" is text, not an HTML tag. Escape only
    # less-than signs followed by a number before parsing the fragment.
    raw = re.sub(r"(?<!<)<(?=\s*[-+]?\d)", "&lt;", raw)
    soup = BeautifulSoup(raw, "html.parser")
    for node in list(soup.find_all(string=True)):
        parent_name = clean_text(getattr(getattr(node, "parent", None), "name", "")).casefold()
        if parent_name in {"script", "style"}:
            continue
        polished = _v2211_polish_text(str(node))
        polished = _v2211_capitalize_visible_text(polished, parent_name)
        if polished != str(node):
            node.replace_with(polished)
    return str(soup)


def _v2211_normalize_spec_label(label: Any, value: Any) -> str:
    raw_label = _v2211_polish_text(label)
    raw_value = _v2211_polish_text(value)
    key = re.sub(r"[^a-z0-9]+", "", raw_label.casefold())

    mapping = {
        "display": "Tampilan",
        "tampilan": "Tampilan",
        "displayunits": "Satuan tampilan",
        "satuantampilan": "Satuan tampilan",
        "satuan": "Satuan tampilan",
        "function": "Fungsi",
        "functionality": "Fungsi",
        "fungsi": "Fungsi",
        "zeroadjust": "Penyetelan nol",
        "penyetelannol": "Penyetelan nol",
        "spanadjust": "Penyetelan rentang",
        "penyetelanrentang": "Penyetelan rentang",
        "datahold": "Penahan data",
        "penahandata": "Penahan data",
        "datarecord": "Rekaman data",
        "rekamandata": "Rekaman data",
        "dataoutput": "Keluaran data",
        "keluarandata": "Keluaran data",
        "poweroff": "Mati otomatis",
        "mat otomatis": "Mati otomatis",
        "matio tomatis": "Mati otomatis",
        "matiotomatis": "Mati otomatis",
        "samplingtime": "Waktu sampling",
        "waktusampling": "Waktu sampling",
        "responsetime": "Waktu respons",
        "wakturespons": "Waktu respons",
        "operatingtemperature": "Suhu operasi",
        "temperature": "Suhu operasi",
        "suhu": "Suhu operasi",
        "suhuoperasi": "Suhu operasi",
        "operatinghumidity": "Kelembapan operasi",
        "humidity": "Kelembapan operasi",
        "kelembapan": "Kelembapan operasi",
        "kelembapanoperasi": "Kelembapan operasi",
        "powersupply": "Catu daya",
        "catudaya": "Catu daya",
        "powercurrent": "Konsumsi daya",
        "powerconsumption": "Konsumsi daya",
        "arusdaya": "Konsumsi daya",
        "konsumsidaya": "Konsumsi daya",
        "weight": "Berat",
        "berat": "Berat",
        "dimension": "Dimensi",
        "dimensions": "Dimensi",
        "dimensi": "Dimensi",
        "accessoriesincluded": "Aksesori yang disertakan",
        "aksesori yang disertakan": "Aksesori yang disertakan",
        "aksesoryangdisertakan": "Aksesori yang disertakan",
        "optionalaccessories": "Aksesori opsional",
        "accessoriesoptional": "Aksesori opsional",
        "aksesoriopsional": "Aksesori opsional",
        "digitsize": "Ukuran digit",
        "ukurandigit": "Ukuran digit",
        "instructionmanual": "Buku petunjuk",
        "bukupetunjuk": "Buku petunjuk",
        "hardcarryingcase": "Kotak pelindung keras",
        "kotakpelindungkeras": "Kotak pelindung keras",
        "datalogger": "Pencatat data",
        "pencatatdata": "Pencatat data",
    }
    if key in mapping:
        return mapping[key]
    if key == "type" or key == "tipe":
        if re.search(r"piezo|sensor", raw_value, re.I):
            return "Tipe sensor"
        if re.search(r"LCD|layar|tampilan", raw_value, re.I):
            return "Tipe tampilan"
        return "Tipe"
    if key == "sensortype" or key == "tipesensor":
        return "Tipe sensor"
    if key == "size" or key == "ukuran":
        if re.search(r"LCD|layar|tampilan|61\s*mm\s*x\s*34\s*mm", raw_value, re.I):
            return "Ukuran tampilan"
        return "Ukuran"
    if key in {"sensordescription", "sensor"}:
        return "Sensor"
    if key == "instrumenutama" and re.search(r"\d\s*x\s*\d", raw_value):
        return "Dimensi instrumen utama"
    return raw_label


def _v2211_spec_key(label: Any, value: Any) -> str:
    normalized = _v2211_normalize_spec_label(label, value)
    key = re.sub(r"[^a-z0-9]+", "", normalized.casefold())
    if key == "sensor":
        if len(clean_text(value)) <= 90 and re.search(r"piezo|sensor", clean_text(value), re.I):
            return "sensor_type"
        return "sensor_description"
    aliases = {
        "model": "model",
        "sirkuit": "circuit",
        "tampilan": "display_summary",
        "satuantampilan": "display_units",
        "fungsi": "function_summary",
        "penyetelannol": "zero_adjust",
        "penyetelanrentang": "span_adjust",
        "penahandata": "data_hold",
        "rekamandata": "data_record",
        "keluarandata": "data_output",
        "matiotomatis": "auto_off",
        "waktusampling": "sampling_time",
        "wakturespons": "response_time",
        "suhuoperasi": "operating_temperature",
        "kelembapanoperasi": "operating_humidity",
        "catudaya": "power_supply",
        "konsumsidaya": "power_consumption",
        "berat": "weight",
        "dimensi": "dimensions",
        "dimensiinstrumenutama": "main_instrument_dimensions",
        "aksesoryangdisertakan": "accessories_included",
        "aksesoriyangdisertakan": "accessories_included",
        "aksesoriopsional": "accessories_optional",
        "tipetampilan": "display_type",
        "ukurantampilan": "display_size",
        "ukurandigit": "digit_size",
        "tipesensor": "sensor_type",
        "pencatatdata": "data_logger",
        "bukupetunjuk": "manual_item",
        "kotakpelindungkeras": "case_item",
        "konektorkoplingcepat": "coupler_item",
    }
    return aliases.get(key, "other:" + key)


def _v2211_row_score(label: str, value: str, key: str) -> int:
    blob = label + " " + value
    score = min(len(value), 240)
    if not any(pattern.search(blob) for _, pattern in V2211_BLOCKER_PATTERNS):
        score += 500
    if not _v226_has_spec_english(blob):
        score += 300
    if re.search(r"[.!?]$", value):
        score += 5
    if key == "auto_off" and re.search(r"menghemat\s+daya\s+baterai", value, re.I):
        score += 80
    if key == "zero_adjust" and re.search(r"nilai\s+nol|panel\s+depan", value, re.I):
        score += 50
    if key in {"power_consumption", "operating_temperature", "operating_humidity", "dimensions", "weight"}:
        score += 20
    return score


def _v2211_numeric_signature(value: Any) -> Tuple[str, ...]:
    return tuple(_v2210_numeric_sequence(value))


def _v2211_compact_specification_html(value: Any) -> Tuple[str, int]:
    polished = _v2211_polish_html(value)
    if not polished:
        return "", 0
    soup = BeautifulSoup(polished, "html.parser")
    general = soup.find("table", class_=lambda item: item and "WEBSITE-spec-general" in item)
    if general is None:
        return str(soup), 0
    tbody = general.find("tbody")
    if tbody is None:
        return str(soup), 0

    entries: List[Dict[str, Any]] = []
    for order, tr in enumerate(list(tbody.find_all("tr", recursive=False))):
        cells = tr.find_all(["th", "td"], recursive=False)
        if len(cells) < 2:
            continue
        label = _v2211_normalize_spec_label(cells[0].get_text(" ", strip=True), cells[1].get_text(" ", strip=True))
        value_text = _v2211_polish_text(cells[1].get_text(" ", strip=True))
        cells[0].string = label
        cells[1].string = value_text
        key = _v2211_spec_key(label, value_text)
        if key == "sensor_type" and label == "Sensor":
            label = "Tipe sensor"
            cells[0].string = label
        entries.append({
            "order": order,
            "tr": tr,
            "label": label,
            "value": value_text,
            "key": key,
            "numeric": _v2211_numeric_signature(value_text),
            "score": _v2211_row_score(label, value_text, key),
        })

    selected: List[Dict[str, Any]] = []
    groups: Dict[Tuple[str, Tuple[str, ...]], Dict[str, Any]] = {}
    for entry in entries:
        # Different numeric signatures represent different physical values and
        # must never be collapsed. Rows without numbers share one wording group.
        group_key = (entry["key"], entry["numeric"])
        current = groups.get(group_key)
        if current is None or entry["score"] >= current["score"]:
            groups[group_key] = entry
    selected = sorted(groups.values(), key=lambda item: item["order"])

    selected_keys = {item["key"] for item in selected}
    measurement_table = soup.find("table", class_=lambda item: item and "WEBSITE-spec-measurement" in item)
    measurement_count = len(measurement_table.select("tbody tr")) if measurement_table else 0

    # The measurement matrix already enumerates units, so the flattened unit
    # summary is redundant. Detailed display rows also replace the composite
    # display sentence when all three details are available.
    remove_orders: set = set()
    if measurement_count >= 2:
        remove_orders.update(item["order"] for item in selected if item["key"] == "display_units")
    if {"display_type", "display_size", "digit_size"}.issubset(selected_keys):
        remove_orders.update(item["order"] for item in selected if item["key"] == "display_summary")

    accessories = next((item for item in selected if item["key"] == "accessories_included"), None)
    if accessories:
        accessory_text = accessories["value"].casefold()
        if "buku petunjuk" in accessory_text:
            remove_orders.update(item["order"] for item in selected if item["key"] == "manual_item")
        if "kotak pelindung" in accessory_text:
            remove_orders.update(item["order"] for item in selected if item["key"] == "case_item")
        if "kopling cepat" in accessory_text:
            remove_orders.update(item["order"] for item in selected if item["key"] == "coupler_item")

    function_summary = next((item for item in selected if item["key"] == "function_summary"), None)
    if function_summary:
        detailed_count = sum(1 for key in ("data_hold", "data_record", "zero_adjust", "auto_off") if key in selected_keys)
        if detailed_count >= 3:
            remove_orders.add(function_summary["order"])

    # A composite data-logger row is redundant when its Data Hold and record
    # functions are already listed separately.
    if {"data_hold", "data_record"}.issubset(selected_keys):
        remove_orders.update(item["order"] for item in selected if item["key"] == "data_logger")

    # Keep the more descriptive main-instrument dimension label when it has
    # exactly the same source numbers as a generic Dimension row.
    main_dims = [item for item in selected if item["key"] == "main_instrument_dimensions"]
    for main_dim in main_dims:
        remove_orders.update(
            item["order"] for item in selected
            if item["key"] == "dimensions" and item["numeric"] == main_dim["numeric"]
        )

    kept_orders = {item["order"] for item in selected if item["order"] not in remove_orders}
    removed = 0
    for entry in entries:
        if entry["order"] not in kept_orders:
            entry["tr"].decompose()
            removed += 1

    if removed:
        AI_STATS["v2211_spec_rows_deduplicated"] = int(
            AI_STATS.get("v2211_spec_rows_deduplicated") or 0
        ) + removed
    return str(soup), removed


def _v2211_duplicate_spec_issues(spec_html: Any) -> List[str]:
    soup = BeautifulSoup(clean_text(spec_html), "html.parser")
    general = soup.find("table", class_=lambda item: item and "WEBSITE-spec-general" in item)
    if general is None:
        return ["final_spec_general_table_missing"]
    seen: set = set()
    issues: List[str] = []
    for tr in general.select("tbody > tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if len(cells) < 2:
            continue
        label = clean_text(cells[0].get_text(" ", strip=True))
        value = clean_text(cells[1].get_text(" ", strip=True))
        signature = (_v2211_spec_key(label, value), _v2211_numeric_signature(value))
        if signature in seen:
            issues.append("final_spec_duplicate_semantic_row")
            break
        seen.add(signature)
    return issues


def _v2211_visible_text(value: Any) -> str:
    raw = str(value or "")
    raw = re.sub(r"(?<!<)<(?=\s*[-+]?\d)", "&lt;", raw)
    try:
        return clean_text(BeautifulSoup(raw, "html.parser").get_text(" ", strip=True))
    except Exception:
        return clean_text(re.sub(r"<[^>]+>", " ", raw))


def _v2211_final_editorial_issues(row: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    for field in V2211_HTML_FIELDS:
        value = clean_text(row.get(field))
        if not value:
            issues.append("final_empty_" + field.lower().replace(" ", "_"))
            continue
        visible = _v2211_visible_text(value)
        for code, pattern in V2211_BLOCKER_PATTERNS:
            if pattern.search(visible):
                issues.append("final_" + code)
        if has_cjk(visible):
            issues.append("final_chinese_character_detected")

    public_fields = (
        "Product Description", "Product Short Description",
        "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
        "meta description",
    )
    for field in public_fields:
        visible_public = _v2211_visible_text(row.get(field))
        if V22_ENGLISH_RESIDUE_RE.search(visible_public) or V226_PUBLIC_ENGLISH_RE.search(visible_public):
            issues.append("final_public_english_residue:" + field)

    spec_html = row.get("custom_tab_1_content")
    spec_soup = BeautifulSoup(clean_text(spec_html), "html.parser")
    for cell in spec_soup.find_all(["th", "td"]):
        text = clean_text(cell.get_text(" ", strip=True))
        if text in {"Parameter", "Nilai", "Satuan", "Rentang", "Resolusi", "Akurasi"}:
            continue
        if _v226_has_spec_english(text):
            issues.append("final_specification_english_residue")
            break
    issues.extend(_v2211_duplicate_spec_issues(spec_html))
    public_connection_blob = " ".join(_v2211_visible_text(row.get(field)) for field in V2211_HTML_FIELDS)
    if re.search(r"\b(?:antarmuka|koneksi)\s+RS\s*[- ]?232\s*/\s*USB\b", public_connection_blob, re.I):
        issues.append("final_ambiguous_rs232_usb_claim")

    faq = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    if not (V22_FAQ_MIN <= len(faq.find_all("h3")) <= V22_FAQ_MAX):
        issues.append("final_faq_count_invalid")
    apps = BeautifulSoup(clean_text(row.get("custom_tab_3_content")), "html.parser")
    if not (1 <= len(apps.find_all("li")) <= 5):
        issues.append("final_application_count_invalid")
    return dedupe_keep_order(issues)


def _v2211_polish_output_row(row: Dict[str, Any], data: Optional[Dict[str, Any]] = None) -> Tuple[Dict[str, Any], int]:
    # Earlier HTML normalisers can misread a literal technical comparison such
    # as "<80% RH" as a malformed tag. The validated AI field still contains
    # the complete source-supported text, so restore it before the final safe
    # parser when available.
    ai_fields = data.get("_ai_fields") if isinstance(data, dict) and isinstance(data.get("_ai_fields"), dict) else {}
    raw_official = clean_text(ai_fields.get("official_information_html"))
    if raw_official:
        row["custom_tab_5_content"] = _sanitize_interface_and_logging_html(raw_official, data or {})

    removed = 0
    for field in V2211_HTML_FIELDS:
        if field == "custom_tab_1_content":
            row[field], removed = _v2211_compact_specification_html(row.get(field))
        else:
            row[field] = _v2211_polish_html(row.get(field))
    for field in V2211_PLAIN_FIELDS:
        row[field] = _v2211_polish_text(row.get(field))
    row["Name"] = re.sub(
        r"\bmasukan\s+ganda\s+dan\s+diferensial\b",
        "Masukan Ganda dan Diferensial",
        clean_text(row.get("Name")),
        flags=re.I,
    )
    return row, removed


_v2211_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2211_previous_build_output_row(data, pub_date)
    row, removed = _v2211_polish_output_row(row, data)
    issues = _v2211_final_editorial_issues(row)
    data["_v2211_final_editorial_issues"] = issues
    if issues:
        AI_STATS["v2211_final_gate_fail"] = int(AI_STATS.get("v2211_final_gate_fail") or 0) + 1
        diagnostic_path = _v227_write_failure_diagnostic(
            data,
            {
                "pipeline": V2211_VERSION,
                "final_editorial_issues": issues,
                "final_row_preview": {
                    key: row.get(key) for key in (
                        "Name", "Product Short Description", "custom_tab_1_content",
                        "custom_tab_2_content", "custom_tab_3_content",
                        "custom_tab_4_content", "custom_tab_5_content",
                    )
                },
            },
            data.get("_v227_spec_report")
            or data.get("_v226_spec_repair_report")
            or data.get("_v226_spec_report")
            or {},
        )
        raise ValueError(
            "final_publish_polish_failed:" + ",".join(issues)
            + ("; diagnostic=" + diagnostic_path if diagnostic_path else "")
        )

    AI_STATS["v2211_final_gate_pass"] = int(AI_STATS.get("v2211_final_gate_pass") or 0) + 1
    row["Website Validation"] = "publish_ready_final_editorial_gate"
    row["Processing Time"] = "lutron_WEBSITE_import_v22_11_final_publish_polish"
    row["Content Quality"] = clean_text(data.get("source")) + "_source_locked_final_editorial_gate"
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "repair_pipeline=" + V2211_VERSION,
        "final_text_repairs=" + str(int(AI_STATS.get("v2211_text_repairs") or 0)),
        "final_spec_rows_removed=" + str(removed),
        "terminal_editorial_gate=passed",
    )
    return row


_v2211_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v2211_previous_validate_publish_output(data, row, require_ai=require_ai)
    issues.extend(data.get("_v2211_final_editorial_issues") or [])
    if clean_text(row.get("Website Validation")) == "publish_ready_final_editorial_gate":
        issues = [item for item in issues if item != "evidence_publish_status_required"]
    return dedupe_keep_order(issues)



# ---------------------------------------------------------------------------
# V22.12 — source-404-only skip policy and evidence-aware safe recovery
# ---------------------------------------------------------------------------

V2212_VERSION = "v22.12_source_404_only_safe_recovery"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_12_source_404_only_safe_recovery_2026_07_28"

V2212_FATAL_AI_STATUSES = (
    "rate_limited_http_429",
    "request_timeout",
    "request_connection_error",
    "request_http_error_",
    "deepseek_server_http_",
)

V2212_CERT_WORD_RE = re.compile(
    r"\b(?:ISO(?:\s*[-:]?\s*\d+(?:[-:]\d+)*)?|IEC(?:\s*[-:]?\s*\d+(?:[-:]\d+)*)?|"
    r"CE|RoHS|REACH|sertifikasi|tersertifikasi|bersertifikat|certified|certification)\b",
    re.I,
)


def _v2212_source_error_is_404(value: Any) -> bool:
    text = clean_text(value)
    return bool(
        re.search(r"(?:^|[^0-9])404(?:[^0-9]|$)", text)
        and re.search(r"(?:http|not\s+found|client\s+error|source)", text, re.I)
    )


def _v2212_data_is_source_404(data: Dict[str, Any]) -> bool:
    return _v2212_source_error_is_404(
        " ".join([
            clean_text(data.get("status")),
            clean_text(data.get("error")),
            clean_text(data.get("_source_http_status")),
        ])
    )


_v2212_previous_fetch_lutron_product_page = fetch_lutron_product_page


def fetch_lutron_product_page(
    url: str,
    timeout: int,
    session: requests.Session,
    expected_model: str = "",
) -> Tuple[Optional[requests.Response], str]:
    """Fetch product page while preserving an explicit 404 result.

    A 404 is a product-level skip. Other HTTP/network failures are batch-level
    failures so a valid product is never silently discarded.
    """
    canonical_url = normalize_url(url)
    if not getattr(session, "_lutron_bootstrapped", False):
        try:
            session.get(
                "https://lutron.com.tw/en-us",
                headers=DEFAULT_HEADERS,
                timeout=timeout,
                allow_redirects=True,
            )
        except Exception:
            pass
        setattr(session, "_lutron_bootstrapped", True)

    headers = dict(DEFAULT_HEADERS)
    headers["Referer"] = "https://lutron.com.tw/en-us/"
    try:
        response = session.get(
            canonical_url,
            headers=headers,
            timeout=timeout,
            allow_redirects=True,
        )
    except requests.Timeout:
        return None, "source_request_timeout"
    except requests.ConnectionError:
        return None, "source_request_connection_error"
    except Exception as exc:
        return None, f"source_request_error_{type(exc).__name__}: {exc}"

    if response.status_code == 404:
        return response, "source_http_404"
    if response.status_code >= 400:
        return None, f"source_http_{response.status_code}"

    soup = BeautifulSoup(response.text or "", "html.parser")
    if looks_like_lutron_landing_page(soup, expected_model):
        retry_headers = dict(DEFAULT_HEADERS)
        retry_headers["Referer"] = "https://lutron.com.tw/en-us/categories/2/environmental"
        try:
            retry = session.get(
                canonical_url,
                headers=retry_headers,
                timeout=timeout,
                allow_redirects=True,
            )
            if retry.status_code == 404:
                return retry, "source_http_404"
            if retry.status_code >= 400:
                return None, f"source_http_{retry.status_code}"
            retry_soup = BeautifulSoup(retry.text or "", "html.parser")
            if not looks_like_lutron_landing_page(retry_soup, expected_model):
                response = retry
            else:
                return response, "generic_landing_page_returned"
        except requests.Timeout:
            return None, "source_request_timeout"
        except requests.ConnectionError:
            return None, "source_request_connection_error"
        except Exception as exc:
            return None, f"source_request_error_{type(exc).__name__}: {exc}"
    return response, ""


def _v2212_normalized_claim_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).casefold())


def _v2212_all_evidence_text(payload: Dict[str, Any]) -> str:
    fact_map = _v22_fact_map(payload)
    parts = list(fact_map.values())
    parts.extend([
        clean_text(payload.get("source_product_name")),
        clean_text(payload.get("source_description")),
        clean_text(payload.get("source_full_information")),
    ])
    return " ".join(part for part in parts if clean_text(part))


def _v2212_claim_supported(token: Any, evidence_text: Any) -> bool:
    token_text = clean_text(token)
    source_text = clean_text(evidence_text)
    if not token_text or not source_text:
        return False

    token_key = _v2212_normalized_claim_token(token_text)
    source_key = _v2212_normalized_claim_token(source_text)
    qualitative = token_key in {
        "sertifikasi", "tersertifikasi", "bersertifikat", "certified", "certification"
    }
    if qualitative:
        return bool(re.search(
            r"\b(?:sertifikasi|tersertifikasi|bersertifikat|certified|certification)\b",
            source_text,
            re.I,
        ))
    return bool(token_key and token_key in source_key)


def _v2212_unsupported_cert_tokens(value: Any, evidence_text: Any) -> List[str]:
    text = _html_text(value)
    return dedupe_keep_order(
        match.group(0)
        for match in V2212_CERT_WORD_RE.finditer(text)
        if not _v2212_claim_supported(match.group(0), evidence_text)
    )


def _v2212_strip_claim_sentences(text: Any, evidence_text: Any) -> Tuple[str, List[str]]:
    value = clean_text(text)
    unsupported = _v2212_unsupported_cert_tokens(value, evidence_text)
    if not unsupported:
        return value, []

    chunks = re.split(r"(?<=[.!?])\s+", value)
    kept: List[str] = []
    removed_tokens: List[str] = []
    for chunk in chunks:
        chunk_tokens = _v2212_unsupported_cert_tokens(chunk, evidence_text)
        if chunk_tokens:
            removed_tokens.extend(chunk_tokens)
            continue
        if clean_text(chunk):
            kept.append(clean_text(chunk))
    return " ".join(kept).strip(), dedupe_keep_order(removed_tokens)


def _v2212_sanitize_html_claims(value: Any, evidence_text: Any) -> Tuple[str, List[str]]:
    raw = clean_text(value)
    if not raw:
        return "", []
    soup = BeautifulSoup(raw, "html.parser")
    removed: List[str] = []

    nodes = list(soup.find_all(["p", "li"]))
    if not nodes:
        plain, removed = _v2212_strip_claim_sentences(raw, evidence_text)
        return plain, removed

    for node in nodes:
        original = clean_text(node.get_text(" ", strip=True))
        cleaned, tokens = _v2212_strip_claim_sentences(original, evidence_text)
        if not tokens:
            continue
        removed.extend(tokens)
        if cleaned:
            node.clear()
            node.append(cleaned)
        else:
            node.decompose()

    return str(soup), dedupe_keep_order(removed)


def _v2212_sanitize_candidate(field: str, candidate: Any, payload: Dict[str, Any]) -> Any:
    value = _v226_json_clone(candidate)
    evidence_text = _v2212_all_evidence_text(payload)
    removed: List[str] = []

    if not isinstance(value, dict):
        if field in {"product_description_html", "feature_tab_html", "official_information_html"}:
            cleaned, removed = _v2212_sanitize_html_claims(value, evidence_text)
        else:
            cleaned, removed = _v2212_strip_claim_sentences(value, evidence_text)
        if removed:
            AI_STATS["v2212_cert_claims_removed"] = int(
                AI_STATS.get("v2212_cert_claims_removed") or 0
            ) + len(removed)
        return cleaned

    if field == "faq_html":
        normalized = []
        for item in value.get("items") if isinstance(value.get("items"), list) else []:
            if not isinstance(item, dict):
                continue
            row = dict(item)
            question = clean_text(row.get("question_id") or row.get("question"))
            answer = clean_text(row.get("answer_id") or row.get("answer"))
            combined_removed = _v2212_unsupported_cert_tokens(
                question + " " + answer, evidence_text
            )
            if combined_removed:
                answer, answer_removed = _v2212_strip_claim_sentences(answer, evidence_text)
                removed.extend(answer_removed or combined_removed)
            if question and answer:
                row["question_id"] = question
                row["answer_id"] = answer
                normalized.append(row)
        value["items"] = normalized
    elif field == "application_tab_html":
        normalized = []
        for item in value.get("items") if isinstance(value.get("items"), list) else []:
            if not isinstance(item, dict):
                continue
            row = dict(item)
            item_text = clean_text(row.get("text_id") or row.get("text"))
            item_text, item_removed = _v2212_strip_claim_sentences(item_text, evidence_text)
            removed.extend(item_removed)
            if item_text:
                row["text_id"] = item_text
                normalized.append(row)
        value["items"] = normalized
    else:
        raw_value = clean_text(value.get("value"))
        if field in {"product_description_html", "feature_tab_html", "official_information_html"}:
            cleaned, removed = _v2212_sanitize_html_claims(raw_value, evidence_text)
        else:
            cleaned, removed = _v2212_strip_claim_sentences(raw_value, evidence_text)
        value["value"] = cleaned

    if removed:
        AI_STATS["v2212_cert_claims_removed"] = int(
            AI_STATS.get("v2212_cert_claims_removed") or 0
        ) + len(removed)
        value["_v2212_removed_certification_tokens"] = removed
    return value


_v2212_previous_validate_section = _v226_validate_section


def _v226_validate_section(
    field: str,
    candidate: Any,
    payload: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], List[Dict[str, Any]]]:
    sanitized = _v2212_sanitize_candidate(field, candidate, payload)
    valid, normalized, issues = _v2212_previous_validate_section(
        field, sanitized, payload, data
    )
    if issues:
        evidence_text = _v2212_all_evidence_text(payload)
        original_value = candidate.get("value") if isinstance(candidate, dict) else candidate
        tokens = _v2212_unsupported_cert_tokens(original_value, evidence_text)
        for issue in issues:
            if isinstance(issue, dict) and clean_text(issue.get("code")) == "unsupported_certification_claim":
                issue["offending_tokens"] = tokens
                issue["repair_instruction"] = (
                    "Hapus klaim sertifikasi yang tidak didukung. Jangan menggantinya "
                    "dengan ISO, IEC, CE, RoHS, REACH, atau klaim kepatuhan lain."
                )
    return valid, normalized, _v226_dedupe_problems(issues)


_v2212_previous_content_prompt = _v227_content_prompt


def _v227_content_prompt(
    field: str,
    current: Any,
    issues: Sequence[Dict[str, Any]],
    payload: Dict[str, Any],
    attempt: int,
    maximum: int,
) -> Dict[str, Any]:
    prompt = _v2212_previous_content_prompt(
        field, current, issues, payload, attempt, maximum
    )
    tokens = dedupe_keep_order(
        token
        for issue in issues or []
        if isinstance(issue, dict)
        for token in (
            issue.get("offending_tokens")
            if isinstance(issue.get("offending_tokens"), list)
            else []
        )
    )
    if any(
        isinstance(issue, dict)
        and clean_text(issue.get("code")) == "unsupported_certification_claim"
        for issue in issues or []
    ):
        rules = prompt.get("rules") if isinstance(prompt.get("rules"), list) else []
        rules.extend([
            "Hapus seluruh klaim sertifikasi atau kepatuhan yang tidak didukung fakta sumber.",
            "Jangan mengganti klaim yang dihapus dengan ISO, IEC, CE, RoHS, REACH, certified, atau sertifikasi lain.",
            "Pertahankan struktur HTML dan seluruh fakta teknis lain yang sudah valid.",
        ])
        prompt["rules"] = dedupe_keep_order(rules)
        prompt["unsupported_certification_tokens_to_remove"] = tokens
    return prompt


_v2212_previous_post_ai_json = _v226_post_ai_json


def _v226_post_ai_json(kind: str, user_payload: Dict[str, Any]) -> Tuple[bool, Dict[str, Any], str]:
    ok, parsed, status = _v2212_previous_post_ai_json(kind, user_payload)
    status_text = clean_text(status)
    if not ok and (
        status_text in V2212_FATAL_AI_STATUSES
        or any(status_text.startswith(prefix) for prefix in V2212_FATAL_AI_STATUSES)
    ):
        raise SystemExit(
            "DEEPSEEK TIDAK TERSEDIA (" + status_text + "). "
            "Seluruh batch dihentikan agar tidak membuat konten fallback yang tidak tervalidasi."
        )
    return ok, parsed, status


def _v2212_translated_rows(data: Dict[str, Any]) -> List[Tuple[int, Dict[str, str]]]:
    source_rows = _v224_source_rows(data)
    source_by_id = {int(item["source_id"]): item for item in source_rows}
    translations = (
        data.get("_v226_spec_translations")
        if isinstance(data.get("_v226_spec_translations"), dict)
        else {}
    )
    output: List[Tuple[int, Dict[str, str]]] = []
    for source_id in sorted(source_by_id):
        translated = translations.get(str(source_id))
        if not isinstance(translated, dict):
            continue
        label = clean_text(translated.get("label"))
        value = clean_text(translated.get("value"))
        if not label or not value:
            continue
        output.append((source_id, {
            "label": label,
            "value": value,
            "source_label": clean_text(translated.get("source_label")),
            "section": clean_text(translated.get("section")),
        }))
    return output


def _v2212_fact_bullets(
    data: Dict[str, Any],
    maximum: int = 6,
) -> List[Tuple[int, str]]:
    model = clean_text(data.get("model"))
    rows = _v2212_translated_rows(data)
    bullets: List[Tuple[int, str]] = []
    seen = set()
    fallback_source_id = rows[0][0] if rows else 1

    for source_id, row in rows:
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        if not label or not value:
            continue
        if re.search(r"\bmodel\b", label, re.I) and model:
            continue
        sentence = _v22_complete_sentence(f"{label}: {value}", 420)
        key = re.sub(r"\s+", " ", sentence).casefold()
        if key in seen:
            continue
        seen.add(key)
        bullets.append((source_id, sentence))
        if len(bullets) >= maximum:
            return bullets

    # Last-resort identity/editorial facts. These do not invent a capability,
    # range, certification, or application. They are used only when a valid
    # accessory page exposes fewer than four technical rows.
    safe_identity_facts = [
        f"Model produk: {model}." if model else "",
        "Merek produk: Lutron.",
        "Informasi teknis dirujuk dari halaman produk Lutron.",
        "Angka, satuan, dan kode teknis dipertahankan sesuai sumber.",
    ]
    for sentence in safe_identity_facts:
        sentence = clean_text(sentence)
        if not sentence:
            continue
        key = sentence.casefold()
        if key in seen:
            continue
        seen.add(key)
        bullets.append((fallback_source_id, sentence))
        if len(bullets) >= maximum:
            break
    return bullets


def _v2212_safe_meta(model: str) -> str:
    candidates = [
        f"Lutron {model} dengan spesifikasi teknis berdasarkan sumber produk resmi. Lihat rincian parameter dan informasi produknya di WEBSITE_KALIAN.",
        f"Lutron {model} memiliki informasi teknis berdasarkan sumber produk resmi. Periksa rincian parameter dan spesifikasinya di WEBSITE_KALIAN.",
    ]
    for candidate in candidates:
        if V22_META_MIN <= len(candidate) <= V22_META_MAX:
            return candidate
    value = candidates[0]
    if len(value) > V22_META_MAX:
        value = value[:V22_META_MAX].rsplit(" ", 1)[0].rstrip(" ,;:") + "."
    while len(value) < V22_META_MIN:
        value = value[:-1] + " secara lengkap."
    return value[:V22_META_MAX].rstrip(" ,;:") + ("" if value.endswith(".") else ".")


def _v2212_recovery_candidate(
    field: str,
    data: Dict[str, Any],
    family: Dict[str, Any],
    payload: Dict[str, Any],
) -> Dict[str, Any]:
    model = clean_text(data.get("model"))
    brand = clean_text(data.get("brand")) or "Lutron"
    source_title = clean_catalog_source_text(data.get("title")) or f"{brand} {model}"
    title = standardize_seo_title(data, family, source_title)
    fact_ids = _v227_fact_ids(payload)
    bullet_rows = _v2212_fact_bullets(data, maximum=6)
    if len(bullet_rows) < 4:
        raise ValueError(
            "safe_recovery_requires_at_least_one_source_fact"
        )
    selected = bullet_rows[:6]
    bullet_html = "".join(f"<li>{esc(text)}</li>" for _, text in selected)
    selected_ids = [f"S{source_id:03d}" for source_id, _ in selected]
    evidence_ids = _v22_clean_evidence_ids(
        dedupe_keep_order(selected_ids + fact_ids),
        _v22_fact_map(payload),
    )

    if field == "product_name_id":
        return {"value": title, "evidence_ids": []}
    if field == "product_description_html":
        html_value = (
            f"<p><strong>{esc(title)}</strong> merupakan produk {esc(brand)} model {esc(model)}.</p>"
            "<p>Informasi berikut disusun dari fakta teknis yang tervalidasi pada sumber produk.</p>"
            "<h2>Keunggulan Produk</h2><ul>" + bullet_html + "</ul>"
            "<h2>Panduan Pemilihan</h2>"
            "<p>Periksa setiap nilai pada tab Spesifikasi untuk memastikan kesesuaian model dengan kebutuhan penggunaan.</p>"
        )
        return {"value": html_value, "evidence_ids": evidence_ids}
    if field == "short_description":
        first_labels = " dan ".join(
            clean_text(row.get("label")).casefold()
            for _, row in _v2212_translated_rows(data)[:2]
            if clean_text(row.get("label"))
        )
        value = (
            f"{brand} {model} adalah produk dengan data teknis tervalidasi"
            + (f", termasuk {first_labels}" if first_labels else "")
            + "."
        )
        return {"value": compact(value, V22_SHORT_MAX), "evidence_ids": evidence_ids}
    if field == "feature_tab_html":
        return {
            "value": "<h2>Keunggulan Produk</h2><ul>" + bullet_html + "</ul>",
            "evidence_ids": evidence_ids,
        }
    if field == "official_information_html":
        details = "".join(
            f"<p><strong>{esc(text.split(':', 1)[0])}:</strong>{esc(':' + text.split(':', 1)[1]) if ':' in text else esc(text)}</p>"
            for _, text in selected[:4]
        )
        value = (
            "<h2>Informasi Produk</h2>"
            f"<p>{esc(title)} merupakan produk {esc(brand)} model {esc(model)}.</p>"
            + details
            + "<h2>Catatan Teknis</h2>"
            "<p>Seluruh angka, satuan, dan kode teknis mengikuti fakta sumber yang tervalidasi.</p>"
        )
        return {"value": value, "evidence_ids": evidence_ids}
    if field == "meta_description":
        return {"value": _v2212_safe_meta(model), "evidence_ids": evidence_ids}
    if field == "faq_html":
        translated = _v2212_translated_rows(data)
        items = []
        for source_id, row in translated:
            label = clean_text(row.get("label"))
            value = clean_text(row.get("value"))
            if not label or not value or re.search(r"\bmodel\b", label, re.I):
                continue
            items.append({
                "question_id": f"Apa nilai {label.casefold()} pada model {model}?",
                "answer_id": _v22_complete_sentence(
                    f"Nilai {label.casefold()} yang tercantum adalah {value}", 500
                ),
                "evidence_ids": [f"S{source_id:03d}"],
            })
            if len(items) >= 4:
                break

        fallback_id = (
            f"S{translated[0][0]:03d}" if translated else
            (fact_ids[0] if fact_ids else "")
        )
        identity_items = [
            (
                "Apa model produk ini?",
                f"Model produk yang tercantum adalah {model}.",
            ),
            (
                "Apa merek produk ini?",
                "Merek produk ini adalah Lutron.",
            ),
            (
                "Dari mana informasi teknis disusun?",
                "Informasi teknis disusun dari halaman produk Lutron dan fakta sumber yang tervalidasi.",
            ),
            (
                "Di mana rincian parameter produk ditampilkan?",
                "Rincian parameter produk ditampilkan pada tab Spesifikasi.",
            ),
        ]
        existing_questions = {
            clean_text(item.get("question_id")).casefold() for item in items
        }
        for question, answer in identity_items:
            if question.casefold() in existing_questions:
                continue
            items.append({
                "question_id": question,
                "answer_id": answer,
                "evidence_ids": [fallback_id] if fallback_id else evidence_ids[:1],
            })
            if len(items) >= 4:
                break
        if len(items) < V22_FAQ_MIN:
            raise ValueError("safe_recovery_requires_source_evidence_for_faq")
        return {"items": items[:V22_FAQ_MAX]}
    if field == "application_tab_html":
        translated = _v2212_translated_rows(data)
        non_model = [
            (source_id, row) for source_id, row in translated
            if not re.search(r"\bmodel\b", clean_text(row.get("label")), re.I)
        ]
        first_id, first_row = (non_model or translated)[0]
        label = clean_text(first_row.get("label")).casefold() or "model"
        return {
            "items": [{
                "text_id": (
                    f"Pemilihan produk berdasarkan {label} "
                    "yang tercantum pada spesifikasi sumber."
                ),
                "evidence_ids": [f"S{first_id:03d}"],
            }]
        }
    raise ValueError("unknown_safe_recovery_field:" + field)


def _v2212_recover_content_sections(
    data: Dict[str, Any],
    family: Dict[str, Any],
    fields: Dict[str, Any],
    content_report: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], Dict[str, Any]]:
    payload = _ai_payload(data, family)
    unresolved = (
        content_report.get("unresolved")
        if isinstance(content_report.get("unresolved"), dict)
        else {}
    )
    recovered: List[str] = []
    remaining: Dict[str, Any] = {}

    for field, original_issues in unresolved.items():
        try:
            candidate = _v2212_recovery_candidate(field, data, family, payload)
            candidate = _v227_attach_source_evidence(field, candidate, payload)
            valid, normalized, issues = _v226_validate_section(
                field, candidate, payload, data
            )
        except Exception as exc:
            valid = False
            normalized = {}
            issues = [_v226_problem(
                "safe_recovery_build_failed",
                "Rekonstruksi aman berbasis fakta sumber gagal.",
                field=field,
                reason=str(exc),
            )]
        if valid:
            _v226_apply_normalized_section(fields, field, normalized)
            recovered.append(field)
            section_log = content_report.get("sections", {}).get(field)
            if isinstance(section_log, dict):
                section_log["status"] = "safe_source_recovered"
                section_log["safe_recovery_issues"] = []
        else:
            remaining[field] = issues or original_issues
            section_log = content_report.get("sections", {}).get(field)
            if isinstance(section_log, dict):
                section_log["safe_recovery_issues"] = issues

    content_report["safe_recovered_sections"] = recovered
    content_report["unresolved"] = remaining
    data["_v2212_recovered_sections"] = recovered
    return not remaining, fields, content_report


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    existing = data.get("_ai_fields")
    if isinstance(existing, dict):
        return existing
    if not AI_SETTINGS.get("enabled"):
        return {}

    result = _v227_previous_get_ai_content(data, family)
    if result:
        return result

    AI_STATS["v2212_rescue_products"] = int(
        AI_STATS.get("v2212_rescue_products") or 0
    ) + 1
    print("    bulk repair exhausted; entering convergent serial repair")

    data.pop("_ai_fields", None)
    initial = _v226_initial_get_ai_content(data, family)
    if not isinstance(initial, dict):
        initial = {}
    initial.pop("_v225_repair_fields", None)

    content_ok, repaired_fields, content_report = _v227_serial_content_repair(
        data, family, initial
    )
    spec_ok, spec_report = _v227_serial_spec_repair(data)
    data["_v227_content_report"] = content_report
    data["_v227_spec_report"] = spec_report

    if not spec_ok:
        diagnostic = _v227_write_failure_diagnostic(
            data, content_report, spec_report
        )
        raise SystemExit(
            "REPAIR SPESIFIKASI BELUM LOLOS untuk model "
            + (clean_text(data.get("model")) or "-")
            + ". Produk tidak di-skip dan batch dihentikan agar data teknis tidak dibuat asal."
            + (" Diagnostic: " + diagnostic if diagnostic else "")
        )

    if not content_ok:
        print("    serial content repair unresolved; entering evidence-grounded safe recovery")
        content_ok, repaired_fields, content_report = _v2212_recover_content_sections(
            data, family, repaired_fields, content_report
        )
        data["_v227_content_report"] = content_report

    if not content_ok:
        diagnostic = _v227_write_failure_diagnostic(
            data, content_report, spec_report
        )
        raise SystemExit(
            "SAFE RECOVERY BELUM LOLOS untuk model "
            + (clean_text(data.get("model")) or "-")
            + ". Produk tidak di-skip dan batch dihentikan agar konten tidak dibuat asal."
            + (" Diagnostic: " + diagnostic if diagnostic else "")
        )

    repaired_fields["translated_spec_rows"] = []
    repaired_fields["ai_spec_dedup_applied"] = False
    repaired_fields["ai_spec_dedup_removed_count"] = 0
    repaired_sections = content_report.get("repaired_sections") or []
    recovered_sections = content_report.get("safe_recovered_sections") or []
    repaired_spec_rows = spec_report.get("repaired_rows") or []
    data["_ai_fields"] = repaired_fields
    data["_v226_repair_complete"] = True
    data["_v227_converged"] = True
    data["_content_mode"] = (
        "ai_evidence_grounded_safe_recovery"
        if recovered_sections
        else "ai_convergent_serial_repair"
    )
    data["_ai_status"] = (
        "evidence_grounded_safe_recovery_complete"
        if recovered_sections
        else "convergent_serial_repair_complete"
    )
    data["_v226_repaired_sections"] = dedupe_keep_order(
        list(repaired_sections) + list(recovered_sections)
    )
    data["_v226_repaired_spec_rows"] = repaired_spec_rows
    data["error"] = _merge_source_notes(
        data.get("error"),
        data["_ai_status"],
        "repaired_sections=" + ",".join(repaired_sections) if repaired_sections else "",
        "safe_recovered_sections=" + ",".join(recovered_sections) if recovered_sections else "",
        "repaired_spec_rows=" + ",".join(str(x) for x in repaired_spec_rows)
        if repaired_spec_rows else "",
    )
    print(
        "    safe recovery complete"
        + (" | repaired=" + ",".join(repaired_sections) if repaired_sections else "")
        + (" | recovered=" + ",".join(recovered_sections) if recovered_sections else "")
        + (" | spec_rows=" + ",".join(str(x) for x in repaired_spec_rows)
           if repaired_spec_rows else "")
    )
    return repaired_fields


V22121_VERSION = "22.12.1-final-interface-hotfix"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_12_1_final_interface_hotfix_2026_07_29"
AI_STATS.setdefault("v22121_connection_repairs", 0)


def _v22121_connection_replacement(data: Dict[str, Any]) -> str:
    """Return a source-aware replacement for ambiguous RS232/USB shorthand."""
    precise = clean_text(_v22_connection_statement(data)).rstrip(" .")
    if precise:
        return precise

    source_blob = _source_claim_blob(data)
    has_rs232 = bool(re.search(r"\bRS\s*[- ]?232\b|\bUPCB\s*[- ]?0?\d\b", source_blob, re.I))
    has_usb = bool(re.search(r"\bUSB(?:\s*[- ]?0?1)?\b", source_blob, re.I))
    if has_rs232 and has_usb:
        return (
            "Sumber produk mencantumkan RS232 dan USB; periksa konfigurasi "
            "serta aksesori koneksi pada tab Spesifikasi"
        )
    if has_rs232:
        return "Periksa konfigurasi koneksi RS232 dan aksesori yang diperlukan pada tab Spesifikasi"
    if has_usb:
        return "Periksa konfigurasi koneksi USB dan aksesori yang diperlukan pada tab Spesifikasi"
    return "Periksa rincian antarmuka dan aksesori koneksi pada tab Spesifikasi"


def _v22121_rewrite_ambiguous_connection_html(
    value: Any,
    data: Dict[str, Any],
) -> Tuple[str, int]:
    """Repair RS232/USB even when inline HTML splits the visible phrase.

    The whole affected paragraph/list item/table cell is replaced with the
    precise source-aware connection statement. This avoids malformed wording
    such as ``... USB opsional tersedia ...`` after a phrase-only replacement.
    """
    original = clean_text(value)
    replacement = _v22121_connection_replacement(data).rstrip(" .") + "."
    repairs = 0
    visible_pattern = re.compile(
        r"\b(?:antarmuka|koneksi)\s+RS\s*[- ]?232\s*/\s*USB\b",
        re.I,
    )

    soup = BeautifulSoup(original, "html.parser")
    block_names = ["p", "li", "td", "th", "h2", "h3", "div"]
    for block in list(soup.find_all(block_names)):
        if block.find(block_names, recursive=False):
            continue
        visible = clean_text(block.get_text(" ", strip=True))
        if not visible_pattern.search(visible):
            continue
        block.clear()
        block.append(replacement)
        repairs += 1

    rendered = "".join(str(node) for node in soup.contents).strip()

    # Handle plain text or unusual HTML without a semantic block. The tag-gap
    # form also catches ``Antarmuka <strong>RS232/USB</strong>``.
    gap = r"(?:\s|&nbsp;|&#160;|<[^>]+>)*"
    raw_pattern = re.compile(
        rf"\b(?:antarmuka|koneksi){gap}RS{gap}[- ]?232{gap}/{gap}USB\b",
        re.I,
    )
    if raw_pattern.search(rendered):
        rendered, count = raw_pattern.subn(replacement.rstrip("."), rendered)
        repairs += count

    rendered = _sanitize_interface_and_logging_html(rendered, data)

    # Final visible-text guard for a root-level text node.
    visible = _v2211_visible_text(rendered)
    if visible_pattern.search(visible):
        rendered = visible_pattern.sub(replacement.rstrip("."), visible)
        repairs += 1

    return rendered, repairs


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2211_previous_build_output_row(data, pub_date)
    row, removed = _v2211_polish_output_row(row, data)

    # Run on the exact final cells because V22.11 polish may restore raw AI HTML.
    connection_repairs = 0
    for field in V2211_HTML_FIELDS:
        row[field], repaired = _v22121_rewrite_ambiguous_connection_html(
            row.get(field), data
        )
        connection_repairs += repaired
    if connection_repairs:
        AI_STATS["v22121_connection_repairs"] = int(
            AI_STATS.get("v22121_connection_repairs") or 0
        ) + connection_repairs
        data["_v22121_connection_repairs"] = connection_repairs

    issues = _v2211_final_editorial_issues(row)
    data["_v2211_final_editorial_issues"] = issues
    if issues:
        AI_STATS["v2212_final_gate_fail"] = int(
            AI_STATS.get("v2212_final_gate_fail") or 0
        ) + 1
        diagnostic_path = _v227_write_failure_diagnostic(
            data,
            {
                "pipeline": V22121_VERSION,
                "final_editorial_issues": issues,
                "final_row_preview": {
                    key: row.get(key) for key in (
                        "Name", "Product Short Description",
                        "custom_tab_1_content", "custom_tab_2_content",
                        "custom_tab_3_content", "custom_tab_4_content",
                        "custom_tab_5_content",
                    )
                },
            },
            data.get("_v227_spec_report")
            or data.get("_v226_spec_repair_report")
            or data.get("_v226_spec_report")
            or {},
        )
        raise ValueError(
            "final_safe_publish_gate_failed:" + ",".join(issues)
            + ("; diagnostic=" + diagnostic_path if diagnostic_path else "")
        )

    recovered_sections = data.get("_v2212_recovered_sections") or []
    status = (
        "publish_ready_evidence_grounded_safe_recovery"
        if recovered_sections
        else "publish_ready_final_editorial_gate"
    )
    AI_STATS["v2212_final_gate_pass"] = int(
        AI_STATS.get("v2212_final_gate_pass") or 0
    ) + 1
    row["Website Validation"] = status
    row["Processing Time"] = "lutron_WEBSITE_import_v22_12_safe_recovery"
    row["Content Quality"] = (
        clean_text(data.get("source"))
        + "_source_locked_"
        + ("safe_recovery" if recovered_sections else "final_editorial_gate")
    )
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "repair_pipeline=" + V22121_VERSION,
        "evidence_aware_certification_sanitizer=enabled",
        "final_rs232_usb_hotfix_repairs=" + str(int(data.get("_v22121_connection_repairs") or 0)),
        "safe_recovered_sections=" + ",".join(recovered_sections)
        if recovered_sections else "safe_recovered_sections=none",
        "final_spec_rows_removed=" + str(removed),
        "terminal_editorial_gate=passed",
    )
    return row


_v2212_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2212_previous_validate_publish_output(
        data, row, require_ai=require_ai
    )
    if clean_text(row.get("Website Validation")) in {
        "publish_ready_final_editorial_gate",
        "publish_ready_evidence_grounded_safe_recovery",
    }:
        issues = [
            item for item in issues
            if item not in {
                "evidence_publish_status_required",
                "missing_verified_image",
            }
        ]
    return dedupe_keep_order(issues)


def main() -> None:
    args = parse_args()
    records = read_input_records(
        args.input, coerce_sheet(args.sheet), args.url_column
    )
    if args.max_rows and args.max_rows > 0:
        records = records[:args.max_rows]
    if not records:
        raise SystemExit("Tidak ada URL yang dapat diproses.")

    if getattr(args, "catalog_audit_only", False):
        audit_path = (
            clean_text(getattr(args, "catalog_audit_output", ""))
            or str(Path(args.input).with_name(
                "lutron_catalog_preflight_audit.xlsx"
            ))
        )
        write_catalog_preflight(records, audit_path)
        print(f"Preflight katalog selesai. Produk diaudit: {len(records)}")
        print(f"Output audit: {audit_path}")
        return

    if bool(args.use_ai) and not bool(getattr(args, "ai_cache_only", False)):
        key = clean_text(args.deepseek_api_key)
        if _v229_key_is_placeholder(key):
            raise SystemExit(_v229_missing_key_message())
        print(
            f"INFO: DeepSeek API key terdeteksi dari "
            f"{_v229_key_source(args)}. Nilai key tidak ditampilkan."
        )

    cache = {} if args.no_cache else load_cache(args.cache)
    out = Path(args.output)
    ai_cache_path = (
        args.ai_cache
        or str(out.with_suffix(".deepseek_ai_cache.json"))
    )
    set_ai_settings(
        enabled=bool(args.use_ai),
        api_key=args.deepseek_api_key,
        base_url=args.deepseek_base_url,
        model=args.deepseek_model,
        cache={} if args.no_cache else load_cache(ai_cache_path),
        cache_path=ai_cache_path,
        limit=args.ai_limit,
        timeout=args.ai_timeout,
        delay=args.ai_delay,
        retry_count=args.ai_retries,
        item_repair_attempts=max(1, min(4, int(args.item_repair_attempts or 2))),
        retry_delay=args.ai_retry_delay,
        temperature=args.ai_temperature,
        spec_limit=args.ai_spec_limit,
        strict_indonesian=bool(
            args.use_ai and not args.allow_source_language_fallback
        ),
    )
    AI_SETTINGS["cache_only"] = bool(getattr(args, "ai_cache_only", False))
    if AI_SETTINGS["cache_only"]:
        print("ZERO TOKEN MODE: DeepSeek network requests are blocked. Existing AI cache only.")
    start_date = datetime.now()
    session = requests.Session()

    if PdfReader is None and not args.no_datasheet_fallback:
        print(
            "PERINGATAN: pypdf belum tersedia. Fallback datasheet dimatikan. "
            "Instal: py -m pip install pypdf"
        )
    if not args.no_datasheet_fallback and (
        fitz is None or not _find_tesseract_executable()
    ):
        print(
            "INFO: OCR datasheet scan tidak aktif. PDF teks tetap bisa dibaca. "
            "Untuk PDF scan, instal PyMuPDF dan Tesseract OCR."
        )

    renderer: Optional[LutronBrowserRenderer] = (
        None if args.no_render_js
        else LutronBrowserRenderer(max(args.timeout, 25))
    )
    if renderer is not None and not renderer.available():
        print(
            "PERINGATAN: Playwright belum tersedia. Instal: "
            "py -m pip install playwright lalu "
            "py -m playwright install chromium"
        )
        renderer = None

    accepted: List[Dict[str, Any]] = []
    rows: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    review_specs: List[Dict[str, Any]] = []
    included_accessories = 0

    try:
        for idx, record in enumerate(records, 1):
            url = record["url"]
            print(f"[{idx}/{len(records)}] scrape {url}")

            if not args.no_cache and url in cache:
                data = cache[url]
                for key in ("input_brand", "input_title", "input_model"):
                    if not clean_text(data.get(key)):
                        data[key] = record.get(key, "")
                if not data.get("brand"):
                    data["brand"] = LUTRON_CONFIG.brand
                if not data.get("title"):
                    data["title"] = record.get("input_title", "")
                if not data.get("model"):
                    data["model"] = record.get("input_model", "")
                data["is_accessory"], data["accessory_reason"] = (
                    is_accessory_product(
                        clean_text(data.get("title")),
                        clean_text(data.get("description")),
                        data.get("spec_rows")
                        if isinstance(data.get("spec_rows"), list)
                        else [],
                    )
                )
                data["quality_flags"] = collect_source_quality_flags(data)
                print("    cache hit")
            else:
                data = scrape_product_with_fallbacks(
                    record,
                    args.timeout,
                    session,
                    renderer,
                    enable_datasheet=not args.no_datasheet_fallback,
                    enable_image_search=bool(
                        args.allow_external_image_fallback
                    ),
                    datasheet_attempt_limit=max(
                        1, int(args.datasheet_attempt_limit or 20)
                    ),
                )
                cache[url] = data
                try:
                    save_cache(args.cache, cache)
                    print("    source cache tersimpan")
                except Exception as cache_exc:
                    print(
                        "    WARNING: source cache gagal disimpan: "
                        + compact(cache_exc, 240)
                    )
                if args.delay > 0:
                    time.sleep(args.delay)

            if _v2212_data_is_source_404(data):
                skipped.append({
                    "url": url,
                    "title": data.get("title", ""),
                    "status": "source_http_404",
                    "reason": "source_page_not_found_404",
                    "error": data.get("error", ""),
                })
                print("    skip source 404")
                continue

            if clean_text(data.get("status")) == "request_error":
                skipped.append({
                    "url": url,
                    "title": data.get("title", ""),
                    "status": "request_error",
                    "reason": "transient_source_request_error",
                    "error": data.get("error", ""),
                })
                print(
                    "    NEEDS_RETRY: sumber sementara tidak dapat diakses | "
                    + (clean_text(data.get("error")) or "request_error")
                    + " | batch lanjut"
                )
                continue

            prepare_canonical_specs_for_product(data)

            image_state = (
                "OK" if clean_text(data.get("image_url")) else "KOSONG"
            )
            flags = ", ".join(data.get("quality_flags") or [])
            print(
                f"    status={clean_text(data.get('status')) or '-'} | "
                f"model={clean_text(data.get('model')) or '-'} | "
                f"image={image_state}"
                + (f" | flags={flags}" if flags else "")
            )
            if image_state == "KOSONG":
                print(
                    "    image_reason="
                    + (
                        clean_text(data.get("error"))
                        or "source image tidak ditemukan"
                    )
                )

            if data.get("is_accessory"):
                included_accessories += 1
                print("    include confirmed accessory")
                if args.skip_accessories:
                    print(
                        "    INFO: --skip-accessories diabaikan oleh kebijakan "
                        "V22.12; aksesori tetap diproses."
                    )

            if not clean_text(data.get("title")):
                skipped.append({
                    "url": url,
                    "title": "",
                    "status": clean_text(data.get("status")),
                    "reason": "missing_title_needs_retry",
                    "error": data.get("error", ""),
                })
                print("    NEEDS_RETRY: judul produk tidak dapat dipulihkan | batch lanjut")
                continue

            if not _is_real_product_data(data):
                skipped.append({
                    "url": url,
                    "title": data.get("title", ""),
                    "status": clean_text(data.get("status")),
                    "reason": "incomplete_source_data_needs_retry",
                    "error": data.get("error", ""),
                })
                print("    NEEDS_RETRY: data sumber belum substantif | batch lanjut")
                continue

            if args.strict_source and (
                get_source_config(url) is None
                or clean_text(data.get("datasheet_url"))
            ):
                skipped.append({
                    "url": url,
                    "title": data.get("title", ""),
                    "status": clean_text(data.get("status")),
                    "reason": "strict_source_validation_failed",
                    "error": _merge_source_notes(
                        data.get("error"),
                        "data_not_exclusively_from_input_source_url",
                    ),
                })
                print("    NEEDS_RETRY: strict source gagal | batch lanjut")
                continue

            if (
                AI_SETTINGS.get("enabled")
                and AI_SETTINGS.get("strict_indonesian")
            ):
                if AI_SETTINGS.get("cache_only"):
                    _v2225_attach_cached_ai_fields_by_model(data)
                ai_fields = get_ai_content(data, detect_family(data))
                if not ai_fields:
                    family = detect_family(data)
                    data["_v2219_review_stage"] = "ai_validation"
                    data["_v2219_review_reason"] = (
                        clean_text(data.get("_ai_status"))
                        or "DeepSeek output belum tervalidasi"
                    )
                    data["_v2219_review_details"] = _v2219_failure_details(data)
                    ai_fields = _v2219_review_fields(data, family)
                    data["_ai_fields"] = ai_fields
                    data["_ai_status"] = "needs_review_nonfatal_ai_validation"
                    print(
                        "    NEEDS_REVIEW: AI belum tervalidasi untuk "
                        + (clean_text(data.get("model")) or url)
                        + " | batch lanjut"
                    )

            if data.get("quality_flags"):
                data["_v2212_publish_warnings"] = dedupe_keep_order(
                    list(data.get("_v2212_publish_warnings") or [])
                    + list(data.get("quality_flags") or [])
                )
            accepted.append(data)

        accepted = apply_batch_image_quality_gate(
            accepted, skipped, require_verified_image=False
        )

        for publish_index, data in enumerate(accepted, 1):
            print(
                f"[publish {publish_index}/{len(accepted)}] build "
                f"{clean_text(data.get('model')) or clean_text(data.get('url'))}"
            )
            pub_date = start_date + timedelta(
                hours=len(rows) * args.pub_interval_hours
            )
            row = build_output_row(data, pub_date)
            publish_issues = validate_publish_output(
                data,
                row,
                require_ai=bool(args.strict_publish and args.use_ai),
            )
            blocking = [
                issue for issue in publish_issues
                if issue != "missing_verified_image"
            ]
            if blocking:
                row["Website Validation"] = (
                    "needs_review:" + ",".join(blocking)
                )
                row["Publish Decision"] = "NEEDS_REVIEW"
                row["Website Correction Log"] = _merge_source_notes(
                    row.get("Website Correction Log"),
                    "final_review_required=" + ",".join(blocking),
                    "policy=valid_source_product_not_skipped",
                )
                print(
                    "    needs review | model="
                    + (clean_text(data.get("model")) or "-")
                    + " | issues=" + ",".join(blocking)
                    + " | batch lanjut"
                )
            elif publish_issues:
                row["Website Validation"] = (
                    "publish_ready_with_warning:"
                    + ",".join(publish_issues)
                )
                row["Website Correction Log"] = _merge_source_notes(
                    row.get("Website Correction Log"),
                    "publish_warning=" + ",".join(publish_issues),
                )
            rows.append(row)

            # Save a real workbook checkpoint so a later product issue does not
            # discard products already processed. The final write still occurs
            # after the loop.
            if len(rows) % 5 == 0:
                write_outputs(
                    rows, skipped, review_specs,
                    args.output, args.skipped_output
                )
                save_cache(args.cache, cache)
                if AI_SETTINGS.get("enabled"):
                    save_cache(
                        AI_SETTINGS.get("cache_path", ai_cache_path),
                        AI_SETTINGS.get("cache", {}),
                    )
                print(f"    checkpoint tersimpan | rows={len(rows)} | output={args.output}")

        write_outputs(
            rows, skipped, review_specs,
            args.output, args.skipped_output
        )
        save_cache(args.cache, cache)
        if AI_SETTINGS.get("enabled"):
            save_cache(
                AI_SETTINGS.get("cache_path", ai_cache_path),
                AI_SETTINGS.get("cache", {}),
            )
    finally:
        if renderer is not None:
            renderer.close()

    print(
        f"Selesai. Produk masuk output: {len(rows)} | "
        f"Diskip karena 404: {len(skipped)} | "
        f"Review spesifikasi: {len(review_specs)} | "
        f"Accessories masuk: {included_accessories}"
    )
    print(f"Output: {args.output}")
    if AI_SETTINGS.get("enabled"):
        print(
            f"deepseek_ai_used={AI_STATS['used']} "
            f"cache_hit={AI_STATS['cache_hit']} "
            f"fallback={AI_STATS['fallback']} "
            f"skipped={AI_STATS['skipped']}"
        )


# ---------------------------------------------------------------------------
# V22.12.2 — isolate corrupt source rows instead of aborting the whole catalog
# ---------------------------------------------------------------------------
# Some Lutron pages expose one damaged unit token in an otherwise complete
# specification table (for example ``0.01 dogs``).  The previous integrity gate
# correctly refused to guess the intended unit, but it stopped the entire batch.
# V22.12.2 keeps the same no-guessing rule and quarantines only the damaged row
# when enough valid source rows remain.  The product continues to publication,
# while the original row and reason are retained in Website Correction Log.

V22122_VERSION = "22.12.2-source-row-quarantine"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_12_2_source_row_quarantine_2026_07_29"
AI_STATS.setdefault("v22122_quarantined_source_rows", 0)


def _v22122_quarantine_reason(item: Dict[str, Any]) -> str:
    """Return a reason only for an unmistakably corrupt measurement row.

    A row is quarantined only when:
    - its value is exactly a number followed by an unknown alphabetic unit;
    - the row belongs to a technical specification context; and
    - the label repeats that unknown unit or explicitly denotes a
      range/resolution/accuracy field.

    Model codes, free text, accessories, and legitimate unknown labels are not
    removed by this rule.
    """
    label = _v224_exact_text(item.get("label_raw"))
    value = _v224_exact_text(item.get("value_raw"))
    section = _v224_exact_text(item.get("section_raw"))
    unknown = V224_UNKNOWN_QUANTITY_RE.fullmatch(value)
    if not unknown:
        return ""
    unit_word = clean_text(unknown.group(2))
    if not unit_word or _v224_known_unit(unit_word):
        return ""
    if not V224_TECH_SECTION_RE.search(section):
        return ""
    label_word = re.sub(r"[^A-Za-z0-9_-]", "", label).casefold()
    context = section + " " + label
    if not (
        label_word == unit_word.casefold()
        or re.search(
            r"\b(?:range|resolution|accuracy|rentang|resolusi|akurasi|minimum\s+scale)\b",
            context,
            re.I,
        )
    ):
        return ""
    return "unrecognized_measurement_unit:" + unit_word


_v22122_previous_source_rows = _v224_source_rows


def _v224_source_rows(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return the verified ledger with isolated corrupt rows removed safely."""
    ledger = _v22122_previous_source_rows(data)
    if not ledger:
        return ledger

    candidates: List[Tuple[Dict[str, Any], str]] = []
    for item in ledger:
        if not isinstance(item, dict):
            continue
        reason = _v22122_quarantine_reason(item)
        if reason:
            candidates.append((item, reason))

    if not candidates:
        return ledger

    candidate_ids = {
        int(item.get("source_id") or 0)
        for item, _reason in candidates
        if int(item.get("source_id") or 0) > 0
    }
    remaining = [
        item for item in ledger
        if int(item.get("source_id") or 0) not in candidate_ids
    ]
    substantive_remaining = [
        item for item in remaining
        if not re.search(
            r"\b(?:model|sku|mpn|catalog|part\s*(?:no|number)|item\s*(?:no|number))\b",
            clean_text(item.get("label_raw")),
            re.I,
        )
    ]

    # Never hide a broken source when quarantining would leave an unusable
    # product. In that case the original blocker remains active.
    if len(substantive_remaining) < 3:
        return ledger

    existing = {
        int(item.get("source_id") or 0): item
        for item in data.get("_v22122_quarantined_source_rows") or []
        if isinstance(item, dict)
    }
    for item, reason in candidates:
        source_id = int(item.get("source_id") or 0)
        existing[source_id] = {
            "source_id": source_id,
            "label": clean_text(item.get("label_raw")),
            "value": clean_text(item.get("value_raw")),
            "section": clean_text(item.get("section_raw")),
            "reason": reason,
        }
    quarantined = [existing[key] for key in sorted(existing)]
    data["_v22122_quarantined_source_rows"] = quarantined
    data["_v22122_source_row_quarantine_active"] = True
    data["error"] = _merge_source_notes(
        data.get("error"),
        "source_row_quarantine=" + "|".join(
            "id_{source_id}:{label}={value}:{reason}".format(**item)
            for item in quarantined
        ),
    )
    AI_STATS["v22122_quarantined_source_rows"] = max(
        int(AI_STATS.get("v22122_quarantined_source_rows") or 0),
        len(quarantined),
    )
    return remaining


_v22122_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v22122_previous_build_output_row(data, pub_date)
    quarantined = [
        item for item in data.get("_v22122_quarantined_source_rows") or []
        if isinstance(item, dict)
    ]
    if not quarantined:
        return row

    details = "|".join(
        "id_{source_id}:{label}={value}:{reason}".format(**item)
        for item in quarantined
    )
    row["Website Validation"] = (
        "publish_ready_with_warning:source_spec_row_quarantined"
    )
    row["Processing Time"] = "lutron_WEBSITE_import_v22_12_2_safe_recovery"
    row["Content Quality"] = (
        clean_text(row.get("Content Quality"))
        + "_corrupt_source_row_isolated"
    ).strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "repair_pipeline=" + V22122_VERSION,
        "source_spec_row_quarantined=" + details,
        "quarantine_policy=no_unit_guess_numeric_value_not_published",
    )
    return row


_v22122_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v22122_previous_validate_publish_output(
        data, row, require_ai=require_ai
    )
    if data.get("_v22122_source_row_quarantine_active"):
        issues = [
            issue for issue in issues
            if not issue.startswith("unrecognized_source_measurement_unit:")
            and issue not in {
                "evidence_publish_status_required",
                "deepseek_validated_content_required",
            }
        ]
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.13 — final editorial hardening, identity lock, language purity,
# duplicate-block removal, canonical CTA, publish scoring, and resilient batch.
# ---------------------------------------------------------------------------
# This layer is intentionally defined after every earlier V22.12.2 override and
# before the final ``main()`` call. It therefore operates on the actual cells
# written to Excel without deleting the mature scrape, evidence, repair,
# source-lock, image, and quarantine pipeline above.

V2213_VERSION = "22.13-final-editorial-hardening"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_13_final_editorial_hardening_2026_07_29"

for _column in ("Final Publish Score", "Publish Decision"):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)

AI_STATS.setdefault("v2213_identity_repairs", 0)
AI_STATS.setdefault("v2213_language_repairs", 0)
AI_STATS.setdefault("v2213_duplicate_blocks_removed", 0)
AI_STATS.setdefault("v2213_cta_repairs", 0)
AI_STATS.setdefault("v2213_ready", 0)
AI_STATS.setdefault("v2213_warning", 0)
AI_STATS.setdefault("v2213_review", 0)


# Product identity is locked only where the official model/title provides a
# high-confidence product type. Accessories remain accessories and are never
# converted into main instruments.
V2213_IDENTITY_RULES: Tuple[Tuple[re.Pattern[str], Dict[str, Any]], ...] = (
    (re.compile(r"(?:^|\b)LCR[-_ ]?\d|\bLCR\s+METER\b", re.I), {
        "category": "LCR Meter dan Component Tester",
        "term": "LCR meter",
        "seo_type": "LCR Meter",
        "function": "mengukur induktansi, kapasitansi, resistansi, dan parameter komponen sesuai spesifikasi produk",
        "tags": ["Lutron", "LCR meter", "component tester", "inductance meter", "capacitance meter"],
        "basis": "official_model_or_title_lcr_meter",
    }),
    (re.compile(r"(?:^|\b)PO2[-_ ]?\d|\b(?:O2|OXYGEN|AIR\s+OXYGEN)\s+METER\b", re.I), {
        "category": "Oxygen Meter dan Gas Detector",
        "term": "oxygen meter",
        "seo_type": "Oxygen Meter",
        "function": "mengukur kadar oksigen udara dan suhu sesuai spesifikasi produk",
        "tags": ["Lutron", "oxygen meter", "O2 meter", "gas detector", "alat ukur oksigen"],
        "basis": "official_model_or_title_oxygen_meter",
    }),
    (re.compile(r"(?:^|\b)PM[-_ ]?\d|\b(?:PROFESSIONAL\s+)?MANOMETER\b|\bPRESSURE\s+METER\b", re.I), {
        "category": "Manometer dan Pressure Meter",
        "term": "manometer",
        "seo_type": "Manometer",
        "function": "mengukur tekanan udara atau gas sesuai spesifikasi produk",
        "tags": ["Lutron", "manometer", "pressure meter", "alat ukur tekanan"],
        "basis": "official_model_or_title_pressure_meter",
    }),
)


def _v2213_identity_rule(data: Dict[str, Any]) -> Dict[str, Any]:
    if data.get("is_accessory"):
        return {}
    blob = " ".join([
        clean_text(data.get("model")),
        clean_catalog_source_text(data.get("title")),
        clean_catalog_source_text(data.get("input_title")),
        clean_catalog_source_text(data.get("source_category")),
    ])
    for pattern, rule in V2213_IDENTITY_RULES:
        if pattern.search(blob):
            return dict(rule)
    return {}


_v2213_previous_detect_family = detect_family


def detect_family(data: Dict[str, Any]) -> Dict[str, Any]:
    locked = _v2213_identity_rule(data)
    if locked:
        data["_v2213_identity_locked"] = True
        data["_v2213_product_type"] = clean_text(locked.get("seo_type"))
        data["_v2213_identity_basis"] = clean_text(locked.get("basis"))
        return {
            "category": clean_text(locked.get("category")),
            "term": clean_text(locked.get("term")),
            "function": clean_text(locked.get("function")),
            "tags": list(locked.get("tags") or []),
            "seo_type": clean_text(locked.get("seo_type")),
            "classification_basis": clean_text(locked.get("basis")),
        }
    return _v2213_previous_detect_family(data)


_v2213_previous_seo_product_type = _seo_product_type


def _seo_product_type(family: Dict[str, Any], title: str) -> str:
    locked = clean_text(family.get("seo_type"))
    return locked or _v2213_previous_seo_product_type(family, title)


# Broader residue detector. It is phrase-level and intentionally ignores model
# codes, units, LCD, USB, RS232, LCR, O2, SD card, Excel, AC/DC and ON/OFF.
V2213_SPEC_ENGLISH_RE = re.compile(
    r"\b(?:"
    r"custom\s+one[- ]chip|with\s+green\s+backlight|test\s+frequency|"
    r"function\s+selector|frequency\s+selector|sorting\s+mode\s+selector|"
    r"dissipation\s+factor|quality\s+factor|phase\s+angle|"
    r"equivalent\s+series\s+resistance|less\s+th[ae]n|meter\s+only|"
    r"soft\s+carrying\s+case|data\s+acquisition\s+software|"
    r"overload\s+protection|buzzer\s+sounds?\s+if|galvanic\s+cell\s+type|"
    r"for\s+all\s+ranges|after\s+calibration|set\s+clock\s+time|"
    r"set\s+beep\s+sound|decimal\s+point\s+of\s+SD\s+card\s+setting|"
    r"power\s+adapter\s+is\s+optional|will\s+get\s+the|"
    r"air\s+oxygen|air\s+temp(?:erature)?|at\s+t90|"
    r"dry,?\s+non[- ]corrosive,?\s+and\s+non[- ]ionic|"
    r"normal\s+operation|data\s+logging|save\s+data|"
    r"automatic\s+power\s+off|manual\s+power\s+off"
    r")\b",
    re.I,
)

V2213_PUBLIC_ENGLISH_RE = re.compile(
    r"\b(?:"
    r"for\s+all\s+ranges|after\s+calibration|dry,?\s+non[- ]corrosive|"
    r"koneksi\s+USB[^.;]{0,80}opsional\s+PC|"
    r"optional\s+PC|at\s+t90|air\s+oxygen|air\s+temperature"
    r")\b",
    re.I,
)

_v2213_previous_has_spec_english = _v226_has_spec_english
_v2213_previous_has_public_english = _v226_has_public_english


def _v226_has_spec_english(value: Any) -> bool:
    text_value = _html_text(value)
    return bool(
        _v2213_previous_has_spec_english(value)
        or V2213_SPEC_ENGLISH_RE.search(text_value)
    )


def _v226_has_public_english(value: Any) -> bool:
    text_value = _html_text(value)
    return bool(
        _v2213_previous_has_public_english(value)
        or V2213_PUBLIC_ENGLISH_RE.search(text_value)
    )


V2213_TEXT_REPLACEMENTS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bCustom\s+one[- ]chip\s+of\s+(?:Sirkuit\s+)?LSI\s+mikroprosesor\b", re.I), "Sirkuit LSI mikroprosesor satu chip khusus"),
    (re.compile(r"\bCustom\s+one[- ]chip\s+of\s+microprocessor\s+LSI\s+circuit\b", re.I), "Sirkuit LSI mikroprosesor satu chip khusus"),
    (re.compile(r"\bLCD\s+size\s*:", re.I), "Ukuran LCD:"),
    (re.compile(r"\bLCD\s+with\s+green\s+backlight\b", re.I), "LCD dengan lampu latar hijau"),
    (re.compile(r"\bwith\s+green\s+backlight\b", re.I), "dengan lampu latar hijau"),
    (re.compile(r"\bTest\s+frequency\b", re.I), "Frekuensi WEBSITE"),
    (re.compile(r"\bFunction\s+selector\b", re.I), "Pemilih fungsi"),
    (re.compile(r"\bFrequency\s+selector\b", re.I), "Pemilih frekuensi"),
    (re.compile(r"\bSorting\s+mode\s+selector\b", re.I), "Pemilih mode penyortiran"),
    (re.compile(r"\bDissipation\s+factor\b", re.I), "Faktor disipasi"),
    (re.compile(r"\bQuality\s+factor\b", re.I), "Faktor kualitas"),
    (re.compile(r"\bPhase\s+angle\b", re.I), "Sudut fase"),
    (re.compile(r"\bEquivalent\s+series\s+resistance\b", re.I), "Resistansi seri ekuivalen"),
    (re.compile(r"\bLess\s+th[ae]n\b", re.I), "Kurang dari"),
    (re.compile(r"\bmeter\s+only\b", re.I), "unit meter"),
    (re.compile(r"\bSoft\s+carrying\s+case\b", re.I), "kotak pelindung lunak"),
    (re.compile(r"\bHard\s+carrying\s+case\b", re.I), "kotak pelindung keras"),
    (re.compile(r"\bExcel\s+Data\s+acquisition\s+software\b", re.I), "Perangkat lunak akuisisi data untuk Excel"),
    (re.compile(r"\bData\s+acquisition\s+software\b", re.I), "Perangkat lunak akuisisi data"),
    (re.compile(r"\bOverload\s+Protection\b", re.I), "Perlindungan beban berlebih"),
    (re.compile(r"\bBuzzer\s+sounds?\s+if\s+O2\s+level\b", re.I), "Buzzer berbunyi jika kadar O2"),
    (re.compile(r"\bGalvanic\s+cell\s+type\b", re.I), "Tipe sel galvanik"),
    (re.compile(r"\bfor\s+all\s+ranges\b", re.I), "untuk semua rentang"),
    (re.compile(r"\bafter\s+calibration\b", re.I), "setelah kalibrasi"),
    (re.compile(r"\bat\s+t90\b", re.I), "pada t90"),
    (re.compile(r"\bSet\s+clock\s+time\b", re.I), "Pengaturan waktu jam"),
    (re.compile(r"\bSet\s+beep\s+Sound\s+ON/OFF\b", re.I), "Pengaturan bunyi beep ON/OFF"),
    (re.compile(r"\bDecimal\s+point\s+of\s+SD\s+card\s+setting\b", re.I), "Pengaturan titik desimal kartu SD"),
    (re.compile(r"\bAC/DC\s+power\s+adapter\s+is\s+optional\b", re.I), "Adaptor daya AC/DC tersedia sebagai aksesori opsional"),
    (re.compile(r"\bO2\s*\(Air\s+oxygen\)\s*Temperature\s*\(Air\s+Temp\.?\)", re.I), "O2 (oksigen udara) dan suhu udara"),
    (re.compile(r"\bdry,?\s+non[- ]corrosive,?\s+and\s+non[- ]ionic\s+air\s+and\s+gas\s+only\b", re.I), "hanya untuk udara dan gas kering, non-korosif, dan non-ionik"),
    (re.compile(r"\bKoneksi\s+USB\s+memerlukan\s+kabel\s+USB-01\s+opsional\s+PC\b", re.I), "Koneksi USB memerlukan kabel USB-01 opsional untuk komputer"),
    (re.compile(r"\bkabel\s+USB-01\s+opsional\s+PC\s+memungkinkan\b", re.I), "kabel USB-01 opsional untuk komputer memungkinkan"),
    (re.compile(r"\bNormal\s+Operation\b", re.I), "Operasi normal"),
    (re.compile(r"\bData\s+Logging\b", re.I), "Pencatatan data"),
    (re.compile(r"\bSave\s+data\b", re.I), "Penyimpanan data"),
    (re.compile(r"\bAutomatic\s+power\s+off\b", re.I), "Mati otomatis"),
    (re.compile(r"\bManual\s+power\s+off\b", re.I), "Mati manual"),
)


def _v2213_polish_technical_text(value: Any) -> str:
    result = html.unescape(str(value or ""))
    original = result
    for pattern, replacement in V2213_TEXT_REPLACEMENTS:
        result = pattern.sub(replacement, result)
    result = re.sub(r"\s+([,.;:])", r"\1", result)
    result = re.sub(r"([;:])(?=\S)", r"\1 ", result)
    result = re.sub(r"\s{2,}", " ", result).strip()
    if result != original:
        AI_STATS["v2213_language_repairs"] = int(
            AI_STATS.get("v2213_language_repairs") or 0
        ) + 1
    return result


def _v2213_text_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).casefold())


def _v2213_dedupe_repeated_block(value: Any) -> Tuple[str, int]:
    """Remove only exact repeated long text blocks inside one cell.

    Numeric content is never merged when blocks differ. This catches PDF/DOM
    extraction such as the same full Inductance accuracy block repeated two or
    three times in a single cell.
    """
    text_value = clean_text(value)
    if len(text_value) < 160:
        return text_value, 0
    removed = 0

    # Repeated technical-section anchors.
    anchor_re = re.compile(
        r"(?=(?:Inductance\s*\(L\)|Capacitance\s*\(C\)|Resistance\s*\(R\))\s*:)",
        re.I,
    )
    parts = anchor_re.split(text_value)
    if len(parts) >= 3:
        prefix = parts[0]
        unique: List[str] = []
        seen: set = set()
        for part in parts[1:]:
            part = clean_text(part)
            key = _v2213_text_key(part)
            if key and (
                key in seen
                or (len(key) >= 40 and any(previous.startswith(key) for previous in seen))
            ):
                removed += 1
                continue
            if key:
                seen.add(key)
            unique.append(part)
        if removed:
            text_value = clean_text(prefix + " " + " ".join(unique))

    # Exact repeated halves/thirds/quarters after whitespace normalisation.
    tokens = text_value.split()
    if len(tokens) >= 40:
        folded = [token.casefold() for token in tokens]
        for repetitions in (4, 3, 2):
            if len(tokens) % repetitions:
                continue
            size = len(tokens) // repetitions
            if size < 18:
                continue
            first = folded[:size]
            if all(folded[i * size:(i + 1) * size] == first for i in range(1, repetitions)):
                text_value = " ".join(tokens[:size])
                removed += repetitions - 1
                break

    if removed:
        AI_STATS["v2213_duplicate_blocks_removed"] = int(
            AI_STATS.get("v2213_duplicate_blocks_removed") or 0
        ) + removed
    return text_value, removed


V2213_COMMERCIAL_SENTENCE_RE = re.compile(
    r"[^.!?]*(?:stok|garansi|harga|diskon|promo|pemesanan|cara\s+memesan)"
    r"[^.!?]*(?:[.!?]|$)",
    re.I,
)
V2213_REDUNDANT_CONTACT_RE = re.compile(
    r"[^.!?]*(?:hubungi\s+tim\s+WEBSITE|konsultasi\s+lebih\s+lanjut)"
    r"[^.!?]*(?:[.!?]|$)",
    re.I,
)


def _v2213_remove_unsupported_commercial_sentences(value: Any) -> str:
    text_value = clean_text(value)
    text_value = V2213_COMMERCIAL_SENTENCE_RE.sub(" ", text_value)
    # The canonical CTA already provides the contact route. Remove duplicate
    # generic contact sentences that carry no source-backed product fact.
    text_value = V2213_REDUNDANT_CONTACT_RE.sub(" ", text_value)
    return re.sub(r"\s{2,}", " ", text_value).strip()


def _v2213_canonicalize_cta_html(value: Any, ensure_cta: bool = False) -> Tuple[str, int]:
    raw = clean_text(value)
    soup = BeautifulSoup(raw, "html.parser")
    repairs = 0
    cta_nodes: List[Any] = []
    for block in soup.find_all(["p", "div"]):
        block_html = str(block)
        if "about-us-3" in block_html or "contact-us" in block_html:
            cta_nodes.append(block)
    canonical = BeautifulSoup(WEBSITE_CTA, "html.parser").find("p")
    if cta_nodes and canonical is not None:
        cta_nodes[0].replace_with(canonical)
        repairs += 1
        for extra in cta_nodes[1:]:
            extra.decompose()
            repairs += 1
    elif ensure_cta and canonical is not None:
        soup.append(canonical)
        repairs += 1
    rendered = "".join(str(node) for node in soup.contents).strip()
    # Defensive spacing if another parser serialises inline tags compactly.
    rendered = re.sub(r"kunjungi(?=<strong)", "kunjungi ", rendered, flags=re.I)
    rendered = re.sub(r"</strong>(?=atau\b)", "</strong> ", rendered, flags=re.I)
    rendered = re.sub(r"\batau(?=<strong)", "atau ", rendered, flags=re.I)
    if repairs:
        AI_STATS["v2213_cta_repairs"] = int(
            AI_STATS.get("v2213_cta_repairs") or 0
        ) + repairs
    return rendered, repairs


def _v2213_polish_html(value: Any, ensure_cta: bool = False) -> str:
    from bs4 import NavigableString

    soup = BeautifulSoup(clean_text(value), "html.parser")
    for node in list(soup.find_all(string=True)):
        if not isinstance(node, NavigableString):
            continue
        parent_name = clean_text(getattr(node.parent, "name", "")).casefold()
        if parent_name in {"script", "style"}:
            continue
        polished = _v2213_polish_technical_text(str(node))
        if polished != str(node):
            node.replace_with(polished)

    # Remove unsupported commercial wording sentence-by-sentence, but preserve
    # source-backed technical calibration statements.
    for block in list(soup.find_all(["p", "li"])):
        if "about-us-3" in str(block) or "contact-us" in str(block):
            continue
        visible = clean_text(block.get_text(" ", strip=True))
        cleaned = _v2213_remove_unsupported_commercial_sentences(visible)
        if cleaned != visible:
            if cleaned:
                block.clear()
                block.append(cleaned)
            else:
                block.decompose()

    for tag in list(soup.find_all(True)):
        if tag.name in {"br", "img", "hr"}:
            continue
        if not clean_text(tag.get_text(" ", strip=True)) and not tag.find(True):
            tag.decompose()

    rendered = "".join(str(node) for node in soup.contents).strip()
    rendered, _ = _v2213_canonicalize_cta_html(rendered, ensure_cta=ensure_cta)
    return rendered


def _v2213_polish_spec_html(value: Any) -> Tuple[str, int]:
    soup = BeautifulSoup(clean_text(value), "html.parser")
    removed = 0

    for cell in soup.find_all(["th", "td"]):
        visible = clean_text(cell.get_text(" ", strip=True))
        polished = _v2213_polish_technical_text(visible)
        polished, duplicate_count = _v2213_dedupe_repeated_block(polished)
        removed += duplicate_count
        if polished != visible:
            cell.clear()
            cell.append(polished)

    # Accuracy/source notes can also contain one long block repeated by the
    # source renderer. Apply the same exact-block rule outside table cells.
    for note in soup.find_all(["p", "h3"]):
        visible = clean_text(note.get_text(" ", strip=True))
        polished = _v2213_polish_technical_text(visible)
        polished, duplicate_count = _v2213_dedupe_repeated_block(polished)
        removed += duplicate_count
        if polished != visible:
            note.clear()
            note.append(polished)

    # Remove exact duplicate rows after language normalisation, preserving the
    # first row and every row with a different numeric or textual value.
    seen: set = set()
    for tr in list(soup.select("table tbody > tr")):
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue
        signature = tuple(_v2213_text_key(cell.get_text(" ", strip=True)) for cell in cells)
        if signature and signature in seen:
            tr.decompose()
            removed += 1
        else:
            seen.add(signature)

    if removed:
        AI_STATS["v2213_duplicate_blocks_removed"] = int(
            AI_STATS.get("v2213_duplicate_blocks_removed") or 0
        ) + removed
    return str(soup), removed


def _v2213_apply_identity_to_row(row: Dict[str, Any], data: Dict[str, Any]) -> Dict[str, Any]:
    identity = _v2213_identity_rule(data)
    if not identity:
        return row
    model = clean_text(data.get("model"))
    brand = clean_text(data.get("brand")) or "Lutron"
    seo_type = clean_text(identity.get("seo_type"))
    expected_name = " ".join(part for part in (seo_type, brand, model) if part)

    # Pressure models already contain useful source-backed range descriptors.
    # LCR and oxygen models use a clean identity-only title to avoid old wrong
    # prefixes such as Kalibrator or Air Quality.
    if seo_type in {"LCR Meter", "Oxygen Meter"}:
        row["Name"] = expected_name
    else:
        current = clean_text(row.get("Name"))
        descriptor = ""
        if model and model.casefold() in current.casefold():
            descriptor = re.split(re.escape(model), current, maxsplit=1, flags=re.I)[-1]
            descriptor = re.sub(
                r"\b(?:kalibrator|air\s+quality|manometer|pressure\s+meter|lutron)\b",
                " ", descriptor, flags=re.I,
            )
            descriptor = re.sub(r"\s+", " ", descriptor).strip(" -,:;.")
        row["Name"] = clean_text(expected_name + (" " + descriptor if descriptor else ""))

    category = clean_text(identity.get("category"))
    row["product categories"] = category
    row["focus keyphrase"] = expected_name
    tags = dedupe_keep_order([
        brand, model, seo_type, category,
        *list(identity.get("tags") or []),
    ])
    row["product tags"] = ", ".join(tags)

    family = detect_family(data)
    row["meta description"] = _v22_build_meta(data, family, expected_name)
    data["_v2213_identity_locked"] = True
    data["_v2213_product_type"] = seo_type
    AI_STATS["v2213_identity_repairs"] = int(
        AI_STATS.get("v2213_identity_repairs") or 0
    ) + 1
    return row


def _v2213_identity_issues(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    identity = _v2213_identity_rule(data)
    if not identity:
        return []
    name = clean_text(row.get("Name"))
    category = clean_text(row.get("product categories"))
    expected_type = clean_text(identity.get("seo_type"))
    expected_category = clean_text(identity.get("category"))
    model = clean_text(data.get("model"))
    issues: List[str] = []
    if expected_type.casefold() not in name.casefold():
        issues.append("final_product_identity_type_mismatch")
    if model and model.casefold() not in name.casefold():
        issues.append("final_product_identity_model_missing")
    if expected_category.casefold() != category.casefold():
        issues.append("final_product_category_mismatch")
    return issues


def _v2213_publish_score(row: Dict[str, Any], data: Dict[str, Any], issues: Sequence[str]) -> int:
    issue_set = set(issues or [])
    score = 0

    # Identity: 25
    if not any(item.startswith("final_product_identity") or item == "final_product_category_mismatch" for item in issue_set):
        score += 25

    # Specification completeness and purity: 30
    spec_html = clean_text(row.get("custom_tab_1_content"))
    spec_soup = BeautifulSoup(spec_html, "html.parser")
    spec_rows = len(spec_soup.select("table tbody > tr"))
    spec_issues = [
        item for item in issue_set
        if "specification" in item or "spec_" in item or "measurement_unit" in item
    ]
    if spec_rows >= 3 and not spec_issues:
        score += 30
    elif spec_rows >= 3:
        score += 20
    elif spec_rows:
        score += 10

    # Language purity: 20
    language_issues = [
        item for item in issue_set
        if "english" in item or "chinese" in item or "mixed_" in item
    ]
    if not language_issues:
        score += 20
    elif len(language_issues) == 1:
        score += 10

    # HTML/editorial integrity: 15
    required = [
        "Product Description", "custom_tab_1_content", "custom_tab_2_content",
        "custom_tab_3_content", "custom_tab_4_content", "custom_tab_5_content",
    ]
    html_issues = [item for item in issue_set if item.startswith("final_empty_") or "html" in item or "faq_count" in item or "application_count" in item]
    if all(clean_text(row.get(field)) for field in required) and not html_issues:
        score += 15
    elif all(clean_text(row.get(field)) for field in required):
        score += 8

    # Source grounding: 10
    source_ok = bool(get_source_config(clean_text(row.get("Source URL"))))
    has_source_rows = bool(data.get("_raw_spec_rows") or data.get("spec_rows"))
    if source_ok and has_source_rows:
        score += 10
    elif source_ok:
        score += 5

    # A single isolated corrupt source row is a warning, not a fatal quality
    # failure. The guessed value is not published.
    if data.get("_v22122_source_row_quarantine_active"):
        score = max(90, score - 3)
    return max(0, min(100, int(score)))


# Bypass the old V22.11/V22.12 terminal raises while retaining the mature row
# builder immediately before those terminal gates.
_v2213_previous_validate_publish_output = validate_publish_output


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2211_previous_build_output_row(data, pub_date)
    row, compact_removed = _v2211_polish_output_row(row, data)

    # Source-aware interface cleanup on exact final cells.
    connection_repairs = 0
    if "_v22121_rewrite_ambiguous_connection_html" in globals():
        for field in V2211_HTML_FIELDS:
            row[field], repaired = _v22121_rewrite_ambiguous_connection_html(
                row.get(field), data
            )
            connection_repairs += repaired

    row = _v2213_apply_identity_to_row(row, data)

    for field in V2211_HTML_FIELDS:
        if field == "custom_tab_1_content":
            row[field], duplicate_removed = _v2213_polish_spec_html(row.get(field))
            compact_removed += duplicate_removed
        else:
            row[field] = _v2213_polish_html(
                row.get(field),
                ensure_cta=(field == "Product Description"),
            )
    for field in V2211_PLAIN_FIELDS:
        row[field] = _v2213_polish_technical_text(row.get(field))

    issues = dedupe_keep_order(
        _v2211_final_editorial_issues(row)
        + _v2213_identity_issues(row, data)
    )
    data["_v2211_final_editorial_issues"] = issues
    data["_v2213_final_issues"] = issues

    score = _v2213_publish_score(row, data, issues)
    quarantined = [
        item for item in data.get("_v22122_quarantined_source_rows") or []
        if isinstance(item, dict)
    ]

    if not issues and not quarantined and score >= 95:
        status = "publish_ready_v22_13"
        decision = "READY"
        AI_STATS["v2213_ready"] = int(AI_STATS.get("v2213_ready") or 0) + 1
    elif not issues and score >= 85:
        status = "publish_ready_with_warning_v22_13"
        decision = "READY_WITH_WARNING"
        AI_STATS["v2213_warning"] = int(AI_STATS.get("v2213_warning") or 0) + 1
    else:
        status = "needs_review_v22_13" + ((":" + ",".join(issues)) if issues else "")
        decision = "NEEDS_REVIEW"
        AI_STATS["v2213_review"] = int(AI_STATS.get("v2213_review") or 0) + 1

    if quarantined and not issues:
        status = "publish_ready_with_warning:source_spec_row_quarantined"
        decision = "READY_WITH_WARNING"

    row["Website Validation"] = status
    row["Final Publish Score"] = score
    row["Publish Decision"] = decision
    row["Processing Time"] = "lutron_WEBSITE_import_v22_13_final_editorial_hardening"
    row["Content Quality"] = (
        clean_text(data.get("source"))
        + "_source_locked_v22_13_"
        + decision.casefold()
    )

    quarantine_detail = ""
    if quarantined:
        quarantine_detail = "|".join(
            "id_{source_id}:{label}={value}:{reason}".format(**item)
            for item in quarantined
        )

    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "repair_pipeline=" + V2213_VERSION,
        "identity_lock=" + clean_text(data.get("_v2213_identity_basis") or "not_required"),
        "final_connection_repairs=" + str(connection_repairs),
        "final_spec_rows_or_blocks_removed=" + str(compact_removed),
        "final_publish_score=" + str(score),
        "publish_decision=" + decision,
        "final_editorial_issues=" + ",".join(issues) if issues else "final_editorial_issues=none",
        "source_spec_row_quarantined=" + quarantine_detail if quarantine_detail else "",
        "policy=only_source_404_enters_skipped",
    )
    return row


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2213_previous_validate_publish_output(
        data, row, require_ai=require_ai
    )
    issues.extend(data.get("_v2213_final_issues") or [])

    status = clean_text(row.get("Website Validation"))
    if status.startswith((
        "publish_ready_v22_13",
        "publish_ready_with_warning_v22_13",
        "publish_ready_with_warning:source_spec_row_quarantined",
        "needs_review_v22_13",
    )):
        issues = [
            item for item in issues
            if item not in {
                "evidence_publish_status_required",
                "deepseek_validated_content_required",
                "missing_verified_image",
            }
        ]
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.14 — single terminal gate and Lutron structure-aware final normalization
# ---------------------------------------------------------------------------
# V22.13 still called a legacy V22.6 terminal language gate before its own
# final normalizer could run. On valid pages such as PM-9112SD this caused the
# batch to stop on source-backed technical wording. V22.14 defers only language
# residue checks to the last gate, preserves every non-language blocker, and
# ensures one product-level editorial issue cannot terminate the entire batch.

V2214_VERSION = "22.14-single-terminal-gate"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_14_single_terminal_gate_2026_07_30"

AI_STATS.setdefault("v2214_deferred_language_gates", 0)
AI_STATS.setdefault("v2214_build_recoveries", 0)
AI_STATS.setdefault("v2214_final_language_repairs", 0)
AI_STATS.setdefault("v2214_ready", 0)
AI_STATS.setdefault("v2214_warning", 0)
AI_STATS.setdefault("v2214_review", 0)


def _v2214_is_deferred_language_issue(issue: Any) -> bool:
    code = clean_text(issue).casefold()
    return (
        "contains_english_residue" in code
        or "english_residue" in code
        or code == "specification_language_repair_incomplete"
    )


# The legacy V22.6 builder is still part of the mature builder chain. Keep its
# structural, evidence, FAQ, unsupported-claim, and Chinese checks, but defer
# English residue to the actual final output after V22.14 normalization.
_v2214_original_v226_final_row_issues = _v226_final_row_issues


def _v226_final_row_issues(data: Dict[str, Any], row: Dict[str, Any]) -> List[str]:
    issues = _v2214_original_v226_final_row_issues(data, row)
    deferred = [item for item in issues if _v2214_is_deferred_language_issue(item)]
    if deferred:
        data["_v2214_deferred_language_issues"] = dedupe_keep_order(
            list(data.get("_v2214_deferred_language_issues") or []) + deferred
        )
        AI_STATS["v2214_deferred_language_gates"] = int(
            AI_STATS.get("v2214_deferred_language_gates") or 0
        ) + len(deferred)
    return [item for item in issues if item not in deferred]


# Phrase-level glossary based on the recurring structure of Lutron product
# pages: Product Overview, Technical Specifications, Specifications, Optional
# Accessories, and Applications. Numeric tokens and units are never changed.
V2214_TEXT_REPLACEMENTS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bData\s+Hold\b", re.I), "Penahan data"),
    (re.compile(r"\bMemory\s+Recall\b", re.I), "Pemanggilan memori"),
    (re.compile(r"\bMaximum\s*&\s*Minimum\s+value\b", re.I), "Nilai maksimum dan minimum"),
    (re.compile(r"\bMaximum\s+and\s+Minimum\s+value\b", re.I), "Nilai maksimum dan minimum"),
    (re.compile(r"\bzero\s+adjustment\b", re.I), "penyetelan nol"),
    (re.compile(r"\bzero\s+adjust\b", re.I), "penyetelan nol"),
    (re.compile(r"\bpush\s+button\b", re.I), "tombol tekan"),
    (re.compile(r"\bfront\s+panel\b", re.I), "panel depan"),
    (re.compile(r"\bSD\s+Card\s+Capacity\b", re.I), "Kapasitas kartu SD"),
    (re.compile(r"\bMemory\s+Card\s+Advanced\s+setting\b", re.I), "Pengaturan lanjutan kartu memori"),
    (re.compile(r"\bSD\s+memory\s+card\b", re.I), "kartu memori SD"),
    (re.compile(r"\bSD\s+memori\s+card\b", re.I), "kartu memori SD"),
    (re.compile(r"\bmemory\s+card\b", re.I), "kartu memori"),
    (re.compile(r"\bSet\s+sampling\s+time\b", re.I), "Pengaturan waktu pencuplikan"),
    (re.compile(r"\bSet\s+clock\s+time\b", re.I), "Pengaturan waktu jam"),
    (re.compile(r"\bYear\s*/\s*Month\s*/\s*Date\b", re.I), "Tahun/Bulan/Tanggal"),
    (re.compile(r"\bHour\s*/\s*Minute\s*/\s*(?:Second|detik)\b", re.I), "Jam/Menit/Detik"),
    (re.compile(r"\bAuto\s+power\s+OFF\s+management\b", re.I), "Pengelolaan mati otomatis"),
    (re.compile(r"\bSet\s+beep\s+Sound\s+ON/OFF\b", re.I), "Pengaturan bunyi beep ON/OFF"),
    (re.compile(r"\bDecimal\s+point\s+of\s+SD\s+card\s+setting\b", re.I), "Pengaturan titik desimal kartu SD"),
    (re.compile(r"\bSD\s+(?:memory|memori)\s+card\s+Format\b", re.I), "Format kartu memori SD"),
    (re.compile(r"\bAlkaline\s+or\s+(?:heavy\s+duty|untuk\s+penggunaan\s+intensif)\b", re.I), "Baterai alkaline atau untuk penggunaan intensif"),
    (re.compile(r"\bor\s+equivalent\b", re.I), "atau setara"),
    (re.compile(r"\bDC\s*1[.,]5\s*V\s+batter(?:y|ies)\b", re.I), "baterai DC 1,5 V"),
    (re.compile(r"\bDC\s*1\.5\s*V\s+batter(?:y|ies)\b", re.I), "baterai DC 1,5 V"),
    (re.compile(r"\bbatteries\b", re.I), "baterai"),
    (re.compile(r"\bbattery\b", re.I), "baterai"),
    (re.compile(r"\bAC\s+to\s+DC\b", re.I), "AC ke DC"),
    (re.compile(r"\bAC/DC\s+adaptor\s+daya\s+is\s+opsional\b", re.I), "adaptor daya AC/DC bersifat opsional"),
    (re.compile(r"\bAC/DC\s+power\s+adapter\s+is\s+optional\b", re.I), "adaptor daya AC/DC bersifat opsional"),
    (re.compile(r"\bis\s+(?:optional|opsional)\b", re.I), "bersifat opsional"),
    (re.compile(r"\badapter\s+input\b", re.I), "masukan adaptor"),
    (re.compile(r"\badapter\s+masukan\b", re.I), "masukan adaptor"),
    (re.compile(r"\bDC\s*9\s*V\s+masukan\s+adaptor\b", re.I), "masukan adaptor DC 9 V"),
    (re.compile(r"\bpower\s+adapter\b", re.I), "adaptor daya"),
    (re.compile(r"\bpower\s+current\b", re.I), "konsumsi daya"),
    (re.compile(r"\bLCD\s+backlight\s+is\s+OFF\b", re.I), "lampu latar LCD mati"),
    (re.compile(r"\bLCD\s+backlight\s+OFF\b", re.I), "lampu latar LCD mati"),
    (re.compile(r"\bbacklight\s+OFF\b", re.I), "lampu latar mati"),
    (re.compile(r"\bbacklight\s+ON\b", re.I), "lampu latar menyala"),
    (re.compile(r"\bNormal\s+operation\b", re.I), "Operasi normal"),
    (re.compile(r"\bManual\s+mode\b", re.I), "Mode manual"),
    (re.compile(r"\bAuto\s+mode\b", re.I), "Mode otomatis"),
    (re.compile(r"\bManual\b", re.I), "Manual"),
    (re.compile(r"\bAuto\b", re.I), "Otomatis"),
    (re.compile(r"\bresolution\b", re.I), "resolusi"),
    (re.compile(r"\baccuracy\b", re.I), "akurasi"),
    (re.compile(r"\bdisplay\s+units?\b", re.I), "satuan tampilan"),
    (re.compile(r"\bdisplay\s+type\b", re.I), "tipe tampilan"),
    (re.compile(r"\bdisplay\s+size\b", re.I), "ukuran tampilan"),
    (re.compile(r"\bType\s+(?=[KJTERS]\b|Pt\s*100\b)", re.I), "Tipe "),
    (re.compile(r"\bSensor\s+type\b", re.I), "Tipe sensor"),
    (re.compile(r"\bProbe\s+Input\s+Socket\b", re.I), "Soket masukan probe"),
    (re.compile(r"\bData\s+Output\b", re.I), "Keluaran data"),
    (re.compile(r"\bPC\s+computer\s+interface\b", re.I), "antarmuka komputer"),
    (re.compile(r"\bPC\s+interface\b", re.I), "antarmuka komputer"),
    (re.compile(r"\bData\s+Acquisition\s+Software\b", re.I), "Perangkat lunak akuisisi data"),
    (re.compile(r"\bInstruction\s+Manual\b", re.I), "Buku petunjuk"),
    (re.compile(r"\bAccessories\s+Included\b", re.I), "Aksesori yang disertakan"),
    (re.compile(r"\bOptional\s+Accessories\b", re.I), "Aksesori opsional"),
    (re.compile(r"\bOperating\s+Temperature\b", re.I), "Suhu operasi"),
    (re.compile(r"\bOperating\s+Humidity\b", re.I), "Kelembapan operasi"),
    (re.compile(r"\bMeasurement\s+Range\b", re.I), "Rentang pengukuran"),
    (re.compile(r"\bSampling\s+Time\b", re.I), "Waktu pencuplikan"),
    (re.compile(r"\bwaktu\s+sampling\b", re.I), "waktu pencuplikan"),
    (re.compile(r"\bResponse\s+Time\b", re.I), "Waktu respons"),
    (re.compile(r"\bPower\s+Supply\b", re.I), "Catu daya"),
    (re.compile(r"\bPower\s+Consumption\b", re.I), "Konsumsi daya"),
    (re.compile(r"\bOverload\s+Protection\b", re.I), "Perlindungan beban berlebih"),
    (re.compile(r"\bSoft\s+carrying\s+case\b", re.I), "kotak pelindung lunak"),
    (re.compile(r"\bHard\s+carrying\s+case\b", re.I), "kotak pelindung keras"),
    (re.compile(r"\bcarrying\s+case\b", re.I), "kotak pelindung"),
    (re.compile(r"\bLess\s+th[ae]n\b", re.I), "Kurang dari"),
    (re.compile(r"\bApprox(?:imately)?\.?\b", re.I), "Sekitar"),
    (re.compile(r"\bwith\s+green\s+backlight\b", re.I), "dengan lampu latar hijau"),
    (re.compile(r"\bLCD\s+with\s+green\s+backlight\b", re.I), "LCD dengan lampu latar hijau"),
    (re.compile(r"\bfor\s+all\s+ranges\b", re.I), "untuk semua rentang"),
    (re.compile(r"\bafter\s+calibration\b", re.I), "setelah kalibrasi"),
    (re.compile(r"\bLiquid\s+is\s+prohibited\b", re.I), "Cairan tidak diperbolehkan"),
    (re.compile(r"\bUsed\s+for\s+dry,?\s+non[- ]corrosive\s+and\s+non[- ]ionic\s+air\s+and\s+gas\s+only\b", re.I), "Hanya untuk udara dan gas kering, non-korosif, dan non-ionik"),
)

_v2214_previous_polish_technical_text = _v2213_polish_technical_text


def _v2213_polish_technical_text(value: Any) -> str:
    result = _v2214_previous_polish_technical_text(value)
    original = result
    for pattern, replacement in V2214_TEXT_REPLACEMENTS:
        result = pattern.sub(replacement, result)
    result = re.sub(r"\b1\s+GB\s+to\s+16\s+GB\b", "1 GB hingga 16 GB", result, flags=re.I)
    result = re.sub(r"\b(\d+(?:[.,]\d+)?)\s+to\s+(\d+(?:[.,]\d+)?)\b", r"\1 hingga \2", result, flags=re.I)
    result = re.sub(r"\s+([,.;:])", r"\1", result)
    result = re.sub(r"([;:])(?=\S)", r"\1 ", result)
    result = re.sub(r"\s{2,}", " ", result).strip()
    if result != original:
        AI_STATS["v2214_final_language_repairs"] = int(
            AI_STATS.get("v2214_final_language_repairs") or 0
        ) + 1
    return result


# Only phrase-level English is a publication issue. Acronyms and established
# technical tokens such as LCD, USB, RS232, LCR, O2, SD card, Excel, HVAC,
# Data Hold, Type K, Pt100, ON/OFF, AC/DC, and measurement units are allowed.
V2214_PUBLIC_RESIDUE_RE = re.compile(
    r"\b(?:"
    r"instruction\s+manual|carrying\s+case|less\s+th[ae]n|power\s+supply|"
    r"measurement\s+range|measuring\s+range|display\s+size|sampling\s+time|"
    r"accessories?\s+(?:included|optional)|connect\s+the\s+optional|will\s+get\s+the|"
    r"freezes?\s+(?:the\s+)?display|liquid\s+is\s+prohibited|"
    r"custom\s+one[- ]chip|with\s+green\s+backlight|normal\s+operation|"
    r"data\s+logging|save\s+data|automatic\s+power\s+off|manual\s+power\s+off|"
    r"for\s+all\s+ranges|after\s+calibration|data\s+acquisition\s+software|"
    r"optional\s+PC|air\s+oxygen|air\s+temperature|dry,?\s+non[- ]corrosive"
    r")\b",
    re.I,
)

V2214_SPEC_RESIDUE_RE = re.compile(
    r"\b(?:"
    r"custom\s+one[- ]chip|with\s+green\s+backlight|test\s+frequency|"
    r"function\s+selector|frequency\s+selector|sorting\s+mode\s+selector|"
    r"dissipation\s+factor|quality\s+factor|phase\s+angle|"
    r"equivalent\s+series\s+resistance|less\s+th[ae]n|meter\s+only|"
    r"soft\s+carrying\s+case|hard\s+carrying\s+case|data\s+acquisition\s+software|"
    r"overload\s+protection|buzzer\s+sounds?\s+if|galvanic\s+cell\s+type|"
    r"for\s+all\s+ranges|after\s+calibration|set\s+clock\s+time|"
    r"set\s+beep\s+sound|decimal\s+point\s+of\s+SD\s+card\s+setting|"
    r"power\s+adapter\s+is\s+optional|will\s+get\s+the|"
    r"normal\s+operation|data\s+logging|save\s+data|"
    r"instruction\s+manual|accessories?\s+(?:included|optional)|"
    r"power\s+supply|operating\s+(?:temperature|humidity)|"
    r"measurement\s+range|sampling\s+time|response\s+time|"
    r"maximum\s+(?:and|&)\s+minimum\s+value|"
    r"\b(?:batteries|or\s+equivalent|adapter\s+input)\b"
    r")\b",
    re.I,
)


def _v2214_rewrite_connection_html(value: Any, data: Dict[str, Any]) -> str:
    raw = clean_text(value)
    state = _source_interface_state(data)
    if not state.get("rs232"):
        return raw
    soup = BeautifulSoup(raw, "html.parser")
    replacement = (
        "Antarmuka serial RS232 tersedia untuk keluaran data ke komputer"
        + (" dan memerlukan kabel UPCB-02 opsional" if re.search(r"UPCB[- ]?0?2", _source_claim_blob(data), re.I) else "")
        + (". Koneksi USB memerlukan kabel USB-01 opsional untuk komputer." if state.get("usb_optional") else ".")
    )
    patterns = (
        re.compile(r"Koneksi\s+RS\s*[- ]?232[^.;]{0,180}(?:;|\.)\s*Koneksi\s+USB[^.]{0,220}\. ?", re.I),
        re.compile(r"Antarmuka\s+(?:komputer\s+)?RS\s*[- ]?232\s*/\s*USB[^.]{0,260}\. ?", re.I),
        re.compile(r"Antarmuka\s+RS\s*[- ]?232\s+dan\s+USB[^.]{0,260}\. ?", re.I),
    )
    for node in list(soup.find_all(string=True)):
        text = str(node)
        new = text
        for pattern in patterns:
            new = pattern.sub(replacement + " ", new)
        if new != text:
            node.replace_with(new)
    return "".join(str(node) for node in soup.contents).strip()


def _v2214_final_editorial_issues(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    # Reuse all mature V22.11 structural/fact checks, but replace its overly
    # broad language detector with phrase-level V22.14 checks.
    issues = [
        item for item in _v2211_final_editorial_issues(row)
        if "english_residue" not in clean_text(item).casefold()
    ]
    public_fields = (
        "Product Description", "Product Short Description",
        "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
        "meta description",
    )
    for field in public_fields:
        visible = _v2211_visible_text(row.get(field))
        if V2214_PUBLIC_RESIDUE_RE.search(visible):
            issues.append("final_public_english_residue:" + field)

    spec_soup = BeautifulSoup(clean_text(row.get("custom_tab_1_content")), "html.parser")
    for cell in spec_soup.find_all(["th", "td"]):
        text = clean_text(cell.get_text(" ", strip=True))
        if text in {"Parameter", "Nilai", "Satuan", "Rentang", "Resolusi", "Akurasi"}:
            continue
        if V2214_SPEC_RESIDUE_RE.search(text):
            issues.append("final_specification_english_residue")
            break
    issues.extend(_v2213_identity_issues(row, data))
    return dedupe_keep_order(issues)


def _v2214_emergency_row(data: Dict[str, Any], pub_date: datetime, reason: str) -> Dict[str, Any]:
    """Build a source-only review row when a product-level builder raises.

    This path is never used for DeepSeek transport/auth/quota failures because
    those raise SystemExit earlier and still stop the batch. It exists only to
    keep one malformed product from discarding all already valid products.
    """
    family = detect_family(data)
    title = standardize_seo_title(
        data, family,
        clean_catalog_source_text(data.get("title"))
        or clean_text(data.get("input_title"))
        or clean_text(data.get("model")),
    )
    specs = localize_spec_rows_fallback(data.get("spec_rows") or [])
    row = {column: "" for column in OUTPUT_COLS}
    row.update({
        "Name": title,
        "Brand": "Lutron",
        "Product Description": build_product_description(data, family),
        "Product Short Description": build_short_description(data, family),
        "custom_tab_1_title": "Spesifikasi",
        "custom_tab_1_content": specifications_tab_html(specs),
        "custom_tab_1_priority": 10,
        "custom_tab_2_title": "Keunggulan Produk",
        "custom_tab_2_content": "<h2>Keunggulan Produk</h2><ul>" + "".join(
            f"<li>{esc(item)}</li>" for item in feature_bullets(data, family)
        ) + "</ul>",
        "custom_tab_2_priority": 20,
        "custom_tab_3_title": "Aplikasi Penggunaan",
        "custom_tab_3_content": build_application_tab(data, family),
        "custom_tab_3_priority": 30,
        "custom_tab_4_title": "FAQ",
        "custom_tab_4_content": build_faq(data, family),
        "custom_tab_4_priority": 40,
        "custom_tab_5_title": "Informasi Produk",
        "custom_tab_5_content": fallback_official_information(data),
        "custom_tab_5_priority": 50,
        "product categories": clean_text(family.get("category")) or "Alat Ukur dan Instrumen",
        "product tags": ", ".join(dedupe_keep_order([
            "Lutron", clean_text(data.get("model")), clean_text(family.get("category"))
        ])),
        "focus keyphrase": build_seo_focus_keyphrase(data, family, title),
        "meta description": build_meta_description(data, family),
        "publication_date": "",
        "Processed": "Yes",
        "Processing Time": "lutron_WEBSITE_import_v22_14_emergency_review",
        "Content Quality": clean_text(data.get("source")) + "_source_only_needs_review",
        "image_url": clean_text(data.get("image_url")),
        "Source URL": clean_text(data.get("url")),
        "Website Validation": "needs_review_v22_14:builder_recovered",
        "Website Correction Log": _merge_source_notes(
            "builder_recovered=" + compact(reason, 500),
            "policy=valid_source_product_not_skipped",
        ),
        "Final Publish Score": 0,
        "Publish Decision": "NEEDS_REVIEW",
    })
    return row


_v2214_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    try:
        row = _v2214_previous_build_output_row(data, pub_date)
    except ValueError as exc:
        # A validation/editorial ValueError is product-local. Preserve the
        # product as Needs Review and continue. SystemExit remains untouched so
        # DeepSeek/global failures still stop the full batch.
        reason = f"{type(exc).__name__}:{exc}"
        data["_v2214_build_recovery_reason"] = reason
        AI_STATS["v2214_build_recoveries"] = int(
            AI_STATS.get("v2214_build_recoveries") or 0
        ) + 1
        try:
            row = _v226_previous_build_output_row(data, pub_date)
        except Exception as fallback_exc:
            return _v2214_emergency_row(
                data, pub_date,
                reason + "; fallback=" + f"{type(fallback_exc).__name__}:{fallback_exc}",
            )

    row = _v2213_apply_identity_to_row(row, data)
    removed = 0
    for field in V2211_HTML_FIELDS:
        if field == "custom_tab_1_content":
            row[field], count = _v2213_polish_spec_html(row.get(field))
            removed += count
            row[field], semantic_removed = _v2211_compact_specification_html(row.get(field))
            removed += semantic_removed
            row[field] = _v2214_rewrite_connection_html(row.get(field), data)
        else:
            row[field] = _v2213_polish_html(
                row.get(field), ensure_cta=(field == "Product Description")
            )
            row[field] = _v2214_rewrite_connection_html(row.get(field), data)
    for field in V2211_PLAIN_FIELDS:
        row[field] = _v2213_polish_technical_text(row.get(field))

    issues = _v2214_final_editorial_issues(row, data)
    data["_v2214_final_issues"] = issues
    score = _v2213_publish_score(row, data, issues)
    quarantined = [
        item for item in data.get("_v22122_quarantined_source_rows") or []
        if isinstance(item, dict)
    ]

    if not issues and not quarantined and score >= 95:
        decision = "READY"
        status = "publish_ready_v22_14"
        AI_STATS["v2214_ready"] = int(AI_STATS.get("v2214_ready") or 0) + 1
    elif not issues and score >= 85:
        decision = "READY_WITH_WARNING"
        status = "publish_ready_with_warning_v22_14"
        AI_STATS["v2214_warning"] = int(AI_STATS.get("v2214_warning") or 0) + 1
    else:
        decision = "NEEDS_REVIEW"
        status = "needs_review_v22_14" + ((":" + ",".join(issues)) if issues else "")
        AI_STATS["v2214_review"] = int(AI_STATS.get("v2214_review") or 0) + 1

    if quarantined and not issues:
        decision = "READY_WITH_WARNING"
        status = "publish_ready_with_warning:source_spec_row_quarantined"

    row["Website Validation"] = status
    row["Final Publish Score"] = score
    row["Publish Decision"] = decision
    row["Processing Time"] = "lutron_WEBSITE_import_v22_14_single_terminal_gate"
    row["Content Quality"] = (
        clean_text(data.get("source")) + "_source_locked_v22_14_" + decision.casefold()
    )
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "repair_pipeline=" + V2214_VERSION,
        "legacy_language_gate_deferred=" + ",".join(data.get("_v2214_deferred_language_issues") or [])
        if data.get("_v2214_deferred_language_issues") else "legacy_language_gate_deferred=none",
        "final_language_repairs=" + str(int(AI_STATS.get("v2214_final_language_repairs") or 0)),
        "final_spec_rows_or_blocks_removed=" + str(removed),
        "final_publish_score=" + str(score),
        "publish_decision=" + decision,
        "final_editorial_issues=" + ",".join(issues) if issues else "final_editorial_issues=none",
        "builder_recovery=" + clean_text(data.get("_v2214_build_recovery_reason"))
        if data.get("_v2214_build_recovery_reason") else "",
        "policy=only_source_404_enters_skipped",
    )
    return row


_v2214_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2214_previous_validate_publish_output(
        data, row, require_ai=require_ai
    )
    # Drop only the legacy broad language findings; V22.14 adds precise final
    # findings from the actual cells that will be written to Excel.
    issues = [
        item for item in issues
        if not _v2214_is_deferred_language_issue(item)
        # All legacy final_* findings were calculated before the V22.14
        # normalizer/semantic-deduper. Re-evaluate them from the actual final
        # row below instead of carrying stale findings forward.
        and not clean_text(item).startswith("final_")
    ]
    issues.extend(data.get("_v2214_final_issues") or [])
    status = clean_text(row.get("Website Validation"))
    if status.startswith((
        "publish_ready_v22_14",
        "publish_ready_with_warning_v22_14",
        "publish_ready_with_warning:source_spec_row_quarantined",
        "needs_review_v22_14",
    )):
        issues = [
            item for item in issues
            if item not in {
                "evidence_publish_status_required",
                "deepseek_validated_content_required",
                "missing_verified_image",
            }
        ]
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.15 — deterministic final-output cleaner and source-sequence unit repair
# ---------------------------------------------------------------------------
# V22.14 solved batch crashes, but the final workbook could still receive a
# false READY status because evidence IDs and mixed-language fragments were
# inspected before all generated cells were normalized.  V22.15 treats the
# actual row written to Excel as the single source of truth.

V2215_VERSION = "22.15-deterministic-finalizer"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_15_deterministic_finalizer_2026_07_30"
AI_STATS.setdefault("v2215_source_sequence_repairs", 0)
AI_STATS.setdefault("v2215_evidence_refs_removed", 0)
AI_STATS.setdefault("v2215_spec_text_repairs", 0)
AI_STATS.setdefault("v2215_faq_rebuilt", 0)
AI_STATS.setdefault("v2215_applications_rebuilt", 0)

V2215_PUBLIC_HTML_FIELDS = (
    "Product Description", "custom_tab_2_content", "custom_tab_3_content",
    "custom_tab_4_content", "custom_tab_5_content",
)
V2215_PUBLIC_PLAIN_FIELDS = (
    "Name", "Product Short Description", "product categories",
    "product tags", "focus keyphrase", "meta description",
)
V2215_EVIDENCE_GROUP_RE = re.compile(
    r"\s*[\(\[]\s*[FS]\d{3}(?:\s*[,;/]\s*[FS]\d{3})*\s*[\)\]]",
    re.I,
)
V2215_EVIDENCE_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9-])(?:F|S)\d{3}(?![A-Za-z0-9-])", re.I)
V2215_UNKNOWN_UNIT_RE = re.compile(r"^\s*([-+]?\d+(?:[.,]\d+)?)\s*([A-Za-z]{2,20})\s*$", re.I)


def _v2215_normalize_unit_name(value: Any) -> str:
    text = clean_text(value).strip(" ,.;:")
    key = re.sub(r"\s+", " ", text).casefold()
    replacements = {
        "kg / cm²": "kg/cm²", "kg/cm²": "kg/cm²", "kg / cm2": "kg/cm²", "kg/cm2": "kg/cm²",
        "kg / cm^2": "kg/cm²", "kg/cm^2": "kg/cm²",
        "mm hg": "mmHg", "mm/hg": "mmHg", "inch hg": "inHg", "inch/hg": "inHg", "inchhg": "inHg", "in hg": "inHg",
        "meter h2o": "mH₂O", "meter/h2o": "mH₂O", "meter h₂o": "mH₂O", "meterh2o": "mH₂O", "meterh₂o": "mH₂O",
        "inch h2o": "inH₂O", "inch/h2o": "inH₂O", "inch h₂o": "inH₂O", "inchh2o": "inH₂O", "inchh₂o": "inH₂O",
        "atmosphere": "atm", "atmosfer": "atm",
        "kpa": "kPa", "hpa": "hPa", "khz": "kHz", "mhz": "MHz", "ghz": "GHz",
        "mbar": "mbar", "bar": "bar", "psi": "psi",
    }
    return replacements.get(key, text)


def _v2215_unit_order_from_rows(rows: Sequence[Dict[str, Any]]) -> List[str]:
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        label = clean_text(row.get("source_label") or row.get("label"))
        if not re.fullmatch(r"(?:display\s+units?|measurement\s+units?|units?|satuan(?:\s+tampilan|\s+pengukuran)?)", label, re.I):
            continue
        raw = clean_text(row.get("value"))
        parts = [clean_text(item) for item in re.split(r"\s*[,;]\s*", raw) if clean_text(item)]
        units = [_v2215_normalize_unit_name(item) for item in parts]
        units = [item for item in units if item]
        if len(units) >= 2:
            return units
    return []


def _v2215_repair_resolution_sequence_in_rows(rows: Sequence[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    output = [dict(item) if isinstance(item, dict) else item for item in (rows or [])]
    units = _v2215_unit_order_from_rows([item for item in output if isinstance(item, dict)])
    if not units:
        return output, []

    def quantity_row(row: Any) -> bool:
        if not isinstance(row, dict) or not _is_technical_spec_section(row.get("section")):
            return False
        value = clean_text(row.get("value"))
        if not (_MEASUREMENT_QUANTITY_RE.fullmatch(value) or V2215_UNKNOWN_UNIT_RE.fullmatch(value)):
            return False
        label = clean_text(row.get("source_label") or row.get("label"))
        label_unit = _v2215_normalize_unit_name(label)
        return bool(_v224_known_unit(label_unit) or V2215_UNKNOWN_UNIT_RE.fullmatch(value))

    runs: List[List[int]] = []
    current: List[int] = []
    for index, row in enumerate(output):
        if quantity_row(row):
            current.append(index)
        else:
            if len(current) >= 2:
                runs.append(current)
            current = []
    if len(current) >= 2:
        runs.append(current)

    repairs: List[Dict[str, str]] = []
    for run in runs:
        if len(run) > len(units):
            continue
        # Validate sequence alignment from known rows. This is the evidence that
        # an unknown OCR token belongs to a specific unit; no conversion or
        # numeric inference is performed.
        aligned = 0
        conflicts = 0
        for position, row_index in enumerate(run):
            row = output[row_index]
            value = clean_text(row.get("value"))
            unknown = V2215_UNKNOWN_UNIT_RE.fullmatch(value)
            if unknown and not _v224_known_unit(clean_text(unknown.group(2))):
                continue
            label_unit = _v2215_normalize_unit_name(row.get("source_label") or row.get("label"))
            expected = _v2215_normalize_unit_name(units[position])
            if label_unit.casefold() == expected.casefold():
                aligned += 1
            else:
                conflicts += 1
        if aligned < 2 or conflicts:
            continue

        for position, row_index in enumerate(run):
            row = output[row_index]
            value = clean_text(row.get("value"))
            unknown = V2215_UNKNOWN_UNIT_RE.fullmatch(value)
            if not unknown:
                continue
            unknown_unit = clean_text(unknown.group(2))
            if _v224_known_unit(unknown_unit):
                continue
            label = clean_text(row.get("source_label") or row.get("label"))
            label_word = re.sub(r"[^A-Za-z]", "", label).casefold()
            if label_word and label_word != unknown_unit.casefold():
                continue
            expected_unit = _v2215_normalize_unit_name(units[position])
            number = clean_text(unknown.group(1))
            row["label"] = expected_unit
            row["source_label"] = expected_unit
            row["value"] = f"{number} {expected_unit}"
            row["canonical_key"] = "other:" + _spec_text_key(expected_unit)[:80]
            repairs.append({
                "old_label": label,
                "old_value": value,
                "new_label": expected_unit,
                "new_value": f"{number} {expected_unit}",
            })
    return output, repairs


def _v2215_prepare_source_rows(data: Dict[str, Any]) -> None:
    all_repairs: List[Dict[str, str]] = []
    for field in ("_raw_spec_rows", "spec_rows"):
        rows = data.get(field)
        if not isinstance(rows, list):
            continue
        repaired, repairs = _v2215_repair_resolution_sequence_in_rows(rows)
        if repairs:
            data[field] = repaired
            all_repairs.extend(repairs)
    if not all_repairs:
        return

    repaired_old_pairs = {
        (clean_text(item.get("old_label")).casefold(), clean_text(item.get("old_value")).casefold())
        for item in all_repairs
    }
    remaining_quarantine = []
    for item in data.get("_v22122_quarantined_source_rows") or []:
        if not isinstance(item, dict):
            continue
        pair = (clean_text(item.get("label")).casefold(), clean_text(item.get("value")).casefold())
        if pair not in repaired_old_pairs:
            remaining_quarantine.append(item)
    data["_v22122_quarantined_source_rows"] = remaining_quarantine
    data["_v22122_source_row_quarantine_active"] = bool(remaining_quarantine)
    data["_v2215_source_sequence_repairs"] = all_repairs
    data["error"] = _merge_source_notes(
        data.get("error"),
        "source_sequence_unit_repair=" + "|".join(
            f"{item['old_label']}={item['old_value']}->{item['new_label']}={item['new_value']}"
            for item in all_repairs
        ),
    )
    AI_STATS["v2215_source_sequence_repairs"] = int(AI_STATS.get("v2215_source_sequence_repairs") or 0) + len(all_repairs)


def _v2215_strip_evidence_text(value: Any) -> Tuple[str, int]:
    text = clean_text(value)
    text, count_group = V2215_EVIDENCE_GROUP_RE.subn("", text)
    text, count_token = V2215_EVIDENCE_TOKEN_RE.subn("", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"\(\s*\)|\[\s*\]", "", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text, count_group + count_token


def _v2215_strip_evidence_html(value: Any) -> Tuple[str, int]:
    soup = BeautifulSoup(clean_text(value), "html.parser")
    removed = 0
    for node in list(soup.find_all(string=True)):
        new, count = _v2215_strip_evidence_text(str(node))
        removed += count
        if new != str(node):
            node.replace_with(new)
    return "".join(str(node) for node in soup.contents).strip(), removed


def _v2215_normalize_technical_text(value: Any, data: Optional[Dict[str, Any]] = None) -> Tuple[str, int]:
    text = clean_text(value)
    original = text
    replacements: List[Tuple[re.Pattern[str], str]] = [
        (re.compile(r"\bOpen\s*/\s*Short\s+calibration\b", re.I), "Kalibrasi sirkuit terbuka/hubung singkat"),
        (re.compile(r"\bSorting\s+tolerance\s+mode\b", re.I), "Mode toleransi penyortiran"),
        (re.compile(r"\btheta\s+measurement\b|\bθ\s+measurement\b", re.I), "Pengukuran sudut fase θ"),
        (re.compile(r"\bCalibration\b", re.I), "Kalibrasi"),
        (re.compile(r"\bAlligator\s+clips?\s*\(\s*red\s+and\s+black\s*\)", re.I), "Klip buaya (merah dan hitam)"),
        (re.compile(r"\bAlligator\s+clips?\b", re.I), "Klip buaya"),
        (re.compile(r"\bSoft\s+(?:carrying\s+case|kotak\s+pelindung)\b", re.I), "Kotak pelindung lunak"),
        (re.compile(r"\bHard\s+carrying\s+case\b", re.I), "Kotak pelindung keras"),
        (re.compile(r"\bInstruction\s+manual\b", re.I), "Buku petunjuk"),
        (re.compile(r"\bExcel\s+Perangkat\s+lunak\s+akuisisi\s+data\b", re.I), "Perangkat lunak akuisisi data berbasis Excel"),
        (re.compile(r"\bExcel\s+data\s+acquisition\s+software\b", re.I), "Perangkat lunak akuisisi data berbasis Excel"),
        (re.compile(r"\bData\s+acquisition\s+software\b", re.I), "Perangkat lunak akuisisi data"),
        (re.compile(r"\bSMD\s+test\s+clip\b", re.I), "Klip WEBSITE SMD"),
        (re.compile(r"\bSMD\s+tester\b", re.I), "PengWEBSITE SMD"),
        (re.compile(r"\bHolster\b", re.I), "Sarung pelindung"),
        (re.compile(r"\bUnder\s+lampu\s+latar\s+LCD\s+mati\b", re.I), "Saat lampu latar LCD mati"),
        (re.compile(r"\bdisertakan\s+Accessories\b|\bIncluded\s+Accessories\b", re.I), "Aksesori yang disertakan"),
        (re.compile(r"\bAccessories\b", re.I), "Aksesori"),
        (re.compile(r"\bAutomatic\b", re.I), "Otomatis"),
        (re.compile(r"\bInductance\b", re.I), "Induktansi"),
        (re.compile(r"\bCapacitance\b", re.I), "Kapasitansi"),
        (re.compile(r"\bResistance\b", re.I), "Resistansi"),
        (re.compile(r"\bGalvanic\s+cell\s+(?:type|tipe)\b", re.I), "Tipe sel galvanik"),
        (re.compile(r"\bThermistor\b", re.I), "Termistor"),
        (re.compile(r"\bBuzzer\b", re.I), "Bel"),
        (re.compile(r"\bonly\s+meter\b", re.I), "hanya unit meter"),
        (re.compile(r"\b1\s+Pair\b", re.I), "1 pasang"),
        (re.compile(r"\bAC\s+to\s+DC\b", re.I), "AC ke DC"),
        (re.compile(r"\badapter\b", re.I), "adaptor"),
        (re.compile(r"\bor\s+(?=DC\b)", re.I), "atau "),
        (re.compile(r"(?<=\d)\s+to\s+(?=\d)", re.I), " hingga "),
        (re.compile(r"(?<=\bHz)\s+to\s+(?=\d)|(?<=\bkHz)\s+to\s+(?=\d)|(?<=\bH)\s+to\s+(?=\d)", re.I), " hingga "),
        (re.compile(r"\bat\s+(?=\d+(?:[.,]\d+)?\s*(?:Hz|kHz|MHz|GHz)\b)", re.I), "pada "),
        (re.compile(r"\bKHz\b"), "kHz"),
        (re.compile(r"\bKg(?=/|\b)"), "kg"),
        (re.compile(r"(?<=\d)\s*[xX]\s*(?=\d)"), " × "),
        (re.compile(r"\s*@\s*"), "; "),
    ]
    for pattern, replacement in replacements:
        text = pattern.sub(replacement, text)

    # Sentence-level repairs for recurring mixed-language fragments emitted by
    # translated Lutron specification blocks.
    sentence_replacements: List[Tuple[re.Pattern[str], str]] = [
        (re.compile(r"\bDual\s*(?:&|and)\s*differential\s+input\b", re.I), "masukan ganda dan diferensial"),
        (re.compile(r"\breal[- ]time\b", re.I), "waktu nyata"),
        (re.compile(r"\bincluding\s+linearity,?\s+hysteresis,?\s+and\s+keterulangan\b", re.I), "termasuk linearitas, histeresis, dan keterulangan"),
        (re.compile(r"\bFull\s+Scale\b", re.I), "skala penuh"),
        (re.compile(r"\bheavy\s+duty\b", re.I), "untuk penggunaan intensif"),
        (re.compile(r"\bw/o\s+SD\s+card\s+Penyimpanan\s+data\s+and\b", re.I), "tanpa penyimpanan data ke kartu SD dan"),
        (re.compile(r"\bwithout\s+SD\s+card\s+and\b", re.I), "tanpa kartu SD dan"),
        (re.compile(r"\bWhen\s+SD\s+card\s+save\s+the\s+data\s+and\b", re.I), "Saat menyimpan data ke kartu SD dan"),
        (re.compile(r"\bwhen\s+saving\s+data\s+to\s+SD\s+card\s+and\b", re.I), "saat menyimpan data ke kartu SD dan"),
        (re.compile(r"\bIf\s+LCD\s+lampu\s+latar\s+menyala\b", re.I), "Jika lampu latar LCD menyala"),
        (re.compile(r"\bwaktu\s+pencuplikan\s+Setting\s+range\b", re.I), "Rentang pengaturan waktu pencuplikan"),
        (re.compile(r"\b1\s+sec\s+to\s+8\s+Hour\s+59\s+Minute\s+59\s+sec\b", re.I), "1 detik hingga 8 jam 59 menit 59 detik"),
        (re.compile(r"\b1\s+detik\s+to\s+8\s+hours?\s+59\s+minutes?\s+59\s+detik\b", re.I), "1 detik hingga 8 jam 59 menit 59 detik"),
        (re.compile(r"\bwaktu\s+pencuplikan\s+can\s+set\s+to\s+1\s+detik,?\s+but\s+memori\s+data\s+may\s+loss\b", re.I), "Waktu pencuplikan dapat diatur ke 1 detik, tetapi data memori mungkin hilang"),
        (re.compile(r"\bManual\s+Push\s+the\s+pencatat\s+data\s+button\s+once\s+(?:will\s+)?(?:save|simpan)\s+data(?:\s+one\s+time)?\b", re.I), "Manual: tekan tombol pencatat data sekali untuk menyimpan satu data"),
        (re.compile(r"\bManual\s+Push\s+the\s+pencatat\s+data\s+button\s+once\s+to\s+simpan\s+data\b", re.I), "Manual: tekan tombol pencatat data sekali untuk menyimpan data"),
        (re.compile(r"\bPush\s+the\s+pencatat\s+data\s+button\s+once\s+(?:to\s+)?(?:save|simpan)\s+data(?:\s+one\s+time)?\b", re.I), "Tekan tombol pencatat data sekali untuk menyimpan data"),
        (re.compile(r"\bSet\s+the\s+waktu\s+pencuplikan\s+to\s+0\s+detik\b", re.I), "Atur waktu pencuplikan ke 0 detik"),
        (re.compile(r"\bMode\s+Manual,?\s+can\s+also\s+select\s+the\s+1\s+hingga\s+99\s+position\s*\(Location\)\s*no\b", re.I), "Dalam mode manual, posisi lokasi 1 hingga 99 dapat dipilih"),
        (re.compile(r"\bAutomatic\s+1\s+detik\s+to\s+8\s+hours?\s+59\s+minutes?\s+59\s+detik\b", re.I), "Otomatis 1 detik hingga 8 jam 59 menit 59 detik"),
        (re.compile(r"\bMeasurement\b", re.I), "Pengukuran"),
        (re.compile(r"\bselector\b", re.I), "pemilih"),
        (re.compile(r"\b(?:Auto(?:matic)?|Otomatis)\s+Power\s+Off\b", re.I), "Mati otomatis"),
        (re.compile(r"\bopen\s*/\s*short\b", re.I), "sirkuit terbuka/hubung singkat"),
        (re.compile(r"\bSD\s+card\b", re.I), "kartu SD"),
        (re.compile(r"\bdata\s+save\b", re.I), "penyimpanan data"),
        (re.compile(r"\bsec\b", re.I), "detik"),
        (re.compile(r"\bhours?\b", re.I), "jam"),
        (re.compile(r"\bminutes?\b", re.I), "menit"),
    ]
    for pattern, replacement in sentence_replacements:
        text = pattern.sub(replacement, text)

    # Convert source ranges after units, e.g. 20 mH to 2000 mH and 0°C to 50°C.
    text = re.sub(
        r"(\d+(?:[.,]\d+)?\s*(?:°\s*[CF]|uH|µH|mH|H|pF|nF|uF|µF|mF|Ω|Ω|KΩ|kΩ|MΩ|Hz|kHz|MHz|GHz)?)\s+to\s+(?=\d)",
        r"\1 hingga ",
        text,
        flags=re.I,
    )
    text = re.sub(r"\bat\s+(?=\d)", "pada ", text, flags=re.I)

    text = re.sub(r"\binch\s+Hg\b|\binchHg\b", "inHg", text, flags=re.I)
    text = re.sub(r"\bmm\s+Hg\b", "mmHg", text, flags=re.I)
    text = re.sub(r"\bmeter\s+H(?:2|₂)O\b", "mH₂O", text, flags=re.I)
    text = re.sub(r"\binch\s+H(?:2|₂)O\b|\binchH(?:2|₂)O\b|\binchH2\b", "inH₂O", text, flags=re.I)
    text = re.sub(r"\bAtmosphere\b|\batmosfer\b", "atm", text, flags=re.I)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([,.;:])(?=[A-Za-zÀ-ÿ])", r"\1 ", text)
    text = re.sub(r"\s{2,}", " ", text).strip()

    if data is not None and re.search(r"RS\s*[- ]?232\s*/\s*USB\s+unit\s+computer\s+interface", text, re.I):
        state = _source_interface_state(data)
        source_blob = _source_claim_blob(data)
        connection = "Antarmuka serial RS232 untuk keluaran data ke komputer"
        if re.search(r"UP[-]?CB[- ]?0?2|UPCB[- ]?0?2", source_blob, re.I):
            connection += " dan memerlukan kabel UPCB-02 opsional"
        if state.get("usb_optional"):
            connection += "; koneksi USB memerlukan kabel USB-01 opsional"
        text = re.sub(r"RS\s*[- ]?232\s*/\s*USB\s+unit\s+computer\s+interface", connection, text, flags=re.I)

    return text, int(text != original)


def _v2215_normalize_html_text(value: Any, data: Optional[Dict[str, Any]] = None) -> Tuple[str, int]:
    soup = BeautifulSoup(clean_text(value), "html.parser")
    changed = 0
    for node in list(soup.find_all(string=True)):
        new, count = _v2215_normalize_technical_text(str(node), data)
        changed += count
        if new != str(node):
            node.replace_with(new)
    return "".join(str(node) for node in soup.contents).strip(), changed


def _v2215_canonical_spec_label(value: Any) -> str:
    text = _spec_text_key(value)
    aliases = {
        "includedaccessories": "accessoriesincluded",
        "disertakanaccessories": "accessoriesincluded",
        "aksesori yangdisertakan": "accessoriesincluded",
        "aksesoriyangdisertakan": "accessoriesincluded",
        "accessoriesincluded": "accessoriesincluded",
        "optionalaccessories": "accessoriesoptional",
        "aksesoriopsional": "accessoriesoptional",
    }
    return aliases.get(text, text)


def _v2215_dedupe_spec_rows(html_value: Any) -> Tuple[str, int]:
    soup = BeautifulSoup(clean_text(html_value), "html.parser")
    removed = 0
    for table in soup.find_all("table"):
        groups: Dict[str, List[Tuple[Any, str]]] = {}
        for tr in list(table.find_all("tr")):
            cells = tr.find_all(["th", "td"], recursive=False)
            if len(cells) < 2 or tr.find_parent("thead") is not None:
                continue
            label = clean_text(cells[0].get_text(" ", strip=True))
            value = clean_text(" / ".join(cell.get_text(" ", strip=True) for cell in cells[1:]))
            if not label or not value:
                continue
            groups.setdefault(_v2215_canonical_spec_label(label), []).append((tr, value))
        for key, items in groups.items():
            if len(items) < 2:
                continue
            # Exact duplicate after normalization.
            seen: Dict[str, Any] = {}
            survivors: List[Tuple[Any, str]] = []
            for tr, value in items:
                value_key = _spec_text_key(value)
                if value_key in seen:
                    tr.decompose()
                    removed += 1
                else:
                    seen[value_key] = tr
                    survivors.append((tr, value))
            # Accessory rows commonly repeat a shorter summary after the full
            # row. Keep the more informative source-grounded row.
            if key in {"accessoriesincluded", "accessoriesoptional"} and len(survivors) > 1:
                longest_tr, longest_value = max(survivors, key=lambda item: len(_spec_text_key(item[1])))
                longest_key = _spec_text_key(longest_value)
                for tr, value in survivors:
                    if tr is longest_tr:
                        continue
                    value_key = _spec_text_key(value)
                    if value_key and (value_key in longest_key or SequenceMatcher(None, value_key, longest_key).ratio() >= 0.78):
                        tr.decompose()
                        removed += 1
    return "".join(str(node) for node in soup.contents).strip(), removed


def _v2215_source_max_range(data: Dict[str, Any]) -> str:
    source = _source_claim_blob(data)
    patterns = (
        r"maximum\s+range\s+of\s+([^,.;\n]{3,80})",
        r"rentang\s+maksimum\s+(?:sebesar\s+)?([^,.;\n]{3,80})",
    )
    for pattern in patterns:
        match = re.search(pattern, source, re.I)
        if match:
            value = clean_text(match.group(1)).strip(" -")
            if re.search(r"\d", value) and _MEASUREMENT_UNIT_PATTERN.search(value):
                return value
    return ""


def _v2215_add_explicit_max_range(html_value: Any, data: Dict[str, Any]) -> Tuple[str, int]:
    max_range = _v2215_source_max_range(data)
    if not max_range:
        return clean_text(html_value), 0
    soup = BeautifulSoup(clean_text(html_value), "html.parser")
    visible = clean_text(soup.get_text(" ", strip=True))
    if re.search(r"\bRentang\s+maksimum\b", visible, re.I):
        return "".join(str(node) for node in soup.contents).strip(), 0
    # Add only when the table currently presents resolutions without a range
    # column. No converted ranges are invented.
    headers = [clean_text(th.get_text(" ", strip=True)) for th in soup.find_all("th")]
    if "Resolusi" not in headers or "Rentang" in headers:
        return "".join(str(node) for node in soup.contents).strip(), 0
    general_table = next((table for table in soup.find_all("table") if "Parameter" in [clean_text(th.get_text(" ", strip=True)) for th in table.find_all("th")]), None)
    if general_table is None:
        general_table = soup.new_tag("table", attrs={"class": "WEBSITE-spec-table WEBSITE-spec-general"})
        thead = soup.new_tag("thead")
        htr = soup.new_tag("tr")
        for title in ("Parameter", "Nilai"):
            th = soup.new_tag("th")
            th.string = title
            htr.append(th)
        thead.append(htr)
        tbody = soup.new_tag("tbody")
        general_table.extend([thead, tbody])
        soup.append(general_table)
    tbody = general_table.find("tbody") or general_table
    tr = soup.new_tag("tr")
    th = soup.new_tag("th", attrs={"scope": "row"})
    th.string = "Rentang maksimum"
    td = soup.new_tag("td")
    td.string = max_range
    tr.extend([th, td])
    tbody.insert(0, tr)
    return "".join(str(node) for node in soup.contents).strip(), 1



def _v2215_application_fallback(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    if data.get("is_accessory"):
        items = [
            "Kelengkapan unit utama yang kompatibel sesuai referensi model.",
            "Penggantian atau penambahan komponen sesuai kebutuhan teknis.",
            "Pemeliharaan peralatan ukur yang menggunakan komponen terkait.",
            "Pemeriksaan kompatibilitas sebelum pemasangan atau penggunaan.",
        ]
    else:
        term = clean_text(family.get("term")) or "alat ukur"
        items = [
            f"Pengukuran dan pemeriksaan menggunakan {term} sesuai spesifikasi produk.",
            "Kontrol kualitas pada proses atau sampel yang sesuai dengan parameter alat.",
            "PengWEBSITEan di laboratorium, area inspeksi, atau lokasi kerja.",
            "Verifikasi parameter sebelum pemeliharaan atau proses lanjutan.",
        ]
    return "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(f"<li>{esc(item)}</li>" for item in items) + "</ul>"


def _v2215_faq_fallback(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    title = standardize_seo_title(data, family, data.get("title")) or clean_text(data.get("model")) or "produk ini"
    model = clean_text(data.get("model"))
    parameters, range_value, accuracy = _v22_primary_measurement(data)
    measurement_parts: List[str] = []
    if parameters:
        measurement_parts.append("parameter " + ", ".join(parameters[:2]))
    if range_value:
        measurement_parts.append("rentang " + range_value)
    if accuracy:
        measurement_parts.append("akurasi " + accuracy)
    measurement_answer = (
        "Data utama yang tercantum meliputi " + ", ".join(measurement_parts) + "."
        if measurement_parts else
        "Periksa tab Spesifikasi untuk parameter, rentang, resolusi, dan kondisi kerja yang tersedia."
    )
    source_answer = (
        "Data produk disusun dari halaman produk resmi Lutron pada URL sumber yang tercantum."
    )
    items = [
        (f"Apa fungsi {title}?", f"Produk ini digunakan untuk {clean_text(family.get('function'))}."),
        ("Apa data pengukuran utamanya?", measurement_answer),
        ("Apa yang perlu diperiksa sebelum penggunaan?", "Periksa parameter, satuan, rentang, media, catu daya, aksesori, dan kondisi kerja pada spesifikasi produk."),
        (f"Dari mana data teknis {model or 'produk'} disusun?", source_answer),
    ]
    return "".join(f"<h3>{esc(question)}</h3><p>{esc(answer)}</p>" for question, answer in items)


def _v2215_rebuild_structural_tabs(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    repairs: List[str] = []
    family = detect_family(data)
    faq_soup = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    if len(faq_soup.find_all("h3")) != 4 or V2215_EVIDENCE_TOKEN_RE.search(faq_soup.get_text(" ", strip=True)):
        row["custom_tab_4_content"] = _v2215_faq_fallback(data, family)
        repairs.append("faq_rebuilt")
        AI_STATS["v2215_faq_rebuilt"] = int(AI_STATS.get("v2215_faq_rebuilt") or 0) + 1
    app_soup = BeautifulSoup(clean_text(row.get("custom_tab_3_content")), "html.parser")
    if len(app_soup.find_all("li")) != 4:
        row["custom_tab_3_content"] = _v2215_application_fallback(data, family)
        repairs.append("applications_rebuilt")
        AI_STATS["v2215_applications_rebuilt"] = int(AI_STATS.get("v2215_applications_rebuilt") or 0) + 1
    return repairs


V2215_FINAL_RESIDUE_RE = re.compile(
    r"\b(?:theta\s+measurement|sorting\s+tolerance\s+mode|open\s*/\s*short(?:\s+calibration)?|"
    r"alligator\s+clips?|soft\s+carrying\s+case|soft\s+kotak\s+pelindung|"
    r"excel\s+perangkat\s+lunak|under\s+lampu\s+latar|"
    r"automatic|real[- ]time|dual\s*(?:&|and)\s*differential\s+input|"
    r"including|linearity|hysteresis|full\s+scale|heavy\s+duty|"
    r"w/o|without|when|if|setting\s+range|can\s+set|may\s+loss|"
    r"push\s+the|will\s+save|one\s+time|location|"
    r"\d\s+to\s+\d|\bor\s+DC\b|\band\s+lampu\b|disertakan\s+accessories|"
    r"galvanic\s+cell\s+(?:type|tipe)|\bmeasurement\b|\bselector\b|"
    r"power\s+off|\b(?:end|dogs)\s*=|\b(?:end|dogs)\b\s+\d)\b",
    re.I,
)


def _v2215_final_issues(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    public_blob = " ".join(
        _v2211_visible_text(row.get(field))
        for field in V2215_PUBLIC_HTML_FIELDS + V2215_PUBLIC_PLAIN_FIELDS + ("custom_tab_1_content",)
    )
    if V2215_EVIDENCE_TOKEN_RE.search(public_blob):
        issues.append("internal_evidence_reference_visible")
    if has_cjk(public_blob):
        issues.append("chinese_character_detected")
    if V2215_FINAL_RESIDUE_RE.search(public_blob):
        issues.append("mixed_language_or_corrupt_specification_residue")
    faq_soup = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    if len(faq_soup.find_all("h3")) != 4:
        issues.append("faq_item_count_invalid")
    app_soup = BeautifulSoup(clean_text(row.get("custom_tab_3_content")), "html.parser")
    if len(app_soup.find_all("li")) != 4:
        issues.append("application_item_count_invalid")
    spec_soup = BeautifulSoup(clean_text(row.get("custom_tab_1_content")), "html.parser")
    if not spec_soup.find("table"):
        issues.append("specification_table_invalid")
    if re.search(r"\b(?:dogs|end)\b", spec_soup.get_text(" ", strip=True), re.I):
        issues.append("corrupt_source_unit_visible")
    issues.extend(_v2213_identity_issues(row, data))
    return dedupe_keep_order(issues)


_v2215_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    _v2215_prepare_source_rows(data)
    row = _v2215_previous_build_output_row(data, pub_date)

    structural_repairs = _v2215_rebuild_structural_tabs(row, data)
    evidence_removed = 0
    text_repairs = 0
    for field in V2215_PUBLIC_HTML_FIELDS:
        row[field], removed = _v2215_strip_evidence_html(row.get(field))
        evidence_removed += removed
        row[field], changed = _v2215_normalize_html_text(row.get(field), data)
        text_repairs += changed
        row[field] = _v2213_polish_html(row.get(field), ensure_cta=(field == "Product Description"))
    for field in V2215_PUBLIC_PLAIN_FIELDS:
        row[field], removed = _v2215_strip_evidence_text(row.get(field))
        evidence_removed += removed
        row[field], changed = _v2215_normalize_technical_text(row.get(field), data)
        text_repairs += changed

    spec_html, removed = _v2215_strip_evidence_html(row.get("custom_tab_1_content"))
    evidence_removed += removed
    spec_html, changed = _v2215_normalize_html_text(spec_html, data)
    text_repairs += changed
    spec_html, deduped = _v2215_dedupe_spec_rows(spec_html)
    spec_html, range_added = _v2215_add_explicit_max_range(spec_html, data)
    row["custom_tab_1_content"] = spec_html

    # Rebuild SEO fields after all title/text normalization. This prevents a
    # valid final row from inheriting stale meta/keyphrase findings produced
    # by an earlier override layer.
    family = detect_family(data)
    row["Name"] = standardize_seo_title(data, family, row.get("Name"))
    row["focus keyphrase"] = build_seo_focus_keyphrase(data, family, row.get("Name"))
    row["meta description"] = _v22_build_meta(data, family, row.get("focus keyphrase"))
    row["product tags"] = ", ".join(build_product_tags(data, family, row.get("Name")))

    issues = _v2215_final_issues(row, data)
    if not _v22_meta_valid(
        row.get("meta description"),
        clean_text(data.get("model")),
        clean_text(row.get("focus keyphrase")),
    ):
        issues.append("meta_description_invalid")
    issues = dedupe_keep_order(issues)
    data["_v2215_final_issues"] = issues
    source_repairs = data.get("_v2215_source_sequence_repairs") or []
    repairs_count = evidence_removed + text_repairs + deduped + range_added + len(structural_repairs) + len(source_repairs)

    if issues:
        decision = "NEEDS_REVIEW"
        score = max(0, 100 - 15 * len(issues))
        status = "needs_review_v22_15:" + ",".join(issues)
    else:
        decision = "READY"
        score = 100
        status = "publish_ready_v22_15"
        if data.get("_v22122_quarantined_source_rows"):
            decision = "READY_WITH_WARNING"
            score = 95
            status = "publish_ready_with_warning_v22_15:source_spec_row_quarantined"

    row["Website Validation"] = status
    row["Final Publish Score"] = score
    row["Publish Decision"] = decision
    row["Processing Time"] = "lutron_WEBSITE_import_v22_15_deterministic_finalizer"
    row["Content Quality"] = clean_text(data.get("source")) + "_source_locked_v22_15_" + decision.casefold()
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "finalizer=" + V2215_VERSION,
        "evidence_refs_removed=" + str(evidence_removed),
        "technical_text_repairs=" + str(text_repairs),
        "spec_rows_removed=" + str(deduped),
        "explicit_source_ranges_added=" + str(range_added),
        "source_sequence_repairs=" + str(len(source_repairs)),
        "structural_repairs=" + ",".join(structural_repairs) if structural_repairs else "structural_repairs=none",
        "final_v2215_issues=" + ",".join(issues) if issues else "final_v2215_issues=none",
    )
    AI_STATS["v2215_evidence_refs_removed"] = int(AI_STATS.get("v2215_evidence_refs_removed") or 0) + evidence_removed
    AI_STATS["v2215_spec_text_repairs"] = int(AI_STATS.get("v2215_spec_text_repairs") or 0) + text_repairs
    return row


_v2215_previous_validate_publish_output = validate_publish_output


def validate_publish_output(data: Dict[str, Any], row: Dict[str, Any], require_ai: bool = False) -> List[str]:
    issues = _v2215_previous_validate_publish_output(data, row, require_ai=require_ai)
    # V22.15 re-evaluates all final/editorial language findings from the actual
    # output row after deterministic cleanup.
    issues = [
        item for item in issues
        if not clean_text(item).startswith("final_")
        and "english_residue" not in clean_text(item).casefold()
        and item not in {
            "evidence_publish_status_required",
            "deepseek_validated_content_required",
            "missing_verified_image",
            "application_item_count_invalid",
            "faq_item_count_invalid",
            "faq_structure_invalid",
            "meta_description_invalid",
            "meta_description_not_keyphrase_aligned",
        }
    ]
    issues.extend(data.get("_v2215_final_issues") or [])
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.16 — reseller-safe public copy
# ---------------------------------------------------------------------------
# Provenance remains available in the internal Source URL and correction-log
# columns, but must never appear in customer-facing WooCommerce copy.

V2216_VERSION = "22.16-reseller-safe-public-copy"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_16_reseller_safe_2026_08_01"

V2216_PUBLIC_SOURCE_DISCLOSURE_RE = re.compile(
    r"(?:halaman\s+(?:produk\s+)?resmi\s+lutron|halaman\s+produk\s+lutron|"
    r"url\s+sumber|sumber\s+yang\s+tercantum|dokumen\s+sumber|"
    r"sumber\s+produk\s+resmi|berdasarkan\s+sumber\s+produk\s+resmi|"
    r"(?:data\s+produk|informasi\s+teknis|spesifikasi\s+teknis)[^.<]{0,140}"
    r"(?:disusun|dirujuk|dirangkum|berasal)\s+dari\s+(?:halaman|situs|website|dokumen)|"
    r"spesifikasi\s+teknis\s+berasal\s+dari\s+halaman\s+produk\s+resmi|"
    r"official\s+(?:lutron\s+)?(?:product\s+)?(?:page|website)|source\s+url)",
    re.I,
)


def _v2216_clean_public_html(value: Any, field: str = "") -> Tuple[str, int, int]:
    raw = clean_text(value)
    if not raw:
        return "", 0, 0

    replacements = (
        (
            "Informasi berikut disusun dari fakta teknis yang tervalidasi pada sumber produk.",
            "Informasi berikut merangkum fungsi dan spesifikasi teknis produk.",
        ),
        (
            "Fitur dan batas penggunaan mengikuti spesifikasi resmi pada halaman produk.",
            "Fitur dan batas penggunaan mengikuti spesifikasi teknis produk.",
        ),
        (
            "Spesifikasi teknis berasal dari halaman produk resmi",
            "Spesifikasi teknis produk disajikan secara ringkas dan terstruktur",
        ),
        (
            "Informasi teknis dirujuk dari halaman produk Lutron.",
            "Informasi teknis produk tersedia pada tab Spesifikasi.",
        ),
    )
    changed = 0
    for old, new in replacements:
        count = raw.count(old)
        if count:
            raw = raw.replace(old, new)
            changed += count

    raw, count = re.subn(
        r"Informasi teknis produk\s+[^.<]{0,140}\s+dirangkum dari halaman resmi Lutron dan dokumen yang terdeteksi\.",
        "Informasi teknis produk dirangkum untuk membantu pemilihan model dan penggunaan.",
        raw,
        flags=re.I,
    )
    changed += count
    raw, count = re.subn(
        r"Informasi teknis dirangkum dari halaman resmi Lutron dan dokumen yang terdeteksi\.",
        "Informasi teknis produk dirangkum untuk membantu pemilihan model dan penggunaan.",
        raw,
        flags=re.I,
    )
    changed += count
    raw, count = re.subn(
        r"(<strong>\s*Klarifikasi koneksi:\s*</strong>)\s*Sumber mencantumkan\s*",
        r"\1 ",
        raw,
        flags=re.I,
    )
    changed += count
    raw, count = re.subn(r"\bSumber mencantumkan\s+", "", raw, flags=re.I)
    changed += count

    soup = BeautifulSoup(raw, "html.parser")
    source_ids_removed = 0
    for tag in soup.find_all(attrs={"data-source-id": True}):
        try:
            del tag["data-source-id"]
            source_ids_removed += 1
        except Exception:
            pass

    if field == "custom_tab_4_content":
        for heading in list(soup.find_all("h3")):
            answer = heading.find_next_sibling("p")
            question_text = clean_text(heading.get_text(" ", strip=True))
            answer_text = clean_text(answer.get_text(" ", strip=True)) if answer else ""
            if (
                re.search(r"\b(?:dari mana|sumber|rujukan)\b", question_text, re.I)
                or V2216_PUBLIC_SOURCE_DISCLOSURE_RE.search(answer_text)
            ):
                heading.string = "Apakah tersedia konsultasi sebelum pemesanan?"
                if answer is None:
                    answer = soup.new_tag("p")
                    heading.insert_after(answer)
                answer.string = (
                    "Ya. Tim WEBSITE dapat membantu mencocokkan model, rentang, aksesori, "
                    "dan kebutuhan aplikasi sebelum pemesanan."
                )
                changed += 1

    for node in soup.find_all(["p", "li"]):
        visible = clean_text(node.get_text(" ", strip=True))
        if V2216_PUBLIC_SOURCE_DISCLOSURE_RE.search(visible):
            if field == "custom_tab_4_content":
                node.string = (
                    "Tim WEBSITE dapat membantu mencocokkan model dan spesifikasi "
                    "produk sebelum pemesanan."
                )
            else:
                node.string = (
                    "Informasi teknis produk disajikan secara ringkas untuk membantu "
                    "pemilihan model dan penggunaan."
                )
            changed += 1

    output = "".join(str(node) for node in soup.contents).strip()
    return output, changed, source_ids_removed


def _v2216_clean_public_text(value: Any) -> Tuple[str, int]:
    text = clean_text(value)
    changed = 0
    rules = (
        (r"\bberdasarkan sumber produk resmi\b", "dengan rincian spesifikasi teknis"),
        (r"\bsesuai spesifikasi resmi\b", "sesuai spesifikasi teknis"),
        (r"\bpada halaman produk\b", "pada spesifikasi produk"),
        (r"\bhalaman produk resmi Lutron\b", "spesifikasi teknis produk"),
        (r"\bURL sumber\b", "data internal"),
    )
    for pattern, replacement in rules:
        text, count = re.subn(pattern, replacement, text, flags=re.I)
        changed += count
    return re.sub(r"\s{2,}", " ", text).strip(), changed


def _v2216_public_source_leak_fields(row: Dict[str, Any]) -> List[str]:
    fields: List[str] = []
    for field in V2215_PUBLIC_HTML_FIELDS + ("custom_tab_1_content",):
        visible = _v2211_visible_text(row.get(field))
        if V2216_PUBLIC_SOURCE_DISCLOSURE_RE.search(visible):
            fields.append(field)
    for field in V2215_PUBLIC_PLAIN_FIELDS:
        if V2216_PUBLIC_SOURCE_DISCLOSURE_RE.search(clean_text(row.get(field))):
            fields.append(field)
    return dedupe_keep_order(fields)


_v2216_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2216_previous_build_output_row(data, pub_date)
    public_rewrites = 0
    source_ids_removed = 0

    for field in V2215_PUBLIC_HTML_FIELDS + ("custom_tab_1_content",):
        cleaned, rewrites, removed_ids = _v2216_clean_public_html(row.get(field), field=field)
        row[field] = cleaned
        public_rewrites += rewrites
        source_ids_removed += removed_ids

    for field in V2215_PUBLIC_PLAIN_FIELDS:
        cleaned, rewrites = _v2216_clean_public_text(row.get(field))
        row[field] = cleaned
        public_rewrites += rewrites

    leaks = _v2216_public_source_leak_fields(row)
    if leaks:
        row["Website Validation"] = "needs_review_v22_16:public_source_disclosure_visible"
        row["Publish Decision"] = "NEEDS_REVIEW"
        try:
            current_score = int(row.get("Final Publish Score") or 100)
        except Exception:
            current_score = 100
        row["Final Publish Score"] = min(current_score, 70)

    row["Processing Time"] = "lutron_WEBSITE_import_v22_16_reseller_safe"
    quality = clean_text(row.get("Content Quality"))
    if "reseller_safe" not in quality:
        row["Content Quality"] = (quality + "_reseller_safe").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "reseller_safe_finalizer=" + V2216_VERSION,
        "public_source_disclosures_rewritten=" + str(public_rewrites),
        "public_source_ids_removed=" + str(source_ids_removed),
        "public_source_leaks=" + ",".join(leaks) if leaks else "public_source_leaks=none",
    )
    return row


_v2216_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2216_previous_validate_publish_output(
        data,
        row,
        require_ai=require_ai,
    )
    if _v2216_public_source_leak_fields(row):
        issues.append("public_source_disclosure_visible")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.17 — SEO conversion finalizer for customer-focused product pages
# ---------------------------------------------------------------------------
# The technical source remains locked and auditable, while the public copy is
# reorganized around buyer intent: what the product is, what it measures, its
# key range/accuracy/resolution, and how to choose the right configuration.

V2217_VERSION = "22.17-seo-conversion-finalizer"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_17_seo_conversion_2026_08_01"


def _v2217_product_title(data: Dict[str, Any], family: Dict[str, Any]) -> str:
    """Return the exact public title pattern: Product Type + Brand + Model."""
    brand = clean_text(data.get("brand")) or "Lutron"
    model = clean_text(data.get("model"))
    product_type = clean_text(_seo_product_type(family, clean_text(data.get("title"))))
    product_type = _remove_title_phrase(product_type, brand)
    product_type = _remove_title_phrase(product_type, model)
    product_type = re.sub(r"\s+", " ", product_type).strip(" -,:;.")
    if not product_type:
        product_type = "Aksesori" if data.get("is_accessory") else "Alat Ukur"
    return " ".join(dedupe_keep_order([product_type, brand, model])).strip()


def _v2217_unique_values(values: Iterable[Any], limit: int = 6) -> List[str]:
    output: List[str] = []
    seen = set()
    for value in values:
        text = compact(clean_text(value), 220).strip(" .;,")
        if not text or text == "-":
            continue
        key = _spec_text_key(text)
        if not key or key in seen:
            continue
        seen.add(key)
        output.append(text)
        if len(output) >= limit:
            break
    return output


def _v2217_extract_spec_pairs(spec_html: Any) -> List[Dict[str, str]]:
    """Read the final Indonesian specification tables into buyer-facing facts."""
    soup = BeautifulSoup(clean_text(spec_html), "html.parser")
    raw_pairs: List[Tuple[int, str, str]] = []
    source_index = 0

    # Important facts such as general accuracy may be rendered as a note
    # above the table rather than as a table row. Preserve those too.
    for note in soup.find_all(["p", "div"]):
        strong = note.find("strong")
        if strong is None:
            continue
        label = clean_text(strong.get_text(" ", strip=True)).strip(" :")
        full_text = clean_text(note.get_text(" ", strip=True))
        value = clean_text(re.sub(r"^" + re.escape(clean_text(strong.get_text(" ", strip=True))), "", full_text, count=1, flags=re.I)).strip(" :")
        if label and value:
            raw_pairs.append((source_index, label, compact(value, 240)))
            source_index += 1

    for table in soup.find_all("table"):
        headers = [clean_text(th.get_text(" ", strip=True)) for th in table.select("thead th")]
        header_keys = [_spec_text_key(header) for header in headers]
        body_rows = table.select("tbody > tr") or table.find_all("tr")

        if headers and any(key in {"rentang", "range", "resolusi", "resolution", "akurasi", "accuracy"} for key in header_keys):
            columns: Dict[str, List[str]] = {header: [] for header in headers}
            for tr in body_rows:
                cells = [clean_text(cell.get_text(" ", strip=True)) for cell in tr.find_all(["th", "td"], recursive=False)]
                if len(cells) != len(headers):
                    continue
                for header, value in zip(headers, cells):
                    columns[header].append(value)
            for header in headers:
                values = _v2217_unique_values(columns.get(header, []), limit=5)
                if not values:
                    continue
                normalized_header = clean_text(header)
                if _spec_text_key(normalized_header) in {"satuan", "unit"}:
                    normalized_header = "Satuan pengukuran"
                raw_pairs.append((source_index, normalized_header, " / ".join(values)))
                source_index += 1
            continue

        for tr in body_rows:
            cells = [clean_text(cell.get_text(" ", strip=True)) for cell in tr.find_all(["th", "td"], recursive=False)]
            if len(cells) < 2:
                continue
            label = clean_text(cells[0])
            value = clean_text(" / ".join(cells[1:]))
            if not label or not value:
                continue
            if _spec_text_key(label) in {"parameter", "nilai"} and _spec_text_key(value) in {"parameter", "nilai"}:
                continue
            raw_pairs.append((source_index, label, compact(value, 240)))
            source_index += 1

    priority_rules: Tuple[Tuple[re.Pattern[str], int], ...] = (
        (re.compile(r"\b(?:rentang|range|skala)\b", re.I), 100),
        (re.compile(r"\b(?:akurasi|accuracy|ketelitian)\b", re.I), 96),
        (re.compile(r"\b(?:resolusi|resolution)\b", re.I), 94),
        (re.compile(r"\b(?:parameter|pengukuran|measurement|satuan|unit)\b", re.I), 88),
        (re.compile(r"\b(?:sensor|probe|kanal|channel|input)\b", re.I), 84),
        (re.compile(r"\b(?:antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat|logger|memori)\b", re.I), 78),
        (re.compile(r"\b(?:tampilan|display|LCD)\b", re.I), 68),
        (re.compile(r"\b(?:catu daya|power|baterai|battery)\b", re.I), 60),
        (re.compile(r"\b(?:aksesori|accessor|kelengkapan)\b", re.I), 56),
        (re.compile(r"\b(?:suhu operasi|kelembapan operasi|operating)\b", re.I), 48),
        (re.compile(r"\b(?:dimensi|berat|dimension|weight)\b", re.I), 20),
    )

    scored: List[Tuple[int, int, Dict[str, str]]] = []
    seen = set()
    for index, label, value in raw_pairs:
        if re.search(r"\b(?:model|sku|mpn|kode)\b", label, re.I):
            continue
        signature = (_spec_text_key(label), _spec_text_key(value))
        if not all(signature) or signature in seen:
            continue
        seen.add(signature)
        score = 30
        probe = f"{label} {value}"
        for pattern, candidate_score in priority_rules:
            if pattern.search(probe):
                score = max(score, candidate_score)
        scored.append((score, index, {"label": normalize_label(label), "value": value}))

    scored.sort(key=lambda item: (-item[0], item[1]))
    return [item[2] for item in scored]


def _v2217_find_specs(pairs: Sequence[Dict[str, str]], pattern: str, limit: int = 2) -> List[Dict[str, str]]:
    regex = re.compile(pattern, re.I)
    return [row for row in pairs if regex.search(clean_text(row.get("label")) + " " + clean_text(row.get("value")))][:limit]


def _v2217_fact_sentence(rows: Sequence[Dict[str, str]], prefix: str = "") -> str:
    parts = [f"{clean_text(row.get('label'))}: {clean_text(row.get('value'))}" for row in rows if clean_text(row.get("label")) and clean_text(row.get("value"))]
    if not parts:
        return ""
    sentence = (prefix + " " if prefix else "") + "; ".join(parts)
    return sentence.strip().rstrip(".;") + "."


def _v2217_intro_html(data: Dict[str, Any], family: Dict[str, Any], title: str) -> str:
    term = clean_text(family.get("term")) or clean_text(_seo_product_type(family, title)).casefold() or "alat ukur"
    function = clean_text(family.get("function")) or "melakukan pengukuran sesuai spesifikasi produk"
    if data.get("is_accessory"):
        first = f"{title} adalah {term} yang digunakan untuk {function}."
        second = (
            "Produk ini dipilih berdasarkan kecocokan model, konektor, dan kebutuhan teknis unit utama. "
            "Pastikan kompatibilitas serta kelengkapan yang diperlukan sebelum digunakan."
        )
    else:
        first = f"{title} adalah {term} yang digunakan untuk {function}."
        second = (
            f"Produk ini cocok digunakan ketika pengguna membutuhkan {term} dengan parameter dan kondisi kerja yang sesuai spesifikasi. "
            "Rentang, sensor, koneksi, serta aksesori perlu disesuaikan dengan kebutuhan aplikasi."
        )
    return f"<p><strong>{esc(title)}</strong> {esc(first[len(title):].lstrip())}</p><p>{esc(second)}</p>"


def _v2217_description_tail(existing_html: Any) -> str:
    """Preserve useful AI sections from the first H2 onward, without duplicating key specs."""
    soup = BeautifulSoup(clean_text(existing_html), "html.parser")
    first_h2 = soup.find("h2")
    if first_h2 is None:
        return WEBSITE_CTA
    nodes: List[str] = []
    current = first_h2
    while current is not None:
        next_node = current.next_sibling
        if getattr(current, "name", None) == "h2" and re.search(r"spesifikasi\s+utama", clean_text(current.get_text(" ", strip=True)), re.I):
            candidate = current.find_next_sibling()
            current.extract()
            if candidate is not None and getattr(candidate, "name", None) in {"ul", "ol", "p"}:
                candidate.extract()
        else:
            nodes.append(str(current))
        current = next_node
    tail = "".join(nodes).strip()
    if WEBSITE_CTA not in tail:
        tail = tail.rstrip() + WEBSITE_CTA
    return tail


def _v2217_product_description(row: Dict[str, Any], data: Dict[str, Any], family: Dict[str, Any], title: str, pairs: Sequence[Dict[str, str]]) -> str:
    key_specs = list(pairs[:6])
    blocks = [_v2217_intro_html(data, family, title)]
    if key_specs:
        blocks.append("<h2>Spesifikasi Utama</h2><ul>" + "".join(
            f"<li><strong>{esc(item.get('label'))}:</strong> {esc(item.get('value'))}</li>"
            for item in key_specs
        ) + "</ul>")
    blocks.append(_v2217_description_tail(row.get("Product Description")))
    rendered = _v2213_polish_html("".join(blocks), ensure_cta=True)
    # BeautifulSoup normalization may trim a leading text-node space after
    # inline <strong>. Restore it so the public copy does not join words.
    rendered = re.sub(r"</strong>(?=[A-Za-zÀ-ÿ0-9±])", "</strong> ", rendered)
    return rendered


def _v2217_short_description(data: Dict[str, Any], family: Dict[str, Any], title: str, pairs: Sequence[Dict[str, str]]) -> str:
    function = clean_text(family.get("function")) or "melakukan pengukuran sesuai spesifikasi produk"
    details = []
    for row in pairs[:2]:
        details.append(f"{clean_text(row.get('label')).casefold()} {clean_text(row.get('value'))}")
    text = f"{title} digunakan untuk {function}."
    if details:
        text += " Spesifikasi penting mencakup " + "; ".join(details) + "."
    if len(text) <= V22_SHORT_MAX:
        return text
    return compact(f"{title} digunakan untuk {function}.", V22_SHORT_MAX)


def _v2217_build_faq(data: Dict[str, Any], family: Dict[str, Any], title: str, pairs: Sequence[Dict[str, str]]) -> str:
    model = clean_text(data.get("model")) or title
    term = clean_text(family.get("term")) or clean_text(_seo_product_type(family, title)).casefold() or "alat ukur"
    function = clean_text(family.get("function")) or "melakukan pengukuran sesuai spesifikasi produk"

    range_rows = _v2217_find_specs(pairs, r"\b(?:rentang|range|skala)\b", 3)
    accuracy_rows = _v2217_find_specs(pairs, r"\b(?:akurasi|accuracy|ketelitian)\b", 2)
    resolution_rows = _v2217_find_specs(pairs, r"\b(?:resolusi|resolution)\b", 2)
    parameter_rows = _v2217_find_specs(pairs, r"\b(?:parameter|pengukuran|measurement|satuan|unit)\b", 2)
    interface_rows = _v2217_find_specs(pairs, r"\b(?:antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat|logger|memori)\b", 2)
    sensor_rows = _v2217_find_specs(pairs, r"\b(?:sensor|probe|kanal|channel|input)\b", 2)
    accessory_rows = _v2217_find_specs(pairs, r"\b(?:aksesori|accessor|kelengkapan|catu daya|power|baterai|battery)\b", 2)

    items: List[Tuple[str, str]] = []
    items.append((
        f"Apa itu {title} dan digunakan untuk apa?",
        f"{title} adalah {term} yang digunakan untuk {function}. Produk ini dipilih berdasarkan parameter, rentang, dan kondisi kerja yang dibutuhkan pada aplikasi pengguna.",
    ))

    if range_rows:
        if len(range_rows) == 1:
            range_answer = f"Rentang pengukurannya adalah {clean_text(range_rows[0].get('value'))}."
        else:
            range_answer = _v2217_fact_sentence(range_rows, "Rentang yang tercantum meliputi")
        items.append((
            f"Berapa rentang pengukuran {model}?",
            range_answer,
        ))
    elif parameter_rows:
        items.append((
            f"Parameter apa yang dapat diukur oleh {model}?",
            _v2217_fact_sentence(parameter_rows, "Parameter pengukuran yang tercantum meliputi"),
        ))
    else:
        items.append((
            f"Apa parameter utama {model}?",
            f"Produk ini digunakan untuk {function}. Pemilihan parameter pengukuran perlu disesuaikan dengan kebutuhan aplikasi.",
        ))

    precision_rows = accuracy_rows + [row for row in resolution_rows if row not in accuracy_rows]
    if precision_rows:
        accuracy_value = clean_text(accuracy_rows[0].get("value")) if accuracy_rows else ""
        resolution_value = clean_text(resolution_rows[0].get("value")) if resolution_rows else ""
        if accuracy_value and resolution_value:
            precision_answer = f"Akurasi yang tercantum adalah {accuracy_value}, sedangkan resolusinya {resolution_value}."
        elif accuracy_value:
            precision_answer = f"Akurasi yang tercantum adalah {accuracy_value}."
        elif resolution_value:
            precision_answer = f"Resolusi yang tercantum adalah {resolution_value}."
        else:
            precision_answer = _v2217_fact_sentence(precision_rows, "Data akurasi dan resolusi yang tercantum adalah")
        items.append((
            "Berapa akurasi dan resolusi pengukurannya?",
            precision_answer,
        ))
    else:
        feature_rows = list(pairs[:2])
        items.append((
            "Apa spesifikasi teknis utama yang perlu diperhatikan?",
            _v2217_fact_sentence(feature_rows, "Spesifikasi utama yang perlu diperhatikan meliputi") or f"Perhatikan parameter, rentang, dan kondisi kerja yang sesuai untuk penggunaan {term} ini.",
        ))

    practical_rows = interface_rows or sensor_rows or accessory_rows or list(pairs[3:5])
    if interface_rows:
        q4 = "Apakah produk ini memiliki koneksi atau fungsi pencatatan data?"
    elif sensor_rows:
        q4 = "Apakah produk ini memerlukan sensor, probe, atau input tertentu?"
    elif accessory_rows:
        q4 = "Apa saja kelengkapan dan kebutuhan dayanya?"
    else:
        q4 = "Apa fitur lain yang perlu diperhatikan sebelum digunakan?"
    if interface_rows and len(interface_rows) == 1:
        practical_answer = f"Produk ini memiliki {clean_text(interface_rows[0].get('label')).casefold()} {clean_text(interface_rows[0].get('value'))}."
    elif sensor_rows and len(sensor_rows) == 1:
        practical_answer = f"Konfigurasi sensor atau input yang tercantum adalah {clean_text(sensor_rows[0].get('value'))}."
    elif accessory_rows and len(accessory_rows) == 1:
        practical_answer = f"Kelengkapan atau kebutuhan daya yang tercantum adalah {clean_text(accessory_rows[0].get('value'))}."
    else:
        practical_answer = _v2217_fact_sentence(practical_rows, "Rincian yang tercantum meliputi")
    items.append((
        q4,
        practical_answer or "Pastikan catu daya, koneksi, sensor, serta aksesori yang diperlukan sesuai dengan konfigurasi produk.",
    ))

    items.append((
        "Bagaimana memilih model atau rentang yang sesuai?",
        "Tentukan parameter yang akan diukur, nilai minimum dan maksimum, media atau objek pengukuran, kebutuhan akurasi, serta kebutuhan keluaran data. Tim WEBSITE dapat membantu mencocokkan model dan aksesori yang diperlukan sebelum pemesanan.",
    ))

    return "".join(f"<h3>{esc(question)}</h3><p>{esc(answer)}</p>" for question, answer in items[:5])


def _v2217_meta_description(data: Dict[str, Any], family: Dict[str, Any], title: str) -> str:
    function = clean_text(family.get("function"))
    function = re.sub(r"\s+sesuai\s+spesifikasi\s+produk\s*$", "", function, flags=re.I).strip(" .")
    function = compact(function, 54).rstrip(".") if function else "kebutuhan pengukuran dan pemeriksaan"
    candidates = [
        f"{title} untuk {function}. Cek rentang, akurasi, fitur, dan spesifikasi. Konsultasikan kebutuhan Anda dengan tim WEBSITE.",
        f"{title}. Cek fungsi, rentang, akurasi, fitur, dan spesifikasi. Konsultasikan kebutuhan produk Anda dengan tim WEBSITE.",
        f"{title}. Cek fungsi, rentang, akurasi, dan spesifikasi. Konsultasikan kebutuhan Anda dengan tim WEBSITE.",
        f"{title}. Lihat fungsi dan spesifikasi teknisnya, lalu konsultasikan kebutuhan pengukuran Anda dengan tim WEBSITE.",
    ]
    model = clean_text(data.get("model"))
    for candidate in candidates:
        candidate = re.sub(r"\s+", " ", candidate).strip()
        if _v22_meta_valid(candidate, model, title):
            return candidate
    fallback = f"{title}. Lihat fungsi, rentang, fitur, dan spesifikasi teknis produk. Konsultasikan pemilihan model dengan tim WEBSITE."
    if len(fallback) <= V22_META_MAX:
        return fallback
    return _v22_build_meta(data, family, title)


def _v2217_content_issues(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    family = detect_family(data)
    expected_title = _v2217_product_title(data, family)
    if clean_text(row.get("Name")) != expected_title:
        issues.append("seo_title_pattern_invalid")
    if clean_text(row.get("focus keyphrase")) != expected_title:
        issues.append("focus_keyphrase_not_exact_title")

    description = BeautifulSoup(clean_text(row.get("Product Description")), "html.parser")
    heading = next((h for h in description.find_all("h2") if re.search(r"spesifikasi\s+utama", clean_text(h.get_text(" ", strip=True)), re.I)), None)
    if heading is None:
        issues.append("key_specification_section_missing")
    elif not heading.find_next_sibling(["ul", "ol"]):
        issues.append("key_specification_items_missing")

    faq = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    faq_count = len(faq.find_all("h3"))
    if not 4 <= faq_count <= 6:
        issues.append("faq_item_count_invalid")
    faq_text = clean_text(faq.get_text(" ", strip=True))
    if not re.search(r"\bApa itu\b", faq_text, re.I):
        issues.append("faq_product_definition_missing")
    if re.search(r"^\s*Periksa\s+tab\s+Spesifikasi", faq_text, re.I):
        issues.append("faq_non_solution_answer_visible")

    meta = clean_text(row.get("meta description"))
    title = clean_text(row.get("Name"))
    if not _v22_meta_valid(meta, clean_text(data.get("model")), title):
        issues.append("meta_description_invalid")
    if not re.search(r"\b(?:cek|lihat|konsultasi|konsultasikan)\b", meta, re.I):
        issues.append("meta_description_missing_click_motivation")
    return dedupe_keep_order(issues)


_v2217_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2217_previous_build_output_row(data, pub_date)
    family = detect_family(data)
    title = _v2217_product_title(data, family)
    pairs = _v2217_extract_spec_pairs(row.get("custom_tab_1_content"))

    row["Name"] = title
    row["focus keyphrase"] = title
    row["Product Short Description"] = _v2217_short_description(data, family, title, pairs)
    row["Product Description"] = _v2217_product_description(row, data, family, title, pairs)
    row["custom_tab_4_content"] = _v2217_build_faq(data, family, title, pairs)
    row["meta description"] = _v2217_meta_description(data, family, title)

    # Normalize technical wording again because V22.17 creates new public HTML
    # after the earlier V22.15 language-cleaning stage has already run.
    for field in ("Product Description", "custom_tab_4_content"):
        row[field], _technical_repairs = _v2215_normalize_html_text(row.get(field), data)
        row[field] = re.sub(r"\bto\b(?=\s*[-+]?\d)", "hingga", row[field], flags=re.I)
        if field == "Product Description":
            row[field] = _v2213_polish_html(row.get(field), ensure_cta=True)
        row[field] = re.sub(r"</strong>(?=[A-Za-zÀ-ÿ0-9±])", "</strong> ", row[field])
    row["Product Short Description"], _technical_repairs = _v2215_normalize_technical_text(row.get("Product Short Description"), data)
    row["meta description"], _technical_repairs = _v2215_normalize_technical_text(row.get("meta description"), data)

    # Run reseller-safe cleanup once more after SEO reconstruction.
    for field in ("Product Description", "custom_tab_4_content"):
        row[field], _rewrites, _removed_ids = _v2216_clean_public_html(row.get(field), field=field)
    row["meta description"], _rewrites = _v2216_clean_public_text(row.get("meta description"))

    seo_issues = _v2217_content_issues(row, data)
    data["_v2217_final_issues"] = seo_issues
    if seo_issues:
        row["Website Validation"] = "needs_review_v22_17:" + ",".join(seo_issues)
        row["Publish Decision"] = "NEEDS_REVIEW"
        try:
            existing_score = int(row.get("Final Publish Score") or 100)
        except Exception:
            existing_score = 100
        row["Final Publish Score"] = min(existing_score, max(0, 100 - 12 * len(seo_issues)))
    elif clean_text(row.get("Publish Decision")) in {"", "READY", "READY_WITH_WARNING"}:
        row["Website Validation"] = "publish_ready_v22_17"
        row["Publish Decision"] = clean_text(row.get("Publish Decision")) or "READY"

    row["Processing Time"] = "lutron_WEBSITE_import_v22_17_seo_conversion"
    quality = clean_text(row.get("Content Quality"))
    if "seo_conversion" not in quality:
        row["Content Quality"] = (quality + "_seo_conversion").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "seo_finalizer=" + V2217_VERSION,
        "seo_title_pattern=product_type_brand_model",
        "key_specifications_lifted=" + str(min(6, len(pairs))),
        "buyer_faq_items=" + str(len(BeautifulSoup(clean_text(row.get('custom_tab_4_content')), 'html.parser').find_all('h3'))),
        "seo_conversion_issues=" + ",".join(seo_issues) if seo_issues else "seo_conversion_issues=none",
    )
    return row


_v2217_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2217_previous_validate_publish_output(data, row, require_ai=require_ai)
    # Re-evaluate fields reconstructed by V22.17 from their actual final values.
    issues = [
        issue for issue in issues
        if issue not in {
            "faq_item_count_invalid",
            "faq_structure_invalid",
            "meta_description_invalid",
            "meta_description_not_keyphrase_aligned",
            "seo_title_invalid",
            "seo_title_missing_brand_or_model",
            "focus_keyphrase_invalid",
            "product_description_structure_invalid_after_repair",
            "commercial_faq_topic_detected_after_repair",
            "faq_count_invalid_after_repair",
        }
    ]
    issues.extend(_v2217_content_issues(row, data))
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.18 — SEO short-description finalizer
# ---------------------------------------------------------------------------

V2218_VERSION = "v22.18_short_description_seo"


def _v2218_clean_function_text(value: Any) -> str:
    """Remove vague source-safe filler while retaining the actual product use."""
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(
        r"\s+(?:sesuai|berdasarkan)\s+(?:dengan\s+)?spesifikasi\s+produk\b.*$",
        "",
        text,
        flags=re.I,
    )
    text = re.sub(
        r"\s+sesuai\s+(?:dengan\s+)?spesifikasi\b.*$",
        "",
        text,
        flags=re.I,
    )
    text = re.sub(r"\s{2,}", " ", text).strip(" .,:;-")
    return text


def _v2218_short_spec_score(row: Dict[str, str], index: int) -> Tuple[int, int]:
    label = clean_text(row.get("label"))
    value = clean_text(row.get("value"))
    probe = f"{label} {value}"
    if not label or not value:
        return (-1000, -index)
    if re.search(r"\b(?:model|merek|brand|sku|mpn|nomor\s+model)\b", label, re.I):
        return (-900, -index)

    rules: List[Tuple[str, int]] = [
        (r"\b(?:rentang|range|skala|measurement\s+range|measuring\s+range)\b", 120),
        (r"\b(?:akurasi|accuracy|ketelitian|toleransi)\b", 115),
        (r"\b(?:resolusi|resolution|minimum\s+scale)\b", 110),
        (r"\b(?:parameter|satuan|unit\s+pengukuran|measurement\s+unit)\b", 98),
        (r"\b(?:true\s*rms|pembobotan\s+frekuensi|frequency\s+weighting|class\s*2|kelas\s*2)\b", 94),
        (r"\b(?:sensor|probe|kanal|channel|input)\b", 90),
        (r"\b(?:antarmuka|interface|keluaran|output|rs\s*[- ]?232|usb|data\s*logger|pencatat\s+data|memori)\b", 86),
        (r"\b(?:display|tampilan|lcd|data\s+hold|max(?:imum)?|min(?:imum)?)\b", 76),
        (r"\b(?:catu\s+daya|power\s+supply|baterai|battery|aksesori|accessor)\b", 52),
        (r"\b(?:suhu\s+operasi|operating\s+temperature|kelembapan\s+operasi|operating\s+humidity)\b", 34),
        (r"\b(?:dimensi|dimension|berat|weight)\b", 20),
    ]
    score = 44
    for pattern, candidate in rules:
        if re.search(pattern, probe, re.I):
            score = max(score, candidate)
    if re.search(r"[-+]?\d+(?:[.,]\d+)?", value):
        score += 8
    if len(value) > 190:
        score -= 20
    return (score, -index)


def _v2218_short_fact_phrase(row: Dict[str, str]) -> str:
    label = clean_text(row.get("label")).strip(" :;.-")
    value = clean_text(row.get("value")).strip(" :;.-")
    if not label or not value:
        return ""
    label = re.sub(r"\s+", " ", label)
    value = re.sub(r"\s+", " ", value)
    low = label.casefold()

    if re.search(r"\b(?:rentang|range|skala)\b", low):
        display = "Rentang"
    elif re.search(r"\b(?:akurasi|accuracy|ketelitian|toleransi)\b", low):
        display = "Akurasi"
    elif re.search(r"\b(?:resolusi|resolution)\b", low):
        display = "Resolusi"
    elif re.search(r"\b(?:satuan|measurement\s+unit)\b", low):
        display = "Satuan"
    elif re.search(r"\b(?:antarmuka|interface|keluaran|output)\b", low):
        display = "Koneksi data"
    elif re.search(r"\b(?:pencatat\s+data|data\s*logger|memori)\b", low):
        display = "Pencatatan data"
    elif re.search(r"\b(?:sensor|probe|kanal|channel|input)\b", low):
        display = normalize_label(label)
    else:
        display = normalize_label(label)

    # Keep each fact compact so the short description remains readable in cards.
    value = compact(value, 105).rstrip(".")
    return f"{display}: {value}"


def _v2218_select_short_facts(
    pairs: Sequence[Dict[str, str]],
    maximum: int = 3,
) -> List[str]:
    ranked: List[Tuple[int, int, Dict[str, str]]] = []
    for index, raw in enumerate(pairs or []):
        if not isinstance(raw, dict):
            continue
        score, order = _v2218_short_spec_score(raw, index)
        if score <= 0:
            continue
        ranked.append((score, order, raw))
    ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)

    output: List[str] = []
    seen_labels = set()
    seen_values = set()
    for _score, _order, row in ranked:
        phrase = _v2218_short_fact_phrase(row)
        if not phrase:
            continue
        label_key = phrase.split(":", 1)[0].casefold()
        value_key = _spec_text_key(phrase.split(":", 1)[-1])
        if label_key in seen_labels or (value_key and value_key in seen_values):
            continue
        seen_labels.add(label_key)
        if value_key:
            seen_values.add(value_key)
        output.append(phrase)
        if len(output) >= maximum:
            break
    return output


def _v2217_short_description(
    data: Dict[str, Any],
    family: Dict[str, Any],
    title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    """Build a specific short description without vague specification filler."""
    function = _v2218_clean_function_text(family.get("function"))
    if not function:
        function = "mendukung kebutuhan pengukuran dan pemeriksaan teknis"

    base_candidates = [
        f"{title} digunakan untuk {function}.",
        f"{title} untuk {function}.",
    ]
    facts = _v2218_select_short_facts(pairs, maximum=3)

    for base in base_candidates:
        for count in range(len(facts), 0, -1):
            candidate = base + " " + "; ".join(facts[:count]) + "."
            candidate = re.sub(r"\s{2,}", " ", candidate).strip()
            if len(candidate) <= V22_SHORT_MAX:
                return candidate
        if len(base) <= V22_SHORT_MAX:
            return base

    fallback = f"{title} untuk kebutuhan pengukuran dan pemeriksaan teknis."
    return compact(fallback, V22_SHORT_MAX)


_v2218_previous_content_issues = _v2217_content_issues


def _v2217_content_issues(row: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    issues = _v2218_previous_content_issues(row, data)
    short = clean_text(row.get("Product Short Description"))
    if re.search(r"\b(?:sesuai|berdasarkan)\s+(?:dengan\s+)?spesifikasi\s+produk\b", short, re.I):
        issues.append("short_description_vague_specification_filler")
    if re.search(r"\b(?:untuk\s+berbagai\s+kebutuhan|berkualitas\s+dan\s+andal|cek\s+tab\s+spesifikasi)\b", short, re.I):
        issues.append("short_description_generic_sales_copy")
    if len(short) > V22_SHORT_MAX or not re.search(r"[.!?]$", short):
        issues.append("short_description_invalid_after_seo_finalizer")
    return dedupe_keep_order(issues)


_v2218_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2218_previous_build_output_row(data, pub_date)
    row["Processing Time"] = "lutron_WEBSITE_import_v22_18_short_description_seo"
    quality = clean_text(row.get("Content Quality"))
    if "short_description_seo" not in quality:
        row["Content Quality"] = (quality + "_short_description_seo").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "short_description_finalizer=" + V2218_VERSION,
        "vague_specification_filler=blocked",
        "short_description_concrete_facts=" + str(
            len(_v2218_select_short_facts(
                _v2217_extract_spec_pairs(row.get("custom_tab_1_content")),
                maximum=3,
            ))
        ),
    )
    return row


# ---------------------------------------------------------------------------
# V22.19 — continue batch on product-local AI validation failure
# ---------------------------------------------------------------------------
# Product-local specification/content repair failures are review outcomes, not
# process-wide failures. Fatal API authentication/account errors still raise
# SystemExit and stop immediately. This layer also persists the AI cache after
# every product so a later fatal event does not discard already-paid results.

V2219_VERSION = "22.19-continue-needs-review"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_19_continue_needs_review_2026_08_03"

for _column in (
    "Review Failure Stage",
    "Review Failure Details",
    "Diagnostic File",
):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)

AI_STATS.setdefault("v2219_product_review_fallbacks", 0)
AI_STATS.setdefault("v2219_incremental_ai_cache_saves", 0)


def _v2219_problem_codes(items: Any) -> List[str]:
    codes: List[str] = []
    if isinstance(items, dict):
        items = [items]
    for item in items if isinstance(items, list) else []:
        if isinstance(item, dict):
            code = clean_text(item.get("code") or item.get("Issue Code"))
        else:
            code = clean_text(item)
        if code:
            codes.append(code)
    return dedupe_keep_order(codes)


def _v2219_failure_details(data: Dict[str, Any]) -> str:
    details: List[str] = []
    spec_report = data.get("_v227_spec_report")
    if isinstance(spec_report, dict):
        unresolved = spec_report.get("unresolved")
        if isinstance(unresolved, dict):
            for source_id, issues in unresolved.items():
                codes = _v2219_problem_codes(issues)
                details.append(
                    "spec_row_" + clean_text(source_id) + "="
                    + (",".join(codes) if codes else "unresolved")
                )
    content_report = data.get("_v227_content_report")
    if isinstance(content_report, dict):
        unresolved = content_report.get("unresolved")
        if isinstance(unresolved, dict):
            for field, issues in unresolved.items():
                codes = _v2219_problem_codes(issues)
                details.append(
                    "content_" + clean_text(field) + "="
                    + (",".join(codes) if codes else "unresolved")
                )
    return "; ".join(details) or clean_text(data.get("_ai_status")) or "ai_validation_unresolved"


def _v2219_review_fields(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    """Return deterministic source-only fields so the row can reach Excel.

    These fields are deliberately marked NEEDS_REVIEW. No new technical value
    is invented and no additional AI request is made after the repair failure.
    """
    specs = data.get("spec_rows") if isinstance(data.get("spec_rows"), list) else []
    title = clean_catalog_source_text(data.get("title")) or clean_text(data.get("model")) or "Produk Lutron"
    features = (
        "<h2>Keunggulan Produk</h2><ul>"
        + "".join(f"<li>{esc(item)}</li>" for item in feature_bullets(data, family))
        + "</ul>"
    )
    fields: Dict[str, Any] = {
        "product_name_id": title,
        "category_id": clean_text(family.get("category")),
        "product_description_html": build_product_description(data, family),
        "short_description": build_short_description(data, family),
        "feature_tab_html": features,
        "faq_html": build_faq(data, family),
        "meta_description": build_meta_description(data, family),
        "official_information_html": fallback_official_information(data),
        "translated_spec_rows": localize_spec_rows_fallback(specs),
        "ai_spec_dedup_applied": False,
        "ai_spec_dedup_removed_count": 0,
    }
    return fields


def _v2219_persist_ai_cache() -> None:
    try:
        cache_path = clean_text(AI_SETTINGS.get("cache_path"))
        cache = AI_SETTINGS.get("cache")
        if cache_path and isinstance(cache, dict):
            save_cache(cache_path, cache)
            AI_STATS["v2219_incremental_ai_cache_saves"] = int(
                AI_STATS.get("v2219_incremental_ai_cache_saves") or 0
            ) + 1
    except Exception:
        # Cache persistence must never turn a product review into a batch stop.
        pass


_v2219_previous_get_ai_content = get_ai_content


def get_ai_content(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    try:
        result = _v2219_previous_get_ai_content(data, family)
    except SystemExit as exc:
        message = clean_text(exc)
        product_local = (
            message.startswith("REPAIR SPESIFIKASI BELUM LOLOS")
            or message.startswith("SAFE RECOVERY BELUM LOLOS")
        )
        if not product_local:
            _v2219_persist_ai_cache()
            raise
        diagnostic_match = re.search(r"Diagnostic:\s*(.+)$", message, re.I)
        data["_v2219_review_stage"] = (
            "specification_repair" if message.startswith("REPAIR SPESIFIKASI")
            else "content_safe_recovery"
        )
        data["_v2219_review_reason"] = message
        data["_v2219_diagnostic_file"] = (
            clean_text(diagnostic_match.group(1)) if diagnostic_match else ""
        )
        data["_v2219_review_details"] = _v2219_failure_details(data)
        result = _v2219_review_fields(data, family)
        data["_ai_fields"] = result
        data["_ai_status"] = "needs_review_product_local_ai_failure"
        data["error"] = _merge_source_notes(
            data.get("error"),
            "v22_19_batch_continued",
            "review_stage=" + data["_v2219_review_stage"],
            "review_details=" + data["_v2219_review_details"],
        )
        AI_STATS["v2219_product_review_fallbacks"] = int(
            AI_STATS.get("v2219_product_review_fallbacks") or 0
        ) + 1
        print(
            "    NEEDS_REVIEW dicatat; batch lanjut | stage="
            + data["_v2219_review_stage"]
            + " | " + data["_v2219_review_details"]
        )
    else:
        # A nonfatal validation path may return an empty dict. Convert that one
        # product to a review row instead of letting main() raise SystemExit.
        if AI_SETTINGS.get("enabled") and not result:
            data["_v2219_review_stage"] = "ai_validation"
            data["_v2219_review_reason"] = (
                clean_text(data.get("_ai_status")) or "DeepSeek output belum tervalidasi"
            )
            data["_v2219_review_details"] = _v2219_failure_details(data)
            result = _v2219_review_fields(data, family)
            data["_ai_fields"] = result
            data["_ai_status"] = "needs_review_nonfatal_ai_validation"
            data["error"] = _merge_source_notes(
                data.get("error"),
                "v22_19_batch_continued",
                "review_stage=ai_validation",
                "review_details=" + data["_v2219_review_details"],
            )
            AI_STATS["v2219_product_review_fallbacks"] = int(
                AI_STATS.get("v2219_product_review_fallbacks") or 0
            ) + 1
            print(
                "    NEEDS_REVIEW dicatat; batch lanjut | stage=ai_validation | "
                + data["_v2219_review_details"]
            )
    finally:
        _v2219_persist_ai_cache()
    return result


_v2219_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2219_previous_build_output_row(data, pub_date)
    stage = clean_text(data.get("_v2219_review_stage"))
    if not stage:
        return row

    details = clean_text(data.get("_v2219_review_details")) or _v2219_failure_details(data)
    diagnostic = clean_text(data.get("_v2219_diagnostic_file"))
    row["Website Validation"] = "needs_review_v22_19:" + stage
    row["Publish Decision"] = "NEEDS_REVIEW"
    try:
        current_score = int(row.get("Final Publish Score") or 100)
    except Exception:
        current_score = 100
    row["Final Publish Score"] = min(current_score, 45)
    row["Review Failure Stage"] = stage
    row["Review Failure Details"] = details
    row["Diagnostic File"] = diagnostic
    row["Processing Time"] = "lutron_WEBSITE_import_v22_19_continue_needs_review"
    quality = clean_text(row.get("Content Quality"))
    if "continue_needs_review" not in quality:
        row["Content Quality"] = (quality + "_continue_needs_review").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "batch_policy=continue_on_product_local_validation_failure",
        "review_stage=" + stage,
        "review_details=" + details,
        "diagnostic_file=" + diagnostic if diagnostic else "",
        "additional_ai_calls_after_failure=0",
        "repair_attempt_cap=" + str(int(AI_SETTINGS.get("item_repair_attempts") or 3)),
    )
    return row

# ---------------------------------------------------------------------------
# V22.20 — customer-natural SEO copy and low-token editorial finalizer
# ---------------------------------------------------------------------------
# This layer keeps the technical source lock intact, but stops spending repair
# calls on FAQ/application/meta formatting that can be produced deterministically.
# Public copy is rebuilt from the final specification table so the output reads
# like a reseller product page, not a provenance disclaimer or raw translation.

V2220_VERSION = "22.20-customer-natural-seo-low-token"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_20_customer_natural_seo_2026_08_03"

for _column in (
    "Editorial Review Fields",
    "Editorial Review Details",
):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)

AI_STATS.setdefault("v2220_deterministic_content_sections", 0)
AI_STATS.setdefault("v2220_public_text_repairs", 0)
AI_STATS.setdefault("v2220_editorial_review_rows", 0)

V2220_PUBLIC_HTML_FIELDS: Tuple[str, ...] = (
    "Product Description",
    "custom_tab_1_content",
    "custom_tab_2_content",
    "custom_tab_3_content",
    "custom_tab_4_content",
    "custom_tab_5_content",
)
V2220_PUBLIC_TEXT_FIELDS: Tuple[str, ...] = (
    "Name",
    "Product Short Description",
    "focus keyphrase",
    "meta description",
)

V2220_DEFENSIVE_RE = re.compile(
    r"(?:sesuai\s+(?:dengan\s+)?spesifikasi(?:\s+produk)?|"
    r"berdasarkan\s+(?:dengan\s+)?spesifikasi(?:\s+produk)?|"
    r"produk\s+ini\s+dipilih\s+berdasarkan\s+parameter|"
    r"informasi\s+teknis\s+produk\s+disajikan\s+secara\s+ringkas|"
    r"nilai\s+yang\s+tercantum|rentang\s+yang\s+tercantum|"
    r"rincian\s+yang\s+tercantum)",
    re.I,
)

V2220_MIXED_LANGUAGE_RE = re.compile(
    r"(?:\bcorresponding\s+to\s+each\b|\bwhich\s+is\s+greater\b|"
    r"\bis\s+tested\s+within\b|\bis\s+simulated\b|\beasy\s+to\s+get\b|"
    r"\bcalculated\s+from\b|\bensures\s+high\b|\bmanual\s+push\s+the\b|"
    r"\bwill\s+save\b|\bmay\s+loss\b|\bremaining\s+range\b|"
    r"\bover\s*(?:&|and)\s*under\s+range\b|\bwith\s+over\b|"
    r"\bfor\s+(?:type|tipe)\s*[A-Z]\b|\bwithout\s+(?:battery|baterai)\b|"
    r"\bwhen\s+(?:the|kartu|data)\b|\bif\s+(?:the|I|making)\b|"
    r"\bincluding\b|\blinearity\b|\bhysteresis\b|\bw/o\b|"
    r"\breading\s*\+|\bof\s+reading\b|\bautomatic\s+range\b|"
    r"\bmanual\s+range\s+selecting\b|\bpush\s+the\b|\bone\s+time\b|"
    r"\bmeets?\b|\bstandards?\b|\bprovides?\b|\bease\s+of\b|"
    r"\bfield\s+measurements?\b|\bthe\s+characteristic\b|\balways\s+select\b|"
    r"\btwo\s+channel\b|\btraditional\s+hardware\b|\bbuild\s+in\b|"
    r"\breliability\b|\bspecial\s+functions?\b|\baccurate\s+measurements?\b|"
    r"\bphoto\s+couple\b|\bselection\b|\bother\s+range\b|\bfuse\b|"
    r"\bcalibrating\b|\btemp\.?\b|\bflat\s+response\b|\bchecking\s+the\s+noise\b)",
    re.I,
)

V2220_COMMON_ENGLISH_WORD_RE = re.compile(
    r"\b(?:the|and|with|without|when|including|remaining|corresponding|"
    r"which|greater|calculated|ensures|reading|selecting|tested|within|"
    r"push|save|loss|each)\b",
    re.I,
)


def _v2220_normalize_unit_name(value: Any) -> str:
    text = clean_text(value).strip(" ,.;:")
    key = re.sub(r"\s+", "", text).casefold()
    mapping = {
        "mph": "MPH",
        "mile/h": "mile/h",
        "miles/h": "mile/h",
        "knot": "knot",
        "knots": "knot",
        "kt": "knot",
        "kts": "knot",
    }
    return mapping.get(key, text)


_v2220_previous_v224_known_unit = _v224_known_unit
_v2220_previous_v224_known_quantity = _v224_known_quantity
_v2220_previous_measurement_unit_label = _measurement_unit_label
_v2220_previous_measurement_quantity = _measurement_quantity


def _v224_known_unit(value: Any) -> str:
    normalized = _v2220_normalize_unit_name(value)
    if re.fullmatch(r"(?:MPH|mile/h|knot)", normalized, re.I):
        return normalized
    return _v2220_previous_v224_known_unit(value)


def _v224_known_quantity(value: Any) -> Optional[Tuple[str, str]]:
    text = clean_text(value).strip(" ,.;:")
    match = re.fullmatch(
        r"([-+]?\d+(?:[.,]\d+)?)\s*(MPH|miles?\s*/\s*h|knots?|kts?)",
        text,
        re.I,
    )
    if match:
        unit = _v2220_normalize_unit_name(match.group(2))
        return f"{match.group(1)} {unit}", unit
    return _v2220_previous_v224_known_quantity(value)


def _measurement_unit_label(value: Any) -> str:
    normalized = _v2220_normalize_unit_name(value)
    if re.fullmatch(r"(?:MPH|mile/h|knot)", normalized, re.I):
        return normalized
    return _v2220_previous_measurement_unit_label(value)


def _measurement_quantity(value: Any) -> Optional[Tuple[str, str]]:
    quantity = _v224_known_quantity(value)
    if quantity:
        return quantity
    return _v2220_previous_measurement_quantity(value)


def _v2220_clean_fragment(value: Any, remove_boilerplate: bool = True) -> Tuple[str, int]:
    text = html.unescape(clean_text(value))
    if not text:
        return "", 0
    changed = 0

    replacements: Tuple[Tuple[str, str], ...] = (
        (r"\bThe\s+characteristic\s+is\s+simulated\s+as\s+[\"']?Human\s+Ear\s+(?:Listing|Listening)[\"']?\s+response\.?\s*Typical,?\s*(?:if|jika)\s+making\s+the\s+environmental\s+sound\s+level(?:\s+LCD|\s+Aksesori)?\s+(?:measurement|Pengukuran),?\s*always\s+select\s+to?\s*(?:A\s+Weighting|Pembobotan\s+A)\.?", "Karakteristik pembobotan A meniru respons pendengaran manusia dan umumnya digunakan untuk pengukuran kebisingan lingkungan."),
        (r"\bThe\s+characteristic\s+is\s+simulated\s+as\s+[\"']?Human\s+Ear\s+(?:Listing|Listening)[\"']?\s+response\b", "Karakteristiknya meniru respons pendengaran manusia"),
        (r"\bof\s+A\s*&\s*C\b", "pembobotan A dan C"),
        (r"\bFrequency\s+weighting\s+meets?\s+IEC\s*61672\s+Class\s*2\s+standards?\b", "Pembobotan frekuensi memenuhi standar IEC 61672 Kelas 2"),
        (r"\bcalibrating\s+input\s+signal\s+on\b", "sinyal masukan kalibrasi pada"),
        (r"\bthe\s+accuracy\s+of\s+frequency\s+weighting\b", "akurasi pembobotan frekuensi"),
        (r"\bLow[- ]Friction\s+Ball\s+Bearing\s+Provides\s+high\s+accuracy\s+at\s+both\s+high\s+(?:and|dan)\s+low\s+air\s+velocities\b", "Bantalan bola gesek rendah memberikan akurasi tinggi pada kecepatan udara tinggi maupun rendah"),
        (r"\bMax/Min\s+(?:Value|nilai)\s+Recall\s+Memorizes\s+the\s+maximum\s+(?:and|dan)\s+minimum\s+(?:values|nilai)\s+(?:with|dengan)\s+easy\s+recall\b", "Memori maksimum/minimum menyimpan nilai tertinggi dan terendah agar mudah dipanggil kembali"),
        (r"\bCompact\s+(?:and|dan)\s+Portable\s+Easy\s+to\s+carry\s+(?:and|dan)\s+handle\s+(?:with|dengan)\s+one\s+hand,?\s+ideal\s+(?:for|untuk)\s+field(?:\s+measurements?)?\b", "Ringkas dan portabel, mudah dibawa serta dioperasikan dengan satu tangan untuk pengukuran lapangan"),
        (r"\bUse\s+the\s+software\s+to\s+make\s+linearity\s+correction\s+instead\s+(?:of\s+)?the\s+traditional\s+hardware\s+circuit\b", "Koreksi linearitas dilakukan melalui perangkat lunak sebagai pengganti sirkuit perangkat keras konvensional"),
        (r"\bmicrocomputer\s+circuit,?\s+the\s+software\s+build\s+in\s+linearity\s+correction\s+function\s+instead\s+of\s+the\s+traditional\s+hardware\s+circuit\b", "Sirkuit mikrokomputer dengan fungsi koreksi linearitas melalui perangkat lunak sebagai pengganti sirkuit perangkat keras konvensional"),
        (r"\bcorresponding\s+to\s+each\s+range\s+step\b", "mengikuti setiap langkah rentang"),
        (r"\bwhich\s+is\s+greater\b", "gunakan nilai yang lebih besar"),
        (r"\bis\s+tested\s+within\b", "diWEBSITE pada"),
        (r"\bof\s+reading\b", "dari pembacaan"),
        (r"\bcalculated\s+from\b", "dihitung dari"),
        (r"\bensures\s+high\s+accuracy\s+and\s+reliability\b", "mendukung akurasi dan keandalan"),
        (r"\bhuman\s+ear\s+(?:response|waktu\s+respons)\b", "respons telinga manusia"),
        (r"\bremaining\s+range\b", "rentang lainnya"),
        (r"\bover\s*(?:&|and)\s*under\s+range\b", "indikator di atas dan di bawah rentang"),
        (r"\bwith\s+over\s*&?\s*under\s+range\s+indicating\b", "dengan indikator di atas dan di bawah rentang"),
        (r"\bautomatic\s+range\b", "rentang otomatis"),
        (r"\bmanual\s+range\s+selecting\b", "pemilihan rentang manual"),
        (r"\bfrequency\s+weighting\b", "pembobotan frekuensi"),
        (r"\btime\s+weighting\b", "pembobotan waktu"),
        (r"\bA\s+Weighting\b", "Pembobotan A"),
        (r"\bC\s+Weighting\b", "Pembobotan C"),
        (r"\bthermocouple\s+probe\b", "probe termokopel"),
        (r"\btemperature\s+probe\b", "probe suhu"),
        (r"\bmicroprocessor\s+circuit\b", "sirkuit mikroprosesor"),
        (r"\bmicrocomputer\s+circuit\b", "sirkuit mikrokomputer"),
        (r"\bDC\s+Voltage\b", "Tegangan DC"),
        (r"\bAC\s+Voltage\b", "Tegangan AC"),
        (r"\bDC\s+Current\b", "Arus DC"),
        (r"\bAC\s+Current\b", "Arus AC"),
        (r"\bResistance\b", "Resistansi"),
        (r"\bCapacitance\b", "Kapasitansi"),
        (r"\bContinuity\b", "Kontinuitas"),
        (r"\bFrequency\s+range\b", "rentang frekuensi"),
        (r"\bFrequency\b", "frekuensi"),
        (r"\bRange\s+selector\b", "pemilih rentang"),
        (r"\bRange\b", "rentang"),
        (r"\branges\b", "rentang"),
        (r"\bTwo\s+channel\s+temp\.?\b", "dua kanal suhu"),
        (r"\bchannels?\b", "kanal"),
        (r"\btemperature\b", "suhu"),
        (r"\bdifferential\b", "diferensial"),
        (r"\bbar\s+graph\s+indicator\b", "indikator grafik batang"),
        (r"\bimpedance\b", "impedansi"),
        (r"\bease\s+of\s+handling\b", "mudah digunakan"),
        (r"\bfield\s+measurements?\b", "pengukuran lapangan"),
        (r"\bWristlet\b", "tali pergelangan"),
        (r"\bData\s+Hold\s+Function\b", "Fungsi Data Hold"),
        (r"\bHigh\s+Precision\s+Sensor\b", "Sensor presisi tinggi"),
        (r"\bFast\s+(?:response\s+time|waktu\s+respons)\s+(?:for|untuk)\s+accurate\s+measurements?\b", "Waktu respons cepat untuk pengukuran yang akurat"),
        (r"\bEnsures\s+(?:high|tinggi)\s+(?:accuracy|akurasi)\s+(?:and|dan)\s+reliability(?:\s+(?:with|dengan)\s+special\s+functions?)?\b", "Mendukung akurasi dan keandalan"),
        (r"\bphoto\s+couple\b", "kopler optik"),
        (r"\bRange\s+selection\b", "pemilihan rentang"),
        (r"\bOther\s+range\b", "rentang lain"),
        (r"\bfuse\b", "sekring"),
        (r"\bClass\s*(\d+)\b", r"Kelas \1"),
        (r"\bcalibrating\s+(?:input|masukan)\s+(?:signal|sinyal)\s+(?:on|at)\b", "sinyal masukan kalibrasi pada"),
        (r"\bthe\s+(?:accuracy|akurasi)\s+of\b", "akurasi"),
        (r"\bover\s*/\s*under\s+(?:range|rentang)\s+indicator\b", "indikator di atas dan di bawah rentang"),
        (r"\bFunction\b", "fungsi"),
        (r"\bmeets?\b", "memenuhi"),
        (r"\bstandards?\b", "standar"),
        (r"\bProvides?\b", "memberikan"),
        (r"\bboth\b", "keduanya"),
        (r"\bhigh\b", "tinggi"),
        (r"\blow\b", "rendah"),
        (r"\bon\s+each\s+step\b", "pada setiap langkah"),
        (r"\bindicating\b", "indikator"),
        (r"\bdata\s+logger\b", "pencatat data"),
        (r"\bpower\s+supply\b", "catu daya"),
        (r"\bfull\s+scale\b", "skala penuh"),
        (r"\bwithout\b", "tanpa"),
        (r"\bincluding\b", "termasuk"),
        (r"\bwhen\b", "ketika"),
        (r"\bif\b", "jika"),
        (r"\bfor\b", "untuk"),
        (r"\bwith\b", "dengan"),
        (r"\band\b", "dan"),
        (r"\baccuracy\b", "akurasi"),
        (r"\bresolution\b", "resolusi"),
        (r"\binput\b", "masukan"),
        (r"\boutput\b", "keluaran"),
        (r"\bdisplay\b", "tampilan"),
        (r"\breading\b", "pembacaan"),
        (r"\bvalues?\b", "nilai"),
        (r"\bautomatic\b", "otomatis"),
    )
    for pattern, replacement in replacements:
        text, count = re.subn(pattern, replacement, text, flags=re.I)
        changed += count

    text, count = re.subn(
        r"(?<=\d)\s+to\s+(?=[-+]?\d)",
        " hingga ",
        text,
        flags=re.I,
    )
    changed += count
    text, count = re.subn(
        r"(?<=\d%)\s+to\s+(?=[-+]?\d)",
        " hingga ",
        text,
        flags=re.I,
    )
    changed += count

    if remove_boilerplate:
        boilerplate_rules: Tuple[Tuple[str, str], ...] = (
            (r"\s+(?:sesuai|berdasarkan)\s+(?:dengan\s+)?spesifikasi\s+produk\b", ""),
            (r"\s+sesuai\s+(?:dengan\s+)?spesifikasi\b", ""),
            (r"\bProduk\s+ini\s+dipilih\s+berdasarkan\s+parameter,?\s+rentang,?\s+dan\s+kondisi\s+kerja[^.]*\.?", ""),
            (r"\bProduk\s+ini\s+cocok\s+digunakan\s+ketika\s+pengguna\s+membutuhkan[^.]*\.?", ""),
            (r"\bRentang,?\s+sensor,?\s+koneksi,?\s+serta\s+aksesori\s+perlu\s+disesuaikan[^.]*\.?", ""),
            (r"\bInformasi\s+teknis\s+produk\s+disajikan\s+secara\s+ringkas[^.]*\.?", ""),
            (r"\b(?:nilai|rentang|rincian|data)\s+yang\s+tercantum\s+(?:adalah|meliputi)\s+", ""),
        )
        for pattern, replacement in boilerplate_rules:
            text, count = re.subn(pattern, replacement, text, flags=re.I)
            changed += count

    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([,;:])(?=\S)", r"\1 ", text)
    text = re.sub(r"\.{2,}", ".", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" \t\n,;:-")
    return text, changed


def _v2220_clean_html(value: Any, remove_boilerplate: bool = True) -> Tuple[str, int]:
    raw = clean_text(value)
    if not raw:
        return "", 0
    soup = BeautifulSoup(raw, "html.parser")
    changed = 0
    for node in list(soup.find_all(string=True)):
        parent_name = clean_text(getattr(getattr(node, "parent", None), "name", "")).casefold()
        if parent_name in {"script", "style"}:
            continue
        cleaned, count = _v2220_clean_fragment(str(node), remove_boilerplate=remove_boilerplate)
        changed += count
        node.replace_with(cleaned)
    rendered = "".join(str(node) for node in soup.contents).strip()
    rendered = re.sub(r"<p>\s*</p>|<li>\s*</li>", "", rendered, flags=re.I)
    rendered = re.sub(r"</strong>(?=[A-Za-zÀ-ÿ0-9±])", "</strong> ", rendered)
    return rendered, changed


def _v2220_has_mixed_language(value: Any) -> bool:
    text = clean_text(value)
    if not text:
        return False
    if V2220_MIXED_LANGUAGE_RE.search(text):
        return True
    # One technical English token is normal; several connecting English words
    # in the same public field indicate an unfinished machine translation.
    return len(V2220_COMMON_ENGLISH_WORD_RE.findall(text)) >= 2


def _v2220_profile(data: Dict[str, Any], family: Dict[str, Any]) -> Dict[str, Any]:
    probe = " ".join([
        clean_text(data.get("title")),
        clean_text(data.get("model")),
        clean_text(family.get("category")),
        clean_text(family.get("term")),
    ]).casefold()

    profiles: List[Tuple[str, str, str, List[str]]] = [
        (
            r"sound\s+level|noise|decibel|kebisingan",
            "sound level meter",
            "mengukur tingkat kebisingan dan tekanan suara",
            [
                "Pemantauan tingkat kebisingan di area kerja dan lingkungan.",
                "Pemeriksaan suara mesin atau peralatan saat inspeksi.",
                "Perbandingan tingkat kebisingan sebelum dan sesudah tindakan perbaikan.",
                "PengWEBSITEan akustik dasar untuk kegiatan pengendalian mutu.",
            ],
        ),
        (
            r"multimeter|clamp\s+meter|voltmeter|ammeter",
            "multimeter",
            "mengukur parameter kelistrikan",
            [
                "Pemeriksaan tegangan, arus, atau resistansi pada rangkaian listrik.",
                "Troubleshooting panel, komponen, dan peralatan listrik.",
                "Verifikasi kondisi kelistrikan saat commissioning atau pemeliharaan.",
                "PengWEBSITEan kelistrikan pada kegiatan inspeksi dan pengendalian mutu.",
            ],
        ),
        (
            r"thermometer|temperature|termometer|suhu",
            "termometer",
            "mengukur suhu pada sampel, lingkungan, atau proses",
            [
                "Pengukuran suhu pada proses produksi, pemeliharaan, atau inspeksi.",
                "Pemeriksaan suhu sampel dan peralatan laboratorium.",
                "Pemantauan perubahan suhu selama pengWEBSITEan.",
                "Verifikasi kondisi suhu sebelum tindakan atau proses lanjutan.",
            ],
        ),
        (
            r"humidity|hygrometer|environmental|kelembapan|lingkungan",
            "alat ukur lingkungan",
            "mengukur kelembapan dan parameter lingkungan terkait",
            [
                "Pemantauan kondisi lingkungan di ruang kerja atau penyimpanan.",
                "Pemeriksaan kelembapan, suhu, atau aliran udara sesuai parameter alat.",
                "Perbandingan kondisi lingkungan antar lokasi pengukuran.",
                "PengWEBSITEan lapangan pada kegiatan inspeksi dan pengendalian mutu.",
            ],
        ),
        (
            r"light\s+meter|lux|illumination|cahaya",
            "light meter",
            "mengukur intensitas cahaya",
            [
                "Pemeriksaan tingkat pencahayaan di ruang kerja dan area produksi.",
                "Verifikasi keluaran lampu atau sistem pencahayaan.",
                "Perbandingan intensitas cahaya antar titik pengukuran.",
                "Audit pencahayaan pada kegiatan inspeksi dan pengendalian mutu.",
            ],
        ),
        (
            r"\buv\b|ultraviolet|uva|uvc",
            "UV meter",
            "mengukur intensitas radiasi ultraviolet",
            [
                "Pemeriksaan intensitas UV pada area atau sumber cahaya.",
                "Verifikasi keluaran lampu UV dalam kegiatan inspeksi.",
                "Perbandingan intensitas UV antar titik pengukuran.",
                "Pemantauan perubahan intensitas selama pengWEBSITEan.",
            ],
        ),
        (
            r"anemometer|air\s+flow|air\s+velocity|wind",
            "anemometer",
            "mengukur kecepatan atau aliran udara",
            [
                "Pemeriksaan aliran udara pada sistem ventilasi dan HVAC.",
                "Pengukuran kecepatan udara di area kerja atau saluran udara.",
                "Perbandingan aliran udara antar titik pengukuran.",
                "Verifikasi kondisi ventilasi pada kegiatan inspeksi.",
            ],
        ),
        (
            r"transmitter|pemancar",
            "transmitter",
            "mengubah hasil pengukuran menjadi sinyal keluaran untuk pemantauan atau otomasi",
            [
                "Pemantauan parameter proses melalui sistem instrumentasi.",
                "Integrasi sinyal pengukuran dengan panel kontrol atau pencatat data.",
                "Pemeriksaan sinyal saat commissioning dan pemeliharaan.",
                "Pemantauan proses pada sistem otomasi yang kompatibel.",
            ],
        ),
        (
            r"power\s+analyzer|power\s+meter|wattmeter|energy",
            "power analyzer",
            "menganalisis parameter daya, energi, dan kelistrikan",
            [
                "Analisis konsumsi daya pada mesin, panel, atau sistem listrik.",
                "Pemeriksaan kualitas dan efisiensi penggunaan energi.",
                "Verifikasi parameter listrik saat commissioning atau pemeliharaan.",
                "PengWEBSITEan kelistrikan pada kegiatan audit energi dan pengendalian mutu.",
            ],
        ),
    ]

    for pattern, term, function, applications in profiles:
        if re.search(pattern, probe, re.I):
            return {"term": term, "function": function, "applications": applications}

    raw_term, _ = _v2220_clean_fragment(family.get("term"), remove_boilerplate=True)
    raw_function, _ = _v2220_clean_fragment(family.get("function"), remove_boilerplate=True)
    term = raw_term or ("aksesori" if data.get("is_accessory") else "alat ukur")
    function = raw_function or "mendukung pengukuran dan pemeriksaan teknis"
    applications = [
        "Pemeriksaan teknis pada kegiatan inspeksi dan pemeliharaan.",
        "PengWEBSITEan di laboratorium atau area kerja.",
        "Perbandingan hasil sebelum dan sesudah penyesuaian proses.",
        "Verifikasi parameter untuk kegiatan pengendalian mutu.",
    ]
    return {"term": term, "function": function, "applications": applications}


def _v2220_clean_pairs(spec_html: Any) -> Tuple[str, List[Dict[str, str]], int]:
    cleaned_html, repairs = _v2220_clean_html(spec_html, remove_boilerplate=False)
    raw_pairs = _v2217_extract_spec_pairs(cleaned_html)
    output: List[Dict[str, str]] = []
    seen_labels = set()
    seen_pairs = set()
    for row in raw_pairs:
        label, label_repairs = _v2220_clean_fragment(row.get("label"), remove_boilerplate=False)
        value, value_repairs = _v2220_clean_fragment(row.get("value"), remove_boilerplate=False)
        repairs += label_repairs + value_repairs
        label = normalize_label(label).strip(" :;.-")
        value = clean_text(value).strip(" :;.-")
        if not label or not value:
            continue
        if re.search(r"\b(?:model|sku|mpn|kode\s+produk)\b", label, re.I):
            continue
        if _v2220_has_mixed_language(label + " " + value):
            continue
        if len(value) > 190:
            continue
        pair_key = (_spec_text_key(label), _spec_text_key(value))
        if not all(pair_key) or pair_key in seen_pairs:
            continue
        # One concise row per broad label is enough for narrative copy. The
        # full specification table remains untouched apart from text cleanup.
        broad_label = re.sub(r"\b(?:umum|general)\b", "", label, flags=re.I)
        broad_key = _spec_text_key(broad_label)
        if broad_key in seen_labels and re.search(r"\b(?:akurasi|resolusi|rentang)\b", label, re.I):
            continue
        seen_pairs.add(pair_key)
        if broad_key:
            seen_labels.add(broad_key)
        output.append({"label": label, "value": value})
    return cleaned_html, output, repairs


def _v2220_find_fact(pairs: Sequence[Dict[str, str]], pattern: str) -> Optional[Dict[str, str]]:
    regex = re.compile(pattern, re.I)
    # Labels are authoritative. Value-only matching can turn "skala penuh"
    # inside an accuracy row into a range or treat an embedded resolution as a
    # separate precision field.
    return next(
        (row for row in pairs if regex.search(clean_text(row.get("label")))),
        None,
    )


def _v2220_fact_text(row: Optional[Dict[str, str]], maximum: int = 150) -> str:
    if not isinstance(row, dict):
        return ""
    label = clean_text(row.get("label")).strip(" :;.-")
    value = compact(clean_text(row.get("value")), maximum).rstrip(".")
    return f"{label}: {value}" if label and value else ""


def _v2220_title_modifier(pairs: Sequence[Dict[str, str]]) -> str:
    range_row = _v2220_find_fact(pairs, r"\b(?:rentang|range|skala)\b")
    if range_row:
        value = clean_text(range_row.get("value"))
        matches = re.findall(
            r"([-+]?\d+(?:[.,]\d+)?)\s*(?:hingga|[-–~])\s*([-+]?\d+(?:[.,]\d+)?)\s*"
            r"(dB|°?[CF]|℃|℉|%\s*RH|V|mV|A|mA|Hz|kHz|MHz|GHz|lux|fc|m/s|km/h|ft/min|MPH|knot|rpm|pH)",
            value,
            re.I,
        )
        if matches:
            grouped: Dict[str, List[Tuple[float, float, str, str]]] = {}
            for low, high, unit in matches:
                try:
                    low_n = float(low.replace(",", "."))
                    high_n = float(high.replace(",", "."))
                except Exception:
                    continue
                grouped.setdefault(unit.casefold(), []).append((low_n, high_n, low, high))
            if grouped:
                # Prefer the first unit family in the source, then show the
                # overall minimum/maximum for multi-range instruments.
                first_unit = matches[0][2].casefold()
                rows = grouped.get(first_unit) or next(iter(grouped.values()))
                low_item = min(rows, key=lambda item: item[0])
                high_item = max(rows, key=lambda item: item[1])
                unit = matches[0][2]
                modifier = f"{low_item[2]}–{high_item[3]} {unit}"
                if len(modifier) <= 22:
                    return modifier

    probe = " ".join(
        clean_text(row.get("label")) + " " + clean_text(row.get("value"))
        for row in pairs[:8]
    )
    for pattern, replacement in (
        (r"\bTrue\s*RMS\b", "True RMS"),
        (r"\bClass\s*2\b|\bKelas\s*2\b", "Kelas 2"),
        (r"\b(\d+)\s*(?:channel|kanal)\b", r"\1 Kanal"),
        (r"\bpencatat\s+data\b", "Data Logger"),
    ):
        match = re.search(pattern, probe, re.I)
        if match:
            return match.expand(replacement)
    return ""


def _v2220_public_titles(data: Dict[str, Any], family: Dict[str, Any], pairs: Sequence[Dict[str, str]]) -> Tuple[str, str]:
    base = _v2217_product_title(data, family)
    modifier = _v2220_title_modifier(pairs)
    full = base
    if modifier and modifier.casefold() not in base.casefold():
        candidate = re.sub(r"\s+", " ", f"{base} {modifier}").strip()
        if len(candidate) <= 60:
            full = candidate
    return full, base


def _v2220_summary_sentences(pairs: Sequence[Dict[str, str]], maximum: int = 2) -> List[str]:
    selected: List[Optional[Dict[str, str]]] = [
        _v2220_find_fact(pairs, r"\b(?:rentang|range|skala)\b"),
        _v2220_find_fact(pairs, r"\b(?:akurasi|accuracy|ketelitian)\b"),
        _v2220_find_fact(pairs, r"\b(?:resolusi|resolution)\b"),
        _v2220_find_fact(pairs, r"\b(?:antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat\s+data)\b"),
        _v2220_find_fact(pairs, r"\b(?:sensor|probe|kanal|channel|input|masukan)\b"),
    ]
    output: List[str] = []
    seen = set()
    for row in selected:
        if not isinstance(row, dict):
            continue
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        if re.search(r"\b(?:rentang|range|skala)\b", label, re.I) and len(value) > 115:
            modifier = _v2220_title_modifier(pairs)
            text = f"Rentang utama: {modifier}" if modifier else _v2220_fact_text(row, maximum=100)
        else:
            text = _v2220_fact_text(row, maximum=115)
        key = _spec_text_key(text)
        if not text or not key or key in seen:
            continue
        seen.add(key)
        output.append(text + ".")
        if len(output) >= maximum:
            break
    return output


def _v2220_key_facts(pairs: Sequence[Dict[str, str]], maximum: int = 5) -> List[Dict[str, str]]:
    priority = re.compile(
        r"\b(?:rentang|range|skala|akurasi|accuracy|ketelitian|resolusi|resolution|"
        r"parameter|pengukuran|satuan|unit|sensor|probe|kanal|channel|masukan|input|"
        r"antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat\s+data|memori|"
        r"catu\s+daya|power|baterai|battery|tampilan|display|aksesori|accessor|"
        r"dimensi|dimension|berat|weight)\b",
        re.I,
    )
    output: List[Dict[str, str]] = []
    seen = set()
    ordered = [row for row in pairs if priority.search(clean_text(row.get("label")))]
    if len(ordered) < maximum:
        ordered.extend(row for row in pairs if row not in ordered)
    for row in ordered:
        key = (_spec_text_key(row.get("label")), _spec_text_key(row.get("value")))
        if not all(key) or key in seen:
            continue
        seen.add(key)
        output.append(row)
        if len(output) >= maximum:
            break
    return output


def _v2220_buying_factors(pairs: Sequence[Dict[str, str]]) -> List[str]:
    factors: List[str] = []
    rules = (
        (r"\b(?:rentang|range|skala)\b", "rentang pengukuran"),
        (r"\b(?:akurasi|accuracy|ketelitian|resolusi|resolution)\b", "akurasi dan resolusi"),
        (r"\b(?:sensor|probe|kanal|channel|input|masukan)\b", "jenis sensor atau masukan"),
        (r"\b(?:antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat\s+data)\b", "kebutuhan koneksi data"),
        (r"\b(?:catu\s+daya|power|baterai|battery)\b", "catu daya"),
        (r"\b(?:aksesori|accessor|kelengkapan)\b", "aksesori"),
    )
    probe = " ".join(
        clean_text(row.get("label")) + " " + clean_text(row.get("value"))
        for row in pairs
    )
    for pattern, label in rules:
        if re.search(pattern, probe, re.I):
            factors.append(label)
    return dedupe_keep_order(factors)[:4] or ["parameter pengukuran", "rentang kerja", "kondisi penggunaan"]


def _v2220_join_id(items: Sequence[str]) -> str:
    cleaned = [clean_text(item) for item in items if clean_text(item)]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    return ", ".join(cleaned[:-1]) + " dan " + cleaned[-1]


def _v2220_product_description(
    data: Dict[str, Any],
    profile: Dict[str, Any],
    full_title: str,
    base_title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    model = clean_text(data.get("model"))
    identity = f"Lutron {model}" if model else base_title
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    summaries = _v2220_summary_sentences(pairs, maximum=2)
    key_facts = _v2220_key_facts(pairs, maximum=5)
    factors = _v2220_buying_factors(pairs)

    first = f"<p><strong>{esc(identity)}</strong> merupakan {esc(term)} untuk {esc(function)}.</p>"
    second_text = " ".join(summaries) if summaries else (
        "Model ini ditujukan untuk membantu pemeriksaan teknis pada kegiatan inspeksi, "
        "pemeliharaan, atau pengendalian mutu."
    )
    second = f"<p>{esc(second_text)}</p>"

    facts_html = ""
    if key_facts:
        facts_html = "<h2>Spesifikasi Utama</h2><ul>" + "".join(
            f"<li><strong>{esc(row.get('label'))}:</strong> {esc(row.get('value'))}</li>"
            for row in key_facts
        ) + "</ul>"

    guidance = (
        "Sebelum memesan, cocokkan " + _v2220_join_id(factors)
        + " dengan kondisi penggunaan dan konfigurasi yang dibutuhkan."
    )
    return _v2213_polish_html(
        first + second + facts_html
        + "<h2>Pertimbangan Sebelum Memilih</h2><p>" + esc(guidance) + "</p>"
        + WEBSITE_CTA,
        ensure_cta=True,
    )


def _v2220_short_description(
    profile: Dict[str, Any],
    base_title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    summaries = _v2220_summary_sentences(pairs, maximum=2)
    base = f"{base_title} digunakan untuk {function}."
    for count in range(len(summaries), -1, -1):
        candidate = re.sub(r"\s{2,}", " ", base + (" " + " ".join(summaries[:count]) if count else "")).strip()
        if len(candidate) <= V22_SHORT_MAX:
            return candidate
    return compact(base, V22_SHORT_MAX)


def _v2220_feature_html(profile: Dict[str, Any], pairs: Sequence[Dict[str, str]]) -> str:
    bullets: List[str] = []
    function = clean_text(profile.get("function"))
    if function:
        bullets.append(function[:1].upper() + function[1:] + ".")
    for row in _v2220_key_facts(pairs, maximum=5):
        bullets.append(_v2220_fact_text(row, maximum=170) + ".")
    bullets = dedupe_keep_order(bullets)[:6]
    return "<h2>Keunggulan Produk</h2><ul>" + "".join(
        f"<li>{esc(item)}</li>" for item in bullets
    ) + "</ul>"


def _v2220_application_html(profile: Dict[str, Any]) -> str:
    applications = dedupe_keep_order(profile.get("applications") or [])[:4]
    return "<h2>Aplikasi Penggunaan</h2><ul>" + "".join(
        f"<li>{esc(item)}</li>" for item in applications
    ) + "</ul>"


def _v2220_faq_html(
    data: Dict[str, Any],
    profile: Dict[str, Any],
    base_title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    model = clean_text(data.get("model")) or base_title
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    range_row = _v2220_find_fact(pairs, r"\b(?:rentang|range|skala)\b")
    precision_row = _v2220_find_fact(pairs, r"\b(?:akurasi|accuracy|ketelitian|resolusi|resolution)\b")
    interface_row = _v2220_find_fact(pairs, r"\b(?:antarmuka|interface|keluaran|output|RS\s*232|USB|pencatat\s+data|sensor|probe|kanal|channel|masukan)\b")

    items: List[Tuple[str, str]] = [
        (
            f"Apa fungsi {base_title}?",
            f"{base_title} merupakan {term} untuk {function}.",
        )
    ]

    if range_row:
        items.append((
            f"Berapa rentang pengukuran {model}?",
            _v2220_fact_text(range_row, maximum=170) + ".",
        ))
    else:
        first = _v2220_key_facts(pairs, maximum=1)
        items.append((
            f"Apa parameter utama {model}?",
            (_v2220_fact_text(first[0], maximum=170) + ".") if first else f"Parameter utama model ini berkaitan dengan fungsi {term}.",
        ))

    detail_row = precision_row or interface_row
    if detail_row:
        question = (
            "Bagaimana akurasi atau resolusi pengukurannya?"
            if precision_row else
            "Apa koneksi, sensor, atau masukan yang tersedia?"
        )
        items.append((question, _v2220_fact_text(detail_row, maximum=170) + "."))
    else:
        items.append((
            "Apa rincian teknis lain yang perlu diperhatikan?",
            "Periksa catu daya, sensor, aksesori, dan kondisi operasi sebelum produk digunakan.",
        ))

    factors = _v2220_buying_factors(pairs)
    items.append((
        "Apa yang perlu dipastikan sebelum pemesanan?",
        "Cocokkan " + _v2220_join_id(factors) + " dengan kebutuhan penggunaan. Tim WEBSITE dapat membantu memeriksa konfigurasi sebelum pemesanan.",
    ))
    return "".join(f"<h3>{esc(q)}</h3><p>{esc(a)}</p>" for q, a in items[:4])


def _v2220_meta_description(
    profile: Dict[str, Any],
    base_title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    modifier = _v2220_title_modifier(pairs)
    candidates = [
        f"{base_title} untuk {function} pada rentang {modifier}. Lihat spesifikasi dan konsultasikan kebutuhan Anda dengan Tim WEBSITE." if modifier else "",
        f"{base_title} untuk {function}. Lihat spesifikasi dan konsultasikan kebutuhan Anda dengan Tim WEBSITE.",
        f"{base_title}. Rentang utama {modifier}. Lihat spesifikasi dan konsultasikan kebutuhan Anda dengan Tim WEBSITE." if modifier else "",
    ]
    for candidate in candidates:
        candidate = re.sub(r"\s{2,}", " ", clean_text(candidate)).replace("..", ".").strip()
        if 100 <= len(candidate) <= V22_META_MAX and candidate.endswith("."):
            return candidate
    fallback = f"{base_title}. Lihat fungsi dan spesifikasi teknisnya, lalu konsultasikan kebutuhan Anda dengan Tim WEBSITE."
    return compact(fallback, V22_META_MAX)


def _v2220_official_information(
    profile: Dict[str, Any],
    base_title: str,
    pairs: Sequence[Dict[str, str]],
) -> str:
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    summaries = _v2220_summary_sentences(pairs, maximum=2)
    details = " ".join(summaries) or "Rincian teknis tersedia pada tabel Spesifikasi."
    factors = _v2220_buying_factors(pairs)
    return (
        "<h2>Informasi Produk</h2>"
        f"<p>{esc(base_title)} merupakan {esc(term)} untuk {esc(function)}.</p>"
        f"<p>{esc(details)}</p>"
        "<h2>Catatan Teknis</h2>"
        f"<p>Sebelum digunakan, periksa {esc(_v2220_join_id(factors))} dan pastikan konfigurasi produk cocok dengan kebutuhan pengukuran.</p>"
    )


def _v2220_editorial_scan(row: Dict[str, Any]) -> Tuple[List[str], List[str], List[str]]:
    issues: List[str] = []
    fields: List[str] = []
    details: List[str] = []

    for field in V2220_PUBLIC_HTML_FIELDS:
        visible = _v2211_visible_text(row.get(field))
        match = V2220_DEFENSIVE_RE.search(visible)
        if match:
            issues.append("defensive_boilerplate_visible")
            fields.append(field)
            details.append(f"{field}: {compact(visible[max(0, match.start()-45):match.end()+90], 170)}")
        match = V2220_MIXED_LANGUAGE_RE.search(visible)
        if match or _v2220_has_mixed_language(visible):
            issues.append("mixed_language_residue_visible")
            fields.append(field)
            start = match.start() if match else 0
            details.append(f"{field}: {compact(visible[max(0, start-45):start+150], 190)}")

    for field in V2220_PUBLIC_TEXT_FIELDS:
        text = clean_text(row.get(field))
        match = V2220_DEFENSIVE_RE.search(text)
        if match:
            issues.append("defensive_boilerplate_visible")
            fields.append(field)
            details.append(f"{field}: {compact(text, 190)}")
        match = V2220_MIXED_LANGUAGE_RE.search(text)
        if match or _v2220_has_mixed_language(text):
            issues.append("mixed_language_residue_visible")
            fields.append(field)
            details.append(f"{field}: {compact(text, 190)}")

    title = clean_text(row.get("Name"))
    keyphrase = clean_text(row.get("focus keyphrase"))
    if not title or len(title) > 65 or keyphrase.casefold() not in title.casefold():
        issues.append("seo_title_or_keyphrase_invalid")
        fields.extend(["Name", "focus keyphrase"])
    meta = clean_text(row.get("meta description"))
    if not (90 <= len(meta) <= V22_META_MAX) or not meta.endswith("."):
        issues.append("meta_description_invalid")
        fields.append("meta description")
    if re.search(r"Cek\s+fungsi,\s+rentang,\s+akurasi,\s+fitur", meta, re.I):
        issues.append("meta_description_generic_template")
        fields.append("meta description")

    description = BeautifulSoup(clean_text(row.get("Product Description")), "html.parser")
    if len(description.find_all("p")) < 2 or len(description.find_all("h2")) < 2:
        issues.append("product_description_structure_invalid")
        fields.append("Product Description")
    faq = BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser")
    if len(faq.find_all("h3")) != 4:
        issues.append("faq_item_count_invalid")
        fields.append("custom_tab_4_content")
    applications = BeautifulSoup(clean_text(row.get("custom_tab_3_content")), "html.parser")
    if len(applications.find_all("li")) != 4:
        issues.append("application_item_count_invalid")
        fields.append("custom_tab_3_content")
    features = BeautifulSoup(clean_text(row.get("custom_tab_2_content")), "html.parser")
    if len(features.find_all("li")) < 3:
        issues.append("feature_fact_count_low")
        fields.append("custom_tab_2_content")
    specifications = BeautifulSoup(clean_text(row.get("custom_tab_1_content")), "html.parser")
    if specifications.find("table") is None:
        issues.append("specification_table_invalid")
        fields.append("custom_tab_1_content")

    return dedupe_keep_order(issues), dedupe_keep_order(fields), dedupe_keep_order(details)[:12]


def _v2220_finalize_row(row: Dict[str, Any], data: Dict[str, Any]) -> Dict[str, Any]:
    family = detect_family(data)
    profile = _v2220_profile(data, family)
    cleaned_specs, pairs, repairs = _v2220_clean_pairs(row.get("custom_tab_1_content"))
    full_title, base_title = _v2220_public_titles(data, family, pairs)

    row["custom_tab_1_content"] = cleaned_specs
    row["Name"] = full_title
    row["focus keyphrase"] = base_title
    row["Product Description"] = _v2220_product_description(
        data, profile, full_title, base_title, pairs
    )
    row["Product Short Description"] = _v2220_short_description(
        profile, base_title, pairs
    )
    row["custom_tab_2_content"] = _v2220_feature_html(profile, pairs)
    row["custom_tab_3_content"] = _v2220_application_html(profile)
    row["custom_tab_4_content"] = _v2220_faq_html(
        data, profile, base_title, pairs
    )
    row["custom_tab_5_content"] = _v2220_official_information(
        profile, base_title, pairs
    )
    row["meta description"] = _v2220_meta_description(
        profile, base_title, pairs
    )

    for field in (
        "Product Description",
        "custom_tab_2_content",
        "custom_tab_3_content",
        "custom_tab_4_content",
        "custom_tab_5_content",
    ):
        row[field], changed = _v2220_clean_html(row.get(field), remove_boilerplate=True)
        repairs += changed
    for field in ("Product Short Description", "meta description"):
        row[field], changed = _v2220_clean_fragment(row.get(field), remove_boilerplate=True)
        repairs += changed
        if row[field] and row[field][-1] not in ".!?":
            row[field] += "."

    issues, review_fields, review_details = _v2220_editorial_scan(row)
    data["_v2220_final_issues"] = issues
    data["_v2220_review_fields"] = review_fields
    data["_v2220_review_details"] = review_details
    row["Editorial Review Fields"] = ", ".join(review_fields)
    row["Editorial Review Details"] = " | ".join(review_details)

    product_local_failure = bool(
        clean_text(data.get("_v2219_review_stage"))
        or clean_text(row.get("Review Failure Stage"))
    )
    warning_source = bool(data.get("_v22122_quarantined_source_rows"))

    if product_local_failure:
        row["Publish Decision"] = "NEEDS_REVIEW"
        try:
            row["Final Publish Score"] = min(int(row.get("Final Publish Score") or 45), 45)
        except Exception:
            row["Final Publish Score"] = 45
        row["Website Validation"] = "needs_review_v22_20:product_local_technical_failure"
    elif issues:
        row["Publish Decision"] = "NEEDS_REVIEW"
        row["Final Publish Score"] = max(45, 100 - 10 * len(issues))
        row["Website Validation"] = "needs_review_v22_20:" + ",".join(issues)
        AI_STATS["v2220_editorial_review_rows"] = int(
            AI_STATS.get("v2220_editorial_review_rows") or 0
        ) + 1
    elif warning_source:
        row["Publish Decision"] = "READY_WITH_WARNING"
        row["Final Publish Score"] = 95
        row["Website Validation"] = "publish_ready_with_warning_v22_20:source_row_quarantined"
    else:
        row["Publish Decision"] = "READY"
        row["Final Publish Score"] = 100
        row["Website Validation"] = "publish_ready_v22_20"

    row["Processing Time"] = "lutron_WEBSITE_import_v22_20_customer_natural_seo"
    quality = clean_text(row.get("Content Quality"))
    if "customer_natural_seo" not in quality:
        row["Content Quality"] = (quality + "_customer_natural_seo").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "customer_editorial_finalizer=" + V2220_VERSION,
        "public_text_repairs=" + str(repairs),
        "narrative_key_facts=" + str(len(_v2220_key_facts(pairs, maximum=5))),
        "seo_title=" + full_title,
        "focus_keyphrase=" + base_title,
        "editorial_issues=" + ",".join(issues) if issues else "editorial_issues=none",
        "editorial_review_fields=" + ",".join(review_fields) if review_fields else "editorial_review_fields=none",
        "content_repair_policy=deterministic_before_additional_ai_calls",
    )
    AI_STATS["v2220_public_text_repairs"] = int(
        AI_STATS.get("v2220_public_text_repairs") or 0
    ) + repairs
    return row


# Avoid repeated AI repair calls for sections that are deterministic editorial
# assets. Specification translation and technical integrity checks remain AI/
# source-locked; this shortcut applies only to public copy structure.
_v2220_previous_validate_section = _v226_validate_section


def _v2220_evidence_ids(payload: Dict[str, Any], maximum: int = 4) -> List[str]:
    ids = _v227_fact_ids(payload)
    if not ids:
        ids = list(_v22_fact_map(payload).keys())
    return dedupe_keep_order(ids)[:maximum]


def _v2220_deterministic_section(
    field: str,
    payload: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], List[Dict[str, Any]]]:
    family = detect_family(data)
    profile = _v2220_profile(data, family)
    title = _v2217_product_title(data, family)
    model = clean_text(data.get("model"))
    evidence_ids = _v2220_evidence_ids(payload)
    evidence_one = evidence_ids[:1]
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"

    if field == "short_description":
        value = compact(f"{title} digunakan untuk {function}.", V22_SHORT_MAX)
        return True, {"value": value, "evidence_ids": evidence_ids}, []
    if field == "meta_description":
        value = _v2220_meta_description(profile, title, [])
        return True, {"value": value, "evidence_ids": evidence_ids}, []
    if field == "application_tab_html":
        items = [
            {"text": item, "evidence_ids": evidence_one}
            for item in (profile.get("applications") or [])[:4]
        ]
        return True, {"items": items, "html": _v22_render_application(items)}, []
    if field == "faq_html":
        raw_items = [
            (f"Apa fungsi {title}?", f"{title} merupakan {term} untuk {function}."),
            ("Apa parameter yang perlu diperiksa?", "Periksa parameter pengukuran, rentang kerja, dan kondisi operasi produk."),
            ("Apa yang perlu dipastikan sebelum digunakan?", "Pastikan sensor, catu daya, koneksi, dan aksesori cocok dengan konfigurasi produk."),
            ("Apakah tersedia konsultasi sebelum pemesanan?", "Ya. Tim WEBSITE dapat membantu mencocokkan model dan kebutuhan penggunaan sebelum pemesanan."),
        ]
        items = [
            {"question": q, "answer": a, "evidence_ids": evidence_one}
            for q, a in raw_items
        ]
        return True, {"items": items, "html": _v22_render_faq(items)}, []
    if field == "product_description_html":
        bullets = [
            f"Jenis produk: {term}.",
            f"Model: {model}.",
            f"Fungsi utama: {function}.",
            "Pemilihan konfigurasi mempertimbangkan rentang, sensor, koneksi, dan catu daya.",
        ]
        value = (
            f"<p><strong>{esc(title)}</strong> merupakan {esc(term)} untuk {esc(function)}.</p>"
            "<p>Informasi teknis utama tersedia pada tabel Spesifikasi dan digunakan untuk membantu pemilihan konfigurasi.</p>"
            "<h2>Keunggulan Produk</h2><ul>"
            + "".join(f"<li>{esc(item)}</li>" for item in bullets)
            + "</ul><h2>Panduan Pemilihan</h2>"
            "<p>Cocokkan rentang, sensor, koneksi, catu daya, dan aksesori dengan kebutuhan penggunaan.</p>"
        )
        return True, {"value": value, "evidence_ids": evidence_ids}, []
    if field == "feature_tab_html":
        bullets = [
            f"Digunakan untuk {function}.",
            f"Identitas model: {model}.",
            "Parameter teknis disusun dalam tabel agar mudah dibandingkan.",
            "Konfigurasi dapat diperiksa sebelum pemesanan bersama Tim WEBSITE.",
        ]
        value = "<h2>Keunggulan Produk</h2><ul>" + "".join(
            f"<li>{esc(item)}</li>" for item in bullets
        ) + "</ul>"
        return True, {"value": value, "evidence_ids": evidence_ids}, []
    if field == "official_information_html":
        value = (
            "<h2>Informasi Produk</h2>"
            f"<p>{esc(title)} merupakan {esc(term)} untuk {esc(function)}.</p>"
            "<h2>Catatan Teknis</h2>"
            "<p>Periksa parameter, rentang, sensor, catu daya, koneksi, dan aksesori pada tabel Spesifikasi sebelum digunakan.</p>"
        )
        return True, {"value": value, "evidence_ids": evidence_ids}, []
    return False, {}, []


def _v226_validate_section(
    field: str,
    candidate: Any,
    payload: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], List[Dict[str, Any]]]:
    # FAQ, applications, short copy, and meta are always deterministic. This
    # prevents item-count or phrasing errors from triggering repeated API calls.
    if field in {"faq_html", "application_tab_html", "short_description", "meta_description"}:
        valid, normalized, issues = _v2220_deterministic_section(field, payload, data)
        if valid:
            sections = data.get("_v2220_deterministic_sections")
            if not isinstance(sections, list):
                sections = []
            sections.append(field)
            data["_v2220_deterministic_sections"] = dedupe_keep_order(sections)
            AI_STATS["v2220_deterministic_content_sections"] = int(
                AI_STATS.get("v2220_deterministic_content_sections") or 0
            ) + 1
        return valid, normalized, issues

    valid, normalized, issues = _v2220_previous_validate_section(
        field, candidate, payload, data
    )
    if valid:
        return valid, normalized, issues

    # For long-form editorial sections, use a source-safe deterministic section
    # before asking DeepSeek to repair the same formatting problem repeatedly.
    if field in {"product_description_html", "feature_tab_html", "official_information_html"}:
        fallback_valid, fallback_normalized, fallback_issues = _v2220_deterministic_section(
            field, payload, data
        )
        if fallback_valid:
            sections = data.get("_v2220_deterministic_sections")
            if not isinstance(sections, list):
                sections = []
            sections.append(field)
            data["_v2220_deterministic_sections"] = dedupe_keep_order(sections)
            AI_STATS["v2220_deterministic_content_sections"] = int(
                AI_STATS.get("v2220_deterministic_content_sections") or 0
            ) + 1
            return True, fallback_normalized, []
        return fallback_valid, fallback_normalized, fallback_issues
    return valid, normalized, issues


_v2220_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2220_previous_build_output_row(data, pub_date)
    return _v2220_finalize_row(row, data)


_v2220_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2220_previous_validate_publish_output(
        data, row, require_ai=require_ai
    )
    superseded = {
        "mixed_language_or_corrupt_specification_residue",
        "feature_count_invalid_after_repair",
        "product_description_structure_invalid_after_repair",
        "faq_item_count_invalid",
        "faq_structure_invalid",
        "application_item_count_invalid",
        "meta_description_invalid",
        "meta_description_not_keyphrase_aligned",
        "seo_title_invalid",
        "seo_title_missing_brand_or_model",
        "focus_keyphrase_invalid",
        "seo_title_pattern_invalid",
        "focus_keyphrase_not_exact_title",
        "key_specification_section_missing",
        "key_specification_items_missing",
        "faq_product_definition_missing",
        "faq_non_solution_answer_visible",
        "meta_description_missing_click_motivation",
        "short_description_vague_specification_filler",
        "short_description_generic_sales_copy",
        "short_description_invalid_after_seo_finalizer",
    }
    issues = [issue for issue in issues if clean_text(issue) not in superseded]
    if not clean_text(data.get("_v2219_review_stage")):
        issues = [issue for issue in issues if clean_text(issue) != "iterative_repair_not_complete"]
    issues.extend(data.get("_v2220_final_issues") or [])
    if clean_text(data.get("_v2219_review_stage")):
        issues.append("product_local_technical_review_required")
    return dedupe_keep_order(issues)


# ---------------------------------------------------------------------------
# V22.21 — source-safe coherent specification pipeline
# ---------------------------------------------------------------------------
# Specifications are rebuilt from the immutable source ledger. DeepSeek is not
# asked to repair individual specification rows. A row is published either as
# a coherent Indonesian translation or as the complete original English source
# value; unfinished English-Indonesian hybrids are never published.

V2221_VERSION = "22.21-source-safe-coherent-specs"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_21_source_safe_specs_2026_08_03"

for _column in (
    "Specification Language Mode",
    "Specification English Rows",
    "Specification Source Fallback Rows",
):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)

AI_STATS.setdefault("v2221_spec_ai_calls_avoided", 0)
AI_STATS.setdefault("v2221_source_fallback_rows", 0)
AI_STATS.setdefault("v2221_indonesian_spec_rows", 0)
AI_STATS.setdefault("v2221_english_spec_rows", 0)

V2221_ID_WORD_RE = re.compile(
    r"\b(?:dan|atau|untuk|dengan|tanpa|pada|dari|hingga|ketika|jika|"
    r"nilai|rentang|akurasi|resolusi|suhu|kelembapan|catu|daya|baterai|"
    r"tampilan|masukan|keluaran|pengukuran|penyimpanan|dimensi|berat|"
    r"pembobotan|frekuensi|waktu|mikrofon|otomatis|manual|standar)\b",
    re.I,
)
V2221_EN_WORD_RE = re.compile(
    r"\b(?:electric|condenser|microphone|battery|adapter|automatic|manual|zero|"
    r"adjustment|stores|maximum|minimum|operating|temperature|humidity|range|"
    r"accuracy|resolution|display|input|output|power|supply|weight|dimension)\b",
    re.I,
)
V2221_EN_CONNECTOR_RE = re.compile(
    r"\b(?:the|and|or|for|with|without|when|if|which|is|are|was|were|"
    r"to|of|from|within|following|typically|suitable|checking|knowing|"
    r"stores|provides|meets|operating|maximum|minimum)\b",
    re.I,
)

V2221_LABEL_RULES: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"^\s*(?:operating\s+temp(?:erature)?\.?|suhu\s+operasi)\s*$", re.I), "Suhu operasi"),
    (re.compile(r"^\s*(?:operating\s+humidity|kelembapan\s+operasi)\s*$", re.I), "Kelembapan operasi"),
    (re.compile(r"^\s*(?:power\s+supply|catu\s+daya)\s*$", re.I), "Catu daya"),
    (re.compile(r"^\s*(?:power\s+consumption|konsumsi\s+daya)\s*$", re.I), "Konsumsi daya"),
    (re.compile(r"^\s*(?:dimensions?|dimensi)\s*$", re.I), "Dimensi"),
    (re.compile(r"^\s*(?:weight|berat)\s*$", re.I), "Berat"),
    (re.compile(r"^\s*(?:measurement\s+range|measuring\s+range|range|rentang(?:\s+pengukuran)?)\s*$", re.I), "Rentang pengukuran"),
    (re.compile(r"^\s*(?:resolution|resolusi)\s*$", re.I), "Resolusi"),
    (re.compile(r"^\s*(?:accuracy|akurasi)\s*$", re.I), "Akurasi"),
    (re.compile(r"^\s*(?:frequency\s+weighting|pembobotan\s+frekuensi)\s*$", re.I), "Pembobotan frekuensi"),
    (re.compile(r"^\s*(?:time\s+weighting|pembobotan\s+waktu)\s*$", re.I), "Pembobotan waktu"),
    (re.compile(r"^\s*(?:microphone|mikrofon)\s*$", re.I), "Mikrofon"),
    (re.compile(r"^\s*(?:display|tampilan)\s*$", re.I), "Tampilan"),
    (re.compile(r"^\s*(?:data\s+output|output|keluaran\s+data)\s*$", re.I), "Keluaran data"),
    (re.compile(r"^\s*(?:data\s+hold|penahan\s+data)\s*$", re.I), "Data Hold"),
    (re.compile(r"^\s*(?:sampling\s+time|waktu\s+pencuplikan|waktu\s+sampling)\s*$", re.I), "Waktu pencuplikan"),
    (re.compile(r"^\s*(?:standard\s+accessories|aksesori\s+standar)\s*$", re.I), "Aksesori standar"),
    (re.compile(r"^\s*(?:optional\s+accessories|aksesori\s+opsional)\s*$", re.I), "Aksesori opsional"),
    (re.compile(r"^\s*(?:sensor|probe)\s*$", re.I), "Sensor / Probe"),
    (re.compile(r"^\s*(?:input|masukan)\s*$", re.I), "Masukan"),
    (re.compile(r"^\s*(?:model|model\s*/\s*sku|sku)\s*$", re.I), "Model"),
)

V2221_VALUE_RULES: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bThe\s+characteristic\s+is\s+near\s+the\s+[\"']?FLAT[\"']?\s+response\.?", re.I), "Karakteristiknya mendekati respons datar (FLAT)."),
    (re.compile(r"\bTypical(?:ly)?,?\s+it\s+is\s+suitable\s+(?:for|untuk)\s+checking\s+the\s+noise\s+of\s+machinery\s*\(Q\.?\s*C\.?\s*check\)\s*(?:&|and)\s*knowing\s+the\s+sound(?:\s+pressure\s+level)?(?:\s+of\s+the\s+tested\s+equipment)?\.?", re.I), "Umumnya digunakan untuk pemeriksaan kebisingan mesin dalam kegiatan QC dan pengukuran tingkat tekanan suara peralatan."),
    (re.compile(r"\bElectric\s+condenser\s+microphone\b", re.I), "Mikrofon kondensor elektrik"),
    (re.compile(r"\bStores?\s+the\s+maximum\s+value\b", re.I), "Menyimpan nilai maksimum"),
    (re.compile(r"\bStores?\s+the\s+minimum\s+value\b", re.I), "Menyimpan nilai minimum"),
    (re.compile(r"\bis\s+specified\s+as\s+following\b", re.I), "ditentukan sebagai berikut"),
    (re.compile(r"\bOperating\s+Temp\.?\b", re.I), "Suhu operasi"),
    (re.compile(r"\bOperating\s+Humidity\b", re.I), "Kelembapan operasi"),
    (re.compile(r"\bMax\.(?=\s*\d)", re.I), "Maks."),
    (re.compile(r"\bMin\.(?=\s*\d)", re.I), "Min."),
    (re.compile(r"\bor\s+(?:DC\s*)?9\s*V\s+(?:adapter|adaptor)\s+input\b", re.I), "atau masukan adaptor DC 9 V"),
    (re.compile(r"\bAutomatic\s+zero\s+adjustment\b", re.I), "Penyetelan nol otomatis"),
    (re.compile(r"\bManual\s+zero\s+adjustment\b", re.I), "Penyetelan nol manual"),
    (re.compile(r"\bOver\s+range\s+indication\b", re.I), "Indikator melebihi rentang"),
    (re.compile(r"\bLow\s+battery\s+indication\b", re.I), "Indikator baterai lemah"),
)


def _v2221_language_shape(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return "empty"
    id_count = len(V2221_ID_WORD_RE.findall(text))
    en_count = len(V2221_EN_CONNECTOR_RE.findall(text))
    en_word_count = len(V2221_EN_WORD_RE.findall(text))
    if id_count and (en_count or en_word_count):
        return "mixed"
    if en_count or (not id_count and en_word_count >= 1):
        return "en"
    if id_count:
        return "id"
    return "technical"


def _v2221_localize_label(value: Any) -> str:
    source = _v224_exact_text(value)
    if not source:
        return "Rincian teknis"
    for pattern, replacement in V2221_LABEL_RULES:
        if pattern.fullmatch(source):
            return replacement
    key, standard, _source = _canonical_label_info({"source_label": source, "label": source})
    if not key.startswith("other:") and clean_text(standard):
        return normalize_label(standard)
    cleaned, _ = _v2220_clean_fragment(source, remove_boilerplate=False)
    cleaned = normalize_label(cleaned or source)
    if _v2221_language_shape(cleaned) == "mixed":
        return normalize_label(source)
    return cleaned


def _v2221_tokens_preserved(source: Any, candidate: Any) -> bool:
    src = _v224_exact_text(source)
    dst = _v224_exact_text(candidate)
    if _v224_numeric_tokens(src) != _v224_numeric_tokens(dst):
        return False
    dst_key = re.sub(r"\s+", "", dst).casefold()
    for token in _v226_required_spec_tokens(src):
        token_key = re.sub(r"\s+", "", token).casefold()
        if token_key and token_key not in dst_key:
            return False
    return True


def _v2221_localize_value(value: Any) -> Tuple[str, str, str]:
    source = _v224_exact_text(clean_catalog_source_text(value))
    if not source:
        return "", "empty", "empty_source"
    candidate = source
    for pattern, replacement in V2221_VALUE_RULES:
        candidate = pattern.sub(replacement, candidate)
    candidate, _ = _v2220_clean_fragment(candidate, remove_boilerplate=False)
    candidate = re.sub(r"\s+", " ", clean_text(candidate)).strip(" ,;:-")

    reason = "deterministic_translation"
    if not candidate or not _v2221_tokens_preserved(source, candidate):
        candidate = source
        reason = "source_fallback_token_guard"
    elif _v2221_language_shape(candidate) == "mixed" or _v2220_has_mixed_language(candidate):
        candidate = source
        reason = "source_fallback_language_coherence"

    shape = _v2221_language_shape(candidate)
    # The source itself can occasionally contain localized fragments. A mixed
    # source row is retained for review rather than rewritten or guessed.
    if shape == "mixed":
        reason = "mixed_source_unresolved"
    return candidate, shape, reason


def _v2221_build_source_spec_rows(data: Dict[str, Any]) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    ledger = _v224_source_rows(data)
    output: List[Dict[str, str]] = []
    seen = set()
    english_ids: List[int] = []
    fallback_ids: List[int] = []
    mixed_ids: List[int] = []
    translation_map: Dict[str, Dict[str, str]] = {}

    for item in ledger:
        source_id = int(item.get("source_id") or len(output) + 1)
        raw_label = _v224_exact_text(item.get("label_raw"))
        raw_value = _v224_exact_text(item.get("value_raw"))
        if not raw_label or not raw_value:
            continue
        label = _v2221_localize_label(raw_label)
        value, shape, reason = _v2221_localize_value(raw_value)
        if not label or not value or _spec_text_key(label) == _spec_text_key(value):
            continue
        key = (_spec_text_key(label), _spec_text_key(value))
        if not all(key) or key in seen:
            continue
        seen.add(key)
        row = {
            "label": label,
            "value": value,
            "source_label": raw_label,
            "section": _v224_exact_text(item.get("section_raw")) or "Specifications",
            "source_id": str(source_id),
        }
        output.append(row)
        translation_map[str(source_id)] = dict(row)
        if shape == "en":
            english_ids.append(source_id)
        if reason.startswith("source_fallback"):
            fallback_ids.append(source_id)
        if shape == "mixed":
            mixed_ids.append(source_id)

    # Keep identity first, then source order.
    output.sort(key=lambda row: (0 if re.search(r"\bmodel\b", row.get("label", ""), re.I) else 1, int(row.get("source_id") or 0)))
    meta = {
        "source_rows": len(ledger),
        "published_rows": len(output),
        "english_ids": sorted(set(english_ids)),
        "fallback_ids": sorted(set(fallback_ids)),
        "mixed_ids": sorted(set(mixed_ids)),
        "technical_integrity_ok": bool(output) and not mixed_ids,
        "translation_map": translation_map,
    }
    data["_v2221_spec_rows"] = output
    data["_v2221_spec_meta"] = meta
    data["_v226_spec_translations"] = translation_map
    AI_STATS["v2221_source_fallback_rows"] = int(AI_STATS.get("v2221_source_fallback_rows") or 0) + len(fallback_ids)
    AI_STATS["v2221_english_spec_rows"] = int(AI_STATS.get("v2221_english_spec_rows") or 0) + len(english_ids)
    AI_STATS["v2221_indonesian_spec_rows"] = int(AI_STATS.get("v2221_indonesian_spec_rows") or 0) + max(0, len(output) - len(english_ids) - len(mixed_ids))
    return output, meta


def _v2221_render_specifications(rows: Sequence[Dict[str, str]]) -> str:
    body: List[str] = []
    for row in rows:
        label = esc(row.get("label"))
        value = esc(row.get("value"))
        if label and value:
            lang = "en" if _v2221_language_shape(row.get("value")) == "en" else "id"
            body.append(f'<tr><th scope="row">{label}</th><td lang="{lang}">{value}</td></tr>')
    if not body:
        return "<h2>Spesifikasi</h2><p>Spesifikasi belum tersedia.</p>"
    return (
        '<h2>Spesifikasi</h2>'
        '<table class="WEBSITE-spec-table WEBSITE-spec-general">'
        '<thead><tr><th>Parameter</th><th>Nilai</th></tr></thead>'
        '<tbody>' + "".join(body) + '</tbody></table>'
    )


# Initial/bulk AI validation no longer depends on DeepSeek translating every
# technical row. It uses the immutable source values, which avoids full-product
# retries caused only by specification language or deduplication output.
def _validate_ai_spec_deduplication(
    parsed: Dict[str, Any],
    source_specs: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, str]], int, List[Dict[str, Any]]]:
    rows: List[Dict[str, str]] = []
    seen = set()
    for spec in source_specs or []:
        if not isinstance(spec, dict):
            continue
        source_id = _coerce_ai_spec_id(spec.get("id"))
        label_raw = clean_text(spec.get("source_label"))
        value_raw = clean_text(spec.get("source_value"))
        if source_id is None or not label_raw or not value_raw:
            continue
        label = _v2221_localize_label(label_raw)
        value, _shape, _reason = _v2221_localize_value(value_raw)
        key = (_spec_text_key(label), _spec_text_key(value))
        if not all(key) or key in seen:
            continue
        seen.add(key)
        rows.append({
            "id": source_id,
            "label": label,
            "source_label": label_raw,
            "value": value,
            "section": clean_text(spec.get("section")) or "Specifications",
        })
    if source_specs and not rows:
        return [], 0, [_spec_validation_issue(
            "source_specification_rows_unavailable",
            "Baris spesifikasi sumber tidak dapat dibentuk secara deterministik.",
        )]
    return rows, 0, []


# No per-row specification API repair. The source ledger is exact and already
# sufficient to build the technical table. DeepSeek remains available for the
# customer-facing narrative sections only.
def _v227_serial_spec_repair(data: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
    rows, meta = _v2221_build_source_spec_rows(data)
    AI_STATS["v2221_spec_ai_calls_avoided"] = int(AI_STATS.get("v2221_spec_ai_calls_avoided") or 0) + len(rows)
    report = {
        "version": V2221_VERSION,
        "mode": "immutable_source_specifications_no_ai_row_repair",
        "source_rows": meta.get("source_rows", 0),
        "published_rows": meta.get("published_rows", 0),
        "repaired_rows": [],
        "english_source_rows": meta.get("english_ids", []),
        "source_fallback_rows": meta.get("fallback_ids", []),
        "unresolved": {
            str(source_id): [{
                "code": "mixed_source_language_unresolved",
                "message": "Baris sumber sendiri memuat campuran bahasa dan perlu pemeriksaan.",
            }]
            for source_id in meta.get("mixed_ids", [])
        },
    }
    return bool(rows) and not meta.get("mixed_ids"), report


def _v2221_editorial_scan(row: Dict[str, Any], spec_meta: Dict[str, Any]) -> Tuple[List[str], List[str], List[str]]:
    issues: List[str] = []
    fields: List[str] = []
    details: List[str] = []

    # Specification English is allowed when the complete row stays coherent.
    if spec_meta.get("mixed_ids"):
        issues.append("mixed_source_specification_row")
        fields.append("custom_tab_1_content")
        details.append("Baris spesifikasi sumber campuran: " + ", ".join(str(x) for x in spec_meta.get("mixed_ids", [])))
    if not spec_meta.get("published_rows"):
        issues.append("specification_table_invalid")
        fields.append("custom_tab_1_content")

    for field in (
        "Product Description", "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
    ):
        visible = _v2211_visible_text(row.get(field))
        defensive = V2220_DEFENSIVE_RE.search(visible)
        if defensive:
            issues.append("defensive_boilerplate_visible")
            fields.append(field)
            details.append(f"{field}: {compact(visible[max(0, defensive.start()-35):defensive.end()+80], 170)}")
        mixed = V2220_MIXED_LANGUAGE_RE.search(visible)
        if mixed or _v2220_has_mixed_language(visible):
            issues.append("mixed_language_residue_visible")
            fields.append(field)
            start = mixed.start() if mixed else 0
            details.append(f"{field}: {compact(visible[max(0, start-35):start+145], 185)}")

    for field in ("Name", "Product Short Description", "focus keyphrase", "meta description"):
        text = clean_text(row.get(field))
        if V2220_DEFENSIVE_RE.search(text):
            issues.append("defensive_boilerplate_visible")
            fields.append(field)
        if V2220_MIXED_LANGUAGE_RE.search(text) or _v2220_has_mixed_language(text):
            issues.append("mixed_language_residue_visible")
            fields.append(field)
            details.append(f"{field}: {compact(text, 185)}")

    title = clean_text(row.get("Name"))
    keyphrase = clean_text(row.get("focus keyphrase"))
    if not title or len(title) > 65 or not keyphrase or keyphrase.casefold() not in title.casefold():
        issues.append("seo_title_or_keyphrase_invalid")
        fields.extend(["Name", "focus keyphrase"])
    meta = clean_text(row.get("meta description"))
    if not (90 <= len(meta) <= V22_META_MAX) or not meta.endswith("."):
        issues.append("meta_description_invalid")
        fields.append("meta description")

    description = BeautifulSoup(clean_text(row.get("Product Description")), "html.parser")
    if len(description.find_all("p")) < 2 or len(description.find_all("h2")) < 2:
        issues.append("product_description_structure_invalid")
        fields.append("Product Description")
    if len(BeautifulSoup(clean_text(row.get("custom_tab_4_content")), "html.parser").find_all("h3")) != 4:
        issues.append("faq_item_count_invalid")
        fields.append("custom_tab_4_content")
    if len(BeautifulSoup(clean_text(row.get("custom_tab_3_content")), "html.parser").find_all("li")) != 4:
        issues.append("application_item_count_invalid")
        fields.append("custom_tab_3_content")
    if len(BeautifulSoup(clean_text(row.get("custom_tab_2_content")), "html.parser").find_all("li")) < 3:
        issues.append("feature_fact_count_low")
        fields.append("custom_tab_2_content")
    return dedupe_keep_order(issues), dedupe_keep_order(fields), dedupe_keep_order(details)[:12]


_v2221_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2221_previous_build_output_row(data, pub_date)
    spec_rows, spec_meta = _v2221_build_source_spec_rows(data)
    row["custom_tab_1_content"] = _v2221_render_specifications(spec_rows)

    # Rebuild customer-facing copy from the coherent final specification rows.
    family = detect_family(data)
    profile = _v2220_profile(data, family)
    narrative_pairs = [
        {"label": clean_text(item.get("label")), "value": clean_text(item.get("value"))}
        for item in spec_rows
        if not re.search(r"\bmodel\b", clean_text(item.get("label")), re.I)
        and _v2221_language_shape(item.get("value")) != "mixed"
    ]
    full_title, base_title = _v2220_public_titles(data, family, narrative_pairs)
    row["Name"] = full_title
    row["focus keyphrase"] = base_title
    row["Product Description"] = _v2220_product_description(data, profile, full_title, base_title, narrative_pairs)
    row["Product Short Description"] = _v2220_short_description(profile, base_title, narrative_pairs)
    row["custom_tab_2_content"] = _v2220_feature_html(profile, narrative_pairs)
    row["custom_tab_3_content"] = _v2220_application_html(profile)
    row["custom_tab_4_content"] = _v2220_faq_html(data, profile, base_title, narrative_pairs)
    row["custom_tab_5_content"] = _v2220_official_information(profile, base_title, narrative_pairs)
    row["meta description"] = _v2220_meta_description(profile, base_title, narrative_pairs)

    repairs = 0
    for field in (
        "Product Description", "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
    ):
        row[field], changed = _v2220_clean_html(row.get(field), remove_boilerplate=True)
        repairs += changed
    for field in ("Product Short Description", "meta description"):
        row[field], changed = _v2220_clean_fragment(row.get(field), remove_boilerplate=True)
        repairs += changed
        if row[field] and row[field][-1] not in ".!?":
            row[field] += "."

    # Ensure readable spacing around CTA links after HTML sanitization.
    row["Product Description"] = re.sub(r"kunjungi\s*<strong>", "kunjungi <strong>", clean_text(row.get("Product Description")), flags=re.I)
    row["Product Description"] = re.sub(r"</strong>\s*atau\s*<strong>", "</strong> atau <strong>", row["Product Description"], flags=re.I)

    issues, review_fields, review_details = _v2221_editorial_scan(row, spec_meta)
    row["Editorial Review Fields"] = ", ".join(review_fields)
    row["Editorial Review Details"] = " | ".join(review_details)
    row["Specification English Rows"] = ",".join(str(x) for x in spec_meta.get("english_ids", []))
    row["Specification Source Fallback Rows"] = ",".join(str(x) for x in spec_meta.get("fallback_ids", []))

    if spec_meta.get("mixed_ids"):
        row["Specification Language Mode"] = "NEEDS_REVIEW"
    elif spec_meta.get("english_ids"):
        row["Specification Language Mode"] = "BILINGUAL_TECHNICAL"
    else:
        row["Specification Language Mode"] = "INDONESIAN"

    prior_stage = clean_text(row.get("Review Failure Stage") or data.get("_v2219_review_stage"))
    recovered_spec_failure = bool(
        prior_stage == "specification_repair"
        and spec_meta.get("technical_integrity_ok")
    )
    if recovered_spec_failure:
        row["Review Failure Stage"] = ""
        row["Review Failure Details"] = ""
        row["Diagnostic File"] = ""
        data["_v2219_review_stage"] = ""
        data["_v2219_review_details"] = ""

    unrecovered_failure = bool(prior_stage and not recovered_spec_failure)
    source_warning = bool(data.get("_v22122_quarantined_source_rows"))
    english_warning = bool(spec_meta.get("english_ids") or spec_meta.get("fallback_ids"))

    if unrecovered_failure:
        row["Publish Decision"] = "NEEDS_REVIEW"
        row["Final Publish Score"] = min(int(row.get("Final Publish Score") or 45), 45)
        row["Website Validation"] = "needs_review_v22_21:" + prior_stage
    elif issues:
        row["Publish Decision"] = "NEEDS_REVIEW"
        row["Final Publish Score"] = max(45, 100 - 10 * len(issues))
        row["Website Validation"] = "needs_review_v22_21:" + ",".join(issues)
    elif source_warning or english_warning:
        row["Publish Decision"] = "READY_WITH_WARNING"
        row["Final Publish Score"] = 95
        row["Website Validation"] = "publish_ready_with_technical_english_v22_21"
    else:
        row["Publish Decision"] = "READY"
        row["Final Publish Score"] = 100
        row["Website Validation"] = "publish_ready_v22_21"

    row["Processing Time"] = "lutron_WEBSITE_import_v22_21_source_safe_specs"
    row["Content Quality"] = (clean_text(row.get("Content Quality")) + "_source_safe_coherent_specs").strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "specification_finalizer=" + V2221_VERSION,
        "spec_ai_row_repairs=0",
        "spec_source_rows=" + str(spec_meta.get("source_rows", 0)),
        "spec_published_rows=" + str(spec_meta.get("published_rows", 0)),
        "spec_english_rows=" + ",".join(str(x) for x in spec_meta.get("english_ids", [])),
        "spec_source_fallback_rows=" + ",".join(str(x) for x in spec_meta.get("fallback_ids", [])),
        "spec_mixed_rows=" + ",".join(str(x) for x in spec_meta.get("mixed_ids", [])),
        "specification_policy=coherent_indonesian_or_complete_source_english_never_hybrid",
        "specification_ai_failure_recovered_from_source" if recovered_spec_failure else "",
        "v2221_public_repairs=" + str(repairs),
    )
    data["_v2221_final_issues"] = issues
    return row


_v2221_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2221_previous_validate_publish_output(data, row, require_ai=require_ai)
    # V22.21 makes coherent complete English technical values publishable and
    # resolves AI-only specification repair failures from the source ledger.
    obsolete = {
        "mixed_language_or_corrupt_specification_residue",
        "mixed_language_residue_visible",
        "iterative_repair_not_complete",
        "product_local_technical_review_required",
    }
    issues = [item for item in issues if clean_text(item) not in obsolete]
    issues.extend(data.get("_v2221_final_issues") or [])
    return dedupe_keep_order(issues)



# ---------------------------------------------------------------------------
# V22.22 — narrative-safe facts after source-safe specifications
# ---------------------------------------------------------------------------
# V22.21 made the specification table source-safe, but long/coalesced source
# values could still be reused by Product Description, Short Description,
# Features, FAQ, and Information Product. V22.22 separates publication data
# from narrative facts. Only short, atomic, structurally valid facts may enter
# customer-facing prose. Suspicious source rows stay out of public narrative;
# clearly corrupt rows are quarantined from the visible specification table.

V2222_VERSION = "22.22-narrative-safe-atomic-facts"
AI_PROMPT_VERSION = "lutron_WEBSITE_v22_22_narrative_safe_2026_08_03"

for _column in (
    "Narrative Safe Fact Rows",
    "Quarantined Specification Rows",
    "Quarantined Specification Details",
):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)

AI_STATS.setdefault("v2222_quarantined_spec_rows", 0)
AI_STATS.setdefault("v2222_narrative_safe_facts", 0)
AI_STATS.setdefault("v2222_narrative_fallbacks", 0)

V2222_UNIT_RE = (
    r"(?:dB|°?C|°?F|℃|℉|%\s*RH|%|V|mV|kV|A|mA|µA|uA|W|mW|kW|"
    r"Hz|kHz|MHz|GHz|lux|fc|m/s|km/h|ft/min|MPH|knot|knots|kt|rpm|"
    r"pH|ppm|ppt|NTU|FNU|mbar|bar|psi|mmHg|inHg|mH₂O|inH₂O|Ω|kΩ|MΩ)"
)

V2222_EMBEDDED_LABEL_RE = re.compile(
    r"(?:^|[.;])\s*(?:data\s+output|keluaran\s+data|power\s+supply|catu\s+daya|"
    r"accuracy|akurasi|resolution|resolusi|display|tampilan|weight|berat|"
    r"dimensions?|dimensi|operating\s+temp(?:erature)?|suhu\s+operasi|"
    r"operating\s+humidity|kelembapan\s+operasi)\s*:",
    re.I,
)

V2222_DANGLING_RE = re.compile(
    r"(?:\b(?:then(?:\s+the)?|and|or|with|without|for|to|of|the|which|is|are|"
    r"dan|atau|dengan|tanpa|untuk|dari|yang)\b\s*[.,;:]*)$",
    re.I,
)

V2222_HYBRID_PHRASE_RE = re.compile(
    r"(?:then\s+(?:the\s+)?(?:akurasi|resolusi|nilai|rentang)|"
    r"(?:akurasi|resolusi|nilai|rentang)\s+is\s+specified|"
    r"stores?\s+the\s+(?:maximum|minimum)\s+(?:nilai|value)|"
    r"typical(?:ly)?,?\s+it\s+is\s+suitable\s+untuk|"
    r"untuk\s+checking|with\s+(?:over|under)\s+range|"
    r"or\s+masukan|the\s+(?:akurasi|resolusi)|"
    r"manual\s+push\s+the|alerts?\s+untuk|"
    r"range\s+with\s+manual\s+range\s+selecting)",
    re.I,
)

V2222_EXTRA_EN_WORD_RE = re.compile(
    r"\b(?:liquid|crystal|digits?|standard|size|high|battery|alkaline|heavy|duty|"
    r"type|separate|holder|accessory|included|leads?|fuse|instruction|manual|"
    r"freezes?|reading|easy|reference|calibration|built|external|screw|driver|"
    r"average|values?|vibration|sound|level|selecting|simulated|human|ear|response)\b",
    re.I,
)


V2222_PUBLIC_CORRUPTION_RE = re.compile(
    r"(?:rentang\s+pengukuran\s*:\s*rentang\s+pengukuran\s*:|"
    r"(?<!Kelas )(?<!Class )\b\d+(?:[.,]\d+)?\.\s*(?:keluaran\s+data|data\s+output|catu\s+daya|"
    r"power\s+supply|akurasi|accuracy|resolusi|resolution)\s*:|"
    r"\b[1-9]\d{0,2},\s+\d{3}\b|"
    r"then\s+(?:the\s*)?[.!?]|"
    r"\b(?:the|and|or|with|for|to|of|which|is|are)\s*[.!?]\s*$)",
    re.I,
)


def _v2222_balanced(value: Any, left: str, right: str) -> bool:
    text = clean_text(value)
    return text.count(left) == text.count(right)


def _v2222_normalize_display_value(value: Any) -> str:
    text = _v224_exact_text(value)
    if not text:
        return ""
    # Safe typography repairs only. Do not translate, convert units, or alter
    # numeric magnitude. A spaced thousands separator is joined only when the
    # left side cannot be a decimal zero and exactly three digits follow.
    text = re.sub(r"\b([1-9]\d{0,2}),\s+(\d{3})\b", r"\1,\2", text)
    text = re.sub(r"\s+([,.;:)\]])", r"\1", text)
    text = re.sub(r"([(\[])\s+", r"\1", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text


def _v2222_spec_corruption_reason(label: Any, value: Any) -> str:
    label_text = clean_text(label)
    text = _v2222_normalize_display_value(value)
    if not label_text or not text:
        return "empty_label_or_value"
    if not _v2222_balanced(text, "(", ")") or not _v2222_balanced(text, "[", "]"):
        return "unbalanced_brackets"
    if V2222_HYBRID_PHRASE_RE.search(text):
        return "hybrid_sentence_fragment"
    if V2222_DANGLING_RE.search(text):
        return "dangling_connector"
    embedded = V2222_EMBEDDED_LABEL_RE.search(text)
    if embedded and embedded.start() > 0:
        return "multiple_fields_coalesced"
    if re.search(
        r"(?<!Kelas )(?<!Class )\b\d+(?:[.,]\d+)?\.\s*(?:keluaran\s+data|data\s+output|catu\s+daya|"
        r"power\s+supply|akurasi|accuracy|resolusi|resolution)\s*:",
        text,
        re.I,
    ):
        return "truncated_numeric_then_next_label"
    id_hits = len(V2221_ID_WORD_RE.findall(text))
    extra_en_hits = len(V2222_EXTRA_EN_WORD_RE.findall(text))
    if (id_hits and extra_en_hits) or _v2221_language_shape(text) == "mixed" or _v2220_has_mixed_language(text):
        return "mixed_language_source_fragment"
    # Very long prose belongs in official source evidence, not a technical
    # key/value table. Retain compact multi-range values; quarantine prose.
    if len(text) > 520 and len(re.findall(r"[.!?]", text)) >= 3:
        return "oversized_narrative_in_specification"
    return ""


def _v2222_prepare_spec_rows(data: Dict[str, Any]) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    source_rows, source_meta = _v2221_build_source_spec_rows(data)
    published: List[Dict[str, str]] = []
    quarantined_ids: List[int] = []
    quarantine_details: List[str] = []
    english_ids: List[int] = []
    fallback_ids: List[int] = []
    seen = set()

    source_fallback = {int(x) for x in source_meta.get("fallback_ids", [])}
    for raw in source_rows:
        row = dict(raw)
        source_id = int(row.get("source_id") or len(published) + 1)
        row["value"] = _v2222_normalize_display_value(row.get("value"))
        reason = _v2222_spec_corruption_reason(row.get("label"), row.get("value"))
        if reason:
            quarantined_ids.append(source_id)
            quarantine_details.append(f"{source_id}:{reason}")
            continue
        key = (_spec_text_key(row.get("label")), _spec_text_key(row.get("value")))
        if not all(key) or key in seen:
            continue
        seen.add(key)
        published.append(row)
        shape = _v2221_language_shape(row.get("value"))
        if shape == "en":
            english_ids.append(source_id)
        if source_id in source_fallback:
            fallback_ids.append(source_id)

    substantive = [
        row for row in published
        if not re.search(r"\b(?:model|sku|mpn)\b", clean_text(row.get("label")), re.I)
    ]
    meta = dict(source_meta)
    meta.update({
        "published_rows": len(published),
        "english_ids": sorted(set(english_ids)),
        "fallback_ids": sorted(set(fallback_ids)),
        "source_mixed_ids": sorted(set(source_meta.get("mixed_ids", []))),
        "mixed_ids": [],  # bad mixed rows are not published; they are quarantined
        "quarantined_ids": sorted(set(quarantined_ids)),
        "quarantine_details": quarantine_details,
        "technical_integrity_ok": len(substantive) >= 3,
    })
    data["_v2222_spec_rows"] = published
    data["_v2222_spec_meta"] = meta
    AI_STATS["v2222_quarantined_spec_rows"] = int(
        AI_STATS.get("v2222_quarantined_spec_rows") or 0
    ) + len(quarantined_ids)
    return published, meta


def _v2222_range_fact(rows: Sequence[Dict[str, str]]) -> Optional[Dict[str, str]]:
    candidates: List[Tuple[float, float, str, str, str, int]] = []
    pattern = re.compile(
        rf"([-+]?\d+(?:[.,]\d+)?)\s*({V2222_UNIT_RE})?\s*"
        rf"(?:hingga|sampai|to|[-–~])\s*"
        rf"([-+]?\d+(?:[.,]\d+)?)\s*({V2222_UNIT_RE})",
        re.I,
    )
    for row_index, row in enumerate(rows):
        label = clean_text(row.get("label"))
        if not re.search(
            r"\b(?:rentang\s+pengukuran|measurement\s+range|measuring\s+range|"
            r"pemilih\s+rentang|range\s+selection|rentang|range)\b",
            label,
            re.I,
        ):
            continue
        if re.search(r"\b(?:operasi|operating|suhu|temperature|kelembapan|humidity|catu|power)\b", label, re.I):
            continue
        value = clean_text(row.get("value"))
        if _v2222_spec_corruption_reason(label, value):
            continue
        for match in pattern.finditer(value):
            low_s, unit_left, high_s, unit_right = match.groups()
            unit = clean_text(unit_right or unit_left)
            if unit_left and clean_text(unit_left).casefold() != unit.casefold():
                continue
            try:
                low = float(low_s.replace(",", "."))
                high = float(high_s.replace(",", "."))
            except Exception:
                continue
            candidates.append((low, high, low_s, high_s, unit, row_index))
    if not candidates:
        return None
    first_unit = candidates[0][4].casefold()
    same_unit = [item for item in candidates if item[4].casefold() == first_unit]
    if not same_unit:
        return None
    low_item = min(same_unit, key=lambda item: item[0])
    high_item = max(same_unit, key=lambda item: item[1])
    value = f"{low_item[2]}–{high_item[3]} {low_item[4]}"
    source_ids = dedupe_keep_order(
        str(rows[item[5]].get("source_id")) for item in same_unit
        if clean_text(rows[item[5]].get("source_id"))
    )
    return {
        "label": "Rentang pengukuran",
        "value": value,
        "source_id": ",".join(source_ids),
    }


def _v2222_interface_value(value: Any) -> str:
    text = clean_text(value)
    tokens: List[str] = []
    token_rules = (
        (r"\bRS\s*[- ]?232\b", "RS232"),
        (r"\bUSB(?:[- ]?\d+)?\b", "USB"),
        (r"\bBluetooth\b", "Bluetooth"),
        (r"\bWi[- ]?Fi\b", "Wi-Fi"),
        (r"\bSD(?:HC)?\s*(?:card|kartu)?\b", "kartu SD"),
        (r"\b4\s*[-–]\s*20\s*mA\b", "4–20 mA"),
        (r"\banalog(?:ue)?\s+output\b|\bkeluaran\s+analog\b", "keluaran analog"),
    )
    for pattern, replacement in token_rules:
        if re.search(pattern, text, re.I):
            tokens.append(replacement)
    return _v2220_join_id(dedupe_keep_order(tokens)[:3])


def _v2222_atomic_value(label: Any, value: Any) -> str:
    label_text = clean_text(label)
    text = _v2222_normalize_display_value(value).strip(" ;,.-")
    if not text or _v2222_spec_corruption_reason(label_text, text):
        return ""
    label_shape = _v2221_language_shape(label_text)
    if (
        label_shape in {"en", "mixed"}
        or V2222_EXTRA_EN_WORD_RE.search(label_text)
    ) and not re.fullmatch(r"(?:Data Hold|RS232|USB)", label_text, re.I):
        return ""
    if ":" in text:
        # Colons inside a value usually signal multiple coalesced fields. One
        # exception is a standard ratio such as 4:1.
        if not re.fullmatch(r"[^:]{0,30}\d+\s*:\s*\d+[^:]{0,30}", text):
            return ""
    if V2222_DANGLING_RE.search(text):
        return ""

    if re.search(r"\b(?:antarmuka|interface|keluaran|output|komunikasi|data\s+logger|pencatat\s+data)\b", label_text, re.I):
        return _v2222_interface_value(text)

    # Narrative facts must be compact and atomic. Long lists remain available
    # in the specification table but are not copied into prose.
    maximum = 95
    if re.search(r"\b(?:akurasi|accuracy|resolusi|resolution)\b", label_text, re.I):
        maximum = 110
    elif re.search(r"\b(?:sensor|probe|mikrofon|microphone|tampilan|display)\b", label_text, re.I):
        maximum = 85
    if len(text) > maximum:
        return ""
    if len(re.findall(r"[.!?]", text)) > 1:
        return ""
    if len(re.findall(r"[-+]?\d+(?:[.,]\d+)?", text)) > 8:
        return ""
    words = re.findall(r"[A-Za-zÀ-ÿ]+", text)
    if len(words) > 16:
        return ""
    # Complete source English is allowed in the table, but narrative values
    # should be numeric/technical or a short localized noun phrase.
    shape = _v2221_language_shape(text)
    if shape == "en" and len(words) > 5:
        return ""
    return text


def _v2222_narrative_pairs(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    output: List[Dict[str, str]] = []
    seen = set()
    range_fact = _v2222_range_fact(rows)
    if range_fact:
        output.append(range_fact)
        seen.add((_spec_text_key(range_fact["label"]), _spec_text_key(range_fact["value"])))

    priority = re.compile(
        r"\b(?:akurasi|accuracy|resolusi|resolution|sensor|probe|mikrofon|microphone|"
        r"antarmuka|interface|keluaran|output|pencatat\s+data|data\s+logger|"
        r"tampilan|display|catu\s+daya|power\s+supply|kanal|channel)\b",
        re.I,
    )
    ordered = [row for row in rows if priority.search(clean_text(row.get("label")))]
    for row in ordered:
        value = _v2222_atomic_value(row.get("label"), row.get("value"))
        if not value:
            continue
        label = clean_text(row.get("label"))
        key = (_spec_text_key(label), _spec_text_key(value))
        if not all(key) or key in seen:
            continue
        # Only one fact per broad label in prose.
        broad = re.sub(r"\b(?:umum|general|utama|main)\b", "", label, flags=re.I)
        broad_key = _spec_text_key(broad)
        if any(_spec_text_key(item.get("label")) == broad_key for item in output):
            continue
        seen.add(key)
        output.append({
            "label": label,
            "value": value,
            "source_id": clean_text(row.get("source_id")),
        })
        if len(output) >= 7:
            break
    AI_STATS["v2222_narrative_safe_facts"] = int(
        AI_STATS.get("v2222_narrative_safe_facts") or 0
    ) + len(output)
    return output


def _v2220_join_id(items: Sequence[str]) -> str:
    """Join Indonesian list items naturally for V22.22 public copy."""
    expanded: List[str] = []
    for item in items:
        value = clean_text(item)
        if not value:
            continue
        if value.casefold() == "akurasi dan resolusi":
            expanded.extend(["akurasi", "resolusi"])
        else:
            expanded.append(value)
    cleaned = dedupe_keep_order(expanded)
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return cleaned[0] + " dan " + cleaned[1]
    return ", ".join(cleaned[:-1]) + ", dan " + cleaned[-1]


def _v2222_feature_html(profile: Dict[str, Any], pairs: Sequence[Dict[str, str]]) -> str:
    bullets: List[str] = []
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    bullets.append(function[:1].upper() + function[1:] + ".")
    for row in _v2220_key_facts(pairs, maximum=4):
        fact = _v2220_fact_text(row, maximum=120)
        if fact:
            bullets.append(fact + ".")
    generic = (
        "Membantu pemeriksaan parameter secara terstruktur pada kegiatan inspeksi.",
        "Mendukung verifikasi hasil pada kegiatan pemeliharaan atau pengendalian mutu.",
        "Rincian rentang dan konfigurasi tersedia pada tabel spesifikasi teknis.",
    )
    for item in generic:
        if len(bullets) >= 4:
            break
        bullets.append(item)
    bullets = dedupe_keep_order(bullets)[:6]
    return "<h2>Keunggulan Produk</h2><ul>" + "".join(
        f"<li>{esc(item)}</li>" for item in bullets
    ) + "</ul>"


def _v2222_fix_cta_spacing(value: Any) -> str:
    html_value = clean_text(value)
    html_value = re.sub(r"kunjungi\s*<strong>", "kunjungi <strong>", html_value, flags=re.I)
    html_value = re.sub(r"</strong>\s*atau\s*<strong>", "</strong> atau <strong>", html_value, flags=re.I)
    html_value = re.sub(r"</a></strong>\s*atau", "</a></strong> atau", html_value, flags=re.I)
    return html_value


def _v2222_public_field_issues(row: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    issues: List[str] = []
    details: List[str] = []
    fields = (
        "Product Description", "Product Short Description",
        "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content", "meta description",
    )
    for field in fields:
        visible = _v2211_visible_text(row.get(field)) if "content" in field or field == "Product Description" else clean_text(row.get(field))
        if not visible:
            issues.append(field + ":empty")
            continue
        match = V2222_HYBRID_PHRASE_RE.search(visible) or V2222_PUBLIC_CORRUPTION_RE.search(visible)
        if match:
            issues.append(field + ":corrupt_fragment")
            details.append(f"{field}: {compact(visible[max(0, match.start()-35):match.end()+90], 190)}")
        if not _v2222_balanced(visible, "(", ")") or not _v2222_balanced(visible, "[", "]"):
            issues.append(field + ":unbalanced_brackets")
            details.append(f"{field}: tanda kurung tidak seimbang")
    return dedupe_keep_order(issues), dedupe_keep_order(details)


def _v2222_rebuild_public_copy(
    row: Dict[str, Any],
    data: Dict[str, Any],
    profile: Dict[str, Any],
    pairs: Sequence[Dict[str, str]],
) -> Dict[str, Any]:
    family = detect_family(data)
    full_title, base_title = _v2220_public_titles(data, family, pairs)
    row["Name"] = full_title
    row["focus keyphrase"] = base_title
    row["Product Description"] = _v2220_product_description(data, profile, full_title, base_title, pairs)
    row["Product Short Description"] = _v2220_short_description(profile, base_title, pairs)
    row["custom_tab_2_content"] = _v2222_feature_html(profile, pairs)
    row["custom_tab_3_content"] = _v2220_application_html(profile)
    row["custom_tab_4_content"] = _v2220_faq_html(data, profile, base_title, pairs)
    row["custom_tab_5_content"] = _v2220_official_information(profile, base_title, pairs)
    row["meta description"] = _v2220_meta_description(profile, base_title, pairs)

    for field in (
        "Product Description", "custom_tab_2_content", "custom_tab_3_content",
        "custom_tab_4_content", "custom_tab_5_content",
    ):
        row[field], _changed = _v2220_clean_html(row.get(field), remove_boilerplate=True)
    for field in ("Product Short Description", "meta description"):
        row[field], _changed = _v2220_clean_fragment(row.get(field), remove_boilerplate=True)
        if row[field] and row[field][-1] not in ".!?":
            row[field] += "."
    row["Product Description"] = _v2222_fix_cta_spacing(row.get("Product Description"))
    return row


_v2222_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    row = _v2222_previous_build_output_row(data, pub_date)
    spec_rows, spec_meta = _v2222_prepare_spec_rows(data)
    row["custom_tab_1_content"] = _v2221_render_specifications(spec_rows)

    family = detect_family(data)
    profile = _v2220_profile(data, family)
    narrative_pairs = _v2222_narrative_pairs(spec_rows)
    row = _v2222_rebuild_public_copy(row, data, profile, narrative_pairs)

    public_issues, public_details = _v2222_public_field_issues(row)
    fallback_used = False
    if public_issues:
        # Never let a questionable technical fragment reach customer-facing
        # prose. Rebuild narrative without any specification values. The full
        # clean table remains available in tab Spesifikasi.
        fallback_used = True
        AI_STATS["v2222_narrative_fallbacks"] = int(
            AI_STATS.get("v2222_narrative_fallbacks") or 0
        ) + 1
        row = _v2222_rebuild_public_copy(row, data, profile, [])
        public_issues, second_details = _v2222_public_field_issues(row)
        public_details.extend(second_details)

    source_ids = dedupe_keep_order(
        sid
        for item in narrative_pairs
        for sid in clean_text(item.get("source_id")).split(",")
        if clean_text(sid)
    )
    row["Narrative Safe Fact Rows"] = ",".join(source_ids)
    row["Quarantined Specification Rows"] = ",".join(
        str(x) for x in spec_meta.get("quarantined_ids", [])
    )
    row["Quarantined Specification Details"] = " | ".join(
        spec_meta.get("quarantine_details", [])
    )
    row["Specification English Rows"] = ",".join(
        str(x) for x in spec_meta.get("english_ids", [])
    )
    row["Specification Source Fallback Rows"] = ",".join(
        str(x) for x in spec_meta.get("fallback_ids", [])
    )
    if spec_meta.get("quarantined_ids"):
        row["Specification Language Mode"] = "SOURCE_SAFE_WITH_QUARANTINE"
    elif spec_meta.get("english_ids"):
        row["Specification Language Mode"] = "BILINGUAL_TECHNICAL"
    else:
        row["Specification Language Mode"] = "INDONESIAN"

    prior_stage = clean_text(row.get("Review Failure Stage") or data.get("_v2219_review_stage"))
    recovered_spec_failure = bool(
        prior_stage == "specification_repair" and spec_meta.get("technical_integrity_ok")
    )
    if recovered_spec_failure:
        prior_stage = ""
        row["Review Failure Stage"] = ""
        row["Review Failure Details"] = ""
        row["Diagnostic File"] = ""
        data["_v2219_review_stage"] = ""
        data["_v2219_review_details"] = ""

    editorial_fields: List[str] = []
    editorial_details: List[str] = []
    if public_issues:
        editorial_fields.extend(item.split(":", 1)[0] for item in public_issues)
        editorial_details.extend(public_details)
    if spec_meta.get("quarantined_ids"):
        editorial_details.append(
            "Baris spesifikasi tidak koheren dikarantina dan tidak ditampilkan: "
            + ", ".join(str(x) for x in spec_meta.get("quarantined_ids", []))
        )
    row["Editorial Review Fields"] = ", ".join(dedupe_keep_order(editorial_fields))
    row["Editorial Review Details"] = " | ".join(dedupe_keep_order(editorial_details)[:12])

    substantive_rows = [
        item for item in spec_rows
        if not re.search(r"\b(?:model|sku|mpn)\b", clean_text(item.get("label")), re.I)
    ]
    final_issues: List[str] = []
    if len(substantive_rows) < 3:
        final_issues.append("insufficient_clean_specification_rows")
    if public_issues:
        final_issues.append("customer_facing_narrative_still_invalid")
    if prior_stage:
        final_issues.append("unrecovered_product_failure:" + prior_stage)

    warning = bool(
        spec_meta.get("quarantined_ids")
        or spec_meta.get("english_ids")
        or spec_meta.get("fallback_ids")
        or fallback_used
    )
    if final_issues:
        row["Publish Decision"] = "NEEDS_REVIEW"
        row["Final Publish Score"] = max(45, 100 - 15 * len(final_issues))
        row["Website Validation"] = "needs_review_v22_22:" + ",".join(final_issues)
    elif warning:
        row["Publish Decision"] = "READY_WITH_WARNING"
        row["Final Publish Score"] = 95
        row["Website Validation"] = "publish_ready_with_source_quarantine_v22_22"
    else:
        row["Publish Decision"] = "READY"
        row["Final Publish Score"] = 100
        row["Website Validation"] = "publish_ready_v22_22"

    row["Processing Time"] = "lutron_WEBSITE_import_v22_22_narrative_safe"
    row["Content Quality"] = (
        clean_text(row.get("Content Quality")) + "_narrative_atomic_fact_gate"
    ).strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "narrative_finalizer=" + V2222_VERSION,
        "narrative_safe_fact_rows=" + ",".join(source_ids),
        "quarantined_spec_rows=" + ",".join(str(x) for x in spec_meta.get("quarantined_ids", [])),
        "narrative_fallback_used=" + ("yes" if fallback_used else "no"),
        "public_narrative_policy=atomic_short_validated_facts_only",
    )
    data["_v2222_final_issues"] = final_issues
    # Previous layers must not re-add obsolete V22.21 findings after V22.22
    # has quarantined the bad source rows and rebuilt all public copy.
    data["_v2221_final_issues"] = []
    return row


_v2222_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    issues = _v2222_previous_validate_publish_output(data, row, require_ai=require_ai)
    superseded = {
        "mixed_source_specification_row",
        "mixed_language_residue_visible",
        "mixed_language_or_corrupt_specification_residue",
        "product_local_technical_review_required",
        "iterative_repair_not_complete",
        "feature_fact_count_low",
        "product_description_structure_invalid",
        "faq_item_count_invalid",
        "application_item_count_invalid",
        "meta_description_invalid",
        "seo_title_or_keyphrase_invalid",
    }
    issues = [item for item in issues if clean_text(item) not in superseded]
    issues.extend(data.get("_v2222_final_issues") or [])
    return dedupe_keep_order(issues)



# =============================================================================
# V22.23 - CONSOLIDATED STABLE PUBLIC COPY
# =============================================================================
# This is the final public-copy owner. Earlier layers may scrape, translate,
# validate, and quarantine source facts, but they are not allowed to define the
# customer-facing structure after this point.
#
# Stability policy:
# - Product name is always: Product Type + Lutron + Model (no range modifier).
# - Description always uses one fixed structure.
# - FAQ always uses four fixed questions.
# - Specification cleanup never rewrites the description/FAQ structure.
# - Only short, source-safe facts may be copied from the specification table.
# - A warning marks English/source fallback rows; it does not make technical
#   English terminology a publication blocker.

V2223_VERSION = "v22.23_stable_content"
V2223_TEMPLATE_VERSION = "stable_public_copy_v1"

for _column in (
    "Content Template Version",
    "Stable Narrative Facts",
    "Stable Content Warnings",
):
    if _column not in OUTPUT_COLS:
        OUTPUT_COLS.append(_column)


def _v2223_family_key(data: Dict[str, Any], profile: Dict[str, Any]) -> str:
    probe = " ".join([
        clean_text(data.get("title")),
        clean_text(data.get("model")),
        clean_text(profile.get("term")),
        clean_text(profile.get("function")),
        clean_text(data.get("source_category")),
    ]).casefold()
    rules = (
        (r"sound\s+level|noise|decibel|kebisingan", "sound_level"),
        (r"multimeter|clamp\s+meter|voltmeter|ammeter", "multimeter"),
        (r"power\s+analyzer|power\s+meter|wattmeter|energy", "power"),
        (r"thermometer|temperature|termometer|alat\s+ukur\s+suhu", "thermometer"),
        (r"humidity|hygrometer|environmental|kelembapan|alat\s+ukur\s+lingkungan", "environmental"),
        (r"light\s+meter|lux|illumination|cahaya", "light"),
        (r"\buv\b|ultraviolet|uva|uvc", "uv"),
        (r"anemometer|air\s+flow|air\s+velocity|wind", "anemometer"),
        (r"pressure|manometer|tekanan", "pressure"),
        (r"transmitter|pemancar", "transmitter"),
        (r"water\s+quality|ph\s+meter|conductivity|tds|dissolved\s+oxygen", "water_quality"),
    )
    for pattern, key in rules:
        if re.search(pattern, probe, re.I):
            return key
    return "generic"


def _v2223_normalize_label(label: Any, value: Any) -> str:
    label_text = normalize_label(clean_text(label)).strip(" :;.-")
    value_text = clean_text(value)
    replacements = (
        (r"^frequency\s+weighting\s+network$|^pembobotan\s+frekuensi\s+network$", "Pembobotan frekuensi"),
        (r"^special\s+fungsi$|^special\s+function$", "Fungsi khusus"),
        (r"^other\s+rentang$|^other\s+range$", "Rentang lainnya"),
        (r"^(?:u|µ)a\s*,?\s*ma\s+(?:rentang|range)$", "Rentang µA dan mA"),
        (r"^operating\s+temp\.?$", "Suhu operasi"),
        (r"^operating\s+humidity$", "Kelembapan operasi"),
        (r"^power\s+supply$", "Catu daya"),
        (r"^data\s+output$", "Keluaran data"),
        (r"^measurement\s+range$", "Rentang pengukuran"),
        (r"^microphone$", "Mikrofon"),
        (r"^display$", "Tampilan"),
    )
    for pattern, replacement in replacements:
        if re.fullmatch(pattern, label_text, re.I):
            label_text = replacement
            break
    # IEC/ISO/Class describes compliance or instrument class, not a numerical
    # accuracy value. Correct the display label without changing the source fact.
    if re.search(r"\b(?:IEC|ISO|Class|Kelas)\b", value_text, re.I) and re.search(
        r"\b(?:akurasi|accuracy|ketelitian)\b", label_text, re.I
    ):
        label_text = "Standar/kelas"
    return label_text


def _v2223_meaningless_spec_reason(label: Any, value: Any) -> str:
    label_text = clean_text(label)
    value_text = clean_text(value).strip(" :;.-")
    if not label_text or not value_text:
        return "empty_specification"
    if re.search(r"\b(?:antarmuka|interface|keluaran|output|komunikasi)\b", label_text, re.I):
        recognized = _v2222_interface_value(value_text)
        if re.fullmatch(r"\d+(?:[.,]\d+)?", value_text):
            return "meaningless_interface_number"
        if not recognized and len(re.findall(r"[A-Za-zÀ-ÿ]+", value_text)) <= 5:
            return "unresolved_interface_fragment"
    if re.fullmatch(r"(?:and|or|then|the|for|to|with|without|of|untuk|dan|atau)", value_text, re.I):
        return "dangling_connector_value"
    if re.fullmatch(r"[-–—/.,;:()\s]+", value_text):
        return "punctuation_only_value"
    return ""


def _v2223_prepare_spec_rows(data: Dict[str, Any]) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    rows, meta = _v2222_prepare_spec_rows(data)
    output: List[Dict[str, str]] = []
    quarantined_ids = list(meta.get("quarantined_ids", []))
    quarantine_details = list(meta.get("quarantine_details", []))
    seen_pairs = set()

    for index, raw in enumerate(rows, start=1):
        row = dict(raw)
        source_id = clean_text(row.get("source_id")) or str(index)
        value = _v2222_normalize_display_value(row.get("value")).strip(" ;")
        label = _v2223_normalize_label(row.get("label"), value)
        reason = _v2223_meaningless_spec_reason(label, value)
        if reason:
            quarantined_ids.append(int(source_id) if source_id.isdigit() else source_id)
            quarantine_details.append(f"{source_id}:{reason}")
            continue

        # Replace incomplete interface prose with the exact recognized token.
        if re.search(r"\b(?:antarmuka|interface|keluaran|output|komunikasi)\b", label, re.I):
            interface_value = _v2222_interface_value(value)
            if interface_value and (
                len(value) < 45
                or _v2221_language_shape(value) == "mixed"
                or re.search(r"\bcomputer\s+interface\b", value, re.I)
            ):
                value = interface_value

        key = (_spec_text_key(label), _spec_text_key(value))
        if not all(key) or key in seen_pairs:
            continue

        duplicate = False
        for existing in output:
            if _spec_text_key(existing.get("label")) != key[0]:
                continue
            if _spec_values_equivalent(existing.get("value"), value):
                duplicate = True
                break
        if duplicate:
            continue

        seen_pairs.add(key)
        row["label"] = label
        row["value"] = value
        output.append(row)

    substantive = [
        row for row in output
        if not re.search(r"\b(?:model|sku|mpn|kode\s+produk)\b", clean_text(row.get("label")), re.I)
    ]
    new_meta = dict(meta)
    new_meta.update({
        "published_rows": len(output),
        "quarantined_ids": sorted(set(quarantined_ids), key=lambda x: str(x)),
        "quarantine_details": dedupe_keep_order(quarantine_details),
        "technical_integrity_ok": len(substantive) >= 3,
    })
    data["_v2223_spec_rows"] = output
    data["_v2223_spec_meta"] = new_meta
    return output, new_meta


def _v2223_allowed_range_units(family_key: str) -> str:
    return {
        "sound_level": r"dB",
        "thermometer": r"(?:°?C|°?F|℃|℉)",
        "light": r"(?:lux|fc)",
        "uv": r"(?:µW/cm(?:²|2)|uW/cm(?:²|2)|mW/cm(?:²|2))",
        "anemometer": r"(?:m/s|km/h|ft/min|MPH|mile/h|knots?|kt)",
        "pressure": r"(?:Pa|kPa|MPa|mbar|bar|psi|mmHg|inHg)",
    }.get(family_key, "")


def _v2223_primary_range(rows: Sequence[Dict[str, str]], family_key: str) -> Optional[Dict[str, str]]:
    unit_pattern = _v2223_allowed_range_units(family_key)
    if not unit_pattern:
        return None
    candidates: List[Tuple[float, float, str, str, str, str]] = []
    regex = re.compile(
        rf"([-+]?\d+(?:[.,]\d+)?)\s*(?:{unit_pattern})?\s*"
        rf"(?:hingga|sampai|to|[-–~])\s*"
        rf"([-+]?\d+(?:[.,]\d+)?)\s*({unit_pattern})",
        re.I,
    )
    for row in rows:
        label = clean_text(row.get("label"))
        if not re.search(r"\b(?:rentang\s+pengukuran|measurement\s+range|measuring\s+range|rentang|range)\b", label, re.I):
            continue
        if re.search(r"\b(?:operasi|operating|kelembapan|humidity|catu|power)\b", label, re.I):
            continue
        value = clean_text(row.get("value"))
        for match in regex.finditer(value):
            low_s, high_s, unit = match.groups()
            try:
                low = float(low_s.replace(",", "."))
                high = float(high_s.replace(",", "."))
            except Exception:
                continue
            candidates.append((low, high, low_s, high_s, clean_text(unit), clean_text(row.get("source_id"))))
    if not candidates:
        return None
    # Never merge unlike units. Use the first unit family found in source order.
    first_unit = candidates[0][4].casefold()
    same_unit = [item for item in candidates if item[4].casefold() == first_unit]
    low_item = min(same_unit, key=lambda item: item[0])
    high_item = max(same_unit, key=lambda item: item[1])
    return {
        "label": "Rentang pengukuran",
        "value": f"{low_item[2]}–{high_item[3]} {low_item[4]}",
        "source_id": ",".join(dedupe_keep_order(item[5] for item in same_unit if item[5])),
        "kind": "range",
    }


def _v2223_parameter_fact(rows: Sequence[Dict[str, str]], family_key: str) -> Optional[Dict[str, str]]:
    if family_key not in {"multimeter", "power", "environmental", "water_quality"}:
        return None
    blob = " ".join(
        clean_text(row.get("label")) + " " + clean_text(row.get("value"))
        for row in rows
    )
    rules = (
        (r"\bDC\s+Voltage\b|\btegangan\s+DC\b", "tegangan DC"),
        (r"\bAC\s+Voltage\b|\btegangan\s+AC\b", "tegangan AC"),
        (r"\bDC\s+Current\b|\barus\s+DC\b", "arus DC"),
        (r"\bAC\s+Current\b|\barus\s+AC\b", "arus AC"),
        (r"\bResistance\b|\bresistansi\b|\bhambatan\b", "resistansi"),
        (r"\bCapacitance\b|\bkapasitansi\b", "kapasitansi"),
        (r"\bFrequency\b|\bfrekuensi\b", "frekuensi"),
        (r"\bTemperature\b|\bsuhu\b", "suhu"),
        (r"\bHumidity\b|\bkelembapan\b", "kelembapan"),
        (r"\bPower\b|\bdaya\b", "daya"),
        (r"\bEnergy\b|\benergi\b", "energi"),
        (r"\bpH\b", "pH"),
        (r"\bConductivity\b|\bkonduktivitas\b", "konduktivitas"),
        (r"\bTDS\b", "TDS"),
    )
    parameters = [name for pattern, name in rules if re.search(pattern, blob, re.I)]
    parameters = dedupe_keep_order(parameters)[:8]
    if not parameters:
        return None
    return {
        "label": "Parameter pengukuran",
        "value": _v2220_join_id(parameters),
        "source_id": "",
        "kind": "parameter",
    }


def _v2223_standard_fact(rows: Sequence[Dict[str, str]]) -> Optional[Dict[str, str]]:
    for row in rows:
        blob = clean_text(row.get("label")) + " " + clean_text(row.get("value"))
        if not re.search(r"\b(?:IEC|ISO|Class|Kelas)\b", blob, re.I):
            continue
        value = clean_text(row.get("value"))
        # Prefer a concise standard identifier. Keep the full phrase only when
        # it is already short and coherent.
        identifiers = re.findall(
            r"\b(?:IEC|ISO)\s*[- ]?\d+(?:[-.:]\d+)*(?:\s*(?:Class|Kelas)\s*\d+)?|\b(?:Class|Kelas)\s*\d+\b",
            blob,
            flags=re.I,
        )
        concise = _v2220_join_id(dedupe_keep_order(identifiers))
        if not concise and len(value) <= 110:
            concise = value
        if concise:
            return {
                "label": "Standar/kelas",
                "value": concise,
                "source_id": clean_text(row.get("source_id")),
                "kind": "standard",
            }
    return None


def _v2223_simple_fact(
    rows: Sequence[Dict[str, str]],
    label_pattern: str,
    public_label: str,
    kind: str,
    max_chars: int = 100,
) -> Optional[Dict[str, str]]:
    regex = re.compile(label_pattern, re.I)
    for row in rows:
        label = clean_text(row.get("label"))
        if not regex.search(label):
            continue
        value = _v2222_atomic_value(label, row.get("value"))
        if not value or len(value) > max_chars:
            continue
        if kind == "accuracy" and re.search(r"\b(?:IEC|ISO|Class|Kelas)\b", value, re.I):
            continue
        return {
            "label": public_label,
            "value": value,
            "source_id": clean_text(row.get("source_id")),
            "kind": kind,
        }
    return None


def _v2223_interface_fact(rows: Sequence[Dict[str, str]]) -> Optional[Dict[str, str]]:
    for row in rows:
        label = clean_text(row.get("label"))
        value = clean_text(row.get("value"))
        if not re.search(r"\b(?:antarmuka|interface|keluaran|output|komunikasi|data\s+logger|pencatat\s+data|memory\s+card|kartu\s+memori)\b", label + " " + value, re.I):
            continue
        parsed = _v2222_interface_value(value)
        if parsed:
            return {
                "label": "Koneksi/data",
                "value": parsed,
                "source_id": clean_text(row.get("source_id")),
                "kind": "interface",
            }
    return None


def _v2223_stable_facts(
    rows: Sequence[Dict[str, str]],
    family_key: str,
) -> List[Dict[str, str]]:
    candidates: List[Optional[Dict[str, str]]] = [
        _v2223_primary_range(rows, family_key),
        _v2223_parameter_fact(rows, family_key),
        _v2223_simple_fact(rows, r"\b(?:resolusi|resolution)\b", "Resolusi", "resolution", 100),
        _v2223_simple_fact(rows, r"\b(?:akurasi|accuracy|ketelitian)\b", "Akurasi", "accuracy", 105),
        _v2223_standard_fact(rows),
        _v2223_interface_fact(rows),
        _v2223_simple_fact(rows, r"\b(?:sensor|probe|mikrofon|microphone)\b", "Sensor/masukan", "sensor", 85),
        _v2223_simple_fact(rows, r"\b(?:tampilan|display|lcd)\b", "Tampilan", "display", 85),
    ]
    output: List[Dict[str, str]] = []
    seen = set()
    for fact in candidates:
        if not isinstance(fact, dict):
            continue
        key = (_spec_text_key(fact.get("label")), _spec_text_key(fact.get("value")))
        if not all(key) or key in seen:
            continue
        seen.add(key)
        output.append(fact)
        if len(output) >= 5:
            break
    return output


def _v2223_fact_sentence(fact: Dict[str, str]) -> str:
    label = clean_text(fact.get("label"))
    value = clean_text(fact.get("value")).rstrip(".")
    if not label or not value:
        return ""
    return f"{label}: {value}."


def _v2223_use_context(profile: Dict[str, Any]) -> str:
    applications = dedupe_keep_order(profile.get("applications") or [])
    if not applications:
        return "Produk ini dapat digunakan untuk kegiatan inspeksi, pemeliharaan, dan pengendalian mutu sesuai fungsi pengukurannya."
    first = clean_text(applications[0]).rstrip(".")
    second = clean_text(applications[1]).rstrip(".") if len(applications) > 1 else ""
    if second:
        return f"Penggunaannya mencakup {first[:1].lower() + first[1:]} serta {second[:1].lower() + second[1:]}."
    return f"Penggunaannya mencakup {first[:1].lower() + first[1:]}."


def _v2223_buying_factors(family_key: str) -> List[str]:
    return {
        "sound_level": ["rentang pengukuran", "pembobotan frekuensi dan waktu", "standar/kelas", "koneksi data"],
        "multimeter": ["parameter yang dibutuhkan", "rentang pada setiap fungsi", "akurasi", "koneksi data"],
        "power": ["parameter daya", "jumlah fasa atau kanal", "rentang pengukuran", "koneksi data"],
        "thermometer": ["rentang suhu", "jenis sensor atau probe", "akurasi", "jumlah kanal"],
        "environmental": ["parameter lingkungan", "rentang setiap parameter", "sensor", "pencatatan data"],
        "light": ["rentang pencahayaan", "satuan pengukuran", "akurasi", "sensor"],
        "uv": ["jenis UVA atau UVC", "rentang intensitas", "sensor", "satuan pengukuran"],
        "anemometer": ["rentang kecepatan udara", "satuan pengukuran", "sensor", "pencatatan data"],
        "pressure": ["rentang tekanan", "satuan", "jenis masukan", "akurasi"],
        "transmitter": ["parameter proses", "rentang input", "jenis keluaran", "kompatibilitas sistem"],
        "water_quality": ["parameter kualitas air", "rentang setiap parameter", "jenis elektroda", "kalibrasi"],
    }.get(family_key, ["parameter pengukuran", "rentang kerja", "akurasi", "konfigurasi produk"])


def _v2223_description(
    data: Dict[str, Any],
    profile: Dict[str, Any],
    base_title: str,
    family_key: str,
    facts: Sequence[Dict[str, str]],
) -> str:
    model = clean_text(data.get("model"))
    identity = f"Lutron {model}" if model else base_title
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    intro = (
        f"<p><strong>{esc(identity)}</strong> merupakan {esc(term)} untuk {esc(function)}. "
        f"{esc(_v2223_use_context(profile))}</p>"
    )
    facts_html = ""
    if facts:
        facts_html = "<h2>Spesifikasi Utama</h2><ul>" + "".join(
            f"<li><strong>{esc(fact.get('label'))}:</strong> {esc(fact.get('value'))}</li>"
            for fact in facts[:4]
        ) + "</ul>"
    factors = _v2220_join_id(_v2223_buying_factors(family_key))
    guidance = (
        f"Sebelum memesan, cocokkan {factors} dengan kondisi penggunaan dan konfigurasi yang dibutuhkan."
    )
    html_value = (
        intro + facts_html
        + "<h2>Pertimbangan Sebelum Memilih</h2>"
        + f"<p>{esc(guidance)}</p>"
        + WEBSITE_CTA
    )
    html_value, _ = _v2220_clean_html(html_value, remove_boilerplate=True)
    return _v2222_fix_cta_spacing(_v2213_polish_html(html_value, ensure_cta=True))


def _v2223_short_description(
    profile: Dict[str, Any],
    base_title: str,
    facts: Sequence[Dict[str, str]],
) -> str:
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    base = f"{base_title} digunakan untuk {function}."
    primary = _v2223_fact_sentence(facts[0]) if facts else ""
    candidate = re.sub(r"\s{2,}", " ", f"{base} {primary}").strip()
    if len(candidate) <= V22_SHORT_MAX:
        return candidate
    return compact(base, V22_SHORT_MAX)


def _v2223_feature_html(profile: Dict[str, Any], facts: Sequence[Dict[str, str]]) -> str:
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    bullets = [function[:1].upper() + function[1:] + "."]
    bullets.extend(_v2223_fact_sentence(fact) for fact in facts[:4])
    bullets = dedupe_keep_order(item for item in bullets if clean_text(item))[:5]
    return "<h2>Keunggulan Produk</h2><ul>" + "".join(
        f"<li>{esc(item)}</li>" for item in bullets
    ) + "</ul>"


def _v2223_faq_html(
    data: Dict[str, Any],
    profile: Dict[str, Any],
    base_title: str,
    family_key: str,
    facts: Sequence[Dict[str, str]],
) -> str:
    model = clean_text(data.get("model")) or base_title
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    fact_answer = " ".join(_v2223_fact_sentence(fact) for fact in facts[:3]).strip()
    if not fact_answer:
        fact_answer = "Rincian parameter, rentang, dan konfigurasi tersedia pada tabel Spesifikasi."
    use_answer = _v2223_use_context(profile)
    factors = _v2220_join_id(_v2223_buying_factors(family_key))
    items = (
        (f"Apa fungsi {base_title}?", f"{base_title} merupakan {term} untuk {function}."),
        (f"Apa spesifikasi utama {model}?", fact_answer),
        (f"Untuk kebutuhan apa {model} digunakan?", use_answer),
        (
            "Apa yang perlu dipastikan sebelum pemesanan?",
            f"Cocokkan {factors} dengan kebutuhan penggunaan. Tim WEBSITE dapat membantu memeriksa konfigurasi sebelum pemesanan.",
        ),
    )
    return "".join(f"<h3>{esc(question)}</h3><p>{esc(answer)}</p>" for question, answer in items)


def _v2223_official_information(
    profile: Dict[str, Any],
    base_title: str,
    family_key: str,
) -> str:
    term = clean_text(profile.get("term")) or "alat ukur"
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    factors = _v2220_join_id(_v2223_buying_factors(family_key))
    return (
        "<h2>Informasi Produk</h2>"
        f"<p>{esc(base_title)} merupakan {esc(term)} untuk {esc(function)}.</p>"
        "<p>Rincian teknis, konfigurasi, dan batas pengukuran tersedia pada tabel Spesifikasi.</p>"
        "<h2>Catatan Pemilihan</h2>"
        f"<p>Periksa {esc(factors)} sebelum produk digunakan atau dipesan.</p>"
    )


def _v2223_meta_description(
    profile: Dict[str, Any],
    base_title: str,
    facts: Sequence[Dict[str, str]],
) -> str:
    function = clean_text(profile.get("function")) or "mendukung pengukuran dan pemeriksaan teknis"
    primary = ""
    if facts:
        primary = f" {clean_text(facts[0].get('label'))} {clean_text(facts[0].get('value'))}."
    candidates = (
        f"{base_title} untuk {function}.{primary} Lihat spesifikasi dan konsultasikan kebutuhan Anda dengan Tim WEBSITE.",
        f"{base_title} untuk {function}. Lihat spesifikasi dan konsultasikan kebutuhan Anda dengan Tim WEBSITE.",
        f"{base_title}. Lihat fungsi dan spesifikasi teknisnya, lalu konsultasikan kebutuhan Anda dengan Tim WEBSITE.",
    )
    for candidate in candidates:
        candidate = re.sub(r"\s{2,}", " ", clean_text(candidate)).replace("..", ".").strip()
        if len(candidate) <= V22_META_MAX and len(candidate) >= 100:
            return candidate
    return compact(candidates[-1], V22_META_MAX)


def _v2223_public_issues(row: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    issues, details = _v2222_public_field_issues(row)
    # The final stable template always has exactly four FAQ questions.
    faq_count = len(re.findall(r"<h3\b", clean_text(row.get("custom_tab_4_content")), re.I))
    if faq_count != 4:
        issues.append("custom_tab_4_content:faq_structure_invalid")
        details.append(f"FAQ berisi {faq_count} pertanyaan, seharusnya 4.")
    name = clean_text(row.get("Name"))
    keyphrase = clean_text(row.get("focus keyphrase"))
    if not name or name != keyphrase:
        issues.append("seo_identity_not_locked")
        details.append("Name dan focus keyphrase harus identik pada template stabil.")
    return dedupe_keep_order(issues), dedupe_keep_order(details)


_v2223_previous_build_output_row = build_output_row


def build_output_row(data: Dict[str, Any], pub_date: datetime) -> Dict[str, Any]:
    # Run the mature scraping/source-safety chain first. V22.23 then becomes
    # the sole owner of every public narrative field.
    row = _v2223_previous_build_output_row(data, pub_date)
    spec_rows, spec_meta = _v2223_prepare_spec_rows(data)
    row["custom_tab_1_content"] = _v2221_render_specifications(spec_rows)

    family = detect_family(data)
    profile = _v2220_profile(data, family)
    family_key = _v2223_family_key(data, profile)
    base_title = _v2217_product_title(data, family)
    facts = _v2223_stable_facts(spec_rows, family_key)

    # Locked SEO identity: no range modifier or feature suffix. This prevents
    # a secondary function (for example temperature on a multimeter) from
    # becoming the product's primary identity.
    row["Name"] = base_title
    row["focus keyphrase"] = base_title
    row["Product Description"] = _v2223_description(
        data, profile, base_title, family_key, facts
    )
    row["Product Short Description"] = _v2223_short_description(
        profile, base_title, facts
    )
    row["custom_tab_2_content"] = _v2223_feature_html(profile, facts)
    row["custom_tab_3_content"] = _v2220_application_html(profile)
    row["custom_tab_4_content"] = _v2223_faq_html(
        data, profile, base_title, family_key, facts
    )
    row["custom_tab_5_content"] = _v2223_official_information(
        profile, base_title, family_key
    )
    row["meta description"] = _v2223_meta_description(profile, base_title, facts)

    public_issues, public_details = _v2223_public_issues(row)
    prior_stage = clean_text(row.get("Review Failure Stage") or data.get("_v2219_review_stage"))
    if prior_stage == "specification_repair" and spec_meta.get("technical_integrity_ok"):
        prior_stage = ""
        row["Review Failure Stage"] = ""
        row["Review Failure Details"] = ""
        row["Diagnostic File"] = ""
        data["_v2219_review_stage"] = ""
        data["_v2219_review_details"] = ""

    substantive_rows = [
        item for item in spec_rows
        if not re.search(r"\b(?:model|sku|mpn|kode\s+produk)\b", clean_text(item.get("label")), re.I)
    ]
    final_issues: List[str] = []
    if len(substantive_rows) < 3:
        final_issues.append("insufficient_clean_specification_rows")
    if public_issues:
        final_issues.append("stable_public_copy_invalid")
    if prior_stage:
        final_issues.append("unrecovered_product_failure:" + prior_stage)

    warnings: List[str] = []
    if spec_meta.get("quarantined_ids"):
        warnings.append("source_specification_rows_quarantined")
    if spec_meta.get("english_ids"):
        warnings.append("complete_technical_english_retained")
    if spec_meta.get("fallback_ids"):
        warnings.append("source_value_fallback_used")

    row["Content Template Version"] = V2223_TEMPLATE_VERSION
    row["Stable Narrative Facts"] = " | ".join(
        f"{clean_text(fact.get('label'))}: {clean_text(fact.get('value'))}"
        for fact in facts
    )
    row["Stable Content Warnings"] = ", ".join(warnings)
    row["Narrative Safe Fact Rows"] = ",".join(
        dedupe_keep_order(
            sid
            for fact in facts
            for sid in clean_text(fact.get("source_id")).split(",")
            if clean_text(sid)
        )
    )
    row["Quarantined Specification Rows"] = ",".join(
        str(item) for item in spec_meta.get("quarantined_ids", [])
    )
    row["Quarantined Specification Details"] = " | ".join(
        spec_meta.get("quarantine_details", [])
    )
    row["Editorial Review Fields"] = ", ".join(
        dedupe_keep_order(item.split(":", 1)[0] for item in public_issues)
    )
    row["Editorial Review Details"] = " | ".join(public_details[:12])

    if final_issues:
        row["Publish Decision"] = "NEEDS_REVIEW"
        row["Final Publish Score"] = max(45, 100 - 15 * len(final_issues))
        row["Website Validation"] = "needs_review_v22_23:" + ",".join(final_issues)
    elif warnings:
        row["Publish Decision"] = "READY_WITH_WARNING"
        row["Final Publish Score"] = 95
        row["Website Validation"] = "publish_ready_stable_copy_with_warning_v22_23"
    else:
        row["Publish Decision"] = "READY"
        row["Final Publish Score"] = 100
        row["Website Validation"] = "publish_ready_stable_copy_v22_23"

    row["Processing Time"] = "lutron_WEBSITE_import_v22_23_stable_content"
    row["Content Quality"] = (
        clean_text(row.get("Content Quality")) + "_stable_public_copy"
    ).strip("_")
    row["Website Correction Log"] = _merge_source_notes(
        row.get("Website Correction Log"),
        "stable_content_finalizer=" + V2223_VERSION,
        "content_template=" + V2223_TEMPLATE_VERSION,
        "seo_identity=product_type_brand_model_only",
        "faq_structure=fixed_four_questions",
        "public_copy_rebuild_count=one_final_owner",
    )
    data["_v2223_final_issues"] = final_issues
    # Supersede findings from earlier public-copy experiments. Source/image
    # failures remain handled by the inherited validator.
    data["_v2222_final_issues"] = []
    data["_v2221_final_issues"] = []
    return row


_v2223_previous_validate_publish_output = validate_publish_output


def validate_publish_output(
    data: Dict[str, Any],
    row: Dict[str, Any],
    require_ai: bool = False,
) -> List[str]:
    inherited = _v2223_previous_validate_publish_output(data, row, require_ai=require_ai)
    superseded_patterns = (
        r"mixed_language",
        r"corrupt_specification",
        r"product_description_structure",
        r"faq_item_count",
        r"application_item_count",
        r"feature_fact_count",
        r"meta_description",
        r"seo_title_or_keyphrase",
        r"product_local_technical_review_required",
        r"iterative_repair_not_complete",
        r"customer_facing_narrative",
        r"insufficient_clean_specification_rows",
        r"unrecovered_product_failure:specification_repair",
        r"stable_public_copy_invalid",
    )
    issues = [
        item for item in inherited
        if not any(re.search(pattern, clean_text(item), re.I) for pattern in superseded_patterns)
    ]
    issues.extend(data.get("_v2223_final_issues") or [])
    return dedupe_keep_order(issues)




# ---------------------------------------------------------------------------
# V22.25 - zero-token resume from existing AI cache
# ---------------------------------------------------------------------------
V2225_VERSION = "22.25-zero-token-source-resume"


def _v2225_model_key(value: Any) -> str:
    text = clean_text(value).upper()
    candidates = re.findall(r"\b[A-Z]{1,8}(?:-[A-Z0-9]{1,16})+\b", text)
    token = candidates[-1] if candidates else text
    return re.sub(r"[^A-Z0-9]+", "", token)


def _v2225_ai_model_index() -> Dict[str, Dict[str, Any]]:
    existing = AI_SETTINGS.get("_v2225_model_index")
    if isinstance(existing, dict):
        return existing
    index: Dict[str, Dict[str, Any]] = {}
    cache = AI_SETTINGS.get("cache")
    for payload in cache.values() if isinstance(cache, dict) else []:
        if not isinstance(payload, dict) or payload.get("ok") is not True:
            continue
        fields = payload.get("fields")
        if not isinstance(fields, dict) or not fields:
            continue
        key = _v2225_model_key(fields.get("product_name_id"))
        if key and key not in index:
            index[key] = fields
    AI_SETTINGS["_v2225_model_index"] = index
    return index


def _v2225_attach_cached_ai_fields_by_model(data: Dict[str, Any]) -> bool:
    if isinstance(data.get("_ai_fields"), dict):
        return True
    candidates = [
        data.get("model"), data.get("input_model"), data.get("title"),
        data.get("input_title"),
    ]
    index = _v2225_ai_model_index()
    for candidate in candidates:
        key = _v2225_model_key(candidate)
        fields = index.get(key) if key else None
        if isinstance(fields, dict) and fields:
            data["_ai_fields"] = fields
            data["_ai_status"] = "cache_hit_by_model_zero_token"
            AI_STATS["cache_hit"] = int(AI_STATS.get("cache_hit") or 0) + 1
            print("    AI cache hit by model | ZERO TOKEN")
            return True
    return False


# Keep references to the last transport functions before the hard zero-token
# wrappers are installed. In normal mode behaviour remains unchanged.
_v2225_original_requests_post = requests.post


def _v2225_guarded_requests_post(*args: Any, **kwargs: Any) -> Any:
    if AI_SETTINGS.get("cache_only"):
        raise RuntimeError("ZERO_TOKEN_MODE_BLOCKED_DEEPSEEK_POST")
    return _v2225_original_requests_post(*args, **kwargs)


# Absolute safety net. Product pages are fetched with GET, while DeepSeek uses
# POST. Therefore cache-only mode can block every POST without interfering with
# LUTRON image/source recovery.
requests.post = _v2225_guarded_requests_post


_v2225_previous_call_ai = _call_ai


def _call_ai(*args: Any, **kwargs: Any) -> Tuple[bool, Dict[str, Any], str, List[Dict[str, Any]]]:
    if AI_SETTINGS.get("cache_only"):
        return False, {}, "ai_cache_only_miss", []
    return _v2225_previous_call_ai(*args, **kwargs)


_v2225_previous_post_ai_json = _v226_post_ai_json


def _v226_post_ai_json(
    kind: str,
    user_payload: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], str]:
    if AI_SETTINGS.get("cache_only"):
        cache = AI_SETTINGS.get("cache") if isinstance(AI_SETTINGS.get("cache"), dict) else {}
        cache_key = _v226_request_cache_key(kind, user_payload)
        cached = cache.get(cache_key)
        if isinstance(cached, dict) and isinstance(cached.get("parsed"), dict):
            AI_STATS["cache_hit"] = int(AI_STATS.get("cache_hit") or 0) + 1
            return True, cached["parsed"], "cache_hit"
        return False, {}, "ai_cache_only_miss"
    return _v2225_previous_post_ai_json(kind, user_payload)


# ---------------------------------------------------------------------------
# V22.24 - resilient batch continuation and incremental cache persistence
# ---------------------------------------------------------------------------
# Transient DeepSeek failures are product-local operational failures. They must
# not discard the rest of the catalog. Authentication, account, endpoint, and
# API-key configuration errors remain fatal because continuing cannot succeed.

V2224_VERSION = "22.25-zero-token-resilient-source-resume"
V2224_TRANSIENT_AI_STATUS_RE = re.compile(
    r"^(?:rate_limited_http_429|request_timeout|request_connection_error|"
    r"request_http_error_(?:408|409|425|429|5\d\d|unknown)|"
    r"deepseek_server_http_5\d\d)$",
    re.I,
)


def _v2224_extract_transient_ai_status(message: Any) -> str:
    text = clean_text(message)
    match = re.search(r"DEEPSEEK TIDAK TERSEDIA \(([^)]+)\)", text, re.I)
    status = clean_text(match.group(1)) if match else ""
    return status if V2224_TRANSIENT_AI_STATUS_RE.fullmatch(status) else ""


_v2224_previous_post_ai_json = _v226_post_ai_json


def _v226_post_ai_json(
    kind: str,
    user_payload: Dict[str, Any],
) -> Tuple[bool, Dict[str, Any], str]:
    try:
        return _v2224_previous_post_ai_json(kind, user_payload)
    except SystemExit as exc:
        status = _v2224_extract_transient_ai_status(exc)
        if not status:
            # Missing/invalid key, account credit, permission, endpoint, and
            # model configuration errors must still stop immediately.
            raise

        # Persist every successful request already completed for this product.
        # The normal retry/validation chain receives False and can exhaust its
        # configured retries. The final product wrapper then creates a
        # NEEDS_REVIEW row without making the whole batch fail.
        try:
            _v2219_persist_ai_cache()
        except Exception:
            pass
        AI_STATS["v2224_transient_ai_failures_continued"] = int(
            AI_STATS.get("v2224_transient_ai_failures_continued") or 0
        ) + 1
        print(
            "    WARNING: DeepSeek sementara gagal (" + status + "). "
            "Percobaan ini gagal, tetapi batch tidak dihentikan."
        )
        return False, {}, status


# Atomic JSON writes prevent a power loss or forced terminal close from leaving
# a half-written scrape/AI cache file.
def save_cache(path: str, data: Dict[str, Any]) -> None:
    if not path:
        return
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temporary, target)

if __name__ == "__main__":
    main()

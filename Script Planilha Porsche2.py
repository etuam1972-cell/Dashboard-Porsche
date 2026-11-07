"""
Porsche Sales Sanitization Agent
================================

Lê um arquivo .xlsx de vendas Porsche (cru) e aplica as regras definidas em
schema.md para gerar um novo arquivo tratado, com cada coluna sanitizada
inserida logo após a sua coluna fonte.

Uso:
    python sanitize_porsche.py [INPUT_XLSX] [OUTPUT_XLSX]

Sem argumentos, usa os defaults definidos em INPUT_PATH e OUTPUT_PATH abaixo.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------
INPUT_PATH = Path("porsche_database.xlsx")
OUTPUT_PATH = Path("porsche_database_sanitized.xlsx")
HTML_OUTPUT_PATH = Path("dashboard-porsche.html")
INVALID = "INVALID"


# ---------------------------------------------------------------------------
# Date sanitization
# ---------------------------------------------------------------------------
MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def _safe_date(y: int, m: int, d: int) -> str:
    try:
        return dt.date(y, m, d).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return INVALID


def sanitize_date(value: Any) -> str:
    if value is None or value == "":
        return INVALID

    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")

    s = str(value).strip()
    if not s:
        return INVALID

    # YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # MM/DD/YYYY or MM-DD-YYYY
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", s)
    if m:
        return _safe_date(int(m.group(3)), int(m.group(1)), int(m.group(2)))

    # MM/DD/YY or MM-DD-YY
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$", s)
    if m:
        yy = int(m.group(3))
        year = 2000 + yy if yy < 70 else 1900 + yy
        return _safe_date(year, int(m.group(1)), int(m.group(2)))

    # Month DDth, YYYY  /  Mon DDth YYYY  /  Month DD, YYYY
    m = re.match(
        r"^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?[,\s]+(\d{4})$",
        s,
    )
    if m:
        mon_key = m.group(1).lower().rstrip(".")
        if mon_key in MONTHS:
            return _safe_date(int(m.group(3)), MONTHS[mon_key], int(m.group(2)))

    return INVALID


# ---------------------------------------------------------------------------
# Porsche model sanitization
# ---------------------------------------------------------------------------
CANONICAL_MODELS = [
    "911 Carrera", "911 Carrera S", "911 Carrera GTS",
    "911 Turbo", "911 Turbo S",
    "911 GT3", "911 GT3 RS",
    "911 Dakar",
    "911 Targa 4", "911 Targa 4S",
    "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster GTS",
    "718 Spyder RS",
    "Cayenne", "Cayenne S", "Cayenne Coupe", "Cayenne E-Hybrid",
    "Cayenne Turbo", "Cayenne Turbo GT",
    "Macan", "Macan S", "Macan T", "Macan GTS", "Macan Electric",
    "Panamera", "Panamera 4", "Panamera 4S",
    "Panamera Turbo", "Panamera Turbo S", "Panamera 4 E-Hybrid",
    "Taycan", "Taycan 4S", "Taycan GTS",
    "Taycan Turbo", "Taycan Turbo S", "Taycan Cross Turismo",
]

# Lookup case-insensitive por chave normalizada
_MODEL_LOOKUP = {m.lower(): m for m in CANONICAL_MODELS}


def _smart_title(s: str) -> str:
    """Title case preservando palavras comuns como E-Hybrid, GTS, RS, etc."""
    keep_upper = {"GT", "GTS", "RS", "GT3", "GT4", "S", "T", "4", "4S", "USA"}
    parts = s.split()
    out = []
    for p in parts:
        token = p.upper()
        if token in keep_upper:
            out.append(token)
        elif "-" in p:
            out.append("-".join(seg.capitalize() for seg in p.split("-")))
        else:
            out.append(p.capitalize())
    return " ".join(out)


def sanitize_model(value: Any) -> str:
    if value is None:
        return INVALID
    s = " ".join(str(value).split()).strip()
    if not s:
        return INVALID
    key = s.lower()
    if key in _MODEL_LOOKUP:
        return _MODEL_LOOKUP[key]
    # Modelos desconhecidos ficam title-case
    return _smart_title(s)


# ---------------------------------------------------------------------------
# Model year sanitization
# ---------------------------------------------------------------------------
WORDS_TO_NUM = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17,
    "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30,
    "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90,
    "hundred": 100, "thousand": 1000,
}


def _words_to_int(text: str) -> int | None:
    """Converte 'twenty twenty four', 'two thousand twenty two', 'eighty two thousand', etc."""
    tokens = re.findall(r"[a-z]+", text.lower())
    if not tokens:
        return None
    if any(t not in WORDS_TO_NUM for t in tokens):
        return None

    # Trata caso "twenty twenty four" (estilo ano: pares de dezenas)
    # quando todas tokens são <= 90 e nenhuma é hundred/thousand
    if all(WORDS_TO_NUM[t] < 100 for t in tokens):
        # tenta ler como dois pares: "twenty twenty four" -> 20*100 + 24
        nums = [WORDS_TO_NUM[t] for t in tokens]
        # combinar pares: ex [20, 20, 4] -> dois números: 20 e 24
        groups: list[int] = []
        i = 0
        while i < len(nums):
            if i + 1 < len(nums) and nums[i] >= 20 and nums[i] % 10 == 0 and nums[i + 1] < 10:
                groups.append(nums[i] + nums[i + 1])
                i += 2
            else:
                groups.append(nums[i])
                i += 1
        if len(groups) == 2:
            return groups[0] * 100 + groups[1]
        if len(groups) == 1:
            return groups[0]

    # Estilo "two thousand twenty two" / "eighty two thousand"
    total = 0
    current = 0
    for t in tokens:
        v = WORDS_TO_NUM[t]
        if v == 100:
            current = max(1, current) * 100
        elif v == 1000:
            current = max(1, current) * 1000
            total += current
            current = 0
        else:
            current += v
    return total + current


def sanitize_year(value: Any) -> str:
    if value is None:
        return INVALID

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        y = int(value)
        return str(y) if 1990 <= y <= 2035 else INVALID

    s = str(value).strip()
    if not s:
        return INVALID

    # 4 dígitos
    m = re.fullmatch(r"\d{4}", s)
    if m:
        y = int(s)
        return str(y) if 1990 <= y <= 2035 else INVALID

    # Formatos como 20-23, 20 24, 20.24
    m = re.fullmatch(r"(\d{2})\s*[-\s.]\s*(\d{2})", s)
    if m:
        y = int(m.group(1) + m.group(2))
        return str(y) if 1990 <= y <= 2035 else INVALID

    # Texto "twenty twenty four"
    if re.search(r"[A-Za-z]", s):
        n = _words_to_int(s)
        if n is not None and 1990 <= n <= 2035:
            return str(n)

    return INVALID


# ---------------------------------------------------------------------------
# Sales price sanitization
# ---------------------------------------------------------------------------
PRICE_WORD_MULTIPLIERS = {
    "thousand": 1_000,
    "million": 1_000_000,
}


def _parse_textual_price(s: str) -> float | None:
    """Parses 'eighty two thousand', 'two hundred thousand', etc."""
    tokens = re.findall(r"[a-z]+", s.lower())
    if not tokens:
        return None
    valid = set(WORDS_TO_NUM) | {"million"}
    if any(t not in valid for t in tokens):
        return None

    total = 0
    current = 0
    for t in tokens:
        if t == "million":
            current = max(1, current) * 1_000_000
            total += current
            current = 0
        elif t == "hundred":
            current = max(1, current) * 100
        elif t == "thousand":
            current = max(1, current) * 1_000
            total += current
            current = 0
        else:
            current += WORDS_TO_NUM[t]
    return float(total + current)


def _parse_numeric_price(num_str: str) -> float | None:
    """Lida com formatos US e europeus."""
    s = num_str.strip()
    if not s:
        return None

    has_dot = "." in s
    has_comma = "," in s

    if has_dot and has_comma:
        # Last separator wins as decimal
        if s.rfind(",") > s.rfind("."):
            # europeu: . é milhar, , é decimal
            s = s.replace(".", "").replace(",", ".")
        else:
            # US: , é milhar, . é decimal
            s = s.replace(",", "")
    elif has_comma:
        # Apenas vírgula. Se houver mais de uma OU os dígitos pós-vírgula não forem 2,
        # tratamos como milhar.
        last = s.rsplit(",", 1)
        digits_after = re.sub(r"\D", "", last[1])
        if s.count(",") > 1 or len(digits_after) != 2:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif has_dot:
        # Apenas ponto. Se múltiplos pontos OU dígitos após o último não forem 2,
        # tratamos como milhar (estilo europeu).
        last = s.rsplit(".", 1)
        digits_after = re.sub(r"\D", "", last[1])
        if s.count(".") > 1 or len(digits_after) != 2:
            s = s.replace(".", "")

    try:
        return float(s)
    except ValueError:
        return None


def sanitize_price(value: Any) -> str:
    if value is None or value == "":
        return INVALID

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{float(value):.2f}"

    s = str(value).strip()
    if not s:
        return INVALID

    # Remover símbolos comuns e palavras-chave (USD, dollars)
    cleaned = s.replace("$", " ")
    cleaned = re.sub(r"\b(usd|dollars?)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()

    # Preço textual: contém apenas letras e espaços
    if re.fullmatch(r"[A-Za-z\s]+", cleaned):
        n = _parse_textual_price(cleaned)
        return f"{n:.2f}" if n is not None else INVALID

    # Multiplicador "k" (mil)
    multiplier = 1.0
    m = re.search(r"([0-9.,]+)\s*k\b", cleaned, flags=re.IGNORECASE)
    if m:
        num_part = m.group(1)
        multiplier = 1000.0
    else:
        # Pega o token numérico
        m = re.search(r"[0-9.,]+", cleaned)
        if not m:
            return INVALID
        num_part = m.group(0)

    n = _parse_numeric_price(num_part)
    if n is None:
        return INVALID
    return f"{n * multiplier:.2f}"


# ---------------------------------------------------------------------------
# Vehicle mileage sanitization
# ---------------------------------------------------------------------------
def sanitize_mileage(value: Any) -> str:
    if value is None:
        return INVALID

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Pode vir como float (ex.: 9.5). Tratamos como inteiro de milhas.
        return str(int(round(float(value))))

    s = str(value).strip()
    if not s:
        return INVALID

    low = s.lower()
    if low in {"new", "new car", "zero miles", "zero", "0 mi", "0 miles", "0mi"}:
        return "0"

    # KM (precisa converter)
    is_km = bool(re.search(r"\bkm\b", low) or low.startswith("km"))

    # Texto puro (sem dígitos): "twelve thousand miles", "fifteen thousand miles"
    digits_present = bool(re.search(r"\d", s))
    if not digits_present:
        # Remover unidade
        textual = re.sub(r"\b(mi|miles|mile|km|kilometers|kilometres)\b", " ", low)
        textual = textual.strip()
        n = _parse_textual_price(textual)  # mesma lógica de palavras
        if n is None:
            return INVALID
        miles = n * (0.621371 if is_km else 1.0)
        return str(int(round(miles)))

    # Capturar parte numérica
    m = re.search(r"[0-9][0-9.,]*", s)
    if not m:
        return INVALID
    num_part = m.group(0)
    n = _parse_numeric_price(num_part)
    if n is None:
        return INVALID
    if is_km:
        n = n * 0.621371
    return str(int(round(n)))


# ---------------------------------------------------------------------------
# Payment method sanitization
# ---------------------------------------------------------------------------
PAYMENT_CANONICAL = [
    "Credit Card", "Debit Card", "Bank Transfer", "Wire Transfer",
    "Financing", "Lease", "Cash", "ACH Payment", "Crypto Payment",
]


def _normalize_text_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[\W_]+", " ", s.lower()).strip()
    s = re.sub(r"\s+", " ", s)
    return s


PAYMENT_LOOKUP = {
    # Credit Card
    "credit card": "Credit Card",
    "creditcard": "Credit Card",
    "credit": "Credit Card",
    "credit card payment": "Credit Card",
    # Debit Card
    "debit card": "Debit Card",
    "debitcard": "Debit Card",
    # Bank Transfer
    "bank transfer": "Bank Transfer",
    "banktransfer": "Bank Transfer",
    "bank": "Bank Transfer",
    # Wire Transfer
    "wire transfer": "Wire Transfer",
    "wiretransfer": "Wire Transfer",
    "wire": "Wire Transfer",
    "bank wire": "Wire Transfer",
    "wire bank": "Wire Transfer",
    # Financing
    "financing": "Financing",
    "finance": "Financing",
    "financing plan": "Financing",
    # Lease
    "lease": "Lease",
    "leasing": "Lease",
    "lease plan": "Lease",
    # Cash
    "cash": "Cash",
    "cash payment": "Cash",
    # ACH
    "ach": "ACH Payment",
    "ach payment": "ACH Payment",
    # Crypto
    "crypto": "Crypto Payment",
    "crypto payment": "Crypto Payment",
    "cryptocurrency": "Crypto Payment",
}


def sanitize_payment(value: Any) -> str:
    if value is None:
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID
    key = _normalize_text_key(s)
    if key in PAYMENT_LOOKUP:
        return PAYMENT_LOOKUP[key]
    # Desconhecidos: title case
    return " ".join(w.capitalize() for w in key.split())


# ---------------------------------------------------------------------------
# City sanitization
# ---------------------------------------------------------------------------
def sanitize_city(value: Any) -> str:
    if value is None:
        return INVALID
    s = " ".join(str(value).split()).strip()
    if not s:
        return INVALID

    parts = s.split(" ")
    out: list[str] = []
    for p in parts:
        if p.lower() == "st." or p.lower() == "st":
            out.append("St.")
        elif p.lower() == "mt." or p.lower() == "mt":
            out.append("Mt.")
        elif "-" in p:
            out.append("-".join(seg.capitalize() for seg in p.split("-")))
        else:
            out.append(p.capitalize())
    return " ".join(out)


# ---------------------------------------------------------------------------
# State sanitization
# ---------------------------------------------------------------------------
US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE",
    "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC",
    "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "district of columbia": "DC",
}
US_STATE_CODES = set(US_STATES.values())


def sanitize_state(value: Any) -> str:
    if value is None:
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID
    upper = s.upper().strip()
    if upper in US_STATE_CODES:
        return upper
    key = " ".join(s.lower().split())
    if key in US_STATES:
        return US_STATES[key]
    return INVALID


# ---------------------------------------------------------------------------
# Delivery status sanitization
# ---------------------------------------------------------------------------
DELIVERY_LOOKUP = {
    "delivered": "Delivered",
    "deliverd": "Delivered",  # typo comum
    "pending": "Pending",
    "in transit": "In Transit",
    "intransit": "In Transit",
    "cancelled": "Cancelled",
    "canceled": "Cancelled",
    "awaiting delivery": "Awaiting Delivery",
    "awaiting pickup": "Awaiting Pickup",
    "pending approval": "Pending Approval",
    "pending review": "Pending Review",
    "shipped": "Shipped",
    "awaiting review": "Awaiting Review",
}


def sanitize_delivery(value: Any) -> str:
    if value is None:
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID
    key = _normalize_text_key(s)
    if key in DELIVERY_LOOKUP:
        return DELIVERY_LOOKUP[key]
    return " ".join(w.capitalize() for w in key.split()) or INVALID


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------
COLUMN_PIPELINE = [
    # (input_col, sanitized_col_name or None, sanitizer or None)
    ("sale_id", None, None),
    ("sale_date", "SaleDateSanitized", sanitize_date),
    ("customer_name", None, None),
    ("porsche_model", "PorscheModelSanitized", sanitize_model),
    ("model_year", "ModelYearSanitized", sanitize_year),
    ("sale_price", "SalesPriceSanitized", sanitize_price),
    ("vehicle_mileage", "VehicleMileageSanitized", sanitize_mileage),
    ("payment_method", "PayMethodSanitized", sanitize_payment),
    ("city", "CitySanitized", sanitize_city),
    ("state", "StateSanitized", sanitize_state),
    ("salesperson", None, None),
    ("delivery_status", "DeliveryStatusSanitized", sanitize_delivery),
]


def process_workbook(input_path: Path, output_path: Path) -> tuple[dict[str, int], list[dict[str, Any]]]:
    wb_in = load_workbook(input_path, data_only=True)
    ws_in = wb_in.active

    rows = list(ws_in.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Empty workbook")

    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    name_to_idx = {h: i for i, h in enumerate(header)}

    # Validar presença das colunas requeridas
    missing = [c for c, _, _ in COLUMN_PIPELINE if c not in name_to_idx]
    if missing:
        raise ValueError(f"Missing required columns in input: {missing}")

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Sanitized"

    # Cabeçalho final
    out_header: list[str] = []
    for src, san_name, _ in COLUMN_PIPELINE:
        out_header.append(src)
        if san_name:
            out_header.append(san_name)
    ws_out.append(out_header)

    # Estilo do cabeçalho
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F2937")
    sanitized_fill = PatternFill("solid", fgColor="065F46")
    for col_idx, name in enumerate(out_header, start=1):
        cell = ws_out.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = sanitized_fill if name.endswith("Sanitized") else header_fill

    # Estatísticas de invalidos
    invalid_counts: dict[str, int] = {}
    records: list[dict[str, Any]] = []

    for r in rows[1:]:
        out_row: list[Any] = []
        rec_dict: dict[str, Any] = {}
        for src, san_name, fn in COLUMN_PIPELINE:
            raw = r[name_to_idx[src]]
            out_row.append(raw)
            rec_dict[src] = raw
            if san_name and fn is not None:
                clean = fn(raw)
                if clean == INVALID:
                    invalid_counts[san_name] = invalid_counts.get(san_name, 0) + 1
                out_row.append(clean)
                rec_dict[san_name] = clean
        ws_out.append(out_row)
        records.append(rec_dict)

    # Largura das colunas
    for col_idx in range(1, len(out_header) + 1):
        letter = get_column_letter(col_idx)
        ws_out.column_dimensions[letter].width = max(14, len(out_header[col_idx - 1]) + 2)
    ws_out.freeze_panes = "A2"

    wb_out.save(output_path)
    return invalid_counts, records



def generate_html_dashboard(records: list[dict[str, Any]], output_html: Path) -> None:
    """Gera uma pagina HTML interativa e responsiva com Dashboard de Vendas Porsche."""
    def json_serial(obj):
        if isinstance(obj, (dt.datetime, dt.date)):
            return obj.strftime("%Y-%m-%d")
        return str(obj)
    json_data = json.dumps(records, default=json_serial, ensure_ascii=False)
    
    html_markup = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Porsche | Executive Sales Dashboard</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {{
    --porsche-black: #0E0E0E;
    --porsche-red: #D5001C;
    --porsche-gold: #D4AF37;
    --dark-card: #161616;
    --dark-border: #262626;
    --gray-text: #A0A0A0;
    --light-gray: #EAEAEA;
    --white: #FFFFFF;
    --success: #00C853;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: #0A0A0A;
    color: var(--light-gray);
    min-height: 100vh;
  }}
  header.topbar {{
    background: linear-gradient(180deg, #181818 0%, #0E0E0E 100%);
    border-bottom: 2px solid var(--porsche-red);
    padding: 20px 40px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 15px;
  }}
  .brand-area {{
    display: flex;
    align-items: center;
    gap: 16px;
  }}
  .brand-icon {{
    font-size: 32px;
    color: var(--porsche-red);
  }}
  .brand-title {{
    font-size: 22px;
    font-weight: 900;
    letter-spacing: 5px;
    color: var(--white);
    text-transform: uppercase;
  }}
  .brand-title span {{ color: var(--porsche-red); }}
  .brand-sub {{
    font-size: 11px;
    letter-spacing: 2px;
    color: var(--porsche-gold);
    text-transform: uppercase;
    margin-top: 3px;
  }}
  .header-actions {{
    display: flex;
    gap: 12px;
    align-items: center;
  }}
  .btn-export {{
    background: var(--porsche-red);
    color: var(--white);
    border: none;
    padding: 10px 22px;
    border-radius: 4px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 8px;
    transition: all 0.2s;
    letter-spacing: 1px;
    text-transform: uppercase;
  }}
  .btn-export:hover {{
    background: #FF1734;
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(213,0,28,0.4);
  }}
  main {{
    max-width: 1440px;
    margin: 0 auto;
    padding: 28px 24px 60px;
  }}
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 18px;
    margin-bottom: 28px;
  }}
  .kpi-card {{
    background: var(--dark-card);
    border: 1px solid var(--dark-border);
    border-radius: 8px;
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
  }}
  .kpi-card::before {{
    content: "";
    position: absolute;
    top: 0; left: 0;
    width: 4px; height: 100%;
    background: var(--porsche-red);
  }}
  .kpi-card.gold::before {{ background: var(--porsche-gold); }}
  .kpi-card.green::before {{ background: var(--success); }}
  .kpi-label {{
    font-size: 12px;
    color: var(--gray-text);
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin-bottom: 10px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .kpi-value {{
    font-size: 26px;
    font-weight: 800;
    color: var(--white);
    letter-spacing: -0.5px;
  }}
  .kpi-detail {{
    font-size: 12px;
    color: var(--porsche-gold);
    margin-top: 6px;
  }}
  .filters-panel {{
    background: var(--dark-card);
    border: 1px solid var(--dark-border);
    border-radius: 8px;
    padding: 18px 22px;
    margin-bottom: 28px;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 14px;
    align-items: flex-end;
  }}
  .filter-group {{
    display: flex;
    flex-direction: column;
    gap: 6px;
  }}
  .filter-group label {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--gray-text);
  }}
  .filter-group select, .filter-group input {{
    background: #0E0E0E;
    color: var(--white);
    border: 1px solid var(--dark-border);
    padding: 9px 12px;
    border-radius: 4px;
    font-size: 13px;
    outline: none;
    transition: border-color 0.2s;
  }}
  .filter-group select:focus, .filter-group input:focus {{
    border-color: var(--porsche-red);
  }}
  .btn-reset {{
    background: #252525;
    color: var(--light-gray);
    border: 1px solid var(--dark-border);
    padding: 9px 16px;
    border-radius: 4px;
    font-size: 12px;
    cursor: pointer;
    font-weight: 600;
    height: 38px;
    transition: background 0.2s;
  }}
  .btn-reset:hover {{
    background: #333333;
  }}
  .charts-grid {{
    display: grid;
    grid-template-columns: 2fr 1fr;
    gap: 20px;
    margin-bottom: 28px;
  }}
  @media (max-width: 992px) {{
    .charts-grid {{ grid-template-columns: 1fr; }}
  }}
  .chart-card {{
    background: var(--dark-card);
    border: 1px solid var(--dark-border);
    border-radius: 8px;
    padding: 20px 22px;
  }}
  .chart-header {{
    font-size: 14px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--white);
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }}
  .chart-header i {{ color: var(--porsche-red); }}
  .chart-wrapper {{
    position: relative;
    height: 340px;
  }}
  .table-card {{
    background: var(--dark-card);
    border: 1px solid var(--dark-border);
    border-radius: 8px;
    padding: 20px 22px;
    overflow-x: auto;
  }}
  table.data-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    text-align: left;
  }}
  table.data-table th {{
    background: #141414;
    color: var(--white);
    padding: 12px 14px;
    border-bottom: 2px solid var(--dark-border);
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    font-size: 11px;
  }}
  table.data-table td {{
    padding: 12px 14px;
    border-bottom: 1px solid #202020;
    color: #CCCCCC;
  }}
  table.data-table tr:hover td {{
    background: #1C1C1C;
    color: var(--white);
  }}
  .badge-status {{
    padding: 4px 10px;
    border-radius: 3px;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    display: inline-block;
  }}
  .badge-delivered {{ background: rgba(0,200,83,0.15); color: #00E676; border: 1px solid #00C853; }}
  .badge-transit {{ background: rgba(212,175,55,0.15); color: #FFD700; border: 1px solid #D4AF37; }}
  .badge-pending {{ background: rgba(213,0,28,0.15); color: #FF5252; border: 1px solid #D5001C; }}
  .badge-other {{ background: rgba(255,255,255,0.1); color: #BBB; border: 1px solid #444; }}
  footer {{
    text-align: center;
    color: var(--gray-text);
    font-size: 12px;
    padding: 24px;
    letter-spacing: 1px;
  }}
</style>
</head>
<body>

<header class="topbar">
  <div class="brand-area">
    <i class="fa-solid fa-car-side brand-icon"></i>
    <div>
      <div class="brand-title">PORSCHE <span>INTELLIGENCE</span></div>
      <div class="brand-sub">Executive Sales &amp; Performance Analytics</div>
    </div>
  </div>
  <div class="header-actions">
    <button class="btn-export" onclick="exportDataExcel()">
      <i class="fa-solid fa-file-excel"></i> Exportar XLSX
    </button>
  </div>
</header>

<main>
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label"><span>Faturamento Total</span> <i class="fa-solid fa-dollar-sign"></i></div>
      <div class="kpi-value" id="kpiRevenue">$0.00</div>
      <div class="kpi-detail" id="kpiAvgPrice">Ticket Medio: $0.00</div>
    </div>
    <div class="kpi-card gold">
      <div class="kpi-label"><span>Veiculos Vendidos</span> <i class="fa-solid fa-car"></i></div>
      <div class="kpi-value" id="kpiTotalUnits">0</div>
      <div class="kpi-detail" id="kpiDeliveredUnits">0 Entregues</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label"><span>Milhagem Media</span> <i class="fa-solid fa-gauge-high"></i></div>
      <div class="kpi-value" id="kpiAvgMileage">0 mi</div>
      <div class="kpi-detail">Media por unidade</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label"><span>Modelo Lider</span> <i class="fa-solid fa-trophy"></i></div>
      <div class="kpi-value" id="kpiTopModel" style="font-size:20px;">-</div>
      <div class="kpi-detail" id="kpiTopModelRevenue">$0.00</div>
    </div>
  </div>

  <div class="filters-panel">
    <div class="filter-group">
      <label><i class="fa-solid fa-car"></i> Modelo</label>
      <select id="filterModel" onchange="applyFilters()">
        <option value="">Todos os Modelos</option>
      </select>
    </div>
    <div class="filter-group">
      <label><i class="fa-solid fa-calendar"></i> Ano Modelo</label>
      <select id="filterYear" onchange="applyFilters()">
        <option value="">Todos os Anos</option>
      </select>
    </div>
    <div class="filter-group">
      <label><i class="fa-solid fa-truck-ramp-box"></i> Status Entrega</label>
      <select id="filterStatus" onchange="applyFilters()">
        <option value="">Todos os Status</option>
      </select>
    </div>
    <div class="filter-group">
      <label><i class="fa-solid fa-credit-card"></i> Pagamento</label>
      <select id="filterPayment" onchange="applyFilters()">
        <option value="">Todas as Formas</option>
      </select>
    </div>
    <div class="filter-group">
      <label><i class="fa-solid fa-magnifying-glass"></i> Busca</label>
      <input type="text" id="filterSearch" placeholder="Cliente, cidade..." oninput="applyFilters()" />
    </div>
    <button class="btn-reset" onclick="resetFilters()">
      <i class="fa-solid fa-rotate-left"></i> Limpar
    </button>
  </div>

  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-header"><i class="fa-solid fa-chart-column"></i> Top 10 Modelos por Faturamento</div>
      <div class="chart-wrapper">
        <canvas id="chartRevenue"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><i class="fa-solid fa-chart-pie"></i> Distribuicao por Status</div>
      <div class="chart-wrapper">
        <canvas id="chartStatus"></canvas>
      </div>
    </div>
  </div>

  <div class="table-card">
    <div class="chart-header" style="margin-bottom:14px;">
      <i class="fa-solid fa-table-list"></i> Detalhamento das Vendas (<span id="tableCount">0</span> registros)
    </div>
    <table class="data-table">
      <thead>
        <tr>
          <th>ID</th>
          <th>Data</th>
          <th>Cliente</th>
          <th>Modelo Porsche</th>
          <th>Ano</th>
          <th>Preco (USD)</th>
          <th>Milhagem</th>
          <th>Pagamento</th>
          <th>Local</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody id="tableBody"></tbody>
    </table>
  </div>
</main>

<footer>
  PORSCHE SALES INTELLIGENCE &bull; Desenvolvido para Santander Coders / DIO
</footer>

<script>
const rawDatabase = {json_data};

let chartRevenueInstance = null;
let chartStatusInstance = null;
let currentFiltered = [...rawDatabase];

function initFilters() {{
  const models = [...new Set(rawDatabase.map(d => d.PorscheModelSanitized || d.porsche_model))].sort();
  const years = [...new Set(rawDatabase.map(d => d.ModelYearSanitized || d.model_year))].filter(y => y && y !== 'INVALID').sort();
  const statuses = [...new Set(rawDatabase.map(d => d.DeliveryStatusSanitized || d.delivery_status))].filter(Boolean).sort();
  const payments = [...new Set(rawDatabase.map(d => d.PayMethodSanitized || d.payment_method))].filter(Boolean).sort();

  const selModel = document.getElementById('filterModel');
  models.forEach(m => selModel.add(new Option(m, m)));

  const selYear = document.getElementById('filterYear');
  years.forEach(y => selYear.add(new Option(y, y)));

  const selStatus = document.getElementById('filterStatus');
  statuses.forEach(s => selStatus.add(new Option(s, s)));

  const selPayment = document.getElementById('filterPayment');
  payments.forEach(p => selPayment.add(new Option(p, p)));
}}

function getPrice(d) {{
  const p = d.SalesPriceSanitized;
  return (p && p !== 'INVALID') ? parseFloat(p) : 0;
}}

function getMileage(d) {{
  const m = d.VehicleMileageSanitized;
  return (m && m !== 'INVALID') ? parseInt(m, 10) : 0;
}}

function applyFilters() {{
  const model = document.getElementById('filterModel').value;
  const year = document.getElementById('filterYear').value;
  const status = document.getElementById('filterStatus').value;
  const payment = document.getElementById('filterPayment').value;
  const search = document.getElementById('filterSearch').value.toLowerCase().trim();

  currentFiltered = rawDatabase.filter(d => {{
    const dModel = d.PorscheModelSanitized || d.porsche_model;
    const dYear = d.ModelYearSanitized || d.model_year;
    const dStatus = d.DeliveryStatusSanitized || d.delivery_status;
    const dPayment = d.PayMethodSanitized || d.payment_method;
    const dClient = String(d.customer_name || '').toLowerCase();
    const dCity = String(d.CitySanitized || d.city || '').toLowerCase();
    const dState = String(d.StateSanitized || d.state || '').toLowerCase();

    if (model && dModel !== model) return false;
    if (year && String(dYear) !== String(year)) return false;
    if (status && dStatus !== status) return false;
    if (payment && dPayment !== payment) return false;
    if (search && !(dClient.includes(search) || dCity.includes(search) || dState.includes(search) || dModel.toLowerCase().includes(search))) return false;
    return true;
  }});

  renderDashboard();
}}

function resetFilters() {{
  document.getElementById('filterModel').value = '';
  document.getElementById('filterYear').value = '';
  document.getElementById('filterStatus').value = '';
  document.getElementById('filterPayment').value = '';
  document.getElementById('filterSearch').value = '';
  applyFilters();
}}

function renderDashboard() {{
  const totalRev = currentFiltered.reduce((acc, d) => acc + getPrice(d), 0);
  const totalUnits = currentFiltered.length;
  const avgPrice = totalUnits > 0 ? totalRev / totalUnits : 0;
  const totalMiles = currentFiltered.reduce((acc, d) => acc + getMileage(d), 0);
  const avgMiles = totalUnits > 0 ? Math.round(totalMiles / totalUnits) : 0;
  const deliveredUnits = currentFiltered.filter(d => (d.DeliveryStatusSanitized || d.delivery_status) === 'Delivered').length;

  const modelStats = {{}};
  currentFiltered.forEach(d => {{
    const m = d.PorscheModelSanitized || d.porsche_model;
    if (!modelStats[m]) modelStats[m] = {{ rev: 0, count: 0 }};
    modelStats[m].rev += getPrice(d);
    modelStats[m].count += 1;
  }});

  let topModel = '-';
  let topModelRev = 0;
  Object.entries(modelStats).forEach(([m, s]) => {{
    if (s.rev > topModelRev) {{
      topModelRev = s.rev;
      topModel = m;
    }}
  }});

  document.getElementById('kpiRevenue').textContent = '$' + totalRev.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
  document.getElementById('kpiAvgPrice').textContent = 'Ticket Medio: $' + avgPrice.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
  document.getElementById('kpiTotalUnits').textContent = totalUnits;
  document.getElementById('kpiDeliveredUnits').textContent = deliveredUnits + ' Entregues';
  document.getElementById('kpiAvgMileage').textContent = avgMiles.toLocaleString('en-US') + ' mi';
  document.getElementById('kpiTopModel').textContent = topModel;
  document.getElementById('kpiTopModelRevenue').textContent = '$' + topModelRev.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});

  renderCharts(modelStats);
  renderTable();
}}

function renderCharts(modelStats) {{
  const sortedModels = Object.entries(modelStats)
    .sort((a, b) => b[1].rev - a[1].rev)
    .slice(0, 10);

  const modelLabels = sortedModels.map(x => x[0]);
  const modelData = sortedModels.map(x => x[1].rev);

  if (chartRevenueInstance) chartRevenueInstance.destroy();
  const ctxRev = document.getElementById('chartRevenue').getContext('2d');
  chartRevenueInstance = new Chart(ctxRev, {{
    type: 'bar',
    data: {{
      labels: modelLabels,
      datasets: [{{
        label: 'Faturamento (USD)',
        data: modelData,
        backgroundColor: '#D5001C',
        borderColor: '#FF1734',
        borderWidth: 1,
        borderRadius: 4
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: (ctx) => '$ ' + ctx.raw.toLocaleString('en-US', {{ minimumFractionDigits: 2 }})
          }}
        }}
      }},
      scales: {{
        x: {{
          ticks: {{ color: '#A0A0A0', font: {{ size: 10 }}, maxRotation: 45, minRotation: 30 }},
          grid: {{ color: '#1E1E1E' }}
        }},
        y: {{
          ticks: {{
            color: '#A0A0A0',
            callback: (v) => '$' + (v >= 1e6 ? (v/1e6).toFixed(1) + 'M' : (v/1e3).toFixed(0) + 'k')
          }},
          grid: {{ color: '#1E1E1E' }}
        }}
      }}
    }}
  }});

  const statusCounts = {{}};
  currentFiltered.forEach(d => {{
    const s = d.DeliveryStatusSanitized || d.delivery_status || 'Outro';
    statusCounts[s] = (statusCounts[s] || 0) + 1;
  }});

  if (chartStatusInstance) chartStatusInstance.destroy();
  const ctxStatus = document.getElementById('chartStatus').getContext('2d');
  chartStatusInstance = new Chart(ctxStatus, {{
    type: 'doughnut',
    data: {{
      labels: Object.keys(statusCounts),
      datasets: [{{
        data: Object.values(statusCounts),
        backgroundColor: [
          '#00C853', '#D5001C', '#D4AF37', '#2979FF', '#FF9100',
          '#9C27B0', '#00B0FF', '#78909C', '#546E7A', '#37474F'
        ],
        borderWidth: 2,
        borderColor: '#181818'
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{
          position: 'right',
          labels: {{ color: '#EAEAEA', boxWidth: 12, font: {{ size: 10 }} }}
        }}
      }}
    }}
  }});
}}

function renderTable() {{
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';
  document.getElementById('tableCount').textContent = currentFiltered.length;

  currentFiltered.slice(0, 50).forEach(d => {{
    const tr = document.createElement('tr');
    const status = d.DeliveryStatusSanitized || d.delivery_status;
    let badgeClass = 'badge-other';
    if (status === 'Delivered') badgeClass = 'badge-delivered';
    else if (status === 'In Transit' || status === 'Shipped') badgeClass = 'badge-transit';
    else if (status === 'Pending' || status === 'Cancelled') badgeClass = 'badge-pending';

    const p = getPrice(d);
    const m = getMileage(d);

    tr.innerHTML = `
      <td>${{d.sale_id}}</td>
      <td>${{d.SaleDateSanitized || d.sale_date}}</td>
      <td><strong>${{d.customer_name}}</strong></td>
      <td>${{d.PorscheModelSanitized || d.porsche_model}}</td>
      <td>${{d.ModelYearSanitized || d.model_year}}</td>
      <td><strong>$${{p.toLocaleString('en-US', {{ minimumFractionDigits: 2 }})}}</strong></td>
      <td>${{m.toLocaleString('en-US')}} mi</td>
      <td>${{d.PayMethodSanitized || d.payment_method}}</td>
      <td>${{d.CitySanitized || d.city}}, ${{d.StateSanitized || d.state}}</td>
      <td><span class="badge-status ${{badgeClass}}">${{status}}</span></td>
    `;
    tbody.appendChild(tr);
  }});
}}

function exportDataExcel() {{
  const exportRows = currentFiltered.map(d => ({{
    "ID Venda": d.sale_id,
    "Data Sanitizada": d.SaleDateSanitized || d.sale_date,
    "Cliente": d.customer_name,
    "Modelo Porsche": d.PorscheModelSanitized || d.porsche_model,
    "Ano Modelo": d.ModelYearSanitized || d.model_year,
    "Preco (USD)": getPrice(d),
    "Milhagem (mi)": getMileage(d),
    "Forma Pagamento": d.PayMethodSanitized || d.payment_method,
    "Cidade": d.CitySanitized || d.city,
    "Estado (UF)": d.StateSanitized || d.state,
    "Vendedor": d.salesperson,
    "Status Entrega": d.DeliveryStatusSanitized || d.delivery_status
  }}));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportRows);
  XLSX.utils.book_append_sheet(wb, ws, "Vendas_Porsche_Filtradas");
  XLSX.writeFile(wb, "porsche_vendas_export.xlsx");
}}

initFilters();
applyFilters();
</script>
</body>
</html>
"""
    output_html.write_text(html_markup, encoding="utf-8")
    print(f"OK -> HTML Dashboard criado com sucesso em: {output_html}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", default=str(INPUT_PATH))
    parser.add_argument("output", nargs="?", default=str(OUTPUT_PATH))
    parser.add_argument("output_html", nargs="?", default=str(HTML_OUTPUT_PATH))
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    out_path = Path(args.output)
    out_html_path = Path(args.output_html)

    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        return 1

    stats, records = process_workbook(in_path, out_path)
    print(f"OK -> {out_path}")
    if stats:
        print("INVALID counts per sanitized column:")
        for k, v in sorted(stats.items()):
            print(f"  {k}: {v}")
    else:
        print("No INVALID values produced.")

    generate_html_dashboard(records, out_html_path)
    return 0

if __name__ == "__main__":
    sys.exit(main())

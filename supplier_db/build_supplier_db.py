#!/usr/bin/env python3
"""Build a clean wine supplier database from Excel price lists only."""

from __future__ import annotations

import csv
import datetime as dt
import glob
import json
import re
import sqlite3
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import yaml


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "suppliers.yaml"
DB_PATH = BASE_DIR / "suppliers.db"
JSON_PATH = BASE_DIR / "suppliers.json"
REJECTED_CSV_PATH = BASE_DIR / "rejected_rows.csv"
REPORT_CSV_PATH = BASE_DIR / "build_report.csv"
ALIASES_JSON_PATH = BASE_DIR / "aliases.json"

EXCEL_SUFFIXES = {".xlsx", ".xls"}
MIN_PRICE_RUB = 500
JSON_MIN_QUALITY = 45

STOP_WORDS = {
    "wine",
    "вино",
    "wines",
    "domaine",
    "domain",
    "chateau",
    "shato",
    "brut",
    "rouge",
    "blanc",
    "doc",
    "docg",
    "igt",
    "aoc",
    "nv",
    "non",
    "vintage",
    "сух",
    "сухое",
    "крас",
    "бел",
    "белое",
    "красное",
    "розовое",
    "игристое",
    "руб",
    "цена",
    "бут",
    "л",
}

REJECT_KEYWORDS = {
    "коробка",
    "подарочная упаковка",
    "упаковка",
    "пакет",
    "набор",
    "бокал",
    "пробка",
    "штопор",
    "декантер",
    "сертификат",
    "доставка",
    "депозит",
    "тара",
    "pos",
    "gift box",
}

HARD_REJECT_KEYWORDS = {
    "подарочные коробки",
    "коробки из дерева",
    "деревянная коробка",
    "набор для вина",
    "штопор",
    "бокал",
    "декантер",
    "сертификат",
    "доставка",
    "депозит",
    "тара",
    "пробка для",
}

SERVICE_NAMES = {
    "золотая матрица",
    "ассортиментная позиция",
    "фокусный общий",
    "фокусныи общии",
    "внутренняя информация",
}

SERVICE_KEYWORDS = {
    "золотая матрица",
    "ассортиментная позиция",
    "фокусн",
}

NON_WINE_KEYWORDS = {
    "водка",
    "коньяк",
    "бренди",
    "виски",
    "whisky",
    "whiskey",
    "gin",
    "джин",
    "ром",
    "rum",
    "текила",
    "tequila",
    "ликер",
    "ликёр",
    "пиво",
    "beer",
    "сидр",
    "cider",
    "вода",
    "water",
    "сок",
    "лимонад",
    "vermouth",
    "вермут",
    "armagnac",
    "арманьяк",
}

WINE_HINTS = {
    "wine",
    "вино",
    "champagne",
    "шамп",
    "sparkling",
    "игрист",
    "brut",
    "rose",
    "роз",
    "rouge",
    "blanc",
    "крас",
    "бел",
    "aoc",
    "doc",
    "docg",
    "igt",
    "cru",
    "riesling",
    "pinot",
    "chardonnay",
    "cabernet",
    "merlot",
    "syrah",
    "sauvignon",
}


@dataclass
class SourceRule:
    supplier: str
    files: list[Path]
    active: bool = True
    updated_at: str | None = None
    columns: dict[str, Any] | None = None


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()


def cell_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return re.sub(r"\s+", " ", text)


def normalize_text(text: str) -> str:
    text = text.lower().replace("ё", "е")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_number(value: Any) -> float | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = cell_text(value)
    if not text:
        return None
    text = text.replace("\u00a0", " ").replace(" ", "")
    text = text.replace(",", ".")
    text = re.sub(r"[^0-9.]", "", text)
    if not text or text.count(".") > 1:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_price(value: Any) -> float | None:
    number = parse_number(value)
    if number is None or number <= 0:
        return None
    if number < MIN_PRICE_RUB:
        return None
    return round(number, 2)


def parse_vintage(*values: Any) -> tuple[int | None, int]:
    text = " ".join(cell_text(v) for v in values if cell_text(v))
    low = normalize_text(text)
    if re.search(r"\b(nv|n v|non vintage|б г|без года|б/г)\b", low):
        return None, 1
    years = [int(y) for y in re.findall(r"\b(20[1-3][0-9])\b", text)]
    years = [y for y in years if 2010 <= y <= 2035]
    return (years[-1], 0) if years else (None, 0)


def parse_volume(*values: Any) -> float | None:
    text = " ".join(cell_text(v) for v in values if cell_text(v))
    if not text:
        return None
    low = text.lower().replace(",", ".")
    ml = re.search(r"\b([1-9][0-9]{2,3})\s*(ml|мл)\b", low)
    if ml:
        return round(int(ml.group(1)) / 1000, 3)
    litre = re.search(r"\b([0-9]+(?:\.[0-9]+)?)\s*(l|л|литр)", low)
    if litre:
        value = float(litre.group(1))
        return round(value, 3) if 0.1 <= value <= 30 else None
    for value in re.findall(r"\b([0-9]+(?:\.[0-9]+)?)\b", low):
        number = float(value)
        if number in {0.187, 0.25, 0.33, 0.375, 0.5, 0.7, 0.75, 1.0, 1.5, 3.0, 6.0}:
            return round(number, 3)
    return None


def token_list(*texts: str) -> list[str]:
    tokens: list[str] = []
    seen: set[str] = set()
    for text in texts:
        for token in normalize_text(text).split():
            if token in STOP_WORDS:
                continue
            if token.isdigit() and len(token) != 4:
                continue
            if len(token) < 3:
                continue
            if token not in seen:
                seen.add(token)
                tokens.append(token)
    return tokens


def header_score(values: list[str]) -> int:
    joined = " ".join(normalize_text(v) for v in values)
    score = 0
    for word in ("наименование", "номенклатура", "designation", "товар назв", "название"):
        if word in joined:
            score += 8
    for word in ("цена", "прайс", "руб", "rub", "price", "бпл"):
        if word in joined:
            score += 6
    for word in ("год", "винтаж", "vintage", "урожая"):
        if word in joined:
            score += 3
    for word in ("объем", "обьем", "литраж", "volume", "емк"):
        if word in joined:
            score += 3
    return score


def find_header_row(df: pd.DataFrame) -> int | None:
    best_row: int | None = None
    best_score = 0
    for idx in range(min(len(df), 80)):
        values = [cell_text(v) for v in df.iloc[idx].tolist()]
        score = header_score(values)
        if score > best_score:
            best_score = score
            best_row = idx
    return best_row if best_score >= 10 else None


def looks_like_headerless_table(df: pd.DataFrame) -> bool:
    useful_rows = 0
    for row_idx in range(min(len(df), 80)):
        row = df.iloc[row_idx]
        has_text = any(len(cell_text(v)) >= 12 and re.search(r"[A-Za-zА-Яа-я]", cell_text(v)) for v in row.tolist())
        has_price = any(parse_price(v) is not None for v in row.tolist())
        if has_text and has_price:
            useful_rows += 1
    return useful_rows >= 3


def column_score(header: str, kind: str) -> int:
    h = normalize_text(header)
    if not h:
        return 0
    patterns = {
        "name": ["наименование", "номенклатура", "designation", "товар назв", "название"],
        "alt_name": ["наименование на английском", "латинское", "английском", "eng name", "name english"],
        "price": ["цена", "прайс", "руб", "rub", "price", "бпл", "руб бут"],
        "vintage": ["год урожая", "винтаж", "vintage", "год"],
        "volume": ["объем", "обьем", "литраж", "volume", "емк", "емкость"],
        "country": ["страна", "country"],
        "region": ["регион", "субрегион", "appellation", "аппел"],
        "color": ["цвет", "тип", "style", "сахар"],
        "producer": ["производитель", "бренд", "producer"],
        "grape": ["сорт", "виноград", "grape"],
    }
    score = 0
    for pattern in patterns[kind]:
        pattern = normalize_text(pattern)
        if pattern in h:
            score += 10 if h == pattern else 6
    if kind == "price" and any(bad in h for bad in ("остаток", "наличие", "stock")):
        score -= 8
    if kind == "name" and any(bad in h for bad in ("фото", "ссылка", "штрихкод")):
        score -= 8
    return score


def detect_columns(df: pd.DataFrame, header_row: int, manual: dict[str, Any] | None = None) -> dict[str, int | None]:
    manual = manual or {}
    headers: list[str] = []
    for col in range(df.shape[1]):
        parts = [cell_text(df.iat[header_row, col])] if header_row >= 0 else []
        if header_row >= 0 and header_row + 1 < len(df):
            nxt = cell_text(df.iat[header_row + 1, col])
            if nxt and header_score([nxt]) > 0:
                parts.append(nxt)
        headers.append(" ".join(parts))

    result: dict[str, int | None] = {}
    for kind in ("name", "alt_name", "price", "vintage", "volume", "country", "region", "color", "producer", "grape"):
        if kind in manual:
            result[kind] = int(manual[kind])
            continue
        scored = sorted(
            ((column_score(h, kind) + content_column_score(df, header_row + 1, i, kind), i) for i, h in enumerate(headers)),
            reverse=True,
        )
        result[kind] = scored[0][1] if scored and scored[0][0] > 0 else None

    if result.get("name") is None:
        result["name"] = infer_name_column(df, header_row + 1)
    if result.get("price") is None:
        result["price"] = infer_price_column(df, header_row + 1, result.get("name"))
    return result


def content_column_score(df: pd.DataFrame, start: int, col: int, kind: str) -> int:
    if kind not in {"price", "name", "volume", "vintage"}:
        return 0
    start = max(start, 0)
    score = 0
    for row in range(start, min(len(df), start + 160)):
        value = df.iat[row, col]
        text = cell_text(value)
        if kind == "price" and parse_price(value) is not None:
            score += 2
        elif kind == "name" and len(text) >= 12 and re.search(r"[A-Za-zА-Яа-я]", text):
            score += 1
        elif kind == "volume" and parse_volume(value) is not None:
            score += 1
        elif kind == "vintage" and parse_vintage(value)[0] is not None:
            score += 1
    return min(score, 20)


def infer_name_column(df: pd.DataFrame, start: int) -> int | None:
    best: tuple[int, int] | None = None
    for col in range(df.shape[1]):
        score = 0
        for row in range(start, min(len(df), start + 80)):
            text = cell_text(df.iat[row, col])
            if len(text) >= 12 and re.search(r"[A-Za-zА-Яа-я]", text):
                score += min(len(text), 80)
        if best is None or score > best[0]:
            best = (score, col)
    return best[1] if best and best[0] > 100 else None


def infer_price_column(df: pd.DataFrame, start: int, name_col: int | None) -> int | None:
    best: tuple[int, int] | None = None
    for col in range(df.shape[1]):
        if col == name_col:
            continue
        score = 0
        for row in range(start, min(len(df), start + 120)):
            price = parse_price(df.iat[row, col])
            if price is not None:
                score += 1
        if best is None or score > best[0]:
            best = (score, col)
    return best[1] if best and best[0] >= 3 else None


def is_section_like(row_values: list[str], raw_name: str, price: float | None) -> bool:
    non_empty = [v for v in row_values if v]
    if price is not None:
        return False
    if len(non_empty) <= 2 and len(raw_name) < 60:
        return True
    if raw_name.isupper() and len(raw_name.split()) <= 6:
        return True
    return False


def reject_reason(raw_name: str, price: float | None, row_values: list[str]) -> str | None:
    if not raw_name:
        return "empty_name"
    if len(raw_name) < 4:
        return "short_name"
    low = normalize_text(raw_name)
    if price is None:
        return "missing_or_invalid_price"
    if price < MIN_PRICE_RUB:
        return "price_below_minimum"
    if is_section_like(row_values, raw_name, price):
        return "section_header"
    has_wine_hint = any(hint in low for hint in WINE_HINTS)
    if low in SERVICE_NAMES or any(word in low for word in SERVICE_KEYWORDS) or any(word in low for word in HARD_REJECT_KEYWORDS):
        return "service_or_accessory"
    if any(word in low for word in REJECT_KEYWORDS) and not has_wine_hint:
        return "service_or_accessory"
    if any(word in low for word in NON_WINE_KEYWORDS) and not has_wine_hint:
        return "non_wine"
    return None


def split_producer_and_name(
    raw_name: str,
    producer_hint: str | None = None,
    alt_name: str | None = None,
) -> tuple[str | None, str | None, str | None]:
    name = re.sub(r"\s+", " ", raw_name).strip(" ,;")
    combined = " ".join(part for part in [name, cell_text(alt_name)] if part)
    producer = cell_text(producer_hint) if producer_hint else ""
    if not producer:
        chapoutier = re.search(r"\b(M\.?\s*Chapoutier|М\.?\s*Шапутье)\b", combined, re.IGNORECASE)
        if chapoutier:
            producer = chapoutier.group(1).replace("M.", "M. ").replace("М.", "М. ")
            producer = re.sub(r"\s+", " ", producer).strip()
        elif "/" in name:
            latin = name.split("/", 1)[1].split(",", 1)[0].strip()
            producer = " ".join(latin.split()[:3])
        else:
            cleaned = re.sub(r"\([^)]*\)", " ", name)
            producer = " ".join(cleaned.split(",")[0].split()[:3])
    wine_name = name
    if producer and normalize_text(wine_name).startswith(normalize_text(producer)):
        wine_name = wine_name[len(producer):].strip(" ,-/")
    cuvee = wine_name.split(",", 1)[0].strip() if wine_name else None
    return producer or None, wine_name or name, cuvee or None


def quality_score(item: dict[str, Any]) -> int:
    score = 0
    score += 25 if item.get("producer") else 0
    score += 25 if item.get("wine_name") or item.get("cuvee") else 0
    score += 15 if item.get("price") else 0
    score += 10 if item.get("vintage") or item.get("is_nv") else 0
    score += 10 if item.get("volume_l") else 0
    score += 5 if item.get("country") else 0
    score += 5 if item.get("region") or item.get("appellation") else 0
    score += 5 if item.get("color") or item.get("style") else 0
    return score


def make_match_key(item: dict[str, Any]) -> str:
    vintage = item["vintage"] if item.get("vintage") else ("nv" if item.get("is_nv") else "")
    parts = [
        item.get("producer_normalized") or "",
        normalize_text(item.get("cuvee") or item.get("wine_name") or ""),
        str(vintage),
        str(item.get("volume_l") or ""),
    ]
    return "|".join(parts)


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    with CONFIG_PATH.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def resolve_price_root(config: dict[str, Any]) -> Path:
    env_root = sys.argv[1] if len(sys.argv) > 1 else None
    env_root = env_root or __import__("os").environ.get("SUPPLIER_PRICELIST_DIR")
    root = env_root or config.get("price_root") or "../wine_finder_bot/pricelists"
    path = Path(root).expanduser()
    if not path.is_absolute():
        path = (BASE_DIR / path).resolve()
    return path


def discover_sources(config: dict[str, Any]) -> list[SourceRule]:
    root = resolve_price_root(config)
    if not root.exists():
        raise FileNotFoundError(f"Price-list directory not found: {root}")

    suppliers = config.get("suppliers") or []
    sources: list[SourceRule] = []
    if suppliers:
        for entry in suppliers:
            if not entry.get("active", True):
                continue
            files: list[Path] = []
            for pattern in entry.get("files", []):
                for match in glob.glob(str(root / pattern)):
                    path = Path(match)
                    if path.suffix.lower() in EXCEL_SUFFIXES:
                        files.append(path)
            if files:
                sources.append(
                    SourceRule(
                        supplier=entry.get("name") or files[0].stem,
                        files=sorted(set(files)),
                        active=True,
                        updated_at=entry.get("updated_at"),
                        columns=entry.get("columns") or {},
                    )
                )
        return sources

    patterns = config.get("include_patterns") or ["*.xlsx", "*.xls"]
    files: set[Path] = set()
    for pattern in patterns:
        for match in glob.glob(str(root / pattern)):
            path = Path(match)
            if path.suffix.lower() in EXCEL_SUFFIXES:
                files.add(path)
    return [SourceRule(supplier=path.stem, files=[path]) for path in sorted(files)]


def create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS supplier_items;
        DROP TABLE IF EXISTS rejected_rows;
        DROP TABLE IF EXISTS aliases;

        CREATE TABLE supplier_items (
          id INTEGER PRIMARY KEY,
          supplier TEXT NOT NULL,
          source_file TEXT NOT NULL,
          source_sheet TEXT,
          source_row INTEGER,
          raw_name TEXT NOT NULL,
          raw_price TEXT,
          price REAL,
          currency TEXT DEFAULT 'RUB',
          producer TEXT,
          producer_normalized TEXT,
          wine_name TEXT,
          cuvee TEXT,
          vintage INTEGER,
          is_nv INTEGER DEFAULT 0,
          volume_l REAL,
          country TEXT,
          region TEXT,
          appellation TEXT,
          color TEXT,
          style TEXT,
          grape_varieties TEXT,
          normalized_name TEXT NOT NULL,
          search_tokens TEXT NOT NULL,
          match_key TEXT,
          quality_score INTEGER DEFAULT 0,
          created_at TEXT NOT NULL,
          updated_at TEXT NOT NULL
        );

        CREATE TABLE rejected_rows (
          id INTEGER PRIMARY KEY,
          supplier TEXT,
          source_file TEXT,
          source_sheet TEXT,
          source_row INTEGER,
          raw_values TEXT,
          reason TEXT NOT NULL,
          created_at TEXT NOT NULL
        );

        CREATE TABLE aliases (
          id INTEGER PRIMARY KEY,
          kind TEXT NOT NULL,
          alias TEXT NOT NULL,
          canonical TEXT NOT NULL,
          source TEXT
        );
        """
    )
    conn.execute("CREATE INDEX idx_supplier_items_match_key ON supplier_items(match_key)")
    conn.execute("CREATE INDEX idx_supplier_items_supplier ON supplier_items(supplier)")


def insert_item(conn: sqlite3.Connection, item: dict[str, Any]) -> None:
    columns = [
        "supplier",
        "source_file",
        "source_sheet",
        "source_row",
        "raw_name",
        "raw_price",
        "price",
        "currency",
        "producer",
        "producer_normalized",
        "wine_name",
        "cuvee",
        "vintage",
        "is_nv",
        "volume_l",
        "country",
        "region",
        "appellation",
        "color",
        "style",
        "grape_varieties",
        "normalized_name",
        "search_tokens",
        "match_key",
        "quality_score",
        "created_at",
        "updated_at",
    ]
    conn.execute(
        f"INSERT INTO supplier_items ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)})",
        [item.get(col) for col in columns],
    )


def insert_rejected(conn: sqlite3.Connection, row: dict[str, Any]) -> None:
    conn.execute(
        """
        INSERT INTO rejected_rows
        (supplier, source_file, source_sheet, source_row, raw_values, reason, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        [
            row["supplier"],
            row["source_file"],
            row["source_sheet"],
            row["source_row"],
            row["raw_values"],
            row["reason"],
            row["created_at"],
        ],
    )


def row_value(row: pd.Series, col: int | None) -> Any:
    if col is None or col >= len(row):
        return None
    return row.iloc[col]


def process_sheet(conn: sqlite3.Connection, source: SourceRule, path: Path, sheet_name: str) -> dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    df = df.dropna(how="all")
    stats = {
        "supplier": source.supplier,
        "source_file": path.name,
        "source_sheet": sheet_name,
        "rows_total": int(len(df)),
        "rows_accepted": 0,
        "rows_rejected": 0,
        "without_price": 0,
        "without_producer": 0,
        "without_vintage_or_nv": 0,
        "without_volume": 0,
        "quality_scores": [],
        "notes": "",
    }
    if df.empty:
        stats["notes"] = "empty_sheet"
        return stats
    header_row = find_header_row(df)
    if header_row is None:
        if looks_like_headerless_table(df):
            header_row = -1
            stats["notes"] = "headerless_table"
        else:
            stats["notes"] = "header_not_found"
            return stats
    columns = detect_columns(df, header_row, source.columns)
    if columns.get("name") is None or columns.get("price") is None:
        stats["notes"] = "required_columns_not_found"
        return stats

    seen_keys: dict[str, float] = {}
    created_at = now_iso()
    for row_idx in range(max(header_row + 1, 0), len(df)):
        row = df.iloc[row_idx]
        row_values = [cell_text(v) for v in row.tolist()]
        if not any(row_values):
            continue

        raw_name = cell_text(row_value(row, columns.get("name")))
        raw_price = cell_text(row_value(row, columns.get("price")))
        price = parse_price(raw_price)
        reason = reject_reason(raw_name, price, row_values)
        if reason:
            stats["rows_rejected"] += 1
            if reason == "missing_or_invalid_price":
                stats["without_price"] += 1
            insert_rejected(
                conn,
                {
                    "supplier": source.supplier,
                    "source_file": path.name,
                    "source_sheet": sheet_name,
                    "source_row": int(row_idx + 1),
                    "raw_values": json.dumps(row_values, ensure_ascii=False),
                    "reason": reason,
                    "created_at": created_at,
                },
            )
            continue

        producer_hint = cell_text(row_value(row, columns.get("producer")))
        alt_name = cell_text(row_value(row, columns.get("alt_name")))
        producer, wine_name, cuvee = split_producer_and_name(raw_name, producer_hint, alt_name)
        vintage, is_nv = parse_vintage(row_value(row, columns.get("vintage")), raw_name)
        volume_l = parse_volume(row_value(row, columns.get("volume")), raw_name)
        country = cell_text(row_value(row, columns.get("country"))) or None
        region = cell_text(row_value(row, columns.get("region"))) or None
        color = cell_text(row_value(row, columns.get("color"))) or None
        grape = cell_text(row_value(row, columns.get("grape"))) or None
        indexed_name = " ".join(part for part in [raw_name, alt_name] if part)
        normalized_name = normalize_text(indexed_name)
        search_tokens = token_list(indexed_name, producer or "", wine_name or "", country or "", region or "")
        item = {
            "supplier": source.supplier,
            "source_file": path.name,
            "source_sheet": sheet_name,
            "source_row": int(row_idx + 1),
            "raw_name": raw_name,
            "raw_price": raw_price,
            "price": price,
            "currency": "RUB",
            "producer": producer,
            "producer_normalized": normalize_text(producer or ""),
            "wine_name": wine_name,
            "cuvee": cuvee,
            "vintage": vintage,
            "is_nv": is_nv,
            "volume_l": volume_l,
            "country": country,
            "region": region,
            "appellation": region,
            "color": color,
            "style": color,
            "grape_varieties": grape,
            "normalized_name": normalized_name,
            "search_tokens": json.dumps(search_tokens, ensure_ascii=False),
            "created_at": created_at,
            "updated_at": created_at,
        }
        item["quality_score"] = quality_score(item)
        item["match_key"] = make_match_key(item)

        dedupe_key = f"{item['supplier']}|{item['match_key']}"
        if dedupe_key in seen_keys and price is not None and seen_keys[dedupe_key] <= price:
            continue
        seen_keys[dedupe_key] = float(price or 0)

        if not item["producer"]:
            stats["without_producer"] += 1
        if not item["vintage"] and not item["is_nv"]:
            stats["without_vintage_or_nv"] += 1
        if not item["volume_l"]:
            stats["without_volume"] += 1
        stats["rows_accepted"] += 1
        stats["quality_scores"].append(item["quality_score"])
        insert_item(conn, item)
    return stats


def write_rejected_csv(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        "SELECT supplier, source_file, source_sheet, source_row, raw_values, reason, created_at FROM rejected_rows"
    ).fetchall()
    with REJECTED_CSV_PATH.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["supplier", "source_file", "source_sheet", "source_row", "raw_values", "reason", "created_at"])
        writer.writerows(rows)


def write_report_csv(report_rows: list[dict[str, Any]]) -> None:
    fields = [
        "supplier",
        "source_file",
        "source_sheet",
        "rows_total",
        "rows_accepted",
        "rows_rejected",
        "without_price",
        "without_producer",
        "without_vintage_or_nv",
        "without_volume",
        "avg_quality_score",
        "notes",
    ]
    with REPORT_CSV_PATH.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        for row in report_rows:
            scores = row.pop("quality_scores", [])
            row["avg_quality_score"] = round(sum(scores) / len(scores), 1) if scores else 0
            writer.writerow({field: row.get(field, "") for field in fields})


def write_json_export(conn: sqlite3.Connection) -> None:
    db_rows = conn.execute(
        """
        SELECT id, supplier, producer, wine_name, cuvee, vintage, is_nv, volume_l, country,
               region, appellation, color, price, currency, raw_name, normalized_name,
               search_tokens, quality_score
        FROM supplier_items
        WHERE quality_score >= ? AND supplier <> '' AND raw_name <> '' AND price IS NOT NULL
        ORDER BY supplier, normalized_name, price
        """,
        [JSON_MIN_QUALITY],
    ).fetchall()
    items = []
    for row in db_rows:
        tokens = json.loads(row[16] or "[]")
        item = {
            "id": row[0],
            "supplier": row[1],
            "company": row[1],
            "producer": row[2],
            "wine_name": row[3],
            "name": row[3] or row[14],
            "name_en": row[14],
            "cuvee": row[4],
            "vintage": row[5],
            "year": row[5],
            "is_nv": bool(row[6]),
            "volume_l": row[7],
            "volume": row[7],
            "country": row[8],
            "region": row[9],
            "appellation": row[10],
            "color": row[11],
            "price": row[12],
            "currency": row[13],
            "raw_name": row[14],
            "raw": row[14],
            "normalized_name": row[15],
            "search_tokens": tokens,
            "tokens": tokens,
        }
        items.append(item)
    payload = {
        "version": 1,
        "generated_at": now_iso(),
        "count": len(items),
        "suppliers": sorted({item["supplier"] for item in items}),
        "companies": sorted({item["supplier"] for item in items}),
        "items": items,
    }
    JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_aliases(config: dict[str, Any], conn: sqlite3.Connection) -> None:
    aliases = []
    for kind, mapping in (config.get("aliases") or {}).items():
        if isinstance(mapping, dict):
            for alias, canonical in mapping.items():
                aliases.append({"kind": kind, "alias": alias, "canonical": canonical, "source": "manual"})
                conn.execute(
                    "INSERT INTO aliases (kind, alias, canonical, source) VALUES (?, ?, ?, ?)",
                    [kind, alias, canonical, "manual"],
                )
    ALIASES_JSON_PATH.write_text(json.dumps(aliases, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    config = load_config()
    sources = discover_sources(config)
    if not sources:
        print("No Excel price lists found.", file=sys.stderr)
        return 2

    if DB_PATH.exists():
        DB_PATH.unlink()
    conn = sqlite3.connect(DB_PATH)
    create_schema(conn)
    report_rows: list[dict[str, Any]] = []

    for source in sources:
        for path in source.files:
            if path.suffix.lower() not in EXCEL_SUFFIXES:
                continue
            print(f"Importing {path.name} as {source.supplier}...")
            try:
                xl = pd.ExcelFile(path)
            except Exception as exc:  # noqa: BLE001
                report_rows.append(
                    {
                        "supplier": source.supplier,
                        "source_file": path.name,
                        "source_sheet": "",
                        "rows_total": 0,
                        "rows_accepted": 0,
                        "rows_rejected": 0,
                        "without_price": 0,
                        "without_producer": 0,
                        "without_vintage_or_nv": 0,
                        "without_volume": 0,
                        "quality_scores": [],
                        "notes": f"open_error: {exc}",
                    }
                )
                continue
            for sheet_name in xl.sheet_names:
                try:
                    stats = process_sheet(conn, source, path, sheet_name)
                except Exception as exc:  # noqa: BLE001
                    stats = {
                        "supplier": source.supplier,
                        "source_file": path.name,
                        "source_sheet": sheet_name,
                        "rows_total": 0,
                        "rows_accepted": 0,
                        "rows_rejected": 0,
                        "without_price": 0,
                        "without_producer": 0,
                        "without_vintage_or_nv": 0,
                        "without_volume": 0,
                        "quality_scores": [],
                        "notes": f"sheet_error: {exc}",
                    }
                report_rows.append(stats)
            conn.commit()

    write_aliases(config, conn)
    write_rejected_csv(conn)
    write_report_csv(report_rows)
    write_json_export(conn)
    conn.commit()
    total = conn.execute("SELECT COUNT(*) FROM supplier_items").fetchone()[0]
    exported = json.loads(JSON_PATH.read_text(encoding="utf-8"))["count"]
    suppliers = conn.execute("SELECT COUNT(DISTINCT supplier) FROM supplier_items").fetchone()[0]
    conn.close()
    print(f"Done: {total} SQLite rows, {exported} JSON rows, {suppliers} suppliers.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

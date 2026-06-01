from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# M5 — Средства федерального казначейства
#
# Скрипт НЕ скачивает данные.
#
# Источники:
# 1. data/m5/cbr_sors/              — Excel ЦБ SORS
# 2. data/m5/roskazna_deposits/     — DOCX/XML Росказны
# 3. data/m5/cbr_bliquidity/        — ground truth ЦБ
#
# На выходе:
# data/m5/result/m5_treasury_signals.xlsx
# data/m5/result/m5_treasury_flow.png
# ============================================================


CBR_SORS_DIR = Path("data/m5/cbr_sors")

ROSKAZNA_DIR = Path("data/m5/roskazna_deposits")
ROSKAZNA_INDEX_FILE = ROSKAZNA_DIR / "files_index.xlsx"
ROSKAZNA_PARSED_FILE = ROSKAZNA_DIR / "roskazna_deposits_parsed.xlsx"

CBR_BLIQUIDITY_DIR = Path("data/m5/cbr_bliquidity")

RESULT_DIR = Path("data/m5/result")
OUTPUT_FILE = RESULT_DIR / "m5_treasury_signals.xlsx"
PLOT_FILE = RESULT_DIR / "m5_treasury_flow.png"

DEBUG_CBR_SORS_FILE = RESULT_DIR / "debug_cbr_sors_matches.xlsx"
DEBUG_ROSKAZNA_FILE = RESULT_DIR / "debug_roskazna_parsed.xlsx"

MAD_WINDOW_MONTHLY = 36
MIN_PERIODS_MONTHLY = 12

MAD_WINDOW_WEEKLY = 156
MIN_PERIODS_WEEKLY = 20

BUDGET_DRAIN_THRESHOLD = -300.0


# ============================================================
# 0. Общие функции
# ============================================================


def clean_text(value) -> str:
    if pd.isna(value):
        return ""

    value = str(value).lower()
    value = value.replace("\xa0", " ")
    value = value.replace("ё", "е")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def clean_number(value):
    if value is None or pd.isna(value):
        return None

    value = str(value).strip()

    if value in ("", "-", "—", "nan", "None", "NaN"):
        return None

    value = value.replace("\xa0", "")
    value = value.replace(" ", "")
    value = value.replace(",", ".")

    value = re.sub(r"[^0-9.\-]", "", value)

    if value in ("", "-", ".", "-."):
        return None

    try:
        return float(value)
    except ValueError:
        return None


def normalize_to_bln(series: pd.Series) -> pd.Series:
    series = pd.to_numeric(series, errors="coerce")
    median_value = series.abs().median()

    if pd.isna(median_value) or median_value == 0:
        return series

    # рубли
    if median_value > 10_000_000:
        return series / 1_000_000_000

    # млн руб.
    if median_value > 10_000:
        return series / 1_000

    # уже млрд руб.
    return series


def normalize_amount_to_bln(value):
    value = clean_number(value)

    if value is None:
        return None

    # рубли
    if value > 10_000_000:
        value = value / 1_000_000_000

    # млн руб.
    elif value > 10_000:
        value = value / 1_000

    return value


def extract_date_from_filename(file_path: Path):
    name = file_path.name

    match_yyyymmdd = re.search(r"(20\d{2})(\d{2})(\d{2})", name)

    if match_yyyymmdd:
        return pd.Timestamp(
            year=int(match_yyyymmdd.group(1)),
            month=int(match_yyyymmdd.group(2)),
            day=int(match_yyyymmdd.group(3)),
        )

    match_ddmmyyyy = re.search(r"(\d{2})[-.](\d{2})[-.](20\d{2})", name)

    if match_ddmmyyyy:
        return pd.Timestamp(
            year=int(match_ddmmyyyy.group(3)),
            month=int(match_ddmmyyyy.group(2)),
            day=int(match_ddmmyyyy.group(1)),
        )

    return pd.NaT


def calculate_mad_score(series: pd.Series, window: int, min_periods: int) -> pd.Series:
    series = pd.to_numeric(series, errors="coerce")

    rolling_median = series.rolling(window=window, min_periods=min_periods).median()

    rolling_mad = series.rolling(window=window, min_periods=min_periods).apply(
        lambda x: (abs(x - x.median())).median(),
        raw=False,
    )

    score = (series - rolling_median) / (1.4826 * rolling_mad)
    score = score.replace([float("inf"), float("-inf")], pd.NA)

    return score


def sum_with_na(series: pd.Series):
    """
    Важно:
    NA = данных нет.
    0 = данные есть, но размещение было нулевым / отбор не состоялся.

    Поэтому нельзя обычный sum() с последующей заменой 0 на NA.
    """
    series = pd.to_numeric(series, errors="coerce")

    if series.notna().sum() == 0:
        return pd.NA

    return series.sum()


def ensure_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for col in columns:
        if col not in df.columns:
            df[col] = pd.NA

    return df


# ============================================================
# 1. ЦБ SORS
# ============================================================


def get_numeric_values_from_row(row_values: list) -> list[float]:
    numbers = []

    for value in row_values:
        number = clean_number(value)

        if number is not None:
            numbers.append(number)

    return numbers


def row_has_budget_keyword(row_text: str) -> bool:
    return (
        ("федеральн" in row_text and "бюджет" in row_text)
        or ("внебюджетн" in row_text and "фонд" in row_text)
        or ("государственн" in row_text and "внебюджетн" in row_text)
    )


def parse_cbr_sors_file(file_path: Path) -> tuple[dict | None, list[dict]]:
    date_value = extract_date_from_filename(file_path)
    debug_rows = []

    try:
        excel = pd.ExcelFile(file_path)
    except Exception as error:
        print(f"{file_path.name}: не удалось открыть Excel")
        print(error)
        return None, debug_rows

    best_result = None

    for sheet_name in excel.sheet_names:
        try:
            raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        except Exception:
            continue

        if raw.empty:
            continue

        federal_value = None
        extra_budget_value = None

        for idx in range(len(raw)):
            row_values = raw.iloc[idx].tolist()
            row_text = " ".join(clean_text(x) for x in row_values)

            if not row_has_budget_keyword(row_text):
                continue

            numbers = get_numeric_values_from_row(row_values)

            debug_rows.append(
                {
                    "file": file_path.name,
                    "sheet": sheet_name,
                    "row": idx,
                    "row_text": row_text[:500],
                    "numbers": str(numbers[:20]),
                }
            )

            if not numbers:
                continue

            # Берём последнее числовое значение в строке.
            value = numbers[-1]

            is_federal = "федеральн" in row_text and "бюджет" in row_text
            is_extra = (
                ("внебюджетн" in row_text and "фонд" in row_text)
                or ("государственн" in row_text and "внебюджетн" in row_text)
            )

            if is_federal and federal_value is None:
                federal_value = value

            if is_extra and extra_budget_value is None:
                extra_budget_value = value

        if federal_value is not None or extra_budget_value is not None:
            total = (federal_value or 0) + (extra_budget_value or 0)

            best_result = {
                "date": date_value,
                "cbr_federal_budget_funds": federal_value,
                "cbr_extra_budgetary_funds": extra_budget_value,
                "cbr_budget_funds_total": total,
                "source_file": file_path.name,
                "source_sheet": sheet_name,
            }

            break

    if best_result is None:
        print(f"{file_path.name}: не нашёл строки с бюджетными средствами")

    return best_result, debug_rows


def load_cbr_sors_budget_funds() -> tuple[pd.DataFrame, pd.DataFrame]:
    files = sorted(CBR_SORS_DIR.glob("*.xlsx"))

    if not files:
        print(f"В папке {CBR_SORS_DIR} нет Excel-файлов ЦБ SORS.")
        return pd.DataFrame(), pd.DataFrame()

    rows = []
    debug_rows_all = []

    print("Читаю Excel ЦБ SORS...")

    for file_path in files:
        result, debug_rows = parse_cbr_sors_file(file_path)

        debug_rows_all.extend(debug_rows)

        if result is not None:
            rows.append(result)

    debug_df = pd.DataFrame(debug_rows_all)

    if not rows:
        print("Не удалось извлечь бюджетные остатки из файлов ЦБ SORS.")
        return pd.DataFrame(), debug_df

    df = pd.DataFrame(rows)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()

    for col in [
        "cbr_federal_budget_funds",
        "cbr_extra_budgetary_funds",
        "cbr_budget_funds_total",
    ]:
        df[col] = normalize_to_bln(df[col])

    df = df.sort_values("date").reset_index(drop=True)

    print(f"ЦБ SORS: извлечено строк: {len(df)}")

    return df, debug_df


def build_cbr_monthly(cbr_df: pd.DataFrame) -> pd.DataFrame:
    if cbr_df.empty:
        return pd.DataFrame(
            columns=[
                "date",
                "cbr_federal_budget_funds",
                "cbr_extra_budgetary_funds",
                "cbr_budget_funds_total",
            ]
        )

    cbr_monthly = cbr_df.copy()
    cbr_monthly["date"] = cbr_monthly["date"].dt.to_period("M").dt.to_timestamp("M")

    cbr_monthly = (
        cbr_monthly.groupby("date", as_index=False)
        .agg(
            cbr_federal_budget_funds=("cbr_federal_budget_funds", "last"),
            cbr_extra_budgetary_funds=("cbr_extra_budgetary_funds", "last"),
            cbr_budget_funds_total=("cbr_budget_funds_total", "last"),
        )
    )

    return cbr_monthly


# ============================================================
# 2. Росказна
# ============================================================


def normalize_path_for_current_os(value: str | Path) -> Path:
    """
    Исправляет Windows-пути из Excel:
    data\\m5\\... -> data/m5/...
    """
    text = str(value).strip().strip('"').strip("'")
    text = text.replace("\\", "/")
    return Path(text)


def parse_russian_date(value):
    if value is None or pd.isna(value):
        return pd.NaT

    text = str(value).strip().lower().replace(",", " ")
    text = re.sub(r"\s+", " ", text)

    dot_match = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)

    if dot_match:
        return pd.Timestamp(
            year=int(dot_match.group(3)),
            month=int(dot_match.group(2)),
            day=int(dot_match.group(1)),
        )

    month_map = {
        "января": 1,
        "февраля": 2,
        "марта": 3,
        "апреля": 4,
        "мая": 5,
        "июня": 6,
        "июля": 7,
        "августа": 8,
        "сентября": 9,
        "октября": 10,
        "ноября": 11,
        "декабря": 12,
    }

    word_match = re.search(
        r"(\d{1,2})\s+"
        r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
        r"\s+(\d{4})",
        text,
    )

    if word_match:
        return pd.Timestamp(
            year=int(word_match.group(3)),
            month=month_map[word_match.group(2)],
            day=int(word_match.group(1)),
        )

    return pd.NaT


def read_docx_text(file_path: Path) -> str:
    try:
        with zipfile.ZipFile(file_path) as archive:
            xml_bytes = archive.read("word/document.xml")

        root = ET.fromstring(xml_bytes)
        texts = []

        for elem in root.iter():
            if elem.tag.endswith("}t") and elem.text:
                texts.append(elem.text)

        return " ".join(texts)
    except Exception:
        return ""


def read_xml_text(file_path: Path) -> str:
    try:
        raw_bytes = file_path.read_bytes()
    except Exception:
        return ""

    for encoding in ["utf-8", "windows-1251", "cp1251"]:
        try:
            return raw_bytes.decode(encoding, errors="ignore")
        except Exception:
            continue

    return ""


def mln_rub_to_bln(value):
    number = clean_number(value)

    if number is None:
        return None

    return number / 1000.0


def get_xml_child_text(node: ET.Element, tag_name: str):
    for child in list(node):
        clean_tag = child.tag.split("}")[-1].lower()

        if clean_tag == tag_name.lower():
            return child.text

    return None


def parse_roskazna_xml_structured(file_path: Path, fallback_date=None) -> list[dict]:
    """
    Парсим XML Росказны по реальным тегам.

    В XML денежные значения идут в млн руб.:
    maxvol      — объявленный лимит;
    totalbid    — спрос банков;
    totalaccept — принятый объём;
    totalsettle — фактическое размещение.

    Для M5 используем placed_volume_bln:
    1. totalsettle, если есть;
    2. totalaccept, если totalsettle нет;
    3. 0, если отбор несостоявшийся;
    4. maxvol как proxy для свежих объявленных отборов.
    """
    if not file_path.exists():
        return []

    text = read_xml_text(file_path)

    if not text:
        return []

    try:
        root = ET.fromstring(text)
    except Exception:
        return []

    auction_nodes = []

    for node in root.iter():
        clean_tag = node.tag.split("}")[-1].lower()

        if clean_tag.startswith("depoauc"):
            auction_nodes.append(node)

    if not auction_nodes:
        auction_nodes = [root]

    rows = []

    for node in auction_nodes:
        aucdate = parse_russian_date(get_xml_child_text(node, "aucdate"))

        if pd.isna(aucdate):
            aucdate = fallback_date

        maxvol_bln = mln_rub_to_bln(get_xml_child_text(node, "maxvol"))
        totalbid_bln = mln_rub_to_bln(get_xml_child_text(node, "totalbid"))
        totalaccept_bln = mln_rub_to_bln(get_xml_child_text(node, "totalaccept"))
        totalsettle_bln = mln_rub_to_bln(get_xml_child_text(node, "totalsettle"))

        comment = get_xml_child_text(node, "Comment") or ""
        failed = "несостояв" in comment.lower()

        if totalsettle_bln is not None:
            placed_bln = totalsettle_bln
            volume_source = "totalsettle"
        elif totalaccept_bln is not None:
            placed_bln = totalaccept_bln
            volume_source = "totalaccept"
        elif failed:
            placed_bln = 0.0
            volume_source = "failed_auction_zero"
        else:
            placed_bln = maxvol_bln
            volume_source = "maxvol_proxy"

        rows.append(
            {
                "date": aucdate,
                "auction_id": get_xml_child_text(node, "id"),
                "funds_placed": get_xml_child_text(node, "FundsPlaced"),
                "term_days": clean_number(get_xml_child_text(node, "term")),
                "firstdate": parse_russian_date(get_xml_child_text(node, "firstdate")),
                "seconddate": parse_russian_date(get_xml_child_text(node, "seconddate")),
                "min_rate": clean_number(get_xml_child_text(node, "minrate")),
                "cutoff_rate": clean_number(get_xml_child_text(node, "cutoffrate")),
                "wa_accept_rate": clean_number(get_xml_child_text(node, "waacceptrate")),
                "maxvol_bln": maxvol_bln,
                "totalbid_bln": totalbid_bln,
                "totalaccept_bln": totalaccept_bln,
                "totalsettle_bln": totalsettle_bln,
                "placed_volume_bln": placed_bln,
                "volume_source": volume_source,
                "cr_bidders": clean_number(get_xml_child_text(node, "crbidders")),
                "accept_cr_bidders": clean_number(get_xml_child_text(node, "acceptcrbidders")),
                "place": get_xml_child_text(node, "place"),
                "comment": comment,
                "is_failed_auction": int(failed),
                "source_file_path": str(file_path.as_posix()),
                "parse_status": "ok" if placed_bln is not None else "amount_not_found",
            }
        )

    return rows


def extract_amount_candidates_from_text(text: str) -> list[dict]:
    """
    Fallback для DOCX/текста, если XML-структура изменилась.
    """
    if not text:
        return []

    text_norm = text.replace("\xa0", " ")
    text_norm = re.sub(r"\s+", " ", text_norm)

    candidates = []

    patterns = [
        r"(?:предельн\w*\s+)?(?:объем|объём|сумма|лимит)\s+(?:средств\s+)?(?:к\s+)?(?:размещени\w*|депозит\w*)?.{0,160}?(\d[\d\s.,]{2,})\s*(?:млн\s*руб|руб|российск\w*\s+рубл)",
        r"(?:maxvol|totalsettle|totalaccept|totalbid)[^0-9]{0,80}(\d[\d\s.,]{2,})",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text_norm, flags=re.IGNORECASE):
            raw_value = match.group(1)
            value_bln = normalize_amount_to_bln(raw_value)

            if value_bln is None:
                continue

            if 0 <= value_bln <= 10_000:
                candidates.append(
                    {
                        "value_bln": value_bln,
                        "raw_value": raw_value,
                        "context": text_norm[max(0, match.start() - 100): match.end() + 100],
                    }
                )

    return candidates


def extract_banks_count_from_text(text: str):
    if not text:
        return None

    text_norm = text.replace("\xa0", " ")
    text_norm = re.sub(r"\s+", " ", text_norm)

    patterns = [
        r"<crbidders>\s*(\d{1,3})\s*</crbidders>",
        r"<acceptcrbidders>\s*(\d{1,3})\s*</acceptcrbidders>",
        r"количеств\w*\s+кредитн\w*\s+организац\w*.{0,80}?(\d{1,3})",
        r"(\d{1,3})\s+кредитн\w*\s+организац\w*",
        r"количеств\w*\s+банк\w*.{0,80}?(\d{1,3})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text_norm, flags=re.IGNORECASE)

        if match:
            value = clean_number(match.group(1))

            if value is not None and 0 < value <= 500:
                return int(value)

    return None


def parse_roskazna_file(row: pd.Series) -> list[dict]:
    file_path = normalize_path_for_current_os(row.get("local_file_path", ""))

    if not file_path.exists():
        return [
            {
                "date": row.get("date"),
                "roskazna_deposit_volume": None,
                "roskazna_planned_volume": None,
                "roskazna_total_bid": None,
                "roskazna_total_accept": None,
                "roskazna_banks_count": None,
                "roskazna_file_name": row.get("file_name"),
                "roskazna_file_type": row.get("file_type"),
                "roskazna_file_path": str(file_path),
                "parse_status": "file_not_found",
                "amount_context": None,
                "volume_source": None,
            }
        ]

    file_type = str(row.get("file_type", "")).lower()

    if file_type == "xml":
        xml_rows = parse_roskazna_xml_structured(file_path, fallback_date=row.get("date"))

        if xml_rows:
            result_rows = []

            for item in xml_rows:
                result_rows.append(
                    {
                        "date": item.get("date"),
                        "roskazna_deposit_volume": item.get("placed_volume_bln"),
                        "roskazna_planned_volume": item.get("maxvol_bln"),
                        "roskazna_total_bid": item.get("totalbid_bln"),
                        "roskazna_total_accept": item.get("totalaccept_bln"),
                        "roskazna_banks_count": item.get("accept_cr_bidders") or item.get("cr_bidders"),
                        "roskazna_file_name": row.get("file_name"),
                        "roskazna_file_type": row.get("file_type"),
                        "roskazna_file_path": str(file_path.as_posix()),
                        "parse_status": item.get("parse_status"),
                        "amount_context": item.get("comment"),
                        "volume_source": item.get("volume_source"),
                        "auction_id": item.get("auction_id"),
                        "term_days": item.get("term_days"),
                        "is_failed_auction": item.get("is_failed_auction"),
                    }
                )

            return result_rows

    # Fallback для DOCX или нестандартного XML.
    if file_type == "docx":
        text = read_docx_text(file_path)
    elif file_type == "xml":
        text = read_xml_text(file_path)
    else:
        text = ""

    if not text:
        return [
            {
                "date": row.get("date"),
                "roskazna_deposit_volume": None,
                "roskazna_planned_volume": None,
                "roskazna_total_bid": None,
                "roskazna_total_accept": None,
                "roskazna_banks_count": None,
                "roskazna_file_name": row.get("file_name"),
                "roskazna_file_type": row.get("file_type"),
                "roskazna_file_path": str(file_path.as_posix()),
                "parse_status": "empty_text",
                "amount_context": None,
                "volume_source": None,
            }
        ]

    candidates = extract_amount_candidates_from_text(text)

    if candidates:
        best_candidate = max(candidates, key=lambda x: x["value_bln"])
        amount = best_candidate["value_bln"]
        context = best_candidate["context"]
    else:
        amount = None
        context = None

    banks_count = extract_banks_count_from_text(text)

    return [
        {
            "date": row.get("date"),
            "roskazna_deposit_volume": amount,
            "roskazna_planned_volume": None,
            "roskazna_total_bid": None,
            "roskazna_total_accept": None,
            "roskazna_banks_count": banks_count,
            "roskazna_file_name": row.get("file_name"),
            "roskazna_file_type": row.get("file_type"),
            "roskazna_file_path": str(file_path.as_posix()),
            "parse_status": "ok" if amount is not None else "amount_not_found",
            "amount_context": context,
            "volume_source": "fallback_text",
        }
    ]


def load_roskazna_deposits() -> pd.DataFrame:
    if ROSKAZNA_PARSED_FILE.exists():
        print(f"Читаю готовый парсинг Росказны: {ROSKAZNA_PARSED_FILE}")

        parsed = pd.read_excel(ROSKAZNA_PARSED_FILE)
        parsed["date"] = pd.to_datetime(parsed["date"], errors="coerce")

        rename_map = {
            "placed_volume_bln": "roskazna_deposit_volume",
            "maxvol_bln": "roskazna_planned_volume",
            "totalbid_bln": "roskazna_total_bid",
            "totalaccept_bln": "roskazna_total_accept",
            "accept_cr_bidders": "roskazna_banks_count",
            "file_name": "roskazna_file_name",
            "source_file_path": "roskazna_file_path",
        }

        parsed = parsed.rename(columns=rename_map)

        for col in [
            "roskazna_deposit_volume",
            "roskazna_planned_volume",
            "roskazna_total_bid",
            "roskazna_total_accept",
            "roskazna_banks_count",
        ]:
            if col not in parsed.columns:
                parsed[col] = pd.NA

            parsed[col] = pd.to_numeric(parsed[col], errors="coerce")

        if "roskazna_file_type" not in parsed.columns:
            parsed["roskazna_file_type"] = "xml"

        if "parse_status" not in parsed.columns:
            parsed["parse_status"] = "ok"

        if "auction_id" not in parsed.columns:
            parsed["auction_id"] = pd.NA

        if "term_days" not in parsed.columns:
            parsed["term_days"] = pd.NA

        parsed = parsed.dropna(subset=["date"]).copy()
        parsed = parsed.sort_values(["date", "auction_id", "term_days"], na_position="last")

        print(f"Росказна: строк из готового парсинга: {len(parsed)}")
        print(f"Росказна: найдено размещений: {parsed['roskazna_deposit_volume'].notna().sum()}")

        return parsed.reset_index(drop=True)

    if not ROSKAZNA_INDEX_FILE.exists():
        print(f"Не найден файл реестра Росказны: {ROSKAZNA_INDEX_FILE}")
        return pd.DataFrame()

    print("Читаю реестр и файлы Росказны...")

    index_df = pd.read_excel(ROSKAZNA_INDEX_FILE)
    index_df["date"] = pd.to_datetime(index_df["date"], errors="coerce")

    for col in ["file_type", "selection_number", "local_file_path"]:
        if col not in index_df.columns:
            index_df[col] = ""

    index_df["priority"] = (
        index_df["file_type"]
        .astype(str)
        .str.lower()
        .map({"xml": 1, "docx": 2})
        .fillna(9)
    )

    index_df = index_df.sort_values(["date", "selection_number", "priority"])

    rows = []

    for _, row in index_df.iterrows():
        rows.extend(parse_roskazna_file(row))

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()

    for col in [
        "roskazna_deposit_volume",
        "roskazna_planned_volume",
        "roskazna_total_bid",
        "roskazna_total_accept",
        "roskazna_banks_count",
    ]:
        if col not in df.columns:
            df[col] = pd.NA

        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "auction_id" in df.columns:
        dedup_cols = ["date", "auction_id", "term_days"]
    else:
        dedup_cols = ["date", "roskazna_file_name"]

    df["priority"] = (
        df["roskazna_file_type"]
        .astype(str)
        .str.lower()
        .map({"xml": 1, "docx": 2})
        .fillna(9)
    )

    df = df.sort_values(dedup_cols + ["priority"])
    df = df.drop_duplicates(subset=dedup_cols, keep="first")

    print(f"Росказна: распарсено строк: {len(df)}")
    print(f"Росказна: найдено размещений: {df['roskazna_deposit_volume'].notna().sum()}")

    return df.reset_index(drop=True)


def build_roskazna_monthly(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "date",
                "roskazna_deposit_volume",
                "roskazna_planned_volume",
                "roskazna_total_bid",
                "roskazna_total_accept",
                "roskazna_banks_count",
            ]
        )

    monthly = (
        df.set_index("date")
        .resample("ME")
        .agg(
            roskazna_deposit_volume=("roskazna_deposit_volume", sum_with_na),
            roskazna_planned_volume=("roskazna_planned_volume", sum_with_na),
            roskazna_total_bid=("roskazna_total_bid", sum_with_na),
            roskazna_total_accept=("roskazna_total_accept", sum_with_na),
            roskazna_banks_count=("roskazna_banks_count", "max"),
        )
        .reset_index()
    )

    return monthly


def build_roskazna_weekly(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "date",
                "roskazna_deposit_volume_weekly",
                "roskazna_planned_volume_weekly",
                "roskazna_total_bid_weekly",
                "roskazna_total_accept_weekly",
                "roskazna_banks_count_weekly",
            ]
        )

    weekly = (
        df.set_index("date")
        .resample("W-FRI")
        .agg(
            roskazna_deposit_volume_weekly=("roskazna_deposit_volume", sum_with_na),
            roskazna_planned_volume_weekly=("roskazna_planned_volume", sum_with_na),
            roskazna_total_bid_weekly=("roskazna_total_bid", sum_with_na),
            roskazna_total_accept_weekly=("roskazna_total_accept", sum_with_na),
            roskazna_banks_count_weekly=("roskazna_banks_count", "max"),
        )
        .reset_index()
    )

    return weekly


# ============================================================
# 3. Ground truth ЦБ
# ============================================================


def load_cbr_bliquidity_ground_truth() -> pd.DataFrame:
    files = sorted(CBR_BLIQUIDITY_DIR.glob("*.xlsx"))

    if not files:
        print("Ground truth ЦБ по ликвидности не найден. Пропускаю.")
        return pd.DataFrame()

    file_path = files[-1]

    print(f"Читаю ground truth ЦБ: {file_path.name}")

    try:
        df = pd.read_excel(file_path, sheet_name="data")
    except Exception:
        try:
            df = pd.read_excel(file_path)
        except Exception as error:
            print("Не удалось прочитать ground truth ЦБ.")
            print(error)
            return pd.DataFrame()

    if "date" not in df.columns:
        print("В ground truth нет колонки date.")
        return pd.DataFrame()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()

    rename_map = {
        "liquidity_deficit_surplus": "ground_truth_liquidity_balance",
        "bank_correspondent_accounts": "ground_truth_corr_accounts",
        "required_reserves_averaging": "ground_truth_required_reserves",
    }

    df = df.rename(columns=rename_map)

    needed_cols = [
        "date",
        "ground_truth_liquidity_balance",
        "ground_truth_corr_accounts",
        "ground_truth_required_reserves",
    ]

    for col in needed_cols:
        if col not in df.columns:
            df[col] = pd.NA

    monthly = (
        df[needed_cols]
        .set_index("date")
        .resample("ME")
        .last()
        .reset_index()
    )

    return monthly


# ============================================================
# 4. Итоговые сигналы M5
# ============================================================


def build_m5_signals() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cbr_df, cbr_debug = load_cbr_sors_budget_funds()
    cbr_monthly = build_cbr_monthly(cbr_df)

    roskazna_raw = load_roskazna_deposits()
    roskazna_monthly = build_roskazna_monthly(roskazna_raw)
    roskazna_weekly = build_roskazna_weekly(roskazna_raw)

    ground_truth = load_cbr_bliquidity_ground_truth()

    frames = []

    if not cbr_monthly.empty:
        frames.append(cbr_monthly)

    if not roskazna_monthly.empty:
        frames.append(roskazna_monthly)

    if not ground_truth.empty:
        frames.append(ground_truth)

    if not frames:
        raise RuntimeError("Не удалось собрать M5: нет данных.")

    min_date = min(frame["date"].min() for frame in frames if not frame.empty)
    max_date = max(frame["date"].max() for frame in frames if not frame.empty)

    timeline = pd.DataFrame(
        {
            "date": pd.date_range(
                min_date.to_period("M").to_timestamp("M"),
                max_date.to_period("M").to_timestamp("M"),
                freq="ME",
            )
        }
    )

    result = timeline.copy()

    for frame in frames:
        result = result.merge(frame, on="date", how="left")

    base_columns = [
        "cbr_federal_budget_funds",
        "cbr_extra_budgetary_funds",
        "cbr_budget_funds_total",
        "roskazna_deposit_volume",
        "roskazna_planned_volume",
        "roskazna_total_bid",
        "roskazna_total_accept",
        "roskazna_banks_count",
        "ground_truth_liquidity_balance",
        "ground_truth_corr_accounts",
        "ground_truth_required_reserves",
    ]

    result = ensure_columns(result, base_columns)

    for col in base_columns:
        result[col] = pd.to_numeric(result[col], errors="coerce")

    # ----------------------------
    # Месячные изменения
    # ----------------------------

    result["cbr_delta_month"] = result["cbr_budget_funds_total"].diff()
    result["roskazna_deposit_volume_delta_month"] = result["roskazna_deposit_volume"].diff()

    # ЦБ: падение бюджетных остатков в банках = стресс, поэтому знак разворачиваем.
    result["mad_score_cbr_raw"] = calculate_mad_score(
        result["cbr_delta_month"],
        window=MAD_WINDOW_MONTHLY,
        min_periods=MIN_PERIODS_MONTHLY,
    )
    result["mad_score_cbr"] = -result["mad_score_cbr_raw"]

    # Росказна: падение размещений ЕКС на депозитах = меньше притока в банки = стресс.
    result["mad_score_roskazna_monthly_raw"] = calculate_mad_score(
        result["roskazna_deposit_volume_delta_month"],
        window=MAD_WINDOW_MONTHLY,
        min_periods=MIN_PERIODS_MONTHLY,
    )
    result["mad_score_roskazna_monthly"] = -result["mad_score_roskazna_monthly_raw"]

    # ----------------------------
    # Недельные изменения Росказны
    # ----------------------------

    if not roskazna_weekly.empty:
        roskazna_weekly["roskazna_deposit_volume_delta_week"] = (
            roskazna_weekly["roskazna_deposit_volume_weekly"].diff()
        )

        roskazna_weekly["mad_score_roskazna_weekly_raw"] = calculate_mad_score(
            roskazna_weekly["roskazna_deposit_volume_delta_week"],
            window=MAD_WINDOW_WEEKLY,
            min_periods=MIN_PERIODS_WEEKLY,
        )

        roskazna_weekly["mad_score_roskazna_weekly"] = -roskazna_weekly[
            "mad_score_roskazna_weekly_raw"
        ]

        roskazna_weekly["month"] = roskazna_weekly["date"].dt.to_period("M").dt.to_timestamp("M")

        roskazna_weekly_monthly_signal = (
            roskazna_weekly.groupby("month", as_index=False)
            .agg(
                mad_score_roskazna_weekly_max=("mad_score_roskazna_weekly", "max"),
                roskazna_deposit_volume_delta_week_min=("roskazna_deposit_volume_delta_week", "min"),
            )
            .rename(columns={"month": "date"})
        )

        result = result.merge(roskazna_weekly_monthly_signal, on="date", how="left")
    else:
        result["mad_score_roskazna_weekly_max"] = pd.NA
        result["roskazna_deposit_volume_delta_week_min"] = pd.NA

    # ----------------------------
    # Финальный MAD-score Росказны
    # ----------------------------
    # Для LSI оставляем название mad_score_roskazna.
    # Это уже не только месячный сигнал, а максимум между месячным и недельным стрессом.
    result["mad_score_roskazna"] = result[
        ["mad_score_roskazna_monthly", "mad_score_roskazna_weekly_max"]
    ].max(axis=1)

    # ----------------------------
    # Флаг бюджетного оттока
    # ----------------------------

    result["flag_budget_drain"] = (
        (result["cbr_delta_month"] <= BUDGET_DRAIN_THRESHOLD)
        | (result["roskazna_deposit_volume_delta_month"] <= BUDGET_DRAIN_THRESHOLD)
        | (result["roskazna_deposit_volume_delta_week_min"] <= BUDGET_DRAIN_THRESHOLD)
        | (result["mad_score_cbr"] >= 2.5)
        | (result["mad_score_roskazna"] >= 2.5)
    ).astype(int)

    final_columns = [
        "date",
        "cbr_federal_budget_funds",
        "cbr_extra_budgetary_funds",
        "cbr_budget_funds_total",
        "cbr_delta_month",
        "roskazna_deposit_volume",
        "roskazna_deposit_volume_delta_month",
        "roskazna_deposit_volume_delta_week_min",
        "roskazna_planned_volume",
        "roskazna_total_bid",
        "roskazna_total_accept",
        "roskazna_banks_count",
        "ground_truth_liquidity_balance",
        "ground_truth_corr_accounts",
        "ground_truth_required_reserves",
        "mad_score_cbr",
        "mad_score_roskazna_monthly",
        "mad_score_roskazna_weekly_max",
        "mad_score_roskazna",
        "flag_budget_drain",
    ]

    result = ensure_columns(result, final_columns)
    result = result[final_columns]

    return result, roskazna_raw, roskazna_weekly, cbr_debug


# ============================================================
# 5. График
# ============================================================


def plot_m5_treasury_flow(result: pd.DataFrame) -> None:
    plot_df = result.dropna(subset=["date"]).copy()

    if plot_df.empty:
        print("Нет данных для графика M5.")
        return

    plt.figure(figsize=(15, 7))

    if plot_df["cbr_delta_month"].notna().sum() > 0:
        plt.bar(
            plot_df["date"],
            plot_df["cbr_delta_month"],
            width=20,
            alpha=0.55,
            label="Изменение бюджетных средств в банках, ЦБ",
        )

    if plot_df["roskazna_deposit_volume_delta_month"].notna().sum() > 0:
        plt.plot(
            plot_df["date"],
            plot_df["roskazna_deposit_volume_delta_month"],
            linewidth=2.0,
            label="Изменение размещений ЕКС на депозитах, Росказна",
        )

    if plot_df["roskazna_deposit_volume_delta_week_min"].notna().sum() > 0:
        plt.plot(
            plot_df["date"],
            plot_df["roskazna_deposit_volume_delta_week_min"],
            linewidth=1.6,
            linestyle="--",
            label="Минимальное недельное изменение размещений Росказны",
        )

    plt.axhline(0, linewidth=1.0)
    plt.axhline(
        BUDGET_DRAIN_THRESHOLD,
        linestyle="--",
        linewidth=1.3,
        label=f"Порог оттока: {abs(BUDGET_DRAIN_THRESHOLD):.0f} млрд руб.",
    )

    stress_points = plot_df[plot_df["flag_budget_drain"] == 1]

    if not stress_points.empty:
        y_values = stress_points["cbr_delta_month"]

        if y_values.notna().sum() == 0:
            y_values = stress_points["roskazna_deposit_volume_delta_month"]

        if y_values.notna().sum() == 0:
            y_values = stress_points["roskazna_deposit_volume_delta_week_min"]

        plt.scatter(
            stress_points["date"],
            y_values,
            s=45,
            marker="v",
            label="Flag_Budget_Drain",
            zorder=5,
        )

    plt.title("M5 — Приток/отток казначейской ликвидности")
    plt.xlabel("Дата")
    plt.ylabel("Изменение, млрд руб.")
    plt.grid(True, alpha=0.25)
    plt.legend()
    plt.tight_layout()

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    plt.savefig(PLOT_FILE, dpi=220)
    plt.close()


# ============================================================
# 6. Сохранение
# ============================================================


def save_outputs(
    result: pd.DataFrame,
    roskazna_raw: pd.DataFrame,
    roskazna_weekly: pd.DataFrame,
    cbr_debug: pd.DataFrame,
) -> None:
    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    description = pd.DataFrame(
        [
            ["date", "Дата, конец месяца"],
            ["cbr_federal_budget_funds", "Средства федерального бюджета в банках, млрд руб."],
            ["cbr_extra_budgetary_funds", "Средства государственных внебюджетных фондов в банках, млрд руб."],
            ["cbr_budget_funds_total", "Суммарные бюджетные средства в банках, млрд руб."],
            ["cbr_delta_month", "Месячное изменение бюджетных средств в банках, млрд руб."],
            ["roskazna_deposit_volume", "Фактически размещённый объём ЕКС на банковских депозитах за месяц, млрд руб."],
            ["roskazna_deposit_volume_delta_month", "Месячное изменение фактических размещений Росказны, млрд руб."],
            ["roskazna_deposit_volume_delta_week_min", "Минимальное недельное изменение размещений Росказны внутри месяца, млрд руб."],
            ["roskazna_planned_volume", "Объявленный лимит / максимальный объём отбора Росказны за месяц, млрд руб."],
            ["roskazna_total_bid", "Спрос кредитных организаций на отборах Росказны за месяц, млрд руб."],
            ["roskazna_total_accept", "Принятый объём заявок на отборах Росказны за месяц, млрд руб."],
            ["roskazna_banks_count", "Количество банков-участников / принятых участников, если найдено"],
            ["ground_truth_liquidity_balance", "Ground truth: дефицит / профицит ликвидности банковского сектора"],
            ["ground_truth_corr_accounts", "Ground truth: средства банков на корсчетах"],
            ["ground_truth_required_reserves", "Ground truth: обязательные резервы к усреднению"],
            ["mad_score_cbr", "MAD-score по месячному изменению бюджетных остатков ЦБ. Положительное значение = стресс."],
            ["mad_score_roskazna_monthly", "Месячный MAD-score по изменению размещений Росказны. Положительное значение = стресс."],
            ["mad_score_roskazna_weekly_max", "Максимальный недельный MAD-score Росказны внутри месяца. Положительное значение = стресс."],
            ["mad_score_roskazna", "Финальный MAD-score Росказны для LSI: максимум между месячным и недельным стрессом."],
            ["flag_budget_drain", "Флаг резкого бюджетного оттока / провала размещений Росказны"],
        ],
        columns=["column_name", "description"],
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="m5_monthly_signals", index=False)
        roskazna_weekly.to_excel(writer, sheet_name="roskazna_weekly", index=False)
        roskazna_raw.to_excel(writer, sheet_name="roskazna_raw_parsed", index=False)
        cbr_debug.to_excel(writer, sheet_name="debug_cbr_sors_matches", index=False)
        description.to_excel(writer, sheet_name="columns_description", index=False)

    if not cbr_debug.empty:
        cbr_debug.to_excel(DEBUG_CBR_SORS_FILE, index=False)

    if not roskazna_raw.empty:
        roskazna_raw.to_excel(DEBUG_ROSKAZNA_FILE, index=False)


# ============================================================
# 7. Запуск
# ============================================================


def main() -> None:
    print("=" * 70)
    print("M5 — Средства федерального казначейства")
    print("Готовлю сигналы по ЦБ SORS, Росказне и ground truth")
    print("=" * 70)

    result, roskazna_raw, roskazna_weekly, cbr_debug = build_m5_signals()

    save_outputs(result, roskazna_raw, roskazna_weekly, cbr_debug)
    plot_m5_treasury_flow(result)

    print()
    print("Готово.")
    print(f"Итоговый Excel сохранён: {OUTPUT_FILE}")
    print(f"График сохранён: {PLOT_FILE}")

    print()
    print(f"Количество месячных строк: {len(result)}")
    print(f"Количество строк Росказны: {len(roskazna_raw)}")
    print(f"Количество недельных строк Росказны: {len(roskazna_weekly)}")
    print(f"Диагностических строк ЦБ SORS: {len(cbr_debug)}")

    if not result.empty:
        print(f"Период: {result['date'].min().date()} — {result['date'].max().date()}")
        print(f"Количество флагов Flag_Budget_Drain: {int(result['flag_budget_drain'].sum())}")

        print()
        print("Заполненность ключевых колонок:")

        for col in [
            "cbr_budget_funds_total",
            "cbr_delta_month",
            "roskazna_deposit_volume",
            "roskazna_deposit_volume_delta_month",
            "roskazna_deposit_volume_delta_week_min",
            "mad_score_cbr",
            "mad_score_roskazna_monthly",
            "mad_score_roskazna_weekly_max",
            "mad_score_roskazna",
            "flag_budget_drain",
        ]:
            if col in result.columns:
                print(f"{col}: {result[col].notna().sum()} из {len(result)}")

        print()
        print("Первые строки итоговой таблицы:")
        print(result.head())


if __name__ == "__main__":
    main()
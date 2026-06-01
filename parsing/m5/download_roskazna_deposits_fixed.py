from pathlib import Path
from urllib.parse import urljoin
import re
import time
import xml.etree.ElementTree as ET

import requests
import urllib3
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# M5 — Росказна: размещение средств ЕКС на банковских депозитах
#
# Скрипт:
# 1. скачивает страницы Росказны;
# 2. собирает ссылки на DOCX/XML;
# 3. скачивает файлы;
# 4. сохраняет files_index.xlsx / files_index.csv;
# 5. сразу парсит XML и сохраняет roskazna_deposits_parsed.xlsx.
#
# Важное:
# - maxvol      — предельный объём отбора, млн руб.;
# - totalbid    — спрос банков, млн руб.;
# - totalaccept — принятый объём, млн руб.;
# - totalsettle — фактически размещённый объём, млн руб.
# Для M5 основной показатель — totalsettle, если он есть.
# ============================================================


BASE_URL = "https://roskazna.gov.ru"

PAGE_URL = (
    "https://roskazna.gov.ru/finansovye-operacii/"
    "razmeshchenie-sredstv-edinogo-kaznachejskogo-scheta/"
    "razmeshchenie-sredstv-edinogo-kaznachejskogo-scheta-na-bankovskih-depozitah"
)

OUTPUT_DIR = Path("data/m5/roskazna_deposits")
RAW_HTML_DIR = OUTPUT_DIR / "raw_html"
RAW_FILES_DIR = OUTPUT_DIR / "raw_files"

INDEX_XLSX_FILE = OUTPUT_DIR / "files_index.xlsx"
INDEX_CSV_FILE = OUTPUT_DIR / "files_index.csv"
PARSED_XLSX_FILE = OUTPUT_DIR / "roskazna_deposits_parsed.xlsx"
PARSED_CSV_FILE = OUTPUT_DIR / "roskazna_deposits_parsed.csv"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
RAW_HTML_DIR.mkdir(parents=True, exist_ok=True)
RAW_FILES_DIR.mkdir(parents=True, exist_ok=True)

START_YEAR = 2021
END_YEAR = 2026

MAX_PAGES_PER_YEAR = 80
REQUEST_PAUSE_SECONDS = 0.4

SSL_VERIFY = False

if not SSL_VERIFY:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


MONTHS_RU = {
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


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0 Safari/537.36"
            ),
            "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        }
    )
    return session


def safe_filename(value: str) -> str:
    value = str(value).strip()
    value = re.sub(r"[^\w.\-а-яА-ЯёЁ]+", "_", value, flags=re.UNICODE)
    value = re.sub(r"_+", "_", value)
    return value.strip("_")


def parse_russian_date(value: str):
    if value is None:
        return pd.NaT

    text = str(value).strip().lower()
    text = text.replace(",", " ")
    text = re.sub(r"\s+", " ", text)

    dot_match = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if dot_match:
        return pd.Timestamp(
            year=int(dot_match.group(3)),
            month=int(dot_match.group(2)),
            day=int(dot_match.group(1)),
        )

    word_match = re.search(
        r"(\d{1,2})\s+"
        r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
        r"\s+(\d{4})",
        text,
    )
    if word_match:
        return pd.Timestamp(
            year=int(word_match.group(3)),
            month=MONTHS_RU[word_match.group(2)],
            day=int(word_match.group(1)),
        )

    return pd.NaT


def extract_date_from_filename(filename: str):
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(filename))
    if not match:
        return pd.NaT

    return pd.Timestamp(
        year=int(match.group(3)),
        month=int(match.group(2)),
        day=int(match.group(1)),
    )


def extract_selection_number(filename: str):
    text = str(filename)
    match = re.search(r"otbor_([0-9_]+)", text, flags=re.IGNORECASE)
    if match:
        return match.group(1)

    if re.search(r"otbor", text, flags=re.IGNORECASE):
        return "0"

    return None


def build_page_params(year: int | None = None, page: int | None = None) -> dict:
    params = {
        "fundPlace": "2",
        "filter_type": "2",
    }

    if year is not None:
        params["filter_year"] = str(year)

    if page is not None and page > 1:
        params["page"] = str(page)

    return params


def download_html(session: requests.Session, params: dict | None = None) -> tuple[str, str]:
    response = session.get(PAGE_URL, params=params, timeout=60, verify=SSL_VERIFY)
    response.raise_for_status()
    response.encoding = "utf-8"
    return response.text, response.url


def find_nearest_date_text(link_tag) -> str | None:
    tr = link_tag.find_parent("tr")
    if tr is not None:
        date_cell = tr.find("td", class_=lambda x: x and "table-date" in x)
        if date_cell is not None:
            return date_cell.get_text(" ", strip=True)

    block = link_tag.find_parent(class_=lambda x: x and "info-archive-table-block" in x)
    if block is not None:
        for p in block.find_all("p"):
            text = p.get_text(" ", strip=True)
            if re.search(r"\d{1,2}\s+[а-яё]+\s+\d{4}", text.lower()):
                return text

    return None


def parse_files_from_html(html: str, source_url: str, source_label: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()

        if "/storage/operation-day-files/" not in href:
            continue

        if not re.search(r"\.(docx|xml)$", href, flags=re.IGNORECASE):
            continue

        file_url = urljoin(BASE_URL, href)
        file_name = a.get("title") or Path(href).name
        file_name = str(file_name).strip()
        file_type = Path(file_name).suffix.lower().replace(".", "")

        date_text = find_nearest_date_text(a)
        date_value = parse_russian_date(date_text)
        if pd.isna(date_value):
            date_value = extract_date_from_filename(file_name)

        rows.append(
            {
                "date": date_value,
                "date_text": date_text,
                "file_name": file_name,
                "file_type": file_type,
                "selection_number": extract_selection_number(file_name),
                "file_url": file_url,
                "source_page_url": source_url,
                "source_label": source_label,
            }
        )

    return rows


def save_html(html: str, filename: str) -> None:
    (RAW_HTML_DIR / filename).write_text(html, encoding="utf-8")


def download_file(session: requests.Session, file_url: str, output_path: Path) -> None:
    response = session.get(file_url, timeout=60, headers={"Accept": "*/*"}, verify=SSL_VERIFY)
    response.raise_for_status()
    output_path.write_bytes(response.content)


def build_local_file_name(row: dict) -> str:
    date_value = row.get("date")
    if pd.notna(date_value):
        date_part = pd.Timestamp(date_value).strftime("%Y-%m-%d")
    else:
        date_part = "unknown_date"

    # На сайте иногда title неполный, поэтому добавляем имя из URL.
    url_name = Path(str(row.get("file_url", ""))).name
    visible_name = row.get("file_name") or url_name
    return f"{date_part}_{safe_filename(visible_name)}"


def normalize_path_for_current_os(value: str | Path) -> Path:
    """Исправляет Windows-пути из Excel: data\m5\... -> data/m5/..."""
    text = str(value).strip().strip('"').strip("'")
    text = text.replace("\\", "/")
    return Path(text)


def clean_number(value):
    if value is None:
        return None

    text = str(value).strip()
    if text in ("", "-", "—", "nan", "None", "NaN"):
        return None

    text = text.replace("\xa0", "")
    text = text.replace(" ", "")
    text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)

    if text in ("", "-", ".", "-."):
        return None

    try:
        return float(text)
    except ValueError:
        return None


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


def parse_roskazna_xml_file(file_path: Path, fallback_date=None) -> list[dict]:
    """Парсит официальный XML Росказны по тегам Depoauc*.

    Возвращает одну или несколько строк, если в файле несколько отборов.
    Денежные поля в XML Росказны указаны в млн руб., поэтому переводим в млрд руб.
    """
    if not file_path.exists():
        return []

    rows = []

    raw_bytes = file_path.read_bytes()
    root = None

    for encoding in ["utf-8", "windows-1251", "cp1251"]:
        try:
            text = raw_bytes.decode(encoding, errors="ignore")
            root = ET.fromstring(text)
            break
        except Exception:
            root = None

    if root is None:
        return []

    auction_nodes = []
    for node in root.iter():
        clean_tag = node.tag.split("}")[-1].lower()
        if clean_tag.startswith("depoauc"):
            auction_nodes.append(node)

    if not auction_nodes:
        auction_nodes = [root]

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
            # Fallback нужен только как proxy, если сайт поменяет XML-структуру.
            placed_bln = maxvol_bln
            volume_source = "maxvol_proxy"

        rows.append(
            {
                "date": aucdate,
                "auction_id": get_xml_child_text(node, "id"),
                "funds_placed": get_xml_child_text(node, "FundsPlaced"),
                "currency": get_xml_child_text(node, "cur"),
                "term_days": clean_number(get_xml_child_text(node, "term")),
                "firstdate": parse_russian_date(get_xml_child_text(node, "firstdate")),
                "seconddate": parse_russian_date(get_xml_child_text(node, "seconddate")),
                "rate_type": get_xml_child_text(node, "ratetype"),
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
            }
        )

    return rows


def parse_downloaded_xml_files(index_df: pd.DataFrame) -> pd.DataFrame:
    if index_df.empty:
        return pd.DataFrame()

    xml_df = index_df[index_df["file_type"].astype(str).str.lower() == "xml"].copy()
    rows = []

    for _, row in xml_df.iterrows():
        local_path = normalize_path_for_current_os(row["local_file_path"])
        parsed_rows = parse_roskazna_xml_file(local_path, fallback_date=row.get("date"))

        for parsed in parsed_rows:
            parsed["file_name"] = row.get("file_name")
            parsed["selection_number"] = row.get("selection_number")
            parsed["file_url"] = row.get("file_url")
            rows.append(parsed)

    parsed_df = pd.DataFrame(rows)

    if parsed_df.empty:
        return parsed_df

    parsed_df["date"] = pd.to_datetime(parsed_df["date"], errors="coerce")
    parsed_df = parsed_df.dropna(subset=["date"]).copy()

    # На всякий случай убираем дубли одного и того же аукциона.
    parsed_df = parsed_df.sort_values(["date", "auction_id", "source_file_path"])
    parsed_df = parsed_df.drop_duplicates(subset=["date", "auction_id", "term_days", "placed_volume_bln"], keep="last")

    return parsed_df.reset_index(drop=True)


def main() -> None:
    print("=" * 70)
    print("M5 — Росказна: скачивание и парсинг депозитов ЕКС")
    print("=" * 70)

    session = make_session()
    all_rows = []

    print("Скачиваю основную страницу Росказны...")
    html, final_url = download_html(session)
    save_html(html, "roskazna_deposits_main.html")

    main_rows = parse_files_from_html(html=html, source_url=final_url, source_label="main_page")
    all_rows.extend(main_rows)
    print(f"На основной странице найдено файлов: {len(main_rows)}")

    for year in range(START_YEAR, END_YEAR + 1):
        print()
        print(f"Обрабатываю архив за {year} год...")

        year_rows_count = 0
        previous_page_urls = set()

        for page in range(1, MAX_PAGES_PER_YEAR + 1):
            params = build_page_params(year=year, page=page)

            try:
                html, final_url = download_html(session, params=params)
            except requests.RequestException as error:
                print(f"{year}, page={page}: ошибка загрузки")
                print(error)
                break

            if final_url in previous_page_urls:
                print(f"{year}, page={page}: страница повторилась, стоп")
                break

            previous_page_urls.add(final_url)
            save_html(html, f"roskazna_deposits_{year}_page_{page}.html")

            rows = parse_files_from_html(
                html=html,
                source_url=final_url,
                source_label=f"archive_{year}_page_{page}",
            )

            filtered_rows = []
            for row in rows:
                if pd.notna(row["date"]) and pd.Timestamp(row["date"]).year != year:
                    continue
                filtered_rows.append(row)

            if not filtered_rows:
                print(f"{year}, page={page}: файлов не найдено, стоп")
                break

            all_rows.extend(filtered_rows)
            year_rows_count += len(filtered_rows)

            print(f"{year}, page={page}: найдено файлов {len(filtered_rows)}")
            time.sleep(REQUEST_PAUSE_SECONDS)

        print(f"Итого за {year}: {year_rows_count}")

    if not all_rows:
        raise RuntimeError("Не удалось найти DOCX/XML файлы Росказны.")

    index_df = pd.DataFrame(all_rows)
    index_df = index_df.drop_duplicates(subset=["file_url"]).copy()
    index_df["date"] = pd.to_datetime(index_df["date"], errors="coerce")
    index_df = index_df.sort_values(["date", "file_name"], na_position="last").reset_index(drop=True)

    print()
    print("Скачиваю DOCX/XML файлы...")

    local_paths = []
    downloaded_count = 0

    for _, row in index_df.iterrows():
        local_file_name = build_local_file_name(row.to_dict())
        local_path = RAW_FILES_DIR / local_file_name
        local_paths.append(local_path.as_posix())

        if local_path.exists():
            print(f"{local_file_name}: уже есть")
            continue

        try:
            print(f"{local_file_name}: скачиваю...")
            download_file(session, row["file_url"], local_path)
            downloaded_count += 1
            time.sleep(REQUEST_PAUSE_SECONDS)
        except requests.RequestException as error:
            print(f"{local_file_name}: ошибка скачивания")
            print(error)

    index_df["local_file_path"] = local_paths

    index_df.to_excel(INDEX_XLSX_FILE, index=False)
    index_df.to_csv(INDEX_CSV_FILE, index=False, encoding="utf-8-sig")

    parsed_df = parse_downloaded_xml_files(index_df)
    parsed_df.to_excel(PARSED_XLSX_FILE, index=False)
    parsed_df.to_csv(PARSED_CSV_FILE, index=False, encoding="utf-8-sig")

    print()
    print("Готово.")
    print(f"Реестр Excel сохранён: {INDEX_XLSX_FILE}")
    print(f"Реестр CSV сохранён: {INDEX_CSV_FILE}")
    print(f"Распарсенные XML сохранены: {PARSED_XLSX_FILE}")
    print(f"HTML сохранены: {RAW_HTML_DIR}")
    print(f"DOCX/XML сохранены: {RAW_FILES_DIR}")
    print(f"Всего уникальных файлов: {len(index_df)}")
    print(f"Скачано новых файлов: {downloaded_count}")
    print(f"Распарсено XML-строк: {len(parsed_df)}")

    if not index_df.empty:
        min_date = index_df["date"].min()
        max_date = index_df["date"].max()
        if pd.notna(min_date) and pd.notna(max_date):
            print(f"Период файлов: {min_date.date()} — {max_date.date()}")

    if not parsed_df.empty:
        print()
        print("Заполненность по Росказне:")
        for col in ["maxvol_bln", "totalbid_bln", "totalaccept_bln", "totalsettle_bln", "placed_volume_bln"]:
            print(f"{col}: {parsed_df[col].notna().sum()} из {len(parsed_df)}")
        print(f"Суммарный фактически размещённый объём: {parsed_df['placed_volume_bln'].sum():.1f} млрд руб.")


if __name__ == "__main__":
    main()

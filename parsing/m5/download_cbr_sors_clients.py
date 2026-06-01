from pathlib import Path
from urllib.parse import urljoin
import re

import requests
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# M5 — Средства федерального казначейства
#
# Источник: ЦБ РФ, SORS
# https://www.cbr.ru/statistics/bank_sector/sors/
#
# Этот скрипт только скачивает Excel-файлы из блока
# "Привлеченные средства".
#
# Важно:
# раньше скачивался только файл 02_19_Funds_clients_branches_*.xlsx,
# но в нём может не быть строк про федеральный бюджет.
# Поэтому сейчас скачиваем все Excel из /Borrowings/ за нужный период.
#
# На выходе:
# data/m5/cbr_sors/
# data/m5/cbr_sors/cbr_sors_files_index.xlsx
# ============================================================


BASE_URL = "https://www.cbr.ru"
PAGE_URL = "https://www.cbr.ru/statistics/bank_sector/sors/"

OUTPUT_DIR = Path("data/m5/cbr_sors")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

RAW_HTML_FILE = OUTPUT_DIR / "cbr_sors_page.html"
INDEX_FILE = OUTPUT_DIR / "cbr_sors_files_index.xlsx"

START_YEAR = 2019
END_YEAR = 2026


def download_html() -> str:
    print("Скачиваю страницу ЦБ SORS...")
    print(PAGE_URL)

    response = requests.get(
        PAGE_URL,
        timeout=60,
        headers={"User-Agent": "Mozilla/5.0"},
    )

    response.raise_for_status()
    response.encoding = "utf-8"

    return response.text


def extract_date_from_name(name: str):
    match = re.search(r"(20\d{2})(\d{2})(\d{2})", name)

    if not match:
        return pd.NaT

    year = int(match.group(1))
    month = int(match.group(2))
    day = int(match.group(3))

    return pd.Timestamp(year=year, month=month, day=day)


def extract_excel_links(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "html.parser")

    rows = []

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()

        if not href.lower().endswith(".xlsx"):
            continue

        # Нас интересуют файлы привлеченных средств.
        # На сайте ЦБ они обычно лежат в BankSector/Borrowings.
        if "/Borrowings/" not in href and "/borrowings/" not in href.lower():
            continue

        file_name = Path(href).name
        file_date = extract_date_from_name(file_name)

        if pd.isna(file_date):
            continue

        if file_date.year < START_YEAR or file_date.year > END_YEAR:
            continue

        full_url = urljoin(BASE_URL, href)

        title = a.get_text(" ", strip=True)
        if not title:
            title = a.get("title", "")

        rows.append(
            {
                "date": file_date,
                "file_name": file_name,
                "url": full_url,
                "title": title,
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df = df.drop_duplicates(subset=["url"]).sort_values(["date", "file_name"])
    df = df.reset_index(drop=True)

    return df


def download_file(url: str, file_path: Path) -> None:
    response = requests.get(
        url,
        timeout=90,
        headers={"User-Agent": "Mozilla/5.0"},
    )

    response.raise_for_status()
    file_path.write_bytes(response.content)


def main() -> None:
    print("=" * 70)
    print("M5 — ЦБ SORS: скачивание Excel по привлечённым средствам")
    print("=" * 70)

    html = download_html()
    RAW_HTML_FILE.write_text(html, encoding="utf-8")

    index_df = extract_excel_links(html)

    if index_df.empty:
        raise RuntimeError("Не нашёл Excel-файлы /Borrowings/ на странице ЦБ SORS.")

    print()
    print(f"Найдено Excel-файлов: {len(index_df)}")

    downloaded = 0

    for _, row in index_df.iterrows():
        file_path = OUTPUT_DIR / row["file_name"]

        if file_path.exists():
            print(f"{row['file_name']}: уже есть, пропускаю")
            continue

        try:
            print(f"{row['file_name']}: скачиваю...")
            download_file(row["url"], file_path)
            downloaded += 1
        except requests.RequestException as error:
            print(f"{row['file_name']}: ошибка скачивания")
            print(error)

    index_df.to_excel(INDEX_FILE, index=False)

    print()
    print("Готово.")
    print(f"HTML страницы сохранён: {RAW_HTML_FILE}")
    print(f"Реестр файлов сохранён: {INDEX_FILE}")
    print(f"Папка с Excel: {OUTPUT_DIR}")
    print(f"Скачано новых файлов: {downloaded}")


if __name__ == "__main__":
    main()
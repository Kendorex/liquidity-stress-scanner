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
# Для M5 нужен НЕ весь раздел "Привлечённые средства",
# а конкретная таблица:
#
# "Бюджетные средства на счетах кредитных организаций"
#
# Файл на сайте ЦБ:
# 02_29_Budget_all.xlsx
# ============================================================


BASE_URL = "https://www.cbr.ru"
PAGE_URL = "https://www.cbr.ru/statistics/bank_sector/sors/"

OUTPUT_DIR = Path("data/m5/cbr_sors")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

RAW_HTML_FILE = OUTPUT_DIR / "cbr_sors_page.html"

INDEX_XLSX_FILE = OUTPUT_DIR / "cbr_sors_files_index.xlsx"
INDEX_CSV_FILE = OUTPUT_DIR / "cbr_sors_files_index.csv"

TARGET_FILE_NAME = "02_29_Budget_all.xlsx"

DIRECT_TARGET_URL = (
    "https://www.cbr.ru/vfs/statistics/banksector/borrowings/02_29_Budget_all.xlsx"
)


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


def clean_text(value) -> str:
    if value is None:
        return ""

    value = str(value).lower()
    value = value.replace("\xa0", " ")
    value = value.replace("ё", "е")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def download_html(session: requests.Session) -> str:
    print("Скачиваю страницу ЦБ SORS...")
    print(PAGE_URL)

    response = session.get(PAGE_URL, timeout=60)
    response.raise_for_status()
    response.encoding = "utf-8"

    return response.text


def is_target_budget_link(href: str, title: str, text: str) -> bool:
    href_low = href.lower()
    file_name = Path(href_low).name

    if not href_low.endswith(".xlsx"):
        return False
    if file_name == TARGET_FILE_NAME.lower():
        return True

    combined = clean_text(f"{title} {text}")

    has_budget_title = (
        "бюджетные средства на счетах кредитных организаций" in combined
        or ("бюджетные средства" in combined and "кредитных организаций" in combined)
    )

    in_borrowings_section = "/borrowings/" in href_low

    return has_budget_title and in_borrowings_section


def extract_target_link_from_html(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "html.parser")

    rows = []

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()

        title = (
            a.get("data-zoom-title")
            or a.get("title")
            or a.get("aria-label")
            or ""
        )

        text = a.get_text(" ", strip=True)

        if not is_target_budget_link(href=href, title=title, text=text):
            continue

        full_url = urljoin(BASE_URL, href)
        file_name = Path(href).name

        rows.append(
            {
                "date_downloaded": pd.Timestamp.today().normalize(),
                "file_name": file_name,
                "url": full_url,
                "title": title or text,
                "source_page_url": PAGE_URL,
                "source_section": "Привлечённые средства / Borrowings",
                "metric": "budget_funds_on_credit_institutions_accounts",
                "metric_ru": "Бюджетные средства на счетах кредитных организаций",
                "module": "M5",
                "is_target_file": 1,
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df = df.drop_duplicates(subset=["url"]).copy()
    df["priority"] = df["file_name"].str.lower().eq(TARGET_FILE_NAME.lower()).map(
        {True: 1, False: 2}
    )

    df = df.sort_values(["priority", "file_name"]).drop(columns=["priority"])
    df = df.reset_index(drop=True)

    return df


def build_fallback_index() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "date_downloaded": pd.Timestamp.today().normalize(),
                "file_name": TARGET_FILE_NAME,
                "url": DIRECT_TARGET_URL,
                "title": "Бюджетные средства на счетах кредитных организаций",
                "source_page_url": PAGE_URL,
                "source_section": "Привлечённые средства / Borrowings",
                "metric": "budget_funds_on_credit_institutions_accounts",
                "metric_ru": "Бюджетные средства на счетах кредитных организаций",
                "module": "M5",
                "is_target_file": 1,
            }
        ]
    )


def validate_xlsx_bytes(content: bytes, url: str) -> None:
    if not content:
        raise RuntimeError(f"Пустой ответ при скачивании: {url}")

    if not content.startswith(b"PK"):
        preview = content[:300].decode("utf-8", errors="ignore")

        raise RuntimeError(
            "Скачанный файл не похож на XLSX. "
            "Возможно, ЦБ вернул HTML-страницу вместо Excel.\n"
            f"URL: {url}\n"
            f"Начало ответа:\n{preview}"
        )


def download_file(session: requests.Session, url: str, file_path: Path) -> None:
    response = session.get(
        url,
        timeout=90,
        headers={"Accept": "*/*"},
    )

    response.raise_for_status()

    content = response.content
    validate_xlsx_bytes(content, url)

    file_path.write_bytes(content)


def main() -> None:
    print("M5 — ЦБ SORS: скачивание бюджетных средств")

    session = make_session()

    try:
        html = download_html(session)
        RAW_HTML_FILE.write_text(html, encoding="utf-8")

        index_df = extract_target_link_from_html(html)

        if index_df.empty:
            print()
            print("В HTML не нашёл ссылку на 02_29_Budget_all.xlsx.")
            print("Использую прямую ссылку на файл ЦБ.")
            index_df = build_fallback_index()

    except requests.RequestException as error:
        print()
        print("Не удалось скачать страницу ЦБ SORS.")
        print(error)
        print("Использую прямую ссылку на файл ЦБ.")
        index_df = build_fallback_index()

    print()
    print(f"Файлов для скачивания: {len(index_df)}")

    local_paths = []
    downloaded = 0

    for _, row in index_df.iterrows():
        file_name = row["file_name"]
        if str(file_name).lower() != TARGET_FILE_NAME.lower():
            file_name = TARGET_FILE_NAME

        file_path = OUTPUT_DIR / file_name
        local_paths.append(str(file_path.as_posix()))

        try:
            if file_path.exists():
                print(f"{file_name}: уже есть, обновляю...")

            else:
                print(f"{file_name}: скачиваю...")

            download_file(session, row["url"], file_path)
            downloaded += 1

        except requests.RequestException as error:
            print(f"{file_name}: ошибка скачивания")
            print(error)

        except RuntimeError as error:
            print(f"{file_name}: ошибка проверки файла")
            print(error)

    index_df["local_file_path"] = local_paths

    index_df.to_excel(INDEX_XLSX_FILE, index=False)
    index_df.to_csv(INDEX_CSV_FILE, index=False, encoding="utf-8-sig")

    print()
    print("Готово.")
    print(f"HTML страницы сохранён: {RAW_HTML_FILE}")
    print(f"Реестр XLSX сохранён: {INDEX_XLSX_FILE}")
    print(f"Реестр CSV сохранён: {INDEX_CSV_FILE}")
    print(f"Папка с Excel: {OUTPUT_DIR}")
    print(f"Скачано / обновлено файлов: {downloaded}")

    print()
    print("Используемый показатель:")
    print(index_df[["file_name", "metric_ru", "url", "local_file_path"]].to_string(index=False))


if __name__ == "__main__":
    main()
from pathlib import Path
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import re

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


PAGE_URL = "https://cbr.ru/hd_base/repo/"
DEFAULT_FROM_DATE = "01.01.2010"
MAX_WORKERS = 6

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SAVE_DIR = PROJECT_ROOT / "data" / "m2" / "repo"

_THREAD_LOCAL = threading.local()


def create_session() -> requests.Session:
    session = requests.Session()
    session.trust_env = False

    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )

    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    return session


def get_thread_session() -> requests.Session:
    if not hasattr(_THREAD_LOCAL, "session"):
        _THREAD_LOCAL.session = create_session()

    return _THREAD_LOCAL.session


def clean_text(text: str) -> str:
    if text is None:
        return ""

    text = str(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def parse_number(value):
    text = clean_text(value)

    if text in ["", "—", "-"]:
        return None

    cleaned = text.replace(" ", "").replace(",", ".")

    try:
        return float(cleaned)
    except ValueError:
        return text


def parse_date(date_str: str) -> datetime:
    return datetime.strptime(date_str, "%d.%m.%Y")


def format_date(date_obj: datetime) -> str:
    return date_obj.strftime("%d.%m.%Y")


def get_headers() -> dict:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        "Referer": PAGE_URL,
        "Connection": "close",
    }


def extract_repo_tables_from_day(html: str, requested_date: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    for table in soup.select("table.data"):
        table_text = clean_text(table.get_text(" ", strip=True))

        if "Тип аукциона" not in table_text:
            continue

        if "Объем спроса на операции репо" not in table_text:
            continue

        row_data = {
            "auction_date": requested_date,
            "auction_datetime": requested_date,
        }

        parent = table.find_parent("div", class_="table-wrapper")
        if parent is not None:
            caption = parent.select_one(".table-caption.gray")
            if caption is not None:
                row_data["auction_datetime"] = clean_text(caption.get_text(" ", strip=True))

        for tr in table.select("tr"):
            cells = tr.find_all(["th", "td"])

            if len(cells) != 2:
                continue

            key = clean_text(cells[0].get_text(" ", strip=True))
            value = clean_text(cells[1].get_text(" ", strip=True))

            if key:
                row_data[key] = parse_number(value)

        rows.append(row_data)

    return rows


def normalize_repo_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "Тип аукциона": "auction_type",
        "Срок, дни": "term_days",
        "Дата исполнения первой части сделки": "first_leg_date",
        "Дата исполнения второй части сделки": "second_leg_date",
        "Объем спроса на операции репо, млн руб.": "repo_demand_mln_rub",
        "Общий объем заключенных сделок репо, млн руб.": "repo_deals_total_mln_rub",
        "Ставка отсечения, % годовых": "cutoff_rate",
        "Средневзвешенная ставка, % годовых": "weighted_average_rate",
        "Минимальная заявленная ставка, % годовых": "min_bid_rate",
        "Максимальная заявленная ставка, % годовых": "max_bid_rate",
        "Объем заключенных сделок репо в рамках лимита, млн руб.": "repo_deals_within_limit_mln_rub",
        "Средневзвешенная ставка по заявкам, удовлетворенным в рамках лимита, % годовых": "weighted_average_rate_within_limit",
    }

    df = df.rename(columns=rename_map)

    if "auction_date" in df.columns:
        df["auction_date"] = pd.to_datetime(df["auction_date"], format="%d.%m.%Y", errors="coerce")

    for column in ["first_leg_date", "second_leg_date"]:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], format="%d.%m.%Y", errors="coerce")

    numeric_columns = [
        "term_days",
        "repo_demand_mln_rub",
        "repo_deals_total_mln_rub",
        "repo_deals_within_limit_mln_rub",
        "cutoff_rate",
        "weighted_average_rate",
        "min_bid_rate",
        "max_bid_rate",
        "weighted_average_rate_within_limit",
    ]

    for column in numeric_columns:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")

    first_columns = [
        "auction_date",
        "auction_datetime",
        "auction_type",
        "term_days",
        "repo_demand_mln_rub",
        "repo_deals_total_mln_rub",
        "repo_deals_within_limit_mln_rub",
        "cutoff_rate",
        "weighted_average_rate",
        "min_bid_rate",
        "max_bid_rate",
        "weighted_average_rate_within_limit",
        "first_leg_date",
        "second_leg_date",
    ]

    existing_first_columns = [column for column in first_columns if column in df.columns]
    other_columns = [column for column in df.columns if column not in existing_first_columns]

    return df[existing_first_columns + other_columns]


def download_one_day_repo(date_str: str, term_filter: int = 0) -> list[dict]:
    session = get_thread_session()

    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": date_str,
        "UniDbQuery.To": date_str,
        "UniDbQuery.P1": str(term_filter),
    }

    response = session.get(PAGE_URL, params=params, headers=get_headers(), timeout=60)
    response.raise_for_status()

    return extract_repo_tables_from_day(response.text, requested_date=date_str)


def build_working_dates(start_date: datetime, end_date: datetime) -> list[str]:
    dates = []
    current_date = start_date

    while current_date <= end_date:
        if current_date.weekday() < 5:
            dates.append(format_date(current_date))

        current_date += timedelta(days=1)

    return dates


def load_date(date_str: str, term_filter: int) -> tuple[str, list[dict], str | None]:
    try:
        rows = download_one_day_repo(date_str=date_str, term_filter=term_filter)
        return date_str, rows, None
    except Exception as error:
        return date_str, [], str(error)


def download_repo_excel(
    from_date: str = DEFAULT_FROM_DATE,
    to_date: str | None = None,
    term_filter: int = 0,
    max_workers: int = MAX_WORKERS,
) -> Path:
    if to_date is None:
        to_date = datetime.today().strftime("%d.%m.%Y")

    SAVE_DIR.mkdir(parents=True, exist_ok=True)

    file_name = f"repo_{from_date.replace('.', '-')}_{to_date.replace('.', '-')}.xlsx"
    save_path = SAVE_DIR / file_name

    print("Скачиваю итоги аукционов РЕПО с сайта ЦБ")
    print(f"Период: {from_date} — {to_date}")
    print(f"Режим: рабочие дни, параллельная загрузка, потоков: {max_workers}")

    start_date = parse_date(from_date)
    end_date = parse_date(to_date)

    if start_date > end_date:
        raise ValueError("from_date не может быть позже to_date")

    dates = build_working_dates(start_date=start_date, end_date=end_date)

    print(f"Дат к проверке: {len(dates)}")
    print("Субботы и воскресенья пропущены, чтобы не делать лишние запросы.")
    print("-" * 80)

    all_rows = []
    errors = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(load_date, date_str, term_filter)
            for date_str in dates
        ]

        for checked_count, future in enumerate(as_completed(futures), start=1):
            date_str, rows, error = future.result()

            if error is not None:
                errors.append((date_str, error))
            elif rows:
                all_rows.extend(rows)
                print(f"{date_str}: найдено строк — {len(rows)}")

            if checked_count % 100 == 0:
                print(f"Проверено дат: {checked_count} из {len(dates)}")

    if not all_rows:
        raise RuntimeError("За выбранный период не удалось найти данные РЕПО")

    df = pd.DataFrame(all_rows)
    df = normalize_repo_columns(df)

    if "auction_date" in df.columns:
        df = df.sort_values("auction_date").reset_index(drop=True)

    df.to_excel(save_path, index=False)

    print("=" * 80)
    print(f"Проверено рабочих дат: {len(dates)}")
    print(f"Строк скачано: {len(df)}")
    print(f"Ошибок при скачивании: {len(errors)}")
    print(f"Файл сохранён: {save_path}")

    if errors:
        print("Первые ошибки для проверки:")
        for date_str, error in errors[:10]:
            print(f"{date_str}: {error}")
    return save_path


if __name__ == "__main__":
    download_repo_excel()
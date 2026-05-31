from pathlib import Path
from datetime import datetime, timedelta
import re
import time

import requests
import pandas as pd
from bs4 import BeautifulSoup

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


PAGE_URL = "https://cbr.ru/hd_base/repo/"

DEFAULT_FROM_DATE = "22.12.2020"


def create_session() -> requests.Session:
    """
    Создаём сессию requests с повторными попытками.
    Отключаем системные proxy-настройки, чтобы снизить риск ошибки WinError 10054.
    """

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


def clean_text(text: str) -> str:
    """
    Убирает лишние пробелы, переносы строк и неразрывные пробелы.
    """

    if text is None:
        return ""

    text = str(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def parse_number(value):
    """
    Переводит русские числа вида '5 630 745,5' в float.
    Даты и текст оставляет как текст.
    """

    if value is None:
        return None

    text = clean_text(value)

    if text in ["", "—", "-"]:
        return None

    cleaned = text.replace(" ", "").replace(",", ".")

    try:
        return float(cleaned)
    except ValueError:
        return text


def parse_date(date_str: str) -> datetime:
    """
    Переводит дату из ДД.ММ.ГГГГ в datetime.
    """

    return datetime.strptime(date_str, "%d.%m.%Y")


def format_date(date_obj: datetime) -> str:
    """
    Переводит datetime в ДД.ММ.ГГГГ.
    """

    return date_obj.strftime("%d.%m.%Y")


def get_headers() -> dict:
    """
    Заголовки, чтобы запрос был похож на обычный браузер.
    """

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
    """
    Достаёт данные РЕПО за один день.
    На странице ЦБ таблица устроена как пары:
    показатель -> значение.
    """

    soup = BeautifulSoup(html, "html.parser")

    rows = []

    # Сначала ищем таблицы, где есть "Тип аукциона"
    tables = soup.select("table.data")

    for table in tables:
        table_text = clean_text(table.get_text(" ", strip=True))

        if "Тип аукциона" not in table_text:
            continue

        if "Объем спроса на операции репо" not in table_text:
            continue

        row_data = {
            "auction_date": requested_date,
            "auction_datetime": None,
        }

        # Пытаемся найти подпись вида "26.05.2026 на 13:30"
        parent = table.find_parent("div", class_="table-wrapper")
        if parent is not None:
            caption = parent.select_one(".table-caption.gray")
            if caption is not None:
                row_data["auction_datetime"] = clean_text(caption.get_text(" ", strip=True))

        # Если подпись не нашлась, оставим просто дату
        if not row_data["auction_datetime"]:
            row_data["auction_datetime"] = requested_date

        for tr in table.select("tr"):
            cells = tr.find_all(["th", "td"])

            if len(cells) != 2:
                continue

            key = clean_text(cells[0].get_text(" ", strip=True))
            value = clean_text(cells[1].get_text(" ", strip=True))

            if not key:
                continue

            row_data[key] = parse_number(value)

        rows.append(row_data)

    return rows


def normalize_repo_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Приводит названия колонок к короткому техническому виду.
    """

    rename_map = {
        "Тип аукциона": "auction_type",
        "Объем спроса на операции репо, млн руб.": "repo_demand_mln_rub",
        "Общий объем заключенных сделок репо, млн руб.": "repo_deals_total_mln_rub",
        "Ставка отсечения, % годовых": "cutoff_rate",
        "Средневзвешенная ставка, % годовых": "weighted_average_rate",
        "Минимальная заявленная ставка, % годовых": "min_bid_rate",
        "Максимальная заявленная ставка, % годовых": "max_bid_rate",
        "Объем заключенных сделок репо в рамках лимита, млн руб.": "repo_deals_within_limit_mln_rub",
        "Средневзвешенная ставка по заявкам, удовлетворенным в рамках лимита, % годовых": "weighted_average_rate_within_limit",
        "Срок, дни": "term_days",
        "Дата исполнения первой части сделки": "first_leg_date",
        "Дата исполнения второй части сделки": "second_leg_date",
    }

    df = df.rename(columns=rename_map)

    if "auction_date" in df.columns:
        df["auction_date"] = pd.to_datetime(
            df["auction_date"],
            format="%d.%m.%Y",
            errors="coerce",
        )

    for col in ["first_leg_date", "second_leg_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col],
                format="%d.%m.%Y",
                errors="coerce",
            )

    numeric_cols = [
        "repo_demand_mln_rub",
        "repo_deals_total_mln_rub",
        "repo_deals_within_limit_mln_rub",
        "cutoff_rate",
        "weighted_average_rate",
        "min_bid_rate",
        "max_bid_rate",
        "weighted_average_rate_within_limit",
        "term_days",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    first_cols = [
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

    existing_first_cols = [col for col in first_cols if col in df.columns]
    other_cols = [col for col in df.columns if col not in existing_first_cols]

    df = df[existing_first_cols + other_cols]

    return df


def download_one_day_repo(
    session: requests.Session,
    date_str: str,
    term_filter: int = 0,
) -> list[dict]:
    """
    Скачивает данные РЕПО за один конкретный день.
    """

    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": date_str,
        "UniDbQuery.To": date_str,
        "UniDbQuery.P1": str(term_filter),
    }

    response = session.get(
        PAGE_URL,
        params=params,
        headers=get_headers(),
        timeout=60,
    )

    response.raise_for_status()

    rows = extract_repo_tables_from_day(response.text, requested_date=date_str)

    return rows


def download_repo_excel(
    from_date: str = DEFAULT_FROM_DATE,
    to_date: str | None = None,
    term_filter: int = 0,
    sleep_seconds: float = 0.15,
) -> Path:
    """
    Скачивает итоги аукционов РЕПО с сайта ЦБ по дням,
    собирает найденные строки и сохраняет результат в Excel.

    term_filter:
    0 — любой срок
    1 — <= 5 дней
    5 — > 5 дней
    """

    if to_date is None:
        to_date = datetime.today().strftime("%d.%m.%Y")

    today_for_folder = datetime.today().strftime("%Y-%m-%d")

    folder_name = f"{today_for_folder}_repo"
    save_dir = Path("data") / "cbr" / folder_name
    save_dir.mkdir(parents=True, exist_ok=True)

    file_name = f"repo_{from_date.replace('.', '-')}_{to_date.replace('.', '-')}.xlsx"
    save_path = save_dir / file_name

    print("Скачиваю итоги аукционов РЕПО с сайта ЦБ...")
    print(f"Период: {from_date} — {to_date}")
    print(f"Фильтр срока: {term_filter}")
    print("Важно: РЕПО скачивается по дням, потому что сайт ЦБ плохо отдаёт большой период сразу.")

    start_date = parse_date(from_date)
    end_date = parse_date(to_date)

    if start_date > end_date:
        raise ValueError("from_date не может быть позже to_date.")

    session = create_session()

    all_rows = []
    current_date = start_date
    checked_days = 0

    while current_date <= end_date:
        date_str = format_date(current_date)

        try:
            rows = download_one_day_repo(
                session=session,
                date_str=date_str,
                term_filter=term_filter,
            )

            if rows:
                all_rows.extend(rows)
                print(f"{date_str}: найдено строк — {len(rows)}")
            else:
                print(f"{date_str}: данных нет")

        except Exception as error:
            print(f"{date_str}: ошибка скачивания — {error}")

        checked_days += 1
        current_date += timedelta(days=1)

        time.sleep(sleep_seconds)

    if not all_rows:
        raise RuntimeError(
            "За выбранный период не удалось найти данные РЕПО. "
            "Проверь даты или попробуй меньший период."
        )

    df = pd.DataFrame(all_rows)
    df = normalize_repo_columns(df)

    if "auction_date" in df.columns:
        df = df.sort_values("auction_date").reset_index(drop=True)

    df.to_excel(save_path, index=False)

    print("=" * 60)
    print(f"Проверено дней: {checked_days}")
    print(f"Строк скачано: {len(df)}")
    print(f"Файл сохранён: {save_path}")
    print("=" * 60)

    return save_path


if __name__ == "__main__":
    download_repo_excel(
        from_date="22.12.2020",
        to_date="26.05.2026",
        term_filter=0,
    )
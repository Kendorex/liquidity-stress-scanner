from pathlib import Path
from datetime import datetime
import requests
import pandas as pd

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


PAGE_URL = "https://cbr.ru/hd_base/KeyRate/"

# Ты решил начинать с этой даты
DEFAULT_FROM_DATE = "22.12.2020"


def create_session() -> requests.Session:
    """
    Создаём сессию requests с повторными попытками.
    Также отключаем системные proxy/VPN-настройки, потому что иногда
    из-за них ЦБ разрывает соединение.
    """

    session = requests.Session()

    # Важно: не брать proxy из Windows / переменных окружения
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


def download_key_rate_excel(
    from_date: str = DEFAULT_FROM_DATE,
    to_date: str | None = None,
) -> Path:
    """
    Скачивает страницу с ключевой ставкой Банка России,
    достаёт HTML-таблицу и сохраняет её в Excel.
    """

    if to_date is None:
        to_date = datetime.today().strftime("%d.%m.%Y")

    today_for_folder = datetime.today().strftime("%Y-%m-%d")

    folder_name = f"{today_for_folder}_key_rate"
    save_dir = Path("data") / "cbr" / folder_name
    save_dir.mkdir(parents=True, exist_ok=True)

    file_name = f"key_rate_{from_date.replace('.', '-')}_{to_date.replace('.', '-')}.xlsx"
    save_path = save_dir / file_name

    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": from_date,
        "UniDbQuery.To": to_date,
    }

    headers = {
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

    print("Скачиваю страницу ключевой ставки с сайта ЦБ...")
    print(f"Период: {from_date} — {to_date}")

    session = create_session()

    response = session.get(
        PAGE_URL,
        params=params,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    print("Страница успешно скачана. Ищу таблицу...")

    tables = pd.read_html(
        response.text,
        decimal=",",
        thousands=" ",
    )

    key_rate_df = None

    for table in tables:
        columns = [str(col).strip() for col in table.columns]

        if "Дата" in columns and "Ставка" in columns:
            key_rate_df = table.copy()
            break

    if key_rate_df is None:
        raise RuntimeError("Не удалось найти таблицу с ключевой ставкой на странице ЦБ.")

    key_rate_df.columns = ["date", "key_rate"]

    key_rate_df["date"] = pd.to_datetime(
        key_rate_df["date"],
        format="%d.%m.%Y",
        errors="coerce",
    )

    key_rate_df["key_rate"] = (
        key_rate_df["key_rate"]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )

    key_rate_df = key_rate_df.dropna(subset=["date", "key_rate"])
    key_rate_df = key_rate_df.sort_values("date").reset_index(drop=True)

    key_rate_df.to_excel(save_path, index=False)

    print(f"Строк скачано: {len(key_rate_df)}")
    print(f"Файл сохранён: {save_path}")

    return save_path


if __name__ == "__main__":
    download_key_rate_excel(
        from_date="22.12.2020",
        to_date=datetime.today().strftime("%d.%m.%Y"),
    )
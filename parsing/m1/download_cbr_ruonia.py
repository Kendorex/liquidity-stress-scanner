from pathlib import Path
from datetime import datetime
import requests

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://cbr.ru"
PAGE_URL = "https://cbr.ru/hd_base/ruonia/dynamics/"
DOWNLOAD_URL = "https://cbr.ru/Queries/UniDbQuery/DownloadExcel/14315"

DEFAULT_FROM_DATE = "11.01.2010"


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


def convert_date_to_cbr_excel_format(date_str: str) -> str:
    date_obj = datetime.strptime(date_str, "%d.%m.%Y")
    return date_obj.strftime("%m/%d/%Y")


def download_ruonia_excel(
    from_date: str = DEFAULT_FROM_DATE,
    to_date: str | None = None,
) -> Path:
    if to_date is None:
        to_date = datetime.today().strftime("%d.%m.%Y")

    save_dir = Path("data") / "m1" / "ruonia"
    save_dir.mkdir(parents=True, exist_ok=True)

    file_name = f"ruonia_{from_date.replace('.', '-')}_{to_date.replace('.', '-')}.xlsx"
    save_path = save_dir / file_name

    params = {
        "FromDate": convert_date_to_cbr_excel_format(from_date),
        "ToDate": convert_date_to_cbr_excel_format(to_date),
        "posted": "False",
        "backUrl": "/hd_base/ruonia/dynamics/",
    }

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
            "application/vnd.ms-excel,*/*"
        ),
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        "Referer": PAGE_URL,
        "Connection": "close",
    }

    print("Скачиваю Excel с RUONIA с сайта ЦБ...")
    print(f"Период: {from_date} — {to_date}")

    session = create_session()

    response = session.get(
        DOWNLOAD_URL,
        params=params,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()

    if "html" in content_type:
        raise RuntimeError(
            "ЦБ вернул HTML вместо Excel. Проверь даты или ссылку скачивания."
        )

    if len(response.content) < 1000:
        raise RuntimeError("Файл слишком маленький. Возможно, скачивание прошло некорректно.")

    save_path.write_bytes(response.content)

    print(f"Файл сохранён: {save_path}")
    print(f"Размер файла: {save_path.stat().st_size / 1024:.1f} КБ")

    return save_path


if __name__ == "__main__":
    download_ruonia_excel(
        from_date="11.01.2010",
        to_date=datetime.today().strftime("%d.%m.%Y"),
    )
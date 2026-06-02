from pathlib import Path
from datetime import datetime
import requests

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://cbr.ru"
PAGE_URL = "https://cbr.ru/hd_base/RReserves/"
DOWNLOAD_URL = "https://cbr.ru/vfs/hd_base/RReserves/required_reserves_table.xlsx"


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


def download_required_reserves_excel() -> Path:
    save_dir = Path("data") / "m1" / "required_reserves"
    save_dir.mkdir(parents=True, exist_ok=True)

    today_for_file = datetime.today().strftime("%Y-%m-%d")
    file_name = f"required_reserves_{today_for_file}.xlsx"
    save_path = save_dir / file_name

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

    print("Скачиваю Excel с обязательными резервами с сайта ЦБ...")
    print(f"Ссылка: {DOWNLOAD_URL}")

    session = create_session()

    response = session.get(
        DOWNLOAD_URL,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()

    if "html" in content_type:
        raise RuntimeError(
            "ЦБ вернул HTML вместо Excel. Возможно, ссылка изменилась или доступ временно ограничен."
        )

    if len(response.content) < 1000:
        raise RuntimeError("Файл слишком маленький. Возможно, скачивание прошло некорректно.")

    save_path.write_bytes(response.content)

    print(f"Файл сохранён: {save_path}")
    print(f"Размер файла: {save_path.stat().st_size / 1024:.1f} КБ")

    return save_path


if __name__ == "__main__":
    download_required_reserves_excel()
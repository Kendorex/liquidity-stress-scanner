from pathlib import Path
from urllib.parse import urljoin
import re
import requests
from bs4 import BeautifulSoup


# ============================================================
# M3 — Размещение ОФЗ
# Скрипт скачивает годовые Excel-файлы Минфина
# с результатами аукционов по размещению государственных ценных бумаг.
#
# Важно:
# этот файл только скачивает исходные данные.
# Расчёт cover ratio, флагов и MAD-score будет в preparation_data/m3.
# ============================================================


BASE_URL = "https://minfin.gov.ru"

SEARCH_URL = (
    "https://minfin.gov.ru/ru/document/"
    "?q_4=Результаты+проведенных+аукционов+по+размещению+государственных+ценных+бумаг"
    "&input_select_search=&input_select_search=&input_select_search="
    "&P_DATE_from_4=&P_DATE_to_4=&M_DATE_from_4=&M_DATE_to_4="
    "&t_4=8387194692553756541"
    "&order_4=&dir_4=desc"
    "&by_doc_number_4=0"
    "&INF_BLOCK_ID_4=0"
)

OUTPUT_DIR = Path("data/m3/ofz_auctions")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def download_html(url: str) -> str:
    print("Скачиваю страницу Минфина с результатами аукционов ОФЗ...")

    response = requests.get(
        url,
        timeout=60,
        headers={
            "User-Agent": "Mozilla/5.0",
        },
    )

    response.raise_for_status()
    return response.text


def extract_excel_links(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")

    found_links = []

    for link in soup.find_all("a", href=True):
        href = link["href"]

        if not href.lower().endswith(".xlsx"):
            continue

        if "Auction_Results" not in href:
            continue

        full_url = urljoin(BASE_URL, href)

        title = link.get("title") or link.get_text(" ", strip=True)

        year_match = re.search(r"(20\d{2})", full_url + " " + title)

        if not year_match:
            continue

        year = int(year_match.group(1))

        found_links.append(
            {
                "year": year,
                "url": full_url,
                "title": title,
            }
        )

    # Убираем дубли по году.
    # Если по одному году есть несколько ссылок, оставляем последнюю найденную.
    unique_by_year = {}

    for item in found_links:
        unique_by_year[item["year"]] = item

    result = list(unique_by_year.values())
    result = sorted(result, key=lambda x: x["year"])

    return result


def download_file(url: str, file_path: Path) -> None:
    response = requests.get(
        url,
        timeout=60,
        headers={
            "User-Agent": "Mozilla/5.0",
        },
    )

    response.raise_for_status()

    file_path.write_bytes(response.content)


def download_ofz_auction_files(links: list[dict]) -> None:
    if not links:
        raise RuntimeError("Не нашёл Excel-файлы с результатами аукционов ОФЗ на странице Минфина.")

    print()
    print("Найдены файлы Минфина:")

    for item in links:
        print(f"{item['year']}: {item['url']}")

    print()
    print("Начинаю скачивание...")

    for item in links:
        year = item["year"]
        url = item["url"]

        file_path = OUTPUT_DIR / f"ofz_auctions_{year}.xlsx"

        try:
            print(f"{year}: скачиваю файл...")

            download_file(url, file_path)

            print(f"{year}: сохранено -> {file_path}")

        except requests.RequestException as error:
            print(f"{year}: ошибка скачивания")
            print(error)

    print()
    print("Скачивание завершено.")


def main() -> None:
    html = download_html(SEARCH_URL)
    links = extract_excel_links(html)

    download_ofz_auction_files(links)

    print()
    print(f"Папка с файлами: {OUTPUT_DIR}")
    print(f"Количество найденных Excel-файлов: {len(links)}")


if __name__ == "__main__":
    main()
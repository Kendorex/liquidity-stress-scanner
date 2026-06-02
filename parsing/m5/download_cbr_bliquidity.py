from pathlib import Path
from urllib.parse import urlencode
import re

import requests
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# M5 — Средства федерального казначейства
#
# Источник: ЦБ РФ
# Таблица: "Дефицит/профицит ликвидности банковского сектора"
# ============================================================


BASE_URL = "https://www.cbr.ru/hd_base/bliquidity/"

DATE_FROM = "01.02.2014"
DATE_TO = "02.03.2026"

OUTPUT_DIR = Path("data/m5/cbr_bliquidity")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

DATE_FROM_FILE = DATE_FROM.replace(".", "-")
DATE_TO_FILE = DATE_TO.replace(".", "-")

RAW_HTML_FILE = OUTPUT_DIR / f"cbr_bliquidity_{DATE_FROM_FILE}_{DATE_TO_FILE}.html"
OUTPUT_XLSX_FILE = OUTPUT_DIR / f"cbr_bliquidity_{DATE_FROM_FILE}_{DATE_TO_FILE}.xlsx"
OUTPUT_CSV_FILE = OUTPUT_DIR / f"cbr_bliquidity_{DATE_FROM_FILE}_{DATE_TO_FILE}.csv"
COLUMNS = [
    "date",
    "liquidity_deficit_surplus",
    "liquidity_deficit_surplus_without_corr_accounts",
    "cbr_claims_total",
    "cbr_claims_repo_fxswap_auctions",
    "cbr_claims_secured_loans_auctions",
    "cbr_claims_repo_fxswap_standing",
    "cbr_claims_secured_loans_standing",
    "cbr_liabilities_total",
    "cbr_liabilities_deposit_auctions",
    "cbr_liabilities_deposit_standing",
    "cbr_liabilities_kobr",
    "non_standard_return_operations",
    "bank_correspondent_accounts",
    "required_reserves_averaging",
]


def build_url() -> str:
    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": DATE_FROM,
        "UniDbQuery.To": DATE_TO,
    }

    return BASE_URL + "?" + urlencode(params)


def download_html(url: str) -> str:
    print("Скачиваю HTML-страницу ЦБ...")
    print(f"Период: {DATE_FROM} — {DATE_TO}")

    response = requests.get(
        url,
        timeout=60,
        headers={
            "User-Agent": "Mozilla/5.0",
        },
    )

    response.raise_for_status()
    response.encoding = "utf-8"

    return response.text


def clean_number(value):
    if value is None:
        return None

    value = str(value).strip()

    if value in ("", "-", "—", "nan", "None"):
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


def parse_html_table(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "html.parser")

    table = soup.select_one("table.data")

    if table is None:
        raise RuntimeError("Не нашёл таблицу table.data на странице ЦБ.")

    rows = []

    for tr in table.find_all("tr"):
        cells = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
        if len(cells) != 15:
            continue

        if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", cells[0]):
            continue

        rows.append(cells)

    if not rows:
        raise RuntimeError("Не удалось извлечь строки данных из HTML-таблицы ЦБ.")

    df = pd.DataFrame(rows, columns=COLUMNS)

    df["date"] = pd.to_datetime(df["date"], errors="coerce", dayfirst=True)

    number_columns = [col for col in COLUMNS if col != "date"]

    for col in number_columns:
        df[col] = df[col].apply(clean_number)
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["date"])
    df = df.sort_values("date").reset_index(drop=True)

    df["source"] = "cbr_bliquidity"
    df["source_url"] = build_url()

    return df


def save_to_excel(df: pd.DataFrame) -> None:
    description = pd.DataFrame(
        [
            ["date", "Дата"],
            ["liquidity_deficit_surplus", "Дефицит (+) / профицит (-) ликвидности"],
            [
                "liquidity_deficit_surplus_without_corr_accounts",
                "Дефицит (+) / профицит (-) ликвидности без учета корсчетов",
            ],
            ["cbr_claims_total", "Требования Банка России к кредитным организациям"],
            ["cbr_claims_repo_fxswap_auctions", "Аукционы: репо и валютный своп"],
            ["cbr_claims_secured_loans_auctions", "Аукционы: обеспеченные кредиты"],
            ["cbr_claims_repo_fxswap_standing", "Операции постоянного действия: репо и валютный своп"],
            ["cbr_claims_secured_loans_standing", "Операции постоянного действия: обеспеченные кредиты"],
            ["cbr_liabilities_total", "Обязательства Банка России перед кредитными организациями"],
            ["cbr_liabilities_deposit_auctions", "Депозиты: аукционы"],
            ["cbr_liabilities_deposit_standing", "Депозиты: операции постоянного действия"],
            ["cbr_liabilities_kobr", "КОБР"],
            [
                "non_standard_return_operations",
                "Операции на возвратной основе, не относящиеся к стандартным инструментам ДКП",
            ],
            ["bank_correspondent_accounts", "Средства банков на корсчетах в Банке России"],
            ["required_reserves_averaging", "Обязательные резервы, подлежащие усреднению на корсчетах"],
            ["source", "Источник"],
            ["source_url", "Ссылка на источник"],
        ],
        columns=["column_name", "description"],
    )

    with pd.ExcelWriter(OUTPUT_XLSX_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="data", index=False)
        description.to_excel(writer, sheet_name="columns_description", index=False)


def main() -> None:
    print()
    print("M5 — ЦБ: дефицит/профицит ликвидности банковского сектора")

    url = build_url()
    html = download_html(url)

    RAW_HTML_FILE.write_text(html, encoding="utf-8")

    df = parse_html_table(html)

    save_to_excel(df)
    df.to_csv(OUTPUT_CSV_FILE, index=False, encoding="utf-8-sig")

    print()
    print("Готово.")
    print(f"HTML сохранён: {RAW_HTML_FILE}")
    print(f"Excel сохранён: {OUTPUT_XLSX_FILE}")
    print(f"CSV сохранён: {OUTPUT_CSV_FILE}")

    print()
    print(f"Количество строк: {len(df)}")
    print(f"Количество колонок: {len(df.columns)}")

    if not df.empty:
        print(f"Период в таблице: {df['date'].min().date()} — {df['date'].max().date()}")

    print()
    print("Первые строки:")
    print(df.head())


if __name__ == "__main__":
    main()
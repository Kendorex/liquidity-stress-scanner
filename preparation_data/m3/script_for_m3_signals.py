from pathlib import Path
import re

import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# M3 — Размещение ОФЗ
#
# Этот скрипт НЕ скачивает данные.
# Он берет уже скачанные Excel-файлы Минфина из data/m3/ofz_auctions/,
# приводит их к единому формату и считает сигналы для модуля M3.
#
# На выходе:
# data/m3/result/ofz_auctions_m3_signals.xlsx
# data/m3/result/m3_cover_ratio_daily.xlsx
# data/m3/result/m3_cover_ratio.png
# ============================================================


RAW_DIR = Path("data/m3/ofz_auctions")
RESULT_DIR = Path("data/m3/result")

OUTPUT_FILE = RESULT_DIR / "ofz_auctions_m3_signals.xlsx"
DAILY_OUTPUT_FILE = RESULT_DIR / "m3_cover_ratio_daily.xlsx"
PLOT_FILE = RESULT_DIR / "m3_cover_ratio.png"
MAD_WINDOW = 156
MIN_PERIODS = 20
START_YEAR = 2016
END_YEAR = 2026


def normalize_column_name(column) -> str:
    column = str(column).strip().lower()
    column = column.replace("\n", " ")
    column = re.sub(r"\s+", " ", column)
    return column


def clean_number(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value in ("", "-", "—", "nan", "None", "NaN"):
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


def find_header_row(raw_df: pd.DataFrame) -> int | None:
    max_rows = min(30, len(raw_df))

    for i in range(max_rows):
        row_text = " ".join(raw_df.iloc[i].astype(str).str.lower().tolist())

        has_date = "дата" in row_text
        has_demand = "спрос" in row_text
        has_placement = "размещ" in row_text or "предлож" in row_text

        if has_date and has_demand and has_placement:
            return i

    return None


def find_column(columns: list[str], keywords: list[str]) -> str | None:
    for column in columns:
        for keyword in keywords:
            if keyword in column:
                return column
    return None


def prepare_one_file(file_path: Path) -> pd.DataFrame:
    print(f"Обрабатываю файл: {file_path.name}")

    source_year_match = re.search(r"(20\d{2})", file_path.name)
    source_year = int(source_year_match.group(1)) if source_year_match else None

    raw = pd.read_excel(file_path, sheet_name=0, header=None)
    header_row = find_header_row(raw)

    if header_row is None:
        print(f"Не нашёл строку с заголовками в файле {file_path.name}. Файл пропущен.")
        return pd.DataFrame()

    df = pd.read_excel(file_path, sheet_name=0, header=header_row)
    df = df.dropna(how="all").copy()

    df.columns = [normalize_column_name(col) for col in df.columns]
    columns = list(df.columns)

    date_col = find_column(columns, ["дата"])
    issue_col = find_column(columns, ["выпуск", "серия", "номер"])
    offer_col = find_column(columns, ["предлож", "объем предложения", "объём предложения"])
    demand_col = find_column(columns, ["спрос"])
    placed_col = find_column(columns, ["размещ"])
    yield_col = find_column(columns, ["средневзвеш", "доход"])

    rename_map = {}

    if date_col:
        rename_map[date_col] = "auction_date"
    if issue_col:
        rename_map[issue_col] = "ofz_issue"
    if offer_col:
        rename_map[offer_col] = "offer_volume"
    if demand_col:
        rename_map[demand_col] = "demand_volume"
    if placed_col:
        rename_map[placed_col] = "placed_volume"
    if yield_col:
        rename_map[yield_col] = "weighted_avg_yield"

    df = df.rename(columns=rename_map)

    needed_columns = [
        "auction_date",
        "ofz_issue",
        "offer_volume",
        "demand_volume",
        "placed_volume",
        "weighted_avg_yield",
    ]

    for col in needed_columns:
        if col not in df.columns:
            df[col] = None

    df = df[needed_columns].copy()

    df["auction_date"] = pd.to_datetime(df["auction_date"], errors="coerce", dayfirst=True)

    number_columns = [
        "offer_volume",
        "demand_volume",
        "placed_volume",
        "weighted_avg_yield",
    ]

    for col in number_columns:
        df[col] = df[col].apply(clean_number)
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["source_year"] = source_year
    df["source_file"] = file_path.name
    df = df.dropna(subset=["auction_date"])
    df = df[df["ofz_issue"].notna()]
    df = df[~df["ofz_issue"].astype(str).str.lower().str.contains("итого|всего|total", na=False)]

    return df


def calculate_mad_score(series: pd.Series, window: int = MAD_WINDOW) -> pd.Series:
    series = pd.to_numeric(series, errors="coerce")

    rolling_median = series.rolling(window=window, min_periods=MIN_PERIODS).median()

    rolling_mad = series.rolling(window=window, min_periods=MIN_PERIODS).apply(
        lambda x: (abs(x - x.median())).median(),
        raw=False,
    )

    score = (series - rolling_median) / (1.4826 * rolling_mad)
    score = score.replace([float("inf"), float("-inf")], pd.NA)

    return score


def build_m3_signals() -> pd.DataFrame:
    files = sorted(RAW_DIR.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(
            f"В папке {RAW_DIR} нет Excel-файлов. "
            f"Сначала запусти parsing/m3/download_minfin_ofz_auctions.py"
        )

    frames = []

    for file_path in files:
        one_file_df = prepare_one_file(file_path)

        if not one_file_df.empty:
            frames.append(one_file_df)

    if not frames:
        raise RuntimeError("Не удалось обработать ни один файл M3.")

    df = pd.concat(frames, ignore_index=True)

    df["auction_date"] = pd.to_datetime(df["auction_date"], errors="coerce")
    df = df[
        (df["auction_date"].dt.year >= START_YEAR)
        & (df["auction_date"].dt.year <= END_YEAR)
    ].copy()

    if df.empty:
        raise RuntimeError(f"После фильтра {START_YEAR}–{END_YEAR} данных не осталось.")

    df = df.sort_values(["auction_date", "ofz_issue"]).reset_index(drop=True)

    numeric_columns = [
        "offer_volume",
        "demand_volume",
        "placed_volume",
        "weighted_avg_yield",
    ]

    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["cover_ratio"] = df["demand_volume"] / df["offer_volume"]

    df.loc[
        df["offer_volume"].isna() | (df["offer_volume"] == 0),
        "cover_ratio",
    ] = pd.NA

    df["cover_ratio"] = pd.to_numeric(df["cover_ratio"], errors="coerce")
    df["flag_nedospros"] = (df["cover_ratio"] < 1.2).astype(int)
    df["flag_perespros"] = (df["cover_ratio"] > 2.0).astype(int)
    df["mad_score_cover_raw"] = calculate_mad_score(df["cover_ratio"])
    df["mad_score_cover"] = -df["mad_score_cover_raw"]
    df["weighted_avg_yield"] = pd.to_numeric(df["weighted_avg_yield"], errors="coerce")

    df["rolling_median_yield"] = (
        df["weighted_avg_yield"]
        .rolling(window=MAD_WINDOW, min_periods=MIN_PERIODS)
        .median()
    )

    df["proxy_yield_spread"] = df["weighted_avg_yield"] - df["rolling_median_yield"]
    df["proxy_yield_spread"] = pd.to_numeric(df["proxy_yield_spread"], errors="coerce")
    df["yield_spread"] = df["proxy_yield_spread"]
    df["mad_score_yield_spread"] = calculate_mad_score(df["proxy_yield_spread"])
    cover_stress_score = pd.to_numeric(df["mad_score_cover"], errors="coerce").fillna(0.0).clip(lower=0.0, upper=4.0) / 4.0 * 100.0
    yield_stress_score = pd.to_numeric(df["mad_score_yield_spread"], errors="coerce").fillna(0.0).clip(lower=0.0, upper=4.0) / 4.0 * 100.0

    df["m3_score"] = (
        0.65 * cover_stress_score
        + 0.25 * yield_stress_score
        + 0.10 * df["flag_nedospros"] * 100.0
    ).clip(lower=0.0, upper=100.0)

    df["m3_flag"] = (
        (df["m3_score"] >= 40)
        | ((df["cover_ratio"] < 0.8) & (yield_stress_score >= 25))
    ).astype(int)

    df["m3_signal"] = df["m3_score"]

    df["m3_comment"] = "proxy_yield_spread: отклонение доходности от собственной скользящей медианы, не от кривой ОФЗ"

    final_columns = [
        "auction_date",
        "ofz_issue",
        "offer_volume",
        "demand_volume",
        "placed_volume",
        "cover_ratio",
        "weighted_avg_yield",
        "rolling_median_yield",
        "proxy_yield_spread",
        "yield_spread",
        "mad_score_cover",
        "mad_score_yield_spread",
        "flag_nedospros",
        "flag_perespros",
        "m3_score",
        "m3_flag",
        "m3_signal",
        "m3_comment",
        "source_year",
        "source_file",
    ]

    df = df[final_columns]

    return df


def build_daily_cover_ratio(df: pd.DataFrame) -> pd.DataFrame:
    daily = (
        df.groupby("auction_date", as_index=False)
        .agg(
            cover_ratio=("cover_ratio", "mean"),
            demand_volume=("demand_volume", "sum"),
            offer_volume=("offer_volume", "sum"),
            placed_volume=("placed_volume", "sum"),
            weighted_avg_yield=("weighted_avg_yield", "mean"),
            proxy_yield_spread=("proxy_yield_spread", "mean"),
            yield_spread=("yield_spread", "mean"),
            mad_score_cover=("mad_score_cover", "mean"),
            mad_score_yield_spread=("mad_score_yield_spread", "mean"),
            flag_nedospros=("flag_nedospros", "max"),
            flag_perespros=("flag_perespros", "max"),
            m3_score=("m3_score", "max"),
            m3_flag=("m3_flag", "max"),
            m3_signal=("m3_signal", "max"),
        )
        .sort_values("auction_date")
        .reset_index(drop=True)
    )

    return daily


def plot_cover_ratio(daily: pd.DataFrame) -> None:
    plot_df = daily.dropna(subset=["auction_date", "cover_ratio"]).copy()

    if plot_df.empty:
        print("Нет данных для построения графика cover ratio.")
        return

    plot_df = plot_df.sort_values("auction_date").reset_index(drop=True)
    plot_df["cover_ratio_smooth"] = (
        plot_df["cover_ratio"]
        .rolling(window=5, min_periods=1)
        .mean()
    )

    upper_limit = plot_df["cover_ratio"].quantile(0.95)
    upper_limit = max(upper_limit, 2.5)
    plot_df["cover_ratio_for_plot"] = plot_df["cover_ratio"].clip(upper=upper_limit)
    plot_df["cover_ratio_smooth_for_plot"] = plot_df["cover_ratio_smooth"].clip(upper=upper_limit)
    nedospros = plot_df[plot_df["flag_nedospros"] == 1]
    perespros = plot_df[plot_df["flag_perespros"] == 1]

    plt.figure(figsize=(16, 7))
    plt.axhspan(
        0,
        1.2,
        alpha=0.12,
        label="Зона недоспроса",
    )

    plt.axhspan(
        2.0,
        upper_limit,
        alpha=0.08,
        label="Зона переспроса",
    )

    plt.plot(
        plot_df["auction_date"],
        plot_df["cover_ratio_for_plot"],
        linewidth=1.0,
        alpha=0.45,
        label="Дневной cover ratio",
    )

    plt.plot(
        plot_df["auction_date"],
        plot_df["cover_ratio_smooth_for_plot"],
        linewidth=2.4,
        label="Сглаженный cover ratio",
    )

    plt.axhline(
        1.2,
        linestyle="--",
        linewidth=1.4,
        label="Порог недоспроса: 1.2",
    )

    plt.axhline(
        2.0,
        linestyle="--",
        linewidth=1.4,
        label="Порог переспроса: 2.0",
    )

    plt.scatter(
        nedospros["auction_date"],
        nedospros["cover_ratio_for_plot"],
        s=45,
        marker="v",
        label="Недоспрос",
        zorder=5,
    )

    plt.scatter(
        perespros["auction_date"],
        perespros["cover_ratio_for_plot"],
        s=35,
        marker="^",
        label="Переспрос",
        zorder=5,
    )

    plt.title(
        f"M3 — Cover ratio аукционов ОФЗ, {START_YEAR}–{END_YEAR}",
        fontsize=15,
        pad=14,
    )
    plt.xlabel("Дата")
    plt.ylabel("Cover ratio: спрос / предложение")

    plt.ylim(0, upper_limit * 1.05)

    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left", frameon=True)

    plt.tight_layout()
    plt.savefig(PLOT_FILE, dpi=220)
    plt.close()


def main() -> None:
    print("M3 — Размещение ОФЗ")
    print("Готовлю сигналы по результатам аукционов Минфина")

    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    df = build_m3_signals()
    daily = build_daily_cover_ratio(df)

    df.to_excel(OUTPUT_FILE, index=False)
    daily.to_excel(DAILY_OUTPUT_FILE, index=False)

    plot_cover_ratio(daily)

    print()
    print("Готово.")
    print(f"Итоговый файл сигналов сохранён: {OUTPUT_FILE}")
    print(f"Дневная таблица для графика сохранена: {DAILY_OUTPUT_FILE}")
    print(f"График сохранён: {PLOT_FILE}")

    print()
    print(f"Количество строк в итоговой таблице: {len(df)}")
    print(f"Количество дат в дневной таблице: {len(daily)}")

    if not df.empty:
        print(f"Период: {df['auction_date'].min().date()} — {df['auction_date'].max().date()}")
        print(f"Количество выпусков/строк с недоспросом: {int(df['flag_nedospros'].sum())}")
        print(f"Количество выпусков/строк с переспросом: {int(df['flag_perespros'].sum())}")

        print()
        print("Первые строки итоговой таблицы:")
        print(df.head())


if __name__ == "__main__":
    main()
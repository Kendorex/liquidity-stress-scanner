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

# Примерно 3 года наблюдений.
# У ОФЗ может быть несколько выпусков в один день, поэтому берем окно по строкам.
MAD_WINDOW = 156
MIN_PERIODS = 20

# Для текущего проекта считаем и рисуем только доступный период 2016–2026.
START_YEAR = 2016
END_YEAR = 2026


def normalize_column_name(column) -> str:
    """Приводим название колонки к удобному виду."""
    column = str(column).strip().lower()
    column = column.replace("\n", " ")
    column = re.sub(r"\s+", " ", column)
    return column


def clean_number(value):
    """Чистим числа из Excel: пробелы, запятые, лишний текст."""
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value in ("", "-", "—", "nan", "None", "NaN"):
        return None

    value = value.replace("\xa0", "")
    value = value.replace(" ", "")
    value = value.replace(",", ".")

    # Оставляем только цифры, точку и минус.
    value = re.sub(r"[^0-9.\-]", "", value)

    if value in ("", "-", ".", "-."):
        return None

    try:
        return float(value)
    except ValueError:
        return None


def find_header_row(raw_df: pd.DataFrame) -> int | None:
    """
    Ищем строку, где находятся заголовки таблицы.
    В файлах Минфина заголовок может быть не в первой строке.
    """
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
    """Ищем колонку по ключевым словам."""
    for column in columns:
        for keyword in keywords:
            if keyword in column:
                return column
    return None


def prepare_one_file(file_path: Path) -> pd.DataFrame:
    """Читает один Excel-файл Минфина и приводит его к единому формату."""
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

    # Убираем служебные строки, итоги и пустые строки.
    df = df.dropna(subset=["auction_date"])
    df = df[df["ofz_issue"].notna()]
    df = df[~df["ofz_issue"].astype(str).str.lower().str.contains("итого|всего|total", na=False)]

    return df


def calculate_mad_score(series: pd.Series, window: int = MAD_WINDOW) -> pd.Series:
    """
    MAD-score = (x - rolling_median) / (1.4826 * rolling_MAD)

    1.4826 нужен, чтобы MAD был сопоставим со стандартным отклонением
    при нормальном распределении.
    """
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

    # Считаем только период 2016–2026.
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

    # Главный показатель M3: спрос / предложение.
    df["cover_ratio"] = df["demand_volume"] / df["offer_volume"]

    df.loc[
        df["offer_volume"].isna() | (df["offer_volume"] == 0),
        "cover_ratio",
    ] = pd.NA

    df["cover_ratio"] = pd.to_numeric(df["cover_ratio"], errors="coerce")

    # Флаги по ТЗ.
    df["flag_nedospros"] = (df["cover_ratio"] < 1.2).astype(int)
    df["flag_perespros"] = (df["cover_ratio"] > 2.0).astype(int)

    # Для стресса важен именно низкий cover_ratio.
    # Поэтому знак разворачиваем: чем ниже cover_ratio относительно нормы,
    # тем выше стрессовый score.
    df["mad_score_cover_raw"] = calculate_mad_score(df["cover_ratio"])
    df["mad_score_cover"] = -df["mad_score_cover_raw"]

    # ------------------------------------------------------------
    # MVP-версия yield_spread
    #
    # В полном варианте по ТЗ yield_spread должен считаться как отклонение
    # доходности размещения от кривой ОФЗ / ближайших выпусков.
    #
    # Пока отдельной кривой ОФЗ нет, поэтому используем proxy:
    # yield_spread = текущая средневзвешенная доходность
    #                − скользящая медианная доходность аукционов.
    #
    # Это не полноценная рыночная кривая ОФЗ, но для MVP дает рабочий
    # индикатор отклонения доходности от собственной исторической нормы.
    # ------------------------------------------------------------

    df["weighted_avg_yield"] = pd.to_numeric(df["weighted_avg_yield"], errors="coerce")

    df["rolling_median_yield"] = (
        df["weighted_avg_yield"]
        .rolling(window=MAD_WINDOW, min_periods=MIN_PERIODS)
        .median()
    )

    df["yield_spread"] = df["weighted_avg_yield"] - df["rolling_median_yield"]
    df["yield_spread"] = pd.to_numeric(df["yield_spread"], errors="coerce")

    df["mad_score_yield_spread"] = calculate_mad_score(df["yield_spread"])

    final_columns = [
        "auction_date",
        "ofz_issue",
        "offer_volume",
        "demand_volume",
        "placed_volume",
        "cover_ratio",
        "weighted_avg_yield",
        "rolling_median_yield",
        "yield_spread",
        "mad_score_cover",
        "mad_score_yield_spread",
        "flag_nedospros",
        "flag_perespros",
        "source_year",
        "source_file",
    ]

    df = df[final_columns]

    return df


def build_daily_cover_ratio(df: pd.DataFrame) -> pd.DataFrame:
    """Собираем дневную таблицу для графика."""
    daily = (
        df.groupby("auction_date", as_index=False)
        .agg(
            cover_ratio=("cover_ratio", "mean"),
            demand_volume=("demand_volume", "sum"),
            offer_volume=("offer_volume", "sum"),
            placed_volume=("placed_volume", "sum"),
            weighted_avg_yield=("weighted_avg_yield", "mean"),
            yield_spread=("yield_spread", "mean"),
            mad_score_cover=("mad_score_cover", "mean"),
            mad_score_yield_spread=("mad_score_yield_spread", "mean"),
            flag_nedospros=("flag_nedospros", "max"),
            flag_perespros=("flag_perespros", "max"),
        )
        .sort_values("auction_date")
        .reset_index(drop=True)
    )

    return daily


def plot_cover_ratio(daily: pd.DataFrame) -> None:
    """Строим аккуратный график Cover ratio ОФЗ."""
    plot_df = daily.dropna(subset=["auction_date", "cover_ratio"]).copy()

    if plot_df.empty:
        print("Нет данных для построения графика cover ratio.")
        return

    plot_df = plot_df.sort_values("auction_date").reset_index(drop=True)

    # Сглаживание, чтобы график не выглядел как шум.
    # Берём короткое окно, потому что аукционы проходят не каждый день.
    plot_df["cover_ratio_smooth"] = (
        plot_df["cover_ratio"]
        .rolling(window=5, min_periods=1)
        .mean()
    )

    # Ограничиваем верхний масштаб, чтобы редкие выбросы не портили картинку.
    upper_limit = plot_df["cover_ratio"].quantile(0.95)
    upper_limit = max(upper_limit, 2.5)

    plot_df["cover_ratio_for_plot"] = plot_df["cover_ratio"].clip(upper=upper_limit)
    plot_df["cover_ratio_smooth_for_plot"] = plot_df["cover_ratio_smooth"].clip(upper=upper_limit)

    nedospros = plot_df[plot_df["flag_nedospros"] == 1]
    perespros = plot_df[plot_df["flag_perespros"] == 1]

    plt.figure(figsize=(16, 7))

    # Зоны для визуальной интерпретации.
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

    # Основная дневная линия — тонкая.
    plt.plot(
        plot_df["auction_date"],
        plot_df["cover_ratio_for_plot"],
        linewidth=1.0,
        alpha=0.45,
        label="Дневной cover ratio",
    )

    # Сглаженная линия — главная.
    plt.plot(
        plot_df["auction_date"],
        plot_df["cover_ratio_smooth_for_plot"],
        linewidth=2.4,
        label="Сглаженный cover ratio",
    )

    # Пороговые уровни.
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

    # Точки сигналов.
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
    print("=" * 70)
    print("M3 — Размещение ОФЗ")
    print("Готовлю сигналы по результатам аукционов Минфина")
    print("=" * 70)

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
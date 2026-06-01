from pathlib import Path
import re

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


PROJECT_ROOT = Path(__file__).resolve().parents[2]

REPO_DIR = PROJECT_ROOT / "data" / "m2" / "repo"
KEY_RATE_DIR = PROJECT_ROOT / "data" / "m2" / "key_rate"
RESULTS_DIR = PROJECT_ROOT / "data" / "m2" / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)


MAIN_TERM_DAYS = 7

# Историческое окно для нормализации: примерно 3 года.
WINDOW_DAYS = 1095

# Минимум прошлых наблюдений, чтобы считать MAD-сигнал осмысленным.
MIN_PERIODS = 20

# Нижние пороги MAD нужны, чтобы спокойная история с почти нулевой
# волатильностью не превращала любое небольшое отклонение в сигнал 100.
COVER_MAD_FLOOR = 0.10
SPREAD_MAD_FLOOR = 0.10


def find_latest_xlsx(folder: Path) -> Path:
    """
    Находит последний нормальный Excel-файл в папке.
    Временные файлы Excel вида ~$file.xlsx игнорируются.
    """

    files = [
        file for file in folder.glob("*.xlsx")
        if not file.name.startswith("~$")
    ]

    if not files:
        raise FileNotFoundError(f"В папке нет Excel-файлов: {folder}")

    return max(files, key=lambda file: file.stat().st_mtime)


def parse_auction_date_from_datetime(value) -> pd.Timestamp:
    """
    Достаёт реальную дату аукциона из строки вида:
    '26.05.2026 на 13:30'.

    Это важно, потому что сайт ЦБ иногда на запрос следующего дня
    возвращает последнюю доступную таблицу. Тогда дата запроса может
    отличаться от настоящей даты аукциона.
    """

    if pd.isna(value):
        return pd.NaT

    text = str(value)
    match = re.search(r"(\d{2}\.\d{2}\.\d{4})", text)

    if not match:
        return pd.NaT

    return pd.to_datetime(match.group(1), format="%d.%m.%Y", errors="coerce")


def to_number(series: pd.Series) -> pd.Series:
    """
    Надёжно переводит числа из Excel в float.
    Поддерживает и обычный float, и русскую запись через запятую.
    """

    return pd.to_numeric(
        series.astype(str)
        .str.replace("\xa0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce",
    )


def calculate_past_mad_score(
    df: pd.DataFrame,
    date_column: str,
    value_column: str,
    mask_column: str | None = None,
    window_days: int = WINDOW_DAYS,
    min_periods: int = MIN_PERIODS,
    mad_floor: float = 0.0,
) -> pd.DataFrame:
    """
    Считает median, MAD и MAD-score только по прошлой истории.

    Это важнее обычного rolling(), потому что для раннего сигнала нельзя
    использовать текущую точку в собственной базе сравнения.
    """

    result = df.copy()

    median_col = f"{value_column}_median_3y"
    mad_col = f"{value_column}_mad_3y"
    mad_used_col = f"{value_column}_mad_used_3y"
    score_col = f"mad_score_{value_column}"

    result[median_col] = np.nan
    result[mad_col] = np.nan
    result[mad_used_col] = np.nan
    result[score_col] = np.nan

    work = result[[date_column, value_column]].copy()
    work["_original_index"] = result.index

    if mask_column is not None:
        work = work.loc[result[mask_column] == 1].copy()

    work = work.dropna(subset=[date_column, value_column]).copy()
    work = work.sort_values([date_column, "_original_index"]).reset_index(drop=True)

    if work.empty:
        return result

    for _, row in work.iterrows():
        current_date = row[date_column]
        current_value = row[value_column]
        original_index = row["_original_index"]

        start_date = current_date - pd.Timedelta(days=window_days)

        history = work[
            (work[date_column] < current_date)
            & (work[date_column] >= start_date)
        ][value_column].dropna()

        if len(history) < min_periods:
            continue

        median_value = float(np.median(history))
        mad_value = float(np.median(np.abs(history - median_value)))
        mad_used = max(mad_value, mad_floor)

        if mad_used == 0:
            continue

        score = (current_value - median_value) / mad_used

        result.loc[original_index, median_col] = median_value
        result.loc[original_index, mad_col] = mad_value
        result.loc[original_index, mad_used_col] = mad_used
        result.loc[original_index, score_col] = score

    return result


def clean_repo() -> pd.DataFrame:
    """Читает и очищает сырые итоги аукционов РЕПО."""

    print("=" * 80)
    print("Очищаю итоги аукционов РЕПО")
    print("=" * 80)

    file_path = find_latest_xlsx(REPO_DIR)
    print(f"Файл: {file_path}")

    df = pd.read_excel(file_path)
    df = df.dropna(how="all").copy()

    if "auction_date" not in df.columns:
        raise ValueError("В файле РЕПО нет колонки auction_date")

    df["auction_date"] = pd.to_datetime(df["auction_date"], errors="coerce")

    if "auction_datetime" in df.columns:
        parsed_dates = df["auction_datetime"].apply(parse_auction_date_from_datetime)
        df["auction_date"] = parsed_dates.combine_first(df["auction_date"])

    for column in ["first_leg_date", "second_leg_date"]:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")

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
            df[column] = to_number(df[column])

    needed_columns = [
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

    existing_columns = [column for column in needed_columns if column in df.columns]
    df = df[existing_columns].copy()

    if "auction_type" in df.columns:
        df = df[
            df["auction_type"]
            .astype(str)
            .str.lower()
            .str.contains("репо", na=False)
        ].copy()

    df = df.dropna(subset=["auction_date", "term_days"]).copy()

    duplicate_subset = [
        "auction_date",
        "auction_datetime",
        "auction_type",
        "term_days",
        "repo_demand_mln_rub",
        "repo_deals_total_mln_rub",
        "repo_deals_within_limit_mln_rub",
        "cutoff_rate",
        "weighted_average_rate",
        "first_leg_date",
        "second_leg_date",
    ]

    duplicate_subset = [column for column in duplicate_subset if column in df.columns]

    before_duplicates = len(df)
    df = df.drop_duplicates(subset=duplicate_subset).copy()
    removed_duplicates = before_duplicates - len(df)

    df["is_main_7d_repo"] = (df["term_days"] == MAIN_TERM_DAYS).astype(int)

    df["repo_demand_bln_rub"] = df["repo_demand_mln_rub"] / 1000
    df["repo_deals_total_bln_rub"] = df["repo_deals_total_mln_rub"] / 1000

    if "repo_deals_within_limit_mln_rub" in df.columns:
        df["repo_deals_within_limit_bln_rub"] = (
            df["repo_deals_within_limit_mln_rub"] / 1000
        )
    else:
        df["repo_deals_within_limit_bln_rub"] = np.nan

    df["cover_ratio"] = df["repo_demand_mln_rub"] / df["repo_deals_total_mln_rub"]
    df.loc[df["repo_deals_total_mln_rub"] <= 0, "cover_ratio"] = np.nan

    df["satisfaction_ratio"] = df["repo_deals_total_mln_rub"] / df["repo_demand_mln_rub"]
    df.loc[df["repo_demand_mln_rub"] <= 0, "satisfaction_ratio"] = np.nan

    df["unmet_demand_bln_rub"] = (
        df["repo_demand_mln_rub"] - df["repo_deals_total_mln_rub"]
    ) / 1000
    df["unmet_demand_bln_rub"] = df["unmet_demand_bln_rub"].clip(lower=0)

    df = df.sort_values(["auction_date", "auction_datetime"]).reset_index(drop=True)

    print(f"Строк после очистки РЕПО: {len(df)}")
    print(f"Удалено дублей: {removed_duplicates}")
    print(f"Из них 7-дневных аукционов: {int(df['is_main_7d_repo'].sum())}")

    return df


def clean_key_rate() -> pd.DataFrame:
    """Читает и очищает ключевую ставку."""

    print("=" * 80)
    print("Очищаю ключевую ставку")
    print("=" * 80)

    file_path = find_latest_xlsx(KEY_RATE_DIR)
    print(f"Файл: {file_path}")

    df = pd.read_excel(file_path)

    if len(df.columns) < 2:
        raise ValueError("В файле ключевой ставки должно быть минимум две колонки")

    df = df.rename(columns={df.columns[0]: "date", df.columns[1]: "key_rate"})

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["key_rate"] = to_number(df["key_rate"])

    df = df.dropna(subset=["date", "key_rate"])
    df = df.drop_duplicates(subset=["date"]).copy()
    df = df.sort_values("date").reset_index(drop=True)

    print(f"Строк после очистки ключевой ставки: {len(df)}")

    return df


def add_key_rate(repo_df: pd.DataFrame, key_rate_df: pd.DataFrame) -> pd.DataFrame:
    """
    Подтягивает ключевую ставку к дате аукциона.

    РЕПО не протягивается по календарю.
    Протягивается только ключевая ставка, потому что она действует
    до следующего изменения.
    """

    repo_df = repo_df.sort_values("auction_date").copy()
    key_rate_df = key_rate_df.sort_values("date").copy()

    merged = pd.merge_asof(
        repo_df,
        key_rate_df,
        left_on="auction_date",
        right_on="date",
        direction="backward",
    )

    merged = merged.drop(columns=["date"], errors="ignore")

    merged["effective_repo_rate"] = merged["cutoff_rate"].combine_first(
        merged["weighted_average_rate"]
    )

    merged["cutoff_spread"] = merged["effective_repo_rate"] - merged["key_rate"]

    return merged


def cover_ratio_level_score(cover_ratio: pd.Series) -> pd.Series:
    """
    Базовая шкала от уровня переспроса.

    Это главный компонент M2:
    - около 1.0 — нормальная ситуация;
    - 1.2–2.0 — нарастает напряжение;
    - выше 2.0 — явный переспрос;
    - выше 3.0 — очень сильный переспрос.
    """

    score = pd.Series(0.0, index=cover_ratio.index)

    low_mask = (cover_ratio > 1.0) & (cover_ratio <= 1.2)
    medium_mask = (cover_ratio > 1.2) & (cover_ratio <= 2.0)
    high_mask = (cover_ratio > 2.0) & (cover_ratio <= 3.0)
    extreme_mask = cover_ratio > 3.0

    score.loc[low_mask] = (
        (cover_ratio.loc[low_mask] - 1.0) / 0.2 * 20
    )

    score.loc[medium_mask] = (
        20
        + (cover_ratio.loc[medium_mask] - 1.2) / 0.8 * 35
    )

    score.loc[high_mask] = (
        55
        + (cover_ratio.loc[high_mask] - 2.0) / 1.0 * 25
    )

    score.loc[extreme_mask] = (
        80
        + (cover_ratio.loc[extreme_mask] - 3.0) / 2.0 * 15
    )

    return score.clip(lower=0, upper=95)


def calculate_m2_signals(m2_df: pd.DataFrame) -> pd.DataFrame:
    """Считает итоговые сигналы M2."""

    print("=" * 80)
    print("Считаю сигналы M2")
    print("=" * 80)

    df = m2_df.copy()
    df = df.sort_values(["auction_date", "auction_datetime"]).reset_index(drop=True)

    df["signal_scope"] = np.where(
        df["is_main_7d_repo"] == 1,
        "main_7d_repo",
        "other_repo_term",
    )

    df = calculate_past_mad_score(
        df=df,
        date_column="auction_date",
        value_column="cover_ratio",
        mask_column="is_main_7d_repo",
        window_days=WINDOW_DAYS,
        min_periods=MIN_PERIODS,
        mad_floor=COVER_MAD_FLOOR,
    )

    df = calculate_past_mad_score(
        df=df,
        date_column="auction_date",
        value_column="cutoff_spread",
        mask_column="is_main_7d_repo",
        window_days=WINDOW_DAYS,
        min_periods=MIN_PERIODS,
        mad_floor=SPREAD_MAD_FLOOR,
    )

    df["positive_mad_score_cover"] = df["mad_score_cover_ratio"].clip(lower=0).fillna(0)
    df["positive_mad_score_cutoff_spread"] = (
        df["mad_score_cutoff_spread"].clip(lower=0).fillna(0)
    )

    df["flag_demand"] = (
        (df["cover_ratio"] > 2.0)
        & (df["is_main_7d_repo"] == 1)
    ).astype(int)

    df["flag_rate_pressure"] = (
        (df["cutoff_spread"] > 0)
        & (df["positive_mad_score_cutoff_spread"] >= 1.0)
        & (df["is_main_7d_repo"] == 1)
    ).astype(int)

    df["stress_pattern_flag"] = (
        (df["flag_demand"] == 1)
        & (df["flag_rate_pressure"] == 1)
    ).astype(int)

    df["signal_ready"] = (
        df["auction_date"].notna()
        & (df["is_main_7d_repo"] == 1)
        & df["cover_ratio"].notna()
        & df["effective_repo_rate"].notna()
        & df["key_rate"].notna()
        & df["mad_score_cover_ratio"].notna()
        & df["mad_score_cutoff_spread"].notna()
    ).astype(int)

    # ------------------------------------------------------------------
    # Итоговый M2 signal
    # ------------------------------------------------------------------
    # Основной фактор — уровень переспроса cover_ratio.
    # MAD по cover_ratio и спред ставки к ключевой используются как усилители,
    # но не ломают шкалу и не превращают нормальный спрос в стресс.
    # ------------------------------------------------------------------

    df["demand_level_score"] = cover_ratio_level_score(df["cover_ratio"])

    df["demand_mad_bonus"] = (
        df["positive_mad_score_cover"].clip(lower=0, upper=4) / 4 * 10
    )

    df["rate_pressure_bonus"] = np.where(
        df["cutoff_spread"] > 0,
        df["positive_mad_score_cutoff_spread"].clip(lower=0, upper=4) / 4 * 15,
        0,
    )

    df["m2_signal"] = (
        df["demand_level_score"]
        + df["demand_mad_bonus"]
        + df["rate_pressure_bonus"]
    ).clip(lower=0, upper=100)

    # Если сигнала по полной методике ещё нельзя считать, ставим 0.
    df.loc[df["signal_ready"] == 0, "m2_signal"] = 0

    # Смысловые ограничения:
    # ставка не может одна сделать стресс, если спрос почти равен размещению.
    almost_normal_demand = (
        (df["signal_ready"] == 1)
        & (df["cover_ratio"] <= 1.2)
    )
    df.loc[almost_normal_demand, "m2_signal"] = (
        df.loc[almost_normal_demand, "m2_signal"].clip(upper=35)
    )

    # До cover_ratio <= 2 это ещё не полноценный переспрос.
    # Может быть напряжение, но не стресс.
    moderate_demand = (
        (df["signal_ready"] == 1)
        & (df["cover_ratio"] > 1.2)
        & (df["cover_ratio"] <= 2.0)
    )
    df.loc[moderate_demand, "m2_signal"] = (
        df.loc[moderate_demand, "m2_signal"].clip(upper=70)
    )

    # Полноценный стрессовый паттерн — переспрос плюс давление по ставке.
    stress_pattern = (
        (df["signal_ready"] == 1)
        & (df["cover_ratio"] > 2.0)
        & (df["flag_rate_pressure"] == 1)
    )
    df.loc[stress_pattern, "m2_signal"] = (
        df.loc[stress_pattern, "m2_signal"] + 5
    ).clip(upper=100)

    # Единый стандарт выхода для LSI.
    df["date"] = df["auction_date"]
    df["m2_score"] = df["m2_signal"]
    df["m2_flag"] = ((df["flag_demand"] == 1) | (df["flag_rate_pressure"] == 1)).astype(int)

    df["m2_signal_zone"] = pd.cut(
        df["m2_signal"],
        bins=[-0.1, 40, 70, 100],
        labels=["норма", "напряжение", "стресс"],
    )

    output_columns = [
        "date",
        "auction_date",
        "auction_datetime",
        "auction_type",
        "term_days",
        "signal_scope",
        "repo_demand_bln_rub",
        "repo_deals_total_bln_rub",
        "repo_deals_within_limit_bln_rub",
        "unmet_demand_bln_rub",
        "cover_ratio",
        "satisfaction_ratio",
        "cutoff_rate",
        "weighted_average_rate",
        "effective_repo_rate",
        "key_rate",
        "cutoff_spread",
        "cover_ratio_median_3y",
        "cover_ratio_mad_3y",
        "cover_ratio_mad_used_3y",
        "mad_score_cover_ratio",
        "cutoff_spread_median_3y",
        "cutoff_spread_mad_3y",
        "cutoff_spread_mad_used_3y",
        "mad_score_cutoff_spread",
        "positive_mad_score_cover",
        "positive_mad_score_cutoff_spread",
        "demand_level_score",
        "demand_mad_bonus",
        "rate_pressure_bonus",
        "flag_demand",
        "flag_rate_pressure",
        "stress_pattern_flag",
        "signal_ready",
        "m2_score",
        "m2_flag",
        "m2_signal",
        "m2_signal_zone",
    ]

    existing_columns = [column for column in output_columns if column in df.columns]
    df = df[existing_columns].copy()

    print("Сигналы M2 рассчитаны")

    return df


def save_results(
    repo_clean: pd.DataFrame,
    key_rate_clean: pd.DataFrame,
    m2_signals: pd.DataFrame,
) -> None:
    """Сохраняет итоговые таблицы M2."""

    signals_path = RESULTS_DIR / "m2_signals.xlsx"
    full_result_path = RESULTS_DIR / "m2_full_result.xlsx"

    m2_signals.to_excel(signals_path, index=False)

    with pd.ExcelWriter(full_result_path) as writer:
        repo_clean.to_excel(writer, sheet_name="repo_clean", index=False)
        key_rate_clean.to_excel(writer, sheet_name="key_rate_clean", index=False)
        m2_signals.to_excel(writer, sheet_name="m2_signals", index=False)

    print("=" * 80)
    print("Файлы сохранены")
    print("=" * 80)
    print(f"Итоговые сигналы M2: {signals_path}")
    print(f"Общий Excel-файл M2: {full_result_path}")


def save_chart(m2_signals: pd.DataFrame) -> None:
    """Сохраняет итоговые графики M2."""

    chart_df = m2_signals[
        (m2_signals["signal_scope"] == "main_7d_repo")
        & (m2_signals["auction_date"].notna())
    ].copy()

    chart_df = chart_df.sort_values("auction_date")

    cover_chart_path = RESULTS_DIR / "m2_cover_ratio_chart.png"
    signal_chart_path = RESULTS_DIR / "m2_signal_chart.png"

    # ------------------------------------------------------------------
    # График 1. Cover ratio
    # ------------------------------------------------------------------

    normal_points = chart_df[
        (chart_df["cover_ratio"].notna())
        & (chart_df["cover_ratio"] <= 2.0)
    ].copy()

    oversubscribed_points = chart_df[
        (chart_df["cover_ratio"].notna())
        & (chart_df["cover_ratio"] > 2.0)
    ].copy()

    fig, ax = plt.subplots(figsize=(15, 7))

    ax.scatter(
        normal_points["auction_date"],
        normal_points["cover_ratio"],
        s=28,
        alpha=0.75,
        label="Обычные аукционы",
    )

    ax.scatter(
        oversubscribed_points["auction_date"],
        oversubscribed_points["cover_ratio"],
        s=55,
        alpha=0.95,
        label="Аукционы с переспросом",
        zorder=5,
    )

    ax.axhline(
        1.0,
        linestyle=":",
        linewidth=1.2,
        label="Спрос равен размещению",
    )

    ax.axhline(
        2.0,
        linestyle="--",
        linewidth=1.2,
        label="Порог переспроса: cover ratio > 2.0",
    )

    ax.set_ylim(0, 6)

    ax.set_title("M2: cover ratio по 7-дневным аукционам РЕПО ЦБ", fontsize=15)
    ax.set_xlabel("Дата аукциона")
    ax.set_ylabel("Cover ratio")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper left", frameon=True)

    plt.savefig(cover_chart_path, dpi=220, bbox_inches="tight")
    plt.close()

    # ------------------------------------------------------------------
    # График 2. Итоговый сигнал M2
    # ------------------------------------------------------------------

    signal_df = chart_df[
        chart_df["signal_ready"] == 1
    ].copy()

    normal_signal = signal_df[
        signal_df["m2_signal"] <= 40
    ].copy()

    tension_signal = signal_df[
        (signal_df["m2_signal"] > 40)
        & (signal_df["m2_signal"] <= 70)
    ].copy()

    stress_signal = signal_df[
        signal_df["m2_signal"] > 70
    ].copy()

    fig, ax = plt.subplots(figsize=(15, 7))

    if not normal_signal.empty:
        ax.scatter(
            normal_signal["auction_date"],
            normal_signal["m2_signal"],
            s=28,
            alpha=0.55,
            label="Норма",
            zorder=3,
        )

    if not tension_signal.empty:
        ax.scatter(
            tension_signal["auction_date"],
            tension_signal["m2_signal"],
            s=70,
            alpha=0.9,
            label="Напряжение",
            zorder=5,
        )

    if not stress_signal.empty:
        ax.scatter(
            stress_signal["auction_date"],
            stress_signal["m2_signal"],
            s=80,
            alpha=0.95,
            label="Стресс",
            zorder=6,
        )

    ax.axhline(
        40,
        linestyle="--",
        linewidth=1.2,
        label="Порог напряжения",
    )

    ax.axhline(
        70,
        linestyle="--",
        linewidth=1.2,
        label="Порог стресса",
    )

    ax.set_ylim(-3, 105)

    ax.set_title("M2: итоговый сигнал по 7-дневным аукционам РЕПО ЦБ", fontsize=15)
    ax.set_xlabel("Дата аукциона")
    ax.set_ylabel("M2 signal, 0–100")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper right", frameon=True)

    plt.savefig(signal_chart_path, dpi=220, bbox_inches="tight")
    plt.close()

    print(f"График cover ratio сохранён: {cover_chart_path}")
    print(f"График M2 signal сохранён: {signal_chart_path}")


def main() -> None:
    print("=" * 80)
    print("Запускаю финальный расчет модуля M2")
    print("=" * 80)
    print(f"Корень проекта: {PROJECT_ROOT}")
    print(f"Папка результатов: {RESULTS_DIR}")

    repo_clean = clean_repo()
    key_rate_clean = clean_key_rate()

    m2_base = add_key_rate(
        repo_df=repo_clean,
        key_rate_df=key_rate_clean,
    )

    m2_signals = calculate_m2_signals(m2_base)

    save_results(
        repo_clean=repo_clean,
        key_rate_clean=key_rate_clean,
        m2_signals=m2_signals,
    )

    save_chart(m2_signals)

    print("=" * 80)
    print("Модуль M2 завершён")
    print("=" * 80)
    print(f"Строк в итоговой таблице: {len(m2_signals)}")
    print(f"Файлы лежат в: {RESULTS_DIR}")


if __name__ == "__main__":
    main()
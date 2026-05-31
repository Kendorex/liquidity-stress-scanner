from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# Корень проекта PRACTIC
PROJECT_ROOT = Path(__file__).resolve().parents[2]

# Входные папки с сырыми Excel-файлами
REQUIRED_RESERVES_DIR = PROJECT_ROOT / "data" / "m1" / "required_reserves"
RUONIA_DIR = PROJECT_ROOT / "data" / "m1" / "ruonia"

# Итоговая папка M1
RESULTS_DIR = PROJECT_ROOT / "data" / "m1" / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)


def find_latest_xlsx(folder: Path) -> Path:
    """
    Находит последний Excel-файл в папке.
    """

    files = list(folder.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(f"В папке нет Excel-файлов: {folder}")

    return max(files, key=lambda file: file.stat().st_mtime)


def clean_required_reserves() -> pd.DataFrame:
    """
    Очищает файл обязательных резервов ЦБ.

    На выходе:
    - period_start
    - period_end
    - actual_corr_accounts
    - required_reserves_avg
    - required_reserves_accounts
    - averaging_period_days
    - report_period
    - regulation_period
    - spread
    """

    print("=" * 80)
    print("Очищаю обязательные резервы")
    print("=" * 80)

    file_path = find_latest_xlsx(REQUIRED_RESERVES_DIR)
    print(f"Файл: {file_path}")

    df = pd.read_excel(
        file_path,
        sheet_name="Обязательные резервы",
        header=2,
    )

    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    df = df.rename(
        columns={
            df.columns[0]: "period_start",
            df.columns[1]: "actual_corr_accounts",
            df.columns[2]: "required_reserves_avg",
            df.columns[3]: "required_reserves_accounts",
            df.columns[4]: "banks_using_averaging",
            df.columns[5]: "active_banks",
            df.columns[6]: "period_start_reference",
            df.columns[7]: "averaging_period_days",
            df.columns[8]: "report_period",
            df.columns[9]: "regulation_period",
        }
    )

    needed_columns = [
        "period_start",
        "actual_corr_accounts",
        "required_reserves_avg",
        "required_reserves_accounts",
        "averaging_period_days",
        "report_period",
        "regulation_period",
    ]

    df = df[needed_columns].copy()

    df["period_start"] = pd.to_datetime(
        df["period_start"],
        errors="coerce",
    )

    numeric_columns = [
        "actual_corr_accounts",
        "required_reserves_avg",
        "required_reserves_accounts",
        "averaging_period_days",
    ]

    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=["period_start"])
    df = df.sort_values("period_start").reset_index(drop=True)

    # Конец периода: сначала считаем через длительность периода.
    df["period_end"] = (
        df["period_start"]
        + pd.to_timedelta(df["averaging_period_days"] - 1, unit="D")
    )

    # Если длительность периода отсутствует, берём день перед следующим периодом.
    df["period_end"] = df["period_end"].fillna(
        df["period_start"].shift(-1) - pd.Timedelta(days=1)
    )

    # Если это последняя строка и нет данных для расчёта конца периода.
    df["period_end"] = df["period_end"].fillna(df["period_start"])

    # Главный показатель M1
    df["spread"] = df["actual_corr_accounts"] - df["required_reserves_avg"]

    output_columns = [
        "period_start",
        "period_end",
        "actual_corr_accounts",
        "required_reserves_avg",
        "required_reserves_accounts",
        "averaging_period_days",
        "report_period",
        "regulation_period",
        "spread",
    ]

    df = df[output_columns].copy()

    print(f"Строк после очистки резервов: {len(df)}")

    return df


def clean_ruonia() -> pd.DataFrame:
    """
    Очищает файл RUONIA.

    На выходе:
    - date
    - ruonia
    - volume
    - deals_count
    - participants_count
    - min_rate
    - percentile_25
    - percentile_75
    - max_rate
    - status_xml
    - date_update
    - rate_range
    """

    print("=" * 80)
    print("Очищаю RUONIA")
    print("=" * 80)

    file_path = find_latest_xlsx(RUONIA_DIR)
    print(f"Файл: {file_path}")

    df = pd.read_excel(
        file_path,
        sheet_name="RC",
    )

    df = df.rename(
        columns={
            "DT": "date",
            "ruo": "ruonia",
            "vol": "volume",
            "T": "deals_count",
            "C": "participants_count",
            "MinRate": "min_rate",
            "Percentile25": "percentile_25",
            "Percentile75": "percentile_75",
            "MaxRate": "max_rate",
            "StatusXML": "status_xml",
            "DateUpdate": "date_update",
        }
    )

    needed_columns = [
        "date",
        "ruonia",
        "volume",
        "deals_count",
        "participants_count",
        "min_rate",
        "percentile_25",
        "percentile_75",
        "max_rate",
        "status_xml",
        "date_update",
    ]

    missing_columns = [column for column in needed_columns if column not in df.columns]

    if missing_columns:
        raise ValueError(f"В RUONIA не найдены колонки: {missing_columns}")

    df = df[needed_columns].copy()

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce",
    )

    df["date_update"] = pd.to_datetime(
        df["date_update"],
        errors="coerce",
    )

    numeric_columns = [
        "ruonia",
        "volume",
        "deals_count",
        "participants_count",
        "min_rate",
        "percentile_25",
        "percentile_75",
        "max_rate",
        "status_xml",
    ]

    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=["date", "ruonia"])
    df = df.sort_values("date").reset_index(drop=True)

    # Разброс ставок внутри дня
    df["rate_range"] = df["max_rate"] - df["min_rate"]

    print(f"Строк после очистки RUONIA: {len(df)}")

    return df


def calculate_rolling_mad_score(
    df: pd.DataFrame,
    date_column: str,
    value_column: str,
    window_days: int = 1095,
    min_periods: int = 12,
) -> pd.DataFrame:
    """
    Считает rolling median, rolling MAD и MAD-score за скользящее окно.

    MAD = median(|x - median(x)|)
    """

    df = df.copy()
    df = df.sort_values(date_column).reset_index(drop=True)

    temp = df[[date_column, value_column]].copy()
    temp = temp.dropna(subset=[date_column, value_column])
    temp = temp.set_index(date_column)

    rolling_window = f"{window_days}D"

    rolling_median = temp[value_column].rolling(
        window=rolling_window,
        min_periods=min_periods,
    ).median()

    rolling_mad = temp[value_column].rolling(
        window=rolling_window,
        min_periods=min_periods,
    ).apply(
        lambda values: np.median(np.abs(values - np.median(values))),
        raw=True,
    )

    result = pd.DataFrame(
        {
            date_column: rolling_median.index,
            f"{value_column}_median_3y": rolling_median.values,
            f"{value_column}_mad_3y": rolling_mad.values,
        }
    )

    df = df.merge(result, on=date_column, how="left")

    median_col = f"{value_column}_median_3y"
    mad_col = f"{value_column}_mad_3y"
    score_col = f"mad_score_{value_column}"

    df[score_col] = (df[value_column] - df[median_col]) / df[mad_col]

    # Если MAD равен 0, score некорректен.
    df.loc[df[mad_col] == 0, score_col] = np.nan

    return df


def add_ruonia_by_period(
    reserves_df: pd.DataFrame,
    ruonia_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Для каждого периода усреднения считает RUONIA внутри периода:
    - mean_ruonia
    - max_ruonia
    - last_ruonia
    - mean_rate_range
    - ruonia_observations
    """

    print("=" * 80)
    print("Накладываю RUONIA на периоды усреднения")
    print("=" * 80)

    result_rows = []

    for _, row in reserves_df.iterrows():
        period_start = row["period_start"]
        period_end = row["period_end"]

        ruonia_period = ruonia_df[
            (ruonia_df["date"] >= period_start)
            & (ruonia_df["date"] <= period_end)
        ].copy()

        row_dict = row.to_dict()

        if len(ruonia_period) == 0:
            row_dict["mean_ruonia"] = np.nan
            row_dict["max_ruonia"] = np.nan
            row_dict["last_ruonia"] = np.nan
            row_dict["mean_rate_range"] = np.nan
            row_dict["ruonia_observations"] = 0
        else:
            ruonia_period = ruonia_period.sort_values("date")

            row_dict["mean_ruonia"] = ruonia_period["ruonia"].mean()
            row_dict["max_ruonia"] = ruonia_period["ruonia"].max()
            row_dict["last_ruonia"] = ruonia_period["ruonia"].iloc[-1]
            row_dict["mean_rate_range"] = ruonia_period["rate_range"].mean()
            row_dict["ruonia_observations"] = len(ruonia_period)

        result_rows.append(row_dict)

    result_df = pd.DataFrame(result_rows)

    print(f"Периодов обработано: {len(result_df)}")

    return result_df


def calculate_m1_signals(m1_df: pd.DataFrame) -> pd.DataFrame:
    """
    Считает итоговые сигналы M1:
    - spread_deviation_from_history
    - mad_score_spread
    - mad_score_mean_ruonia
    - flag_end_of_period
    - stress_pattern_flag
    - m1_signal
    """

    print("=" * 80)
    print("Считаю сигналы M1")
    print("=" * 80)

    df = m1_df.copy()
    df = df.sort_values("period_start").reset_index(drop=True)

    historical_median_spread = df["spread"].median()
    df["spread_deviation_from_history"] = df["spread"] - historical_median_spread

    df = calculate_rolling_mad_score(
        df=df,
        date_column="period_start",
        value_column="spread",
        window_days=1095,
        min_periods=12,
    )

    df = calculate_rolling_mad_score(
        df=df,
        date_column="period_start",
        value_column="mean_ruonia",
        window_days=1095,
        min_periods=12,
    )

    # В этой версии таблица периодическая:
    # каждая строка уже является итогом периода усреднения.
    # Поэтому флаг конца периода ставим 1 на уровне period-level.
    df["flag_end_of_period"] = 1
    df["flag_end_of_period_comment"] = (
        "period_level: строка отражает итог периода усреднения; "
        "для daily-версии нужно отдельно отметить последние 3-5 дней"
    )

    df["positive_mad_score_spread"] = df["mad_score_spread"].clip(lower=0).fillna(0)
    df["positive_mad_score_ruonia"] = df["mad_score_mean_ruonia"].clip(lower=0).fillna(0)

    # Флаг готовности сигнала:
    # последняя незавершенная строка резервов может не иметь actual_corr_accounts и spread.
    df["signal_ready"] = (
        df["actual_corr_accounts"].notna()
        & df["required_reserves_avg"].notna()
        & df["spread"].notna()
        & df["mean_ruonia"].notna()
        & df["mad_score_spread"].notna()
        & df["mad_score_mean_ruonia"].notna()
    ).astype(int)

    # Паттерн из ТЗ:
    # высокий спред + высокая RUONIA + конец периода.
    df["stress_pattern_flag"] = (
        (df["positive_mad_score_spread"] >= 1.5)
        & (df["positive_mad_score_ruonia"] >= 1.0)
        & (df["flag_end_of_period"] == 1)
        & (df["signal_ready"] == 1)
    ).astype(int)

    # Базовый сигнал:
    # спред важнее, но RUONIA нужна как подтверждение межбанковского стресса.
    base_signal = (
        0.7 * df["positive_mad_score_spread"]
        + 0.3 * df["positive_mad_score_ruonia"]
    )

    df["m1_signal"] = (base_signal / 5 * 100).clip(lower=0, upper=100)

    # Если RUONIA не подтверждает стресс, не даем сигналу сразу улетать в 100.
    # Это важно по ТЗ: стресс = рост спреда + рост RUONIA.
    no_ruonia_confirmation_mask = (
        (df["positive_mad_score_ruonia"] < 1.0)
        & (df["signal_ready"] == 1)
    )

    df.loc[no_ruonia_confirmation_mask, "m1_signal"] = (
        df.loc[no_ruonia_confirmation_mask, "m1_signal"].clip(upper=70)
    )

    # Если сработал стрессовый паттерн, усиливаем сигнал.
    stress_mask = df["stress_pattern_flag"] == 1

    df.loc[stress_mask, "m1_signal"] = (
        df.loc[stress_mask, "m1_signal"] * 1.15
    ).clip(upper=100)

    # Если данных не хватает, сигнал не считаем.
    df.loc[df["signal_ready"] == 0, "m1_signal"] = 0

    # Уровень интерпретации сигнала
    df["m1_signal_zone"] = pd.cut(
        df["m1_signal"],
        bins=[-0.1, 40, 70, 100],
        labels=["норма", "напряжение", "стресс"],
    )

    output_columns = [
        "period_start",
        "period_end",
        "actual_corr_accounts",
        "required_reserves_avg",
        "required_reserves_accounts",
        "averaging_period_days",
        "report_period",
        "regulation_period",
        "spread",
        "spread_deviation_from_history",
        "spread_median_3y",
        "spread_mad_3y",
        "mad_score_spread",
        "mean_ruonia",
        "max_ruonia",
        "last_ruonia",
        "mean_rate_range",
        "ruonia_observations",
        "mean_ruonia_median_3y",
        "mean_ruonia_mad_3y",
        "mad_score_mean_ruonia",
        "flag_end_of_period",
        "flag_end_of_period_comment",
        "positive_mad_score_spread",
        "positive_mad_score_ruonia",
        "stress_pattern_flag",
        "signal_ready",
        "m1_signal",
        "m1_signal_zone",
    ]

    existing_columns = [column for column in output_columns if column in df.columns]
    df = df[existing_columns].copy()

    print("Сигналы M1 рассчитаны")

    return df


def save_results(
    required_reserves_clean: pd.DataFrame,
    ruonia_clean: pd.DataFrame,
    m1_signals: pd.DataFrame,
) -> None:
    """
    Сохраняет только финальные результаты M1.

    Отдельные файлы required_reserves_clean.xlsx и ruonia_clean.xlsx не создаются.
    Они остаются внутри кода как промежуточные таблицы.

    Сохраняем:
    - m1_signals.xlsx
    - m1_full_result.xlsx
    """

    signals_path = RESULTS_DIR / "m1_signals.xlsx"
    full_result_path = RESULTS_DIR / "m1_full_result.xlsx"

    m1_signals.to_excel(signals_path, index=False)

    with pd.ExcelWriter(full_result_path) as writer:
        required_reserves_clean.to_excel(
            writer,
            sheet_name="required_reserves_clean",
            index=False,
        )

        ruonia_clean.to_excel(
            writer,
            sheet_name="ruonia_clean",
            index=False,
        )

        m1_signals.to_excel(
            writer,
            sheet_name="m1_signals",
            index=False,
        )

    print("=" * 80)
    print("Файлы сохранены")
    print("=" * 80)
    print(f"Итоговые сигналы M1: {signals_path}")
    print(f"Общий Excel-файл M1: {full_result_path}")


def save_chart(m1_signals: pd.DataFrame) -> None:
    """
    Сохраняет цветной график 'Спред + RUONIA' для дашборда.
    """

    chart_df = m1_signals.dropna(subset=["period_start"]).copy()
    chart_df = chart_df.sort_values("period_start")

    chart_path = RESULTS_DIR / "m1_spread_ruonia_chart.png"

    spread_color = "#2563eb"   # синий
    ruonia_color = "#f97316"   # оранжевый
    stress_color = "#dc2626"   # красный
    tension_color = "#facc15"  # жёлтый

    fig, ax1 = plt.subplots(figsize=(15, 7))

    # Спред
    ax1.plot(
        chart_df["period_start"],
        chart_df["spread"],
        color=spread_color,
        linewidth=2.0,
        label="Спред усреднения, млрд руб.",
    )

    ax1.set_xlabel("Дата начала периода усреднения")
    ax1.set_ylabel("Спред усреднения, млрд руб.", color=spread_color)
    ax1.tick_params(axis="y", labelcolor=spread_color)
    ax1.grid(True, alpha=0.25)

    # RUONIA на второй оси
    ax2 = ax1.twinx()

    ax2.plot(
        chart_df["period_start"],
        chart_df["mean_ruonia"],
        color=ruonia_color,
        linewidth=2.0,
        label="Средняя RUONIA за период, %",
    )

    ax2.set_ylabel("Средняя RUONIA за период, %", color=ruonia_color)
    ax2.tick_params(axis="y", labelcolor=ruonia_color)

    # Точки напряжения и стресса по итоговому сигналу
    tension_points = chart_df[
        (chart_df["m1_signal"] > 40)
        & (chart_df["m1_signal"] <= 70)
        & (chart_df["signal_ready"] == 1)
    ]

    stress_points = chart_df[
        (chart_df["m1_signal"] > 70)
        & (chart_df["signal_ready"] == 1)
    ]

    if len(tension_points) > 0:
        ax1.scatter(
            tension_points["period_start"],
            tension_points["spread"],
            color=tension_color,
            edgecolor="black",
            linewidth=0.5,
            s=50,
            label="Напряжение по M1",
            zorder=5,
        )

    if len(stress_points) > 0:
        ax1.scatter(
            stress_points["period_start"],
            stress_points["spread"],
            color=stress_color,
            edgecolor="black",
            linewidth=0.5,
            s=65,
            label="Стресс по M1",
            zorder=6,
        )

    # Дополнительно выделяем именно паттерн из ТЗ
    pattern_points = chart_df[
        (chart_df["stress_pattern_flag"] == 1)
        & (chart_df["signal_ready"] == 1)
    ]

    if len(pattern_points) > 0:
        ax1.scatter(
            pattern_points["period_start"],
            pattern_points["spread"],
            facecolors="none",
            edgecolors=stress_color,
            linewidth=2.0,
            s=120,
            label="Паттерн: высокий спред + высокая RUONIA",
            zorder=7,
        )

    fig.suptitle("M1: спред обязательных резервов и RUONIA", fontsize=15)
    fig.tight_layout()

    lines_1, labels_1 = ax1.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()

    ax1.legend(
        lines_1 + lines_2,
        labels_1 + labels_2,
        loc="upper left",
        frameon=True,
    )

    plt.savefig(chart_path, dpi=220, bbox_inches="tight")
    plt.close()

    print(f"График сохранён: {chart_path}")


def main():
    print("=" * 80)
    print("Запускаю финальный расчет модуля M1")
    print("=" * 80)
    print(f"Корень проекта: {PROJECT_ROOT}")
    print(f"Папка результатов: {RESULTS_DIR}")

    required_reserves_clean = clean_required_reserves()
    ruonia_clean = clean_ruonia()

    m1_base = add_ruonia_by_period(
        reserves_df=required_reserves_clean,
        ruonia_df=ruonia_clean,
    )

    m1_signals = calculate_m1_signals(m1_base)

    save_results(
        required_reserves_clean=required_reserves_clean,
        ruonia_clean=ruonia_clean,
        m1_signals=m1_signals,
    )

    save_chart(m1_signals)

    print("=" * 80)
    print("Модуль M1 завершён")
    print("=" * 80)
    print(f"Периодов в итоговой таблице: {len(m1_signals)}")
    print(f"Файлы лежат в: {RESULTS_DIR}")


if __name__ == "__main__":
    main()
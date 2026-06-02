from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CALENDAR_DIR = PROJECT_ROOT / "data" / "m4" / "tax_calendar"
RESULTS_DIR = PROJECT_ROOT / "data" / "m4" / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)


def find_latest_xlsx(folder: Path) -> Path:
    files = [file for file in folder.glob("*.xlsx") if not file.name.startswith("~$")]

    if not files:
        raise FileNotFoundError(f"В папке нет Excel-файлов: {folder}")

    return max(files, key=lambda file: file.stat().st_mtime)


def clean_tax_calendar() -> pd.DataFrame:
    print("Очищаю календарь налоговых дат")

    file_path = find_latest_xlsx(CALENDAR_DIR)
    print(f"Файл: {file_path}")

    df = pd.read_excel(file_path)

    if "date" not in df.columns:
        raise ValueError("В календаре M4 нет колонки date")

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()
    df = df.sort_values("date").reset_index(drop=True)

    flag_columns = [
        "is_weekend",
        "tax_notification_flag",
        "tax_payment_flag",
        "ndfl_second_payment_flag",
        "end_of_month_flag",
        "end_of_quarter_flag",
        "tax_payment_window_flag",
        "tax_notification_window_flag",
        "tax_week_flag",
    ]

    for column in flag_columns:
        if column not in df.columns:
            df[column] = 0
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0).astype(int)

    numeric_columns = [
        "direct_event_weight",
        "days_to_nearest_tax_payment",
        "days_to_nearest_tax_notification",
    ]

    for column in numeric_columns:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")

    df["tax_event_name"] = df.get("tax_event_name", "").fillna("").astype(str)
    df["tax_event_type"] = df.get("tax_event_type", "").fillna("").astype(str)

    print(f"Строк после очистки: {len(df)}")

    return df


def calculate_m4_signals(calendar_df: pd.DataFrame) -> pd.DataFrame:
    print("Считаю сезонный фактор M4")

    df = calendar_df.copy()
    df["tax_pressure_score"] = 0.0
    df["tax_pressure_score"] += df["tax_payment_flag"] * 45
    df["tax_pressure_score"] += df["ndfl_second_payment_flag"] * 25
    df["tax_pressure_score"] += df["tax_notification_flag"] * 12
    payment_window_bonus = np.where(
        df["tax_payment_window_flag"] == 1,
        18 - np.minimum(np.abs(df["days_to_nearest_tax_payment"].fillna(99)), 5) * 3,
        0,
    )
    df["tax_pressure_score"] += np.maximum(payment_window_bonus, 0)

    notification_window_bonus = np.where(
        df["tax_notification_window_flag"] == 1,
        8
        - np.minimum(np.abs(df["days_to_nearest_tax_notification"].fillna(99)), 3) * 2,
        0,
    )
    df["tax_pressure_score"] += np.maximum(notification_window_bonus, 0)
    tax_week_bonus = np.where(
        (df["tax_week_flag"] == 1)
        & (df["tax_payment_window_flag"] == 0)
        & (df["tax_notification_window_flag"] == 0),
        4,
        0,
    )
    df["tax_pressure_score"] += tax_week_bonus
    df["tax_pressure_score"] += df["end_of_month_flag"] * 8
    df["tax_pressure_score"] += df["end_of_quarter_flag"] * 15
    df["tax_pressure_score"] = df["tax_pressure_score"].clip(lower=0, upper=100)
    df["seasonal_factor"] = 1.0 + 0.25 * df["tax_pressure_score"] / 100
    df["seasonal_factor"] = df["seasonal_factor"].clip(lower=1.0, upper=1.25)
    df["m4_seasonal_factor"] = df["seasonal_factor"]
    df["m4_score"] = df["tax_pressure_score"]
    df["m4_flag"] = df["tax_week_flag"].astype(int)
    df["m4_signal"] = df["m4_seasonal_factor"]

    df["m4_signal_zone"] = pd.cut(
        df["seasonal_factor"],
        bins=[0.99, 1.05, 1.12, 1.20, 1.25],
        labels=[
            "обычный день",
            "слабый фактор",
            "налоговое давление",
            "сильное давление",
        ],
        include_lowest=True,
    )

    output_columns = [
        "date",
        "year",
        "month",
        "quarter",
        "weekday",
        "is_weekend",
        "tax_event_name",
        "tax_event_type",
        "tax_notification_flag",
        "tax_payment_flag",
        "ndfl_second_payment_flag",
        "end_of_month_flag",
        "end_of_quarter_flag",
        "tax_payment_window_flag",
        "tax_notification_window_flag",
        "tax_week_flag",
        "days_to_nearest_tax_payment",
        "days_to_nearest_tax_notification",
        "tax_pressure_score",
        "seasonal_factor",
        "m4_seasonal_factor",
        "m4_score",
        "m4_flag",
        "m4_signal",
        "m4_signal_zone",
    ]

    existing_columns = [column for column in output_columns if column in df.columns]
    result = df[existing_columns].copy()

    print("Сезонный фактор M4 рассчитан")

    return result


def save_results(calendar_clean: pd.DataFrame, m4_signals: pd.DataFrame) -> None:
    signals_path = RESULTS_DIR / "m4_signals.xlsx"
    full_result_path = RESULTS_DIR / "m4_full_result.xlsx"

    m4_signals.to_excel(signals_path, index=False)

    with pd.ExcelWriter(full_result_path) as writer:
        calendar_clean.to_excel(writer, sheet_name="tax_calendar_clean", index=False)
        m4_signals.to_excel(writer, sheet_name="m4_signals", index=False)

    print("Файлы сохранены")
    print(f"Итоговые сигналы M4: {signals_path}")
    print(f"Общий Excel-файл M4: {full_result_path}")


def save_chart(m4_signals: pd.DataFrame) -> None:
    chart_df = m4_signals.copy()
    chart_df["date"] = pd.to_datetime(chart_df["date"], errors="coerce")
    chart_df = chart_df.dropna(subset=["date"]).sort_values("date")

    full_chart_path = RESULTS_DIR / "m4_seasonal_factor_chart_full.png"
    recent_chart_path = RESULTS_DIR / "m4_seasonal_factor_chart_recent.png"
    fig, ax = plt.subplots(figsize=(16, 7))

    ax.plot(
        chart_df["date"],
        chart_df["seasonal_factor"],
        linewidth=1.6,
        label="Seasonal factor M4",
    )

    strong_points = chart_df[chart_df["seasonal_factor"] >= 1.20].copy()
    if not strong_points.empty:
        ax.scatter(
            strong_points["date"],
            strong_points["seasonal_factor"],
            s=45,
            label="Сильное налоговое давление",
            zorder=5,
        )

    ax.axhline(1.00, linestyle=":", linewidth=1.2, label="Обычный день")
    ax.axhline(1.12, linestyle="--", linewidth=1.2, label="Налоговое давление")
    ax.axhline(1.20, linestyle="--", linewidth=1.2, label="Сильное давление")

    ax.set_title(
        "M4: календарный сезонный фактор налоговых периодов (полный период)",
        fontsize=16,
    )
    ax.set_xlabel("Дата")
    ax.set_ylabel("Seasonal factor")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="lower right", frameon=True)

    plt.savefig(full_chart_path, dpi=220, bbox_inches="tight")
    plt.close()
    recent_start_date = chart_df["date"].max() - pd.DateOffset(years=4)
    recent_df = chart_df[chart_df["date"] >= recent_start_date].copy()

    fig, ax = plt.subplots(figsize=(16, 7))

    ax.plot(
        recent_df["date"],
        recent_df["seasonal_factor"],
        linewidth=1.6,
        label="Seasonal factor M4",
    )

    strong_points_recent = recent_df[recent_df["seasonal_factor"] >= 1.20].copy()
    if not strong_points_recent.empty:
        ax.scatter(
            strong_points_recent["date"],
            strong_points_recent["seasonal_factor"],
            s=45,
            label="Сильное налоговое давление",
            zorder=5,
        )

    ax.axhline(1.00, linestyle=":", linewidth=1.2, label="Обычный день")
    ax.axhline(1.12, linestyle="--", linewidth=1.2, label="Налоговое давление")
    ax.axhline(1.20, linestyle="--", linewidth=1.2, label="Сильное давление")

    ax.set_title(
        "M4: календарный сезонный фактор налоговых периодов (последние 4 года)",
        fontsize=16,
    )
    ax.set_xlabel("Дата")
    ax.set_ylabel("Seasonal factor")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="lower right", frameon=True)

    plt.savefig(recent_chart_path, dpi=220, bbox_inches="tight")
    plt.close()

    print(f"Полный график сохранён: {full_chart_path}")
    print(f"Приближенный график сохранён: {recent_chart_path}")


def print_summary(m4_signals: pd.DataFrame) -> None:
    print("Сводка M4")
    print(
        f"Период: {m4_signals['date'].min().date()} — {m4_signals['date'].max().date()}"
    )
    print(f"Строк в итоговой таблице: {len(m4_signals)}")
    print("Распределение зон:")
    print(m4_signals["m4_signal_zone"].value_counts(dropna=False).to_string())


def main() -> None:
    print("=" * 80)
    print("Запускаю расчет модуля M4")
    print("=" * 80)
    print(f"Корень проекта: {PROJECT_ROOT}")
    print(f"Папка результатов: {RESULTS_DIR}")

    calendar_clean = clean_tax_calendar()
    m4_signals = calculate_m4_signals(calendar_clean)

    save_results(calendar_clean, m4_signals)
    save_chart(m4_signals)
    print_summary(m4_signals)

    print("=" * 80)
    print("Модуль M4 завершён")
    print("=" * 80)


if __name__ == "__main__":
    main()

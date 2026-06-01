
from pathlib import Path
import warnings
import time

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# LSI — Liquidity Stress Index
#
# Верхний слой над M1–M5:
# 1) читает готовые сигналы модулей;
# 2) приводит их к дневной шкале;
# 3) нормализует сигналы в sub-index 0–100;
# 4) считает единый LSI 0–100;
# 5) показывает вклад каждого модуля.
#
# Важно:
# - M4 не является отдельным стресс-модулем, а работает как сезонный множитель.
# - Вклады contribution_m1...contribution_m5 после clipping нормируются так,
#   чтобы их сумма была равна итоговому LSI.
# ============================================================


# ------------------------------------------------------------
# Пути
# ------------------------------------------------------------

M1_FILE_CANDIDATES = [
    Path("data/m1/results/m1_signals.xlsx"),
    Path("data/m1/result/m1_signals.xlsx"),
]

M2_FILE_CANDIDATES = [
    Path("data/m2/results/m2_signals.xlsx"),
    Path("data/m2/result/m2_signals.xlsx"),
]

M3_FILE_CANDIDATES = [
    Path("data/m3/result/ofz_auctions_m3_signals.xlsx"),
    Path("data/m3/results/ofz_auctions_m3_signals.xlsx"),
    Path("data/m3/result/m3_cover_ratio_daily.xlsx"),
    Path("data/m3/results/m3_cover_ratio_daily.xlsx"),
]

M4_FILE_CANDIDATES = [
    Path("data/m4/results/m4_signals.xlsx"),
    Path("data/m4/result/m4_signals.xlsx"),
]

M5_FILE_CANDIDATES = [
    Path("data/m5/result/m5_treasury_signals.xlsx"),
    Path("data/m5/results/m5_treasury_signals.xlsx"),
]

RESULT_DIR = Path("data/lsi/results")
OUTPUT_FILE = RESULT_DIR / "lsi_signals.xlsx"
LSI_CHART_FILE = RESULT_DIR / "lsi_chart.png"
CONTRIBUTIONS_CHART_FILE = RESULT_DIR / "lsi_contributions_chart.png"


# ------------------------------------------------------------
# Методология
# ------------------------------------------------------------

BASE_WEIGHTS = {
    "m1": 0.25,
    "m2": 0.25,
    "m3": 0.20,
    "m5": 0.30,
}

TAX_WEEK_WEIGHT_DISCOUNT = 0.85
MAX_SEASONAL_MULTIPLIER = 1.25
AGGREGATION_SCALE = 1.8
MAD_CAP = 4.0
M3_CARRY_DAYS = 7

# Обрезаем будущие / неполные будущие даты.
CUT_FUTURE_DATES = True


# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------

def find_existing_path(candidates: list[Path], module_name: str) -> Path | None:
    for path in candidates:
        if path.exists():
            return path

    print(f"{module_name}: файл не найден.")
    print("Проверенные пути:")
    for path in candidates:
        print(f"  - {path}")

    return None


def read_excel_safe(path: Path | None, module_name: str, sheet_name=0) -> pd.DataFrame:
    if path is None:
        return pd.DataFrame()

    print(f"{module_name}: читаю {path}")

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as error:
        warnings.warn(f"{module_name}: не удалось прочитать {path}, лист {sheet_name}: {error}")
        return pd.DataFrame()


def read_excel_try_sheets(path: Path | None, module_name: str, preferred_sheets=None) -> pd.DataFrame:
    if path is None:
        return pd.DataFrame()

    preferred_sheets = preferred_sheets or [0]

    print(f"{module_name}: читаю {path}")

    try:
        excel = pd.ExcelFile(path)
    except Exception as error:
        warnings.warn(f"{module_name}: не удалось открыть {path}: {error}")
        return pd.DataFrame()

    sheet_order = []
    for sheet in preferred_sheets:
        if isinstance(sheet, int):
            if 0 <= sheet < len(excel.sheet_names):
                sheet_order.append(excel.sheet_names[sheet])
        elif sheet in excel.sheet_names:
            sheet_order.append(sheet)

    for sheet in excel.sheet_names:
        if sheet not in sheet_order:
            sheet_order.append(sheet)

    for sheet in sheet_order:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            if not df.empty:
                df.attrs["source_sheet"] = sheet
                return df
        except Exception:
            continue

    return pd.DataFrame()


def to_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.normalize()


def num(series, default=0.0) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(series, errors="coerce").fillna(default)


def first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def positive_mad_to_score(series, cap: float = MAD_CAP) -> pd.Series:
    values = pd.to_numeric(series, errors="coerce").fillna(0.0)
    values = values.clip(lower=0.0, upper=cap)
    return values / cap * 100.0


def negative_mad_to_score(series, cap: float = MAD_CAP) -> pd.Series:
    values = pd.to_numeric(series, errors="coerce").fillna(0.0)
    values = (-values).clip(lower=0.0, upper=cap)
    return values / cap * 100.0


def clip_0_100(series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0).clip(0.0, 100.0)


def status_from_lsi(value: float) -> str:
    if value < 40:
        return "ЗЕЛЁНЫЙ"
    if value < 70:
        return "ЖЁЛТЫЙ"
    return "КРАСНЫЙ"


def normalize_weights(weights: dict[str, float]) -> dict[str, float]:
    total = sum(weights.values())
    if total == 0:
        return weights
    return {key: value / total for key, value in weights.items()}


def build_daily_timeline(frames: list[pd.DataFrame]) -> pd.DataFrame:
    dates = []

    for frame in frames:
        if frame is None or frame.empty or "date" not in frame.columns:
            continue

        clean_dates = pd.to_datetime(frame["date"], errors="coerce").dropna()
        if not clean_dates.empty:
            dates.append((clean_dates.min().normalize(), clean_dates.max().normalize()))

    if not dates:
        raise RuntimeError("Не удалось построить дневную шкалу: нет дат во входных модулях.")

    start = min(item[0] for item in dates)
    end = max(item[1] for item in dates)

    if CUT_FUTURE_DATES:
        today = pd.Timestamp.today().normalize()
        end = min(end, today)

    return pd.DataFrame({"date": pd.date_range(start, end, freq="D")})


def safe_bool_flag(series) -> pd.Series:
    return (pd.to_numeric(series, errors="coerce").fillna(0.0) > 0).astype(int)


# ------------------------------------------------------------
# Загрузка M1–M5
# ------------------------------------------------------------

def prepare_m1() -> pd.DataFrame:
    path = find_existing_path(M1_FILE_CANDIDATES, "M1")
    df = read_excel_try_sheets(path, "M1")
    if df.empty:
        return pd.DataFrame(columns=["date", "m1_score", "m1_flag", "m1_ready"])

    date_col = first_existing_column(
        df,
        [
            "period_end",
            "period_start",
            "averaging_period_end",
            "averaging_period_start",
            "end_date",
            "start_date",
            "date",
        ],
    )

    if date_col is None:
        raise RuntimeError(f"M1: не нашёл колонку даты. Колонки: {list(df.columns)}")

    df["date"] = to_date(df[date_col])

    spread_score = positive_mad_to_score(df["mad_score_spread"] if "mad_score_spread" in df.columns else None)
    ruonia_score = positive_mad_to_score(df["mad_score_mean_ruonia"] if "mad_score_mean_ruonia" in df.columns else None)

    if "stress_pattern_flag" in df.columns:
        flag = safe_bool_flag(df["stress_pattern_flag"])
    elif "flag_stress" in df.columns:
        flag = safe_bool_flag(df["flag_stress"])
    else:
        flag = pd.Series(0, index=df.index)

    df["m1_score"] = clip_0_100(0.45 * spread_score + 0.35 * ruonia_score + 0.20 * flag * 100.0)
    df["m1_flag"] = flag.astype(int)

    if "signal_ready" in df.columns:
        df["m1_ready"] = safe_bool_flag(df["signal_ready"])
    else:
        df["m1_ready"] = 1

    result = df[["date", "m1_score", "m1_flag", "m1_ready"]].dropna(subset=["date"])
    return result


def prepare_m2() -> pd.DataFrame:
    path = find_existing_path(M2_FILE_CANDIDATES, "M2")
    df = read_excel_try_sheets(path, "M2")
    if df.empty:
        return pd.DataFrame(columns=["date", "m2_score", "m2_flag", "m2_ready"])

    date_col = first_existing_column(df, ["auction_date", "date"])
    if date_col is None:
        raise RuntimeError(f"M2: не нашёл колонку даты. Колонки: {list(df.columns)}")

    df["date"] = to_date(df[date_col])

    cover_score = positive_mad_to_score(df["mad_score_cover_ratio"] if "mad_score_cover_ratio" in df.columns else None)
    rate_score = positive_mad_to_score(df["mad_score_cutoff_spread"] if "mad_score_cutoff_spread" in df.columns else None)

    demand_flag = safe_bool_flag(df["flag_demand"] if "flag_demand" in df.columns else pd.Series(0, index=df.index))
    rate_flag = safe_bool_flag(df["flag_rate_pressure"] if "flag_rate_pressure" in df.columns else pd.Series(0, index=df.index))
    flag = ((demand_flag + rate_flag) > 0).astype(int)

    df["m2_score"] = clip_0_100(0.50 * cover_score + 0.30 * rate_score + 0.20 * flag * 100.0)
    df["m2_flag"] = flag

    if "signal_ready" in df.columns:
        df["m2_ready"] = safe_bool_flag(df["signal_ready"])
    else:
        df["m2_ready"] = 1

    daily = (
        df[["date", "m2_score", "m2_flag", "m2_ready"]]
        .dropna(subset=["date"])
        .groupby("date", as_index=False)
        .agg(
            m2_score=("m2_score", "max"),
            m2_flag=("m2_flag", "max"),
            m2_ready=("m2_ready", "max"),
        )
    )

    return daily


def prepare_m3() -> pd.DataFrame:
    path = find_existing_path(M3_FILE_CANDIDATES, "M3")
    df = read_excel_try_sheets(path, "M3")
    if df.empty:
        return pd.DataFrame(columns=["date", "m3_score", "m3_flag", "m3_ready"])

    date_col = first_existing_column(df, ["auction_date", "date"])
    if date_col is None:
        raise RuntimeError(f"M3: не нашёл колонку даты. Колонки: {list(df.columns)}")

    df["date"] = to_date(df[date_col])

    if "m3_score" in df.columns:
        df["m3_score"] = clip_0_100(df["m3_score"])
    else:
        # В актуальном M3 mad_score_cover уже развернут: положительное значение = стресс.
        # Старую ошибку с negative_mad_to_score здесь не используем.
        low_cover_score = positive_mad_to_score(df["mad_score_cover"] if "mad_score_cover" in df.columns else None)
        yield_score = positive_mad_to_score(df["mad_score_yield_spread"] if "mad_score_yield_spread" in df.columns else None)
        if "flag_nedospros" in df.columns:
            nedospros_flag = safe_bool_flag(df["flag_nedospros"])
        elif "cover_ratio" in df.columns:
            nedospros_flag = (pd.to_numeric(df["cover_ratio"], errors="coerce") < 1.2).fillna(False).astype(int)
        else:
            nedospros_flag = pd.Series(0, index=df.index)
        df["m3_score"] = clip_0_100(0.50 * low_cover_score + 0.20 * yield_score + 0.30 * nedospros_flag * 100.0)

    if "m3_flag" in df.columns:
        df["m3_flag"] = safe_bool_flag(df["m3_flag"])
    elif "flag_nedospros" in df.columns:
        df["m3_flag"] = safe_bool_flag(df["flag_nedospros"])
    elif "cover_ratio" in df.columns:
        df["m3_flag"] = (pd.to_numeric(df["cover_ratio"], errors="coerce") < 1.2).fillna(False).astype(int)
    else:
        df["m3_flag"] = 0

    df["m3_ready"] = 1

    daily = (
        df[["date", "m3_score", "m3_flag", "m3_ready"]]
        .dropna(subset=["date"])
        .groupby("date", as_index=False)
        .agg(
            m3_score=("m3_score", "max"),
            m3_flag=("m3_flag", "max"),
            m3_ready=("m3_ready", "max"),
        )
    )

    return daily


def prepare_m4() -> pd.DataFrame:
    path = find_existing_path(M4_FILE_CANDIDATES, "M4")
    df = read_excel_try_sheets(path, "M4")
    if df.empty:
        return pd.DataFrame(columns=["date", "tax_week_flag", "seasonal_factor", "tax_pressure_score"])

    date_col = first_existing_column(df, ["date", "calendar_date"])
    if date_col is None:
        raise RuntimeError(f"M4: не нашёл колонку даты. Колонки: {list(df.columns)}")

    df["date"] = to_date(df[date_col])
    df["tax_week_flag"] = safe_bool_flag(df["tax_week_flag"] if "tax_week_flag" in df.columns else pd.Series(0, index=df.index))
    df["seasonal_factor"] = pd.to_numeric(
        df["seasonal_factor"] if "seasonal_factor" in df.columns else 1.0,
        errors="coerce",
    ).fillna(1.0)

    if "tax_pressure_score" in df.columns:
        df["tax_pressure_score"] = pd.to_numeric(df["tax_pressure_score"], errors="coerce").fillna(0.0)
    else:
        df["tax_pressure_score"] = df["tax_week_flag"] * 100.0

    return df[["date", "tax_week_flag", "seasonal_factor", "tax_pressure_score"]].dropna(subset=["date"])


def prepare_m5() -> pd.DataFrame:
    path = find_existing_path(M5_FILE_CANDIDATES, "M5")
    df = read_excel_try_sheets(path, "M5", preferred_sheets=["m5_monthly_signals", 0])
    if df.empty:
        return pd.DataFrame(columns=["date", "m5_score", "m5_flag", "m5_ready"])

    date_col = first_existing_column(df, ["date", "period_end", "month"])
    if date_col is None:
        raise RuntimeError(f"M5: не нашёл колонку даты. Колонки: {list(df.columns)}")

    df["date"] = to_date(df[date_col])

    if "m5_score" in df.columns:
        df["m5_score"] = clip_0_100(df["m5_score"])
    else:
        cbr_score = positive_mad_to_score(df["mad_score_cbr"] if "mad_score_cbr" in df.columns else None)
        if "mad_score_roskazna" in df.columns:
            roskazna_mad = df["mad_score_roskazna"]
        elif "mad_score_roskazna_weekly_max" in df.columns:
            roskazna_mad = df["mad_score_roskazna_weekly_max"]
        elif "mad_score_roskazna_monthly" in df.columns:
            roskazna_mad = df["mad_score_roskazna_monthly"]
        else:
            roskazna_mad = None
        roskazna_score = positive_mad_to_score(roskazna_mad)
        flag = safe_bool_flag(df["flag_budget_drain"] if "flag_budget_drain" in df.columns else pd.Series(0, index=df.index))
        df["m5_score"] = clip_0_100(0.70 * cbr_score + 0.15 * roskazna_score + 0.15 * flag * 100.0)

    if "m5_flag" in df.columns:
        df["m5_flag"] = safe_bool_flag(df["m5_flag"])
    else:
        df["m5_flag"] = safe_bool_flag(df["flag_budget_drain"] if "flag_budget_drain" in df.columns else pd.Series(0, index=df.index))

    signal_cols = [col for col in ["m5_score", "mad_score_cbr", "mad_score_roskazna", "flag_budget_drain"] if col in df.columns]
    if signal_cols:
        df["m5_ready"] = df[signal_cols].notna().any(axis=1).astype(int)
    else:
        df["m5_ready"] = 0

    return df[["date", "m5_score", "m5_flag", "m5_ready"]].dropna(subset=["date"])


def prepare_ground_truth() -> pd.DataFrame:
    path = find_existing_path(M5_FILE_CANDIDATES, "M5 ground truth")
    df = read_excel_try_sheets(path, "M5 ground truth", preferred_sheets=["m5_monthly_signals", 0])
    if df.empty or "ground_truth_liquidity_balance" not in df.columns:
        return pd.DataFrame(columns=["date", "ground_truth_liquidity_balance", "ground_truth_stress_flag"])

    date_col = first_existing_column(df, ["date", "period_end", "month"])
    if date_col is None:
        return pd.DataFrame(columns=["date", "ground_truth_liquidity_balance", "ground_truth_stress_flag"])

    df["date"] = to_date(df[date_col])
    df["ground_truth_liquidity_balance"] = pd.to_numeric(df["ground_truth_liquidity_balance"], errors="coerce")

    valid = df["ground_truth_liquidity_balance"].dropna()
    if valid.empty:
        df["ground_truth_stress_flag"] = 0
    else:
        # Для проверки: стресс = нижний квартиль исторических значений или баланс < 0.
        threshold = min(0.0, valid.quantile(0.25))
        df["ground_truth_stress_flag"] = (df["ground_truth_liquidity_balance"] <= threshold).astype(int)

    return df[["date", "ground_truth_liquidity_balance", "ground_truth_stress_flag"]].dropna(subset=["date"])


# ------------------------------------------------------------
# Объединение
# ------------------------------------------------------------

def merge_modules() -> pd.DataFrame:
    m1 = prepare_m1()
    m2 = prepare_m2()
    m3 = prepare_m3()
    m4 = prepare_m4()
    m5 = prepare_m5()
    gt = prepare_ground_truth()

    print()
    print("Готовлю единую дневную шкалу...")
    timeline = build_daily_timeline([m1, m2, m3, m4, m5, gt])
    result = timeline.copy()

    # M1 и M5 — низкая частота. Берём последнее известное значение.
    for frame in [m1, m5, gt]:
        if not frame.empty:
            result = pd.merge_asof(
                result.sort_values("date"),
                frame.sort_values("date"),
                on="date",
                direction="backward",
            )

    # M2 — в дни без аукциона давление = 0.
    if not m2.empty:
        result = result.merge(m2, on="date", how="left")

    # M3 — держим сигнал M3 несколько дней после аукциона.
    if not m3.empty:
        m3_daily = timeline.merge(m3, on="date", how="left")
        for col in ["m3_score", "m3_flag", "m3_ready"]:
            if col not in m3_daily.columns:
                m3_daily[col] = 0

        m3_daily["m3_score"] = pd.to_numeric(m3_daily["m3_score"], errors="coerce").fillna(0.0)
        m3_daily["m3_flag"] = pd.to_numeric(m3_daily["m3_flag"], errors="coerce").fillna(0).astype(int)
        m3_daily["m3_ready"] = pd.to_numeric(m3_daily["m3_ready"], errors="coerce").fillna(0).astype(int)

        m3_daily["m3_score"] = m3_daily["m3_score"].rolling(M3_CARRY_DAYS, min_periods=1).max()
        m3_daily["m3_flag"] = m3_daily["m3_flag"].rolling(M3_CARRY_DAYS, min_periods=1).max().astype(int)
        m3_daily["m3_ready"] = m3_daily["m3_ready"].rolling(M3_CARRY_DAYS, min_periods=1).max().astype(int)

        result = result.merge(
            m3_daily[["date", "m3_score", "m3_flag", "m3_ready"]],
            on="date",
            how="left",
        )

    # M4 — календарь строго по дате.
    if not m4.empty:
        result = result.merge(m4, on="date", how="left")

    fill_defaults = {
        "m1_score": 0.0,
        "m2_score": 0.0,
        "m3_score": 0.0,
        "m5_score": 0.0,
        "m1_flag": 0,
        "m2_flag": 0,
        "m3_flag": 0,
        "m5_flag": 0,
        "m1_ready": 0,
        "m2_ready": 0,
        "m3_ready": 0,
        "m5_ready": 0,
        "tax_week_flag": 0,
        "seasonal_factor": 1.0,
        "tax_pressure_score": 0.0,
        "ground_truth_liquidity_balance": np.nan,
        "ground_truth_stress_flag": 0,
    }

    for col, default in fill_defaults.items():
        if col not in result.columns:
            result[col] = default
        result[col] = result[col].fillna(default)

    return result


# ------------------------------------------------------------
# LSI
# ------------------------------------------------------------

def calculate_lsi(df: pd.DataFrame, weights: dict[str, float] | None = None) -> pd.DataFrame:
    weights = normalize_weights(weights or BASE_WEIGHTS)
    result = df.copy()

    for module, weight in weights.items():
        result[f"weight_{module}"] = weight

    # В налоговые недели уменьшаем M1/M2/M5, чтобы не было двойного счёта календарного фактора.
    tax_mask = result["tax_week_flag"].astype(int) == 1
    for module in ["m1", "m2", "m5"]:
        if f"weight_{module}" in result.columns:
            result.loc[tax_mask, f"weight_{module}"] = (
                result.loc[tax_mask, f"weight_{module}"] * TAX_WEEK_WEIGHT_DISCOUNT
            )

    weight_cols = [f"weight_{module}" for module in weights]
    weight_sum = result[weight_cols].sum(axis=1).replace(0, np.nan)

    for module in weights:
        result[f"weight_{module}"] = result[f"weight_{module}"] / weight_sum

    # Базовые вклады до сезонного множителя.
    for module in weights:
        result[f"base_contribution_{module}"] = (
            result[f"{module}_score"] * result[f"weight_{module}"] * AGGREGATION_SCALE
        )

    result["base_lsi_before_seasonal"] = sum(result[f"base_contribution_{module}"] for module in weights)

    result["seasonal_multiplier"] = (
        pd.to_numeric(result["seasonal_factor"], errors="coerce")
        .fillna(1.0)
        .clip(lower=1.0, upper=MAX_SEASONAL_MULTIPLIER)
    )

    result["seasonal_factor"] = result["seasonal_multiplier"]

    result["raw_lsi_before_clip"] = result["base_lsi_before_seasonal"] * result["seasonal_multiplier"]
    result["lsi"] = result["raw_lsi_before_clip"].clip(0.0, 100.0)

    # Вклад M4 — только дополнительное усиление от сезонного множителя.
    result["base_contribution_m4"] = (
        result["base_lsi_before_seasonal"] * (result["seasonal_multiplier"] - 1.0)
    ).clip(lower=0.0)

    # До clipping сумма этих вкладов = raw_lsi_before_clip.
    preclip_cols = []
    for module in weights:
        result[f"preclip_contribution_{module}"] = result[f"base_contribution_{module}"]
        preclip_cols.append(f"preclip_contribution_{module}")

    result["preclip_contribution_m4"] = result["base_contribution_m4"]
    preclip_cols.append("preclip_contribution_m4")

    preclip_sum = result[preclip_cols].sum(axis=1).replace(0, np.nan)
    clip_factor = (result["lsi"] / preclip_sum).replace([np.inf, -np.inf], np.nan).fillna(0.0)

    # Финальные вклады. Их сумма равна LSI.
    for module in weights:
        result[f"contribution_{module}"] = result[f"preclip_contribution_{module}"] * clip_factor

    result["contribution_m4"] = result["preclip_contribution_m4"] * clip_factor

    # Контрольная колонка для проверки интерпретируемости.
    contribution_cols = [
        "contribution_m1",
        "contribution_m2",
        "contribution_m3",
        "contribution_m4",
        "contribution_m5",
    ]
    for col in contribution_cols:
        if col not in result.columns:
            result[col] = 0.0

    result["contribution_sum_check"] = result[contribution_cols].sum(axis=1)
    result["contribution_gap"] = result["lsi"] - result["contribution_sum_check"]

    result["status"] = result["lsi"].apply(status_from_lsi)
    result["event_stress_flag"] = result["date"].apply(is_known_stress_event).astype(int)
    result["active_flags"] = result.apply(build_active_flags_text, axis=1)
    result["top_driver"] = result.apply(get_top_driver, axis=1)

    return result


def is_known_stress_event(date_value) -> int:
    """Ручной event-flag для проверки на известных стрессовых окнах."""
    date = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(date):
        return 0

    episodes = [
        (pd.Timestamp("2014-12-01"), pd.Timestamp("2014-12-31")),
        (pd.Timestamp("2022-02-01"), pd.Timestamp("2022-03-31")),
        (pd.Timestamp("2025-01-01"), pd.Timestamp("2025-01-31")),
    ]

    return int(any(start <= date <= end for start, end in episodes))


def build_active_flags_text(row: pd.Series) -> str:
    flags = []

    if int(row.get("m1_flag", 0)) == 1:
        flags.append("M1: резервы/RUONIA")

    if int(row.get("m2_flag", 0)) == 1:
        flags.append("M2: спрос на repo")

    if int(row.get("m3_flag", 0)) == 1:
        flags.append("M3: недоспрос ОФЗ")

    if int(row.get("tax_week_flag", 0)) == 1:
        flags.append("M4: налоговая неделя")

    if int(row.get("m5_flag", 0)) == 1:
        flags.append("M5: бюджетный отток")

    return "; ".join(flags) if flags else "нет активных флагов"


def get_top_driver(row: pd.Series) -> str:
    values = {
        "M1": row.get("contribution_m1", 0.0),
        "M2": row.get("contribution_m2", 0.0),
        "M3": row.get("contribution_m3", 0.0),
        "M4": row.get("contribution_m4", 0.0),
        "M5": row.get("contribution_m5", 0.0),
    }

    return max(values, key=values.get)


# ------------------------------------------------------------
# Backtest / sensitivity / quality
# ------------------------------------------------------------

def build_backtest(lsi_df: pd.DataFrame) -> pd.DataFrame:
    episodes = [
        ("Декабрь 2014", "2014-12-01", "2014-12-31"),
        ("Февраль–март 2022", "2022-02-01", "2022-03-31"),
        ("Отложенная проверка: январь 2025", "2025-01-01", "2025-01-31"),
    ]

    rows = []

    for name, start, end in episodes:
        mask = (lsi_df["date"] >= pd.Timestamp(start)) & (lsi_df["date"] <= pd.Timestamp(end))
        period = lsi_df.loc[mask].copy()

        if period.empty:
            rows.append(
                {
                    "episode": name,
                    "start": start,
                    "end": end,
                    "comment": "нет данных",
                }
            )
            continue

        idxmax = period["lsi"].idxmax()

        rows.append(
            {
                "episode": name,
                "start": start,
                "end": end,
                "days": len(period),
                "mean_lsi": period["lsi"].mean(),
                "max_lsi": period["lsi"].max(),
                "red_days": int((period["status"] == "КРАСНЫЙ").sum()),
                "yellow_or_red_days": int((period["lsi"] >= 40).sum()),
                "max_lsi_date": period.loc[idxmax, "date"],
                "top_driver_at_max": period.loc[idxmax, "top_driver"],
                "ground_truth_stress_days": int(
                    period.get("ground_truth_stress_flag", pd.Series(0, index=period.index)).sum()
                ),
                "event_stress_days": int(
                    period.get("event_stress_flag", pd.Series(0, index=period.index)).sum()
                ),
                "comment": "ok",
            }
        )

    return pd.DataFrame(rows)


def build_sensitivity(base_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    base_lsi = calculate_lsi(base_df, BASE_WEIGHTS)
    base_mean = base_lsi["lsi"].mean()
    base_max = base_lsi["lsi"].max()

    for module in BASE_WEIGHTS:
        for shock in [-0.2, 0.2]:
            shocked_weights = BASE_WEIGHTS.copy()
            shocked_weights[module] = shocked_weights[module] * (1.0 + shock)
            shocked_weights = normalize_weights(shocked_weights)

            shocked = calculate_lsi(base_df, shocked_weights)

            rows.append(
                {
                    "module_shocked": module.upper(),
                    "weight_change": f"{shock:+.0%}",
                    "new_weight_m1": shocked_weights["m1"],
                    "new_weight_m2": shocked_weights["m2"],
                    "new_weight_m3": shocked_weights["m3"],
                    "new_weight_m5": shocked_weights["m5"],
                    "mean_lsi": shocked["lsi"].mean(),
                    "mean_lsi_change": shocked["lsi"].mean() - base_mean,
                    "max_lsi": shocked["lsi"].max(),
                    "max_lsi_change": shocked["lsi"].max() - base_max,
                }
            )

    return pd.DataFrame(rows)


def build_quality_report(lsi_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for module in ["m1", "m2", "m3", "m5"]:
        score_col = f"{module}_score"
        ready_col = f"{module}_ready"
        flag_col = f"{module}_flag"

        rows.append(
            {
                "module": module.upper(),
                "nonzero_score_days": int((lsi_df[score_col] > 0).sum()) if score_col in lsi_df.columns else 0,
                "flag_days": int((lsi_df[flag_col] > 0).sum()) if flag_col in lsi_df.columns else 0,
                "mean_score": lsi_df[score_col].mean() if score_col in lsi_df.columns else 0,
                "max_score": lsi_df[score_col].max() if score_col in lsi_df.columns else 0,
                "ready_days": int((lsi_df[ready_col] > 0).sum()) if ready_col in lsi_df.columns else 0,
            }
        )

    rows.append(
        {
            "module": "M4",
            "nonzero_score_days": int((lsi_df["tax_pressure_score"] > 0).sum()),
            "flag_days": int((lsi_df["tax_week_flag"] > 0).sum()),
            "mean_score": lsi_df["tax_pressure_score"].mean(),
            "max_score": lsi_df["tax_pressure_score"].max(),
            "ready_days": int(lsi_df["seasonal_factor"].notna().sum()),
        }
    )

    rows.append(
        {
            "module": "CONTRIBUTIONS_CHECK",
            "nonzero_score_days": "",
            "flag_days": "",
            "mean_score": "",
            "max_score": float(lsi_df["contribution_gap"].abs().max()),
            "ready_days": "max_abs_gap_should_be_0",
        }
    )

    return pd.DataFrame(rows)


# ------------------------------------------------------------
# Графики и сохранение
# ------------------------------------------------------------

def plot_lsi(lsi_df: pd.DataFrame) -> None:
    plot_df = lsi_df.dropna(subset=["date"]).copy()

    if plot_df.empty:
        return

    plt.figure(figsize=(16, 7))
    plt.plot(plot_df["date"], plot_df["lsi"], linewidth=1.8, label="LSI")
    plt.axhline(40, linestyle="--", linewidth=1.1, label="Порог жёлтой зоны = 40")
    plt.axhline(70, linestyle="--", linewidth=1.1, label="Порог красной зоны = 70")

    stress_episodes = [
        ("2014-12-01", "2014-12-31", "декабрь 2014"),
        ("2022-02-01", "2022-03-31", "февраль–март 2022"),
    ]

    for start, end, label in stress_episodes:
        plt.axvspan(pd.Timestamp(start), pd.Timestamp(end), alpha=0.12, label=label)

    plt.title("Liquidity Stress Index, 0–100")
    plt.xlabel("Дата")
    plt.ylabel("LSI")
    plt.ylim(0, 105)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    plt.savefig(LSI_CHART_FILE, dpi=220)
    plt.close()


def plot_contributions(lsi_df: pd.DataFrame) -> None:
    # Рисуем только последние 365 дней, чтобы график быстро строился.
    plot_df = lsi_df.dropna(subset=["date"]).tail(365).copy()

    if plot_df.empty:
        return

    for col in ["contribution_m1", "contribution_m2", "contribution_m3", "contribution_m4", "contribution_m5"]:
        if col not in plot_df.columns:
            plot_df[col] = 0.0

    plt.figure(figsize=(16, 7))
    plt.stackplot(
        plot_df["date"],
        plot_df["contribution_m1"],
        plot_df["contribution_m2"],
        plot_df["contribution_m3"],
        plot_df["contribution_m4"],
        plot_df["contribution_m5"],
        labels=["M1 резервы", "M2 repo", "M3 ОФЗ", "M4 сезонность", "M5 казначейство"],
        alpha=0.85,
    )
    plt.plot(plot_df["date"], plot_df["lsi"], linewidth=1.5, label="LSI")

    plt.title("Вклад модулей в LSI — последние 365 дней")
    plt.xlabel("Дата")
    plt.ylabel("Пункты LSI")
    plt.ylim(0, 105)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    plt.savefig(CONTRIBUTIONS_CHART_FILE, dpi=220)
    plt.close()


def save_outputs(
    lsi_df: pd.DataFrame,
    backtest: pd.DataFrame,
    sensitivity: pd.DataFrame,
    quality: pd.DataFrame,
) -> Path:
    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    config = pd.DataFrame(
        [
            ["method", "Интерпретируемая взвешенная сумма + сезонный множитель M4"],
            [
                "base_formula",
                "raw = (w1*M1 + w2*M2 + w3*M3 + w5*M5) * scale * seasonal_factor; LSI = clip(raw, 0, 100)",
            ],
            ["weights", str(BASE_WEIGHTS)],
            [
                "tax_week_adjustment",
                f"при Tax_Week_Flag=1 веса M1/M2/M5 умножаются на {TAX_WEEK_WEIGHT_DISCOUNT}",
            ],
            ["seasonal_multiplier_cap", MAX_SEASONAL_MULTIPLIER],
            ["aggregation_scale", AGGREGATION_SCALE],
            ["mad_cap", MAD_CAP],
            ["m3_carry_days", M3_CARRY_DAYS],
            ["status_green", "0 <= LSI < 40"],
            ["status_yellow", "40 <= LSI < 70"],
            ["status_red", "70 <= LSI <= 100"],
            ["contribution_rule", "после clipping вклады нормируются так, чтобы сумма вкладов равнялась LSI"],
        ],
        columns=["parameter", "value"],
    )

    selected_cols = [
        "date",
        "lsi",
        "raw_lsi_before_clip",
        "status",
        "top_driver",
        "active_flags",
        "m1_score",
        "m2_score",
        "m3_score",
        "tax_pressure_score",
        "seasonal_factor",
        "m5_score",
        "contribution_m1",
        "contribution_m2",
        "contribution_m3",
        "contribution_m4",
        "contribution_m5",
        "contribution_sum_check",
        "contribution_gap",
        "ground_truth_liquidity_balance",
        "ground_truth_stress_flag",
        "event_stress_flag",
    ]
    selected_cols = [col for col in selected_cols if col in lsi_df.columns]

    output_path = OUTPUT_FILE

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            lsi_df[selected_cols].to_excel(writer, sheet_name="lsi_daily", index=False)
            lsi_df.to_excel(writer, sheet_name="lsi_full", index=False)
            backtest.to_excel(writer, sheet_name="backtest", index=False)
            sensitivity.to_excel(writer, sheet_name="sensitivity", index=False)
            quality.to_excel(writer, sheet_name="data_quality", index=False)
            config.to_excel(writer, sheet_name="method_config", index=False)

        return output_path

    except PermissionError:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_path = RESULT_DIR / f"lsi_signals_{timestamp}.xlsx"

        print()
        print(f"Файл {OUTPUT_FILE} открыт или заблокирован.")
        print(f"Сохраняю результат в новый файл: {output_path}")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            lsi_df[selected_cols].to_excel(writer, sheet_name="lsi_daily", index=False)
            lsi_df.to_excel(writer, sheet_name="lsi_full", index=False)
            backtest.to_excel(writer, sheet_name="backtest", index=False)
            sensitivity.to_excel(writer, sheet_name="sensitivity", index=False)
            quality.to_excel(writer, sheet_name="data_quality", index=False)
            config.to_excel(writer, sheet_name="method_config", index=False)

        return output_path


def main() -> None:
    print("=" * 70)
    print("LSI — Liquidity Stress Index")
    print("Агрегационный слой M1–M5")
    print("=" * 70)

    base_df = merge_modules()

    print("Считаю LSI...")
    lsi_df = calculate_lsi(base_df, BASE_WEIGHTS)

    backtest = build_backtest(lsi_df)
    sensitivity = build_sensitivity(base_df)
    quality = build_quality_report(lsi_df)

    print("Сохраняю Excel...")
    output_path = save_outputs(lsi_df, backtest, sensitivity, quality)

    print("Строю графики...")
    plot_lsi(lsi_df)
    plot_contributions(lsi_df)

    print()
    print("Готово.")
    print(f"Итоговый Excel: {output_path}")
    print(f"График LSI: {LSI_CHART_FILE}")
    print(f"График вкладов: {CONTRIBUTIONS_CHART_FILE}")

    print()
    print("Последние 10 значений LSI:")
    print(lsi_df[["date", "lsi", "status", "top_driver", "active_flags"]].tail(10).to_string(index=False))

    print()
    print("Backtest:")
    print(backtest.to_string(index=False))

    print()
    print("Проверка заполненности данных:")
    print(quality.to_string(index=False))


if __name__ == "__main__":
    main()

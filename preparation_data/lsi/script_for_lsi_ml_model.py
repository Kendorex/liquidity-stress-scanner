from pathlib import Path
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler
    from sklearn.pipeline import Pipeline
    from sklearn.metrics import roc_auc_score, average_precision_score, classification_report
except ImportError as error:
    raise RuntimeError(
        "Не установлен scikit-learn. Установи его командой: pip install scikit-learn"
    ) from error


# ============================================================
# LSI ML — дополнительная ML-калибровка Liquidity Stress Index
#
# Это НЕ замена основному интерпретируемому LSI.
# Основной LSI считается явной формулой с весами.
# Этот скрипт добавляет отдельный ML-слой поверх уже собранных признаков.
#
# Идея:
# 1. Берём готовый результат основного LSI:
#    data/lsi/results/lsi_signals.xlsx, лист lsi_full.
# 2. Используем признаки M1–M5, флаги, M4 seasonal/tax.
# 3. Целевая переменная — ground_truth_stress_flag из таблицы ЦБ
#    «Ликвидность банковского сектора».
# 4. Обучаем интерпретируемую LogisticRegression.
# 5. Получаем ML_LSI = вероятность стресса * 100.
# 6. Для каждого дня считаем вклад признаков в ML-сигнал.
#
# На выходе:
# data/lsi/results/lsi_ml_signals.xlsx
# data/lsi/results/lsi_ml_chart.png
# data/lsi/results/lsi_ml_contributions_chart.png
# ============================================================


INPUT_FILE = Path("data/lsi/results/lsi_signals.xlsx")
OUTPUT_DIR = Path("data/lsi/results")
OUTPUT_FILE = OUTPUT_DIR / "lsi_ml_signals.xlsx"
ML_CHART_FILE = OUTPUT_DIR / "lsi_ml_chart.png"
ML_CONTRIBUTIONS_CHART_FILE = OUTPUT_DIR / "lsi_ml_contributions_chart.png"

# Отложенная выборка: всё, начиная с этой даты, не используется при обучении.
# Это нужно, чтобы не обучать модель на тех же данных, на которых проверяем.
HOLDOUT_START = "2024-01-01"

# Порог вероятности стресса для статусов.
# ML_LSI = probability * 100.
GREEN_THRESHOLD = 40
RED_THRESHOLD = 70

# Основные признаки. Скрипт сам оставит только те, которые реально есть в Excel.
FEATURE_COLUMNS = [
    # нормализованные scores модулей
    "m1_score",
    "m2_score",
    "m3_score",
    "m5_score",
    # M4 календарный контекст
    "tax_pressure_score",
    "seasonal_factor",
    "tax_week_flag",
    # флаги модулей
    "m1_flag",
    "m2_flag",
    "m3_flag",
    "m5_flag",
    # уже рассчитанный экспертный LSI как агрегированный признак
    "lsi",
]

TARGET_COLUMN = "ground_truth_stress_flag"
DATE_COLUMN = "date"


# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------


def status_from_ml_lsi(value: float) -> str:
    if value < GREEN_THRESHOLD:
        return "ЗЕЛЁНЫЙ"
    if value < RED_THRESHOLD:
        return "ЖЁЛТЫЙ"
    return "КРАСНЫЙ"


def read_lsi_full() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Не найден файл {INPUT_FILE}. Сначала запусти основной LSI-скрипт."
        )

    print(f"Читаю основной LSI: {INPUT_FILE}")

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name="lsi_full")
    except Exception:
        df = pd.read_excel(INPUT_FILE, sheet_name="lsi_daily")

    if DATE_COLUMN not in df.columns:
        raise RuntimeError("В lsi_signals.xlsx нет колонки date.")

    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors="coerce").dt.normalize()
    df = df.dropna(subset=[DATE_COLUMN]).copy()
    df = df.sort_values(DATE_COLUMN).reset_index(drop=True)

    if TARGET_COLUMN not in df.columns:
        raise RuntimeError(
            f"В lsi_signals.xlsx нет целевой колонки {TARGET_COLUMN}. "
            "Проверь, что основной LSI подтягивает ground truth из M5."
        )

    return df


def prepare_ml_dataset(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    existing_features = [col for col in FEATURE_COLUMNS if col in df.columns]

    if not existing_features:
        raise RuntimeError("Не нашёл ни одного признака для ML-модели.")

    data = df.copy()

    for col in existing_features:
        data[col] = pd.to_numeric(data[col], errors="coerce")

    data[TARGET_COLUMN] = pd.to_numeric(data[TARGET_COLUMN], errors="coerce")

    # Для признаков пропуски заменяем нулём: нет сигнала = нет давления.
    data[existing_features] = data[existing_features].fillna(0.0)

    # Для таргета пропуски удаляем: модель нельзя учить без ground truth.
    data = data.dropna(subset=[TARGET_COLUMN]).copy()
    data[TARGET_COLUMN] = data[TARGET_COLUMN].astype(int)

    # На всякий случай убираем строки, где target не 0/1.
    data = data[data[TARGET_COLUMN].isin([0, 1])].copy()

    if data.empty:
        raise RuntimeError("После очистки не осталось строк с ground truth.")

    target_counts = data[TARGET_COLUMN].value_counts().to_dict()
    print(f"Ground truth классы: {target_counts}")

    if data[TARGET_COLUMN].nunique() < 2:
        raise RuntimeError(
            "В ground truth есть только один класс. LogisticRegression нельзя обучить. "
            "Нужно, чтобы ground_truth_stress_flag содержал и 0, и 1."
        )

    return data, existing_features


def split_train_test(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    holdout_start = pd.Timestamp(HOLDOUT_START)

    train = data[data[DATE_COLUMN] < holdout_start].copy()
    test = data[data[DATE_COLUMN] >= holdout_start].copy()

    if train.empty or test.empty:
        warnings.warn(
            "Не удалось сделать нормальный time-based split. "
            "Использую последние 25% наблюдений как test."
        )
        split_idx = int(len(data) * 0.75)
        train = data.iloc[:split_idx].copy()
        test = data.iloc[split_idx:].copy()

    if train[TARGET_COLUMN].nunique() < 2:
        warnings.warn(
            "В train только один класс. Расширяю train за счёт всей доступной истории. "
            "Это хуже для честной проверки, но позволяет обучить модель."
        )
        train = data.copy()

    return train, test


def train_model(train: pd.DataFrame, feature_cols: list[str]) -> Pipeline:
    x_train = train[feature_cols]
    y_train = train[TARGET_COLUMN]

    model = Pipeline(
        steps=[
            ("scaler", StandardScaler()),
            (
                "logit",
                LogisticRegression(
                    class_weight="balanced",
                    max_iter=2000,
                    C=0.8,
                    random_state=42,
                ),
            ),
        ]
    )

    model.fit(x_train, y_train)
    return model


def predict_all(model: Pipeline, df: pd.DataFrame, feature_cols: list[str]) -> pd.DataFrame:
    result = df.copy()

    for col in feature_cols:
        result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0.0)

    probabilities = model.predict_proba(result[feature_cols])[:, 1]

    result["ml_stress_probability"] = probabilities
    result["ml_lsi"] = (probabilities * 100.0).clip(0.0, 100.0)
    result["ml_status"] = result["ml_lsi"].apply(status_from_ml_lsi)
    result["ml_stress_flag_40"] = (result["ml_lsi"] >= GREEN_THRESHOLD).astype(int)
    result["ml_stress_flag_70"] = (result["ml_lsi"] >= RED_THRESHOLD).astype(int)

    return result


def evaluate_model(model: Pipeline, data: pd.DataFrame, feature_cols: list[str], sample_name: str) -> dict:
    if data.empty:
        return {"sample": sample_name, "comment": "empty"}

    y_true = data[TARGET_COLUMN].astype(int)
    y_prob = model.predict_proba(data[feature_cols])[:, 1]
    y_pred = (y_prob >= 0.5).astype(int)

    row = {
        "sample": sample_name,
        "rows": len(data),
        "stress_days": int(y_true.sum()),
        "mean_probability": float(np.mean(y_prob)),
        "max_probability": float(np.max(y_prob)),
    }

    if y_true.nunique() == 2:
        row["roc_auc"] = float(roc_auc_score(y_true, y_prob))
        row["average_precision"] = float(average_precision_score(y_true, y_prob))
    else:
        row["roc_auc"] = np.nan
        row["average_precision"] = np.nan

    report = classification_report(y_true, y_pred, output_dict=True, zero_division=0)
    row["precision_stress"] = report.get("1", {}).get("precision", np.nan)
    row["recall_stress"] = report.get("1", {}).get("recall", np.nan)
    row["f1_stress"] = report.get("1", {}).get("f1-score", np.nan)
    row["comment"] = "ok"

    return row


def build_coefficients_table(model: Pipeline, feature_cols: list[str]) -> pd.DataFrame:
    scaler = model.named_steps["scaler"]
    logit = model.named_steps["logit"]

    coef_scaled = logit.coef_[0]

    # Коэффициенты в исходной шкале признаков.
    # Это удобно для интерпретации: вклад = coef_original * feature_value.
    coef_original = coef_scaled / scaler.scale_
    intercept_original = logit.intercept_[0] - np.sum(coef_scaled * scaler.mean_ / scaler.scale_)

    table = pd.DataFrame(
        {
            "feature": feature_cols,
            "coef_scaled": coef_scaled,
            "coef_original_scale": coef_original,
            "abs_coef_scaled": np.abs(coef_scaled),
            "direction": np.where(coef_scaled >= 0, "усиливает стресс", "снижает стресс"),
        }
    )

    table = table.sort_values("abs_coef_scaled", ascending=False).reset_index(drop=True)

    intercept_row = pd.DataFrame(
        [
            {
                "feature": "intercept",
                "coef_scaled": logit.intercept_[0],
                "coef_original_scale": intercept_original,
                "abs_coef_scaled": abs(logit.intercept_[0]),
                "direction": "базовый уровень",
            }
        ]
    )

    return pd.concat([table, intercept_row], ignore_index=True)


def add_feature_contributions(
    result: pd.DataFrame,
    model: Pipeline,
    feature_cols: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Для LogisticRegression естественный вклад считается в log-odds:
    logit(p) = intercept + beta_1*x_1 + ... + beta_n*x_n.

    Для дашборда дополнительно переводим положительные вклады в шкалу ML_LSI:
    contribution_feature_i суммируются примерно до ML_LSI.
    Это не SHAP, но интерпретируемая декомпозиция линейной модели.
    """
    out = result.copy()

    scaler = model.named_steps["scaler"]
    logit = model.named_steps["logit"]

    x = out[feature_cols].copy()
    for col in feature_cols:
        x[col] = pd.to_numeric(x[col], errors="coerce").fillna(0.0)

    x_scaled = scaler.transform(x)
    coef = logit.coef_[0]

    raw_contrib = x_scaled * coef

    contrib_cols = []
    for idx, feature in enumerate(feature_cols):
        col_name = f"ml_contribution_{feature}"
        out[col_name] = raw_contrib[:, idx]
        contrib_cols.append(col_name)

    out["ml_intercept_contribution"] = logit.intercept_[0]
    out["ml_logit_raw"] = out["ml_intercept_contribution"] + out[contrib_cols].sum(axis=1)

    # Положительные вклады распределяем в пункты ML_LSI.
    # Если все вклады <= 0, показываем нулевой вклад признаков.
    positive = out[contrib_cols].clip(lower=0.0)
    positive_sum = positive.sum(axis=1).replace(0, np.nan)

    display_rows = []
    for feature in feature_cols:
        raw_col = f"ml_contribution_{feature}"
        points_col = f"ml_points_{feature}"
        out[points_col] = (out[raw_col].clip(lower=0.0) / positive_sum * out["ml_lsi"]).fillna(0.0)
        display_rows.append(points_col)

    out["ml_points_sum"] = out[display_rows].sum(axis=1)
    out["ml_top_driver"] = out[display_rows].idxmax(axis=1).str.replace("ml_points_", "", regex=False)

    feature_map = pd.DataFrame(
        {
            "feature": feature_cols,
            "raw_contribution_column": [f"ml_contribution_{f}" for f in feature_cols],
            "points_contribution_column": [f"ml_points_{f}" for f in feature_cols],
        }
    )

    return out, feature_map


def build_backtest(df: pd.DataFrame) -> pd.DataFrame:
    episodes = [
        ("Декабрь 2014", "2014-12-01", "2014-12-31"),
        ("Февраль–март 2022", "2022-02-01", "2022-03-31"),
        ("Август 2023", "2023-08-01", "2023-08-31"),
        ("Отложенная проверка: январь 2025", "2025-01-01", "2025-01-31"),
    ]

    rows = []

    for name, start, end in episodes:
        mask = (df[DATE_COLUMN] >= pd.Timestamp(start)) & (df[DATE_COLUMN] <= pd.Timestamp(end))
        period = df.loc[mask].copy()

        if period.empty:
            rows.append({"episode": name, "start": start, "end": end, "comment": "нет данных"})
            continue

        max_idx = period["ml_lsi"].idxmax()

        rows.append(
            {
                "episode": name,
                "start": start,
                "end": end,
                "days": len(period),
                "mean_ml_lsi": period["ml_lsi"].mean(),
                "max_ml_lsi": period["ml_lsi"].max(),
                "max_ml_lsi_date": period.loc[max_idx, DATE_COLUMN],
                "top_driver_at_max": period.loc[max_idx, "ml_top_driver"],
                "yellow_or_red_days": int((period["ml_lsi"] >= GREEN_THRESHOLD).sum()),
                "red_days": int((period["ml_lsi"] >= RED_THRESHOLD).sum()),
                "ground_truth_stress_days": int(period.get(TARGET_COLUMN, pd.Series(0, index=period.index)).sum()),
                "comment": "ok",
            }
        )

    return pd.DataFrame(rows)


def build_sensitivity(model: Pipeline, df: pd.DataFrame, feature_cols: list[str]) -> pd.DataFrame:
    """
    Sensitivity для ML делаем не по весам, а по ключевым признакам:
    смотрим, как меняется средний ML_LSI при шоке признака ±20%.
    Это близко к требованию ТЗ про устойчивость к весам, но адаптировано для ML.
    """
    rows = []
    base_prob = model.predict_proba(df[feature_cols])[:, 1]
    base_mean = float(np.mean(base_prob) * 100.0)
    base_max = float(np.max(base_prob) * 100.0)

    for feature in feature_cols:
        for shock in [-0.2, 0.2]:
            shocked = df[feature_cols].copy()
            shocked[feature] = shocked[feature] * (1.0 + shock)
            prob = model.predict_proba(shocked)[:, 1]

            rows.append(
                {
                    "feature_shocked": feature,
                    "shock": f"{shock:+.0%}",
                    "mean_ml_lsi": float(np.mean(prob) * 100.0),
                    "mean_ml_lsi_change": float(np.mean(prob) * 100.0 - base_mean),
                    "max_ml_lsi": float(np.max(prob) * 100.0),
                    "max_ml_lsi_change": float(np.max(prob) * 100.0 - base_max),
                }
            )

    return pd.DataFrame(rows)


def plot_ml_lsi(df: pd.DataFrame) -> None:
    plot_df = df.dropna(subset=[DATE_COLUMN]).copy()
    if plot_df.empty:
        return

    plt.figure(figsize=(16, 7))
    plt.plot(plot_df[DATE_COLUMN], plot_df["ml_lsi"], linewidth=1.6, label="ML_LSI")

    if "lsi" in plot_df.columns:
        plt.plot(plot_df[DATE_COLUMN], plot_df["lsi"], linewidth=1.1, alpha=0.65, label="Базовый LSI")

    plt.axhline(GREEN_THRESHOLD, linestyle="--", linewidth=1.0, label="Порог жёлтой зоны = 40")
    plt.axhline(RED_THRESHOLD, linestyle="--", linewidth=1.0, label="Порог красной зоны = 70")

    stress_episodes = [
        ("2014-12-01", "2014-12-31", "декабрь 2014"),
        ("2022-02-01", "2022-03-31", "февраль–март 2022"),
        ("2023-08-01", "2023-08-31", "август 2023"),
    ]

    for start, end, label in stress_episodes:
        plt.axvspan(pd.Timestamp(start), pd.Timestamp(end), alpha=0.10, label=label)

    plt.title("ML Liquidity Stress Index, 0–100")
    plt.xlabel("Дата")
    plt.ylabel("ML_LSI")
    plt.ylim(0, 105)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(ML_CHART_FILE, dpi=220)
    plt.close()


def plot_ml_contributions(df: pd.DataFrame, feature_cols: list[str]) -> None:
    point_cols = [f"ml_points_{feature}" for feature in feature_cols if f"ml_points_{feature}" in df.columns]

    if not point_cols:
        return

    plot_df = df.tail(365).copy()
    if plot_df.empty:
        return

    plt.figure(figsize=(16, 7))
    plt.stackplot(
        plot_df[DATE_COLUMN],
        *[plot_df[col] for col in point_cols],
        labels=[col.replace("ml_points_", "") for col in point_cols],
        alpha=0.85,
    )
    plt.plot(plot_df[DATE_COLUMN], plot_df["ml_lsi"], linewidth=1.4, label="ML_LSI")
    plt.title("Вклад признаков в ML_LSI — последние 365 дней")
    plt.xlabel("Дата")
    plt.ylabel("Пункты ML_LSI")
    plt.ylim(0, 105)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left", ncol=2)
    plt.tight_layout()
    plt.savefig(ML_CONTRIBUTIONS_CHART_FILE, dpi=220)
    plt.close()


def save_outputs(
    result: pd.DataFrame,
    train: pd.DataFrame,
    test: pd.DataFrame,
    metrics: pd.DataFrame,
    coefficients: pd.DataFrame,
    backtest: pd.DataFrame,
    sensitivity: pd.DataFrame,
    feature_map: pd.DataFrame,
    feature_cols: list[str],
) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    method_config = pd.DataFrame(
        [
            ["model", "LogisticRegression(class_weight='balanced')"],
            ["target", TARGET_COLUMN],
            ["features", ", ".join(feature_cols)],
            ["holdout_start", HOLDOUT_START],
            ["ml_lsi_formula", "ML_LSI = P(ground_truth_stress_flag=1) * 100"],
            ["interpretability", "коэффициенты логистической регрессии + вклад признаков в log-odds"],
            ["status_green", "0 <= ML_LSI < 40"],
            ["status_yellow", "40 <= ML_LSI < 70"],
            ["status_red", "70 <= ML_LSI <= 100"],
            ["important_note", "ML-слой является дополнительной калибровкой, основной LSI остаётся интерпретируемой weighted-sum моделью"],
        ],
        columns=["parameter", "value"],
    )

    selected_cols = [
        DATE_COLUMN,
        "ml_lsi",
        "ml_status",
        "ml_stress_probability",
        "ml_top_driver",
        "ml_stress_flag_40",
        "ml_stress_flag_70",
        "lsi",
        "status",
        TARGET_COLUMN,
    ]
    selected_cols += [col for col in feature_cols if col not in selected_cols]
    selected_cols = [col for col in selected_cols if col in result.columns]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result[selected_cols].to_excel(writer, sheet_name="ml_lsi_daily", index=False)
        result.to_excel(writer, sheet_name="ml_lsi_full", index=False)
        coefficients.to_excel(writer, sheet_name="model_coefficients", index=False)
        metrics.to_excel(writer, sheet_name="model_metrics", index=False)
        backtest.to_excel(writer, sheet_name="backtest", index=False)
        sensitivity.to_excel(writer, sheet_name="sensitivity", index=False)
        feature_map.to_excel(writer, sheet_name="feature_contribution_map", index=False)
        train[[DATE_COLUMN, TARGET_COLUMN] + feature_cols].to_excel(writer, sheet_name="train_sample", index=False)
        test[[DATE_COLUMN, TARGET_COLUMN] + feature_cols].to_excel(writer, sheet_name="test_sample", index=False)
        method_config.to_excel(writer, sheet_name="method_config", index=False)


def main() -> None:
    print("=" * 70)
    print("LSI ML — Logistic Regression calibration")
    print("Дополнительный ML-слой поверх M1–M5")
    print("=" * 70)

    raw = read_lsi_full()
    data, feature_cols = prepare_ml_dataset(raw)
    train, test = split_train_test(data)

    print()
    print(f"Признаки модели: {feature_cols}")
    print(f"Train: {len(train)} строк, стресс-дней: {int(train[TARGET_COLUMN].sum())}")
    print(f"Test: {len(test)} строк, стресс-дней: {int(test[TARGET_COLUMN].sum())}")

    model = train_model(train, feature_cols)

    # Предсказания считаем по всей исходной таблице LSI, чтобы получить ежедневный ML_LSI.
    result = raw.copy()
    for col in feature_cols:
        result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0.0)

    result = predict_all(model, result, feature_cols)
    result, feature_map = add_feature_contributions(result, model, feature_cols)

    metrics = pd.DataFrame(
        [
            evaluate_model(model, train, feature_cols, "train"),
            evaluate_model(model, test, feature_cols, "test_holdout"),
            evaluate_model(model, data, feature_cols, "all_labeled_data"),
        ]
    )

    coefficients = build_coefficients_table(model, feature_cols)
    backtest = build_backtest(result)
    sensitivity = build_sensitivity(model, data, feature_cols)

    save_outputs(
        result=result,
        train=train,
        test=test,
        metrics=metrics,
        coefficients=coefficients,
        backtest=backtest,
        sensitivity=sensitivity,
        feature_map=feature_map,
        feature_cols=feature_cols,
    )

    plot_ml_lsi(result)
    plot_ml_contributions(result, feature_cols)

    print()
    print("Готово.")
    print(f"Excel: {OUTPUT_FILE}")
    print(f"График ML_LSI: {ML_CHART_FILE}")
    print(f"График вкладов ML: {ML_CONTRIBUTIONS_CHART_FILE}")

    print()
    print("Метрики модели:")
    print(metrics.to_string(index=False))

    print()
    print("Коэффициенты модели:")
    print(coefficients.to_string(index=False))

    print()
    print("Backtest:")
    print(backtest.to_string(index=False))

    print()
    print("Последние 10 значений ML_LSI:")
    print(
        result[[DATE_COLUMN, "ml_lsi", "ml_status", "ml_top_driver", "ml_stress_probability"]]
        .tail(10)
        .to_string(index=False)
    )


if __name__ == "__main__":
    main()

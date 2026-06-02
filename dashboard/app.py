from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Liquidity Stress Scanner",
    page_icon="📊",
    layout="wide",
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]

LSI_PATH = PROJECT_ROOT / "data" / "lsi" / "results" / "lsi_signals.xlsx"
BACKTEST_PATH = PROJECT_ROOT / "data" / "lsi" / "results" / "lsi_backtest.xlsx"
QUALITY_PATH = PROJECT_ROOT / "data" / "lsi" / "results" / "lsi_quality_report.xlsx"

MODULE_PATHS = {
    "M1 резервы": PROJECT_ROOT / "data" / "m1" / "results" / "m1_signals.xlsx",
    "M2 репо": PROJECT_ROOT / "data" / "m2" / "results" / "m2_signals.xlsx",
    "M3 ОФЗ": PROJECT_ROOT / "data" / "m3" / "result" / "ofz_auctions_m3_signals.xlsx",
    "M4 сезонность": PROJECT_ROOT / "data" / "m4" / "results" / "m4_signals.xlsx",
    "M5 казначейство": PROJECT_ROOT
    / "data"
    / "m5"
    / "result"
    / "m5_treasury_signals.xlsx",
}

@st.cache_data
def load_lsi_data(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл: {path}")

    df = pd.read_excel(path)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)

    numeric_cols = [
        "lsi",
        "m1_score",
        "m2_score",
        "m3_score",
        "m5_score",
        "tax_pressure_score",
        "seasonal_factor",
        "contribution_m1",
        "contribution_m2",
        "contribution_m3",
        "contribution_m4",
        "contribution_m5",
        "ground_truth_liquidity_balance",
        "ground_truth_stress_flag",
        "event_stress_flag",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@st.cache_data
def load_excel_if_exists(path: Path) -> pd.DataFrame | None:
    if not path.exists():
        return None

    df = pd.read_excel(path)

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.sort_values("date")

    return df


def build_backtest_from_lsi(lsi_df: pd.DataFrame) -> pd.DataFrame:
    episodes = [
        {
            "Эпизод": "Декабрь 2014",
            "Начало": "2014-12-01",
            "Конец": "2014-12-31",
            "Комментарий": "Реакция может быть слабее из-за более ограниченного покрытия ранних данных.",
        },
        {
            "Эпизод": "Февраль–март 2022",
            "Начало": "2022-02-01",
            "Конец": "2022-03-31",
            "Комментарий": "Основной стрессовый период для проверки реакции индекса.",
        },
        {
            "Эпизод": "Отложенная проверка: январь 2025",
            "Начало": "2025-01-01",
            "Конец": "2025-01-31",
            "Комментарий": "Период используется как дополнительная проверка вне основных исторических шоков.",
        },
    ]

    rows = []

    for episode in episodes:
        start = pd.Timestamp(episode["Начало"])
        end = pd.Timestamp(episode["Конец"])

        part = lsi_df[(lsi_df["date"] >= start) & (lsi_df["date"] <= end)].copy()

        if part.empty:
            rows.append(
                {
                    "Эпизод": episode["Эпизод"],
                    "Начало": episode["Начало"],
                    "Конец": episode["Конец"],
                    "Дней в данных": 0,
                    "Средний LSI": "-",
                    "Максимальный LSI": "-",
                    "Жёлтых/красных дней": "-",
                    "Красных дней": "-",
                    "Главный драйвер": "-",
                    "Комментарий": "Нет данных за период.",
                }
            )
            continue

        driver = (
            part["top_driver"].dropna().astype(str).value_counts().index[0]
            if "top_driver" in part.columns and part["top_driver"].notna().any()
            else "-"
        )

        rows.append(
            {
                "Эпизод": episode["Эпизод"],
                "Начало": episode["Начало"],
                "Конец": episode["Конец"],
                "Дней в данных": len(part),
                "Средний LSI": round(part["lsi"].mean(), 2),
                "Максимальный LSI": round(part["lsi"].max(), 2),
                "Жёлтых/красных дней": int((part["lsi"] >= 40).sum()),
                "Красных дней": int((part["lsi"] >= 70).sum()),
                "Главный драйвер": driver,
                "Комментарий": episode["Комментарий"],
            }
        )

    return pd.DataFrame(rows)


def build_quality_report_from_files() -> pd.DataFrame:
    rows = []

    for module_name, path in MODULE_PATHS.items():
        if not path.exists():
            rows.append(
                {
                    "Модуль": module_name,
                    "Строк": 0,
                    "Период с": "-",
                    "Период по": "-",
                    "Средний score": np.nan,
                    "Максимальный score": np.nan,
                    "Кол-во флагов": np.nan,
                    "Комментарий": "Файл не найден",
                }
            )
            continue

        try:
            df = pd.read_excel(path)

            possible_date_cols = [
                "date",
                "auction_date",
                "Дата",
                "period_date",
                "month",
                "report_date",
            ]

            date_col = next(
                (col for col in possible_date_cols if col in df.columns), None
            )

            preferred_score_cols = {
                "M1 резервы": "m1_score",
                "M2 репо": "m2_score",
                "M3 ОФЗ": "m3_score",
                "M4 сезонность": "m4_seasonal_factor",
                "M5 казначейство": "m5_score",
            }

            preferred_flag_cols = {
                "M1 резервы": "m1_flag",
                "M2 репо": "m2_flag",
                "M3 ОФЗ": "m3_flag",
                "M4 сезонность": "m4_flag",
                "M5 казначейство": "m5_flag",
            }

            score_col = preferred_score_cols.get(module_name)
            flag_col = preferred_flag_cols.get(module_name)

            if score_col not in df.columns:
                score_cols = [col for col in df.columns if col.endswith("_score")]
                score_col = score_cols[0] if score_cols else None

            if flag_col not in df.columns:
                flag_cols = [col for col in df.columns if col.endswith("_flag")]
                flag_col = flag_cols[0] if flag_cols else None

            if date_col:
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
                date_min = df[date_col].min()
                date_max = df[date_col].max()
            else:
                date_min = pd.NaT
                date_max = pd.NaT

            if score_col:
                score = pd.to_numeric(df[score_col], errors="coerce")

                if module_name.startswith("M4"):
                    score = score.clip(lower=1.0, upper=1.25)

                mean_score = score.mean()
                max_score = score.max()
            else:
                mean_score = np.nan
                max_score = np.nan

            if module_name.startswith("M4"):
                flag_count = None
            elif flag_col:
                flags = pd.to_numeric(df[flag_col], errors="coerce").fillna(0)
                flag_count = int((flags > 0).sum())
            else:
                flag_count = None

            comment = "ОК"

            if module_name.startswith("M3"):
                comment = "Используется proxy-spread доходности"
            elif module_name.startswith("M5"):
                comment = "ЦБ основной, Росказна доп."
            elif module_name.startswith("M4"):
                comment = "Множитель LSI, cap 1.25"

            rows.append(
                {
                    "Модуль": module_name,
                    "Строк": len(df),
                    "Период с": (
                        date_min.strftime("%Y-%m-%d") if pd.notna(date_min) else "-"
                    ),
                    "Период по": (
                        date_max.strftime("%Y-%m-%d") if pd.notna(date_max) else "-"
                    ),
                    "Тип показателя": (
                        "сезонный множитель"
                        if module_name.startswith("M4")
                        else "стресс-сигнал"
                    ),
                    "Средний score": (
                        round(mean_score, 2) if pd.notna(mean_score) else None
                    ),
                    "Максимальный score": (
                        round(max_score, 2) if pd.notna(max_score) else None
                    ),
                    "Кол-во флагов": flag_count,
                    "Комментарий": comment,
                }
            )

        except Exception as error:
            rows.append(
                {
                    "Модуль": module_name,
                    "Строк": 0,
                    "Период с": "-",
                    "Период по": "-",
                    "Средний score": "-",
                    "Максимальный score": "-",
                    "Кол-во флагов": "-",
                    "Комментарий": f"Ошибка чтения: {error}",
                }
            )

    return pd.DataFrame(rows)

def status_color(status: str) -> str:
    if isinstance(status, str):
        status_upper = status.upper()
        if "КРАС" in status_upper:
            return "🔴"
        if "ЖЁЛ" in status_upper or "ЖЕЛ" in status_upper:
            return "🟡"
        if "ЗЕЛ" in status_upper:
            return "🟢"
    return "⚪"


def format_number(value, digits: int = 1) -> str:
    if pd.isna(value):
        return "—"
    return f"{value:.{digits}f}"


def filter_by_period(df: pd.DataFrame, period_name: str) -> pd.DataFrame:
    max_date = df["date"].max()

    if period_name == "Последние 30 дней":
        start_date = max_date - pd.Timedelta(days=30)
    elif period_name == "Последние 90 дней":
        start_date = max_date - pd.Timedelta(days=90)
    elif period_name == "Последние 365 дней":
        start_date = max_date - pd.Timedelta(days=365)
    elif period_name == "С 2022 года":
        start_date = pd.Timestamp("2022-01-01")
    else:
        start_date = df["date"].min()

    return df[df["date"] >= start_date].copy()


def make_lsi_chart(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df["date"],
            y=df["lsi"],
            mode="lines",
            name="LSI",
            line=dict(width=2),
            hovertemplate="Дата: %{x|%Y-%m-%d}<br>LSI: %{y:.1f}<extra></extra>",
        )
    )

    fig.add_hline(
        y=40,
        line_dash="dash",
        annotation_text="Порог жёлтой зоны = 40",
        annotation_position="top left",
    )

    fig.add_hline(
        y=70,
        line_dash="dash",
        annotation_text="Порог красной зоны = 70",
        annotation_position="top left",
    )

    stress_periods = [
        ("Декабрь 2014", pd.Timestamp("2014-12-01"), pd.Timestamp("2014-12-31")),
        ("Февраль–март 2022", pd.Timestamp("2022-02-01"), pd.Timestamp("2022-03-31")),
    ]

    chart_min_date = df["date"].min()
    chart_max_date = df["date"].max()

    for name, start, end in stress_periods:
        if end < chart_min_date or start > chart_max_date:
            continue

        fig.add_vrect(
            x0=max(start, chart_min_date),
            x1=min(end, chart_max_date),
            fillcolor="LightSkyBlue",
            opacity=0.25,
            layer="below",
            line_width=0,
            annotation_text=name,
            annotation_position="top left",
        )

    fig.update_layout(
        title="Liquidity Stress Index, 0–100",
        xaxis_title="Дата",
        yaxis_title="LSI",
        xaxis=dict(range=[chart_min_date, chart_max_date]),
        yaxis=dict(range=[0, 105]),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
        height=520,
    )

    return fig


def make_contribution_chart(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    contribution_cols = [
        ("contribution_m1", "M1 резервы"),
        ("contribution_m2", "M2 репо"),
        ("contribution_m3", "M3 ОФЗ"),
        ("contribution_m4", "M4 сезонность"),
        ("contribution_m5", "M5 казначейство"),
    ]

    for col, name in contribution_cols:
        if col in df.columns:
            fig.add_trace(
                go.Scatter(
                    x=df["date"],
                    y=df[col].fillna(0),
                    mode="lines",
                    stackgroup="one",
                    name=name,
                    hovertemplate=f"{name}: " + "%{y:.1f}<extra></extra>",
                )
            )

    fig.add_trace(
        go.Scatter(
            x=df["date"],
            y=df["lsi"],
            mode="lines",
            name="LSI",
            line=dict(width=2),
            hovertemplate="LSI: %{y:.1f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="Вклад модулей в итоговый LSI",
        xaxis_title="Дата",
        yaxis_title="Пункты LSI",
        yaxis=dict(range=[0, 105]),
        hovermode="x unified",
        legend=dict(orientation="h"),
        height=520,
    )

    return fig


def make_modules_bar_chart(latest_row: pd.Series) -> go.Figure:
    modules = ["M1 резервы", "M2 репо", "M3 ОФЗ", "M5 казначейство"]
    values = [
        latest_row.get("m1_score", 0),
        latest_row.get("m2_score", 0),
        latest_row.get("m3_score", 0),
        latest_row.get("m5_score", 0),
    ]

    fig = go.Figure(
        go.Bar(
            x=modules,
            y=values,
            text=[format_number(v, 1) for v in values],
            textposition="auto",
        )
    )

    fig.update_layout(
        title="Текущие стресс-сигналы по модулям",
        yaxis_title="Score, 0–100",
        yaxis=dict(range=[0, 100]),
        height=420,
    )

    return fig


def make_status_distribution_chart(df: pd.DataFrame) -> go.Figure:
    status_counts = df["status"].fillna("Нет статуса").value_counts().reset_index()
    status_counts.columns = ["status", "count"]

    fig = go.Figure(
        go.Pie(
            labels=status_counts["status"],
            values=status_counts["count"],
            hole=0.45,
        )
    )

    fig.update_layout(
        title="Распределение дней по статусам",
        height=420,
    )

    return fig

st.title("📊 Liquidity Stress Scanner")
st.caption(
    "Интерпретируемая система раннего выявления стресса ликвидности на основе модулей М1–М5"
)

try:
    lsi_df = load_lsi_data(LSI_PATH)
except Exception as error:
    st.error(f"Не удалось загрузить LSI-файл: {error}")
    st.stop()


st.sidebar.header("Настройки")

period_name = st.sidebar.selectbox(
    "Период отображения",
    [
        "Последние 30 дней",
        "Последние 90 дней",
        "Последние 365 дней",
        "С 2022 года",
        "Вся история",
    ],
    index=2,
)

filtered_df = filter_by_period(lsi_df, period_name)

st.sidebar.markdown("---")
st.sidebar.write("Файл LSI:")
st.sidebar.code(str(LSI_PATH.relative_to(PROJECT_ROOT)))

st.sidebar.write("Период данных:")
st.sidebar.write(f"{lsi_df['date'].min().date()} — {lsi_df['date'].max().date()}")

latest = lsi_df.iloc[-1]
previous = lsi_df.iloc[-2] if len(lsi_df) >= 2 else latest

latest_lsi = latest["lsi"]
previous_lsi = previous["lsi"]
delta_lsi = latest_lsi - previous_lsi

latest_status = latest.get("status", "Нет статуса")
latest_driver = latest.get("top_driver", "—")
latest_flags = latest.get("active_flags", "—")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Текущий LSI",
        value=format_number(latest_lsi, 1),
        delta=format_number(delta_lsi, 1),
    )

with col2:
    st.metric(
        label="Статус",
        value=f"{status_color(latest_status)} {latest_status}",
    )

with col3:
    st.metric(
        label="Главный вклад в LSI",
        value=str(latest_driver),
    )

with col4:
    st.metric(
        label="Сезонный фактор",
        value=format_number(latest.get("seasonal_factor", np.nan), 2),
    )

st.info(
    f"Модульные флаги на последнюю дату: **{latest_flags if pd.notna(latest_flags) else 'нет'}**. "
    "Флаг отдельного модуля не означает автоматический переход LSI в жёлтую или красную зону."
)

tab_overview, tab_modules, tab_backtest, tab_quality, tab_data = st.tabs(
    [
        "Обзор LSI",
        "Модули",
        "Backtest",
        "Качество данных",
        "Таблицы",
    ]
)
with tab_overview:
    st.subheader("Динамика индекса")

    st.plotly_chart(make_lsi_chart(filtered_df), width="stretch")

    st.subheader("Вклад модулей")

    st.plotly_chart(make_contribution_chart(filtered_df), width="stretch")

    col_left, col_right = st.columns(2)

    with col_left:
        st.plotly_chart(make_modules_bar_chart(latest), width="stretch")

    with col_right:
        st.plotly_chart(make_status_distribution_chart(filtered_df), width="stretch")

    st.subheader("Последние 30 дней")

    last_30 = lsi_df.tail(30).copy()
    show_cols = [
        "date",
        "lsi",
        "status",
        "top_driver",
        "active_flags",
        "m1_score",
        "m2_score",
        "m3_score",
        "m5_score",
        "seasonal_factor",
    ]

    available_cols = [col for col in show_cols if col in last_30.columns]

    round_cols = [
        "lsi",
        "m1_score",
        "m2_score",
        "m3_score",
        "m5_score",
        "seasonal_factor",
    ]

    for col in round_cols:
        if col in last_30.columns:
            last_30[col] = pd.to_numeric(last_30[col], errors="coerce").round(2)

    st.dataframe(
        last_30[available_cols].sort_values("date", ascending=False),
        width="stretch",
        hide_index=True,
    )

with tab_modules:
    st.subheader("Текущие значения модулей")

    module_cols = st.columns(4)

    with module_cols[0]:
        st.metric("M1 резервы", format_number(latest.get("m1_score", np.nan), 1))

    with module_cols[1]:
        st.metric("M2 репо", format_number(latest.get("m2_score", np.nan), 1))

    with module_cols[2]:
        st.metric("M3 ОФЗ", format_number(latest.get("m3_score", np.nan), 1))

    with module_cols[3]:
        st.metric("M5 казначейство", format_number(latest.get("m5_score", np.nan), 1))

    st.markdown("""
        **Как читать модули:**

        - **M1** — давление по обязательным резервам и корсчетам.
        - **M2** — спрос банков на операции репо ЦБ.
        - **M3** — спрос на аукционах ОФЗ, используется proxy-spread доходности.
        - **M4** — налоговый календарь и сезонность, используется как множитель.
        - **M5** — движение бюджетных средств, основной источник — данные ЦБ, Росказна используется дополнительно.
        """)

    st.subheader("Динамика score по модулям")

    score_cols = {
        "m1_score": "M1 резервы",
        "m2_score": "M2 репо",
        "m3_score": "M3 ОФЗ",
        "m5_score": "M5 казначейство",
    }

    fig = go.Figure()

    for col, name in score_cols.items():
        if col in filtered_df.columns:
            fig.add_trace(
                go.Scatter(
                    x=filtered_df["date"],
                    y=filtered_df[col],
                    mode="lines",
                    name=name,
                )
            )

    fig.update_layout(
        title="Score модулей, 0–100",
        xaxis_title="Дата",
        yaxis_title="Score",
        yaxis=dict(range=[0, 100]),
        hovermode="x unified",
        height=520,
    )

    st.plotly_chart(fig, width="stretch")

with tab_backtest:
    st.subheader("Проверка индекса на стрессовых периодах")

    backtest_df = load_excel_if_exists(BACKTEST_PATH)

    if backtest_df is None:
        backtest_df = build_backtest_from_lsi(lsi_df)

    backtest_df = backtest_df.convert_dtypes()

    backtest_table = backtest_df.drop(columns=["Комментарий"], errors="ignore")

    st.dataframe(
        backtest_table,
        width="stretch",
        hide_index=True,
    )

    st.markdown("""
    **Комментарий по периодам:**  
    декабрь 2014 может определяться слабее из-за ограниченного покрытия ранних данных;  
    февраль–март 2022 используется как основной стрессовый период;  
    январь 2025 используется как отложенная проверка.
    """)

    st.markdown("""
        **Как читать backtest:**

        - `Средний LSI` показывает общий уровень давления в выбранном периоде.
        - `Максимальный LSI` показывает пик стресса.
        - `Жёлтых/красных дней` показывает, сколько дней индекс выходил из спокойной зоны.
        - Ранние периоды могут определяться слабее из-за ограниченного покрытия исторических данных.
        """)

with tab_quality:
    st.subheader("Качество и покрытие данных")

    quality_df = load_excel_if_exists(QUALITY_PATH)

    if quality_df is None:
        quality_df = build_quality_report_from_files()

    st.caption("Отчёт строится автоматически на основе итоговых файлов модулей М1–М5.")

    if quality_df is None or quality_df.empty:
        st.warning(
            "Не удалось сформировать таблицу качества данных. Проверь пути к файлам модулей."
        )
    else:
        quality_display = quality_df.copy()

        for col in quality_display.columns:
            quality_display[col] = quality_display[col].apply(
                lambda x: "—" if pd.isna(x) else str(x)
            )

        st.dataframe(
            quality_display,
            width="stretch",
            hide_index=True,
        )

    st.markdown("""
        **Ключевые ограничения:**

        - В М3 используется proxy-spread доходности, а не полноценная рыночная кривая ОФЗ.
        - В М5 основным источником является ЦБ; данные Росказны используются как дополнительный индикатор.
        - По ранним периодам покрытие данных слабее, поэтому исторический backtest нужно читать осторожно.
        - Основной LSI является интерпретируемой скоринговой моделью, а не black-box ML.
        """)
    
with tab_data:
    st.subheader("Итоговая таблица LSI")

    selected_cols = st.multiselect(
        "Выбери колонки",
        options=list(lsi_df.columns),
        default=[
            col
            for col in [
                "date",
                "lsi",
                "status",
                "top_driver",
                "active_flags",
                "m1_score",
                "m2_score",
                "m3_score",
                "m5_score",
                "seasonal_factor",
            ]
            if col in lsi_df.columns
        ],
    )

    table_df = filtered_df[selected_cols].sort_values("date", ascending=False)

    st.dataframe(
        table_df,
        width="stretch",
        hide_index=True,
    )

    csv = table_df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="Скачать выбранную таблицу CSV",
        data=csv,
        file_name="lsi_filtered.csv",
        mime="text/csv",
    )

st.markdown("---")
st.caption(
    "LSI рассчитывается как интерпретируемая взвешенная сумма сигналов М1, М2, М3 и М5 "
    "с сезонной корректировкой М4. Диапазон индекса: 0–100."
)

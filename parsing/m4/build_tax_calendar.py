from pathlib import Path
from datetime import datetime
import calendar

import pandas as pd
import numpy as np


PROJECT_ROOT = Path(__file__).resolve().parents[2]
SAVE_DIR = PROJECT_ROOT / "data" / "m4" / "tax_calendar"
SAVE_DIR.mkdir(parents=True, exist_ok=True)

START_DATE = "2010-01-01"


def next_business_day(date_value: pd.Timestamp) -> pd.Timestamp:
    date_value = pd.Timestamp(date_value).normalize()

    while date_value.weekday() >= 5:
        date_value += pd.Timedelta(days=1)

    return date_value


def month_last_day(year: int, month: int) -> pd.Timestamp:
    return pd.Timestamp(year=year, month=month, day=calendar.monthrange(year, month)[1])


def build_tax_events(start_date: str, end_date: str) -> pd.DataFrame:
    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize()

    events = []

    for year in range(start.year, end.year + 1):
        for month in range(1, 13):
            month_start = pd.Timestamp(year=year, month=month, day=1)

            if month_start > end or month_last_day(year, month) < start:
                continue

            notification_date = next_business_day(pd.Timestamp(year=year, month=month, day=25))
            payment_date = next_business_day(pd.Timestamp(year=year, month=month, day=28))
            last_day = month_last_day(year, month)

            quarter_end_flag = month in [3, 6, 9, 12]
            quarter_payment_flag = month in [1, 4, 7, 10]

            events.append(
                {
                    "event_date": notification_date,
                    "event_type": "tax_notification",
                    "event_name": "Уведомления / отчетность по налогам",
                    "event_weight": 15,
                }
            )

            events.append(
                {
                    "event_date": payment_date,
                    "event_type": "main_tax_payment",
                    "event_name": "Основной срок уплаты налогов / ЕНП",
                    "event_weight": 45 if not quarter_payment_flag else 55,
                }
            )
            if month == 12:
                ndfl_date = next_business_day(last_day)
            else:
                next_month = month + 1
                next_year = year
                ndfl_date = next_business_day(pd.Timestamp(year=next_year, month=next_month, day=5))

            events.append(
                {
                    "event_date": ndfl_date,
                    "event_type": "ndfl_second_payment",
                    "event_name": "Дополнительный срок уплаты НДФЛ",
                    "event_weight": 25,
                }
            )

            events.append(
                {
                    "event_date": last_day,
                    "event_type": "end_of_month",
                    "event_name": "Конец месяца",
                    "event_weight": 10,
                }
            )

            if quarter_end_flag:
                events.append(
                    {
                        "event_date": last_day,
                        "event_type": "end_of_quarter",
                        "event_name": "Конец квартала",
                        "event_weight": 20,
                    }
                )

    events_df = pd.DataFrame(events)
    events_df = events_df.dropna(subset=["event_date"])
    events_df["event_date"] = pd.to_datetime(events_df["event_date"]).dt.normalize()
    events_df = events_df[(events_df["event_date"] >= start) & (events_df["event_date"] <= end)].copy()
    events_df = events_df.sort_values(["event_date", "event_type"]).reset_index(drop=True)

    return events_df


def build_daily_calendar(start_date: str = START_DATE, end_date: str | None = None) -> pd.DataFrame:
    if end_date is None:
        end_date = datetime.today().strftime("%Y-%m-%d")

    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize()

    days = pd.DataFrame({"date": pd.date_range(start=start, end=end, freq="D")})
    days["weekday"] = days["date"].dt.weekday
    days["is_weekend"] = (days["weekday"] >= 5).astype(int)
    days["month"] = days["date"].dt.month
    days["year"] = days["date"].dt.year
    days["quarter"] = days["date"].dt.quarter

    events = build_tax_events(start_date=start_date, end_date=end_date)

    event_lists = (
        events.groupby("event_date")
        .agg(
            tax_event_name=("event_name", lambda values: "; ".join(sorted(set(map(str, values))))),
            tax_event_type=("event_type", lambda values: "; ".join(sorted(set(map(str, values))))),
            direct_event_weight=("event_weight", "sum"),
        )
        .reset_index()
        .rename(columns={"event_date": "date"})
    )

    calendar_df = days.merge(event_lists, on="date", how="left")
    calendar_df["tax_event_name"] = calendar_df["tax_event_name"].fillna("")
    calendar_df["tax_event_type"] = calendar_df["tax_event_type"].fillna("")
    calendar_df["direct_event_weight"] = calendar_df["direct_event_weight"].fillna(0)

    calendar_df["tax_notification_flag"] = calendar_df["tax_event_type"].str.contains("tax_notification", na=False).astype(int)
    calendar_df["tax_payment_flag"] = calendar_df["tax_event_type"].str.contains("main_tax_payment", na=False).astype(int)
    calendar_df["ndfl_second_payment_flag"] = calendar_df["tax_event_type"].str.contains("ndfl_second_payment", na=False).astype(int)
    calendar_df["end_of_month_flag"] = calendar_df["tax_event_type"].str.contains("end_of_month", na=False).astype(int)
    calendar_df["end_of_quarter_flag"] = calendar_df["tax_event_type"].str.contains("end_of_quarter", na=False).astype(int)
    payment_dates = events.loc[events["event_type"].isin(["main_tax_payment", "ndfl_second_payment"]), "event_date"].drop_duplicates()
    notification_dates = events.loc[events["event_type"].eq("tax_notification"), "event_date"].drop_duplicates()

    calendar_df["days_to_nearest_tax_payment"] = np.nan if False else np.nan
    calendar_df["days_to_nearest_tax_notification"] = np.nan if False else np.nan

    payment_values = payment_dates.to_numpy(dtype="datetime64[ns]")
    notification_values = notification_dates.to_numpy(dtype="datetime64[ns]")

    def nearest_delta_days(date_value: pd.Timestamp, event_values) -> float:
        if len(event_values) == 0:
            return np.nan
        deltas = (event_values - np.datetime64(date_value)).astype("timedelta64[D]").astype(int)
        nearest_index = np.argmin(np.abs(deltas))
        return float(deltas[nearest_index])

    calendar_df["days_to_nearest_tax_payment"] = calendar_df["date"].apply(
        lambda value: nearest_delta_days(value, payment_values)
    )
    calendar_df["days_to_nearest_tax_notification"] = calendar_df["date"].apply(
        lambda value: nearest_delta_days(value, notification_values)
    )

    calendar_df["tax_payment_window_flag"] = calendar_df["days_to_nearest_tax_payment"].between(-1, 3).astype(int)
    calendar_df["tax_notification_window_flag"] = calendar_df["days_to_nearest_tax_notification"].between(-1, 2).astype(int)
    calendar_df["tax_week_flag"] = calendar_df["days_to_nearest_tax_payment"].between(-2, 5).astype(int)

    output_columns = [
        "date",
        "year",
        "month",
        "quarter",
        "weekday",
        "is_weekend",
        "tax_event_name",
        "tax_event_type",
        "direct_event_weight",
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
    ]

    return calendar_df[output_columns].copy()


def main() -> None:
    end_date = datetime.today().strftime("%Y-%m-%d")

    print(f"Период: {START_DATE} — {end_date}")

    calendar_df = build_daily_calendar(start_date=START_DATE, end_date=end_date)

    file_name = f"tax_calendar_{START_DATE}_{end_date}.xlsx".replace(":", "-")
    save_path = SAVE_DIR / file_name

    calendar_df.to_excel(save_path, index=False)

    print(f"Строк в календаре: {len(calendar_df)}")
    print(f"Файл сохранён: {save_path}")


if __name__ == "__main__":
    main()

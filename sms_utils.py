import pandas as pd
from datetime import datetime, timedelta


WEEKDAYS_KR = ["월", "화", "수", "목", "금", "토", "일"]


def format_weekday_kr(date_obj):
    return WEEKDAYS_KR[date_obj.weekday()]


def load_send_logs(gc, spreadsheet_id, sheet_name="발송로그"):
    try:
        sheet = gc.open_by_key(spreadsheet_id)
        worksheet = sheet.worksheet(sheet_name)
        records = worksheet.get_all_records()

        if not records:
            return pd.DataFrame()

        df = pd.DataFrame(records)
        df["발송일시"] = pd.to_datetime(df["발송일시"], errors="coerce")
        df["회차"] = pd.to_numeric(df["회차"], errors="coerce")
        df = df.dropna(subset=["발송일시"])
        df = df.sort_values("발송일시", ascending=False).reset_index(drop=True)
        return df

    except Exception:
        return pd.DataFrame()


def filter_by_period(df, period):
    if df.empty:
        return df

    period_map = {
        "최근 1주": timedelta(days=7),
        "최근 2주": timedelta(days=14),
        "최근 1개월": timedelta(days=30),
        "최근 3개월": timedelta(days=90),
        "전체": None,
    }

    delta = period_map.get(period)
    if delta:
        cutoff = datetime.now() - delta
        return df[df["발송일시"] >= cutoff]
    return df


def filter_by_conditions(df, result_filter, trigger_filter):
    if df.empty:
        return df
    df = df[df["결과"].isin(result_filter)]
    df = df[df["트리거"].isin(trigger_filter)]
    return df


def get_previous_period_stats(full_df, period):
    period_days = {
        "최근 1주": 7,
        "최근 2주": 14,
        "최근 1개월": 30,
        "최근 3개월": 90,
        "전체": None,
    }
    days = period_days.get(period)
    if not days or full_df.empty:
        return None

    now = datetime.now()
    prev_start = now - timedelta(days=days * 2)
    prev_end = now - timedelta(days=days)

    prev_df = full_df[
        (full_df["발송일시"] >= prev_start) & (full_df["발송일시"] < prev_end)
    ]
    if prev_df.empty:
        return None

    prev_total = len(prev_df)
    prev_success = len(prev_df[prev_df["결과"] == "성공"])
    prev_fail = len(prev_df[prev_df["결과"] == "실패"])
    prev_skipped = len(prev_df[prev_df["결과"] == "건너뜀"])
    prev_no_contact = len(prev_df[prev_df["결과"] == "연락처 없음"])
    prev_sent = prev_total - prev_skipped - prev_no_contact

    return {
        "total": prev_total,
        "success": prev_success,
        "fail": prev_fail,
        "skipped": prev_skipped,
        "success_rate": round(prev_success / max(prev_sent, 1) * 100, 1),
    }


def get_resendable_failures(full_df, failures):
    successes = full_df[full_df["결과"] == "성공"]

    resendable = []
    for _, fail_row in failures.iterrows():
        already_sent = successes[
            (successes["회차"] == fail_row["회차"])
            & (successes["수신번호"] == fail_row["수신번호"])
            & (successes["발송일시"] > fail_row["발송일시"])
        ]
        if already_sent.empty:
            resendable.append(fail_row)

    if not resendable:
        return pd.DataFrame()
    return pd.DataFrame(resendable)

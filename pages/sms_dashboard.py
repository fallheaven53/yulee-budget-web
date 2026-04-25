import streamlit as st
import pandas as pd
import json
import os
import requests
import altair as alt
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from sms_utils import (
    load_send_logs, filter_by_period, filter_by_conditions,
    format_weekday_kr, get_previous_period_stats, get_resendable_failures,
)

st.set_page_config(page_title="SMS 발송 현황", page_icon="📱", layout="wide")
st.title("SMS 발송 현황")

SMS_LOG_SPREADSHEET_ID = "1CHTo8BALaSyS8K1TKMVM22rUbG5lbqVRDgsMiOMPCoQ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
FUNCTION_URL = (
    "https://asia-northeast3-nice-abbey-473900-e6"
    ".cloudfunctions.net/auto-sms-scheduler"
)


def _get_gc():
    creds_dict = None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
    except Exception:
        pass

    if creds_dict:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        app_dir = os.path.join(os.path.dirname(__file__), "..")
        creds_path = os.path.join(app_dir, "credentials.json")
        if not os.path.exists(creds_path):
            st.error("Google 인증 정보가 없습니다.")
            st.stop()
        creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)

    return gspread.authorize(creds)


@st.cache_data(ttl=300)
def get_cached_logs():
    gc = _get_gc()
    return load_send_logs(gc, SMS_LOG_SPREADSHEET_ID)


# ── 사이드바 필터 ──
with st.sidebar:
    st.subheader("조회 조건")

    period = st.selectbox(
        "조회 기간",
        ["최근 1주", "최근 2주", "최근 1개월", "최근 3개월", "전체"],
        index=2,
    )

    result_filter = st.multiselect(
        "결과 필터",
        ["성공", "실패", "건너뜀", "연락처 없음"],
        default=["성공", "실패", "건너뜀", "연락처 없음"],
    )

    trigger_filter = st.multiselect(
        "발송 구분",
        ["5day", "1day", "manual_resend"],
        default=["5day", "1day", "manual_resend"],
    )

    st.markdown("---")
    if st.button("데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()


# ── 데이터 로드 ──
full_df = get_cached_logs()

if full_df.empty:
    st.warning("발송로그 데이터가 없습니다.")
    st.stop()

filtered_df = filter_by_period(full_df, period)
filtered_df = filter_by_conditions(filtered_df, result_filter, trigger_filter)


# ── 1. 요약 카드 ──
def render_summary_cards(df, full_df, period):
    if df.empty:
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.metric("발송", "0건")
        with c2: st.metric("성공률", "-")
        with c3: st.metric("실패", "0건")
        with c4: st.metric("건너뜀", "0건")
        with c5: st.metric("연락처 없음", "0건")
        return

    total = len(df)
    success = len(df[df["결과"] == "성공"])
    fail = len(df[df["결과"] == "실패"])
    skipped = len(df[df["결과"] == "건너뜀"])
    no_contact = len(df[df["결과"] == "연락처 없음"])
    sent = total - skipped - no_contact
    success_rate = round(success / max(sent, 1) * 100, 1)

    prev = get_previous_period_stats(full_df, period)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        d = total - prev["total"] if prev else None
        st.metric("발송", f"{total}건", delta=f"{d}건" if d is not None else None)
    with c2:
        d = round(success_rate - prev["success_rate"], 1) if prev else None
        st.metric("성공률", f"{success_rate}%", delta=f"{d}%p" if d is not None else None)
    with c3:
        d = fail - prev["fail"] if prev else None
        st.metric("실패", f"{fail}건", delta=f"{d}건" if d is not None else None, delta_color="inverse")
    with c4:
        d = skipped - prev["skipped"] if prev else None
        st.metric("건너뜀", f"{skipped}건", delta=f"{d}건" if d is not None else None, delta_color="off")
    with c5:
        st.metric("연락처 없음", f"{no_contact}건")


render_summary_cards(filtered_df, full_df, period)

st.markdown("---")


# ── 2. 발송 추이 차트 ──
def render_weekly_chart(df):
    if df.empty:
        st.info("조회 기간 내 발송 데이터가 없습니다.")
        return

    st.subheader("발송 추이")

    chart_df = df.copy()
    chart_df["날짜"] = chart_df["발송일시"].apply(
        lambda x: f"{x.month}/{x.day} ({format_weekday_kr(x)})"
    )
    chart_df["날짜_sort"] = chart_df["발송일시"].dt.date

    pivot = chart_df.groupby(["날짜_sort", "날짜", "결과"]).size().reset_index(name="건수")

    color_map = {"성공": "#4CAF50", "실패": "#F44336", "건너뜀": "#9E9E9E", "연락처 없음": "#FF9800"}

    chart = (
        alt.Chart(pivot)
        .mark_bar()
        .encode(
            x=alt.X("건수:Q", title="건수"),
            y=alt.Y(
                "날짜:N",
                sort=alt.EncodingSortField(field="날짜_sort", order="ascending"),
                title="",
            ),
            color=alt.Color(
                "결과:N",
                scale=alt.Scale(
                    domain=list(color_map.keys()),
                    range=list(color_map.values()),
                ),
            ),
            tooltip=["날짜", "결과", "건수"],
        )
        .properties(height=max(len(pivot["날짜"].unique()) * 40, 200))
    )

    st.altair_chart(chart, use_container_width=True)


render_weekly_chart(filtered_df)

st.markdown("---")


# ── 3. 실패 사유 분류 ──
def render_failure_analysis(df):
    failures = df[df["결과"] == "실패"] if not df.empty else pd.DataFrame()

    if failures.empty:
        st.success("조회 기간 내 실패 건이 없습니다.")
        return

    st.subheader("실패 사유 분류")

    if "실패사유" in failures.columns:
        reason_counts = failures["실패사유"].value_counts().reset_index()
        reason_counts.columns = ["사유", "건수"]

        chart = (
            alt.Chart(reason_counts)
            .mark_bar(color="#F44336")
            .encode(
                x=alt.X("건수:Q", title="건수"),
                y=alt.Y("사유:N", sort="-x", title=""),
                tooltip=["사유", "건수"],
            )
            .properties(height=max(len(reason_counts) * 35, 100))
        )
        st.altair_chart(chart, use_container_width=True)

    st.markdown("---")
    st.caption("최근 실패 상세")

    cols = ["발송일시", "트리거", "단체명", "수신자", "역할", "채널", "실패사유"]
    available = [c for c in cols if c in failures.columns]
    fail_display = failures[available].copy()
    fail_display["발송일시"] = fail_display["발송일시"].dt.strftime("%m/%d %H:%M")
    st.dataframe(fail_display, use_container_width=True, hide_index=True)


render_failure_analysis(filtered_df)

st.markdown("---")


# ── 4. 수동 재발송 ──
def render_manual_resend(df):
    failures = df[df["결과"] == "실패"] if not df.empty else pd.DataFrame()
    if failures.empty:
        return

    st.subheader("수동 재발송")

    resendable = get_resendable_failures(df, failures)
    if resendable.empty:
        st.info("재발송 가능한 실패 건이 없습니다. (모두 재발송 완료)")
        return

    options = []
    for _, row in resendable.iterrows():
        label = (
            f"{row['발송일시'].strftime('%m/%d %H:%M')} | "
            f"{row['단체명']} | {row['수신자']} ({row['역할']}) | "
            f"{row.get('실패사유', '')}"
        )
        options.append({"label": label, "row": row})

    selected_labels = st.multiselect(
        "재발송할 건을 선택하세요",
        [o["label"] for o in options],
    )

    if not selected_labels:
        st.caption("재발송할 건을 선택하면 실행 버튼이 나타납니다.")
        return

    st.warning(f"{len(selected_labels)}건을 재발송합니다.")

    if st.button("재발송 실행", type="primary"):
        for label in selected_labels:
            option = next(o for o in options if o["label"] == label)
            row = option["row"]

            payload = {
                "trigger_type": "manual_resend",
                "round": int(row["회차"]),
                "troupe": row["단체명"],
                "recipient": row["수신자"],
                "role": row["역할"],
            }

            try:
                resp = requests.post(FUNCTION_URL, json=payload, timeout=30)
                if resp.status_code == 200:
                    result = resp.json()
                    if result.get("success"):
                        st.success(f"{row['단체명']} {row['수신자']}: 재발송 성공")
                    else:
                        st.error(f"{row['단체명']} {row['수신자']}: {result.get('message', '실패')}")
                else:
                    st.error(f"{row['단체명']} {row['수신자']}: HTTP {resp.status_code}")
            except Exception as e:
                st.error(f"{row['단체명']} {row['수신자']}: 오류 - {e}")

        st.cache_data.clear()


render_manual_resend(full_df)

st.markdown("---")


# ── 5. 발송 상세 테이블 ──
def render_detail_table(df):
    if df.empty:
        st.info("조회 기간 내 발송 데이터가 없습니다.")
        return

    st.subheader("발송 상세")

    cols = ["발송일시", "트리거", "회차", "단체명", "수신자", "역할", "결과", "채널"]
    available = [c for c in cols if c in df.columns]
    display_df = df[available].copy()
    display_df["발송일시"] = display_df["발송일시"].dt.strftime("%Y-%m-%d %H:%M")

    def color_result(val):
        if val == "성공":
            return "color: #4CAF50"
        elif val == "실패":
            return "color: #F44336; font-weight: bold"
        elif val == "건너뜀":
            return "color: #9E9E9E"
        elif val == "연락처 없음":
            return "color: #FF9800"
        return ""

    try:
        styled = display_df.style.map(color_result, subset=["결과"])
    except AttributeError:
        styled = display_df.style.applymap(color_result, subset=["결과"])

    st.dataframe(styled, use_container_width=True, hide_index=True, height=400)


render_detail_table(filtered_df)

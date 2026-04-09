"""
율이공방 — 예산 관리 (Web)
Streamlit 기반 예산 집행 현황 관리기
"""

import streamlit as st
import pandas as pd
import os, io
from datetime import datetime, date
from data_manager import (
    DataManager, ProjectData, COMMON_PROJECT,
    DEFAULT_PROJECTS, clean_num, fmt_won
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(APP_DIR, "예산집행_DB.xlsx")

# ── 비밀번호 설정 ──
PASSWORD = os.environ.get("APP_PASSWORD", "yulee0328")


# ═══════════════════════════════════════════
#  세션 초기화
# ═══════════════════════════════════════════

def get_dm() -> DataManager:
    """DataManager를 세션에 유지"""
    if "dm" not in st.session_state:
        st.session_state.dm = DataManager(DB_FILE)
    return st.session_state.dm


def reload_dm():
    """데이터 강제 리로드"""
    st.session_state.dm = DataManager(DB_FILE)


# ═══════════════════════════════════════════
#  로그인
# ═══════════════════════════════════════════

def login_page():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## 율이공방 — 예산 관리")
        st.markdown("##### 로그인")
        pw = st.text_input("비밀번호", type="password", key="login_pw")
        if st.button("로그인", use_container_width=True, type="primary"):
            if pw == PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("비밀번호가 틀립니다.")


# ═══════════════════════════════════════════
#  사이드바
# ═══════════════════════════════════════════

def render_sidebar():
    dm = get_dm()

    st.sidebar.markdown("### 율이공방 — 예산 관리")
    st.sidebar.divider()

    # 사업 선택
    st.sidebar.markdown("**사업 선택**")
    idx = dm.project_names.index(dm.current_project) if dm.current_project in dm.project_names else 0
    selected = st.sidebar.selectbox(
        "사업", dm.project_names, index=idx,
        key="sb_project", label_visibility="collapsed")

    if selected != dm.current_project:
        dm.switch_project(selected)
        st.rerun()

    # 사업 관리 버튼
    c1, c2, c3 = st.sidebar.columns(3)
    with c1:
        if st.button("추가", key="btn_add_pj", use_container_width=True):
            st.session_state.show_add_pj = True
    with c2:
        if st.button("수정", key="btn_rename_pj", use_container_width=True):
            st.session_state.show_rename_pj = True
    with c3:
        if st.button("삭제", key="btn_del_pj", use_container_width=True):
            st.session_state.show_del_pj = True

    # 사업 추가 다이얼로그
    if st.session_state.get("show_add_pj"):
        with st.sidebar.form("add_pj_form"):
            new_name = st.text_input("새 사업명")
            if st.form_submit_button("추가"):
                if new_name.strip():
                    if new_name.strip() in dm.projects:
                        st.sidebar.error("이미 존재하는 사업입니다.")
                    else:
                        dm.add_project(new_name.strip())
                        dm.save()
                        st.session_state.show_add_pj = False
                        st.rerun()
                else:
                    st.sidebar.warning("사업명을 입력해주세요.")

    # 사업명 수정 다이얼로그
    if st.session_state.get("show_rename_pj"):
        with st.sidebar.form("rename_pj_form"):
            new_name = st.text_input("새 사업명", value=dm.current_project)
            if st.form_submit_button("변경"):
                nn = new_name.strip()
                if nn and nn != dm.current_project:
                    if dm.rename_project(dm.current_project, nn):
                        dm.save()
                        st.session_state.show_rename_pj = False
                        st.rerun()
                    else:
                        st.sidebar.error("이미 존재하는 사업명입니다.")
                else:
                    st.session_state.show_rename_pj = False
                    st.rerun()

    # 사업 삭제 확인
    if st.session_state.get("show_del_pj"):
        st.sidebar.warning(f"'{dm.current_project}' 사업을 삭제합니다.")
        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("삭제 확인", type="primary", use_container_width=True):
                if len(dm.project_names) <= 1:
                    st.sidebar.error("최소 1개의 사업이 필요합니다.")
                else:
                    dm.delete_project(dm.current_project)
                    dm.save()
                    st.session_state.show_del_pj = False
                    st.rerun()
        with c2:
            if st.button("취소", use_container_width=True):
                st.session_state.show_del_pj = False
                st.rerun()

    st.sidebar.divider()

    # 연도 / 총예산 (관 공통이 아닐 때)
    if not dm.is_common:
        st.sidebar.markdown("**연도 / 총예산**")
        new_year = st.sidebar.number_input("연도", value=dm.year, min_value=2020,
                                            max_value=2040, key="sb_year")
        new_budget = st.sidebar.text_input("총예산(원)", value=f"{dm.total_budget:,}",
                                            key="sb_budget")
        if new_year != dm.year:
            dm.year = new_year
            dm.save()
        budget_val = clean_num(new_budget)
        if budget_val != dm.total_budget:
            dm.total_budget = budget_val
            dm.save()

    st.sidebar.divider()

    # 데이터 관리
    st.sidebar.markdown("**데이터 관리**")

    # 백업 다운로드
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            st.sidebar.download_button(
                "데이터 백업 (다운로드)",
                f.read(),
                file_name=f"예산집행_DB_백업_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    # 데이터 복원 (업로드)
    uploaded = st.sidebar.file_uploader("데이터 복원 (업로드)", type=["xlsx"],
                                         key="restore_file")
    if uploaded is not None:
        if st.sidebar.button("복원 실행", type="primary", use_container_width=True):
            with open(DB_FILE, "wb") as f:
                f.write(uploaded.read())
            reload_dm()
            st.sidebar.success("데이터가 복원되었습니다.")
            st.rerun()

    st.sidebar.divider()
    st.sidebar.caption("율이공방 — 예산 관리 (Web) v1.0")


# ═══════════════════════════════════════════
#  탭1: 예산 계획
# ═══════════════════════════════════════════

def render_plan_tab():
    dm = get_dm()

    if dm.is_common:
        st.info("관 공통(부대경비)은 편성 예산이 없습니다. 집행 내역 탭에서 바로 등록하세요.")
        return

    st.markdown(f"**{dm.current_project}** — {dm.year}년 | 총예산 {fmt_won(dm.total_budget)}")

    # ── 편성목 목록 ──
    st.markdown("#### 편성목")

    if dm.categories:
        cat_data = []
        for cat in dm.categories:
            sp = dm.cat_spent(cat["name"])
            cat_data.append({
                "편성목명": cat["name"],
                "코드": cat["code"],
                "예산(원)": f"{cat['budget']:,}",
                "집행액": f"{sp:,}",
                "잔액": f"{cat['budget'] - sp:,}",
            })
        st.dataframe(pd.DataFrame(cat_data), use_container_width=True, hide_index=True)
    else:
        st.caption("등록된 편성목이 없습니다.")

    # 편성목 추가
    with st.expander("편성목 추가 / 삭제"):
        with st.form("add_cat_form"):
            cc1, cc2, cc3 = st.columns([3, 2, 2])
            with cc1:
                new_cat_name = st.text_input("편성목명")
            with cc2:
                new_cat_code = st.text_input("코드")
            with cc3:
                new_cat_budget = st.text_input("예산(원)", value="0")
            if st.form_submit_button("편성목 추가"):
                if new_cat_name.strip():
                    if dm.get_cat(new_cat_name.strip()):
                        st.error("이미 존재하는 편성목입니다.")
                    else:
                        cat = {"name": new_cat_name.strip(),
                               "code": new_cat_code.strip(),
                               "budget": clean_num(new_cat_budget),
                               "items": []}
                        dm.categories.append(cat)
                        dm.save()
                        st.rerun()

        # 편성목 삭제
        if dm.categories:
            cat_names = [c["name"] for c in dm.categories]
            del_cat = st.selectbox("삭제할 편성목", cat_names, key="del_cat_sel")
            if st.button("편성목 삭제", key="btn_del_cat"):
                dm.categories = [c for c in dm.categories if c["name"] != del_cat]
                if del_cat in dm.monthly:
                    del dm.monthly[del_cat]
                dm.save()
                st.rerun()

    # ── 세부항목 ──
    if dm.categories:
        st.markdown("#### 세부항목")
        sel_cat_name = st.selectbox("편성목 선택", [c["name"] for c in dm.categories],
                                     key="plan_cat_sel")
        cat = dm.get_cat(sel_cat_name)
        if cat:
            if cat["items"]:
                item_data = []
                for item in cat["items"]:
                    sp = dm.item_spent(cat["name"], item["name"])
                    item_data.append({
                        "세부항목명": item["name"],
                        "예산(원)": f"{item['budget']:,}",
                        "집행액": f"{sp:,}",
                        "잔액": f"{item['budget'] - sp:,}",
                    })
                st.dataframe(pd.DataFrame(item_data), use_container_width=True, hide_index=True)

                total_item = sum(it["budget"] for it in cat["items"])
                if total_item != cat["budget"]:
                    diff = total_item - cat["budget"]
                    st.warning(f"세부항목 합계({total_item:,}원)가 편성목 예산({cat['budget']:,}원)과 "
                               f"{abs(diff):,}원 차이납니다.")
            else:
                st.caption("등록된 세부항목이 없습니다.")

            # 세부항목 추가
            with st.expander("세부항목 추가 / 삭제"):
                with st.form("add_item_form"):
                    ic1, ic2 = st.columns([3, 2])
                    with ic1:
                        new_item_name = st.text_input("세부항목명")
                    with ic2:
                        new_item_budget = st.text_input("예산(원)", value="0",
                                                         key="new_item_budget")
                    submitted = st.form_submit_button("세부항목 추가")
                    if submitted and new_item_name.strip():
                        cat["items"].append({
                            "name": new_item_name.strip(),
                            "budget": clean_num(new_item_budget)
                        })
                        dm.sync_cat_budget(cat["name"])
                        dm.save()
                        st.rerun()

                if cat["items"]:
                    del_item = st.selectbox("삭제할 세부항목",
                                             [i["name"] for i in cat["items"]],
                                             key="del_item_sel")
                    if st.button("세부항목 삭제", key="btn_del_item"):
                        cat["items"] = [i for i in cat["items"] if i["name"] != del_item]
                        dm.sync_cat_budget(cat["name"])
                        dm.save()
                        st.rerun()

    # ── 월별 배분 ──
    if dm.categories:
        st.markdown("#### 월별 배분 계획")

        sel_month_cat = st.selectbox("편성목 선택", [c["name"] for c in dm.categories],
                                      key="monthly_cat_sel")

        md = dm.monthly.get(sel_month_cat, {})
        with st.form("monthly_form"):
            cols = st.columns(6)
            new_monthly = {}
            for i, m in enumerate(range(1, 13)):
                with cols[i % 6]:
                    val = st.number_input(f"{m}월", value=md.get(m, 0),
                                           min_value=0, step=100000,
                                           key=f"month_{m}")
                    new_monthly[m] = val

            if st.form_submit_button("월별 배분 저장"):
                dm.monthly[sel_month_cat] = new_monthly
                dm.save()
                st.success("저장되었습니다.")
                st.rerun()

        # 월별 요약 테이블
        total_plan = sum(md.get(m, 0) for m in range(1, 13))
        if total_plan > 0:
            st.caption(f"연간 배분 합계: {total_plan:,}원")


# ═══════════════════════════════════════════
#  탭2: 집행 내역
# ═══════════════════════════════════════════

def render_records_tab():
    dm = get_dm()

    st.markdown(f"**{dm.current_project}** — 집행 내역")

    # ── 수정 모드: 등록 대신 수정 폼 표시 ──
    if st.session_state.get("edit_mode", False):
        target_id = st.session_state.get("edit_target_id")
        # dm.records에서 직접 원본 데이터 가져오기
        target_rec = None
        for r in dm.records:
            if r["id"] == target_id:
                target_rec = r
                break

        if target_rec is None:
            st.session_state["edit_mode"] = False
            st.rerun()
        else:
            # 편성목 옵션
            if dm.is_common:
                cat_list = dm.all_project_cat_names()
            else:
                cat_list = [c["name"] for c in dm.categories]
            cat_val = target_rec["cat"]
            cat_idx = cat_list.index(cat_val) if cat_val in cat_list else 0

            st.markdown(f"##### 집행 내역 수정 — ID {target_id}")
            with st.form(f"edit_form_{target_id}"):
                ec1, ec2 = st.columns(2)
                with ec1:
                    e_date = st.text_input("집행일", value=str(target_rec["date"]))
                with ec2:
                    e_cat = st.selectbox("편성목", cat_list, index=cat_idx)

                ec3, ec4 = st.columns(2)
                with ec3:
                    e_item = st.text_input("세부항목", value=str(target_rec["item"]))
                with ec4:
                    e_detail = st.text_input("세부내용", value=str(target_rec["detail"]))

                ec5, ec6, ec7 = st.columns(3)
                with ec5:
                    e_amount = st.text_input("금액(원)", value=f"{target_rec['amount']:,}")
                with ec6:
                    e_round = st.text_input("회차", value=str(target_rec["round_"]))
                with ec7:
                    e_memo = st.text_input("비고", value=str(target_rec["memo"]))

                bc1, bc2 = st.columns(2)
                with bc1:
                    submitted = st.form_submit_button("수정 저장", type="primary",
                                                      use_container_width=True)
                with bc2:
                    cancelled = st.form_submit_button("취소", use_container_width=True)

                if cancelled:
                    st.session_state["edit_mode"] = False
                    st.rerun()

                if submitted:
                    if not e_cat:
                        st.error("편성목을 선택해주세요.")
                    elif not e_detail.strip():
                        st.error("세부내용을 입력해주세요.")
                    elif clean_num(e_amount) <= 0:
                        st.error("금액은 0보다 커야 합니다.")
                    else:
                        date_str = e_date.strip().replace("/", "-")
                        dm.update_record(target_id, {
                            "date": date_str,
                            "cat": e_cat,
                            "item": e_item.strip() if e_item else "",
                            "detail": e_detail.strip(),
                            "amount": clean_num(e_amount),
                            "round_": e_round.strip(),
                            "memo": e_memo.strip(),
                        })
                        st.session_state["edit_mode"] = False
                        st.success("수정 완료")
                        st.rerun()

    # ── 신규 등록 (수정 모드가 아닐 때만 표시) ──
    else:
      with st.expander("집행 내역 등록"):
        with st.form("add_record_form"):
            ac1, ac2 = st.columns(2)
            with ac1:
                add_date = st.date_input("집행일", value=date.today(), key="add_date")
            with ac2:
                if dm.is_common:
                    add_cat_opts = dm.all_project_cat_names()
                else:
                    add_cat_opts = [c["name"] for c in dm.categories]
                add_cat = st.selectbox("편성목", [""] + add_cat_opts, key="add_cat")

            ac3, ac4 = st.columns(2)
            with ac3:
                if dm.is_common:
                    add_item = st.text_input("세부항목 (선택)", key="add_item")
                else:
                    add_item_opts = []
                    if add_cat:
                        cat_obj = dm.get_cat(add_cat)
                        if cat_obj:
                            add_item_opts = [i["name"] for i in cat_obj["items"]]
                    add_item = st.selectbox("세부항목", [""] + add_item_opts, key="add_item")
            with ac4:
                add_detail = st.text_input("세부내용", key="add_detail")

            ac5, ac6, ac7 = st.columns(3)
            with ac5:
                add_amount = st.text_input("금액(원)", key="add_amount")
            with ac6:
                add_round = st.text_input("회차", key="add_round")
            with ac7:
                add_memo = st.text_input("비고", key="add_memo")

            if st.form_submit_button("등록", type="primary", use_container_width=True):
                if not add_cat:
                    st.error("편성목을 선택해주세요.")
                elif not dm.is_common and not add_item:
                    st.error("세부항목을 선택해주세요.")
                elif not add_detail.strip():
                    st.error("세부내용을 입력해주세요.")
                elif clean_num(add_amount) <= 0:
                    st.error("금액은 0보다 커야 합니다.")
                else:
                    dm.add_record({
                        "date": add_date.strftime("%Y-%m-%d"),
                        "cat": add_cat,
                        "item": add_item if add_item else "",
                        "detail": add_detail.strip(),
                        "amount": clean_num(add_amount),
                        "round_": add_round.strip(),
                        "memo": add_memo.strip(),
                    })
                    st.success("등록 완료")
                    st.rerun()

    # ── 필터링 ──
    fl1, fl2, fl3 = st.columns(3)
    with fl1:
        if dm.is_common:
            flt_cats = ["전체"] + dm.all_project_cat_names()
        else:
            flt_cats = ["전체"] + [c["name"] for c in dm.categories]
        flt_cat = st.selectbox("편성목 필터", flt_cats, key="flt_cat")
    with fl2:
        flt_months = ["전체"] + [f"{m}월" for m in range(1, 13)]
        flt_month = st.selectbox("월 필터", flt_months, key="flt_month")
    with fl3:
        if dm.is_common:
            memo_vals = sorted(set(r.get("memo", "").strip() for r in dm.records if r.get("memo", "").strip()))
            flt_memos = ["전체"] + memo_vals
            flt_memo = st.selectbox("비고 필터", flt_memos, key="flt_memo")
        else:
            flt_memo = "전체"

    # ── 목록 ──
    records = dm.records.copy()
    if flt_cat != "전체":
        records = [r for r in records if r["cat"] == flt_cat]
    if flt_month != "전체":
        m = flt_month.replace("월", "").zfill(2)
        records = [r for r in records if r["date"] and r["date"][5:7] == m]
    if flt_memo != "전체":
        records = [r for r in records if r.get("memo", "").strip() == flt_memo]
    records = sorted(records, key=lambda r: r.get("date", ""))

    if records:
        total_amount = sum(r["amount"] for r in records)
        st.caption(f"조회 결과: {len(records)}건 | 합계: {total_amount:,}원")

        display_data = []
        for r in records:
            display_data.append({
                "ID": r["id"],
                "집행일": r["date"],
                "편성목": r["cat"],
                "세부항목": r["item"],
                "세부내용": r["detail"],
                "금액(원)": f"{r['amount']:,}",
                "회차": r["round_"],
                "비고": r["memo"],
            })
        df = pd.DataFrame(display_data)
        st.dataframe(df, use_container_width=True, hide_index=True)

        # ── 수정 / 삭제 ──
        st.markdown("##### 수정 / 삭제")
        rec_ids = [r["id"] for r in records]
        sel_id = st.selectbox("ID 선택", rec_ids, key="action_rec_id")

        ac1, ac2 = st.columns(2)
        with ac1:
            if st.button("수정", key="btn_edit_rec", use_container_width=True):
                st.session_state["edit_mode"] = True
                st.session_state["edit_target_id"] = sel_id
                st.rerun()
        with ac2:
            if st.button("삭제", key="btn_del_rec", use_container_width=True):
                st.session_state.confirm_del_rec = sel_id

        if st.session_state.get("confirm_del_rec"):
            rid = st.session_state.confirm_del_rec
            st.warning(f"ID {rid} 집행 내역을 삭제합니다.")
            dc1, dc2 = st.columns(2)
            with dc1:
                if st.button("삭제 확인", type="primary", key="confirm_del_yes"):
                    dm.delete_record(rid)
                    st.session_state.confirm_del_rec = None
                    st.rerun()
            with dc2:
                if st.button("취소", key="confirm_del_no"):
                    st.session_state.confirm_del_rec = None
                    st.rerun()
    else:
        st.info("등록된 집행 내역이 없습니다.")


# ═══════════════════════════════════════════
#  탭3: 현황 대시보드
# ═══════════════════════════════════════════

def render_dashboard_tab():
    dm = get_dm()

    st.markdown(f"**{dm.current_project}** — 현황 대시보드")

    total_sp = dm.total_spent()

    if dm.is_common:
        # ── 관 공통 대시보드 ──
        rec_count = len(dm.records)

        mc1, mc2 = st.columns(2)
        with mc1:
            st.metric("집행 누계액", f"{total_sp:,}원")
        with mc2:
            st.metric("집행 건수", f"{rec_count}건")

        # 비고별 분류
        if dm.records:
            st.markdown("##### 비고별 집행 현황")
            memo_sums = {}
            for r in dm.records:
                memo = r.get("memo", "").strip() or "(미분류)"
                if memo not in memo_sums:
                    memo_sums[memo] = {"count": 0, "amount": 0}
                memo_sums[memo]["count"] += 1
                memo_sums[memo]["amount"] += r["amount"]

            memo_df = pd.DataFrame([
                {"비고(관련사업)": k, "건수": v["count"], "집행액(원)": f"{v['amount']:,}"}
                for k, v in sorted(memo_sums.items(), key=lambda x: -x[1]["amount"])
            ])
            st.dataframe(memo_df, use_container_width=True, hide_index=True)

            # 편성목별 현황
            st.markdown("##### 편성목별 집행 현황")
            cat_sums = {}
            for r in dm.records:
                cn = r["cat"]
                if cn not in cat_sums:
                    cat_sums[cn] = 0
                cat_sums[cn] += r["amount"]

            cat_df = pd.DataFrame([
                {"편성목": k, "집행액(원)": f"{v:,}"}
                for k, v in sorted(cat_sums.items(), key=lambda x: -x[1])
            ])
            st.dataframe(cat_df, use_container_width=True, hide_index=True)

    else:
        # ── 일반 사업 대시보드 ──
        total_bud = dm.total_budget
        total_rem = total_bud - total_sp
        total_rate = total_sp / total_bud * 100 if total_bud else 0

        # 요약 카드
        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.metric("연간 총 예산", f"{total_bud:,}원")
        with mc2:
            st.metric("총 집행액", f"{total_sp:,}원")
        with mc3:
            st.metric("총 잔액", f"{total_rem:,}원",
                      delta=f"{total_rem:,}원 남음" if total_rem >= 0 else f"{abs(total_rem):,}원 초과",
                      delta_color="normal" if total_rem >= 0 else "inverse")
        with mc4:
            st.metric("전체 집행률", f"{total_rate:.1f}%")

        # 진행률 바
        st.progress(min(total_rate / 100, 1.0))

        # 편성목별 현황
        st.markdown("##### 편성목별 현황")
        if dm.categories:
            cat_rows = []
            for cat in dm.categories:
                sp = dm.cat_spent(cat["name"])
                rem = cat["budget"] - sp
                rate = f"{sp/cat['budget']*100:.1f}%" if cat["budget"] else "-"
                cat_rows.append({
                    "편성목": cat["name"],
                    "코드": cat["code"],
                    "예산(원)": f"{cat['budget']:,}",
                    "집행액(원)": f"{sp:,}",
                    "잔액(원)": f"{rem:,}",
                    "집행률": rate,
                })
            cat_total_bud = sum(c["budget"] for c in dm.categories)
            cat_total_sp = sum(dm.cat_spent(c["name"]) for c in dm.categories)
            cat_total_rem = cat_total_bud - cat_total_sp
            cat_total_rate = (f"{cat_total_sp/cat_total_bud*100:.1f}%"
                              if cat_total_bud else "-")
            cat_rows.append({
                "편성목": "합  계",
                "코드": "",
                "예산(원)": f"{cat_total_bud:,}",
                "집행액(원)": f"{cat_total_sp:,}",
                "잔액(원)": f"{cat_total_rem:,}",
                "집행률": cat_total_rate,
            })
            st.dataframe(pd.DataFrame(cat_rows), use_container_width=True, hide_index=True)

        # 월별 현황
        st.markdown("##### 월별 집행 현황")
        month_rows = []
        monthly_plans = []
        monthly_spents = []
        for m in range(1, 13):
            plan = dm.monthly_plan(m)
            sp = dm.monthly_spent(m)
            diff = sp - plan
            rate = f"{sp/plan*100:.1f}%" if plan else "-"
            monthly_plans.append(plan)
            monthly_spents.append(sp)
            month_rows.append({
                "월": f"{m}월",
                "계획(원)": f"{plan:,}" if plan else "-",
                "집행(원)": f"{sp:,}",
                "차이(원)": f"{diff:+,}" if plan else "-",
                "달성률": rate,
            })
        st.dataframe(pd.DataFrame(month_rows), use_container_width=True, hide_index=True)

        # 그래프
        st.markdown("##### 월별 집행 추이")
        import plotly.graph_objects as go

        months_labels = [f"{m}월" for m in range(1, 13)]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=months_labels, y=monthly_spents,
            name="월별 집행", marker_color="#89b4fa", opacity=0.6))
        fig.add_trace(go.Scatter(
            x=months_labels, y=monthly_spents,
            name="집행 추이", mode="lines+markers",
            line=dict(color="#89b4fa", width=2),
            marker=dict(size=7)))
        if any(p > 0 for p in monthly_plans):
            fig.add_trace(go.Scatter(
                x=months_labels, y=monthly_plans,
                name="계획", mode="lines+markers",
                line=dict(color="#cdd6f4", width=1.5, dash="dash"),
                marker=dict(size=5, symbol="square")))

        fig.update_layout(
            plot_bgcolor="#313244", paper_bgcolor="#1e1e2e",
            font=dict(color="#cdd6f4"),
            legend=dict(bgcolor="#313244"),
            margin=dict(l=20, r=20, t=30, b=20),
            height=300,
            yaxis=dict(gridcolor="#45475a"),
            xaxis=dict(gridcolor="#45475a"),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── 내보내기 ──
    st.markdown("##### 내보내기")
    ex1, ex2 = st.columns(2)

    with ex1:
        if st.button("정산표 생성", use_container_width=True):
            wb = dm.export_settlement_wb()
            buf = io.BytesIO()
            wb.save(buf)
            st.session_state.settlement_data = buf.getvalue()
            st.session_state.settlement_name = (
                f"{dm.current_project}_정산표_{datetime.now().strftime('%Y%m%d')}.xlsx")

        if st.session_state.get("settlement_data"):
            st.download_button(
                "정산표 다운로드",
                st.session_state.settlement_data,
                file_name=st.session_state.settlement_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    with ex2:
        if st.button("월별 현황 생성", use_container_width=True):
            wb = dm.export_monthly_wb()
            buf = io.BytesIO()
            wb.save(buf)
            st.session_state.monthly_data = buf.getvalue()
            st.session_state.monthly_name = (
                f"{dm.current_project}_월별현황_{datetime.now().strftime('%Y%m%d')}.xlsx")

        if st.session_state.get("monthly_data"):
            st.download_button(
                "월별 현황 다운로드",
                st.session_state.monthly_data,
                file_name=st.session_state.monthly_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)


# ═══════════════════════════════════════════
#  메인
# ═══════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="율이공방 — 예산 관리",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="auto",
    )

    # 로그인 체크
    if not st.session_state.get("logged_in"):
        login_page()
        return

    # 다크 테마 CSS
    st.markdown("""
    <style>
    .stApp { background-color: #1e1e2e; }
    .stMetric label { color: #888 !important; }
    .stMetric [data-testid="stMetricValue"] { color: #f5c542 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #313244; color: #cdd6f4;
        border-radius: 8px 8px 0 0; padding: 8px 20px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #f5c542 !important; color: #1e1e2e !important;
    }
    [data-testid="stSidebar"] { background-color: #181825; }
    .stDataFrame { border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

    render_sidebar()

    # 탭
    tab1, tab2, tab3 = st.tabs(["📋 예산 계획", "📝 집행 내역", "📊 현황 대시보드"])

    with tab1:
        render_plan_tab()

    with tab2:
        render_records_tab()

    with tab3:
        render_dashboard_tab()


if __name__ == "__main__":
    main()

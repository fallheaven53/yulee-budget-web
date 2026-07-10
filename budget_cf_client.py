# -*- coding: utf-8 -*-
"""웹앱 → 예산 중계 CF(#2026-101W) 클라이언트. HMAC-SHA256 서명 호출.

app_key는 Streamlit Secrets(budget_app_key)에서만 읽는다(코드·리포 미기재).
쓰기는 이 클라이언트를 경유해 CF로 통일 → 사업목록 상태열·sheet_rev를 CF가 일관 관리.
읽기는 당분간 시트 직접 유지(gsheet_sync). 실패 시 BudgetCFError."""
import json
import time
import hmac
import hashlib
import uuid

import requests

CF_URL = "https://budget-relay-6qpfysu3yq-du.a.run.app"
_TIMEOUT = 30


class BudgetCFError(Exception):
    pass


def _app_key():
    import streamlit as st
    return st.secrets["budget_app_key"]


def _sign(method, path, body, ts, nonce, app_key):
    bh = hashlib.sha256(body or b"").hexdigest()
    msg = f"{method}\n{path}\n{bh}\n{ts}\n{nonce}"
    return hmac.new(app_key.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).hexdigest()


def _call(method, path, obj=None):
    app_key = _app_key()
    body = json.dumps(obj, ensure_ascii=False).encode("utf-8") if obj is not None else b""
    ts = str(int(time.time()))
    nonce = uuid.uuid4().hex
    sig = _sign(method, path, body, ts, nonce, app_key)
    headers = {"X-App-Id": "budget-web", "X-Timestamp": ts, "X-Nonce": nonce,
               "X-Signature": sig, "Content-Type": "application/json"}
    try:
        r = requests.request(method, CF_URL + path, data=body, headers=headers, timeout=_TIMEOUT)
    except requests.RequestException as e:
        raise BudgetCFError(f"연결 실패: {e}")
    try:
        j = r.json()
    except ValueError:
        raise BudgetCFError(f"HTTP {r.status_code}: 비정상 응답")
    if not j.get("ok"):
        raise BudgetCFError(j.get("error", f"HTTP {r.status_code}"))
    return j


# ── 조회 ──
def list_projects():
    return _call("GET", "/projects")["data"]


def read_project(name):
    # path는 원문(디코딩)으로 서명. requests가 전송 시 URL 인코딩하고 CF는 디코딩된 request.path로 검증.
    return _call("GET", "/project/" + name)


# ── 집행 ──
def record_add(project, rec):
    return _call("POST", "/record/add", {"project": project, "record": rec})


def record_update(project, rid, rec):
    return _call("POST", "/record/update", {"project": project, "id": rid, "record": rec})


def record_delete(project, rid):
    return _call("POST", "/record/delete", {"project": project, "id": rid})


# ── 계획 ──
def plan_settings(project, year, total_budget):
    return _call("POST", "/plan/settings", {"project": project, "year": year, "total_budget": total_budget})


def plan_cat_add(project, category):
    return _call("POST", "/plan/cat_add", {"project": project, "category": category})


def plan_cat_del(project, cat_name):
    return _call("POST", "/plan/cat_del", {"project": project, "cat_name": cat_name})


def plan_item_add(project, cat_name, item):
    return _call("POST", "/plan/item_add", {"project": project, "cat_name": cat_name, "item": item})


def plan_item_del(project, cat_name, item_name):
    return _call("POST", "/plan/item_del", {"project": project, "cat_name": cat_name, "item_name": item_name})


def plan_monthly(project, cat_name, months):
    return _call("POST", "/plan/monthly", {"project": project, "cat_name": cat_name, "months": months})


# ── 사업 ──
def project_add(name):
    return _call("POST", "/project/add", {"project": name})


def project_rename(old, new):
    return _call("POST", "/project/rename", {"old": old, "new": new})


def project_delete(name):
    """물리 삭제(#2026-101W 조정): 사업목록 행+사업별 3탭+rev 완전 제거. 동기 확인."""
    return _call("POST", "/project/delete", {"project": name})

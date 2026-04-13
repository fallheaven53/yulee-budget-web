"""
율이공방 — 구글 스프레드시트 동기화 모듈
로컬(tkinter)과 웹앱(Streamlit) 모두에서 사용
"""

import json, os
from data_manager import ProjectData, clean_num, DEFAULT_PROJECTS, COMMON_PROJECT

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 시트 이름 규칙: 로컬 엑셀과 동일
# - "사업목록"
# - "{사업명}_계획", "{사업명}_집행내역", "{사업명}_월별배분"


class GoogleSheetSync:
    """구글 스프레드시트 읽기/쓰기"""

    def __init__(self, credentials_path=None, credentials_dict=None,
                 spreadsheet_id=None):
        """
        credentials_path: 서비스 계정 JSON 파일 경로 (로컬용)
        credentials_dict: 서비스 계정 JSON dict (Streamlit secrets용)
        spreadsheet_id:   구글 시트 ID
        """
        if not HAS_GSPREAD:
            raise ImportError("gspread 패키지가 필요합니다. pip install gspread")

        self.spreadsheet_id = spreadsheet_id

        if credentials_dict:
            creds = Credentials.from_service_account_info(credentials_dict,
                                                           scopes=SCOPES)
        elif credentials_path and os.path.exists(credentials_path):
            creds = Credentials.from_service_account_file(credentials_path,
                                                           scopes=SCOPES)
        else:
            raise FileNotFoundError("구글 서비스 계정 인증 정보가 없습니다.")

        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)

    # ═══════════════════════════════════════════
    #  시트 헬퍼
    # ═══════════════════════════════════════════

    def _get_or_create_sheet(self, title, rows=100, cols=20):
        """시트가 있으면 반환, 없으면 생성"""
        try:
            return self.spreadsheet.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            return self.spreadsheet.add_worksheet(title=title,
                                                   rows=rows, cols=cols)

    def _clear_and_write(self, ws, data):
        """시트를 지우고 데이터를 한번에 쓰기"""
        ws.clear()
        if data:
            ws.update(data, value_input_option="RAW")

    # ═══════════════════════════════════════════
    #  업로드 (DataManager → 구글 시트)
    # ═══════════════════════════════════════════

    def upload_all(self, dm):
        """DataManager의 전체 데이터를 구글 시트에 업로드"""

        # 1) 사업목록 시트
        ws_pj = self._get_or_create_sheet("사업목록")
        pj_data = [["사업명"]]
        for pn in dm.project_names:
            pj_data.append([pn])
        self._clear_and_write(ws_pj, pj_data)

        # 2) 사업별 시트
        for pn in dm.project_names:
            p = dm.projects[pn]
            ps = pn.replace("/", "_")

            # 계획
            ws_plan = self._get_or_create_sheet(f"{ps}_계획")
            plan_data = self._build_plan_data(p)
            self._clear_and_write(ws_plan, plan_data)

            # 집행내역
            ws_rec = self._get_or_create_sheet(f"{ps}_집행내역")
            rec_data = self._build_records_data(p)
            self._clear_and_write(ws_rec, rec_data)

            # 월별배분
            ws_mon = self._get_or_create_sheet(f"{ps}_월별배분")
            mon_data = self._build_monthly_data(p)
            self._clear_and_write(ws_mon, mon_data)

        # 불필요한 시트 정리 (Sheet1 등)
        self._cleanup_default_sheets()

    def _build_plan_data(self, p):
        rows = [["연도", "총예산", "편성목명", "편성목코드", "편성목예산",
                 "세부항목명", "세부항목예산"]]
        first = True
        for cat in p.categories:
            for item in cat["items"]:
                rows.append([
                    p.year if first else "",
                    p.total_budget if first else "",
                    cat["name"], cat["code"], cat["budget"],
                    item["name"], item["budget"]
                ])
                first = False
            if not cat["items"]:
                rows.append([
                    p.year if first else "",
                    p.total_budget if first else "",
                    cat["name"], cat["code"], cat["budget"], "", ""
                ])
                first = False
        return rows

    def _build_records_data(self, p):
        rows = [["ID", "집행일", "편성목", "세부항목", "세부내용",
                 "금액", "회차", "비고"]]
        for r in p.records:
            rows.append([r["id"], r["date"], r["cat"], r["item"],
                         r["detail"], r["amount"], r["round_"], r["memo"]])
        return rows

    def _build_monthly_data(self, p):
        rows = [["편성목명"] + [str(m) for m in range(1, 13)]]
        for cat in p.categories:
            md = p.monthly.get(cat["name"], {})
            rows.append([cat["name"]] + [md.get(m, 0) for m in range(1, 13)])
        return rows

    def _cleanup_default_sheets(self):
        """기본 Sheet1 삭제 (시트가 2개 이상일 때만)"""
        try:
            sheets = self.spreadsheet.worksheets()
            if len(sheets) > 1:
                for s in sheets:
                    if s.title in ("Sheet1", "시트1"):
                        self.spreadsheet.del_worksheet(s)
                        break
        except Exception:
            pass

    # ═══════════════════════════════════════════
    #  다운로드 (구글 시트 → DataManager)
    # ═══════════════════════════════════════════

    def download_all(self, dm):
        """구글 시트에서 전체 데이터를 DataManager로 로드"""

        dm.project_names = []
        dm.projects = {}

        # 1) 사업목록 읽기
        try:
            ws_pj = self.spreadsheet.worksheet("사업목록")
            rows = ws_pj.get_all_values()
            for row in rows[1:]:  # 헤더 건너뛰기
                pn = row[0].strip() if row else ""
                if pn:
                    dm.project_names.append(pn)
                    dm.projects[pn] = ProjectData()
        except gspread.exceptions.WorksheetNotFound:
            return False

        # 2) 사업별 데이터 읽기
        for pn in list(dm.project_names):
            p = dm.projects[pn]
            ps = pn.replace("/", "_")

            try:
                ws = self.spreadsheet.worksheet(f"{ps}_계획")
                self._load_plan_from_rows(p, ws.get_all_values())
            except gspread.exceptions.WorksheetNotFound:
                pass

            try:
                ws = self.spreadsheet.worksheet(f"{ps}_집행내역")
                self._load_records_from_rows(p, ws.get_all_values())
            except gspread.exceptions.WorksheetNotFound:
                pass

            try:
                ws = self.spreadsheet.worksheet(f"{ps}_월별배분")
                self._load_monthly_from_rows(p, ws.get_all_values())
            except gspread.exceptions.WorksheetNotFound:
                pass

        if dm.project_names:
            dm.current_project = dm.project_names[0]

        return True

    def _load_plan_from_rows(self, p, rows):
        if len(rows) < 2:
            return
        header = rows[0]
        p.categories = []
        cat_map = {}

        # 첫 데이터 행에서 연도/총예산
        first_row = rows[1] if len(rows) > 1 else []
        try:
            p.year = int(first_row[0]) if first_row[0] else p.year
        except (ValueError, IndexError):
            pass
        try:
            p.total_budget = int(float(str(first_row[1]).replace(",", ""))) if first_row[1] else 0
        except (ValueError, IndexError):
            pass

        for row in rows[1:]:
            cn = row[2].strip() if len(row) > 2 else ""
            cc = row[3].strip() if len(row) > 3 else ""
            cb = clean_num(row[4]) if len(row) > 4 else 0
            iname = row[5].strip() if len(row) > 5 else ""
            ibud = clean_num(row[6]) if len(row) > 6 else 0

            if cn and cn not in cat_map:
                cat = {"name": cn, "code": cc, "budget": cb, "items": []}
                p.categories.append(cat)
                cat_map[cn] = cat
            if cn in cat_map and iname:
                cat_map[cn]["items"].append({"name": iname, "budget": ibud})

    def _load_records_from_rows(self, p, rows):
        if len(rows) < 2:
            return
        p.records = []
        p._next_id = 1

        for row in rows[1:]:
            rid = row[0].strip() if len(row) > 0 else ""
            if not rid:
                continue
            try:
                rid_int = int(rid)
                if rid_int >= p._next_id:
                    p._next_id = rid_int + 1
            except ValueError:
                pass
            p.records.append({
                "id":     rid,
                "date":   row[1].strip() if len(row) > 1 else "",
                "cat":    row[2].strip() if len(row) > 2 else "",
                "item":   row[3].strip() if len(row) > 3 else "",
                "detail": row[4].strip() if len(row) > 4 else "",
                "amount": clean_num(row[5]) if len(row) > 5 else 0,
                "round_": row[6].strip() if len(row) > 6 else "",
                "memo":   row[7].strip() if len(row) > 7 else "",
            })

    def _load_monthly_from_rows(self, p, rows):
        if len(rows) < 2:
            return
        p.monthly = {}
        for row in rows[1:]:
            cn = row[0].strip() if len(row) > 0 else ""
            if not cn:
                continue
            p.monthly[cn] = {}
            for m in range(1, 13):
                p.monthly[cn][m] = clean_num(row[m]) if len(row) > m else 0

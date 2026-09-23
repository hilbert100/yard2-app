import io
import re
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------
# 0. 페이지 기본 설정
# ---------------------------------------------------------
st.set_page_config(
    page_title="제2야적장 입출고 관리",
    page_icon="📲",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------------------------------------------------
# 1. 시트 스키마 정의
#    Items 시트가 곧 "현재 상태"이고, 상태값(상태 컬럼)에 따라
#    재고현황(IN_STOCK) / 출고예정내역(PENDING_DISPATCH) / 출고내역(DISPATCHED)
#    화면에 나눠서 보여줍니다. 품번은 전체 시스템에서 영구히 유일합니다.
# ---------------------------------------------------------
ITEMS_HEADERS = ["품번", "종류", "규격", "위치", "상태", "입고일", "출고일", "배송지", "특기사항"]
ITEMS_COL = {name: i + 1 for i, name in enumerate(ITEMS_HEADERS)}  # 1-based 열 번호

LOG_HEADERS = ["품번", "종류", "규격", "위치", "입고일", "출고일", "배송지", "특기사항", "기록시각"]
CATEGORIES_HEADERS = ["종류"]
DESTINATIONS_HEADERS = ["배송지"]

STATUS_LABEL = {
    "IN_STOCK": "재고현황",
    "PENDING_DISPATCH": "출고예정내역",
    "DISPATCHED": "출고내역(출고완료)",
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


LEVEL_ALIASES = {"상": "상", "상단": "상", "위": "상", "top": "상", "하": "하", "하단": "하", "아래": "하", "bottom": "하"}


def normalize_part_no(raw):
    """품번 맨 앞의 고정 접두사 'F'를 자동으로 붙여준다.
    이미 F로 시작하면 그대로 두고, 없으면 앞에 붙인다. (예: 'B2-2B-004-3' -> 'FB2-2B-004-3')"""
    s = (raw or "").strip().upper()
    if not s:
        return s
    if not s.startswith("F"):
        s = "F" + s
    return s


def format_location_code(zone, row, col, level=""):
    """level이 비어있으면 기존과 동일하게 'A-01-01' (대부분의 1단 적재).
    상/하를 지정하면 'A-01-01-상' 처럼 뒤에 층 구분이 붙는다(2단 적재용)."""
    base = f"{str(zone).strip().upper()}-{int(row):02d}-{int(col):02d}"
    level_clean = LEVEL_ALIASES.get(str(level).strip().lower()) if level else ""
    # 위 딕셔너리는 소문자 top/bottom만 lower() 매칭되므로 한글은 원문 그대로도 시도
    if not level_clean and level:
        level_clean = LEVEL_ALIASES.get(str(level).strip(), "")
    return f"{base}-{level_clean}" if level_clean else base


LOCATION_INPUT_RE = re.compile(
    r"^\s*([A-Za-z가-힣]+)\s*[-,/\s]+\s*(\d+)\s*[-,/\s]+\s*(\d+)"
    r"(?:\s*[-,/\s]+\s*(상단|하단|상|하|위|아래|top|bottom))?\s*$",
    re.IGNORECASE,
)


def normalize_location_input(raw):
    """'A-01-01' 처럼 정확한 형식이 아니어도, 구역/열/행(+선택적으로 상/하 층)을
    공백·쉼표·슬래시·하이픈 아무 구분자로나 입력하면 자동 변환. 파싱 실패 시 None."""
    if not raw:
        return None
    m = LOCATION_INPUT_RE.match(str(raw))
    if not m:
        return None
    zone, row, col, level = m.groups()
    return format_location_code(zone, row, col, level or "")


# ---------------------------------------------------------
# 2. 구글 시트 연결 (연결 객체 자체만 캐시 — 데이터는 매번 새로 읽어 실시간성 유지)
# ---------------------------------------------------------
@st.cache_resource
def get_gsheet_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource
def get_spreadsheet():
    client = get_gsheet_client()
    return client.open_by_key(st.secrets["sheet_id"])


def _get_or_create_ws(ss, name, headers):
    try:
        ws = ss.worksheet(name)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=1000, cols=len(headers) + 2)
        ws.append_row(headers)
        return ws
    if ws.row_values(1) != headers:
        ws.update("A1", [headers])
    return ws


@st.cache_resource
def get_worksheets():
    """탭(Items/Categories/Destinations/DispatchLog) 확보 + 기본값 시딩.
    이 배포본에서 딱 한 번만 실행되어(모든 사용자/재실행이 공유) API 호출을 크게 아낍니다."""
    ss = get_spreadsheet()
    ws_items = _get_or_create_ws(ss, "Items", ITEMS_HEADERS)
    ws_categories = _get_or_create_ws(ss, "Categories", CATEGORIES_HEADERS)
    ws_destinations = _get_or_create_ws(ss, "Destinations", DESTINATIONS_HEADERS)
    ws_log = _get_or_create_ws(ss, "DispatchLog", LOG_HEADERS)
    if len(ws_categories.col_values(1)) <= 1:
        ws_categories.append_rows([["열연강판"], ["냉연강판"], ["후판"]])
    if len(ws_destinations.col_values(1)) <= 1:
        ws_destinations.append_rows([["본사 공장"], ["A현장"], ["하청1"]])
    return ws_items, ws_categories, ws_destinations, ws_log


def _cache_version():
    return st.session_state.get("_sheet_cache_version", 0)


def _bump_cache_version():
    """쓰기 작업 후 호출 — 캐시된 읽기 결과를 무효화해서 다음 조회에 최신값이 반영되게 함."""
    st.session_state["_sheet_cache_version"] = _cache_version() + 1


@st.cache_data(ttl=8, show_spinner=False)
def _cached_all_records(sheet_name, _version):
    ss = get_spreadsheet()
    ws = ss.worksheet(sheet_name)
    return ws.get_all_records()


@st.cache_data(ttl=8, show_spinner=False)
def _cached_col_values(sheet_name, col, _version):
    ss = get_spreadsheet()
    ws = ss.worksheet(sheet_name)
    return ws.col_values(col)


# ---------------------------------------------------------
# 3. 구글 시트 DB 엔진
# ---------------------------------------------------------
class SteelYardSheetDB:
    def __init__(self):
        self.ws_items, self.ws_categories, self.ws_destinations, self.ws_log = get_worksheets()

    # ---------------- Master 정보 (짧게 캐싱해서 반복 조회 시 API 호출 절약) ----------------
    def get_categories(self):
        vals = _cached_col_values("Categories", 1, _cache_version())
        return [v.strip() for v in vals[1:] if v.strip()]

    def add_category(self, name):
        name = name.strip()
        if not name:
            return
        if name in self.get_categories():
            raise ValueError(f"[{name}] 는 이미 등록된 종류입니다.")
        self.ws_categories.append_row([name])
        _bump_cache_version()

    def delete_category(self, name):
        try:
            cell = self.ws_categories.find(name, in_column=1)
        except Exception:
            cell = None
        if cell is None:
            raise ValueError(f"[{name}] 종류를 찾을 수 없습니다.")
        self.ws_categories.delete_rows(cell.row)
        _bump_cache_version()

    def get_destinations(self):
        vals = _cached_col_values("Destinations", 1, _cache_version())
        return [v.strip() for v in vals[1:] if v.strip()]

    def add_destination(self, name):
        name = name.strip()
        if not name:
            return
        if name in self.get_destinations():
            raise ValueError(f"[{name}] 는 이미 등록된 배송지입니다.")
        self.ws_destinations.append_row([name])
        _bump_cache_version()

    def delete_destination(self, name):
        try:
            cell = self.ws_destinations.find(name, in_column=1)
        except Exception:
            cell = None
        if cell is None:
            raise ValueError(f"[{name}] 배송지를 찾을 수 없습니다.")
        self.ws_destinations.delete_rows(cell.row)
        _bump_cache_version()

    # ---------------- 내부 헬퍼 ----------------
    def _items_df(self):
        records = _cached_all_records("Items", _cache_version())
        df = pd.DataFrame(records)
        if df.empty:
            return pd.DataFrame(columns=ITEMS_HEADERS)
        for col in ITEMS_HEADERS:
            if col not in df.columns:
                df[col] = ""
        return df[ITEMS_HEADERS].astype(str).replace("nan", "")

    def _find_item_row(self, part_no):
        """품번으로 Items 시트에서 실제 행 번호(헤더 포함, 1-based)를 찾음. 없으면 None."""
        cell = self.ws_items.find(part_no, in_column=ITEMS_COL["품번"])
        if cell is None:
            return None
        return cell.row

    # ---------------- 1단계: 입고 -> 재고현황(IN_STOCK) ----------------
    def register_inbound(self, part_no, category_name, spec, zone, row, col, inbound_date, remarks, level=""):
        clean_part_no = normalize_part_no(part_no)
        loc_code = format_location_code(zone, row, col, level)

        if category_name not in self.get_categories():
            raise ValueError(f"[{category_name}] 는 등록되지 않은 종류입니다. Master 관리에서 먼저 추가해주세요.")

        # 품번은 영구 고유번호 -> 상태 불문하고 이미 존재하면 거부
        existing_row = self._find_item_row(clean_part_no)
        if existing_row:
            status = self.ws_items.cell(existing_row, ITEMS_COL["상태"]).value
            raise ValueError(
                f"[{clean_part_no}] 품번은 이미 사용된 고유번호입니다. "
                f"(현재 상태: {STATUS_LABEL.get(status, status)}) 품번을 다시 확인해 주세요."
            )

        df = self._items_df()
        active = df[df["상태"].isin(["IN_STOCK", "PENDING_DISPATCH"])] if not df.empty else df
        if not active.empty and (active["위치"] == loc_code).any():
            raise ValueError(f"[{loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다. (2단 적재라면 상/하를 다르게 지정해주세요)")

        clean_spec = spec.strip().upper() if spec else ""
        self.ws_items.append_row([
            clean_part_no, category_name, clean_spec, loc_code, "IN_STOCK",
            str(inbound_date), "", "", remarks or ""
        ])
        _bump_cache_version()

    def bulk_register_inbound(self, rows):
        """일괄 업로드 전용. rows: [{part_no, category_name, spec, zone, row, col,
        inbound_date, remarks, level}, ...]
        구글 시트를 딱 한 번만 읽어 전체를 메모리에서 검증한 뒤, 유효한 항목을
        한 번의 append_rows 호출로 몰아서 저장한다 (API 호출 수를 줄여 할당량 초과 방지).
        반환: [(품번, 성공여부, 사유), ...] — rows와 같은 순서."""
        existing_df = self._items_df()
        existing_parts = set(existing_df["품번"].tolist()) if not existing_df.empty else set()
        active_locations = set()
        if not existing_df.empty:
            active_locations = set(
                existing_df[existing_df["상태"].isin(["IN_STOCK", "PENDING_DISPATCH"])]["위치"].tolist()
            )
        valid_categories = set(self.get_categories())

        results = []
        new_sheet_rows = []
        seen_parts = set()
        seen_locations = set()

        for r in rows:
            try:
                clean_part_no = normalize_part_no(r.get("part_no", ""))
                if not clean_part_no:
                    raise ValueError("품번이 비어있습니다.")
                if clean_part_no in existing_parts or clean_part_no in seen_parts:
                    raise ValueError(f"[{clean_part_no}] 품번은 이미 사용된 고유번호입니다.")

                category_name = r.get("category_name", "")
                if category_name not in valid_categories:
                    raise ValueError(f"[{category_name}] 는 등록되지 않은 종류입니다. Master 관리에서 먼저 추가해주세요.")

                loc_code = format_location_code(r["zone"], r["row"], r["col"], r.get("level", ""))
                if loc_code in active_locations or loc_code in seen_locations:
                    raise ValueError(f"[{loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다.")

                clean_spec = (r.get("spec") or "").strip().upper()
                new_sheet_rows.append([
                    clean_part_no, category_name, clean_spec, loc_code, "IN_STOCK",
                    str(r["inbound_date"]), "", "", r.get("remarks") or ""
                ])
                seen_parts.add(clean_part_no)
                seen_locations.add(loc_code)
                results.append((clean_part_no, True, ""))
            except Exception as e:
                results.append((r.get("part_no", "-"), False, str(e)))

        if new_sheet_rows:
            self.ws_items.append_rows(new_sheet_rows)
            _bump_cache_version()

        return results

    # ---------------- 적재위치 변경 ----------------
    def update_location(self, part_no, new_zone, new_row, new_col, level=""):
        clean_part_no = normalize_part_no(part_no)
        new_loc_code = format_location_code(new_zone, new_row, new_col, level)

        row_idx = self._find_item_row(clean_part_no)
        if not row_idx:
            raise ValueError("존재하지 않는 품번입니다.")

        df = self._items_df()
        active = df[(df["상태"].isin(["IN_STOCK", "PENDING_DISPATCH"])) & (df["품번"] != clean_part_no)]
        if not active.empty and (active["위치"] == new_loc_code).any():
            raise ValueError(f"[{new_loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다.")

        self.ws_items.update_cell(row_idx, ITEMS_COL["위치"], new_loc_code)
        _bump_cache_version()

    # ---------------- 2단계: 재고현황 체크 -> 출고예정(PENDING_DISPATCH) ----------------
    def move_to_pending_dispatch(self, part_numbers):
        for p in part_numbers:
            row_idx = self._find_item_row(p)
            if row_idx:
                self.ws_items.update_cell(row_idx, ITEMS_COL["상태"], "PENDING_DISPATCH")
        _bump_cache_version()

    # ---------------- 3단계: 출고확정 -> 출고내역(DISPATCHED) ----------------
    def confirm_outbound(self, part_no, outbound_date, dest_name, remarks):
        clean_part_no = normalize_part_no(part_no)
        row_idx = self._find_item_row(clean_part_no)
        if not row_idx:
            raise ValueError("존재하지 않는 품번입니다.")

        row_vals = self.ws_items.row_values(row_idx)
        row_vals += [""] * (len(ITEMS_HEADERS) - len(row_vals))

        existing_remarks = row_vals[ITEMS_COL["특기사항"] - 1]
        final_remarks = existing_remarks
        if remarks:
            final_remarks = f"{existing_remarks} / [출고비고] {remarks}" if existing_remarks else remarks

        inbound_date_val = row_vals[ITEMS_COL["입고일"] - 1]

        # 상태(E) ~ 특기사항(I)까지 한 번에 갱신 (API 호출 최소화)
        self.ws_items.update(
            f"E{row_idx}:I{row_idx}",
            [["DISPATCHED", inbound_date_val, str(outbound_date), dest_name, final_remarks]]
        )

        # 출고 이력을 별도 로그에 영구 기록
        self.ws_log.append_row([
            clean_part_no,
            row_vals[ITEMS_COL["종류"] - 1],
            row_vals[ITEMS_COL["규격"] - 1],
            row_vals[ITEMS_COL["위치"] - 1],
            inbound_date_val,
            str(outbound_date),
            dest_name,
            final_remarks,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
        _bump_cache_version()

    # ---------------- 4단계: 보류 체크 -> 재고현황(IN_STOCK) 복원 ----------------
    def restore_to_in_stock(self, part_numbers):
        if isinstance(part_numbers, str):
            part_numbers = [part_numbers]
        for p in part_numbers:
            row_idx = self._find_item_row(p)
            if row_idx:
                self.ws_items.update_cell(row_idx, ITEMS_COL["상태"], "IN_STOCK")
        _bump_cache_version()

    # ---------------- 재고현황 표에서 셀 직접 수정 ----------------
    def update_item(self, original_part_no, new_part_no, category_name, spec, location_code, remarks):
        """재고현황 표에서 종류/품번/규격/위치/특기사항을 한 번에 수정."""
        row_idx = self._find_item_row(original_part_no)
        if not row_idx:
            raise ValueError(f"[{original_part_no}] 품번을 찾을 수 없습니다.")

        new_part_no_clean = normalize_part_no(new_part_no)
        if not new_part_no_clean:
            raise ValueError("품번은 비워둘 수 없습니다.")

        if new_part_no_clean != original_part_no:
            existing_row = self._find_item_row(new_part_no_clean)
            if existing_row:
                raise ValueError(f"[{new_part_no_clean}] 품번은 이미 다른 곳에서 사용 중입니다.")

        raw_loc = (location_code or "").strip()
        if not raw_loc:
            raise ValueError("위치는 비워둘 수 없습니다.")

        loc_code = normalize_location_input(raw_loc)
        if loc_code is None:
            raise ValueError(
                f"위치 형식을 알아볼 수 없습니다: '{raw_loc}'. "
                "구역과 열, 행을 순서대로 입력해주세요 (예: A 1 1, A-1-1, A,1,1)."
            )

        df = self._items_df()
        active = df[
            (df["상태"].isin(["IN_STOCK", "PENDING_DISPATCH"])) & (df["품번"] != original_part_no)
        ]
        if not active.empty and (active["위치"] == loc_code).any():
            raise ValueError(f"[{loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다.")

        clean_spec = (spec or "").strip().upper()
        clean_remarks = remarks or ""

        # 품번(A)~위치(D), 특기사항(I) 갱신
        self.ws_items.update(
            f"A{row_idx}:D{row_idx}",
            [[new_part_no_clean, category_name, clean_spec, loc_code]],
        )
        self.ws_items.update_cell(row_idx, ITEMS_COL["특기사항"], clean_remarks)
        _bump_cache_version()

    # ---------------- 품목 삭제 (재고현황 / 출고내역 공통) ----------------
    def delete_item(self, part_no):
        """품번을 시스템에서 완전히 삭제 (상태 무관). 출고 이력(DispatchLog)은 영향받지 않음."""
        row_idx = self._find_item_row(part_no)
        if not row_idx:
            raise ValueError(f"[{part_no}] 품번을 찾을 수 없습니다.")
        self.ws_items.delete_rows(row_idx)
        _bump_cache_version()

    # ---------------- 조회 ----------------
    def search_current_stock(self, category_name="전체", part_kw="", spec_kw=""):
        df = self._items_df()
        cols = ["종류", "품번", "규격", "적재위치", "입고일", "특기사항"]
        if df.empty:
            return pd.DataFrame(columns=cols)

        df = df[df["상태"] == "IN_STOCK"].copy()
        if category_name and category_name != "전체":
            df = df[df["종류"] == category_name]
        if part_kw:
            df = df[df["품번"].str.contains(part_kw.strip().upper(), na=False, regex=False)]
        if spec_kw:
            df = df[df["규격"].str.contains(spec_kw.strip().upper(), na=False, regex=False)]

        df = df.rename(columns={"위치": "적재위치"})
        if df.empty:
            return pd.DataFrame(columns=cols)
        return df[cols].sort_values(["종류", "적재위치", "품번"]).reset_index(drop=True)

    def search_pending_dispatch(self):
        df = self._items_df()
        cols = ["종류", "품번", "규격", "적재위치", "입고일", "특기사항"]
        if df.empty:
            return pd.DataFrame(columns=cols)
        df = df[df["상태"] == "PENDING_DISPATCH"].copy().rename(columns={"위치": "적재위치"})
        if df.empty:
            return pd.DataFrame(columns=cols)
        return df[cols].sort_values("품번").reset_index(drop=True)

    def search_dispatched(self):
        df = self._items_df()
        cols = ["종류", "품번", "규격", "입고일", "출고일", "배송지", "특기사항"]
        if df.empty:
            return pd.DataFrame(columns=cols)
        df = df[df["상태"] == "DISPATCHED"].copy()
        if df.empty:
            return pd.DataFrame(columns=cols)
        return df[cols].sort_values(["출고일", "품번"], ascending=[False, True]).reset_index(drop=True)


# ---------------------------------------------------------
# 엑셀 다운로드 생성기
# ---------------------------------------------------------
def render_html_table(df, checkbox_col=None):
    """조회 전용 화면(이용자)에서 쓰는, 목업과 같은 스타일의 컴팩트한 표."""
    if df.empty:
        return
    cols = [c for c in df.columns if c != checkbox_col]
    header_html = "".join(
        f'<th style="text-align:center;padding:5px 8px;color:var(--text-muted);'
        f'border-bottom:0.5px solid var(--border);white-space:nowrap;">{c}</th>'
        for c in cols
    )
    rows_html = ""
    for _, row in df.iterrows():
        cells = "".join(
            f'<td style="padding:6px 8px;border-bottom:0.5px solid var(--border);'
            f'white-space:nowrap;text-align:center;">{row[c] if str(row[c]).strip() else "-"}</td>'
            for c in cols
        )
        rows_html += f"<tr>{cells}</tr>"

    html = f"""
    <div style="background:var(--surface-2);border:0.5px solid var(--border);
                border-radius:12px;padding:10px;">
      <div style="overflow-x:auto;">
        <table style="border-collapse:collapse;font-size:12px;width:100%;">
          <thead><tr>{header_html}</tr></thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div>
      <div style="font-size:10px;color:var(--text-muted);margin-top:8px;">
        ← 좌우로 스와이프하면 모든 컬럼을 볼 수 있습니다
      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def generate_excel_report(df, title="재고현황"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=title)
        worksheet = writer.sheets[title]

        header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        data_font = Font(name="맑은 고딕", size=10)
        border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")

        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center

        for row_num in range(2, len(df) + 2):
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.font = data_font
                cell.border = border
                cell.alignment = align_center if col_num <= 7 else align_left

        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 6, 12)

    output.seek(0)
    return output


# ---------------------------------------------------------
# 4. 컨트롤러 (Streamlit UI)
# ---------------------------------------------------------
try:
    db = SteelYardSheetDB()
except Exception as e:
    st.error(
        "❌ 구글 시트 연결에 실패했습니다. secrets 설정(gcp_service_account, sheet_id)과 "
        "시트 공유 권한(서비스 계정 이메일에 편집자 권한 공유)을 확인해 주세요."
    )
    st.exception(e)
    st.stop()

# 모바일에서 종류/품번/규격/위치/입고일/특기사항/체크박스가
# 가로 스크롤 없이 최대한 한 화면에 들어오도록 표 폰트와 셀 여백을 압축
st.markdown("""
<style>
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
    font-size: 12px;
}
[data-testid="stDataFrame"] div[role="gridcell"],
[data-testid="stDataEditor"] div[role="gridcell"] {
    padding-left: 4px !important;
    padding-right: 4px !important;
    font-size: 12px !important;
    justify-content: center !important;
    text-align: center !important;
}
[data-testid="stDataFrame"] div[role="columnheader"],
[data-testid="stDataEditor"] div[role="columnheader"] {
    padding-left: 4px !important;
    padding-right: 4px !important;
    font-size: 12px !important;
    font-weight: 600;
    justify-content: center !important;
    text-align: center !important;
}
[data-testid="stDataFrame"] div[role="columnheader"] > div,
[data-testid="stDataEditor"] div[role="columnheader"] > div {
    justify-content: center !important;
    text-align: center !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    """
    <div style="display:flex; align-items:baseline; gap:8px; white-space:nowrap;
                overflow:hidden; margin:0 0 12px 0;">
        <h1 style="font-size:clamp(15px, 5vw, 26px); text-overflow:ellipsis;
                   overflow:hidden; margin:0;">
            📲 제2야적장 입출고 관리
        </h1>
        <span style="font-family:'Pretendard','Noto Sans KR',sans-serif;
                     font-size:clamp(10px, 2.8vw, 14px); font-weight:600;
                     color:#1B3A6B; flex-shrink:0;">
            서진로지스
        </span>
    </div>
    """,
    unsafe_allow_html=True,
)

try:
    ADMIN_PIN = st.secrets["admin_pin"]
except Exception:
    ADMIN_PIN = "1234"  # secrets.toml 미설정 시 임시값 — 운영 전 반드시 교체

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

c_role, c_pin = st.columns([2, 3])
role_selection = c_role.radio("사용자 권한", ["이용자 (조회만)", "관리자"], horizontal=True)

if role_selection == "관리자":
    if not st.session_state.is_admin:
        with c_pin.form("admin_login_form", clear_on_submit=False):
            input_pin = st.text_input("🔑 관리자 비밀번호", type="password", placeholder="비밀번호 입력")
            submitted = st.form_submit_button("로그인")
        if submitted:
            if input_pin == ADMIN_PIN:
                st.session_state.is_admin = True
                st.success("🔓 관리자 인증 성공!")
                st.rerun()
            else:
                st.error("❌ 비밀번호가 올바르지 않습니다.")
    else:
        if c_pin.button("🔒 관리자 로그아웃"):
            st.session_state.is_admin = False
            st.rerun()
else:
    st.session_state.is_admin = False

is_admin = st.session_state.is_admin

tab_in, tab_stock, tab_pending, tab_history, tab_manage = st.tabs(
    ["📥 입고 입력", "🔍 실시간 재고현황", "🚚 출고예정내역", "📜 출고내역", "⚙️ Master 관리"]
)

# TAB 1: 입고 등록
with tab_in:
    if not is_admin:
        st.warning("🔒 입고 입력은 관리자 권한이 필요합니다. 상단에서 관리자 비밀번호를 인증해 주세요.")
    else:
        st.subheader("📥 현장 즉시 입고 등록")
        with st.form("realtime_in_form", clear_on_submit=True):
            col_cat, col_part, col_spec, col_loc, col_date, col_rem = st.columns([1.5, 3.5, 2, 2.5, 1.5, 2])

            categories = db.get_categories()
            if not categories:
                st.error("등록된 강판 종류가 없습니다. Master 관리 탭에서 추가해 주세요.")
            else:
                with col_cat:
                    cat_name = st.selectbox("종류", categories)

                with col_part:
                    st.markdown("**품번 (4단 분할, 맨 앞 F는 자동으로 붙습니다)**")
                    p1, p2, p3, p4 = st.columns(4)
                    part_1 = p1.text_input("1단", placeholder="B2", label_visibility="collapsed")
                    part_2 = p2.text_input("2단", placeholder="2B", label_visibility="collapsed")
                    part_3 = p3.text_input("3단", placeholder="004", label_visibility="collapsed")
                    part_4 = p4.text_input("4단", placeholder="3", label_visibility="collapsed")

                with col_spec:
                    spec = st.text_input("규격", placeholder="예: 12T x 1500 x 6000")

                with col_loc:
                    st.markdown("**적재위치 (2단 적재 시 층 선택)**")
                    z_col, r_col, c_col, lv_col = st.columns(4)
                    zone = z_col.text_input("구역", value="A", label_visibility="collapsed")
                    row_n = r_col.number_input("열", min_value=1, value=1, label_visibility="collapsed")
                    col_n = c_col.number_input("행", min_value=1, value=1, label_visibility="collapsed")
                    level = lv_col.selectbox("층", ["", "상", "하"], label_visibility="collapsed")

                with col_date:
                    in_date = st.date_input("입고일", datetime.now())

                with col_rem:
                    remarks = st.text_input("특기사항", placeholder="예: 측면 스크래치")

                st.divider()

                if st.form_submit_button("⚡ 실시간 입고 등록", use_container_width=True):
                    part_components = [p.strip() for p in [part_1, part_2, part_3, part_4] if p.strip()]
                    if len(part_components) < 4:
                        st.error("품번 4단 칸을 모두 입력해주세요.")
                    else:
                        full_part_no = "-".join(part_components).upper()
                        try:
                            db.register_inbound(full_part_no, cat_name, spec, zone, row_n, col_n, in_date, remarks, level)
                            saved_part_no = normalize_part_no(full_part_no)
                            saved_loc = format_location_code(zone, row_n, col_n, level)
                            st.success(f"✅ [{saved_part_no}] 입고 등록 완료! (위치: {saved_loc})")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 등록 실패: {e}")

        st.divider()
        st.subheader("📤 엑셀 일괄 업로드")
        st.caption(
            "정해진 양식(종류 / 품번1~4단 / 규격 / 구역·열·행·층 / 입고일 / 특기사항)의 "
            "엑셀 파일을 올리면 한 줄씩 자동으로 입고 등록합니다."
        )

        uploaded_file = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=["xlsx"], key="bulk_upload")

        if uploaded_file is not None:
            try:
                # dtype=str: '004' 같은 앞자리 0이 있는 품번/숫자 칸이 자동으로 숫자로
                # 변환되어 0이 사라지는 것을 방지 (모든 칸을 텍스트 그대로 읽음)
                upload_df = pd.read_excel(
                    uploaded_file, sheet_name="입고데이터", header=0, engine="openpyxl", dtype=str
                )
            except Exception as e:
                st.error(f"❌ 엑셀 파일을 읽을 수 없습니다: {e}. [입고데이터] 시트가 있는지 확인해주세요.")
                upload_df = None

            if upload_df is not None:
                upload_df = upload_df.dropna(how="all")
                st.write(f"총 {len(upload_df)}줄을 읽었습니다.")

                if st.button("⚡ 일괄 업로드 실행", type="primary", use_container_width=True):
                    # 1단계: 엑셀 형식만 미리 정리 (아직 구글 시트에 접근하지 않음 — API 호출 없음)
                    skip_results = []  # (excel_row_no, "-", "⚠️ 건너뜀", 사유)
                    shape_error_results = []  # (excel_row_no, 품번, "❌ 실패", 사유) — 열/행 숫자 오류 등
                    candidate_rows = []  # bulk_register_inbound에 넘길 유효 후보들
                    row_no_by_part = {}  # 결과를 다시 엑셀 행 번호와 매칭하기 위한 매핑

                    for i, row in upload_df.iterrows():
                        excel_row_no = i + 3  # 헤더(1) + 예시(2) 다음부터 시작하는 실제 엑셀 행 번호

                        def g(col):
                            val = row.get(col)
                            return "" if pd.isna(val) else str(val).strip()

                        cat = g("종류")
                        p1, p2, p3, p4 = g("품번1단"), g("품번2단"), g("품번3단"), g("품번4단")
                        spec = g("규격")
                        zone = g("구역")
                        row_n_raw = g("열")
                        col_n_raw = g("행")
                        level = g("층(선택)")
                        remarks = g("특기사항")

                        part_components = [p for p in [p1, p2, p3, p4] if p]
                        if not (cat and len(part_components) == 4 and zone and row_n_raw and col_n_raw):
                            if cat or p1 or p2 or p3 or p4 or zone or row_n_raw or col_n_raw:
                                skip_results.append((excel_row_no, "-", "⚠️ 건너뜀", "필수 항목(종류/품번1~4단/구역/열/행) 중 비어있는 칸이 있습니다."))
                            continue

                        full_part_no = "-".join(part_components).upper()

                        try:
                            row_n = int(float(row_n_raw))
                            col_n = int(float(col_n_raw))
                        except ValueError:
                            shape_error_results.append((excel_row_no, full_part_no, "❌ 실패", f"열/행은 숫자여야 합니다 (입력값: {row_n_raw}, {col_n_raw})"))
                            continue

                        in_date_val = row.get("입고일")
                        if pd.isna(in_date_val) or str(in_date_val).strip() == "":
                            in_date = datetime.now().date()
                        else:
                            parsed = pd.to_datetime(in_date_val, errors="coerce")
                            in_date = parsed.date() if not pd.isna(parsed) else datetime.now().date()

                        candidate_rows.append({
                            "part_no": full_part_no, "category_name": cat, "spec": spec,
                            "zone": zone, "row": row_n, "col": col_n,
                            "inbound_date": in_date, "remarks": remarks, "level": level,
                        })
                        row_no_by_part[normalize_part_no(full_part_no)] = excel_row_no

                    # 2단계: 구글 시트는 딱 두 번만 호출 (읽기 1회 + 쓰기 1회)해서 후보들을 실제로 등록
                    bulk_results = db.bulk_register_inbound(candidate_rows) if candidate_rows else []

                    final_results = list(skip_results) + list(shape_error_results)
                    for part_no, ok, msg in bulk_results:
                        normalized = normalize_part_no(part_no)
                        excel_row_no = row_no_by_part.get(normalized, "-")
                        final_results.append((excel_row_no, normalized, "✅ 성공" if ok else "❌ 실패", "" if ok else msg))

                    final_results.sort(key=lambda r: (r[0] if isinstance(r[0], int) else 9999))

                    success_count = sum(1 for r in final_results if r[2] == "✅ 성공")
                    fail_count = sum(1 for r in final_results if r[2] == "❌ 실패")
                    skip_count = sum(1 for r in final_results if r[2] == "⚠️ 건너뜀")

                    st.success(f"✅ 성공 {success_count}건 · ❌ 실패 {fail_count}건 · ⚠️ 건너뜀 {skip_count}건")

                    if final_results:
                        result_df = pd.DataFrame(final_results, columns=["엑셀 행", "품번", "결과", "사유"])
                        st.dataframe(result_df, use_container_width=True, hide_index=True)

                    if success_count:
                        st.info("성공한 항목은 [실시간 재고현황] 탭에서 확인하실 수 있습니다.")

# TAB 2: 실시간 재고현황
with tab_stock:
    st.subheader("🔍 실시간 재고현황")

    st.markdown("##### 🎯 조건별 검색")
    sf_col1, sf_col2, sf_col3 = st.columns([1.5, 2, 2])

    cat_list = ["전체"] + db.get_categories()
    sel_cat = sf_col1.selectbox("종류 선택", cat_list)
    sel_part = sf_col2.text_input("품번 검색", placeholder="예: FA-3B 또는 01")
    sel_spec = sf_col3.text_input("규격 검색", placeholder="예: 12T 또는 1500")

    stock_df = db.search_current_stock(category_name=sel_cat, part_kw=sel_part, spec_kw=sel_spec)
    st.metric("조회된 재고 수량", f"{len(stock_df)}건")

    if not stock_df.empty:
        excel_data = generate_excel_report(stock_df, title="실시간재고현황")
        st.download_button(
            "📊 현재 조회 내역 엑셀 다운로드 (.xlsx)",
            data=excel_data,
            file_name=f"실시간_재고현황_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.divider()

        if is_admin:
            with st.expander("🔄 [니구리 작업] 적재위치 수시 변경 (클릭)", expanded=False):
                with st.form("loc_update_form", clear_on_submit=True):
                    u_col1, u_col2, u_col3, u_col4, u_col5 = st.columns([2.2, 1, 1, 1, 1])
                    u_part = u_col1.selectbox("위치 변경 대상 품번", stock_df["품번"].tolist())
                    u_zone = u_col2.text_input("신규 구역", value="A")
                    u_row = u_col3.number_input("신규 열", min_value=1, value=1)
                    u_col = u_col4.number_input("신규 행", min_value=1, value=1)
                    u_level = u_col5.selectbox("층", ["", "상", "하"])
                    if st.form_submit_button("⚡ 적재위치 변경 저장", use_container_width=True):
                        try:
                            db.update_location(u_part, u_zone, u_row, u_col, u_level)
                            new_loc = format_location_code(u_zone, u_row, u_col, u_level)
                            st.success(f"✅ [{u_part}] 위치가 [{new_loc}](으)로 수정되었습니다!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 위치 변경 실패: {e}")

            st.info(
                "💡 **셀을 탭해서 값을 고치고 다른 칸으로 넘어가면(또는 Enter) 바로 저장**됩니다 — "
                "따로 저장 버튼을 안 누르셔도 됩니다. "
                "품번 맨 앞의 F는 안 쓰셔도 저장할 때 자동으로 붙습니다. "
                "위치는 정확한 형식(A-01-01) 없이 구역·열·행을 순서대로 입력하시면 "
                "자동으로 하이픈이 붙습니다 (예: `A 1 1`, `A,1,1`, `A-1-1` 모두 A-01-01로 저장됨). "
                "2단 적재라면 맨 뒤에 상 또는 하를 추가로 입력하세요 (예: `A 1 1 상` → A-01-01-상). "
                "보통은 비워두시면 됩니다. "
                "출고 예정으로 보낼 품번은 **[출고]**, 완전히 지울 품번은 **[삭제]** 체크박스를 체크한 뒤 아래 버튼을 눌러주세요."
            )

            stock_df_display = stock_df.reset_index(drop=True).copy()
            stock_df_display["출고 선택"] = False
            stock_df_display["삭제 체크"] = False

            cat_options = sorted(set(db.get_categories()) | set(stock_df_display["종류"].unique().tolist()))

            def _auto_save_stock_edits():
                """종류/품번/규격/위치/특기사항 셀을 고치고 포커스를 벗어나는 즉시(별도 버튼 없이) 저장."""
                editor_state = st.session_state.get("editor_stock", {})
                edited_rows = editor_state.get("edited_rows", {})
                if not edited_rows:
                    return
                field_cols = ["종류", "품번", "규격", "적재위치", "특기사항"]
                successes, errors = [], []
                for idx_str, changes in edited_rows.items():
                    if not any(k in changes for k in field_cols):
                        continue  # 출고 선택/삭제 체크만 바뀐 경우는 여기서 처리 안 함
                    idx = int(idx_str)
                    orig = stock_df_display.loc[idx]
                    try:
                        db.update_item(
                            original_part_no=orig["품번"],
                            new_part_no=changes.get("품번", orig["품번"]),
                            category_name=changes.get("종류", orig["종류"]),
                            spec=changes.get("규격", orig["규격"]),
                            location_code=changes.get("적재위치", orig["적재위치"]),
                            remarks=changes.get("특기사항", orig["특기사항"]),
                        )
                        successes.append(orig["품번"])
                    except Exception as e:
                        errors.append(f"[{orig['품번']}] {e}")
                st.session_state["_stock_edit_feedback"] = (successes, errors)

            edited_df = st.data_editor(
                stock_df_display,
                use_container_width=True,
                hide_index=True,
                disabled=["입고일"],
                column_config={
                    "종류": st.column_config.SelectboxColumn("종류", width="small", options=cat_options),
                    "품번": st.column_config.TextColumn("품번", width="large"),
                    "규격": st.column_config.TextColumn("규격", width="small"),
                    "적재위치": st.column_config.TextColumn("위치", width="small"),
                    "입고일": st.column_config.DateColumn("입고일", width="small", format="MM/DD"),
                    "특기사항": st.column_config.TextColumn("특기사항", width="small"),
                    "출고 선택": st.column_config.CheckboxColumn("출고", width="small", default=False),
                    "삭제 체크": st.column_config.CheckboxColumn("삭제", width="small", default=False),
                },
                key="editor_stock",
                on_change=_auto_save_stock_edits,
            )

            _feedback = st.session_state.pop("_stock_edit_feedback", None)
            if _feedback:
                _successes, _errors = _feedback
                if _successes:
                    st.success(f"✅ 자동 저장됨: {', '.join(_successes)}")
                for _err in _errors:
                    st.error(f"❌ {_err}")

            col_move, col_del = st.columns(2)

            with col_move:
                selected_parts = edited_df[edited_df["출고 선택"] == True]["품번"].tolist()
                if st.button("🚚 선택 항목 출고예정으로 이동", type="primary", use_container_width=True):
                    if not selected_parts:
                        st.warning("출고 예정으로 이동할 품번을 1개 이상 선택해주세요.")
                    else:
                        db.move_to_pending_dispatch(selected_parts)
                        st.success(f"✅ 총 {len(selected_parts)}건이 [출고예정내역]으로 이동되었습니다: {', '.join(selected_parts)}")
                        st.rerun()

            with col_del:
                delete_parts = edited_df[edited_df["삭제 체크"] == True]["품번"].tolist()
                if st.button("🗑️ 선택 항목 삭제", use_container_width=True, key="delete_stock_btn"):
                    if not delete_parts:
                        st.warning("삭제할 품번을 1개 이상 선택해주세요.")
                    else:
                        del_errors = []
                        del_success = []
                        for p in delete_parts:
                            try:
                                db.delete_item(p)
                                del_success.append(p)
                            except Exception as e:
                                del_errors.append(f"[{p}] {e}")
                        if del_success:
                            st.success(f"🗑️ 총 {len(del_success)}건 삭제 완료: {', '.join(del_success)}")
                        for err in del_errors:
                            st.error(f"❌ {err}")
                        if del_success:
                            st.rerun()
        else:
            render_html_table(stock_df)
    else:
        st.info("조건에 일치하는 재고가 없습니다.")

# TAB 3: 출고예정내역
with tab_pending:
    st.subheader("🚚 출고예정내역")

    pending_df = db.search_pending_dispatch()
    st.metric("출고 대기 수량", f"{len(pending_df)}건")

    if not pending_df.empty:
        if not is_admin:
            render_html_table(pending_df)
            st.info("💡 보류·확정 처리는 관리자 권한이 필요합니다.")
        else:
            destinations = db.get_destinations()
            if not destinations:
                st.error("등록된 배송지가 없습니다. Master 관리 탭에서 추가해 주세요.")
            else:
                st.info(
                    "💡 **보류**를 체크하면 [실시간 재고현황]으로 복원되고, **확정**을 체크하면 "
                    "옆의 배송지로 [출고내역]으로 이동합니다 (출고일자는 오늘 날짜로 자동 입력). "
                    "체크 후 아래 [적용] 버튼을 눌러주세요."
                )

                pending_df_display = pending_df.reset_index(drop=True).copy()
                pending_df_display["보류 체크"] = False
                pending_df_display["확정 체크"] = False
                pending_df_display["배송지"] = destinations[0]

                edited_pending_df = st.data_editor(
                    pending_df_display,
                    use_container_width=True,
                    hide_index=True,
                    disabled=["종류", "품번", "규격", "적재위치", "입고일", "특기사항"],
                    column_config={
                        "종류": st.column_config.TextColumn("종류", width="small"),
                        "품번": st.column_config.TextColumn("품번", width="large"),
                        "규격": st.column_config.TextColumn("규격", width="small"),
                        "적재위치": st.column_config.TextColumn("위치", width="small"),
                        "입고일": st.column_config.DateColumn("입고일", width="small", format="MM/DD"),
                        "특기사항": st.column_config.TextColumn("특기사항", width="small"),
                        "보류 체크": st.column_config.CheckboxColumn("보류", width="small", default=False),
                        "확정 체크": st.column_config.CheckboxColumn("확정", width="small", default=False),
                        "배송지": st.column_config.SelectboxColumn("배송지", width="small", options=destinations),
                    },
                    key="editor_pending",
                )

                if st.button("⚡ 적용", type="primary", use_container_width=True):
                    hold_rows = edited_pending_df[edited_pending_df["보류 체크"] == True]
                    confirm_rows = edited_pending_df[
                        (edited_pending_df["확정 체크"] == True) & (edited_pending_df["보류 체크"] == False)
                    ]

                    if hold_rows.empty and confirm_rows.empty:
                        st.warning("보류 또는 확정으로 체크된 품번이 없습니다.")
                    else:
                        today = datetime.now().date()

                        if not hold_rows.empty:
                            hold_parts = hold_rows["품번"].tolist()
                            db.restore_to_in_stock(hold_parts)
                            st.success(f"↩️ 총 {len(hold_parts)}건 보류 → [실시간 재고현황]으로 복원되었습니다: {', '.join(hold_parts)}")

                        confirm_errors = []
                        confirm_success = []
                        for _, row in confirm_rows.iterrows():
                            try:
                                db.confirm_outbound(row["품번"], today, row["배송지"], "")
                                confirm_success.append(row["품번"])
                            except Exception as e:
                                confirm_errors.append(f"[{row['품번']}] {e}")

                        if confirm_success:
                            st.success(f"🎉 총 {len(confirm_success)}건 확정 → [출고내역]으로 이동되었습니다: {', '.join(confirm_success)}")
                        for err in confirm_errors:
                            st.error(f"❌ {err}")

                        st.rerun()
    else:
        st.info("현재 출고 예정인 내역이 없습니다.")

# TAB 4: 출고내역
with tab_history:
    st.subheader("📜 출고 완료 내역")

    dispatched_df = db.search_dispatched()
    st.metric("총 누적 출고 수량", f"{len(dispatched_df)}건")

    if not dispatched_df.empty:
        excel_out_data = generate_excel_report(dispatched_df, title="출고내역")
        st.download_button(
            "📊 출고내역 엑셀 다운로드 (.xlsx)",
            data=excel_out_data,
            file_name=f"출고내역_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.divider()

        if not is_admin:
            render_html_table(dispatched_df)
        else:
            st.info("💡 잘못 등록된 출고 건은 **[삭제]** 체크박스를 체크한 뒤 아래 버튼을 눌러 삭제할 수 있습니다.")

            dispatched_df_display = dispatched_df.reset_index(drop=True).copy()
            dispatched_df_display["삭제 체크"] = False

            edited_dispatched_df = st.data_editor(
                dispatched_df_display,
                use_container_width=True,
                hide_index=True,
                disabled=["종류", "품번", "규격", "입고일", "출고일", "배송지", "특기사항"],
                column_config={
                    "삭제 체크": st.column_config.CheckboxColumn("삭제", width="small", default=False),
                },
                key="editor_dispatched",
            )

            hist_delete_parts = edited_dispatched_df[edited_dispatched_df["삭제 체크"] == True]["품번"].tolist()

            if st.button("🗑️ 선택 항목 삭제", use_container_width=True, key="delete_dispatched_btn"):
                if not hist_delete_parts:
                    st.warning("삭제할 품번을 1개 이상 선택해주세요.")
                else:
                    hd_errors = []
                    hd_success = []
                    for p in hist_delete_parts:
                        try:
                            db.delete_item(p)
                            hd_success.append(p)
                        except Exception as e:
                            hd_errors.append(f"[{p}] {e}")
                    if hd_success:
                        st.success(f"🗑️ 총 {len(hd_success)}건 삭제 완료: {', '.join(hd_success)}")
                    for err in hd_errors:
                        st.error(f"❌ {err}")
                    if hd_success:
                        st.rerun()
    else:
        st.info("출고 처리된 내역이 없습니다.")

# TAB 5: Master 항목 관리
with tab_manage:
    if not is_admin:
        st.warning("🔒 기준 정보 관리는 관리자 권한이 필요합니다. 상단에서 관리자 비밀번호를 인증해 주세요.")
    else:
        st.subheader("⚙️ 기준 정보(Master) 관리")
        col_cat, col_dest = st.columns(2)

        with col_cat:
            st.markdown("##### 🏷️ 강판 종류 추가")
            new_cat = st.text_input("신규 강판 종류", placeholder="예: 도금강판", key="new_cat_input")
            if st.button("종류 추가", use_container_width=True):
                if new_cat:
                    try:
                        db.add_category(new_cat)
                        st.success(f"[{new_cat}] 추가 완료!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"추가 실패: {e}")

            st.markdown("##### 🗑️ 강판 종류 삭제")
            existing_cats = db.get_categories()
            if existing_cats:
                cat_to_delete = st.selectbox("삭제할 종류 선택", existing_cats, key="del_cat_select")
                if st.button("종류 삭제", use_container_width=True):
                    try:
                        db.delete_category(cat_to_delete)
                        st.success(f"[{cat_to_delete}] 삭제 완료! (기존 재고 데이터는 그대로 유지됩니다)")
                        st.rerun()
                    except Exception as e:
                        st.error(f"삭제 실패: {e}")
            else:
                st.caption("등록된 종류가 없습니다.")

        with col_dest:
            st.markdown("##### 🚚 배송지 추가")
            new_dest = st.text_input("신규 배송지", placeholder="예: B공장", key="new_dest_input")
            if st.button("배송지 추가", use_container_width=True):
                if new_dest:
                    try:
                        db.add_destination(new_dest)
                        st.success(f"[{new_dest}] 추가 완료!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"추가 실패: {e}")

            st.markdown("##### 🗑️ 배송지 삭제")
            existing_dests = db.get_destinations()
            if existing_dests:
                dest_to_delete = st.selectbox("삭제할 배송지 선택", existing_dests, key="del_dest_select")
                if st.button("배송지 삭제", use_container_width=True):
                    try:
                        db.delete_destination(dest_to_delete)
                        st.success(f"[{dest_to_delete}] 삭제 완료! (기존 출고 데이터는 그대로 유지됩니다)")
                        st.rerun()
                    except Exception as e:
                        st.error(f"삭제 실패: {e}")
            else:
                st.caption("등록된 배송지가 없습니다.")

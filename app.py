import io
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
ITEMS_HEADERS = ["품번", "종류", "규격", "위치", "상태", "입고일", "출고일", "배송지", "특이사항"]
ITEMS_COL = {name: i + 1 for i, name in enumerate(ITEMS_HEADERS)}  # 1-based 열 번호

LOG_HEADERS = ["품번", "종류", "규격", "위치", "입고일", "출고일", "배송지", "특이사항", "기록시각"]
CATEGORIES_HEADERS = ["종류"]
DESTINATIONS_HEADERS = ["배송지"]

STATUS_LABEL = {
    "IN_STOCK": "재고현황",
    "PENDING_DISPATCH": "출고예정내역",
    "DISPATCHED": "출고내역(출고완료)",
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def format_location_code(zone, row, col):
    return f"{str(zone).strip().upper()}-{int(row):02d}-{int(col):02d}"


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


# ---------------------------------------------------------
# 3. 구글 시트 DB 엔진
# ---------------------------------------------------------
class SteelYardSheetDB:
    def __init__(self):
        self.ss = get_spreadsheet()
        self.ws_items = self._get_or_create_ws("Items", ITEMS_HEADERS)
        self.ws_categories = self._get_or_create_ws("Categories", CATEGORIES_HEADERS)
        self.ws_destinations = self._get_or_create_ws("Destinations", DESTINATIONS_HEADERS)
        self.ws_log = self._get_or_create_ws("DispatchLog", LOG_HEADERS)
        self._seed_defaults()

    def _get_or_create_ws(self, name, headers):
        try:
            ws = self.ss.worksheet(name)
        except WorksheetNotFound:
            ws = self.ss.add_worksheet(title=name, rows=1000, cols=len(headers) + 2)
            ws.append_row(headers)
            return ws
        if ws.row_values(1) != headers:
            ws.update("A1", [headers])
        return ws

    def _seed_defaults(self):
        if len(self.ws_categories.col_values(1)) <= 1:
            self.ws_categories.append_rows([["열연강판"], ["냉연강판"], ["후판"]])
        if len(self.ws_destinations.col_values(1)) <= 1:
            self.ws_destinations.append_rows([["본사 공장"], ["A현장"], ["하청1"]])

    # ---------------- Master 정보 ----------------
    def get_categories(self):
        return [v.strip() for v in self.ws_categories.col_values(1)[1:] if v.strip()]

    def add_category(self, name):
        name = name.strip()
        if not name:
            return
        if name in self.get_categories():
            raise ValueError(f"[{name}] 는 이미 등록된 종류입니다.")
        self.ws_categories.append_row([name])

    def get_destinations(self):
        return [v.strip() for v in self.ws_destinations.col_values(1)[1:] if v.strip()]

    def add_destination(self, name):
        name = name.strip()
        if not name:
            return
        if name in self.get_destinations():
            raise ValueError(f"[{name}] 는 이미 등록된 배송지입니다.")
        self.ws_destinations.append_row([name])

    # ---------------- 내부 헬퍼 ----------------
    def _items_df(self):
        records = self.ws_items.get_all_records()
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
    def register_inbound(self, part_no, category_name, spec, zone, row, col, inbound_date, remarks):
        clean_part_no = part_no.strip().upper()
        loc_code = format_location_code(zone, row, col)

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
            raise ValueError(f"[{loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다 (1단 적재만 허용).")

        clean_spec = spec.strip().upper() if spec else ""
        self.ws_items.append_row([
            clean_part_no, category_name, clean_spec, loc_code, "IN_STOCK",
            str(inbound_date), "", "", remarks or ""
        ])

    # ---------------- 적재위치 변경 ----------------
    def update_location(self, part_no, new_zone, new_row, new_col):
        clean_part_no = part_no.strip().upper()
        new_loc_code = format_location_code(new_zone, new_row, new_col)

        row_idx = self._find_item_row(clean_part_no)
        if not row_idx:
            raise ValueError("존재하지 않는 품번입니다.")

        df = self._items_df()
        active = df[(df["상태"].isin(["IN_STOCK", "PENDING_DISPATCH"])) & (df["품번"] != clean_part_no)]
        if not active.empty and (active["위치"] == new_loc_code).any():
            raise ValueError(f"[{new_loc_code}] 위치에는 이미 다른 강판이 적재되어 있습니다.")

        self.ws_items.update_cell(row_idx, ITEMS_COL["위치"], new_loc_code)

    # ---------------- 2단계: 재고현황 체크 -> 출고예정(PENDING_DISPATCH) ----------------
    def move_to_pending_dispatch(self, part_numbers):
        for p in part_numbers:
            row_idx = self._find_item_row(p)
            if row_idx:
                self.ws_items.update_cell(row_idx, ITEMS_COL["상태"], "PENDING_DISPATCH")

    # ---------------- 3단계: 출고확정 -> 출고내역(DISPATCHED) ----------------
    def confirm_outbound(self, part_no, outbound_date, dest_name, remarks):
        clean_part_no = part_no.strip().upper()
        row_idx = self._find_item_row(clean_part_no)
        if not row_idx:
            raise ValueError("존재하지 않는 품번입니다.")

        row_vals = self.ws_items.row_values(row_idx)
        row_vals += [""] * (len(ITEMS_HEADERS) - len(row_vals))

        existing_remarks = row_vals[ITEMS_COL["특이사항"] - 1]
        final_remarks = existing_remarks
        if remarks:
            final_remarks = f"{existing_remarks} / [출고비고] {remarks}" if existing_remarks else remarks

        inbound_date_val = row_vals[ITEMS_COL["입고일"] - 1]

        # 상태(E) ~ 특이사항(I)까지 한 번에 갱신 (API 호출 최소화)
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

    # ---------------- 4단계: 보류 체크 -> 재고현황(IN_STOCK) 복원 ----------------
    def restore_to_in_stock(self, part_numbers):
        if isinstance(part_numbers, str):
            part_numbers = [part_numbers]
        for p in part_numbers:
            row_idx = self._find_item_row(p)
            if row_idx:
                self.ws_items.update_cell(row_idx, ITEMS_COL["상태"], "IN_STOCK")

    # ---------------- 조회 ----------------
    def search_current_stock(self, category_name="전체", part_kw="", spec_kw=""):
        df = self._items_df()
        cols = ["종류", "품번", "규격", "적재위치", "입고일", "특이사항"]
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
        cols = ["종류", "품번", "규격", "적재위치", "입고일", "특이사항"]
        if df.empty:
            return pd.DataFrame(columns=cols)
        df = df[df["상태"] == "PENDING_DISPATCH"].copy().rename(columns={"위치": "적재위치"})
        if df.empty:
            return pd.DataFrame(columns=cols)
        return df[cols].sort_values("품번").reset_index(drop=True)

    def search_dispatched(self):
        df = self._items_df()
        cols = ["종류", "품번", "규격", "적재위치", "입고일", "출고일", "배송지", "특이사항"]
        if df.empty:
            return pd.DataFrame(columns=cols)
        df = df[df["상태"] == "DISPATCHED"].copy().rename(columns={"위치": "적재위치"})
        if df.empty:
            return pd.DataFrame(columns=cols)
        return df[cols].sort_values(["출고일", "품번"], ascending=[False, True]).reset_index(drop=True)


# ---------------------------------------------------------
# 엑셀 다운로드 생성기
# ---------------------------------------------------------
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

# 모바일에서 종류/품번/규격/위치/입고일/특이사항/체크박스가
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
}
[data-testid="stDataFrame"] div[role="columnheader"],
[data-testid="stDataEditor"] div[role="columnheader"] {
    padding-left: 4px !important;
    padding-right: 4px !important;
    font-size: 12px !important;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

st.title("📲 제2야적장 입출고 관리 (Google Sheets 연동)")

try:
    ADMIN_PIN = st.secrets["admin_pin"]
except Exception:
    ADMIN_PIN = "1234"  # secrets.toml 미설정 시 임시값 — 운영 전 반드시 교체

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

c_role, c_pin = st.columns([2, 3])
role_selection = c_role.radio("사용자 권한", ["일반 작업자 (조회만)", "관리자"], horizontal=True)

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
                    st.markdown("**품번 (4단 분할)**")
                    p1, p2, p3, p4 = st.columns(4)
                    part_1 = p1.text_input("1단", placeholder="FA", label_visibility="collapsed")
                    part_2 = p2.text_input("2단", placeholder="3B", label_visibility="collapsed")
                    part_3 = p3.text_input("3단", placeholder="01", label_visibility="collapsed")
                    part_4 = p4.text_input("4단", placeholder="02", label_visibility="collapsed")

                with col_spec:
                    spec = st.text_input("규격", placeholder="예: 12T x 1500 x 6000")

                with col_loc:
                    st.markdown("**적재위치 (1단 원칙)**")
                    z_col, r_col, c_col = st.columns(3)
                    zone = z_col.text_input("구역", value="A", label_visibility="collapsed")
                    row_n = r_col.number_input("열", min_value=1, value=1, label_visibility="collapsed")
                    col_n = c_col.number_input("행", min_value=1, value=1, label_visibility="collapsed")

                with col_date:
                    in_date = st.date_input("입고일", datetime.now())

                with col_rem:
                    remarks = st.text_input("특이사항", placeholder="예: 측면 스크래치")

                st.divider()

                if st.form_submit_button("⚡ 실시간 입고 등록", use_container_width=True):
                    part_components = [p.strip() for p in [part_1, part_2, part_3, part_4] if p.strip()]
                    if len(part_components) < 4:
                        st.error("품번 4단 칸을 모두 입력해주세요.")
                    else:
                        full_part_no = "-".join(part_components).upper()
                        try:
                            db.register_inbound(full_part_no, cat_name, spec, zone, row_n, col_n, in_date, remarks)
                            st.success(f"✅ [{full_part_no}] 입고 등록 완료! (위치: {zone.upper()}-{row_n:02d}-{col_n:02d})")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 등록 실패: {e}")

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
                    u_col1, u_col2, u_col3, u_col4 = st.columns([2.5, 1, 1, 1])
                    u_part = u_col1.selectbox("위치 변경 대상 품번", stock_df["품번"].tolist())
                    u_zone = u_col2.text_input("신규 구역", value="A")
                    u_row = u_col3.number_input("신규 열", min_value=1, value=1)
                    u_col = u_col4.number_input("신규 행", min_value=1, value=1)
                    if st.form_submit_button("⚡ 적재위치 변경 저장", use_container_width=True):
                        try:
                            db.update_location(u_part, u_zone, u_row, u_col)
                            st.success(f"✅ [{u_part}] 위치가 [{u_zone.upper()}-{u_row:02d}-{u_col:02d}](으)로 수정되었습니다!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 위치 변경 실패: {e}")

            st.info("💡 출고 예정으로 보낼 품번의 **[출고 선택]** 체크박스를 체크한 후 아래 버튼을 누르세요.")

            stock_df_display = stock_df.copy()
            stock_df_display["출고 선택"] = False

            edited_df = st.data_editor(
                stock_df_display,
                use_container_width=True,
                hide_index=True,
                disabled=["종류", "품번", "규격", "적재위치", "입고일", "특이사항"],
                column_config={
                    "종류": st.column_config.TextColumn("종류", width="small"),
                    "품번": st.column_config.TextColumn("품번", width="small"),
                    "규격": st.column_config.TextColumn("규격", width="small"),
                    "적재위치": st.column_config.TextColumn("위치", width="small"),
                    "입고일": st.column_config.DateColumn("입고일", width="small", format="MM/DD"),
                    "특이사항": st.column_config.TextColumn("특이사항", width="small"),
                    "출고 선택": st.column_config.CheckboxColumn("출고", width="small", default=False),
                },
                key="editor_stock",
            )
            selected_parts = edited_df[edited_df["출고 선택"] == True]["품번"].tolist()

            if st.button("🚚 선택 항목 출고예정으로 이동", type="primary", use_container_width=True):
                if not selected_parts:
                    st.warning("출고 예정으로 이동할 품번을 1개 이상 선택해주세요.")
                else:
                    db.move_to_pending_dispatch(selected_parts)
                    st.success(f"✅ 총 {len(selected_parts)}건이 [출고예정내역]으로 이동되었습니다: {', '.join(selected_parts)}")
                    st.rerun()
        else:
            st.dataframe(
                stock_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "종류": st.column_config.TextColumn("종류", width="small"),
                    "품번": st.column_config.TextColumn("품번", width="small"),
                    "규격": st.column_config.TextColumn("규격", width="small"),
                    "적재위치": st.column_config.TextColumn("위치", width="small"),
                    "입고일": st.column_config.DateColumn("입고일", width="small", format="MM/DD"),
                    "특이사항": st.column_config.TextColumn("특이사항", width="small"),
                },
            )
    else:
        st.info("조건에 일치하는 재고가 없습니다.")

# TAB 3: 출고예정내역
with tab_pending:
    st.subheader("🚚 출고예정내역")

    pending_df = db.search_pending_dispatch()
    st.metric("출고 대기 수량", f"{len(pending_df)}건")

    if not pending_df.empty:
        if is_admin:
            st.info(
                "💡 **보류**할 품번은 [보류 체크]를 체크한 뒤 [↩️ 선택 항목 보류 → 재고현황 복원] 버튼을 누르세요. "
                "출고를 확정할 품번은 아래 출고확정 양식을 이용하세요."
            )

            pending_df_display = pending_df.copy()
            pending_df_display["보류 체크"] = False

            edited_pending_df = st.data_editor(
                pending_df_display,
                use_container_width=True,
                hide_index=True,
                disabled=["종류", "품번", "규격", "적재위치", "입고일", "특이사항"],
                column_config={
                    "종류": st.column_config.TextColumn("종류", width="small"),
                    "품번": st.column_config.TextColumn("품번", width="small"),
                    "규격": st.column_config.TextColumn("규격", width="small"),
                    "적재위치": st.column_config.TextColumn("위치", width="small"),
                    "입고일": st.column_config.DateColumn("입고일", width="small", format="MM/DD"),
                    "특이사항": st.column_config.TextColumn("특이사항", width="small"),
                    "보류 체크": st.column_config.CheckboxColumn("보류", width="small", default=False),
                },
                key="editor_pending",
            )
            hold_parts = edited_pending_df[edited_pending_df["보류 체크"] == True]["품번"].tolist()

            if st.button("↩️ 선택 항목 보류 → 재고현황 복원", use_container_width=True):
                if not hold_parts:
                    st.warning("보류 처리할 품번을 1개 이상 선택해주세요.")
                else:
                    db.restore_to_in_stock(hold_parts)
                    st.success(f"↩️ 총 {len(hold_parts)}건이 보류 처리되어 [실시간 재고현황]으로 복원되었습니다: {', '.join(hold_parts)}")
                    st.rerun()

            st.divider()
            st.markdown("#### ✅ 출고 확정 처리")

            destinations = db.get_destinations()
            if not destinations:
                st.error("등록된 배송지가 없습니다. Master 관리 탭에서 추가해 주세요.")
            else:
                with st.form("pending_confirm_form", clear_on_submit=True):
                    col_p, col_dst, col_dt, col_rem = st.columns([2.5, 2, 1.5, 2])
                    with col_p:
                        target_part = st.selectbox("출고 확정 대상 품번", pending_df["품번"].tolist())
                    with col_dst:
                        dest_name = st.selectbox("출고 배송지", destinations)
                    with col_dt:
                        out_date = st.date_input("출고일", datetime.now())
                    with col_rem:
                        out_remarks = st.text_input("출고 비고", placeholder="예: A현장 1차")

                    if st.form_submit_button("✅ 최종출고 확정", use_container_width=True, type="primary"):
                        try:
                            db.confirm_outbound(target_part, out_date, dest_name, out_remarks)
                            st.success(f"🎉 [{target_part}] 최종 출고 처리 완료! (출고내역으로 이동, 재고현황에서 제외)")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 출고 실패: {e}")
        else:
            st.dataframe(pending_df, use_container_width=True, hide_index=True)
            st.info("💡 출고 확정 및 보류 처리는 관리자 권한이 필요합니다.")
    else:
        st.info("현재 출고 예정인 내역이 없습니다.")

# TAB 4: 출고내역
with tab_history:
    st.subheader("📜 출고 완료 내역")

    dispatched_df = db.search_dispatched()
    st.metric("총 누적 출고 수량", f"{len(dispatched_df)}건")

    if not dispatched_df.empty:
        excel_out_data = generate_excel_report(dispatched_df, title="출고완료내역")
        st.download_button(
            "📊 출고내역 엑셀 다운로드 (.xlsx)",
            data=excel_out_data,
            file_name=f"출고완료내역_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.divider()
        st.dataframe(dispatched_df, use_container_width=True, hide_index=True)
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

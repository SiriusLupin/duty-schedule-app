import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook

# ====== Google Drive API（Service Account）套件 ======
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# ============================================================
# 0) 使用者可編輯簡化對照表（預設值）
# ============================================================
default_rules = [
    {"原始關鍵字": "調劑複核", "簡化後": "C"},
    {"原始關鍵字": "處方判讀", "簡化後": "判讀"},
    {"原始關鍵字": "藥物諮詢", "簡化後": "諮詢"},
    {"原始關鍵字": "門診藥局調劑", "簡化後": "門診"},
    {"原始關鍵字": "中正 2樓", "簡化後": "中2"},
    {"原始關鍵字": "中正13樓", "簡化後": "中13"},
    {"原始關鍵字": "思源樓", "簡化後": "思源"},
    {"原始關鍵字": "長青樓", "簡化後": "長青"},
    {"原始關鍵字": "抗凝藥師門診", "簡化後": "抗凝門診"},
    {"原始關鍵字": "移植藥師門診", "簡化後": "移植門診"},
    {"原始關鍵字": "中藥局調劑", "簡化後": "中藥局"},
    {"原始關鍵字": "非常班之諮詢與藥動服務", "簡化後": "假日oncall"},
]


# ============================================================
# 1) Google Drive 下載/列檔工具（Service Account）
# ============================================================
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

def build_drive_service():
    """
    用 Streamlit secrets 內的 service account 建立 Drive API client。
    你必須先在 Streamlit Cloud 的 Secrets 或 .streamlit/secrets.toml 放入
    [gcp_service_account] 區塊（type/project_id/private_key/client_email/token_uri...）。
    """
    if "gcp_service_account" not in st.secrets:
        st.error("❌ 找不到 st.secrets['gcp_service_account']，請先設定 Streamlit Secrets。")
        st.stop()

    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=DRIVE_SCOPES
    )
    return build("drive", "v3", credentials=creds)


def extract_drive_file_id(url: str) -> str | None:
    """
    從使用者貼上的 Google Drive / Google Sheet 連結中抽出 file_id。
    支援常見格式：
    - https://docs.google.com/spreadsheets/d/<ID>/edit...
    - https://drive.google.com/file/d/<ID>/view...
    - https://drive.google.com/open?id=<ID>
    - ...?id=<ID>
    """
    if not url:
        return None

    patterns = [
        r"/d/([a-zA-Z0-9-_]+)",      # .../d/<id>/...
        r"[?&]id=([a-zA-Z0-9-_]+)",  # ...?id=<id> 或 &id=<id>
        r"open\?id=([a-zA-Z0-9-_]+)",
        r"file/d/([a-zA-Z0-9-_]+)",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def download_drive_file_as_bytes(file_id: str) -> io.BytesIO:
    """
    下載 Google Drive 檔案成 BytesIO（記憶體檔案），供 pandas/openpyxl 讀取。
    同時支援：
    A) Google 試算表（原生） -> export 成 xlsx
    B) 真正 .xlsx 檔 -> get_media 直接下載
    """
    service = build_drive_service()
    meta = service.files().get(fileId=file_id, fields="name,mimeType").execute()
    mime = meta.get("mimeType", "")

    bio = io.BytesIO()

    # Google Sheets -> 匯出成 XLSX
    if mime == "application/vnd.google-apps.spreadsheet":
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        # 例如 .xlsx
        request = service.files().get_media(fileId=file_id)

    downloader = MediaIoBaseDownload(bio, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    bio.seek(0)
    return bio


def list_recent_drive_files(months_approx_days: int = 92, page_size: int = 100):
    """
    列出近三個月（約 92 天）內有更新的：
    - Google 試算表
    - Excel .xlsx

    注意：Service Account 只看得到「自己建立」或「別人共享給它」的檔案。
    """
    service = build_drive_service()

    since_dt = datetime.now(timezone.utc) - timedelta(days=months_approx_days)
    since_str = since_dt.isoformat().replace("+00:00", "Z")

    q = (
        "("
        "mimeType='application/vnd.google-apps.spreadsheet' OR "
        "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
        ") "
        f"AND modifiedTime >= '{since_str}' "
        "AND trashed=false"
    )

    resp = service.files().list(
        q=q,
        fields="files(id,name,mimeType,modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=page_size
    ).execute()

    return resp.get("files", [])


def get_excel_bio(source_choice: str, uploaded_file, selected_drive_file, drive_url_backup: str):
    """
    統一回傳 BytesIO，讓後續解析只寫一套。
    source_choice：
      - 上傳 Excel
      - Google Drive（近3個月下拉選）
      - Google Drive（貼連結備援）
    """
    if source_choice == "上傳 Excel":
        if not uploaded_file:
            return None
        data = uploaded_file.read()
        bio = io.BytesIO(data)
        bio.seek(0)
        return bio

    if source_choice == "現有共用班表檔案(3個月內)":
        if not selected_drive_file:
            return None
        return download_drive_file_as_bytes(selected_drive_file["id"])

    # 貼連結備援
    if not drive_url_backup:
        return None
    file_id = extract_drive_file_id(drive_url_backup)
    if not file_id:
        st.error("❌ 無法從連結解析檔案 ID，請確認貼的是 Drive/Sheet 分享連結。")
        st.stop()

    try:
        return download_drive_file_as_bytes(file_id)
    except Exception as e:
        st.error(f"❌ 從 Google Drive 下載失敗：{e}")
        st.stop()


# ============================================================
# 2) 灰底假日判斷：第二列日期底色（灰色=假日）
# ============================================================
def build_holiday_map(excel_bio: io.BytesIO) -> dict[int, bool]:
    """
    用 openpyxl 讀取 Excel：
    - 第二列（row=2）日期列的底色（灰底代表假日）
    回傳 holiday_map：{ openpyxl_column_index(1-based): is_holiday }
    """
    excel_bio.seek(0)
    wb = load_workbook(excel_bio, data_only=True)
    ws = wb.active

    # 你目前使用的灰底 RGB
    gray_rgb = "FFD9D9D9"

    holiday_map = {}
    for col in range(2, ws.max_column + 1):  # B欄開始（A欄是工作內容）
        cell = ws.cell(row=2, column=col)
        fg = cell.fill.fgColor
        is_gray = (fg.type == "rgb" and fg.rgb == gray_rgb)
        holiday_map[col] = is_gray

    return holiday_map


# ============================================================
# 3) 套用時間規則（含你新增的中2藥局發藥括號時間）
# ============================================================
def apply_time_rules(df, holiday_map, column_map):
    """
    df 欄位應含：日期、星期、工作內容、簡化後內容、Start Time、End Time
    holiday_map：欄位底色假日判定
    column_map： (日期, 星期) -> Excel 欄位 index（B=2 起）
    """
    prescription_time_map = {
        "上午": ("08:00", "12:00"),
        "下午": ("13:30", "17:30"),
        "小夜1hr": ("17:30", "18:30"),
        "小夜": ("17:30", "21:30")
    }

    extra_rows = []

    for idx, row in df.iterrows():
        content = row["工作內容"]
        weekday = str(row["星期"]).strip()

        key = (row["日期"], weekday)
        col_idx = column_map.get(key, None)
        is_holiday = holiday_map.get(col_idx, False)

        # 1) 調劑複核（平日 vs 假日）
        if "調劑複核" in content:
            if is_holiday:
                df.at[idx, "Start Time"] = "11:00"
                df.at[idx, "End Time"] = "15:00"
            else:
                df.at[idx, "Start Time"] = "13:30"
                df.at[idx, "End Time"] = "15:00"

        # 2) 門診藥局調劑（括號時間）
        elif "門診藥局調劑" in content:
            match = re.search(r"\((\d{1,2}:\d{2})-(\d{1,2}:\d{2})\)", content)
            if match:
                df.at[idx, "Start Time"] = match.group(1)
                df.at[idx, "End Time"] = match.group(2)

        # 2.5) 中2藥局發藥（括號時間）
        elif "中2藥局" in content:
            match = re.search(r"\((\d{1,2}:\d{2})-(\d{1,2}:\d{2})\)", content)
            if match:
                df.at[idx, "Start Time"] = match.group(1)
                df.at[idx, "End Time"] = match.group(2)

        # 3) 處方判讀 / 化療處方判讀 / 藥物諮詢 / PreESRD（依上午/下午/小夜）
        elif any(k in content for k in ["處方判讀", "化療處方判讀", "藥物諮詢", "PreESRD"]):
            for key_word, (start, end) in prescription_time_map.items():
                if key_word in content:
                    df.at[idx, "Start Time"] = start
                    df.at[idx, "End Time"] = end
                    break

        # 4) 抗凝藥師門診：週二上午 / 週三下午
        elif "抗凝藥師門診" in content:
            if weekday == "二":
                df.at[idx, "Start Time"] = "08:30"
                df.at[idx, "End Time"] = "12:00"
            elif weekday == "三":
                df.at[idx, "Start Time"] = "13:30"
                df.at[idx, "End Time"] = "17:00"

        # 5) 移植藥師門診：目前只有上午
        # 若未來有下午，請在此補 elif "下午" in content: ...
        elif "移植藥師門診" in content and "上午" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        # 6) 中藥局調劑：目前固定 08:30-12:00（你可再加 weekday == "三" 的限制）
        elif "中藥局調劑" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        # 7) 瑞德西偉審核：08:00-20:00
        elif "瑞德西偉審核" in content:
            df.at[idx, "Start Time"] = "08:00"
            df.at[idx, "End Time"] = "20:00"

        # 8) 平日：若工作為「處方判讀 7-住院」，額外新增「非常班之諮詢與藥動服務」17:30-21:30
        if "處方判讀 7-住院" in content and not is_holiday:
            extra_rows.append({
                "日期": row["日期"],
                "星期": row["星期"],
                "工作內容": "非常班之諮詢與藥動服務",
                "簡化後內容": "非常班之諮詢與藥動服務",  # 後面仍會做簡化 replace
                "Start Time": "17:30",
                "End Time": "21:30"
            })

        # 9) 假日：「非常班之諮詢與藥動服務」三班
        if "非常班之諮詢與藥動服務" in content and is_holiday:
            if "上午" in content:
                df.at[idx, "Start Time"] = "08:00"
                df.at[idx, "End Time"] = "12:30"
            elif "下午" in content:
                df.at[idx, "Start Time"] = "12:30"
                df.at[idx, "End Time"] = "17:00"
            elif "晚上" in content:
                df.at[idx, "Start Time"] = "17:00"
                df.at[idx, "End Time"] = "21:00"

    if extra_rows:
        df = pd.concat([df, pd.DataFrame(extra_rows)], ignore_index=True)

    return df


# ============================================================
# 4) Streamlit UI：排版順序 1代號 2來源 3縮寫表
# ============================================================
st.title("📆 班表轉換工具（支援假日底色與字詞縮寫對照表）")

# 操作說明（下載）
try:
    with open("班表轉換操作說明v2.pdf", "rb") as f:
        st.download_button("📘 下載操作說明 PDF", data=f.read(), file_name="班表轉換操作說明v2.pdf")
except FileNotFoundError:
    st.caption("（找不到操作說明 PDF 檔案；若在 Streamlit Cloud 請確認已放入 Repo）")

# 1) 班表代號
code = st.text_input("請輸入班表代號：")

# 2) 班表來源：上傳 / Drive 下拉 / Drive 連結備援
st.subheader("📁 班表來源")
source = st.radio(
    "選擇班表來源：",
    ["上傳 Excel", "現有共用班表檔案(3個月內)", "試算表連結"],
    horizontal=False
)

uploaded_file = None
selected_drive_file = None
drive_url_backup = ""

if source == "上傳 Excel":
    uploaded_file = st.file_uploader("請上傳 Excel 班表（.xlsx）")

elif source == "現有共用班表檔案(3個月內)":
    # 近三個月清單
    try:
        files = list_recent_drive_files(months_approx_days=92, page_size=100)
    except Exception as e:
        st.error(f"❌ 無法列出 Google Drive 檔案：{e}")
        files = []

    if not files:
        st.warning("目前 Service Account 近3個月內看不到任何 Excel/試算表。請確認：主管有共享檔案給服務帳號，且檔案近期有更新。")
    else:
        def pretty_label(f):
            typ = "Google試算表" if f["mimeType"] == "application/vnd.google-apps.spreadsheet" else "Excel(.xlsx)"
            mt = f.get("modifiedTime", "")
            return f'{f["name"]} ｜ {typ} ｜ {mt}'

        options = {pretty_label(f): f for f in files}
        chosen = st.selectbox("請選擇班表檔案（近3個月更新）：", list(options.keys()))
        selected_drive_file = options[chosen]

else:
    drive_url_backup = st.text_input("請貼上 Google Drive / Google 試算表連結")


# 3) 簡化對照表（不需要等上傳才顯示）
st.subheader("🔧 字詞縮寫表")
st.markdown(
    """<p style='color:black; font-size:16px; font-weight:bold;'>
    您可以自行修改想要的縮寫，並可由下方表格預覽。<br>
    也可點選右上角的「+」新增欄位自訂縮寫
    </p>""",
    unsafe_allow_html=True
)
st.markdown(
    "<p style='color:red; font-size:18px; font-weight:bold;'>🗑️⚠ 注意！若留有空行程式可能發生錯誤，請將空行右側方框勾選後，右上角點選刪除。</p>",
    unsafe_allow_html=True
)

df_rules = pd.DataFrame(default_rules)
edited_rules = st.data_editor(df_rules, use_container_width=True, num_rows="dynamic")
simplify_map = dict(zip(edited_rules["原始關鍵字"], edited_rules["簡化後"]))


# ============================================================
# 5) 主流程：讀檔 -> 假日底色 -> 解析日期/星期 -> 找代號 -> 縮寫 -> 時間 -> 輸出
# ============================================================
excel_bio = get_excel_bio(source, uploaded_file, selected_drive_file, drive_url_backup)

if code and excel_bio:
    # (A) 讀 Excel 成 DataFrame
    excel_bio.seek(0)
    df = pd.read_excel(excel_bio, header=None)

    # (B) 底色判斷假日
    holiday_map = build_holiday_map(excel_bio)

    # (C) 從第一列標題抓民國年與月份（例如：113年4月班表）
    title = str(df.iat[0, 0])
    m = re.search(r"(\d{2,3})年(\d{1,2})月", title)
    if not m:
        st.error("❌ 無法擷取年份與月份，請確認標題格式如『113年4月班表』")
        st.stop()

    year = int(m.group(1)) + 1911
    month = int(m.group(2))
    year_month = f"{year}{month:02d}"

    # (D) 第二、三列為日期與星期（B欄開始）
    dates = df.iloc[1, 1:].tolist()
    weekdays = df.iloc[2, 1:].tolist()

    # date_mapping：每一欄對應的（日期、星期）
    date_mapping = [
        {"日期": f"{year}-{month:02d}-{int(d):02d}", "星期": weekdays[i]}
        for i, d in enumerate(dates)
        if str(d).strip().isdigit()
    ]

    # col_index_map：給底色查詢用 (日期, 星期) -> Excel 欄 index（B=2 起）
    col_index_map = {
        (entry["日期"], entry["星期"]): i + 2
        for i, entry in enumerate(date_mapping)
    }

    # (E) 掃描 A 欄工作內容，找出含「代號」的日期欄
    results = []
    for row_idx in range(3, df.shape[0]):
        raw = df.iat[row_idx, 0]

        # 正確判斷 nan：先判斷原始值，再轉字串
        if pd.isna(raw):
            continue

        content = str(raw).strip()
        if not content:
            continue
        if content.lower() == "nan":
            continue
        if "附　註" in content:
            continue

        for col_idx in range(1, len(date_mapping) + 1):
            cell = df.iat[row_idx, col_idx]
            cell_str = "" if pd.isna(cell) else str(cell)

            if code in cell_str:
                # (F) 先移除括號時間（你原本規則）
                simplified = re.sub(r"\(\d{1,2}:\d{2}-\d{1,2}:\d{2}\)", "", content)

                # (G) 再依縮寫表 replace（避免空值造成錯誤）
                for k, v in simplify_map.items():
                    if pd.notna(k) and pd.notna(v):
                        simplified = simplified.replace(str(k), str(v))

                results.append({
                    "日期": date_mapping[col_idx - 1]["日期"],
                    "星期": date_mapping[col_idx - 1]["星期"],
                    "工作內容": content,
                    "簡化後內容": simplified,
                })

    df_result = pd.DataFrame(results)
    if df_result.empty:
        st.warning("找不到符合此代號的班表內容。請確認代號是否正確，或該月未排班。")
        st.stop()

    # (H) 套用時間規則
    df_result["Start Time"] = ""
    df_result["End Time"] = ""
    df_result = apply_time_rules(df_result, holiday_map, col_index_map)

    # (I) 輸出 Google Calendar CSV 欄位
    df_output = df_result.rename(columns={"簡化後內容": "Subject", "日期": "Start Date"})
    df_output["End Date"] = df_output["Start Date"]
    df_output = df_output[["Subject", "Start Date", "Start Time", "End Date", "End Time"]]

    # (J) 匯出 CSV：UTF-8 with BOM（Excel 打開較不容易亂碼）
    csv_text = df_output.to_csv(index=False, encoding="utf-8-sig")

    st.success("✅ 轉換完成")
    st.subheader("內容預覽")
    st.dataframe(df_output, use_container_width=True)

    st.markdown(
        "<p style='color:red; font-size:18px; font-weight:bold;'>⚠ CSV 檔案直接開啟內容可能為亂碼，但不影響匯入，請先確認上方資料無誤後再點選下方按鈕下載。</p>",
        unsafe_allow_html=True
    )

    st.download_button(
        label=f"📥 下載 {year_month}個人班表({code}).csv",
        data=csv_text,
        file_name=f"{year_month}個人班表({code}).csv",
        mime="text/csv"
    )

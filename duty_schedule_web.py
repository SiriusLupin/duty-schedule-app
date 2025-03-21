import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ---------- 時間解析規則 ---------- #
def apply_time_rules(df):
    prescription_time_map = {
        "上午": ("08:00", "12:00"),
        "下午": ("13:30", "17:30"),
        "小夜1hr": ("17:30", "18:30"),
        "小夜": ("17:30", "21:30")
    }

    extra_rows = []
    df = df[~df["工作內容"].str.strip().isin(["", "nan", "附　註"])]

    for idx, row in df.iterrows():
        content = row["工作內容"]
        weekday = str(row["星期"]).strip()

        if "調劑複核" in content:
            if weekday in ["一", "二", "三", "四", "五"]:
                df.at[idx, "Start Time"] = "13:30"
                df.at[idx, "End Time"] = "15:00"
            elif weekday in ["六", "日"]:
                df.at[idx, "Start Time"] = "11:00"
                df.at[idx, "End Time"] = "15:00"

        elif "門診藥局調劑" in content:
            match = re.search(r"\((\d{1,2}:\d{2})-(\d{1,2}:\d{2})\)", content)
            if match:
                df.at[idx, "Start Time"] = match.group(1)
                df.at[idx, "End Time"] = match.group(2)

        elif any(k in content for k in ["處方判讀", "化療處方判讀", "藥物諮詢", "PreESRD"]):
            for key, (start, end) in prescription_time_map.items():
                if key in content:
                    df.at[idx, "Start Time"] = start
                    df.at[idx, "End Time"] = end
                    break

        elif "抗凝藥師門診" in content:
            if weekday == "二":
                df.at[idx, "Start Time"] = "08:30"
                df.at[idx, "End Time"] = "12:00"
            elif weekday == "三":
                df.at[idx, "Start Time"] = "13:30"
                df.at[idx, "End Time"] = "17:00"

        elif "移植藥師門診" in content and "上午" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        elif "PreESRD" in content and "上午" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        elif "中藥局調劑" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        elif "瑞德西偉審核" in content:
            df.at[idx, "Start Time"] = "08:00"
            df.at[idx, "End Time"] = "20:00"

        if "處方判讀 7-住院" in content and weekday in ["一", "二", "三", "四", "五"]:
            extra_rows.append({
                "日期": row["日期"],
                "星期": row["星期"],
                "工作內容": "非常班之諮詢與藥動服務",
                "簡化後內容": "非常班之諮詢與藥動服務",
                "Start Time": "17:30",
                "End Time": "21:30"
            })

        if "非常班之諮詢與藥動服務" in content and weekday in ["六", "日"]:
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

# ---------- Streamlit App ---------- #
st.set_page_config(page_title="班表轉換工具", layout="centered")
st.title("📆 班表轉換工具（Google 日曆格式）")

with open("班表轉換成google日曆檔操作說明.pdf", "rb") as f:
    st.download_button("📘 下載操作說明 PDF", data=f.read(), file_name="班表轉換操作說明.pdf")

code = st.text_input("請輸入班表代號：")
file = st.file_uploader("請上傳班表 Excel 檔（.xlsx）")

if file and code:
    df = pd.read_excel(file, header=None)
    title = str(df.iat[0, 0])
    match = re.search(r"(\d{2,3})年(\d{1,2})月", title)
    if match:
        roc_year = int(match.group(1))
        month = int(match.group(2))
        year = roc_year + 1911
        year_month = f"{year}{month:02d}"

        dates = df.iloc[1, 1:].tolist()
        weekdays = df.iloc[2, 1:].tolist()
        date_mapping = [
            {"日期": f"{year}-{month:02d}-{int(day):02d}", "星期": weekdays[i]}
            for i, day in enumerate(dates) if str(day).strip().isdigit()
        ]

        results = []
        for row_idx in range(3, df.shape[0]):
            content = str(df.iat[row_idx, 0]).strip()
            if not content or content.lower() == "nan" or "附　註" in content:
                continue
            for col_idx in range(1, len(date_mapping) + 1):
                cell = str(df.iat[row_idx, col_idx])
                if code in cell:
                    simplified = re.sub(r"\(\d{1,2}:\d{2}-\d{1,2}:\d{2}\)", "", content)
                    simplified = simplified.replace("調劑複核", "C")
                    results.append({
                        "日期": date_mapping[col_idx - 1]["日期"],
                        "星期": date_mapping[col_idx - 1]["星期"],
                        "工作內容": content,
                        "簡化後內容": simplified,
                    })

        df_result = pd.DataFrame(results)
        df_result["Start Time"] = ""
        df_result["End Time"] = ""
        df_result = apply_time_rules(df_result)

        df_output = df_result.rename(columns={"簡化後內容": "Subject", "日期": "Start Date"})
        df_output["End Date"] = df_output["Start Date"]
        df_output = df_output[["Subject", "Start Date", "Start Time", "End Date", "End Time"]]


        # 將 DataFrame 轉為 CSV
        csv = df_output.to_csv(index=False, encoding='utf-8')

        st.success("轉換完成，請點擊下方按鈕下載 CSV 檔")
        st.dataframe(df_output, use_container_width=True)
        
        st.markdown(
    "<p style='color:red; font-size:18px; font-weight:bold;'>⚠ CSV 檔案直接開啟內容可能為亂碼，但不影響匯入，請先確認上方資料無誤後再點選下方按鈕下載。</p>",
    unsafe_allow_html=True
        st.download_button(
            label=f"📥 下載 {year_month}個人班表({code}).csv",
            data=csv,
            file_name=f"{year_month}個人班表({code}).csv",
            mime="text/csv"
        )
    else:
        st.error("無法從標題中擷取年份與月份。請確認第一列格式是否正確（例如：113年4月班表）")

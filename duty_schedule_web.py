import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os

# ---------- 時間解析規則 ---------- #
# 對照表：處方判讀、藥物諮詢、化療處方判讀、PreESRD 時間區段
def apply_time_rules(df):
    prescription_time_map = {
        "上午": ("08:00", "12:00"),
        "下午": ("13:30", "17:30"),
        "小夜1hr": ("17:30", "18:30"),
        "小夜": ("17:30", "21:30")
    }

    extra_rows = []

    # 移除 "附　註"（全形空格）開頭或完全為空的列
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
         # 門診藥局調劑 - 擷取括號中的時間格式
        elif "門診藥局調劑" in content:
            match = re.search(r"\((\d{1,2}:\d{2})-(\d{1,2}:\d{2})\)", content)
            if match:
                df.at[idx, "Start Time"] = match.group(1)
                df.at[idx, "End Time"] = match.group(2)
        # 當上下午時間段相同於處方判讀 / 化療處方判讀 / 藥物諮詢 / PreESRD（依關鍵字）者，可直接加入此欄位，詳細時間段則見上面時間對照表
        elif any(keyword in content for keyword in ["處方判讀", "化療處方判讀", "藥物諮詢", "PreESRD"]):
            for key, (start, end) in prescription_time_map.items():
                if key in content:
                    df.at[idx, "Start Time"] = start
                    df.at[idx, "End Time"] = end
                    break
        # 抗凝藥師門診 - 星期二與三不同時段
        elif "抗凝藥師門診" in content:
            if weekday == "二":
                df.at[idx, "Start Time"] = "08:30"
                df.at[idx, "End Time"] = "12:00"
            elif weekday == "三":
                df.at[idx, "Start Time"] = "13:30"
                df.at[idx, "End Time"] = "17:00"
        # 移植藥師門診 - 上午固定，下午預留
        elif "移植藥師門診" in content:
            if "上午" in content:
                df.at[idx, "Start Time"] = "08:30"
                df.at[idx, "End Time"] = "12:00"
            # 預留未來有下午時段使用
            # elif "下午" in content:
            #     df.at[idx, "Start Time"] = "13:30"
            #     df.at[idx, "End Time"] = "17:00"

        elif "PreESRD" in content:
            if "上午" in content:
                df.at[idx, "Start Time"] = "08:30"
                df.at[idx, "End Time"] = "12:00"
            # 預留未來有下午時段使用
            # elif "下午" in content:
            #     df.at[idx, "Start Time"] = "13:30"
            #     df.at[idx, "End Time"] = "17:00"

        # 中藥局調劑 - 不限定星期三早上
        elif "中藥局調劑" in content:
            df.at[idx, "Start Time"] = "08:30"
            df.at[idx, "End Time"] = "12:00"

        elif "瑞德西偉審核" in content:
            df.at[idx, "Start Time"] = "08:00"
            df.at[idx, "End Time"] = "20:00"

        # 非常班之諮詢與藥動服務-1：平日由處方判讀 7-住院自動新增非常班
        if "處方判讀 7-住院" in content and weekday in ["一", "二", "三", "四", "五"]:
            extra_rows.append({
                "日期": row["日期"],
                "星期": row["星期"],
                "工作內容": "非常班之諮詢與藥動服務",
                "簡化後內容": "非常班之諮詢與藥動服務",
                "Start Time": "17:30",
                "End Time": "21:30"
            })

        # 非常班之諮詢與藥動服務-2：假日已列出的非常班班別時間解析
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

# ---------- GUI 主程式 ---------- #
def run_gui():
    def select_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            file_path.set(path)

    def execute():
        name = code_entry.get().strip()
        file = file_path.get()
        if not name or not file:
            messagebox.showerror("錯誤", "請輸入班表代號並選擇檔案")
            return

        try:
            df = pd.read_excel(file, header=None)
            title = str(df.iat[0, 0])
            match = re.search(r"(\d{2,3})年(\d{1,2})月", title)
            if not match:
                raise ValueError("無法從標題擷取年月")
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
                    if name in cell:
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

            default_filename = f"{year_month}個人班表({name}).csv"
            path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=default_filename, filetypes=[("CSV files", "*.csv")])
            if not path:
                return

            df_output.to_csv(path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("成功", f"CSV 已匯出：{path}")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    root = tk.Tk()
    root.title("班表轉換工具")
    root.geometry("350x300")
    file_path = tk.StringVar()

    tk.Label(root, text="班表代號：",font=("微軟正黑體", 12)).pack(pady=5)
    code_entry = tk.Entry(root,font=("微軟正黑體", 12), width=10)
    code_entry.pack(pady=5)

    tk.Button(root, text="選擇班表檔案", command=select_file, font=("微軟正黑體", 12)).pack(pady=10)
    tk.Label(root, textvariable=file_path, font=("微軟正黑體", 12), wraplength=300, justify="left", fg="blue").pack(pady=2)

    tk.Button(root, text="執行轉換並儲存CSV", command=execute).pack(pady=10)
    tk.Button(root, text="關閉", command=root.destroy).pack(pady=10)
    root.mainloop()

if __name__ == '__main__':
    run_gui()
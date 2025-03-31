import streamlit as st
import pandas as pd
import os
import tempfile

st.set_page_config(page_title="Regular Report Creator", layout="wide")
st.title("📊 Semi-automatic Regular Report Creator")

st.markdown("""
**Pre-work：**

Before updating your files, please:
1. Download the latest files from **Acolaid** using conditions 
   - StatClass=(from Q1 to Q6 seperately) and；
   - Decision Date Is Null
2. Add column **Meeting Date** and its data manually.
2. Ensure files are named as `Q1.csv`, `Q2.csv`, etc.
3. Only CSV files are accepted.

---

""")

uploaded_files = st.file_uploader("Please select and upload multiple CSV files", type=["csv"], accept_multiple_files=True)

file_paths = {}
if uploaded_files:
    st.success(f"✅ Total uploaded {len(uploaded_files)} file(s)")
    st.write("file name：", [file.name for file in uploaded_files])
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        if filename.upper().startswith("Q") and filename.upper().endswith(".CSV"):
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            temp_file.write(uploaded_file.read())
            temp_file.close()
            file_paths[filename[:-4]] = temp_file.name  # 用 Q1、Q2 作键名

    if st.button("Create a Regular Report"):
        # 你的处理代码入口，从这里开始调用 file_paths 去读取数据
        # 你原始的脚本可以直接套进来，file_paths 已自动生成
        st.info("Script Running")

        # 插入原始处理脚本代码开始
        #!/usr/bin/env python
        # coding: utf-8
        
        
        import pandas as pd
        from datetime import datetime
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment, Font, PatternFill
        
        
        
        # === 第一步：读取 CSV 文件并保存为 UTF-8 ===

        
        
        dataframes = {}
        for name, path in file_paths.items():
            # 读取 ISO-8859-1 编码的 CSV 并直接保存在内存中（UTF-8 处理由 pandas 内部完成）
            df = pd.read_csv(path, encoding="ISO-8859-1")
            dataframes[name] = df
        
        # === 第二步：合并表格并计算时间列 ===
        merged_df = pd.concat(dataframes.values(), ignore_index=True)
        today = datetime.today()
        merged_df['Reg Date'] = pd.to_datetime(merged_df['Reg Date'], errors='coerce')
        merged_df['No. of weeks in system'] = merged_df['Reg Date'].apply(
            lambda x: (today - x).days // 7 if pd.notnull(x) else 'N/A'
        )
        merged_df['Expiry Date'] = pd.to_datetime(merged_df['Expiry Date'], errors='coerce')
        merged_df['No. of weeks past expiry date'] = merged_df['Expiry Date'].apply(
            lambda x: (today - x).days // 7 if pd.notnull(x) else 'N/A'
        )
        merged_df['Meeting Date'] = pd.to_datetime(merged_df['Meeting Date'], errors='coerce')
        merged_df['No. of weeks past meeting date'] = merged_df['Meeting Date'].apply(
            lambda x: (today - x).days // 7 if pd.notnull(x) else 'N/A'
        )
        
        # === 第三步：删除 PAS 类型行 ===
        filtered_df = merged_df[merged_df['App Type'] != 'PAS'].copy()
        
        
        # === 第四步：提取和排序字段 ===
        desired_columns = [
            "Application Number", "Application Address", "Officer", "Reception Date", "Reg Date",
            "No. of weeks in system", "Expiry Date", "No. of weeks past expiry date", "Meeting Date", 
            "No. of weeks past meeting date", "PPA",
            "App Type", "Validation Code", "Finalised Decision Level", "StatClass",
            "Agent Name", "Applicant Name", "Proposal"
        ]
       # 自动识别 Rec Date / Reception Date / Received Date 等等
         for col in filtered_df.columns:
             if "rec" in col.lower() and "date" in col.lower():
                 filtered_df = filtered_df.rename(columns={col: "Reception Date"})

       
        column_renames = {
            "CaseFullRef": "Application Number",
            "Application Address": "Application Address",
            "Officer": "Officer",
            "Rec Date": "Reception Date",
            "Reg Date": "Reg Date",
            "No. of weeks in system": "No. of weeks in system",
            "Expiry Date": "Expiry Date",
            "No. of weeks past expiry date": "No. of weeks past expiry date",
            "Meeting Date": "Meeting Date",
            "No. of weeks past meeting date": "No. of weeks past meeting date",
            "PPA.1": "PPA",
            "App Type": "App Type",
            "CaseResub": "Validation Code",
            "Decision Level": "Finalised Decision Level",
            "StatClass": "StatClass",
            "Agent Name": "Agent Name",
            "Applicant Name": "Applicant Name",
            "Proposal": "Proposal",
        }
        processed_df = filtered_df.rename(columns=column_renames)
        for col in desired_columns:
            if col not in processed_df.columns:
                processed_df[col] = ""
        processed_df = processed_df[desired_columns]
        

        # === 第五步：格式化日期并排序 ===
        date_format = "%d %b %Y"
        for col in ['Reception Date', 'Reg Date', 'Expiry Date', 'Meeting Date']:
            processed_df[col] = pd.to_datetime(processed_df[col], errors='coerce') \
                .dt.strftime(date_format).fillna("")
        processed_df['Reception Date (parsed)'] = pd.to_datetime(processed_df['Reception Date'], format="%d %b %Y", errors='coerce')
        processed_df = processed_df.sort_values(by='Reception Date (parsed)', ascending=True).drop(columns=['Reception Date (parsed)'])

        
        # === 第六步：导出初始 Excel 文件 ===
        output_path = "Final_Planning_Table.xlsx"
        processed_df.to_excel(output_path, index=False, engine='openpyxl')
        

        
        # === 第七步：高亮 Proposal 列中符合条件的单元格（黄色）===
        wb = load_workbook(output_path)
        ws = wb.active
        header = [cell.value for cell in ws[1]]
        proposal_col_index = header.index("Proposal") + 1
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for row in ws.iter_rows(min_row=2, min_col=proposal_col_index, max_col=proposal_col_index):
            cell = row[0]
            if isinstance(cell.value, str):
                value = cell.value.lower()
                contains_target = any(keyword in value for keyword in ['commercial', 'student', 'student accommodation'])
                lacks_dwell = not any(word in value for word in ['dwell', 'dwelling', 'resident', 'residential', 'c3'])
                if contains_target and lacks_dwell:
                    cell.fill = yellow_fill
        

        
        # === 第八步：对三列中最大的前10个数值做浅红高亮 ===
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        target_columns = [
            "No. of weeks in system",
            "No. of weeks past expiry date",
            "No. of week past meeting date"
        ]
        for col_name in target_columns:
            if col_name in header:
                col_idx = header.index(col_name) + 1
                values_with_cells = []
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    cell = row[0]
                    try:
                        value = float(cell.value)
                        values_with_cells.append((value, cell))
                    except:
                        continue
                top_cells = sorted(values_with_cells, key=lambda x: x[0], reverse=True)[:10]
                for _, cell in top_cells:
                    cell.fill = light_red_fill
        
        # === 第九步：设置格式为表格（浅蓝绿色，Style Light 12）===
        if ws.tables:
            for tbl in list(ws.tables.values()):
                del ws.tables[tbl.name]
        max_row = ws.max_row
        max_col = ws.max_column
        table_range = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName="PlanningDataTable", ref=table_range)
        style = TableStyleInfo(
            name="TableStyleLight12",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)
        
        
        last_row = ws.max_row
        last_col = ws.max_column
        
        # 第三步：合并最后一行下方的整行
        merge_range = f"A{last_row + 1}:{get_column_letter(last_col)}{last_row + 1}"
        ws.merge_cells(merge_range)
        
        # 第四步：设置单元格样式和文本
        footer_cell = ws.cell(row=last_row + 1, column=1)
        footer_cell.value = f"This report covers all applications received up to {datetime.today().strftime('%d %b %Y')}."
        footer_cell.alignment = Alignment(horizontal="center", vertical="center")
        footer_cell.font = Font(bold=True, italic=True)
        footer_cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # === 第十步：保存最终版本 ===
        wb.save(output_path)
        print(f"✅ Excel file have been saved：{output_path}")
        
        
        # 插入原始处理脚本代码结束
        st.success("✅ Data processing completed!")

        today_str = datetime.today().strftime("%d_%b_%Y")
        output_filename = f"Regular_Report_on_Pending_Application_{today_str}.xlsx"

        if os.path.exists(output_path):
           with open(output_path, "rb") as f:
              st.download_button(
                 "Click here to download the result file",
                 f,
                 file_name=output_filename
              )

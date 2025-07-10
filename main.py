import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


# 欲尋找的欄位名稱
target_columns = ['Process', 'Customer', 'Machine require(set/Wk)', 'Machine O/H(set)', 'Idling tester(set)']

# 讀取 Excel 檔案（保留巨集）
source = '250702162359805.xlsm'
wb = load_workbook(source, keep_vba=True)

# 要處理的來源工作表
source_sheets = ['CP Summary', 'FT Summary']
ws_source = wb['CP Summary']
output_data = []

# for sheet_name in source_sheets:
#     ws = wb[sheet_name]
    
#     col_pos = {}  # 儲存找到的欄位名稱 → (row, column)
    
#     # 遍歷整個工作表找欄位名稱
#     for row in ws.iter_rows():
#         for cell in row:
#             if cell.value in target_columns and cell.value not in col_pos:
#                 col_pos[cell.value] = (cell.row, cell.column)
    
#     # 如果沒找到所有欄位，就略過這張工作表
#     #if len(col_pos) != len(target_columns):
#     #    print(f"[跳過] 無法在 {sheet_name} 找到所有欄位：{target_columns}")
#     #    continue

#     # 取出欄號並依照 target_columns 排列順序
#     col_indices = [col_pos[col][1] for col in target_columns]
#     start_row = max([col_pos[col][0] for col in target_columns]) + 1

#     # 從 start_row 開始往下抓資料
#     for r in range(start_row, ws.max_row + 1):
#         row_data = []
#         empty_row = True
#         for col in col_indices:
#             value = ws.cell(row=r, column=col).value
#             row_data.append(value)
#             if value is not None:
#                 empty_row = False
#         if empty_row:
#             break  # 遇到空白列就停止
        # output_data.append(row_data)

# 清除或建立寫入用工作表
if 'Sheet3' in wb.sheetnames:
    wb.remove(wb['Sheet3'])
ws_target = wb.create_sheet('Sheet3')

# -------- 標題樣板設定 --------
def write_table_block(ws, start_row, data_row, month_data):
    # ----------- Step 1：先填入所有標題與資料內容 -----------

    # 標題列（主欄位）
    ws.cell(row=start_row, column=1, value="Process")
    ws.cell(row=start_row, column=2, value="Tester")
    ws.cell(row=start_row, column=3, value="Customer")
    ws.cell(row=start_row, column=4, value="Month")
    ws.cell(row=start_row, column=5, value="6M FCST (pcs/wk)")
    ws.cell(row=start_row, column=11, value="Summary")

    # 標題列（次欄位）
    ws.cell(row=start_row+2, column=4, value="MachineO/H")
    ws.cell(row=start_row+3, column=4, value="Machine require(set/Wk)")
    ws.cell(row=start_row+4, column=4, value="idling tester")

    # 月份欄位
    for col in range(5, 11):
        for i in range(0, 6):
            ws.cell(row=start_row+1, column=col, value=month_data[0][i])
            
    ws.cell(row=start_row+1, column=5, value="Wk1") # 首月   值從表格抓
    ws.cell(row=start_row+1, column=6, value="Wk2") # 次月
    ws.cell(row=start_row+1, column=7, value="Wk3") # 第三月
    ws.cell(row=start_row+1, column=8, value="Wk3") # 第四月
    ws.cell(row=start_row+1, column=9, value="Wk3") # 第五月
    ws.cell(row=start_row+1, column=10, value="Wk3") # 第六月

    # 資料列（你自己的資料）
    for col in range(0, 10):
        ws.cell(row=start_row+2, column=col+1, value="")

    ws.cell(row=start_row+1, column=1, value=data_row[0])  # Process
    ws.cell(row=start_row+1, column=2, value=data_row[1])  # Tester
    ws.cell(row=start_row+1, column=3, value=data_row[2])  # Customer

    ws.cell(row=start_row+2, column=4, value=data_row[3])  # Machine O/H
    ws.cell(row=start_row+3, column=4, value=data_row[4])  # Machine require
    ws.cell(row=start_row+4, column=4, value=data_row[5])  # Idling tester

    ws.cell(row=start_row+1, column=5, value=data_row[6])  # Wk1
    ws.cell(row=start_row+1, column=6, value=data_row[7])  # Wk2
    ws.cell(row=start_row+1, column=7, value=data_row[8])  # Wk3
    ws.cell(row=start_row+1, column=8, value=data_row[9])  # Summary

    # ----------- Step 2：再合併儲存格（一定要最後做） -----------
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+2, end_column=1)  # Process
    ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row+2, end_column=2)  # Tester
    ws.merge_cells(start_row=start_row, start_column=3, end_row=start_row+2, end_column=3)  # Customer
    ws.merge_cells(start_row=start_row, start_column=4, end_row=start_row, end_column=4)    # Month title
    ws.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=7)    # 6M FCST
    ws.merge_cells(start_row=start_row, start_column=11, end_row=start_row+2, end_column=8)  # Summary

    # ----------- Step 3：設定樣式（置中對齊 + 粗體） -----------
    for r in range(start_row, start_row + 4):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if r == start_row:
                cell.font = Font(bold=True)

    # ----------- Step 4：填入該筆資料（依照欄位順序） -----------
    ws.cell(row=start_row+1, column=1, value=data_row[0])  # Process
    ws.cell(row=start_row+1, column=2, value=data_row[1])  # Tester
    ws.cell(row=start_row+1, column=3, value=data_row[2])  # Customer

    ws.cell(row=start_row+2, column=5, value=data_row[3])  # MachineO/H
    ws.cell(row=start_row+3, column=5, value=data_row[4])  # Machine require
    ws.cell(row=start_row+4, column=5, value=data_row[5])  # idling tester

    ws.cell(row=start_row+1, column=5, value=data_row[6])  # Wk1
    ws.cell(row=start_row+1, column=6, value=data_row[7])  # Wk2
    ws.cell(row=start_row+1, column=7, value=data_row[8])  # Wk3
    ws.cell(row=start_row+1, column=8, value=data_row[9])  # Summary

# -------- 逐筆寫入表格區塊 --------
current_row = 1
for row in ws_source.iter_rows(min_row=2, values_only=True):
    if all(cell is None for cell in row):
        continue  # 跳過空列

    write_table_block(ws_target, current_row, row)
    current_row += 5  # 表格高 4 列，空 1 列

# # 寫入資料
# for row in output_data:
#     ws3.append(row)

# 儲存結果
output_file = f"{source.split('.')[0]}_export.xlsm"
wb.save(output_file)




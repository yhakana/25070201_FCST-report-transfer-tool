from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 欄位寬度、格式
col_widths = [14, 14, 16, 25, 12, 12, 12, 12, 12, 12, 14]
header_fill = PatternFill(start_color="FDE8D7", end_color="FDE8D7", fill_type="solid")
bold_font = Font(bold=True)
thin = Side(border_style="thin", color="000000")
border = Border(top=thin, left=thin, right=thin, bottom=thin)

def find_header_row(ws, target_headers, search_rows=20):
    """
    在前 search_rows 行中尋找同時包含所有 target_headers 的那一列
    回傳 (標題row的index, header_map)
    """
    for row in ws.iter_rows(min_row=1, max_row=search_rows, values_only=True):
        header_map = {}
        for col_idx, cell_value in enumerate(row, 1):
            if cell_value in target_headers:
                header_map[cell_value] = col_idx
        # 若該列至少包含所有必須欄位
        if all(h in header_map for h in target_headers):
            return ws.iter_rows(min_row=1, max_row=search_rows).index(row)+1, header_map
    raise ValueError("找不到符合所有標題的那一列")

def find_multilevel_header(ws, target_headers, month_candidates, search_limit=30):
    """
    自動找出多層標題的起始row index（回傳: (row_idx1, row_idx2)）
    key_headers: 主要標題名清單（如 Process、Tester、Machine O/H...）
    month_candidates: 月份標題可能的字串list
    search_limit: 最多搜尋前幾列
    """
    for r in range(1, search_limit):
        row1 = [str(cell.value).strip() if cell.value else "" for cell in ws[r]]
        hits = [h for h in target_headers if h in row1]
        if len(hits) >= 2:  # 至少2個以上大標題
            row2 = [str(cell.value).strip() if cell.value else "" for cell in ws[r+1]]
            month_hits = [m for m in month_candidates if m in row2]
            if len(month_hits) >= 2:  # 至少2個月份
                return r, r+1
    raise ValueError("找不到雙列標題的位置")

def get_header_col_map(ws):
    header_map = {}
    for col, cell in enumerate(ws[1], 1):  # ws[1] 代表第一列
        header_map[cell.value] = col
    return header_map

def get_all_data(ws, header_row_idx, header_map, target_headers):
    data_rows = []
    for row in ws.iter_rows(min_row=header_row_idx+1, values_only=True):
        data = {}
        for h in target_headers:
            col_idx = header_map.get(h)
            if col_idx:
                data[h] = row[col_idx-1]
            else:
                data[h] = None
        data_rows.append(data)
    return data_rows

def fill_table(ws, start_row, data_dict):
    # 填入標準欄位，這裡以 Process, Tester, Customer, Month, Summary 為例
    ws.cell(row=start_row+2, column=1, value=data_dict.get('Process'))    # A3
    ws.cell(row=start_row+2, column=2, value=data_dict.get('Tester'))     # B3
    ws.cell(row=start_row+2, column=3, value=data_dict.get('Customer'))   # C3
    ws.cell(row=start_row+2, column=4, value=data_dict.get('Month'))      # D3
    ws.cell(row=start_row+2, column=11, value=data_dict.get('Summary'))   # K3
    # 依需求也可填入其他欄位

def draw_table(ws, start_row):
    try:
        # 設定欄寬
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width

        # 合併儲存格
        try:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+1, end_column=1)   # A1:A2
            ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row+1, end_column=2)   # B1:B2
            ws.merge_cells(start_row=start_row, start_column=3, end_row=start_row+1, end_column=3)   # C1:C2
            ws.merge_cells(start_row=start_row, start_column=4, end_row=start_row+1, end_column=4)   # D1:D2
            ws.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=10)    # E1:J1
            ws.merge_cells(start_row=start_row, start_column=11, end_row=start_row+1, end_column=11) # K1:K2
            ws.merge_cells(start_row=start_row+2, start_column=1, end_row=start_row+4, end_column=1)    # A3:A5
            ws.merge_cells(start_row=start_row+2, start_column=2, end_row=start_row+4, end_column=2)    # B3:B5
            ws.merge_cells(start_row=start_row+2, start_column=3, end_row=start_row+4, end_column=3)    # C3:C5
            ws.merge_cells(start_row=start_row+2, start_column=11, end_row=start_row+4, end_column=11)  # K3:K5
        except Exception as e:
            print(f"合併儲存格失敗，起始列 {start_row}，錯誤訊息：{e}")

        # 標題內容與格式
        try:
            ws.cell(row=start_row, column=1, value="Process")
            ws.cell(row=start_row, column=2, value="Tester")
            ws.cell(row=start_row, column=3, value="Customer")
            ws.cell(row=start_row, column=4, value="Month")
            ws.cell(row=start_row, column=5, value="6M FCST ( pcs/wk )")
            ws.cell(row=start_row, column=11, value="Summary")

            for col in range(1, 12):
                cell = ws.cell(row=start_row, column=col)
                cell.fill = header_fill
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for col in range(5, 11):
                cell = ws.cell(row=start_row+1, column=col)
                cell.fill = header_fill
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for col in range(1, 12):
                cell = ws.cell(row=start_row+1, column=col)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
        except Exception as e:
            print(f"設定標題內容或格式失敗，起始列 {start_row}，錯誤訊息：{e}")

        # 直排內容與格式
        try:
            ws.cell(row=start_row+2, column=4, value="MachineO/H")
            ws.cell(row=start_row+3, column=4, value="Machine require(set/Wk)")
            ws.cell(row=start_row+4, column=4, value="idling tester")
            for row in range(start_row+2, start_row+5):
                cell = ws.cell(row=row, column=4)
                cell.fill = header_fill
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = border
        except Exception as e:
            print(f"設定直排內容或格式失敗，起始列 {start_row}，錯誤訊息：{e}")

        # 補其餘區塊邊框
        try:
            for r in range(start_row, start_row+5):
                for c in range(1, 12):
                    ws.cell(row=r, column=c).border = border
        except Exception as e:
            print(f"設定邊框失敗，起始列 {start_row}，錯誤訊息：{e}")

    except Exception as e:
        print(f"draw_table 發生不可預期錯誤，起始列 {start_row}，錯誤訊息：{e}")

# 主程式區段
source = '250702162359805.xlsm'
try:
    wb = load_workbook(source, keep_vba=True)
except Exception as e:
    print(f"檔案讀取失敗：{e}")
    raise

if 'SummaryTable' in wb.sheetnames:
    wb.remove(wb['SummaryTable'])
ws_summary_table = wb.create_sheet('SummaryTable')
ws_cp = wb['CP Summary']
ws_ft = wb['FT Summary']

target_headers = ['Process', 'Tester', 'Customer', '6M FCST ( pcs/wk )', 'Summary', 'Machine require(set/Wk)','Machine O/H (set)', 'Idling tester (set)']
month_candidates = ["Jun'25", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "2026-Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"]

# 找欄位index

# header_row_cp, header_map_cp = find_multilevel_header(ws_cp, target_headers, month_candidates)
# cp_data = get_all_data(ws_cp, header_row_cp, header_map_cp, target_headers)

# header_row_ft, header_map_ft = find_header_row(ws_ft, target_headers)
# ft_data = get_all_data(ws_ft, header_row_ft, header_map_ft, target_headers)

# 決定總表格數
# all_data = cp_data + ft_data  # 若要分兩區也可分開

all_data = 5
table_height = 5
for idx in range(0, 10):
    base_row = 1 + idx * (table_height + 1)
    draw_table(ws_summary_table, base_row)
    #fill_table(ws_summary_table, base_row, data_dict)

# 儲存檔案
output_file = f"{source.rsplit('.', 1)[0]}_test.xlsm"
try:
    wb.save(output_file)
except Exception as e:
    print(f"儲存檔案時發生錯誤：{e}")

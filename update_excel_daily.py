"""
TSMC Excel 每日更新腳本 - 2026-07-01
執行：python update_excel_daily.py
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date

EXCEL_PATH = r"C:\Users\K748\OneDrive - 財團法人中華民國對外貿易發展協會\FET\Stock分析\TSMC_股市分析報告.xlsx"

def make_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(cell, bg="FFFFFF", align="center", bold=False, color="000000"):
    cell.font = Font(name="Arial", size=10, bold=bold, color=color)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = make_border()

today_str = '2026-10-06'
tw_price = 'NT$2,575（10/5收，+75；開2,550高2,580低2,545，收盤與盤中高皆創歷史新高）；10/2收NT$2,500（-10）'
change_pct = '+3.00%（10/5收，+75.00，領先大盤0.45pp；加權+2.55%創收盤新高）；10/2為-0.40%（-10.00，落後大盤0.65pp）'
nyse_price = 'US$485.80（10/5收，+2.75%，近兩年最高收盤；ADR溢價+19.95%）；10/2收US$472.78（+2.96%）'
volume = '26,800（10/5收，+69.7%，高於5日均量24,887張7.7%；2026年182日中由大到小第138）；10/2為15,792張（2026年第9低）'
news_summary = """報告日2026-10-06。本報涵蓋10/2與10/5兩個台股交易日，10/5之日報未產製。10/2收2,500(-0.40%)，外資賣超5,913張觸發退回中性；10/5收2,575(+3.00%)創歷史新高、量26,800張、外資買超9,770張，上調條件成立、回到中性偏多，因短線過熱不再上調。"""
change_color = '00B050'
try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb["每日更新記錄"]

    # 若最後一行日期 = 今日，更新該行；否則新增一行（避免重複 row）
    last_row = ws.max_row
    last_date = ws.cell(last_row, 1).value
    target_row = last_row if last_date == today_str else last_row + 1
    mode = "更新" if target_row == last_row else "新增"

    row_data = [
        today_str,
        tw_price,
        change_pct,
        nyse_price,
        volume,
        news_summary,
        "Yahoo Finance / Bloomberg"
    ]
    bg = "E9F2FB" if target_row % 2 == 0 else "FFFFFF"
    for col, val in enumerate(row_data, 1):
        c = ws.cell(row=target_row, column=col, value=val)
        if col == 3:
            style_cell(c, bg=bg, align="center", color=change_color, bold=True)
        elif col == 6:
            style_cell(c, bg=bg, align="left")
        else:
            style_cell(c, bg=bg)

    ws["A2"] = f"最後更新：{today_str}"
    wb["封面總覽"]["A3"] = f"報告更新日期：{today_str}"
    wb.save(EXCEL_PATH)
    print(f"Excel {mode}成功：第 {target_row} 行（{today_str}）")
except Exception as e:
    print(f"Excel 更新失敗：{e}")

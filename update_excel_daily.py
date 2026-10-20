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

today_str = '2026-10-20'
tw_price = 'NT$2,420（10/19收，平盤0.00、0.00%；開2,395高2,420低2,390，開低25元、盤中2,390跌破季線MA60 2,421.5後尾盤拉回收在全日最高的紅K，2,450–2,490缺口第3日未回補，收盤連2日低於MA60 1.5元）；10/16收NT$2,420（-0.82%）'
change_pct = '0.00%（10/19收，平盤，相對大盤+0.61pp；加權-289.18點、-0.61%收46,822.08，全市場成交8,667.12億）；10/16為-0.82%'
nyse_price = 'US$420.22（10/19收，+1.84%；終止連3跌、五日-5.43%；ADR溢價擴大至+12.70%，新台幣收32.445、取32.45）；10/16收US$412.61（-2.28%）'
volume = '27,695（10/19收，-35.7%；低於前報基準38,800張28.6%、低於新5日均量39,627張30.1%；10/12以來最低，2026年191日中第138）；10/16為43,047張'
news_summary = """10/19盤中2,390破季線後拉回平盤2,420，量縮27,695張，外資賣超縮至9,834張連9賣、投信買3,884張，維持看空；費半+1.61%終止連8跌、ADR+1.84%，夜盤中性。"""
change_color = '000000'
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

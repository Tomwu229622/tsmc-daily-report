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

today_str = '2026-11-02'
tw_price = 'NT$2,370（10/30收，-45、-1.86%；開低2,380即全日最高、低2,360，實體10元黑K；跌破半年線MA120 2,389.83（結束連78日在上）、跌回MA5 2,399之下，收在五條均線之下，連9日低於MA60 2,431.58（差61.58元）；留2,380–2,390向下缺口，2,450–2,490缺口第10日未回補）；10/29收NT$2,415（-0.41%）'
change_pct = '-1.86%（10/30收，相對大盤+0.63pp（精算+0.6308）；加權-1,156.50點、-2.49%收45,211.83，全市場成交9,396.05億）；10/29為-0.41%'
nyse_price = 'US$406.13（10/30收，-5.95%；連2日下跌；ADR溢價同日對齊+11.32%（ADR 10/30對台股10/30與匯率32.48），10/29同日對齊+16.19%；盤後408.79（+0.65%，19:59 EDT））；10/29收US$431.81（-1.75%）'
volume = '36,616（10/30收，+98.96%；高於前報基準27,067張35.28%、新5日均量28,971張；2026年198日中第86）；10/29為18,404張'
news_summary = """10/30收2,370跌1.86%，放量跌破半年線、收在五條均線下；外資賣超8,632張連16賣，維持看空；英特爾-28.58%、費半-7.20%、ADR-5.95%、溢價+11.32%。"""
change_color = 'FF0000'
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

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

today_str = '2026-11-27'
tw_price = 'NT$2,145（11/26收，-5、-0.23%；開2,150（平開）／高2,170／低2,120，實體5元的小黑K；盤中最高恰觸及2,170未回補2,170–2,230缺口；守在MA5 2,111之上、續在MA10 2,156.50之下；MACD柱-7.33續負收斂、KD黃金交叉延續（K 40.63／D 37.88）、RSI 42.81；外資賣超18,134張（仍在門檻之外、較前日縮小8,837張）、投信買超1,176張、自營買超1,461張、三大合計-15,497張；量縮下跌（字面與實質一致）；雙條件0/2、維持看空、轉弱程度改判持平）'
change_pct = '-0.23%（11/26收，相對大盤-2.74pp（精算-2.7365），轉為落後；加權+956.38點、+2.50%收39,151.81，全市場成交8,453.37億）；11/25為-3.59%'
nyse_price = 'US$342.34（11/25收盤沿用；11/26美股感恩節休市、無新收盤；11/25為-1.96%；費半11/25為-0.65%；ADR溢價混合日期口徑+5.18%（ADR 11/25對台股11/26與匯率32.95），同日口徑11/25為+5.50%；11/25盤後約347（約+1.5%，19:59 ET）；11/27美股提早收盤）'
volume = '45,893（11/26收，-45.95%；為前報基準62,450張的0.73倍＝量縮下跌，剔除11/24鎖漲停極小量的對照基準亦為量縮下跌；新5日均量60,312張；2026年217日中由大到小第57）；11/25為84,912張'
news_summary = """11/26加權大漲2.50%，台積電平開小跌收2,145(-0.23%)、相對轉落後、盤中最高觸及2,170未回補缺口；外資賣超18,134張仍在門檻之外、投信與自營買超；量縮下跌；外資台指期淨空再創窗口最大；美股11/26感恩節休市（ADR沿用11/25）；行政院稱美301擬議稅率15%不疊加MFN；維持看空、轉弱程度改判持平。"""
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

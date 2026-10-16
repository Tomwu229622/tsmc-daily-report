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

today_str = '2026-10-16'
tw_price = 'NT$2,440（10/15收，-50、-2.01%；開2,420高2,450低2,420，跳空開低70元、開盤即最低、收盤較開盤+20紅K，留2,450–2,490缺口）；10/14收NT$2,490（平盤）'
change_pct = '-2.01%（10/15收，相對大盤+1.40pp；加權-1,655.96點、-3.41%收46,969.87，全市場成交1兆4,357.55億）；10/14為0.00%'
nyse_price = 'US$422.25（10/15收，-1.80%，法說會後首個美股盤、盤中-5.3%後收斂；連2跌、9/16以來最低、五日-7.80%；ADR溢價回升至+11.69%）；10/14收US$430.00（-3.40%）'
volume = '58,487（10/15收，+108.0%；高於前報基準26,510張120.6%、高於新5日均量34,819張68.0%；成交筆數512,176筆、2026年第7多；2026年189日中第15、7/31以來最大）；10/14為28,112張'
news_summary = """報告日2026-10-16（台美皆為10/15）。10/15跳空開低收2,440（-2.01%），放量58,487張、外資賣超34,111張（2026年第4大、連7賣），跌破MA20中軌、MA5破MA10、MACD柱-17.43，外資台指期淨空-97,582口創2026年新深，立場維持看空；收盤後Q3法說：營收US$46.72B超指引上緣、毛利率66.4%、EPS NT$29.43創高、Q4財測$52.3–54.3B、全年成長上修至約44%；ADR-1.80%收422.25、費半連7跌；夜盤個股期2,453偏強。"""
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

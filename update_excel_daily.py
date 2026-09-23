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

today_str = '2026-09-23'
tw_price = 'NT$2,460（9/22收，收在全日最低；盤中高2,510為6/23以來最高）'
change_pct = '-0.81%（9/22收，-20.00，落後大盤+0.17%）'
nyse_price = 'US$452.00（9/22收，+1.54%，6/30以來最高收盤）'
volume = '22,010（9/22收，+36.8%，量增但收黑K）'
news_summary = """報告日2026-09-23(Wed)，台股資料為2026-09-22(Tue)官方收盤；ADR、費半與美股個股、韓股為9/22正式收盤；期貨為TAIFEX 9/22日盤，另含今晨夜盤(9/22 15:00至9/23約05:00)。

(1)⚠️⚠️開高走低收最低：台股2330開2,505(跳空+25)、盤中高2,510(6/23以來最高盤中價)、收2,460＝全日最低(-20、-0.81%)，開收跌幅45元為7/29以來最大；⚠️惟「跳空開高、收最低」在2026年已出現18次，非罕見型態。媒體歸因於9/25中秋、9/28教師節連假前獲利了結(本報無法證實動機)。

(2)⭐⭐前報六項檢核：量能字面通過(22,010張>17,585、收盤守2,445)但量增出現在黑K、實質存疑；布林上軌未通過(盤中越2,490上軌、收盤退回)；相對強弱未通過且觸發確認(連3日落後加權+0.17%，落後0.98pp)；外資連四買未通過(轉賣4,329張約NT$106.51億，而全市場外資淨買超NT$453.21億)；夜盤只兌現開盤方向。

(3)⭐技術：MA5 2,441穿越MA10 2,430(8/26以來首度)、五線恢復完整多頭排列(9/14以來首度)、收盤連4日站上全部均線；MACD柱+6.87續擴大(屬EMA慣性)、RSI 56.30中止連3升、J 80.98回落。

(4)⭐籌碼與衍生品：投信-405張、自營+467張(連7日)、三大法人-4,267張中止連3買；個股期CDF日盤2,489、基差+29；外資台指期淨空擴大至-75,568口；今晨夜盤CDF收2,515(+26)。ADR溢價擴大至+16.64%，新台幣31.7384連3日升值。

(5)立場自「中性偏多」下修為「中性」。次日檢核：收盤能否守2,445(跌破則升級為短線反轉)、外資賣超是否擴大；9/23週結算、9/24連假前最後交易日、美光美東9/30財報。"""
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

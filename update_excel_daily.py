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

today_str = "2026-07-23"
tw_price = "NT$2,400（7/22收）"   # 台股2330 7/22 Wed官方收NT$2,400(-10、-0.41% vs 7/21 NT$2,410、開2,440/高2,445/低2,385、量31,653張TWSE口徑、成交額762.06億、成交136,194筆、TWSE確認)——開高走低尾盤翻黑、反彈第三日高檔震盪、20MA得而復失惟守住5MA
change_pct = "-0.41%（7/22收）"    # 7/22官方漲跌幅-0.41%（TWSE、跳空開高+30後獲利了結賣壓壓回、高低震盪60點）
nyse_price = "US$421.21（7/22收）" # NYSE TSM 7/22 Wed收$421.21(-$3.40、-0.80% vs 7/21 $424.61、Yahoo API確認)——連日大漲後縮量整理；SOX同日+0.44%收12,410.67連2紅(自6/22史高-15.2%)
volume = "31,653（7/22收）"    # 台股7/22官方成交量31,653張（TWSE口徑、持平前日、續攻量能未放大；成交筆數136,194大增、換手活絡）、成交額762.06億
news_summary = "報告日2026-07-23(Thu)。📊反彈第三日轉高檔震盪：台股2330 7/22官方收NT$2,400(-10、-0.41%、開2,440/高2,445/低2,385、量31,653張、額762.06億、136,194筆、TWSE確認)——跳空開高一度站上20MA 2,414、獲利了結賣壓壓回尾盤翻黑、高低震盪60點；惟低點守住5MA 2,378屬強勢整理。⭐大盤同日續漲+592.91點(+1.34%)收44,825.78、兩日收復7/17崩跌逾七成——惟中小型股領漲(櫃買+3.46%)、2330落後大盤。💰頭條：CFO黃仁昭定調2nm已於Q2貢獻營收、Q3起放大成最新成長引擎——與2027漲價傳聞形成量價齊揚敘事。📊籌碼兩層分化：外資對2330回賣-3,539張(T86官方、7/21轉買僅一日)、投信+1,562(連5買擴大)、自營-1,059(終結連9買)、三大合計-3,036張；⭐全市場BFI82U：三大法人買超338.5億——外資+173.4億正式終結連13賣、投信+189.0億。✅NYSE TSM 7/22縮量整理收$421.21(-0.80%)、SOX +0.44%收12,410.67連2紅(自史高-15.2%)——客戶強(NVDA+2.30%/AVGO+2.67%)設備弱(AMAT-1.88%/MU-1.17%)換手分化。🚀台股供應鏈補漲：華邦電漲停+9.97%、聯發科+4.90%收3,850、日月光+3.63%；韓股7/23早盤SK+4.9%/三星+3.6%續強。📊ADR溢價+13.6%(匯率~32.35)。技術面(官方收盤計算)：⭐KDJ低檔金叉正式兌現(K49.4上穿D44.8、J58.7)、RSI 50.7持穩中軸、MACD綠柱三日連續收斂-12.8；支撐2,385/2,378(5MA)/2,345/2,337(60MA)、壓力2,414(20MA)/2,440-2,445/2,470/2,519-2,535——站回20MA與放量4萬張為續航雙確認。⚠️保留：量能連2日僅3.2萬張、外資買賣反覆、分析師籲獲利了結。"
change_color = "FF0000"        # 紅色（7/22官方-0.41%下跌）

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

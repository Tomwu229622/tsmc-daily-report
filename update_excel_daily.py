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

today_str = "2026-07-24"
tw_price = "NT$2,405（7/23收）"   # 台股2330 7/23 Thu官方收NT$2,405(+5、+0.21% vs 7/22 NT$2,400、開2,385/高2,405/低2,370、量28,001張TWSE口徑、成交額668.34億、成交142,112筆、TWSE確認)——平低開後盤中低2,370一度跌破前日5MA、尾盤拉升收在當日最高、反彈第四日量縮整理
change_pct = "+0.21%（7/23收）"    # 7/23官方漲跌幅+0.21%（TWSE、低開走高尾盤拉升收最高、連2日收斂於2,400-2,405窄幅）
nyse_price = "US$415.58（7/23收）" # NYSE TSM 7/23 Thu收$415.58(-$5.63、-1.34% vs 7/22 $421.21、Yahoo API確認)——AI權值回檔；SOX同日-0.54%收12,343.84終結連2紅(自6/22史高-15.7%)、記憶體/設備逆勢強MU+3.20%
volume = "28,001（7/23收）"    # 台股7/23官方成交量28,001張（TWSE口徑、量續縮前日31,653張、續攻量能仍未放大；成交筆數142,112續增、小單換手活絡）、成交額668.34億
news_summary = "報告日2026-07-24(Fri)。📊反彈第四日量縮整理、尾盤拉升守穩：台股2330 7/23官方收NT$2,405(+5、+0.21%、開2,385/高2,405/低2,370、量28,001張、額668.34億、142,112筆、TWSE確認)——平低開後盤中低2,370一度跌破前日5MA、尾盤買盤進場拉升收在當日最高；連2日收斂於2,400-2,405窄幅、賣壓遞減但攻擊量能同樣缺席。⭐大盤同日尾盤翻紅+25.03點(+0.06%)收44,850.81連3紅——成交額縮至9,306億(前日1.026兆)、2330尾盤拉升為指數翻紅主力、反彈自普漲進入輪動整理期。📊籌碼兩層分化延續：外資對2330連2賣且擴大-4,614張(前日-3,539、T86官方)、投信-943終結連5買、自營+38、三大合計-5,519張；⭐惟全市場BFI82U：三大法人買超185.2億——外資+69.6億連2買、投信+73.7億、自營+41.9億——資金留在台股、只是不在權值。⚠️NYSE TSM 7/23收$415.58(-1.34%)、SOX -0.54%收12,343.84終結連2紅(自史高-15.7%)——板塊再換手：記憶體/設備強(MU+3.20%/KLAC+1.88%/AMAT+1.60%)vs客戶端回吐(NVDA-1.56%/AMD-2.29%/AAPL-1.30%)。⚠️韓股7/23大漲(SK海力士+4.86%/三星+3.65%)後7/24早盤雙雙重挫(SK-5.1%/三星-4.3%)——今日台股半導體開盤逆風；VIX+12.4%回升18.70、WTI急漲~$92。📊ADR溢價自+13.6%收斂至+11.6%(匯率~32.30)。技術面(官方收盤計算)：KDJ金叉後開口擴大(K54.2/D47.9/J66.8)、RSI 51.1中軸上續升、MACD綠柱四日連續收斂-10.5；支撐2,370(7/23低)/2,365(5MA)/2,345/2,340(60MA)、壓力2,415(20MA)/2,440-2,445/2,470/2,519-2,535——站上20MA(距僅10點)與放量4萬張為續航雙確認。⚠️保留：量能連3日2.8-3.2萬張未放大、外資對2330連2賣、韓股7/24早盤重挫。"
change_color = "00B050"        # 綠色（7/23官方+0.21%上漲）

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

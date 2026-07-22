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

today_str = "2026-07-22"
tw_price = "NT$2,410（7/21收）"   # 台股2330 7/21 Tue官方收NT$2,410(+90、+3.88% vs 7/20 NT$2,320、開2,350/高2,410/低2,345、收在當日最高、量31,606張TWSE口徑、成交額754.30億、成交88,066筆、TWSE確認)——漲價題材引爆反彈第二日、站回5MA/60MA
change_pct = "+3.88%（7/21收）"    # 7/21官方漲跌幅+3.88%（TWSE、日經漲價報導催化、外資終結連8賣轉買）
nyse_price = "US$424.61（7/21收）" # NYSE TSM 7/21 Tue收$424.61(+$22.31、+5.55% vs 7/20 $402.30、Yahoo API確認)——漲價報導激勵；SOX同日+5.21%收12,356.16(自6/22史高收斂至-15.6%、脫離熊市線)
volume = "31,606（7/21收）"    # 台股7/21官方成交量31,606張（TWSE口徑、僅7/20的57%、量縮價漲惜售式反彈）、成交額754.30億
news_summary = "報告日2026-07-22(Wed)。🚀漲價題材引爆反彈第二日：台股2330 7/21官方收NT$2,410(+90、+3.88%、開2,350/高2,410/低2,345、收在當日最高、量31,606張、額754.30億、TWSE確認)站回5MA 2,386/60MA 2,334——大盤同日飆+1,783.17點(+4.20%)收44,232.87創史上最大單日漲點、收復7/17崩跌六成(台積電領軍、聯發科+9.88%、日月光+6.03%)。💰頭條：日經亞洲報導TSMC已通知客戶2027年初起先進+成熟製程全面漲價5-10%、HPC超額追加訂單另加價10-15%(Apple/NVIDIA/AMD/Qualcomm均適用)——定價權最強證據。📊籌碼最強訊號兌現：外資7/21對2330終結連8賣、正式轉買+4,976張(T86官方)、投信+974(連4買)、自營+214(連9買)、三大法人+6,163張連2日買超；全市場三大法人買超162.9億(BFI82U：投信+171.1億、外資仍小賣-43.3億)。✅NYSE TSM 7/21大漲+5.55%收$424.61、SOX +5.21%收12,356.16脫離熊市線(自史高-15.6%)；SOXX/SOXL單日吸金逾$21億；MU+12.2%、INTC+8.6%、AMD+8.1%、設備股全面大漲；韓股SK+4.08%/三星+6.15%、7/22早盤續強。📊ADR溢價擴大至+13.9%(匯率32.33)、台股7/22具補漲題材。⚠️保留：量縮反彈(僅前日57%)、續攻需量能放大。技術面(官方收盤計算)：RSI 51.5站回中軸、MACD綠柱收斂-15.2、KDJ K40.1/D40.2金叉臨界(J 5.0→39.8)；支撐2,386(5MA)/2,345/2,334(60MA)/2,320、壓力2,419-2,424(20MA)/2,470/2,528-2,535——站穩20MA為反彈第二道確認。"
change_color = "00B050"        # 綠色（7/21官方+3.88%上漲）

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

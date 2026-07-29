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

today_str = "2026-07-29"
tw_price = "NT$2,280（7/28收）"   # 台股2330 7/28 Tue官方收NT$2,280(-70、-2.98% vs 7/27 NT$2,350、開2,270/高2,305/低2,270、量45,333張TWSE口徑、成交額1,036.24億、成交463,044筆、TWSE確認)——韓股熔斷日跳空低開-80、相對抗跌惟60MA 2,347季線與布林下軌2,289雙雙失守
change_pct = "-2.98%（7/28收）"    # 7/28官方漲跌幅-2.98%（TWSE、跌幅小於大盤-4.65%與供應鏈-9~-10%、相對抗跌）
nyse_price = "US$392.31（7/28收）" # NYSE TSM 7/28 Tue收$392.31(-$6.78、-1.70% vs 7/27 $399.09、Yahoo API確認)——亞洲股災外溢下相對抗跌；SOX同日-4.49%收11,035.68自史高-24.6%熊市加深、惟NVDA +0.25%/AAPL +0.94%翻紅、VIX反降18.21
volume = "45,333（7/28收）"    # 台股7/28官方成交量45,333張（TWSE口徑、量放大57%；成交筆數自92,753暴增至463,044、恐慌小單再現）、成交額1,036.24億
news_summary = "報告日2026-07-29(Wed)。⚠️韓股熔斷引爆亞股股災：KOSPI 7/28暴跌-10.8%（近20年最大跌幅、今年第8次熔斷）——三星-13.39%、SK海力士-14.65%（盤中一度-30%）——導火線為AI基建融資疑慮＋中國記憶體產能競爭＋雙倍槓桿ETF斷頭。⚠️台股7/28史上第3大跌點：大盤-2,030.83點(-4.65%)收41,603.36、電子-5.11%；2330官方收NT$2,280(-70、-2.98%、開2,270/高2,305/低2,270、量45,333張、額1,036.24億、463,044筆、TWSE確認)——跌幅小於大盤與供應鏈（聯發科-9.92%、日月光-8.88%、聯電-9.92%、華邦電跌停）、相對抗跌，惟60MA 2,347季線、7/17低2,290、布林下軌2,289全數失守。⚠️籌碼再轉空：外資賣超擴大-14,659張(連5賣、T86官方)、投信+284連3買、自營+5,123連4買大幅低接；全市場BFI82U三大法人大賣-1,176.0億（外資-874.8億）。📊美股7/28分化：TSM -1.70%收$392.31相對抗跌、SOX -4.49%收11,035.68自史高-24.6%——記憶體/設備重挫（MU -8.85%、AMAT -7.82%、AMKR -24.74%）、AMD -8.15%；惟NVDA +0.25%/AAPL +0.94%翻紅、VIX反降-2.46%至18.21——未見美股系統性恐慌。⭐今晨(7/29)韓股初步止穩：三星盤中~22.8萬(+3.6%)、SK持平；SK海力士今日Q2財報（估營益~64兆韓元創史高）電話會HBM定調為風向球。📊ADR溢價回擴至+11.2%（匯率32.30）。技術面(官方收盤計算)：RSI 40.8、MACD綠柱急擴-33.9且DIF -9.3深陷零軸下（中期空方確立）、KDJ死叉深化且J 9.5超賣（K29.1/D38.9）；支撐2,270/2,230/2,200、壓力2,289-2,290/2,305/2,320/2,347-2,357。⭐觀察：韓股止穩延續性、守2,270／收復布林下軌2,289、外資賣超能否收斂、今晚美股MSFT/META財報。"
change_color = "FF0000"        # 紅色（7/28官方-2.98%下跌）

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

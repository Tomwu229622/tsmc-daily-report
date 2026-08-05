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

today_str = "2026-08-05"
tw_price = "NT$2,320（8/4收）"   # 台股2330 8/4 Tue官方收NT$2,320(-50、-2.11% vs 8/3 NT$2,370、開2,335/高2,360/低2,310、量41,021張TWSE口徑、成交額954.55億、成交277,787筆、TWSE確認)——連2日獲利了結、跌破20MA 2,371與60MA 2,350；大盤同日僅-0.06%收43,360.66
change_pct = "-2.11%（8/4收）"    # 8/4官方-2.11%（TWSE、大盤同日-25.75點/-0.06%收43,360.66——中小型科技（矽光子/CPO/低軌衛星）漲停潮接棒、上漲578家>下跌406家）
nyse_price = "US$417.17（8/4收）" # NYSE TSM 8/4 Tue收$417.17(+$11.06、+2.72% vs 8/3 $406.11、Yahoo API確認、連4日收高創本輪新高)——美股8/4半導體史詩級大漲：SOX+6.55%收12,179.26、ARM+17.36%、INTC+10.84%、AMKR+9.89%、LRCX+7.85%、MU+7.62%、AVGO+6.61%；S&P+1.79%收7,736.52創收盤新高；VIX 16.50
volume = "41,021（8/4收）"    # 台股8/4官方成交量41,021張（TWSE口徑、仍低於5日均量53,044——連2日拉回量未放大、賣壓有限）、成交額954.55億
news_summary = "報告日2026-08-05(Wed)。📊台美短線背離、今日補漲條件齊備：台股8/4連2日獲利了結——2330收NT$2,320(-50、-2.11%、開2,335/高2,360/低2,310、量41,021張低於5日均量53,044、額954.55億、277,787筆、TWSE確認)、跌破20MA 2,371與60MA 2,350（季線失守）；大盤僅-25.75點(-0.06%)收43,360.66、上漲578家>下跌406家——矽光子/CPO/低軌衛星漲停潮、中小型接棒；華邦電+9.79%續強、日月光-4.10%、聯發科-1.15%拉回換手。⚠️籌碼：外資連2賣-12,430張（自-9,717擴大、「一日翻多」確認失敗）、投信+434連8買、自營-2,327、三大合計-14,323張（T86官方）；全市場BFI82U合計+20.1億（外資-57.3億、投信+271.7億大買、自營-194.3億）——外資調節集中權值、內資持續進場。⭐美股8/4半導體史詩級大漲：SOX+6.55%收12,179.26（自史高收斂至-16.8%）——ARM+17.36%收$280.56、INTC+10.84%站上$100、AMKR+9.89%、LRCX+7.85%、MU+7.62%、KLAC+6.95%、AVGO+6.61%、AMAT+5.48%、ASML+4.22%、NVDA+2.56%收$211.94；S&P 500+1.79%收7,736.52創收盤新高、那指+2.6%、Palantir+29%；TSM+2.72%收$417.17連4日收高創本輪新高；AMZN-2.32%、META-0.39%——資金自hyperscaler轉入半導體；VIX 16.50。📊AMD Q2財報（8/4盤後）全面超標：營收$11.53B(+50% YoY、超共識~$11.3B)創新高、EPS $1.66超$1.62、資料中心$6.7B(+107%、占58%)、Q3指引$13B±0.3B遠超共識$12.52B——惟盤後重挫約-7~-9%（媒體報導、高「耳語數字」回吐）——需求數字對TSMC先進製程/CoWoS實質正面。⭐韓股V型再起：8/4尾盤收復（盤中-3%→三星收+0.21% 24.0萬、SK+0.64% 157.7萬）、今晨8/5盤中暴漲——SK+7.67%至169.8萬、三星+4.79%、KOSPI 6,670.67（較8/3+6.6%；8/4官方指數Yahoo缺漏）。📊BWIBBU官方P/E 31.19x/P/B 10.21x/殖利率0.95%；ADR溢價暴衝至+16.37%（$417.17 vs理論值$358.50、匯率32.36）——近期最大、今日補漲壓力極強。技術面依官方收盤重算（至8/4）：RSI 47.3、MACD綠柱-12.4續收斂（DIF-17.8零軸下）、KDJ金叉延續（K53.9/D44.2/J73.3）、5MA 2,304/120MA 2,152仍站上；支撐2,310/2,304/2,215、壓力2,345/2,350/2,360/2,371/2,425/2,535。整體信號：中性偏多。⭐觀察：跳空收復2,350(60MA)/2,371(20MA)與否、外資連2賣能否翻買、AMD盤後跌勢對亞股AI情緒干擾（韓股今晨已大漲、暫未兌現）、232條款（仍未查得公布）與台美關稅談判（傳20%→~15%）。全程TWSE/Yahoo官方API取數。"
change_color = "FF0000"        # 紅色（8/4官方-2.11%下跌）

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

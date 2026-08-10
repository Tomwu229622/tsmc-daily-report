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

today_str = "2026-08-10"
tw_price = "NT$2,370（8/7收）"   # 台股2330 8/7 Fri官方收NT$2,370(+5、+0.21% vs 8/6 NT$2,365、開2,390/高2,395/低2,355、量24,414張TWSE口徑、成交額579.47億、成交64,670筆、TWSE確認)——小漲站回5MA 2,366、重新站上全部均線；大盤同日-0.38%收44,225.91、2330相對強勢
change_pct = "+0.21%（8/7收）"    # 8/7官方+0.21%（TWSE、大盤同日-170.79點/-0.38%收44,225.91——2330相對強勢；量24,414張僅5日均量32,593的75%、近月新低、連3日萎縮）
nyse_price = "US$420.04（8/7收）" # NYSE TSM 8/7 Fri收$420.04(+$1.84、+0.44% vs 8/6 $418.20、Yahoo API確認)——連2日創本波收盤新高、重站$420；SOX+2.56%收12,356.79創本波新高(連3日站上12,000)；NVDA+2.27%收$223.96創史高、QCOM+4.66%、KLAC+2.53%、AMAT+2.21%、ASML+2.15%、INTC+1.84%收復$100；AMD-1.21%、ARM-1.43%、MU-0.44%；S&P+0.62%收7,757.64創收盤新高；VIX 14.90本波最低
volume = "24,414（8/7收）"    # 台股8/7官方成交量24,414張（TWSE口徑、僅5日均量32,593張的75%、近月新低、連3日萎縮——多空觀望等待7月營收與美股方向）、成交額579.47億
news_summary = '報告日2026-08-10(Mon)。⭐外部條件全面轉多、補漲條件齊備：台股8/7小漲站回全均線——2330開2,390高開後回測2,355(貼近60MA 2,357獲支撐)、尾盤收NT$2,370(+5、+0.21%、高2,395/低2,355、量24,414張僅5日均量32,593的75%近月新低、額579.47億、64,670筆、TWSE確認)——站回5MA 2,366、重新站上全部均線(20MA 2,362/60MA 2,357/120MA 2,167)；大盤-170.79點(-0.38%)收44,225.91、2330相對強勢。✅籌碼翻多、三大法人同步站買方：外資回買+642張(終結1日賣超)、投信+61張(終結連2賣)、自營+1,243、合計+1,946張(T86官方)——幅度溫和需連2買確認；⚠️全市場BFI82U合計-428.9億(外資-407.2億大賣、投信-12.0億)——集中調節中小型與記憶體(華邦電外資-43,331張、-4.39%連5紅中止；日月光-3,054張、-1.68%)、資金回流台積電等權值避風。⭐美股8/7(五)半導體全面大漲：SOX+2.56%收12,356.79創本波新高(連3日站上12,000、自史高收斂至-15.6%)；NVDA+2.27%收$223.96創歷史新高；QCOM+4.66%、KLAC+2.53%、AMAT+2.21%、ASML+2.15%創本波新高、AMKR+2.05%、INTC+1.84%收復$100——設備股全面反攻；TSM+0.44%收$420.04連2日創本波收盤新高；AMD-1.21%、ARM-1.43%、MU-0.44%少數逆勢；S&P+0.62%收7,757.64創收盤新高、那指+1.30%、VIX降至14.90(本波最低)；台指期夜盤大漲約736點(媒體報導)。✅韓股記憶體風暴緩和：SK 8/7 -4.88%收142.2萬韓元後、今晨8/10反彈+2.46%至145.7萬、三星+1.19%、KOSPI+1.00%(Yahoo 09:37 KST快照)——二次崩跌未發生；⚠️惟台股記憶體補跌已現(華邦電-4.39%)。📅7月合併月營收截至今晨未公布(pr.tsmc.com查證)、預計今日(8/10)前後公布——投顧預估上看4,500億(預期值、非官方)；6月4,426.8億(+67.9%YoY)基期極高。📊「台積電條款」今日上路：千金股注意標準價差門檻100→300元、處置期10→5日——高價權值股流動性利多。📊BWIBBU官方P/E 31.86x/P/B 10.43x/殖利率0.93%；市值約NT$61.4兆；ADR溢價自+13.98%續擴至+14.28%($420.04 vs理論值$367.56、匯率32.24)——連3日擴大、歷史上多以台股補漲收斂。技術面依官方收盤重算(至8/7)：RSI 50.6中軸上方、MACD黃金交叉續強(DIF-7.1/DEA-10.3、柱+6.3較前日+3.7續放大、連3日紅柱擴張)、KDJ金叉延續(K72.2/D62.5/J91.7偏高)；均線5MA 2,366/20MA 2,362/60MA 2,357/120MA 2,167全數站上；支撐2,366/2,362/2,357/2,355/2,320/2,217、壓力2,390/2,395/2,415/2,425/2,507/2,535。整體信號：中性偏多。⭐觀察：(1)補漲幅度與量能——ADR溢價+14.28%＋SOX創本波新高＋夜盤+736點、開高機率高、關鍵在量能能否自24,414張明顯回升並站穩2,395-2,415壓力區；(2)7月營收公布時點與數字；(3)外資回買+642張延續性(連2買才確認)；(4)232條款仍未公布(未能證實)。全程TWSE/Yahoo官方API取數。'
change_color = "00B050"        # 綠色（8/7官方+0.21%上漲）

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

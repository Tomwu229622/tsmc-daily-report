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

today_str = "2026-08-07"
tw_price = "NT$2,365（8/6收）"   # 台股2330 8/6 Thu官方收NT$2,365(-40、-1.66% vs 8/5 NT$2,405、開2,395/高2,395/低2,360、量25,537張TWSE口徑、成交額606.01億、成交89,846筆、TWSE確認)——縮量拉回、回吐8/5大漲的一半，惟仍站20MA 2,364與60MA 2,354；大盤同日-0.48%收44,396.70
change_pct = "-1.66%（8/6收）"    # 8/6官方-1.66%（TWSE、大盤同日-214.90點/-0.48%收44,396.70——縮量整理；量25,537張僅5日均量41,605的61%、近月最低）
nyse_price = "US$418.20（8/6收）" # NYSE TSM 8/6 Thu收$418.20(+$4.20、+1.01% vs 8/5 $414.00、Yahoo API與StockAnalysis雙源確認)——創本波收盤新高、與台股同日-1.66%形成背離；SOX+0.33%收12,048.69連2日站穩12,000；ARM+4.41%、QCOM+1.82%、ASML+1.56%、AMD+1.50%收$489.28；MU-1.31%、INTC-1.24%；NVDA-0.10%收$218.99；S&P-0.18%；VIX 15.15本波最低
volume = "25,537（8/6收）"    # 台股8/6官方成交量25,537張（TWSE口徑、僅5日均量41,605張的61%、近月最低——縮量整理而非爆量出貨）、成交額606.01億
news_summary = '報告日2026-08-07(Fri)。📊台美背離重現、記憶體風暴未及代工：台股8/6縮量拉回——2330開2,395後全日走低、收NT$2,365(-40、-1.66%、高2,395/低2,360、量25,537張僅5日均量41,605的61%近月最低、額606.01億、89,846筆、TWSE確認)——回吐8/5大漲的一半、跌破5MA 2,377，惟仍站20MA 2,364(貼線+1元)與60MA 2,354(季線未再失守)；大盤-214.90點(-0.48%)收44,396.70；族群分歧：華邦電+1.18%連5紅、日月光+0.34%，聯發科-2.00%自4,000回落。⚠️籌碼翻空、外資翻買僅1日：2330外資轉賣-4,011張(8/5的+12,251張未延續、驗證「ADR溢價套利一次性回補」疑慮)、投信-82張連2賣、自營-583、三大合計-4,676張(T86官方)；⭐惟全市場BFI82U仍+51.2億(投信+94.1億續買、外資+20.2億、自營合計-63.1億)——資金輪出台積電而非撤離台股。⭐美股8/6反向走強、TSM創本波收盤新高：SOX+0.33%收12,048.69連2日站穩12,000；TSM+1.01%收$418.20(Yahoo與StockAnalysis雙源確認)——在台股同日-1.66%下逆勢走高、台美背離重現；ARM+4.41%領漲、QCOM+1.82%、ASML+1.56%、AMD+1.50%收$489.28(財報賣壓僅1日、BofA維持買進)；⚠️記憶體與設備續弱：MU-1.31%、INTC-1.24%失守$100、AMAT-1.27%；NVDA-0.10%收$218.99高檔整理；S&P-0.18%、那指-0.06%、道瓊-0.85%、VIX降至15.15(本波最低)。⚠️⭐韓股記憶體8/6崩跌但屬「記憶體利空、非代工利空」：SK海力士-10.37%收149.5萬韓元、三星-6.30%收23.05萬、KOSPI重挫逾4%觸發sidecar熔斷(採Yahoo官方收盤序列；媒體盤中快照報-5.7~-8.3%為不同時點)——導火線為SanDisk 8/5盤後FY27Q1營收指引$103-108億低於共識$108億、加上Western Digital財測平淡，兩者盤後跌逾8-10%(TradingKey/KED Global)；⭐台積電無HBM/NAND曝險，當晚SOX反收紅、TSM創新高即為市場分流定價；高盛/摩通維持韓廠買進(遠期本益比僅3.5-3.6倍)；今晨8/7韓股分歧(三星+2.82%、SK-1.34%)。📊產業：2nm報價漲幅確認10-20%(TrendForce、遠低於傳言50%、3-7nm漲3-10%、未經台積電官方證實)；熊本JASM廠震後全面復產(TechNews 8/4)；232條款8/7晨仍未查得公布(未能證實)；7月月營收預計8/10前後公布(6月4,426.8億、+67.9%YoY史上新高、基期極高、公布時點與金額未能證實)。📊BWIBBU官方P/E 31.80x/P/B 10.41x/殖利率0.93%；市值約NT$61.3兆；ADR溢價自+10.97%再度擴大至+13.98%($418.20 vs理論值$366.89、匯率32.23)——台股跌/ADR漲雙向背離重現、歷史上此型態多以台股補漲收斂。技術面依官方收盤重算(至8/6)：⭐出現「價跌指標未壞」溫和背離——MACD黃金交叉續強(DIF-9.2/DEA-11.1、柱+3.7較前日+0.5明顯放大、連2日紅柱擴張)、RSI 50.3貼中軸、KDJ金叉延續(K69.5/D57.6/J93.4仍偏高)；均線5MA 2,377(跌破)/20MA 2,364/60MA 2,354/120MA 2,162(仍站上)；支撐2,364/2,360/2,354/2,320/2,217、壓力2,377/2,395/2,415/2,425/2,511/2,535。整體信號：中性。⭐觀察：(1)ADR溢價+13.98%＋TSM創本波新高＋SOX站穩12,000、補漲條件再度具備；(2)外資是否連2賣；(3)量能能否自25,537張回升(守2,364且量增為多方、放量破2,354季線則轉弱)；(4)韓股記憶體利空是否外溢。全程TWSE/Yahoo官方API取數。'
change_color = "FF0000"        # 紅色（8/6官方-1.66%下跌）

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

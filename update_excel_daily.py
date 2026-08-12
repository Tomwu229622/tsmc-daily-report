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

today_str = "2026-08-12"
tw_price = "NT$2,395（8/11收）"   # 台股2330 8/11 Tue官方收NT$2,395(+15、+0.63% vs 8/10 NT$2,380、開2,390/高2,405/低2,375、量18,248張TWSE口徑、成交額436.67億、成交50,065筆、TWSE確認)——連2紅收復2,390逼近2,400；大盤同日+0.43%收45,120.72續創8月反彈新高、2330相對強勢
change_pct = "+0.63%（8/11收）"    # 8/11官方+0.63%（TWSE、大盤同日+191.96點/+0.43%收45,120.72——2330相對強勢；量18,248張僅5日均量25,296的72%、近月新低、連5日萎縮；8/10 +0.42%收2,380一併補記——8/11排程未執行）
nyse_price = "US$422.06（8/11收）" # NYSE TSM 8/11 Tue收$422.06(+$3.59、+0.86% vs 8/10 $418.47、Yahoo API確認)——創本波收盤新高且逆大盤(S&P同日-0.32%收7,728.20)；SOX 8/10 -2.94%回檔後8/11+0.87%收12,098.47守穩12,000；設備股領漲：ASML+3.80%收$1,799.38創本波新高、KLAC+4.01%首站$200、LRCX+1.64%；NVDA平盤$217.50(FT報導攜華爾街大行籌$5,000億AI基建資金)；AAPL-1.09%、AVGO-1.50%；VIX 15.28
volume = "18,248（8/11收）"    # 台股8/11官方成交量18,248張（TWSE口徑、僅5日均量25,296張的72%、再創近月新低、連5日萎縮：36,782→25,537→24,414→21,498→18,248——無量緩漲為最大隱憂）、成交額436.67億
news_summary = '報告日2026-08-12(Wed)。⚠️8/11排程未執行、本次一併補記8/10行情。⭐基本面雙重利多落地、量縮緩攻逼近2,400：(1)7月營收4,675.81億(+44.69%YoY、+5.62%MoM、8/10盤後公布)連3月創單月新高、優於投顧預期4,500億；1-7月累計2.87兆(+37.0%YoY)、Q3指引達成無虞；(2)8/11董事會四大議案：Q2配息NT$7.0/股(12/10除息、12/16基準、2027-01-07發放、⭐首推外資美元領息選項)、核准資本預算約US$294.4億、與Sony半導體合資7,470億日圓於熊本設新公司開發次世代影像感測器(台積電出資2,820億日圓約NT$572億、2029量產)、通過Q2財報(營收1.270兆、純益7,065.6億、H1 EPS 49.33)。📊台股連2紅：8/10 +0.42%收2,380(大盤+1.59%大漲702.85點收44,928.76——2330補漲落後、資金衝記憶體/封裝：華邦電+9.79%、日月光+7.69%)；8/11 +0.63%收2,395(開2,390/高2,405/低2,375、TWSE確認)收復2,390逼近2,400；大盤+0.43%收45,120.72續創8月反彈新高、2330轉相對強勢。⚠️量能連5日萎縮：8/11僅18,248張(5日均量25,296的72%、近月新低：36,782→25,537→24,414→21,498→18,248)、額436.67億——無量緩漲、突破2,400-2,410壓力區需量能確認。✅籌碼續翻多：2330外資連3買且加碼(+642→+119→+3,498張)、投信連3買(+61→+153→+382)、自營-189、8/11三大合計+3,691張(T86官方)；全市場BFI82U 8/10大買+736.1億(外資+517.4億)、8/11續買+280.1億(外資+223.3億)——資金全面回流。📊美股8/10回檔8/11反彈：SOX 8/10 -2.94%收11,993.86(油價通膨疑慮＋AI資本支出雜訊、NVDA-2.86%、INTC-4.06%破$100)→8/11 +0.87%收12,098.47(FT報導NVIDIA攜華爾街大行籌$5,000億AI基建資金；設備股領漲ASML+3.80%、KLAC+4.01%首站$200)；TSM 8/11 +0.86%收$422.06創本波收盤新高且逆大盤(S&P-0.32%收7,728.20、VIX 15.28)。✅韓股今晨(8/12)大漲：三星+3.55%至24.8萬韓元、SK+2.32%至145.8萬、KOSPI+2.06%(盤中快照)——記憶體行情重啟、華邦電外資2日回補+56,916張。📊BWIBBU官方P/E 32.20x/P/B 10.54x/殖利率0.92%；市值約NT$62.1兆；ADR溢價自+14.28%收斂至+13.56%($422.06 vs理論值$371.66、匯率32.22)——以台股補漲收斂、符合歷史型態。技術面依官方收盤重算(至8/11)：RSI 52.5續升、MACD黃金交叉續強(DIF-1.7逼近零軸/DEA-8.2、柱+12.9連5日擴張：+0.5→+3.7→+6.3→+9.6→+12.9)、KDJ金叉延續(K79.3/D70.9/J96.1偏高)；均線5MA 2,383/10MA 2,344/20MA 2,358/60MA 2,364/120MA 2,183全數站上且5MA上彎；支撐2,390/2,383/2,375/2,364/2,358/2,320、壓力2,400/2,405/2,410/2,425/2,497/2,535。整體信號：中性偏多。⭐觀察：(1)2,400-2,410攻防與量能能否終結連5縮；(2)📅今晚(美東8/12)美國7月CPI——短線最大宏觀變數(8/10美股回檔即因通膨疑慮)；(3)外資連4買將確立中期翻多；(4)232條款半導體專章仍未公布(未能證實)。全程TWSE/Yahoo官方API取數。'
change_color = "00B050"        # green: 8/11 official +0.63% up

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

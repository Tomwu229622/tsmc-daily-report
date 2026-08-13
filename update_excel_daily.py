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

today_str = "2026-08-13"
tw_price = "NT$2,415（8/12收）"   # 台股2330 8/12 Wed官方收NT$2,415(+20、+0.84% vs 8/11 NT$2,395、開2,405/高2,415/低2,390、量19,448張TWSE口徑、成交額467.98億、成交63,548筆、TWSE確認)——⭐突破2,400整數與2,405/2,410壓力帶且收在當日最高、創7/31以來收盤新高；大盤同日+0.88%收45,518.07續創8月反彈新高
change_pct = "+0.84%（8/12收）"    # 8/12官方+0.84%（TWSE、大盤同日+397.35點/+0.88%收45,518.07；量19,448張為5日均量21,829的89%、終結連5日萎縮(+6.6% vs 8/11)惟仍未放大；連3日收高、站上全部均線）
nyse_price = "US$429.15（8/12收）" # NYSE TSM 8/12 Wed收$429.15(+$7.09、+1.68% vs 8/11 $422.06、Yahoo API確認)——連3日創本波收盤新高；美國7月CPI年增3.4%/核心2.5%符合預期且低於前值、CME FedWatch 9月升息機率降至42%，SOX+2.49%收12,399.38創本波新高；NVDA+3.03%收$224.09創收盤新高；設備與記憶體全面領漲：MU+4.92%、LRCX+4.72%、AMAT+4.29%、KLAC+3.88%、INTC+3.32%收復$100；ASML+0.59%收$1,810.07再創史高；AAPL-0.87%、AVGO-0.01%背離；S&P+0.26%、VIX 14.55
volume = "19,448（8/12收）"    # 台股8/12官方成交量19,448張（TWSE口徑、為5日均量21,829張的89%、較8/11的18,248張回升+6.6%終結連5日萎縮，惟量能仍未明顯放大）、成交額467.98億、63,548筆
news_summary = '報告日2026-08-13(Thu)。⭐⭐突破2,400收最高、MACD DIF上穿零軸、今晨夜盤大漲：(1)⭐台股攻堅收在最高——2330 8/12收2,415(+20、+0.84%、開2,405/高2,415/低2,390、TWSE確認)一舉突破2,400整數與2,405(8/11高)/2,410(8/10高)三道壓力，收盤價與當日最高同價為強勢作收、創7/31(收2,425)以來收盤新高；連3日收高、站上全部均線；大盤+397.35點(+0.88%)收45,518.07續創8月反彈新高。(2)⭐⭐技術面關鍵確認成立——MACD DIF自8/11的-1.36上穿零軸至+2.78(前一日日報即標記「DIF上穿零軸將為中期修復的關鍵確認訊號」、本日正式成立)、DEA-5.51、紅柱+16.57連6日擴張(+0.53→+3.68→+6.31→+9.09→+12.43→+16.57)；RSI 54.0中軸上方續升；⚠️惟KDJ J值99.8逼近100(K83.3/D75.1)、短線過熱。(3)⭐美國7月CPI符合預期、緊縮預期降溫：CPI月增0.1%/年增3.4%(前值3.5%)、核心月增0.2%/年增2.5%(前值2.6%)均符合預期且低於前值(BLS官方、CNBC 8/12)；居住項貢獻約2/3漲幅、能源月減1.5%(年增仍14.7%)；CME FedWatch 9月升息機率自45%降至42%——SOX+2.49%收12,399.38創本波新高、NVDA+3.03%收$224.09創收盤新高、TSM+1.68%收$429.15連3日創本波新高、VIX降至14.55；AI概念股財報全面告捷(CoreWeave+20%、Super Micro+15%、Nebius+12.5%)。(4)⭐⭐今晨夜盤大漲、開高機率極高：台指期(TX 202608)夜盤收46,117(+601點、+1.32%、量35,208口、TAIFEX官方)對現貨45,518.07正價差+599；台積電個股期貨(CDF 202608)夜盤收2,439(+26、+1.08%、高2,449)對現貨2,415正價差+24。(5)✅籌碼連4買：2330外資+2,627張(連4買：+642→+119→+3,498→+2,627)、投信+587張(連4買且逐日加碼：+61→+153→+382→+587)、自營-290(自行-235/避險-55)、三大合計+2,924張(T86官方)；全市場BFI82U連3日買超(8/10 +736.1億、8/11 +280.1億、8/12 +220.0億、外資+110.2億)。(6)⭐韓股記憶體全面爆發：8/12三星+6.68%收25.55萬韓元、SK海力士+5.54%收150.4萬；今晨8/13續飆——三星+3.72%至26.5萬、SK+7.11%至161.1萬、KOSPI 6,826.14(盤中快照)；⚠️惟台股記憶體/封裝8/12反向回檔(華邦電-0.56%收177.0、日月光-1.27%收621、聯發科-0.12%收4,015)——資金自中小型輪回權值台積電、與8/10方向相反。📊BWIBBU官方P/E 32.47x/P/B 10.63x/殖利率0.91%；市值約NT$62.6兆；ADR溢價自+13.56%擴大至+14.48%($429.15 vs理論值$374.88、匯率32.21)——ADR漲幅大於台股、台股續有補漲空間。技術面依TWSE官方收盤序列(2026-01-02~08-12共146個交易日)重算：RSI 54.0、MACD DIF+2.78上穿零軸/DEA-5.51/柱+16.57連6日擴張、KDJ K83.3/D75.1/J99.8(逼近超買)；均線5MA 2,385/10MA 2,365/20MA 2,356/60MA 2,364/120MA 2,183全數站上；布林2,493/2,356/2,220；支撐2,405/2,400/2,390/2,385/2,365/2,364、壓力2,425/2,440/2,449/2,493/2,535。⚠️衍生品註記：8/12為週三週選結算日、TXO未平倉大幅結清重建(買權83,111→57,021口、賣權102,705→60,479口)，P/C比自1.24降至1.06屬結算技術性因素而非情緒逆轉，勿與8/11直接比較。整體信號：中性偏多。⭐觀察：(1)夜盤正價差+24指向開高——能否放量站穩2,415並挑戰2,425(7/31收)為第一檢核點；(2)⚠️量能能否突破21,829張(5日均量)——8/12僅89%均量、無量衝高提防沖高回落；(3)KDJ J 99.8過熱的消化方式；(4)外資連5買延續性；(5)232條款半導體專章仍未公布(未能證實)。全程TWSE/TAIFEX/Yahoo官方API取數。'
change_color = "00B050"        # green: 8/12 official +0.84% up

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

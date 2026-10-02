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

today_str = '2026-10-02'
tw_price = 'NT$2,510（10/1收，+30；開2,490低2,485收最高2,510，平6/22年高收盤，量縮39.7%）'
change_pct = '+1.21%（10/1收，+30.00，領先大盤0.35pp；加權+0.86%創收盤新高）'
nyse_price = 'US$459.20（10/1收，+0.66%；費半+1.59%；ADR溢價+16.50%）'
volume = '20,666（10/1收，-39.7%，低於5日均量23,843張13.3%、低於9月中位數3.1%）'
news_summary = """報告日2026-10-02(Fri)。台股資料為2026-10-01(Thu，第四季首個交易日)TWSE官方收盤；ADR、費半與美股個股為10/1正式收盤(Yahoo meta 16:00 ET)；韓股為10/1收盤(Yahoo日K，確有交易；陸媒「國軍節休市」說法判為錯誤)；期貨為TAIFEX 10/1日盤＋今晨夜盤(10/1 15:00至10/2約05:00)；匯率採台北外匯10/1收盤31.842；全市場法人金額採BFI82U官方。

(1)⭐⭐10/1台股2330開高走高收全日最高2,510(+30、+1.21%)，平6/22的2026年最高收盤、首度收在2,495–2,510壓力區上緣(開2,490／低2,485，缺口2,480–2,485未回補)；加權+413.36點(+0.86%)收48,353.49創收盤新高，台積電領先0.35pp。⚠️量20,666張萎縮39.7%(金額NT$517.02億、57,634筆)，低於5日均量23,843張13.3%、低於9月中位數3.1%，為縮量創高。

(2)⭐前報六項檢核：守2,475通過(9/30黑K頭部警訊解除)；回撤條件(MA10 2,452／外資賣逾5,000張)未觸發；上調條件「收≥2,510且量≥24,112張」只過價格、量能不過→依判準不上調；夜盤2,505指示開+10收+30兩端兌現(累計開6/7收4/7)；美光10/1正式收盤+3.03%收$1,097.39、僅兌現盤後約+12%的四分之一(利多出盡部分應驗)；外資連2買且集中度自5.6%回升至35.4%；P/C新週期第2日82.54%。

(3)⭐籌碼：外資+3,090張(買14,212,490股／賣11,121,760股，連2買、2日合計3,792張，占全市場外資買超219.37億的35.4%)、投信+596(連3買)、自營+602(連12買)，三大法人官方合計+4,289張；全市場三大法人+NT$294.62億(外資+219.37／投信+58.65／自營自行+32.37／避險-15.78，BFI82U官方)。⚠️外資台指期淨空-79,554口(再擴大1,403口，空單91,927口為追蹤以來最高)，現貨買、期貨空背離重現。

(4)⭐技術(TWSE官方序列180日自算)：RSI 61.37(超越9/23的60.97、7/01以來最高)、MACD柱+11.75(中止連3日縮小；DIF 23.91連9升／DEA 18.04)、K 82.68／D 72.49／J 103.05(重返100之上，9/23以來首度)；MA5 2,488／MA10 2,465／MA20 2,443／MA60 2,403／MA120 2,338五線多頭連6日；布林2,359–2,443–2,526、寬度6.81%。衍生品：個股期日盤收2,530(+30)、OI 25,362、基差+20，今晨夜盤收2,515(較結算-15、較現貨+5)；台指期48,685(+0.73%)、夜盤48,475(-210)；CDO近月call 317(單日+127、集中2,650–2,700)／put 110口、最大痛點2,500。

(5)⭐美股10/1費半+200.38點(+1.59%)收12,829.00、ADR +0.66%收$459.20(溢價+16.50%、自17.17%收斂)；19檔供應鏈14漲5跌：科林+3.53%、應材+3.50%、科磊+2.77%、美光+3.03%、SK海力士+3.21%、三星+2.79%、輝達+1.09%；博通-2.15%、高通-1.06%(連4黑)、蘋果-0.81%、艾司摩爾-0.18%、英特爾-0.19%；台股聯發科+1.22%、日月光+1.00%、華邦電+0.28%。消息：美10年期殖利率盤中5.34%創2002年以來新高、收5.24%【媒體數】；9月ISM製造業PMI 54.5；美光26份長約、承諾US$320億【媒體數】；台積電德州第二園區評估延燒、董事會未拍板。BWIBBU P/E 29.09x(追蹤以來最高)／P/B 10.12x／殖利率0.88%。

(6)立場維持「中性偏多」、不上調，首要保留改為「縮量創高」與「現貨買、期貨空」背離。次日(10/2週五)檢核：收≥2,515且量≥23,843張→上調；跌回2,495之下且量增、或跌破MA10 2,465、或外資單日賣逾5,000張→退回中性；J過100後回落風險；夜盤「持平至小幅偏弱」指示兌現；外資能否連3買、期貨空單是否續擴大。9月營收約10/10前後(官方未公告；10/9國慶補假休市【待核對】)、緘默期10/5起、Q3法說10/15。"""
change_color = '00B050'
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

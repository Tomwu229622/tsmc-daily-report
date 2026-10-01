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

today_str = '2026-10-01'
tw_price = 'NT$2,480（9/30收，+5；開2,495高2,510收最低2,480，放量黑K）'
change_pct = '+0.20%（9/30收，+5.00，落後大盤0.45pp；加權+0.65%）'
nyse_price = 'US$456.19（9/30收，-0.16%；費半-0.00%；ADR溢價+17.17%）'
volume = '34,282（9/30收，+27.5%，高於9月中位數21,326張、5日均量1.42倍）'
news_summary = """報告日2026-10-01(Thu)。台股資料為2026-09-30(Wed，第三季最後交易日)TWSE官方收盤；ADR、費半與美股個股為9/30正式收盤(Yahoo meta 16:00 ET)；韓股為9/30收盤；期貨為TAIFEX 9/30日盤＋今晨夜盤(9/30 15:00至10/1約05:00)；匯率採台北外匯9/30收盤31.852；全市場法人金額採BFI82U官方。

(1)⚠️9/30台股2330放量34,282張(+27.5%、金額NT$853.89億、83,797筆)開高走低：開2,495、盤中2,510(＝6/22全年最高收盤價位)後收全日最低2,480(+5、+0.20%)，上漲黑K；加權+308.17點(+0.65%)收47,940.13，台積電落後0.45pp；2,495–2,510自9/22起三度上攻未站穩，壓力區依前報判準確立。

(2)⭐⭐前報六項檢核：收復2,500半通過(盤中收復、收盤未站上，但站回MA5 2,478)；回撤條件(MA10 2,442／外資第3日賣逾5,000張)未觸發；夜盤2,520指示開+20收+5兩端方向兌現(累計開5/6收3/6)；投信+761連2買、自營+399連11買延續；美光FQ4財報大超預期(營收US$542.3億、資料中心年增逾11倍、指引US$615億、盤後約+12%【媒體數】)；週選結算後台指選擇權P/C 80.57%。

(3)⭐籌碼：外資轉買702張(買25,569,109股／賣24,866,515股，中止連2日賣超8,403張)，三大法人合計+1,864張(9/23以來首度淨買超)；全市場三大法人+NT$403.88億(外資+308.57／投信+82.91／自營+12.40，BFI82U官方)，台積電僅占外資買超5.6%；外資台指期淨空-78,151口(收斂878口、中止連4日擴大)。

(4)⭐技術(TWSE官方序列179日自算)：RSI 57.54、MACD柱+9.34連3日縮小(DIF 21.23連6升／DEA 16.57)、K 74.02／D 67.40／J 87.25(連3日回落)；MA5 2,478／MA10 2,452／MA20 2,439／MA60 2,403／MA120 2,333五線多頭連5日，收盤站回MA5；布林2,362–2,439–2,516；5日均量24,112張。衍生品：個股期日盤收2,500(+9)、OI 25,063、基差+20，今晨夜盤收2,505(+5)；台指期48,330(+1.15%)、夜盤48,298；CDO近月call 190／put 109口、最大痛點2,500(樣本小)。

(5)⭐美股9/30費半-0.54點(-0.00%)收12,628.62、ADR -0.16%收$456.19(溢價+17.17%、自17.72%收斂)；19檔供應鏈10漲9跌：科林+1.44%、英特爾+3.71%、蘋果+1.10%、輝達+0.51%；艾司摩爾-1.24%、科磊-0.81%、Amkor-3.25%、Arm-1.37%(五日-12.90%)；台股日月光+2.33%、華邦電+3.16%、聯發科+0.20%；韓股SK海力士+0.62%、三星-1.47%。消息：美10年期殖利率盤中5.29%創2007年來新高【媒體數】、8月PCE年增3.4%降溫但道瓊-0.86%；AMD 82億美元收購World Labs；新台幣9月貶0.59%、外資9月賣超台股429億【媒體數】。BWIBBU P/E 28.75x／P/B 10.00x／殖利率0.89%。

(6)立場維持「中性偏多」，首要保留改為2,495–2,510壓力區三度受阻與放量黑K(前報頭部警訊字面成立、定為弱化版)。次日檢核：守2,475(破且量增→頭部確認、退回中性)；跌破MA10 2,452或外資單日賣逾5,000張→退回中性；收≥2,510且量≥24,112張→上調；夜盤2,505指示兌現；美光10/1正式收盤核對；外資能否連2買且集中回台積電。9月營收約10/10(官方未公告)、緘默期10/5起、Q3法說10/15。"""
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

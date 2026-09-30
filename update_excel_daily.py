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

today_str = '2026-09-30'
tw_price = 'NT$2,475（9/29收，連假後首日平盤；開＝低＝收2,475、高2,495 T字線）'
change_pct = '0.00%（9/29收，0.00，領先大盤+0.82pp；加權-0.82%）'
nyse_price = 'US$456.94（9/29收，+0.90%；費半+1.32%；ADR溢價+17.72%）'
volume = '26,893（9/29收，+84.7%，高於9月中位數21,025張、5日均量1.31倍）'
news_summary = """報告日2026-09-30(Wed)。台股資料為2026-09-29(Tue，連假後首個交易日)TWSE官方收盤；ADR、費半與美股個股為9/29正式收盤；韓股為9/29收盤；期貨為TAIFEX 9/29日盤，另含今晨夜盤(9/29 15:00至9/30約05:00)。今日9/30為週選結算日、美光美東9/30盤後財報。

(1)⭐9/29台股2330開＝低＝收2,475(高2,495、0.00%)，量26,893張(+84.7%、成交金額NT$666.59億、62,248筆)，加權-392.64點(-0.82%)收47,631.96、聯發科-7.10%、日月光-1.72%，台積電領先大盤0.82pp。外資續賣3,736張(連2日合計8,403張、約NT$92.6億；全市場外資賣超NT$632.01億【媒體數】、台積電占約14.7%)、投信結束連3賣轉買628張、自營連10買117張，三大法人合計-2,991張。P/E 28.69x、P/B 9.98x、殖利率0.89%不變。

(2)⭐⭐前報七項檢核：「守2,475且量能>21,025張→回升中性偏多」通過，本報依判準上調為「中性偏多」；「收盤≥2,475延續五線紀錄」字面通過但MA5已上移至2,478、收盤低於MA5三元，紀錄止於6日；跌破2,460未觸發；外資未回補；台積電領先大盤兌現；美光財報、週選結算如期提醒。回撤條件：收盤跌破MA10 2,442或外資第3日賣超逾5,000張即退回中性。

(3)⭐技術：RSI 56.88不變(平盤數學結果)、MACD柱+10.33連2日縮小(DIF 20.56續升)、K 72.13/D 64.09/J 88.23；MA5 2,478/MA10 2,442/MA20 2,435/MA60 2,402/MA120 2,329五線多頭排列連4日；布林2,359-2,435-2,511、寬6.26%；5日均量20,473張。

(4)⭐美股9/29費半+1.32%收12,629.16(回補9/28跌幅八成)，設備四檔全數大漲(應材+5.19%、科磊+3.89%、艾司摩爾+3.56%、科林+2.99%，皆創1個月新高)，Arm+3.65%、美光+1.05%、博通+1.58%；蘋果-2.66%、高通-1.80%、輝達-0.72%偏弱；ADR+0.90%收$456.94。韓股9/29三星+0.93%、SK海力士-0.17%。新台幣9/29收31.876(貶9.6分)。今晨夜盤個股期收2,520(+29、+1.16%)、台指期+698點(+1.46%)同向偏多；外資台指期淨空-79,029口連4日擴大；選擇權PCR 0.75(週結算前收斂)。

(5)⭐消息：外電傳台積電美國第二園區落腳德州規劃6座廠(未經證實)；美10年期殖利率逼近5.2%、市場定價Fed 10月升息(^TNX官方收盤未取得)；美國CR已延長至12/11、關門風險低；Q3法說會10/15、9月營收約10/10。

(6)立場上調為「中性偏多」(附保留)。次日檢核：收復2,500並站上MA5→上影線視為洗盤；跌破2,442或外資第3日賣逾5,000張→退回中性；夜盤2,520/+698指示能否兌現；投信轉買、自營連10買能否延續；美光財報(台北10/1凌晨)；週選結算後PCR與外資部位重讀。"""
change_color = '000000'
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

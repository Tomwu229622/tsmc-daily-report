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

today_str = '2026-09-29'
tw_price = 'NT$2,475（9/24收，連假前最後交易日；低開2,480、收盤恰落在分界2,475）'
change_pct = '-1.00%（9/24收，-25.00，落後大盤-0.28%）'
nyse_price = 'US$452.88（9/28收，+0.50%；連假三日累計+1.41%；美債10年期殖利率升至5.24%）'
volume = '14,558（9/24收，-36.2%，2026年第6低、低於5日均量37.4%）'
news_summary = """報告日2026-09-29(Tue)，連假後首個交易日。台股資料為2026-09-24(Thu)官方收盤(9/25中秋、9/28教師節休市；⚠️9/25排程未執行、本則一併補記9/24行情)；ADR、費半與美股個股為9/28正式收盤(另記連假期間美股9/24、9/25、9/28三日)；韓股為9/28收盤(9/24-25休市)；期貨為TAIFEX 9/24日盤，另含今晨夜盤(9/28 15:00至9/29約05:00)，期交所9/24晚間至9/28全無交易。

(1)⚠️9/24台股2330低開2,480、收2,475(-25、-1.00%)，恰落在前報第一順位分界2,475上；量14,558張為2026年第6低(-36.2%)，外資轉賣4,667張(約NT$115.5億，占全市場外資淨賣超NT$329.65億約35.0%)，投信連3賣1,287張、自營連9買242張，三大法人淨賣5,712張；落後加權-0.28%達0.72pp，聯發科+1.93%(連4日近2年新高)、日月光+0.87%逆勢。

(2)⭐⭐前報六項檢核：分界2,475字面守住但量縮無效檢驗→本報首次在字面達標時不依判準上調、維持「中性」並說明理由；外資轉賣且全市場同向→總經面普遍撤出；相對強弱未通過；夜盤開盤與收盤方向皆兌現(4日實證開盤4/4、收盤2/4)；技術過熱如期提醒；連假提醒適用。

(3)⭐技術：MACD柱+11.61翻正以來首度縮小(DIF 19.92續升)、RSI 56.88、J自100.11退回93.37；收盤恰等於MA5 2,475，五線多頭排列連3日但站在臨界點；布林2,358-2,433-2,507、寬度6.10%。

(4)⚠️⚠️連假期間美股三日：10年期殖利率5.162%→5.184%(9/25盤中5.230%為2007年以來最高)→5.240%；費半-0.33%/+1.41%/-1.61%三日累計-0.55%；9/28 AI安全疑慮(OpenAI代理式模型逃出容器)使Arm-8.70%、高通-7.17%、英特爾-5.67%、超微-3.61%、美光-2.61%，輝達+1.68%(加碼1,500億美元庫藏股)、ADR+0.50%收$452.88、艾司摩爾+1.58%逆勢；韓股9/28三星-5.43%、SK海力士-5.05%。ADR溢價擴至+16.30%(時差效果)。今晨夜盤CDF收2,505(+9)、台指期-167點；外資台指期淨空連3日擴大至-77,031口。

(5)⭐消息：經濟日報9/28報導大廠追加2奈米訂單10-20%、年底月產能上看12萬片；Digitimes 9/24證實高通旗艦晶片採台積電2奈米；台積電赴美投資第7度核准累計440億美元；Q3法說會官方確認10/15(四)14:00-15:30、緘默期10/5-10/14。

(6)立場維持「中性」。次日檢核：收盤守2,475且量能>9月中位數21,025張則回升中性偏多；跌破2,460則五線紀錄中斷、下調展望；外資是否回補；台積電能否領先大盤(夜盤指示個股偏多、大盤偏空)；美光美東9/30財報、週選9/30結算、9月營收約10/10。"""
change_color = 'FF0000'
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

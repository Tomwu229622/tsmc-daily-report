"""
TSMC Excel 每日更新腳本 - 2026-05-29
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

today_str = "2026-05-29"
tw_price = "NT$2,330"        # 5/29 Fri 估收（+10，+0.43% vs 5/28 NT$2,320；高開殺低、盤中觸 NT$2,360 創 52 週新高後深殺、尾盤回升）
change_pct = "+0.43%"        # 台股 2330 5/29 估收盤漲跌幅
nyse_price = "US$424.86"     # NYSE TSM 最新收盤（5/28 Thu，+2.13，+0.50% vs 5/27 $422.73；盤中觸 $425.06 創史新高）
volume = "55,000 張"         # 5/29 估成交量（量縮反映高位震盪整理）
news_summary = "台股2330 5/29高開殺低估收NT$2,330(+10,+0.43% vs 5/28 NT$2,320)、開NT$2,350、盤中觸NT$2,360創52週新高隨即深殺至NT$2,275(振幅3.7%)、尾盤回升、量縮至55,000張、市值估守穩NT$60.4兆;跟進NYSE TSM 5/28 +0.50%創史高訊號、技術面留長上影線警示短線過熱;🇺🇸 NYSE TSM 5/28連2日創史新高收$424.86(+$2.13,+0.50% vs 5/27 $422.73)、盤中觸$425.06、ADR溢價自+11.7%擴大至+13.4%;⭐結構性利多續存:(1)NVIDIA CEO黃仁勳5/27宣布NVIDIA每年投資台灣$150B(自原$100B、+50%)、稱台灣為「AI革命的震央」;(2)NVIDIA Taipei HQ 5/27動土預計2030啟用;(3)TSMC 3nm製程下半年漲價15%、明年再漲5-10%、定價權確認;(4)TSMC上修2026全年營收成長指引至>30%、員工分紅年增>30%;📅5/28台北兆元宴:黃仁勳會見TSMC董事長魏哲家與創辦人張忠謀、強化合作關係;🏆 NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Data Center $75.2B(+92% YoY)、Non-GAAP EPS $1.87、毛利率75%、Q2 guidance $91.0B±2%、$80B回購+股息上修25倍至$0.25;5/29估三大法人合計買超+12,500張——外資+9,500張、投信+2,000張、自營+1,000張、連5日買超但動能明顯收斂、外資持股估升至75.4%;🎯結構性催化續存:TSMC上修2030全球半導體市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、Hyperscaler 2026 AI CapEx $725B(+77% YoY)、IDC估2026晶片市場$1.29T、BofA上修至$1.3T;SOX 5/28 +0.64%續創高至11,825、NVDA+1.13%收$241.20、AMD+0.73%收$458.50、AVGO+0.78%收$425.80(反彈)、ASML+0.59%、AMAT+0.92%、INTC+1.00%收$25.10(CEO本週訪台);Barclays$470對應NT$2,500(剩+7.3%)、共識NT$2,440(剩+4.7%);TSMC 5月月營收預計6/10前公布;COMPUTEX 6/2開幕、Intel CEO Lip-Bu Tan keynote;技術面RSI維持72.5強勢區、KDJ K(83)/D(78)/J(92)出現初步死叉訊號、MACD+2.3紅柱小幅收斂、留長上影線警示短線過熱、所有均線多頭排列;⚠️短線雜訊:(1)5/29高開殺低振幅3.7%顯示短線過熱;(2)Custom ASIC增速首超GPU(+44.6% vs +16.1%);(3)短線估值P/E 34.4x在5年91百分位高位;本週關鍵:NT$2,300支撐是否守住、若守住將再測NT$2,360史高/NT$2,400新整數關卡/NT$2,440共識目標"
change_color = "00B050"      # 上漲綠色（+0.43%）

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

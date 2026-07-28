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

today_str = "2026-07-28"
tw_price = "NT$2,350（7/27收）"   # 台股2330 7/27 Mon官方收NT$2,350(0、0.00% vs 7/24 NT$2,350、開2,330/高2,365/低2,330、量28,939張TWSE口徑、成交額678.85億、成交92,753筆、TWSE確認)——低開-20後盤中翻紅高2,365、尾盤收平盤、守住60MA 2,345、美股重挫下明顯抗跌
change_pct = "0.00%（7/27收）"    # 7/27官方漲跌幅0.00%（TWSE、低開翻紅收平盤、下影線守穩）
nyse_price = "US$399.09（7/27收）" # NYSE TSM 7/27 Mon收$399.09(-$4.32、-1.07% vs 7/24 $403.41、Yahoo API確認)——失守$400但相對抗跌；SOX同日-2.23%收11,554.88自史高-21.0%正式跌入熊市、NVDA -5.00%失守$200
volume = "28,939（7/27收）"    # 台股7/27官方成交量28,939張（TWSE口徑、量回增前日24,810張；成交筆數自195,859驟降至92,753、恐慌小單退潮）、成交額678.85億
news_summary = "報告日2026-07-28(Tue)。📊個股抗跌vs外部熊市共振：台股2330 7/27官方收NT$2,350(0、0.00%、開2,330/高2,365/低2,330、量28,939張、額678.85億、92,753筆、TWSE確認)——低開-20後盤中翻紅高2,365、尾盤收平盤、守住60MA 2,345；成交筆數自195,859驟降至92,753、恐慌小單退潮。✅大盤同日僅小跌-20.65點(-0.05%)收43,634.19——週五恐慌未蔓延。✅籌碼明顯改善：外資對2330賣超自-9,637大幅收斂至-2,662張(連4賣但減速逾七成、T86官方)、投信+567連2買、自營+1,410連3買、三大合計僅-685張；全市場BFI82U由賣轉買：三大法人買超6.6億——外資+80.4億轉買、投信+4.8億。⚠️美股7/27 AI高Beta續挫：TSM -1.07%收$399.09失守$400(相對抗跌)、SOX -2.23%收11,554.88自史高-21.0%正式跌入熊市——NVDA -5.00%失守$200、AMD -5.17%、ASML -5.79%；惟AAPL +1.17%連2漲、AVGO/QCOM/ARM翻紅——市場評論指屬槓桿部位強制賣出、跌勢集中AI權值。⚠️韓股7/28早盤重挫：SK海力士財報日開盤急殺逾-10%、三星-8.9%(盤中)——SK Q2市場預期營益率75-77%創紀錄仍遭拋售、「財報優仍被賣」極端化、今日台股開盤最大逆風。⭐油價續回落WTI~$82/布蘭特~$88、VIX 18.67持平、10Y美債4.64%。📊ADR溢價連3日收斂至+9.7%(匯率~32.31)。技術面(官方收盤計算)：RSI 46.4持平中軸下、MACD綠柱-12.7且DIF -0.8跌破零軸(中期動能第一道裂痕)、KDJ死叉延續(K41.1/D43.7/J35.9)；支撐2,345(60MA)/2,330(7/27低)/2,320/2,307(布林下軌)、壓力2,383(5MA)/2,400-2,405/2,414(20MA)/2,440-2,445。⭐觀察：韓股重挫外溢下守60MA、外資賣壓收斂能否延續、SK財報電話會HBM定調、MSFT/META(7/29)AAPL/AMZN(7/30)財報。"
change_color = "000000"        # 黑色（7/27官方0.00%平盤）

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

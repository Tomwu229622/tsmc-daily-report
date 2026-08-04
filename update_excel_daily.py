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

today_str = "2026-08-04"
tw_price = "NT$2,370（8/3收）"   # 台股2330 8/3 Mon官方收NT$2,370(-55、-2.27% vs 7/31漲停NT$2,425、開2,390/高2,395/低2,365、量35,210張TWSE口徑、成交額836.73億、成交174,489筆、TWSE確認)——漲停後首日量縮健康換手、大盤逆勢+0.62%收43,386.41
change_pct = "-2.27%（8/3收）"    # 8/3官方-2.27%（TWSE、大盤同日+266.66點/+0.62%收43,386.41——資金輪動：聯發科/華邦電雙漲停、日月光+9.91%）
nyse_price = "US$406.11（8/3收）" # NYSE TSM 8/3 Mon收$406.11(+$1.86、+0.46% vs 7/31 $404.25、Yahoo API確認、連3日收高)——美股大型科技暴漲：MSFT+4.93%、META+6.02%、GOOGL+4.88%、AMZN+4.58%；NVDA+2.93%收$206.64；AAPL-1.78%；SOX+1.05%收11,430.35、VIX 15.86
volume = "35,210（8/3收）"    # 台股8/3官方成交量35,210張（TWSE口徑、僅7/31漲停量69,478之半——拉回量縮屬健康換手）、成交額836.73億
news_summary = "報告日2026-08-03(Mon)；8/1-8/2週末休市、7/31(Fri)日報缺漏、本列一併補述7/30-31兩交易日。⭐股災後V型反轉：台股7/30先止穩（大盤-105.88收39,933.30、2330 +5收2,205、量51,372張）；7/31史詩級反彈——大盤+3,186.45點(+7.98%)收43,119.75創史上最大單日漲點（7月全月仍累跌約3,006點）；2330跳空高開2,350、尾盤鎖漲停收NT$2,425(+220、+9.98%、開2,350/高2,425/低2,345、量69,478張為7/17以來最大、額1,666.62億、215,037筆、TWSE確認)——媒體稱台股史上第3次2330漲停；供應鏈全面漲停：聯發科+9.89%、日月光+9.90%、聯電+10.00%、華邦電+9.70%。✅籌碼一日翻多（最關鍵訊號）：7/31外資買超+13,780張終結連7賣（7/30已收斂至-3,958）、投信+7,063連6買爆量放大7.5倍、自營+2,724連7買、三大合計+23,568張（T86官方）；全市場BFI82U自7/30 -495.4億翻正至+873.1億（外資+675.5億、投信+360.2億）。⭐美股7/30暴力反彈：LRCX財報後+17.99%（1999年來最佳單日）、MU+18.36%、MSFT+15.51%、AMAT+14.97%、AMD+13.0%、TSM+7.64%收$403.31、SOX+8.19%；7/31高檔續穩：TSM+0.23%收$404.25、SOX收11,311.08自史高-22.7%、NVDA+2.93%收復$200、VIX兩日-22.6%降至15.99。📊7/30盤後財報兩樣情：AMZN營收$200.6B(+20%)、AWS+37%至$42.2B創2021來最快、全年CapEx上看$220B——7/31 +15.32%（2012年來最大漲幅）；AAPL營收$109.4B/EPS$2.02雙超預期惟9月季指引低於共識＋Cook坦承「晶片供應吃緊、低估自身需求」——7/31 -7.35%收$308.91（Cook最後法說、9/1 Ternus接任）——AI CapEx連續驗證、AAPL供應吃緊反證需求強勁、對TSMC淨效果偏正面。⭐韓股7/31暴漲：KOSPI收6,595.45(+17.9%、Yahoo API)、SK海力士+29.95%、三星+26.81%（SK會長加碼買股＋空頭/槓桿ETF回補）；⚠️今晨(8/3 09:40)回吐：KOSPI -4.78%、SK -7.39%、三星-7.62%。📊BWIBBU官方P/E 32.60x/P/B 10.67x/殖利率0.91%；ADR溢價收斂至+8.02%（匯率32.40、今晨升至32.29）。技術面依官方收盤重算：漲停一舉收復全部均線（5MA 2,292/20MA 2,382/60MA 2,348/120MA 2,143）、KDJ黃金交叉(K42.6/D32.9/J62.1)、RSI 54.6回中性、MACD綠柱自-53.4收斂至-25.9（DIF -21.2仍零軸下）；支撐2,382/2,350/2,345/2,292、壓力2,445/2,535/2,543。整體信號：中性偏多（漲停後首日驗證）。📈今晨(8/3 09:03官方MIS)：大盤43,007.35(-0.26%)、2330開2,390現2,370(-2.27%)溫和拉回、遠小於韓股。⭐觀察：守2,382(20MA)/2,345(7/31低)、外資買超延續性、韓股回吐幅度、232條款（7月底窗口已過仍未查得公布）與台美關稅談判（傳20%→~15%）。全程TWSE/Yahoo官方API取數。"
change_color = "FF0000"        # 紅色（8/3官方-2.27%下跌）

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

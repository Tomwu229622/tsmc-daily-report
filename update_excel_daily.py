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

today_str = "2026-08-06"
tw_price = "NT$2,405（8/5收）"   # 台股2330 8/5 Wed官方收NT$2,405(+85、+3.66% vs 8/4 NT$2,320、開2,385/高2,415/低2,370、量36,782張TWSE口徑、成交額881.58億、成交119,876筆、TWSE確認)——補漲兌現、一舉收復20MA 2,369與60MA 2,352；大盤同日大漲+2.89%收44,611.60
change_pct = "+3.66%（8/5收）"    # 8/5官方+3.66%（TWSE、大盤同日+1,250.94點/+2.89%收44,611.60——補漲大反攻、完全回應前日SOX+6.55%與ADR溢價+16.4%訊號）
nyse_price = "US$414.00（8/5收）" # NYSE TSM 8/5 Wed收$414.00(-$3.17、-0.76% vs 8/4 $417.17、Yahoo API確認、連4日收高中止)——美股8/5半導體暴漲後回吐：SOX-1.40%收12,008.88守12,000、LRCX-3.25%、AMKR-3.60%、AMAT-2.26%；NVDA逆勢+3.44%收$219.22（SpaceX獨家採用）；AMD財報定價-7.04%收$482.05；S&P-0.17%；VIX 15.81
volume = "36,782（8/5收）"    # 台股8/5官方成交量36,782張（TWSE口徑、低於5日均量46,773——大漲量未爆、惜售型上漲）、成交額881.58億
news_summary = "報告日2026-08-06(Thu)。⭐補漲兌現、季線失而復得：台股8/5大反攻——2330跳空開2,385、收NT$2,405(+85、+3.66%、高2,415/低2,370、量36,782張低於5日均量46,773、額881.58億、119,876筆、TWSE確認)、一舉收復20MA 2,369與60MA 2,352（季線失守僅1日）；大盤大漲+1,250.94點(+2.89%)收44,611.60——完全回應前日SOX+6.55%與ADR溢價+16.4%的補漲訊號；聯發科+3.49%站上4,000、華邦電+7.64%連4大漲、日月光+1.37%。✅籌碼翻多：外資翻買+12,251張（終結連2賣、T86官方）、投信-1,205張（連8買中止、獲利調節）、自營-1,099、三大合計+9,948張；全市場BFI82U合計+1,030.3億大買（外資+903.1億近期最大單日回補、投信+112.7億）。📊美股8/5暴漲後回吐整理：SOX-1.40%收12,008.88守住12,000——設備股回落（LRCX-3.25%、AMKR-3.60%、AMAT-2.26%、ASML-1.97%、ARM-2.13%）；NVDA逆勢+3.44%收$219.22連5高——SpaceX宣布獨家採用NVIDIA處理器、2026年底算力擴至2GW+、Vera Rubin NVL72獲點名；AMD財報後正式定價-7.04%收$482.05（Q2全面超標仍遭高估值賣壓、P/E約53倍）；S&P-0.17%、那指-0.83%、VIX降至15.81；TSM-0.76%收$414.00連4高中止。⚠️韓股今晨大幅回吐：8/5大漲（SK+5.77%收166.8萬、三星+2.50%收24.6萬、市場關注股東回饋方案）後、今晨8/6開盤重挫——SK盤中-6.83%至155.4萬、三星-2.64%（Yahoo 08:33快照）——今日台股情緒逆風。📊產業：2nm報價傳漲10-20%（自傳言50%收斂、未官方證實）；熊本JASM廠震後全面復產；ABF基板傳被鎖產能至2028。📊BWIBBU官方P/E 32.33x/P/B 10.59x/殖利率0.91%；ADR溢價自+16.38%大幅收斂至+10.97%（$414.00 vs理論值$373.09、匯率32.23）——背離健康解除。技術面依官方收盤重算（至8/5）：RSI 53.0回中軸上、MACD黃金交叉（DIF-11.3上穿DEA-11.5、柱+0.5翻紅）、KDJ金叉延續（K66.6/D51.7/J96.3偏高過熱）、站回全部均線（5MA 2,345/20MA 2,369/60MA 2,352/120MA 2,157）；支撐2,371/2,370/2,352/2,345/2,320、壓力2,415/2,425/2,522/2,535。整體信號：中性偏多。⭐觀察：韓股重挫下守2,370-2,371缺口與否、外資翻買延續性、今晚SOX守12,000、232條款（仍未查得公布、未能證實）。全程TWSE/Yahoo官方API取數。"
change_color = "00B050"        # 綠色（8/5官方+3.66%上漲）

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

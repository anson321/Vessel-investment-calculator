import io
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# --- 1. 頁面基本配置 ---
st.set_page_config(
    page_title="船東 S&P 投資快速試算器", page_icon="🚢", layout="wide"
)

# 極簡金融風格 CSS
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #1F2937;
    }
    .disclaimer-banner {
        font-size: 0.78rem;
        color: #64748B;
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 4px;
        padding: 6px 12px;
        margin-bottom: 12px;
    }
    .main-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 0.85rem;
        color: #4B5563;
        margin-bottom: 1.2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        background-color: #F1F5F9;
        border-radius: 4px 4px 0px 0px;
        color: #4B5563;
        font-size: 0.9rem;
        font-weight: 500;
        padding: 0px 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        font-weight: 600;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 4px;
        border: none;
        font-size: 0.85rem;
        font-weight: 500;
        padding: 6px 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 免責聲明
st.markdown(
    '<div class="disclaimer-banner">'
    '<b>免責聲明 Disclaimer：</b>本計算器僅供 S&P 投資快速評估參考，實際結果受市場運價、利息變動與船舶實際狀況影響，投資人須自行承擔相關風險。'
    '</div>',
    unsafe_allow_html=True
)

# --- 2. 側邊欄：極簡 5 大核心輸入 ---
st.sidebar.header("🎯 投資 5 大核心參數")

vessel_name = st.sidebar.text_input("船舶 / 項目名稱", value="64,000 DWT Ultramax")
purchase_price = st.sidebar.number_input("1. 船舶購入總價 (USD)", value=33000000, step=500000)
ltv_pct = st.sidebar.slider("2. 銀行融資成數 LTV (%)", min_value=0, max_value=90, value=60, step=5, help="拉到 0 代表全額現金購船")
gross_tc_rate = st.sidebar.number_input("3. 預估平均毛日租金 Gross TC (USD/day)", value=17000, step=500)
lease_years = st.sidebar.slider("4. 專案持有與融資年限 (Years)", min_value=1, max_value=15, value=5)

exit_mode = st.sidebar.radio("5. 期末處分估算方式", ["預估二手市場賣價", "輕排水量 (LDT) 拆船價"])

if exit_mode == "預估二手市場賣價":
    expected_resale_value = st.sidebar.number_input("期末預估二手賣價 (USD)", value=22000000, step=500000)
else:
    vessel_ldt = st.sidebar.number_input("船舶輕排水量 LDT (MT)", value=10000, step=500)
    scrap_price_per_lt = st.sidebar.number_input("拆船單價 (USD/LT)", value=550, step=10)
    long_tons = vessel_ldt / 1.016
    expected_resale_value = long_tons * scrap_price_per_lt
    st.sidebar.info(f"💡 折算期末拆船殘值：**${expected_resale_value:,.0f}** USD")

# 微調參數（預設最貼近現實狀況的標準值，收折不打擾）
with st.sidebar.expander("⚙️ 實務營運與融資基準設定", expanded=False):
    daily_opex = st.number_input("每日營運成本 OPEX (USD/day, 含船管費)", value=5500, step=100)
    commission_pct = st.number_input("租費傭金 (Brokerage Comm %)", value=3.75, step=0.25)
    annual_off_hire_days = st.number_input("每年預估停航天數 Off-hire (天)", value=10, step=1)
    
    if ltv_pct > 0:
        interest_rate_pct = st.number_input("貸款年利率 (SOFR + Margin %)", value=6.25, step=0.25)
        balloon_pct = st.number_input("貸款期末尾款比例 Balloon (% of Debt)", value=30, step=5, help="期末一次還清之尾款佔總貸款比例")
    else:
        interest_rate_pct, balloon_pct = 0.0, 0.0
    
    hurdle_rate_pct = st.number_input("船東目標折現率 WACC (%)", value=8.0, step=0.5)

# --- 3. 後台財務計算引擎 ---
debt_amount = purchase_price * (ltv_pct / 100.0)
equity_injected = purchase_price - debt_amount

# 營運天數與淨日租金
operating_days_per_year = 365.0 - annual_off_hire_days
net_tc_rate = gross_tc_rate * (1.0 - commission_pct / 100.0)

# 年營運收入與 OPEX
annual_revenue = net_tc_rate * operating_days_per_year
annual_opex = daily_opex * 365.0
annual_net_operating_income = annual_revenue - annual_opex  # EBITDA 概念

# 融資還本付息 (Amortization Schedule)
total_months = lease_years * 12
monthly_interest_rate = (interest_rate_pct / 100.0) / 12.0

if ltv_pct > 0 and lease_years > 0:
    balloon_amount = debt_amount * (balloon_pct / 100.0)
    principal_to_amortize = debt_amount - balloon_amount
    monthly_principal_pay = principal_to_amortize / total_months
else:
    balloon_amount = 0.0
    monthly_principal_pay = 0.0

# 建立月度現金流與還本付息表
monthly_records = []
monthly_net_cashflows = []
current_debt = debt_amount

for m in range(1, total_months + 1):
    start_debt = current_debt
    monthly_interest = start_debt * monthly_interest_rate
    monthly_principal = monthly_principal_pay if start_debt >= monthly_principal_pay else start_debt
    end_debt = max(0.0, start_debt - monthly_principal)
    current_debt = end_debt
    
    monthly_rev = annual_revenue / 12.0
    monthly_op_cost = annual_opex / 12.0
    monthly_debt_service = monthly_principal + monthly_interest
    monthly_net_cf = monthly_rev - monthly_op_cost - monthly_debt_service
    
    monthly_net_cashflows.append(monthly_net_cf)
    monthly_records.append({
        "月份": f"第 {m} 個月",
        "期初未還本金": start_debt,
        "月度租金收入": monthly_rev,
        "月度 OPEX": monthly_op_cost,
        "本金支付": monthly_principal,
        "利息支付": monthly_interest,
        "月度還本付息 (BBC)": monthly_debt_service,
        "月度淨現金流": monthly_net_cf,
        "期末未還本金": end_debt,
    })

df_monthly = pd.DataFrame(monthly_records)

# 逐年現金流彙整
yearly_records = []
annual_net_cashflows = []

for y in range(1, lease_years + 1):
    m_start = (y - 1) * 12
    m_end = y * 12
    y_df = df_monthly.iloc[m_start:m_end]
    
    y_rev = y_df["月度租金收入"].sum()
    y_opex = y_df["月度 OPEX"].sum()
    y_debt_service = y_df["月度還本付息 (BBC)"].sum()
    y_net_cf = y_df["月度淨現金流"].sum()
    y_start_debt = y_df["期初未還本金"].iloc[0]
    
    annual_net_cashflows.append(y_net_cf)
    yearly_records.append({
        "年份": f"第 {y} 年",
        "期初未還本金": y_start_debt,
        "年度淨租金收入": y_rev,
        "年度 OPEX": y_opex,
        "年度還本付息": y_debt_service,
        "年度營運淨現金流": y_net_cf,
    })

df_yearly = pd.DataFrame(yearly_records)

# 期末清算 (Terminal Exit)
net_terminal_exit = expected_resale_value - balloon_amount
total_cash_profit = sum(annual_net_cashflows) + net_terminal_exit - equity_injected

# 計算 IRR 與 NPV
project_cfs = [-equity_injected] + annual_net_cashflows.copy()
project_cfs[-1] += net_terminal_exit

r = hurdle_rate_pct / 100.0
npv_value = sum(cf / ((1 + r) ** t) for t, cf in enumerate(project_cfs))

def compute_irr(cfs, iterations=1000):
    if equity_injected <= 0:
        return None
    rate = 0.10
    for _ in range(iterations):
        npv = sum(cf / ((1 + rate) ** t) for t, cf in enumerate(cfs))
        d_npv = sum(-t * cf / ((1 + rate) ** (t + 1)) for t, cf in enumerate(cfs))
        if abs(d_npv) < 1e-10:
            break
        new_rate = rate - npv / d_npv
        if abs(new_rate - rate) < 1e-7:
            return new_rate * 100.0
        rate = new_rate
    return rate * 100.0

irr_value = compute_irr(project_cfs)

# 首年保本運價計算 (Gross Breakeven)
y1_debt_service = df_yearly.loc[0, "年度還本付息"]
first_year_cost = y1_debt_service + annual_opex
gross_breakeven = (first_year_cost / operating_days_per_year) / (1.0 - commission_pct / 100.0)

# --- 4. 生成 Excel 匯出檔 ---
def generate_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S&P Investment Model"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Segoe UI", size=10)
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    border_thin = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))

    ws.cell(row=1, column=1, value=f"{vessel_name} - S&P 投資財務評估表").font = font_title
    ws.cell(row=2, column=1, value=f"自備款: ${equity_injected:,.0f} | IRR: {irr_value:.2f}% | 美金純利: ${total_cash_profit:,.0f}").font = font_body

    headers = ["月份", "期初未還本金", "月度租金收入", "月度 OPEX", "本金支付", "利息支付", "月度還本付息", "月度淨現金流", "期末未還本金"]
    for col_num, h_text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, r in enumerate(monthly_records):
        row = idx + 5
        ws.cell(row=row, column=1, value=r["月份"]).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2, value=round(r["期初未還本金"]))
        ws.cell(row=row, column=3, value=round(r["月度租金收入"]))
        ws.cell(row=row, column=4, value=round(r["月度 OPEX"]))
        ws.cell(row=row, column=5, value=round(r["本金支付"]))
        ws.cell(row=row, column=6, value=round(r["利息支付"]))
        ws.cell(row=row, column=7, value=round(r["月度還本付息 (BBC)"]))
        ws.cell(row=row, column=8, value=round(r["月度淨現金流"]))
        ws.cell(row=row, column=9, value=round(r["期末未還本金"]))

        for col_num in range(1, 10):
            cell = ws.cell(row=row, column=col_num)
            cell.font = font_body
            cell.border = border_thin
            if col_num > 1:
                cell.number_format = '$#,##0;($#,##0);"-"'

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. 主頁面抬頭與下載按鈕 ---
title_col, btn_col = st.columns([3.5, 1])

with title_col:
    st.markdown(f'<div class="main-title">🚢 {vessel_name} - S&P 投資決策快速評估</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">極簡高質感海事金融試算模組｜快速上手與回報評估</div>', unsafe_allow_html=True)

with btn_col:
    st.download_button(
        label="📊 下載完整還本付息表 (.xlsx)",
        data=generate_excel(),
        file_name=f"{vessel_name}_Investment_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")

# --- 6. 核心 KPI 數據卡片 (6 欄平鋪) ---
def render_kpi(col, title, value, subtext="", status="neutral"):
    if status == "profit":
        bg, border, text, sub = "#ECFDF5", "#A7F3D0", "#047857", "#065F46"
    elif status == "loss":
        bg, border, text, sub = "#FEE2E2", "#FCA5A5", "#991B1B", "#B91C1C"
    else:
        bg, border, text, sub = "#F8FAFC", "#E2E8F0", "#1E3A8A", "#64748B"

    card_html = f"""
    <div style="background-color: {bg}; border: 1px solid {border}; padding: 10px 12px; border-radius: 6px; height: 95px; display: flex; flex-direction: column; justify-content: space-between;">
        <div style="font-size: 0.78rem; color: #4B5563; font-weight: 600;">{title}</div>
        <div style="font-size: 1.25rem; color: {text}; font-weight: 700;">{value}</div>
        <div style="font-size: 0.72rem; color: {sub};">{subtext}</div>
    </div>
    """
    col.markdown(card_html, unsafe_allow_html=True)

col1, col2, col3, col4, col5, col6 = st.columns(6)

equity_sub_text = f"全額現金購船" if ltv_pct == 0 else f"貸款 {ltv_pct}% (${debt_amount:,.0f})"
render_kpi(col1, "自有資金投入 (Net Equity)", f"${equity_injected:,.0f}", equity_sub_text)
render_kpi(col2, "首年保本運價 (Breakeven)", f"${gross_breakeven:,.0f}/天", f"預設租金 ${gross_tc_rate:,.0f}/天")

irr_display = f"{irr_value:.2f}%" if irr_value is not None else "N/A"
render_kpi(col3, "股權內部報酬率 (IRR)", irr_display, f"目標 WACC {hurdle_rate_pct}%")

render_kpi(col4, "專案淨現值 (NPV)", f"${npv_value:,.0f}", "現值折現淨效益")

if total_cash_profit >= 0:
    render_kpi(col5, "實質美金純利 (Cash Profit)", f"+${total_cash_profit:,.0f}", f"持有 {lease_years} 年淨賺美金", status="profit")
else:
    render_kpi(col5, "實質美金虧損 (Cash Loss)", f"-${abs(total_cash_profit):,.0f}", f"持有 {lease_years} 年預估虧損", status="loss")

render_kpi(col6, "期末淨回收金額", f"${net_terminal_exit:,.0f}", f"賣船 ${expected_resale_value:,.0f} 扣尾款")

st.markdown("<br/>", unsafe_allow_html=True)

# --- 7. 功能分頁 ---
tab1, tab2, tab3 = st.tabs([
    "📊 逐年營運財務預測 (Yearly Forecast)",
    "🗓️ 月度還本付息與現金流明細 (Monthly Schedule)",
    "💡 投資成數與評估分析 (Investment Pitch)",
])

# TAB 1: 逐年預測表
with tab1:
    st.subheader("逐年營運財務預測明細表")
    
    df_yearly_disp = df_yearly.copy()
    for col in ["期初未還本金", "年度淨租金收入", "年度 OPEX", "年度還本付息", "年度營運淨現金流"]:
        df_yearly_disp[col] = df_yearly_disp[col].apply(lambda x: f"${x:,.0f}")
        
    st.table(df_yearly_disp)
    
    st.markdown("##### 📝 核心計算邏輯與公式說明：")
    st.markdown(f"""
*   **年度淨租金收入**：以預設毛日租金 **${gross_tc_rate:,.0f}/天** 扣除 **{commission_pct}%** 傭金後，按每年營運 **{operating_days_per_year:.0f} 天**（扣除 {annual_off_hire_days} 天 Off-hire）計算。
*   **首年保本運價**：當年度維繫「不虧損」所需的最低市場開價門檻，公式為：
""")
    st.latex(r"\text{保本日租金} = \frac{(\text{每日還本付息} + \text{每日 OPEX}) \times 365}{\text{每年營運天數} \times (1 - \text{傭金 \%})}")

# TAB 2: 月度細節與折線圖
with tab2:
    st.subheader("累積現金流回收曲線與月度明細")
    
    x_m = [r["月份"] for r in monthly_records]
    cum_cfs = np.cumsum([-equity_injected] + monthly_net_cashflows)[1:]
    
    fig_cum = go.Figure()
    fig_cum.add_trace(go.Scatter(x=x_m, y=cum_cfs, mode="lines", name="累積現金流 (USD)", line=dict(color="#1E3A8A", width=2.5)))
    fig_cum.add_hline(y=0, line_dash="dash", line_color="#DC2626", annotation_text="保本平衡線")
    fig_cum.update_layout(title="累積營運現金流走勢圖", template="plotly_white")
    st.plotly_chart(fig_cum, use_container_width=True)
    
    st.markdown("##### 🗓️ 月度詳細還本付息表")
    df_m_disp = df_monthly.copy()
    for col in ["期初未還本金", "月度租金收入", "月度 OPEX", "本金支付", "利息支付", "月度還本付息 (BBC)", "月度淨現金流", "期末未還本金"]:
        df_m_disp[col] = df_m_disp[col].apply(lambda x: f"${x:,.0f}")
    st.dataframe(df_m_disp, use_container_width=True, height=350)

# TAB 3: 投資評估建議
with tab3:
    st.subheader("💡 船東與投資人戰略評估重點")
    
    if irr_value is not None and irr_value >= hurdle_rate_pct:
        st.success(f"🟢 **投資可行性極高**：專案內部報酬率 (IRR) 高達 **{irr_value:.2f}%**，超越目標資本成本門檻 ({hurdle_rate_pct}%)。")
    elif irr_value is not None:
        st.warning(f"🟡 **投資回報偏低**：專案 IRR 為 **{irr_value:.2f}%**，未達目標回報門檻 ({hurdle_rate_pct}%)，建議向賣方議價購船總價。")
    else:
        st.error("🔴 **現金流無法覆蓋成本**：現行運價與融資條件下無法產生正向內部報酬率。")

    st.markdown("##### 🔍 核心財務結論：")
    st.markdown(f"""
*   **資金門檻**：投資人需準備 **${equity_injected:,.0f} 美金** 之自有資金。
*   **營運利差**：預設運價 **${gross_tc_rate:,.0f}/天** 高於首年保本點 **${gross_breakeven:,.0f}/天**，每日享有 **+${gross_tc_rate - gross_breakeven:,.0f}/天** 的安全邊際利差。
*   **期末清算**：持有 {lease_years} 年期滿賣船，扣除銀行尾款 (${balloon_amount:,.0f}) 後，預估可淨拿回 **${net_terminal_exit:,.0f} 美金** 現金。
*   **全期美金純利**：營運累積淨租金加上賣船淨回收，全期專案共為股權投資人淨賺 **+${total_cash_profit:,.0f} 美金**。
""")
import io
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

# --- ReportLab PDF 套件與繁體中文字型註冊 ---
pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))

# --- 1. 頁面基本設定 ---
st.set_page_config(
    page_title="船東 S&P 投資決策計算器", page_icon="🚢", layout="wide"
)

# --- 2. 高質感極簡海事金融 CSS ---
st.markdown(
    """
    <style>
    /* 全局字體與配色 */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #1F2937;
    }
    
    /* 頂部免責聲明小字區塊 */
    .disclaimer-banner {
        font-size: 0.78rem;
        color: #64748B;
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 4px;
        padding: 6px 12px;
        margin-bottom: 12px;
        line-height: 1.4;
    }

    /* 主標題樣式 */
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

    /* 單色極簡分頁 (Tabs) 樣式重構 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        white-space: pre-wrap;
        background-color: #F1F5F9;
        border-radius: 4px 4px 0px 0px;
        color: #4B5563;
        font-size: 0.9rem;
        font-weight: 500;
        padding: 0px 16px;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        font-weight: 600;
    }
    
    /* 按鈕樣式極簡化 */
    .stButton>button, .stDownloadButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 4px;
        border: none;
        font-size: 0.85rem;
        font-weight: 500;
        padding: 6px 12px;
        width: 100%;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #1E40AF;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- 3. 頁面最上方免責聲明區塊 ---
st.markdown(
    '<div class="disclaimer-banner">'
    '<b>免責聲明 Disclaimer：</b>本報告僅供參考，實際結果可能有所不同，投資人須自行承擔所有相關風險。投資前請務必評估自身的風險承受能力。<br/>'
    'This report is for reference only. Actual results may vary, and investors assume all associated risks. Please assess your personal risk tolerance prior to investing.'
    '</div>',
    unsafe_allow_html=True
)

# --- 側邊欄：0. 船型評估模式 ---
st.sidebar.header("0. 船型評估模式 (Vessel Profile)")

vessel_profile = st.sidebar.radio(
    "選擇評估船型模式",
    [
        "通用散裝船 (Dry Bulk Carrier)",
        "水泥船 / 利基特種船 (Cement Carrier)",
    ],
    help="通用散裝船上限 30 歲；水泥船/特種船支援最大 45 歲船齡與專用設備殘值加成。"
)

is_cement_mode = "水泥船" in vessel_profile
max_age_limit = 45 if is_cement_mode else 30

# --- 側邊欄：1. 船舶與融資參數 ---
st.sidebar.header("1. 船舶與融資參數 (Financing)")

vessel_name_default = "20,000 DWT Cement Carrier" if is_cement_mode else "64,500 DWT Ultramax"
vessel_name = st.sidebar.text_input(
    "船舶名稱 / 項目名稱",
    value=vessel_name_default,
    help="項目或船舶名稱，用於報表抬頭識別。"
)

purchase_price_default = 28000000 if is_cement_mode else 33100000
purchase_price = st.sidebar.number_input(
    "船舶購入總價 (USD)",
    value=purchase_price_default,
    step=500000,
)

delivery_cost = st.sidebar.number_input(
    "交接與啟航總費用 Delivery Cost (USD)",
    value=500000,
    step=50000,
    help="交船日 (Day 0) 必須支付之費用：含殘油買回 (ROB)、船員遣送安排、改旗改級規費。"
)

working_capital_days = st.sidebar.slider(
    "Day 0 預留營運資金天數 Working Capital (天)",
    min_value=0,
    max_value=60,
    value=30,
    step=5,
    help="預留 15~30 天之每日 OPEX 與還本付息作為 Day 0 營運資金墊款，防止交船初期租金入帳位移 (Collection Lag) 造成現金流斷裂。期末全額回收。"
)

ltv_pct = st.sidebar.slider(
    "融資成數 LTV (%)",
    min_value=0,
    max_value=120,
    value=54,
    step=1,
)

lease_years = st.sidebar.slider(
    "融資與持有年限 Hold Period (Years)",
    min_value=1,
    max_value=10,
    value=5,
    help="專案投資持有與融資期限（期末自動進行賣船結算與現金流清算）。"
)

if ltv_pct > 0:
    sofr_rate = st.sidebar.number_input("預估基準利率 SOFR (%)", value=4.25, step=0.25)
    margin_rate = st.sidebar.number_input("加碼利差 Margin (%)", value=2.00, step=0.25)
    facility_fee_pct = st.sidebar.number_input("前端融資手續費 Facility Fee (%)", value=1.00, step=0.25)
    purchase_obligation = st.sidebar.number_input(
        "貸款期末尾款 Balloon (Purchase Option) (USD)",
        value=10000000 if is_cement_mode else 11000000,
        step=500000,
        help="銀行貸款期末剩餘未還本金尾款 (Balloon)，或光船租賃合約期末過戶買回金額 (Purchase Option / Obligation)。"
    )
else:
    sofr_rate, margin_rate, facility_fee_pct, purchase_obligation = 0.0, 0.0, 0.0, 0.0

# --- 側邊欄：2. 營運成本與進塢特檢 ---
st.sidebar.header("2. 營運成本與進塢特檢 (OPEX & Drydock)")

opex_base_default = 6000 if is_cement_mode else 5000
opex_base = st.sidebar.number_input("基礎每日營運成本 Base OPEX (USD/day)", value=opex_base_default, step=100)

management_fee_monthly = st.sidebar.number_input(
    "委外船管費 Technical Management Fee (USD/month)",
    value=9000,
    step=500,
)
daily_management_fee = (management_fee_monthly * 12.0) / 365.0

opex_escalation_pct = st.sidebar.number_input("OPEX 年度通膨率 (%)", value=2.0, step=0.5)

drydock_mode = st.sidebar.radio(
    "交船時進塢狀況 (Drydock Status at Delivery)",
    ["剛完成塢特檢 (Fresh Out of SS/IS)", "常態均攤預提 (Flat Daily Reserve)"],
    help="若選擇『剛完成塢特檢』，前2.5年 (M1-M30) 免平滑預提進塢費；第30個月發生一次性中間檢支出，第31個月起開始攤提下次大特檢。"
)

ss_cost_5yr = st.sidebar.number_input("預估 5 年大特檢總費用 SS (USD)", value=1100000 if is_cement_mode else 1000000, step=100000)
is_cost_25yr = st.sidebar.number_input("預估 2.5 年中間檢總費用 IS (USD)", value=500000 if is_cement_mode else 400000, step=50000)

daily_is_reserve = is_cost_25yr / (2.5 * 365)
daily_ss_reserve = ss_cost_5yr / (5.0 * 365)
daily_drydock_reserve_total = daily_is_reserve + daily_ss_reserve

if drydock_mode == "剛完成塢特檢 (Fresh Out of SS/IS)":
    st.sidebar.success("💡 模式：前2.5年免進塢預提！第30個月單點支出 IS 中間檢，第31個月起開始預提 SS。")
else:
    st.sidebar.caption(f"💡 模式：每日固定平滑折合進塢費 ${daily_drydock_reserve_total:,.0f}/天。")

# --- 側邊欄：3. 商業條款與船齡 ---
st.sidebar.header("3. 商業條款與船齡 (Commercial)")

vessel_age_default = 15 if is_cement_mode else 0
vessel_age = st.sidebar.number_input(f"目前船齡 (Years)", min_value=0, max_value=max_age_limit, value=vessel_age_default, step=1)
default_gross_tc_default = 18000 if is_cement_mode else 16500
default_gross_tc_rate = st.sidebar.number_input("預設市場毛日租金 Gross TC Rate (USD/day)", value=default_gross_tc_default, step=500)
total_commission_pct = st.sidebar.number_input("租費總傭金 Brokerage Comm. (%)", value=3.75, step=0.25)
off_hire_days = st.sidebar.number_input("預估每年 Off-hire 停航天數", value=5, min_value=0, max_value=30)

# --- 側邊欄：4. 期末處分與 Book Value ---
st.sidebar.header("4. 期末處分與帳面價值 (Exit)")

exit_option = st.sidebar.radio("選擇期末處分方式", ["二手船市場價格 (Market Resale)", "拆船估算價格 (Scrap)"])
vessel_ldt_default = 6500 if is_cement_mode else 10000
vessel_ldt = st.sidebar.number_input("船舶輕排水量 LDT (公噸 MT)", value=vessel_ldt_default, step=500)
scrap_price_per_lt = st.sidebar.number_input("拆船單價 (USD/LT)", value=600, step=10)

long_tons = vessel_ldt / 1.016
base_scrap_steel_value = long_tons * scrap_price_per_lt

if is_cement_mode:
    equipment_residual_value = st.sidebar.number_input("水泥專用設備殘值 (USD)", value=3500000, step=250000)
    scrap_value = base_scrap_steel_value + equipment_residual_value
else:
    equipment_residual_value = 0.0
    scrap_value = base_scrap_steel_value

if exit_option == "二手船市場價格 (Market Resale)":
    market_resale_default = 16000000 if is_cement_mode else 22000000
    expected_resale_value = st.sidebar.number_input("預估期末二手船價 (USD)", value=market_resale_default, step=500000)
else:
    expected_resale_value = scrap_value

depreciation_useful_life = st.sidebar.number_input("會計折舊年限 Useful Life (Years)", value=30 if is_cement_mode else 25, step=1)
remaining_life_for_dep = max(1.0, float(depreciation_useful_life - vessel_age))
annual_depreciation = max(0.0, (purchase_price - scrap_value) / remaining_life_for_dep)

# --- 側邊欄：5. 折現率與運價 ---
st.sidebar.header("5. 資金成本折現率 (Hurdle Rate)")
hurdle_default = 7.0 if is_cement_mode else 8.0
hurdle_rate_pct = st.sidebar.number_input("船東目標折現率 WACC (%)", value=hurdle_default, step=0.5)
next_ss_years = st.sidebar.slider("距離下次特檢 (Special Survey) 年數", min_value=1, max_value=5, value=3 if drydock_mode=="常態均攤預提 (Flat Daily Reserve)" else 5)

eval_years = lease_years
use_uniform_rate = st.sidebar.checkbox("全期使用統一基準運價", value=True)

yearly_gross_tc_rates = []
if use_uniform_rate:
    yearly_gross_tc_rates = [float(default_gross_tc_rate)] * eval_years
else:
    with st.sidebar.expander("自訂每年不同毛日租金", expanded=True):
        for y in range(1, eval_years + 1):
            rate_input = st.number_input(f"第 {y} 年毛日租金 (USD/day)", value=float(default_gross_tc_rate), step=500.0, key=f"tc_rate_y_{y}")
            yearly_gross_tc_rates.append(rate_input)

# --- 後台財務與試算邏輯 ---
future_age_end_of_lease = vessel_age + eval_years
age_off_hire_penalty = (vessel_age - 15) * 1.5 if vessel_age > 15 else 0.0
dynamic_off_hire_days = off_hire_days + age_off_hire_penalty
annual_operating_days = max(300.0, 365.0 - dynamic_off_hire_days)
age_opex_factor = 1.0 + ((vessel_age - 15) * 0.03) if vessel_age > 15 else 1.0
dynamic_opex_base = opex_base * age_opex_factor

drydock_reserve_y1 = 0.0 if drydock_mode == "剛完成塢特檢 (Fresh Out of SS/IS)" else daily_drydock_reserve_total
daily_opex_y1 = dynamic_opex_base + daily_management_fee + drydock_reserve_y1

if ltv_pct > 0 and lease_years > 0:
    debt_amount = purchase_price * (ltv_pct / 100.0)
    facility_fee_amount = debt_amount * (facility_fee_pct / 100.0)
    total_interest_rate = (sofr_rate + margin_rate) / 100.0
    principal_to_pay = debt_amount - purchase_obligation
    daily_principal = principal_to_pay / (lease_years * 365)
    daily_interest_y1 = (debt_amount * total_interest_rate) / 365
    daily_bbc_y1 = daily_principal + daily_interest_y1
else:
    debt_amount, facility_fee_amount = 0.0, 0.0
    daily_principal, total_interest_rate = 0.0, 0.0
    daily_bbc_y1 = 0.0

working_capital_buffer = (daily_opex_y1 + daily_bbc_y1) * working_capital_days
equity_amount = (purchase_price - debt_amount) + facility_fee_amount + delivery_cost + working_capital_buffer

if equity_amount < 0:
    st.sidebar.caption(f"💡 融資試算：貸款 ${debt_amount:,.0f} / 套現 +${abs(equity_amount):,.0f}")
else:
    st.sidebar.caption(f"💡 融資試算：貸款 ${debt_amount:,.0f} / 自有資金 ${equity_amount:,.0f} (含營運資金 ${working_capital_buffer:,.0f})")

yearly_records = []
cum_cashflow = 0.0
payback_years_op = None
annual_net_cashflows = []

for y in range(1, eval_years + 1):
    current_vessel_age = vessel_age + (y - 1)
    opex_base_y = dynamic_opex_base * ((1.0 + opex_escalation_pct / 100.0) ** (y - 1))
    daily_management_fee_y = daily_management_fee * ((1.0 + opex_escalation_pct / 100.0) ** (y - 1))
    
    if drydock_mode == "剛完成塢特檢 (Fresh Out of SS/IS)":
        if y <= 2:
            drydock_reserve_y = 0.0
        elif y == 3:
            drydock_reserve_y = is_cost_25yr / 365.0
        else:
            drydock_reserve_y = daily_ss_reserve
    else:
        drydock_reserve_y = daily_drydock_reserve_total

    daily_total_opex_y = opex_base_y + daily_management_fee_y + drydock_reserve_y
    
    debt_start_of_year = max(0.0, debt_amount - (daily_principal * 365 * (y - 1)))
    daily_interest_y = (debt_start_of_year * total_interest_rate) / 365
    daily_bbc_y = daily_principal + daily_interest_y
    
    annual_cost_y = (daily_bbc_y + daily_total_opex_y) * 365.0
    gross_breakeven_y = (annual_cost_y / annual_operating_days) / (1.0 - total_commission_pct / 100.0)
    
    gross_tc_rate_y = yearly_gross_tc_rates[y - 1]
    net_tc_rate_y = gross_tc_rate_y * (1.0 - total_commission_pct / 100.0)
    
    annual_revenue_y = net_tc_rate_y * annual_operating_days
    annual_opex_y = daily_total_opex_y * 365.0
    annual_debt_service_y = daily_bbc_y * 365.0
    annual_cf_y = annual_revenue_y - annual_opex_y - annual_debt_service_y
    
    annual_net_cashflows.append(annual_cf_y)
    
    prev_cum = cum_cashflow
    cum_cashflow += annual_cf_y
    
    if equity_amount > 0 and prev_cum < equity_amount and cum_cashflow >= equity_amount:
        needed = equity_amount - prev_cum
        fraction = needed / annual_cf_y if annual_cf_y > 0 else 0
        payback_years_op = (y - 1) + fraction

    book_value_end_of_year = max(scrap_value, purchase_price - (annual_depreciation * y))

    yearly_records.append({
        "年份 (*1)": f"第 {y} 年 (船齡 {current_vessel_age}歲)",
        "市場毛運價 (*2)": gross_tc_rate_y,
        "未還本金 (*3)": debt_start_of_year,
        "當期 OPEX (含通膨/船管)": opex_base_y + daily_management_fee_y,
        "市場保本毛運價 (*4)": gross_breakeven_y,
        "每日淨利潤 (純利) (*5)": net_tc_rate_y - (annual_cost_y / annual_operating_days),
        "年度淨現金流 (*6)": annual_cf_y,
        "期末帳面價值 (Book Value) (*7)": book_value_end_of_year,
        "累積營運現金流 (*8)": cum_cashflow,
    })

df_yearly = pd.DataFrame(yearly_records)
end_of_eval_book_value = max(scrap_value, purchase_price - (annual_depreciation * eval_years))
accounting_gain_or_loss = expected_resale_value - end_of_eval_book_value

net_terminal_exit = (expected_resale_value - purchase_obligation) + working_capital_buffer

total_cash_recovered = cum_cashflow + net_terminal_exit
total_net_cash_profit = total_cash_recovered - equity_amount

project_cashflows = [-equity_amount] + annual_net_cashflows.copy()
project_cashflows[-1] += net_terminal_exit

r = hurdle_rate_pct / 100.0
npv_value = sum(cf / ((1 + r) ** t) for t, cf in enumerate(project_cashflows))

def compute_irr(cfs, iterations=1000):
    if equity_amount <= 0:
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

irr_value = compute_irr(project_cashflows)

total_months = eval_years * 12
op_days_per_month = annual_operating_days / 12.0
total_days_per_month = 365.0 / 12.0

monthly_records = []

monthly_records.append({
    "月份": "Day 0 (交船日)",
    "期初未還本金": debt_amount,
    "月度租金淨收入": 0.0,
    "本金支付": 0.0,
    "利息支付": 0.0,
    "月度總還本付息": 0.0,
    "月度總 OPEX": 0.0,
    "月度淨現金流": -equity_amount,
    "累積淨現金流": -equity_amount,
    "期末未還本金": debt_amount,
})

monthly_cum_cf = -equity_amount
monthly_principal_pay = (principal_to_pay / total_months) if (ltv_pct > 0 and lease_years > 0) else 0.0

for m in range(1, total_months + 1):
    current_year_idx = (m - 1) // 12
    current_gross_tc = yearly_gross_tc_rates[current_year_idx]
    current_net_tc = current_gross_tc * (1.0 - total_commission_pct / 100.0)
    current_opex_base = dynamic_opex_base * ((1.0 + opex_escalation_pct / 100.0) ** current_year_idx)
    current_management_fee = daily_management_fee * ((1.0 + opex_escalation_pct / 100.0) ** current_year_idx)

    monthly_is_capex = 0.0
    if drydock_mode == "剛完成塢特檢 (Fresh Out of SS/IS)":
        if m < 30:
            current_daily_drydock = 0.0
        elif m == 30:
            current_daily_drydock = 0.0
            monthly_is_capex = is_cost_25yr
        else:
            current_daily_drydock = daily_ss_reserve
    else:
        current_daily_drydock = daily_drydock_reserve_total

    current_daily_opex = current_opex_base + current_management_fee + current_daily_drydock

    start_debt = max(0.0, debt_amount - (monthly_principal_pay * (m - 1)))
    end_debt = max(0.0, start_debt - monthly_principal_pay)

    monthly_interest_pay = (start_debt * total_interest_rate) / 12.0
    monthly_revenue = current_net_tc * op_days_per_month
    monthly_opex = (current_daily_opex * total_days_per_month) + monthly_is_capex
    monthly_debt_service = monthly_principal_pay + monthly_interest_pay
    monthly_net_cf = monthly_revenue - monthly_debt_service - monthly_opex
    monthly_cum_cf += monthly_net_cf

    monthly_records.append({
        "月份": f"第 {m} 個月 (Y{(m-1)//12 + 1}-M{(m-1)%12 + 1})",
        "期初未還本金": start_debt,
        "月度租金淨收入": monthly_revenue,
        "本金支付": monthly_principal_pay,
        "利息支付": monthly_interest_pay,
        "月度總還本付息": monthly_debt_service,
        "月度總 OPEX": monthly_opex,
        "月度淨現金流": monthly_net_cf,
        "累積淨現金流": monthly_cum_cf,
        "期末未還本金": end_debt,
    })

df_monthly = pd.DataFrame(monthly_records)

exit_score = 0
reasons = []

if drydock_mode == "剛完成塢特檢 (Fresh Out of SS/IS)":
    reasons.append("剛完成特檢優勢：買船即享 2.5 年 (M1-M30) 免進塢營運期，前 30 個月無進塢費用預提，現金流極度充沛！")

if accounting_gain_or_loss >= 0:
    reasons.append(f"期末賣船會計資本利得 (Gain on Sale)：第 {eval_years} 年期末處分估值 ${expected_resale_value:,.0f} 高於會計帳面價值（Book Value ${end_of_eval_book_value:,.0f}），財報可認列 +${accounting_gain_or_loss:,.0f} 的處分利潤。")
else:
    reasons.append(f"期末賣船會計資產減損 (Loss on Sale)：期末處分估值 ${expected_resale_value:,.0f} 低於會計帳面價值（Book Value ${end_of_eval_book_value:,.0f}），財報需認列 -${abs(accounting_gain_or_loss):,.0f} 的資產減損。")

if npv_value > 0:
    reasons.append(f"財務效益優異 (NPV > 0)：在折現率 {hurdle_rate_pct}% 下，專案淨現值達 ${npv_value:,.0f}，IRR 高達 {irr_value:.2f}%，超過資金成本門檻。")
else:
    exit_score += 20
    reasons.append(f"未達目標回報率 (NPV < 0)：在折現率 {hurdle_rate_pct}% 下，NPV 為 ${npv_value:,.0f}。建議向賣方議價或調高融資 LTV。")

if working_capital_buffer > 0:
    reasons.append(f"Day 0 現金防禦充衛：已預留 {working_capital_days} 天營運資金（${working_capital_buffer:,.0f}），能有效化解交船初期租金支付位移（Collection Lag）之風險，且該資金將於期末全額回收。")

if is_cement_mode and equipment_residual_value > 0:
    reasons.append(f"水泥專用設備溢價保護：內含預估 ${equipment_residual_value:,.0f} 之氣動卸料設備殘值。總殘值保底達 ${scrap_value:,.0f}。")

if future_age_end_of_lease >= (42 if is_cement_mode else 28):
    exit_score += 50
    reasons.append(f"接近船齡極限（期末船齡 {future_age_end_of_lease} 歲），建議考慮執行拆船或清算處分。")

def generate_excel_with_formulas():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly Cashflow Schedule"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="4B5563")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Segoe UI", size=10)
    font_disclaimer = Font(name="Segoe UI", size=9, italic=True, color="64748B")

    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    border_thin = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))

    ws.cell(row=1, column=1, value=f"{vessel_name} ({vessel_profile}) - 月度現金流明細表").font = font_title
    ws.cell(row=2, column=1, value=f"IRR: {irr_value:.2f}% | NPV (@{hurdle_rate_pct}%): ${npv_value:,.0f} | 實質美金純利: ${total_net_cash_profit:,.0f} | 處分: {exit_option} (${expected_resale_value:,.0f})").font = font_subtitle
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18

    headers = ["月份", "期初未還本金 (B)", "月度租金淨收入 (C)", "本金支付 (D)", "利息支付 (E)", "月度總還本付息 (F)", "月度總 OPEX (G)", "月度淨現金流 (H)", "累積淨現金流 (I)", "期末未還本金 (J)"]
    ws.row_dimensions[4].height = 28
    for col_num, h_text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    monthly_interest_rate = total_interest_rate / 12.0
    op_monthly_records = [r for r in monthly_records if r["月份"] != "Day 0 (交船日)"]

    for idx, r in enumerate(op_monthly_records):
        row = idx + 5
        ws.row_dimensions[row].height = 20
        ws.cell(row=row, column=1, value=r["月份"]).alignment = align_center

        if row == 5:
            ws.cell(row=row, column=2, value=round(debt_amount))
        else:
            ws.cell(row=row, column=2, value=f"=J{row-1}")

        ws.cell(row=row, column=3, value=round(r["月度租金淨收入"]))
        ws.cell(row=row, column=4, value=round(r["本金支付"]))
        ws.cell(row=row, column=5, value=f"=ROUND(B{row}*{monthly_interest_rate}, 0)")
        ws.cell(row=row, column=6, value=f"=D{row}+E{row}")
        ws.cell(row=row, column=7, value=round(r["月度總 OPEX"]))
        ws.cell(row=row, column=8, value=f"=C{row}-F{row}-G{row}")

        if row == 5:
            ws.cell(row=row, column=9, value=f"=H{row}")
        else:
            ws.cell(row=row, column=9, value=f"=I{row-1}+H{row}")

        ws.cell(row=row, column=10, value=f"=MAX(0, B{row}-D{row})")

        for col_num in range(1, 11):
            cell = ws.cell(row=row, column=col_num)
            cell.font = font_body
            cell.border = border_thin
            if col_num > 1:
                cell.alignment = align_right
                cell.number_format = '$#,##0;($#,##0);"-"'

    last_row = len(op_monthly_records) + 6
    ws.cell(row=last_row, column=1, value="【免責聲明 Disclaimer】本報告僅供參考，實際結果可能有所不同，投資人須自行承擔所有相關風險。投資前請務必評估自身的風險承受能力。").font = font_disclaimer
    ws.cell(row=last_row + 1, column=1, value="This report is for reference only. Actual results may vary, and investors assume all associated risks. Please assess your personal risk tolerance prior to investing.").font = font_disclaimer

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def generate_pdf_report():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontName="MSung-Light", fontSize=16, leading=20, textColor=colors.HexColor("#1E3A8A"))
    subtitle_style = ParagraphStyle("SubTitleStyle", parent=styles["Normal"], fontName="MSung-Light", fontSize=9, leading=13, textColor=colors.HexColor("#4B5563"))
    heading_style = ParagraphStyle("HeadingStyle", parent=styles["Heading2"], fontName="MSung-Light", fontSize=11, leading=15, textColor=colors.HexColor("#1E3A8A"), spaceBefore=10, spaceAfter=4)
    body_style = ParagraphStyle("BodyStyle", parent=styles["Normal"], fontName="MSung-Light", fontSize=8.5, leading=11)
    disclaimer_style = ParagraphStyle("DisclaimerStyle", parent=styles["Normal"], fontName="MSung-Light", fontSize=7.5, leading=10, textColor=colors.HexColor("#64748B"))

    story.append(Paragraph(f"船舶 S&P 投資決策與 Exit 戰略分析報告", title_style))
    story.append(Paragraph(f"船舶/項目: {vessel_name} ({vessel_profile[:10]}) | 購入價: ${purchase_price:,.0f} | 融資 LTV: {ltv_pct}% | 營運資金預留: ${working_capital_buffer:,.0f} ({working_capital_days}天) | 期末 Book Value: ${end_of_eval_book_value:,.0f}", subtitle_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("1. 核心財務指標 (Executive Summary)", heading_style))

    first_year_gross_breakeven = df_yearly.loc[0, "市場保本毛運價 (*4)"]
    payback_str = "0年 (首日套現)" if equity_amount <= 0 else (f"{int(payback_years_op)}年{int(round((payback_years_op - int(payback_years_op)) * 12))}月" if payback_years_op is not None else f">{eval_years}年")
    irr_str = f"{irr_value:.2f}%" if irr_value is not None else "N/A"
    net_cash_profit_str = f"+${total_net_cash_profit:,.0f}" if total_net_cash_profit >= 0 else f"-${abs(total_net_cash_profit):,.0f}"

    kpi_data = [
        ["自有資金總投入 (Net Equity)", f"${equity_amount:,.0f}", "首年保本毛運價 (Breakeven)", f"${first_year_gross_breakeven:,.0f} /天"],
        ["自備款回本期 (Payback)", payback_str, "股權內部報酬率 (IRR)", irr_str],
        [f"專案淨現值 (NPV @{hurdle_rate_pct}%)", f"${npv_value:,.0f}", "專案全期實質美金純利", net_cash_profit_str],
    ]

    t_kpi = Table(kpi_data, colWidths=[130, 120, 160, 130])
    t_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1E3A8A")),
        ("FONTNAME", (0, 0), (-1, -1), "MSung-Light"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#1E3A8A")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 10))

    story.append(Paragraph("2. 逐年營運財務預測 (Yearly Financial Forecast)", heading_style))
    table_headers = ["年份", "市場毛運價", "未還本金", "當期 OPEX", "保本毛運價", "年度淨現金流", "累積現金流"]
    yearly_rows = [table_headers]

    for r in yearly_records:
        yearly_rows.append([
            str(r["年份 (*1)"]),
            f"${r['市場毛運價 (*2)']:,.0f}",
            f"${r['未還本金 (*3)']:,.0f}",
            f"${r['當期 OPEX (含通膨/船管)']:,.0f}",
            f"${r['市場保本毛運價 (*4)']:,.0f}",
            f"${r['年度淨現金流 (*6)']:,.0f}",
            f"${r['累積營運現金流 (*8)']:,.0f}",
        ])

    t_yearly = Table(yearly_rows, colWidths=[80, 65, 80, 75, 85, 80, 85])
    t_yearly.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "MSung-Light"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_yearly)
    story.append(Spacer(1, 10))

    story.append(Paragraph("3. S&P 資產處分與 Exit 戰略建議", heading_style))
    for reason in reasons:
        story.append(Paragraph(f"• {reason}", body_style))
        story.append(Spacer(1, 2))

    story.append(Spacer(1, 10))
    story.append(Paragraph("4. 船東資產獲利與營運戰略矩陣 (Shipowner Strategy Matrix)", heading_style))

    gain_loss_text = (
        f"期末處分預估 ${expected_resale_value:,.0f}，高於帳面價值 ${end_of_eval_book_value:,.0f}，財報可認列處分利潤 +${accounting_gain_or_loss:,.0f}。"
        if accounting_gain_or_loss >= 0
        else f"期末處分估值 ${expected_resale_value:,.0f} 低於會計帳面 ${end_of_eval_book_value:,.0f}（認列會計減損 -${abs(accounting_gain_or_loss):,.0f}）。但營運期間累積淨租金強勁，全期專案實質淨賺美金 {net_cash_profit_str}！"
    )
    leverage_text = f"融資成數 LTV 為 {ltv_pct}%（自備款 ${equity_amount:,.0f}，含 ${working_capital_buffer:,.0f} 營運資金）。" + ("超額融資套現模式。" if ltv_pct > 100 else f"高槓桿將股權 IRR 拉升至 {irr_str}。" if ltv_pct >= 60 else "低槓桿穩健架構。")
    
    safety_margin = default_gross_tc_rate - first_year_gross_breakeven
    breakeven_text = f"預設毛日租金 ${default_gross_tc_rate:,.0f}/天 vs 首年保本 ${first_year_gross_breakeven:,.0f}/天。" + (f"每日享 +${safety_margin:,.0f}/天 純利差。" if safety_margin > 0 else f"每日虧損 -${abs(safety_margin):,.0f}/天。")
    
    total_rate = sofr_rate + margin_rate
    interest_text = f"融資總利率 {total_rate:.2f}%，首年利息 ${debt_amount * (total_rate / 100.0):,.0f}。Margin 若調降 0.5%，每年省 ${debt_amount * 0.005:,.0f} 利息。"

    strat_table_data = [
        [Paragraph("<b>戰略維度</b>", body_style), Paragraph("<b>戰略評估與實務建議</b>", body_style)],
        [Paragraph("<b>1. 資產處分與會計損益</b>", body_style), Paragraph(gain_loss_text, body_style)],
        [Paragraph("<b>2. 負債槓桿與資本結構</b>", body_style), Paragraph(leverage_text, body_style)],
        [Paragraph("<b>3. 運價保本與安全邊際</b>", body_style), Paragraph(breakeven_text, body_style)],
        [Paragraph("<b>4. 利率與融資成本避險</b>", body_style), Paragraph(interest_text, body_style)],
    ]

    t_strat = Table(strat_table_data, colWidths=[150, 390])
    t_strat.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "MSung-Light"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t_strat)
    story.append(Spacer(1, 15))

    story.append(Paragraph("<b>【免責聲明 Disclaimer】</b>", disclaimer_style))
    story.append(Paragraph("本報告僅供參考，實際結果可能有所不同，投資人須自行承擔所有相關風險。投資前請務必評估自身的風險承受能力。", disclaimer_style))
    story.append(Paragraph("This report is for reference only. Actual results may vary, and investors assume all associated risks. Please assess your personal risk tolerance prior to investing.", disclaimer_style))

    doc.build(story)
    return buffer.getvalue()

title_col, btn_col1, btn_col2 = st.columns([3.2, 1, 1])

with title_col:
    st.markdown('<div class="main-title">🚢 船東 S&P 投報、IRR/NPV 與動態回本期決策計算器</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">專為 S&P 經紀人與船東設計｜極簡高質感海事金融決策終端</div>', unsafe_allow_html=True)

excel_bytes = generate_excel_with_formulas()
pdf_bytes = generate_pdf_report()

with btn_col1:
    st.download_button(
        label="📊 下載現金流分析報告 (.xlsx)",
        data=excel_bytes,
        file_name=f"{vessel_name}_Cashflow_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with btn_col2:
    st.download_button(
        label="📄 下載投資分析報告 (.pdf)",
        data=pdf_bytes,
        file_name=f"{vessel_name}_Investment_Report.pdf",
        mime="application/pdf",
    )

st.markdown("---")

def render_kpi_card(col, title, value, subtext="", status="neutral"):
    if status == "loss":
        bg_color = "#FEE2E2"
        border_color = "#FCA5A5"
        title_color = "#991B1B"
        value_color = "#991B1B"
        sub_color = "#B91C1C"
    elif status == "profit":
        bg_color = "#ECFDF5"
        border_color = "#A7F3D0"
        title_color = "#065F46"
        value_color = "#047857"
        sub_color = "#065F46"
    else:
        bg_color = "#F8FAFC"
        border_color = "#E2E8F0"
        title_color = "#4B5563"
        value_color = "#1E3A8A"
        sub_color = "#64748B"

    card_html = f"""
    <div style="
        background-color: {bg_color};
        border: 1px solid {border_color};
        padding: 10px 12px;
        border-radius: 6px;
        height: 100px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-sizing: border-box;
    ">
        <div style="font-size: 0.78rem; color: {title_color}; font-weight: 600; white-space: normal; line-height: 1.2;">{title}</div>
        <div style="font-size: 1.2rem; color: {value_color}; font-weight: 700; line-height: 1.2;">{value}</div>
        <div style="font-size: 0.72rem; color: {sub_color}; white-space: normal; line-height: 1.2;">{subtext}</div>
    </div>
    """
    col.markdown(card_html, unsafe_allow_html=True)

col1, col2, col3, col4, col5, col6 = st.columns(6)

equity_sub = "含手續費/交接費/營運資金" if (facility_fee_amount > 0 or delivery_cost > 0 or working_capital_buffer > 0) else ("全額現金" if ltv_pct == 0 else f"融資 {ltv_pct}%")
render_kpi_card(col1, "自有資金總投入", f"${equity_amount:,.0f}" if equity_amount >= 0 else f"-${abs(equity_amount):,.0f}", equity_sub, status="neutral")

first_year_gross_breakeven = df_yearly.loc[0, "市場保本毛運價 (*4)"]
render_kpi_card(col2, "首年保本毛運價", f"${first_year_gross_breakeven:,.0f}/天", "運價保本門檻", status="neutral")

payback_display = "0 年 (首日套現)" if equity_amount <= 0 else (f"{int(payback_years_op)}年{int(round((payback_years_op - int(payback_years_op)) * 12))}月" if payback_years_op is not None else f">{eval_years} 年")
render_kpi_card(col3, "自備款回本期", payback_display, "營運資金回收", status="neutral")

irr_disp = f"{irr_value:.2f}%" if irr_value is not None else "N/A"
render_kpi_card(col4, "股權內部報酬率", irr_disp, "槓桿後 IRR", status="neutral")

if total_net_cash_profit >= 0:
    render_kpi_card(col5, "實質美金純利 (Cash Profit)", f"+${total_net_cash_profit:,.0f}", "真金白銀純利", status="profit")
else:
    render_kpi_card(col5, "實質美金純利 (Cash Loss)", f"-${abs(total_net_cash_profit):,.0f}", "⚠️ 實質營運虧損", status="loss")

if accounting_gain_or_loss >= 0:
    render_kpi_card(col6, "會計帳面損益 (Book Gain)", f"+${accounting_gain_or_loss:,.0f}", f"期末帳面 ${end_of_eval_book_value:,.0f}", status="profit")
else:
    render_kpi_card(col6, "會計帳面損益 (Book Loss)", f"-${abs(accounting_gain_or_loss):,.0f}", f"期末帳面 ${end_of_eval_book_value:,.0f}", status="loss")

st.markdown("<br/>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "各年度動態財務與公式說明",
    "累積現金流與回本交點圖",
    "月度詳細現金流與還本付息表",
    "AI 賣船與處分建議",
    "船東戰略與獲利策略矩陣",
])

with tab1:
    st.subheader(f"逐年動態財務預測表 ({vessel_profile})")

    df_display = df_yearly.copy()
    for col in [
        "市場毛運價 (*2)",
        "未還本金 (*3)",
        "當期 OPEX (含通膨/船管)",
        "市場保本毛運價 (*4)",
        "每日淨利潤 (純利) (*5)",
        "年度淨現金流 (*6)",
        "期末帳面價值 (Book Value) (*7)",
        "累積營運現金流 (*8)",
    ]:
        df_display[col] = df_display[col].apply(lambda x: f"${x:,.0f}")

    st.table(df_display[[
        "年份 (*1)",
        "市場毛運價 (*2)",
        "未還本金 (*3)",
        "當期 OPEX (含通膨/船管)",
        "市場保本毛運價 (*4)",
        "每日淨利潤 (純利) (*5)",
        "年度淨現金流 (*6)",
        "期末帳面價值 (Book Value) (*7)",
        "累積營運現金流 (*8)",
    ]])

    st.markdown("##### 📝 表格計算公式說明（Formulas & Footnotes）")
    st.markdown(f"""
*   **`*1` 年份 (y)**：評估期之第 y 個營運年度，括號內標註當期船齡。
*   **`*2` 市場毛運價 (Gross TC)**：市場預估開出或水泥 COA 長約折算之每日毛租金。
*   **`*3` 未還本金 (Debt)**：第 y 年年初剩餘之融資未還本金餘額。
*   **`*4` 市場保本毛運價 (Gross Breakeven)**：考慮傭金 ({total_commission_pct}%) 與老化校正後之 Off-hire ({dynamic_off_hire_days:.1f} 天) 後，當年在市場上維繫不虧損所需的最低開價門檻。
""")

    st.latex(
        r"""GrossBreakeven_y = \frac{(\text{Daily BBC}_y + \text{Daily OPEX}_y) \times 365}{(365 - \text{DynamicOffHire}) \times (1 - \text{Comm \%})}"""
    )

    st.markdown(f"""
*   **`*5` 每日淨利潤 (純利)**：實際扣除傭金抽成、Off-hire 停航天數折算、OPEX (含通膨、委外船管費與老化加成) 及每日融資還本付息 (BBC) 後，船舶每天為船東創造的實拿純淨利。
*   **`*6` 年度淨現金流 (Annual CF)**：當年度營運實際進帳之純營運現金流淨額。
*   **`*7` 期末帳面價值 (Book Value)**：以會計直線折舊法算至該年年底之船舶資產剩餘帳面價值（購船成本扣除累積折舊）。
*   **`*8` 累積營運現金流 (Cum CF)**：截至第 y 年年底止，累積進帳之營運純現金流總額。
""")

with tab2:
    st.subheader("累積營運現金流 vs 自有資金回收曲線")

    years_list = [0] + list(range(1, eval_years + 1))
    cum_cf_list = [0.0] + list(df_yearly["累積營運現金流 (*8)"])

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=years_list,
            y=cum_cf_list,
            mode="lines+markers",
            name="累積營運淨現金流 (USD)",
            line=dict(color="#1E3A8A", width=3),
        )
    )

    if equity_amount > 0:
        fig.add_hline(
            y=equity_amount,
            line_dash="dash",
            annotation_text=f"自備款門檻 (${equity_amount:,.0f})",
            line_color="#DC2626",
        )

    fig.update_layout(
        title="累積現金流回收趨勢圖",
        xaxis_title="年份 (Years)",
        yaxis_title="美金 (USD)",
        template="plotly_white",
    )
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("月度營運現金流、融資還本付息與資產殘值明細")
    
    x_labels = [r["月份"] for r in monthly_records]
    inflows = [r["月度租金淨收入"] for r in monthly_records]
    outflows = [r["月度總還本付息"] + r["月度總 OPEX"] for r in monthly_records]
    cum_cfs = [r["累積淨現金流"] for r in monthly_records]

    fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])

    fig_monthly.add_trace(
        go.Bar(
            x=x_labels,
            y=inflows,
            name="月度租金淨收入 (Inflow)",
            marker_color="#1E3A8A",
            opacity=0.8,
        ),
        secondary_y=False,
    )
    fig_monthly.add_trace(
        go.Bar(
            x=x_labels,
            y=outflows,
            name="月度總開支 (Outflow)",
            marker_color="#94A3B8",
            opacity=0.8,
        ),
        secondary_y=False,
    )
    fig_monthly.add_trace(
        go.Scatter(
            x=x_labels,
            y=cum_cfs,
            name="累積淨現金流 (Cum. Cashflow)",
            line=dict(color="#059669", width=2.5),
        ),
        secondary_y=True,
    )

    fig_monthly.update_layout(
        title="月度現金流走勢圖 (含 Day 0 自備款總投入與第 30 個月特檢)",
        barmode="group",
        xaxis_title="營運月份 (Months)",
        template="plotly_white",
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

    df_m_display = df_monthly.copy()
    for col in [
        "期初未還本金",
        "月度租金淨收入",
        "本金支付",
        "利息支付",
        "月度總還本付息",
        "月度總 OPEX",
        "月度淨現金流",
        "累積淨現金流",
        "期末未還本金",
    ]:
        df_m_display[col] = df_m_display[col].apply(lambda x: f"${x:,.0f}")

    st.dataframe(df_m_display, use_container_width=True, height=380)

with tab4:
    st.subheader("AI 船東資產處分、IRR/NPV 與 Exit 評估建議")

    if (accounting_gain_or_loss < 0) and (total_net_cash_profit > 0):
        st.info(
            f"💡 **【機構級投資解讀：會計紙上減損，實質現金盈利】**\n\n"
            f"本專案第 {eval_years} 年賣船估值低於會計帳面價值（認列 **-${abs(accounting_gain_or_loss):,.0f}** 之紙上處分減損）。\n\n"
            f"**然而從股權投資角度看：** 船舶營運期間累積產出 **+${cum_cashflow:,.0f}** 之強勁淨租金。扣除初始自備款與貸款尾款後，全期專案共為船東創造 **+${total_net_cash_profit:,.0f} 美金的實質純現金利潤**，槓桿後股權 IRR 達 **{irr_disp}**，屬於非常健康的現金流獲利項目！"
        )

    st.write(f"##### 🎯 處分/賣船建議指數：**{min(exit_score, 100)} / 100**")
    st.progress(min(exit_score, 100) / 100)

    st.markdown("##### 🔍 S&P 經紀人資產分析報告：")
    for r in reasons:
        st.markdown(f"• {r}")

with tab5:
    st.subheader("船東（Shipowner）資產獲利與營運戰略矩陣")
    
    st.markdown("##### 📊 專案現金流勾稽與會計損益還原對照表 (Reconciliation Table)")
    
    recon_data = [
        {"財務維度項目": "1. Day 0 初始自有資金總投入 (Net Equity)", "實質美金現金流 (Cash Flow)": f"-${equity_amount:,.0f}", "財報 P&L 損益影響": "—", "實務商業說明": "含手續費、交接費與營運資金預留"},
        {"財務維度項目": f"2. 持有 {eval_years} 年累積營運淨租金收入", "實質美金現金流 (Cash Flow)": f"+${cum_cashflow:,.0f}", "財報 P&L 損益影響": f"+${cum_cashflow:,.0f}", "實務商業說明": "每日租金扣除 OPEX、船管費與還本付息"},
        {"財務維度項目": "3. 期末賣船處分回收金額", "實質美金現金流 (Cash Flow)": f"+${expected_resale_value:,.0f}", "財報 P&L 損益影響": f"+${expected_resale_value:,.0f}", "實務商業說明": "二手市場轉售或拆船回收現金"},
        {"財務維度項目": "4. 清償銀行貸款尾款 (Balloon / Purchase Option)", "實質美金現金流 (Cash Flow)": f"-${purchase_obligation:,.0f}", "財報 P&L 損益影響": "—", "實務商業說明": "結清銀行尾款或支付租賃公司買回過戶金額"},
        {"財務維度項目": "5. 預留營運資金 (Working Capital) 全額回收", "實質美金現金流 (Cash Flow)": f"+${working_capital_buffer:,.0f}", "財報 P&L 損益影響": "—", "實務商業說明": "Day 0 預留之墊款原封不動解凍收回"},
        {"財務維度項目": f"6. 扣除：期末會計帳面價值 (Book Value)", "實質美金現金流 (Cash Flow)": "—", "財報 P&L 損益影響": f"-${end_of_eval_book_value:,.0f}", "實務商業說明": "會計直線折舊算出之紙上剩餘價值"},
        {"財務維度項目": "🎯 最終結算總獲利 (Final Result)", "實質美金現金流 (Cash Flow)": f"+${total_net_cash_profit:,.0f}" if total_net_cash_profit>=0 else f"-${abs(total_net_cash_profit):,.0f}", "財報 P&L 損益影響": f"+${accounting_gain_or_loss:,.0f}" if accounting_gain_or_loss>=0 else f"-${abs(accounting_gain_or_loss):,.0f}", "實務商業說明": "實質真金白銀純利 vs. 財報賣船損益"},
    ]
    st.table(pd.DataFrame(recon_data))

    st.markdown("---")

    strat_col1, strat_col2 = st.columns(2)

    with strat_col1:
        st.markdown("##### 1. 📈 資產買賣與會計處分損益")
        st.write(f"**預估處分價**：${expected_resale_value:,.0f} | **期末 Book Value**：${end_of_eval_book_value:,.0f}")
        if accounting_gain_or_loss >= 0:
            st.success(f"💎 會計資本利得：處分價格高於帳面價值 **+${accounting_gain_or_loss:,.0f}**！")
        else:
            st.warning(f"📉 會計資產減損：處分價格低於帳面價值 **-${abs(accounting_gain_or_loss):,.0f}**（全期美金純利 **+${total_net_cash_profit:,.0f}**）")

    with strat_col2:
        st.markdown("##### 2. 🏦 負債槓桿與資本結構優化")
        st.write(f"**目前融資成數 (LTV)**：{ltv_pct}% | **初始投入淨自備款**：${equity_amount:,.0f}")
        if ltv_pct >= 60:
            st.info(f"⚖️ 高效財務槓桿：配合 {ltv_pct}% LTV，將股權 IRR 拉升至 **{irr_disp}**（目標折現率 {hurdle_rate_pct}%）。")
        else:
            st.success("🛡️ 低槓桿/保守型架構：自備款比例高，財務結構極度穩健。")

    st.markdown("---")
    strat_col3, strat_col4 = st.columns(2)

    with strat_col3:
        st.markdown("##### 3. 🛡️ 運價保本點與安全邊際")
        safety_margin = default_gross_tc_rate - first_year_gross_breakeven
        st.write(f"**預設毛日租金**：${default_gross_tc_rate:,.0f}/天 | **首年保本毛運價**：${first_year_gross_breakeven:,.0f}/天")
        if safety_margin > 0:
            st.success(f"🟢 安全邊際：每日純利差 **+${safety_margin:,.0f} /天**")
        else:
            st.error(f"🔴 營運虧損警示：每日虧損利差 **-${abs(safety_margin):,.0f} /天**")

    with strat_col4:
        st.markdown("##### 4. 💵 利率與融資成本避險")
        total_rate = sofr_rate + margin_rate
        st.write(f"**融資總利率**：{total_rate:.2f}% (SOFR {sofr_rate}% + Margin {margin_rate}%)")
        st.info(f"💡 利率敏感度：若 Margin 談判調降 **0.5%**，每年可直接省下 **${debt_amount * 0.005:,.0f}** 利息支出！")
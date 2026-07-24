import os
from io import BytesIO
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.figure_factory as ff
import plotly.express as px

st.set_page_config(page_title="跨境库存ERP计划系统", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"


# ===================== 文件配置 =====================
SKU_FILE = DATA_DIR / "sku.xlsx"
SUPPLY_FILE = DATA_DIR / "supply_chain.xlsx"
DEFAULT_SUPPLY_IMPORT = DATA_DIR / "shipping_batches.xlsx"

SKU_COLUMNS = [
    "SKU编码", "产品名称", "采购交期(天)", "安全库存天数", "日均销量",
    "深圳仓库存", "海外仓存", "Full仓库存",
    "国内发海外仓在途", "送Full仓在途", "采购在途", "计划返单数量"
]

SUPPLY_COLUMNS = [
    "中文品名", "公司SKU", "MSKU", "备注", "FULL仓", "入仓号/留仓号",
    "Shipment ID", "是否入仓", "海外仓数量", "剩余数量", "物流商", "渠道",
    "交货日期", "实际送达日期", "时效", "预估送达日期"
]

STATUS_ORDER = [
    "采购途中", "海运途中", "入full途中", "海外仓",
    "已入仓", "转运营", "丢失赔付", "其他"
]

DATE_COLUMNS = ["交货日期", "实际送达日期", "预估送达日期"]
NUMERIC_SUPPLY_COLUMNS = ["海外仓数量", "剩余数量", "时效"]


# ===================== 通用工具 =====================
def safe_num(value, default=0.0):
    value = pd.to_numeric(value, errors="coerce")
    return default if pd.isna(value) else float(value)


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_status(value):
    text = clean_text(value)
    lower = text.lower().replace(" ", "")

    mapping = {
        "采购途中": "采购途中",
        "采购中": "采购途中",
        "海运途中": "海运途中",
        "海运中": "海运途中",
        "入full途中": "入full途中",
        "入full仓途中": "入full途中",
        "送full仓在途": "入full途中",
        "海外仓": "海外仓",
        "已入仓": "已入仓",
        "转运营": "转运营",
        "丢失赔付": "丢失赔付",
    }
    if lower in mapping:
        return mapping[lower]
    if "采购" in lower:
        return "采购途中"
    if "海运" in lower:
        return "海运途中"
    if "full" in lower and ("途中" in lower or "在途" in lower):
        return "入full途中"
    if "海外仓" in lower:
        return "海外仓"
    if "入仓" in lower:
        return "已入仓"
    if "赔付" in lower or "丢失" in lower:
        return "丢失赔付"
    if "运营" in lower:
        return "转运营"
    return "其他"


def batch_quantity(row):
    """批次数量优先取剩余数量，缺失时取海外仓数量。"""
    remaining = safe_num(row.get("剩余数量"))
    overseas = safe_num(row.get("海外仓数量"))
    return remaining if remaining > 0 else overseas


def to_excel_bytes(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


# ===================== SKU数据 =====================
def init_sku_table():
    try:
        df = pd.read_excel(SKU_FILE)
    except Exception:
        df = pd.DataFrame(columns=SKU_COLUMNS)
        df.to_excel(SKU_FILE, index=False)

    for col in SKU_COLUMNS:
        if col not in df.columns:
            df[col] = "" if col in ["SKU编码", "产品名称"] else 0

    df = df[SKU_COLUMNS].copy()
    df["SKU编码"] = df["SKU编码"].fillna("").astype(str).str.strip()
    df["产品名称"] = df["产品名称"].fillna("").astype(str).str.strip()

    for col in SKU_COLUMNS[2:]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def save_sku_df(df):
    df[SKU_COLUMNS].to_excel(SKU_FILE, index=False)


def batch_import_sku(upload_file):
    new_data = pd.read_excel(upload_file) if upload_file.name.lower().endswith(".xlsx") else pd.read_csv(upload_file)
    for col in SKU_COLUMNS:
        if col not in new_data.columns:
            new_data[col] = "" if col in ["SKU编码", "产品名称"] else 0

    new_data = new_data[SKU_COLUMNS].copy()
    new_data["SKU编码"] = new_data["SKU编码"].fillna("").astype(str).str.strip()
    new_data = new_data[new_data["SKU编码"] != ""]

    old_df = init_sku_table()
    merged = pd.concat([old_df, new_data], ignore_index=True)
    merged = merged.drop_duplicates(subset="SKU编码", keep="last")
    save_sku_df(merged)
    return merged


# ===================== 供应链批次数据 =====================
def prepare_supply_df(df):
    df = df.copy()

    # 只保留真实字段，忽略Excel空白尾列
    for col in SUPPLY_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NaT if col in DATE_COLUMNS else ""

    df = df[SUPPLY_COLUMNS].copy()

    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in NUMERIC_SUPPLY_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    text_cols = [c for c in SUPPLY_COLUMNS if c not in DATE_COLUMNS + NUMERIC_SUPPLY_COLUMNS]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["是否入仓"] = df["是否入仓"].apply(normalize_status)
    df = df[df["公司SKU"] != ""].copy()
    return df


def init_supply_table():
    if os.path.exists(SUPPLY_FILE):
        try:
            return prepare_supply_df(pd.read_excel(SUPPLY_FILE))
        except Exception:
            pass

    # 首次运行时，自动读取放在项目目录中的原始物流表
    if os.path.exists(DEFAULT_SUPPLY_IMPORT):
        try:
            df = prepare_supply_df(pd.read_excel(DEFAULT_SUPPLY_IMPORT))
            save_supply_df(df)
            return df
        except Exception:
            pass

    df = pd.DataFrame(columns=SUPPLY_COLUMNS)
    df.to_excel(SUPPLY_FILE, index=False)
    return prepare_supply_df(df)


def supply_unique_key(df):
    key_parts = [
        df["Shipment ID"].fillna("").astype(str).str.strip(),
        df["公司SKU"].fillna("").astype(str).str.strip(),
        df["入仓号/留仓号"].fillna("").astype(str).str.strip(),
        df["MSKU"].fillna("").astype(str).str.strip(),
    ]
    key = key_parts[0]
    for part in key_parts[1:]:
        key = key + "|" + part
    return key


def save_supply_df(df):
    output = df[SUPPLY_COLUMNS].copy()
    for col in DATE_COLUMNS:
        output[col] = pd.to_datetime(output[col], errors="coerce")
    output.to_excel(SUPPLY_FILE, index=False)


def batch_import_supply(upload_file):
    incoming = pd.read_excel(upload_file)
    incoming = prepare_supply_df(incoming)
    old = init_supply_table()

    incoming["_key"] = supply_unique_key(incoming)
    old["_key"] = supply_unique_key(old)

    merged = pd.concat([old, incoming], ignore_index=True)
    merged = merged.drop_duplicates(subset="_key", keep="last").drop(columns="_key")
    save_supply_df(merged)

    # 自动把供应链表里新出现的SKU补充到SKU主数据
    sku_df = init_sku_table()
    existing = set(sku_df["SKU编码"].astype(str))
    new_skus = []
    for _, row in incoming.iterrows():
        code = clean_text(row["公司SKU"])
        if code and code not in existing:
            new_skus.append({
                "SKU编码": code,
                "产品名称": clean_text(row["中文品名"]),
                "采购交期(天)": 20,
                "安全库存天数": 15,
                "日均销量": 0,
                "深圳仓库存": 0,
                "海外仓存": 0,
                "Full仓库存": 0,
                "国内发海外仓在途": 0,
                "送Full仓在途": 0,
                "采购在途": 0,
                "计划返单数量": 0,
            })
            existing.add(code)

    if new_skus:
        sku_df = pd.concat([sku_df, pd.DataFrame(new_skus)], ignore_index=True)
        save_sku_df(sku_df)

    return merged, len(incoming)


def summarize_supply_by_sku(supply_df):
    """把统一批次表按状态汇总成库存计划所需字段。"""
    if supply_df.empty:
        return pd.DataFrame(columns=[
            "SKU编码", "批次采购途中", "批次海运途中", "批次送Full途中",
            "批次海外仓库存", "批次已入仓", "批次丢失赔付"
        ])

    work = supply_df.copy()
    work["批次数量"] = work.apply(batch_quantity, axis=1)

    pivot = (
        work.pivot_table(
            index="公司SKU",
            columns="是否入仓",
            values="批次数量",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
        .rename(columns={"公司SKU": "SKU编码"})
    )

    result = pd.DataFrame({"SKU编码": pivot["SKU编码"]})
    result["批次采购途中"] = pivot["采购途中"] if "采购途中" in pivot.columns else 0
    result["批次海运途中"] = pivot["海运途中"] if "海运途中" in pivot.columns else 0
    result["批次送Full途中"] = pivot["入full途中"] if "入full途中" in pivot.columns else 0
    result["批次海外仓库存"] = pivot["海外仓"] if "海外仓" in pivot.columns else 0
    result["批次已入仓"] = pivot["已入仓"] if "已入仓" in pivot.columns else 0
    result["批次丢失赔付"] = pivot["丢失赔付"] if "丢失赔付" in pivot.columns else 0
    return result



def expected_available_date_for_batch(row, status, base_today, purchase_days, ship_days, customs_days):
    """
    兼容旧接口，使用新的默认日期推算逻辑。
    注意：
    - 海运途中无开船日期：默认7天后开船，再43天到达；
    - 采购途中无下单日期：默认7天完成采购，再50天到达；
    - 1900年日期统一视为无效空值。
    """
    return infer_batch_available_date(
        row=row,
        status=status,
        base_today=base_today,
        default_ship_wait_days=7,
        ocean_days_without_sailing=43,
        purchase_ready_days=7,
        purchase_to_arrival_days=50,
        customs_days=customs_days,
    )


def calculate_reorder_connection_plan(
    selected_sku,
    supply_df,
    current_available,
    daily_sales,
    base_today,
    purchase_days,
    ship_days,
    customs_days,
    overlap_days=10,
    platform_stock_days=20,
    reorder_turnover_days=30,
):
    """
    接货式返单：
    优先衔接最新采购在途；没有采购在途时衔接最新海运在途；
    新返单在上一批销售结束前 overlap_days 天到货。
    """
    lead_days = purchase_days + ship_days + customs_days
    sku_batches = supply_df[
        supply_df["公司SKU"].fillna("").astype(str).str.strip()
        == str(selected_sku).strip()
    ].copy()

    source_label = "当前现货"
    source_qty = float(current_available)
    source_available = base_today

    for status in ["采购途中", "海运途中"]:
        candidates = sku_batches[sku_batches["是否入仓"] == status].copy()
        if candidates.empty:
            continue

        candidates["预计可售日期_计算"] = candidates.apply(
            lambda r: expected_available_date_for_batch(
                r, status, base_today, purchase_days, ship_days, customs_days
            ),
            axis=1,
        )
        candidates["批次数量_计算"] = candidates.apply(batch_quantity, axis=1)
        candidates = candidates.sort_values("预计可售日期_计算", kind="stable")
        latest = candidates.iloc[-1]

        source_label = status
        source_qty = float(latest["批次数量_计算"])
        source_available = latest["预计可售日期_计算"]
        break

    if daily_sales > 0:
        sales_days = source_qty / daily_sales
        sales_end = source_available + timedelta(days=sales_days)
        planned_arrival = sales_end - timedelta(days=overlap_days)
        reorder_date = planned_arrival - timedelta(days=lead_days)
    else:
        sales_days = 0
        sales_end = None
        planned_arrival = None
        reorder_date = None

    return {
        "衔接来源": source_label,
        "衔接库存数量": source_qty,
        "衔接库存预计可售日期": source_available,
        "衔接库存可销售天数": sales_days,
        "衔接库存预计销售结束日期": sales_end,
        "新返单计划到货日期": planned_arrival,
        "新返单最晚下单日期": reorder_date,
        "安全重叠天数": overlap_days,
        "完整前置时间": lead_days,
        "平台仓在库天数": platform_stock_days,
        "返单周转天数": reorder_turnover_days,
        "平台仓目标库存": daily_sales * platform_stock_days if daily_sales > 0 else 0,
        "返单周转目标库存": daily_sales * reorder_turnover_days if daily_sales > 0 else 0,
    }



def valid_business_date(value):
    """
    将Excel日期转换为有效业务日期。
    1900年前后通常是Excel空值/0被误解析，统一视为未填写。
    """
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return None
    if dt.year < 2000:
        return None
    return dt.date()


MAX_TIMELINE_DATE = date(2100, 12, 31)
MAX_TIMELINE_DAYS = 36500


def safe_future_timestamp(start_value, days_value):
    """安全计算甘特图结束日期，防止极小日销量导致日期溢出或页面黑屏。"""
    start_ts = pd.Timestamp(start_value)
    try:
        days_float = float(days_value)
    except (TypeError, ValueError):
        days_float = 0.0

    if pd.isna(days_float) or days_float < 0:
        days_float = 0.0

    days_float = min(days_float, float(MAX_TIMELINE_DAYS))
    max_ts = pd.Timestamp(MAX_TIMELINE_DATE)
    result = start_ts + pd.Timedelta(days=days_float)
    return min(result, max_ts)


def infer_batch_available_date(
    row,
    status,
    base_today,
    default_ship_wait_days=7,
    ocean_days_without_sailing=43,
    purchase_ready_days=7,
    purchase_to_arrival_days=50,
    customs_days=0,
):
    """
    批次预计可售日期默认规则：

    海运途中：
    - 有有效开船/交货日期：该日期 + 43天到达
    - 没有日期：今天 + 7天开船 + 43天到达

    采购途中：
    - 有有效下单/交货日期：该日期 + 7天采购完成 + 50天到达
    - 没有日期：今天 + 7天采购完成 + 50天到达

    入FULL途中：
    - 有预估/实际送达日期优先使用
    - 否则今天 + 清关预约天数
    """
    estimated = valid_business_date(row.get("预估送达日期"))
    actual = valid_business_date(row.get("实际送达日期"))
    handover = valid_business_date(row.get("交货日期"))

    if actual:
        return actual
    if estimated:
        return estimated

    if status == "海运途中":
        sailing_date = handover or (base_today + timedelta(days=default_ship_wait_days))
        return sailing_date + timedelta(days=ocean_days_without_sailing)

    if status == "采购途中":
        order_date = handover or base_today
        return order_date + timedelta(days=purchase_ready_days + purchase_to_arrival_days)

    if status == "入full途中":
        return base_today + timedelta(days=customs_days)

    if status in ["海外仓", "已入仓"]:
        return base_today

    return base_today


# ===================== 库存计算 =====================
def calc_sell_out_day(avail_stock, daily_sale, today):
    if daily_sale <= 0 or avail_stock <= 0:
        return None
    return today + timedelta(days=float(avail_stock) / float(daily_sale))


def calc_reorder_deadline(sell_out_date, purchase_d, ship_d, customs_d, watch_d, today):
    if sell_out_date is None:
        return today
    return sell_out_date - timedelta(days=float(purchase_d + ship_d + customs_d + watch_d))


def build_inventory_calc(
    sku_df, supply_df, base_today,
    global_purchase, global_ship, global_customs,
    global_safety_days, platform_stock_days, reorder_turnover_days
):
    batch_summary = summarize_supply_by_sku(supply_df)
    merged = sku_df.merge(batch_summary, on="SKU编码", how="left")

    batch_cols = [
        "批次采购途中", "批次海运途中", "批次送Full途中",
        "批次海外仓库存", "批次已入仓", "批次丢失赔付"
    ]
    for col in batch_cols:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)

    rows = []
    for _, row in merged.iterrows():
        daily_sale = safe_num(row["日均销量"])
        safe_days = max(safe_num(row["安全库存天数"], 15), float(global_safety_days))

        # 海外仓批次数据存在时使用批次汇总；没有时沿用SKU表手工库存
        overseas_stock = (
            safe_num(row["批次海外仓库存"])
            if safe_num(row["批次海外仓库存"]) > 0
            else safe_num(row["海外仓存"])
        )

        purchase_transit = (
            safe_num(row["批次采购途中"])
            if safe_num(row["批次采购途中"]) > 0
            else safe_num(row["采购在途"])
        )

        ocean_transit = (
            safe_num(row["批次海运途中"])
            if safe_num(row["批次海运途中"]) > 0
            else safe_num(row["国内发海外仓在途"])
        )

        full_transit = (
            safe_num(row["批次送Full途中"])
            if safe_num(row["批次送Full途中"]) > 0
            else safe_num(row["送Full仓在途"])
        )

        available = (
            safe_num(row["深圳仓库存"])
            + overseas_stock
            + safe_num(row["Full仓库存"])
            + safe_num(row["批次已入仓"])
        )
        total_stock = available + purchase_transit + ocean_transit + full_transit

        sell_day = calc_sell_out_day(available, daily_sale, base_today)
        reorder_deadline = calc_reorder_deadline(
            sell_day, global_purchase, global_ship, global_customs, 0, base_today
        )
        remain_days = (reorder_deadline - base_today).days
        turnover_days = total_stock / daily_sale if daily_sale > 0 else None
        platform_target_stock = daily_sale * platform_stock_days
        reorder_cycle_target_stock = daily_sale * reorder_turnover_days
        safety_stock_qty = daily_sale * global_safety_days
        suggested_reorder_qty = max(
            0,
            platform_target_stock + reorder_cycle_target_stock + safety_stock_qty - total_stock
        )

        new_row = row.copy()
        new_row["海外仓库存(批次合并)"] = overseas_stock
        new_row["采购途中(批次合并)"] = purchase_transit
        new_row["海运途中(批次合并)"] = ocean_transit
        new_row["入Full途中(批次合并)"] = full_transit
        new_row["可售现货库存"] = available
        new_row["全链路总库存(含在途)"] = total_stock
        new_row["预计售罄日期"] = sell_day
        new_row["返单最晚截止日期"] = reorder_deadline
        new_row["距离必须返单剩余天数"] = remain_days
        new_row["库存周转天数"] = turnover_days
        new_row["安全库存数量"] = safety_stock_qty
        new_row["平台仓目标库存"] = platform_target_stock
        new_row["返单周转目标库存"] = reorder_cycle_target_stock
        new_row["建议返单数量"] = suggested_reorder_qty
        new_row["预警-即将断货"] = bool(daily_sale > 0 and available <= daily_sale * safe_days)
        new_row["预警-急需立即返单"] = bool(sell_day is not None and remain_days <= 0)
        rows.append(new_row)

    return pd.DataFrame(rows)


# ===================== 页面 =====================
st.title("📦 跨境库存 ERP 计划系统")
st.caption("SKU库存、采购途中、海运途中、入FULL途中及海外仓批次统一管理")

with st.sidebar:
    st.header("⚙️ 计划参数")
    base_today = st.date_input("计划基准日期", value=date.today())
    global_purchase = st.number_input("采购交期(天)", min_value=1, value=20)
    global_ship = st.number_input("海运天数", min_value=1, value=50)
    global_customs = st.number_input("清关+预约天数", min_value=1, value=4)
    global_safety_days = st.number_input("安全库存天数", min_value=1, value=10)
    platform_stock_days = st.number_input("平台仓在库天数", min_value=1, value=20)
    reorder_turnover_days = st.number_input("返单周转天数", min_value=1, value=30)
    total_plan_days = (
        global_purchase
        + global_ship
        + global_customs
        + global_safety_days
        + platform_stock_days
        + reorder_turnover_days
    )
    st.info(f"完整供应链总周期：{total_plan_days} 天")
    st.caption(
        f"采购{global_purchase}天 + 海运{global_ship}天 + 清关预约{global_customs}天 "
        f"+ 安全库存{global_safety_days}天 + 平台仓在库{platform_stock_days}天 "
        f"+ 返单周转{reorder_turnover_days}天"
    )
    st.caption(
        "日期缺失默认：海运途中7天后开船、43天到达；采购途中7天完成采购、再50天到达。"
    )

    st.divider()
    st.header("📥 数据导入")

    sku_upload = st.file_uploader("上传SKU库存表", type=["xlsx", "csv"], key="sku_upload")
    if sku_upload is not None and st.button("导入SKU数据", width="stretch"):
        result = batch_import_sku(sku_upload)
        st.success(f"SKU导入完成，共 {len(result)} 个SKU")
        st.rerun()

    supply_upload = st.file_uploader(
        "上传《海运及转运货物明细》",
        type=["xlsx"],
        key="supply_upload"
    )
    if supply_upload is not None and st.button("导入并更新供应链批次", width="stretch"):
        result, imported_count = batch_import_supply(supply_upload)
        st.success(f"已读取 {imported_count} 行，合并后共 {len(result)} 条批次")
        st.rerun()

sku_df = init_sku_table()
supply_df = init_supply_table()

df_calc = build_inventory_calc(
    sku_df, supply_df, base_today,
    global_purchase, global_ship, global_customs,
    global_safety_days, platform_stock_days, reorder_turnover_days
)

tab_dashboard, tab_supply, tab_sku, tab_gantt, tab_export = st.tabs([
    "经营驾驶舱", "供应链批次总表", "SKU库存计划", "供应链甘特图", "导入导出"
])

# ===================== 经营驾驶舱 =====================
with tab_dashboard:
    if df_calc.empty:
        st.warning("暂无SKU数据，请先导入SKU库存表或供应链批次表。")
    else:
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("SKU总数", len(df_calc))
        k2.metric("可售现货", int(df_calc["可售现货库存"].sum()))
        k3.metric("全链路总库存", int(df_calc["全链路总库存(含在途)"].sum()))
        k4.metric("急需返单SKU", int(df_calc["预警-急需立即返单"].sum()))
        k5.metric("低于安全库存SKU", int(df_calc["预警-即将断货"].sum()))
        k6.metric("建议返单总量", int(df_calc["建议返单数量"].sum()))

        st.subheader("状态库存概览")
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("采购途中", int(df_calc["采购途中(批次合并)"].sum()))
        s2.metric("海运途中", int(df_calc["海运途中(批次合并)"].sum()))
        s3.metric("入FULL途中", int(df_calc["入Full途中(批次合并)"].sum()))
        s4.metric("海外仓", int(df_calc["海外仓库存(批次合并)"].sum()))
        s5.metric("丢失赔付", int(df_calc["批次丢失赔付"].sum()))

        st.subheader("🔴 急需立即返单")
        urgent = df_calc[df_calc["预警-急需立即返单"]]
        if urgent.empty:
            st.success("当前没有急需立即返单的SKU")
        else:
            st.dataframe(
                urgent[[
                    "SKU编码", "产品名称", "可售现货库存", "日均销量",
                    "预计售罄日期", "返单最晚截止日期", "距离必须返单剩余天数"
                ]],
                width="stretch",
                hide_index=True,
            )

# ===================== 供应链批次总表 =====================
with tab_supply:
    st.subheader("🚢 海运及转运货物明细")
    st.caption("采购、海运、海外仓、入FULL途中等状态集中在同一张表中。")

    if supply_df.empty:
        st.info("暂无供应链批次数据，请上传《海运及转运货物明细.xlsx》。")
    else:
        counts = supply_df["是否入仓"].value_counts().to_dict()
        status_options = [
            status for status in STATUS_ORDER
            if status in counts
        ]

        c1, c2, c3 = st.columns([2.2, 1.4, 1.4])
        with c1:
            keyword = st.text_input(
                "搜索",
                placeholder="输入公司SKU、MSKU、中文品名、Shipment ID、入仓号、物流商"
            )
        with c2:
            selected_status = st.multiselect(
                "状态筛选",
                options=status_options,
                default=status_options,
                format_func=lambda x: f"{x} ({counts.get(x, 0)})"
            )
        with c3:
            channel_options = sorted([x for x in supply_df["渠道"].unique().tolist() if x])
            selected_channels = st.multiselect("渠道", channel_options, default=channel_options)

        filtered = supply_df.copy()
        if selected_status:
            filtered = filtered[filtered["是否入仓"].isin(selected_status)]
        else:
            filtered = filtered.iloc[0:0]

        if selected_channels:
            filtered = filtered[filtered["渠道"].isin(selected_channels)]
        elif channel_options:
            filtered = filtered.iloc[0:0]

        if keyword.strip():
            q = keyword.strip().lower()
            search_cols = [
                "中文品名", "公司SKU", "MSKU", "备注", "入仓号/留仓号",
                "Shipment ID", "物流商", "渠道", "FULL仓"
            ]
            mask = pd.Series(False, index=filtered.index)
            for col in search_cols:
                mask = mask | filtered[col].fillna("").astype(str).str.lower().str.contains(q, regex=False)
            filtered = filtered[mask]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("筛选批次数", len(filtered))
        m2.metric("海外仓数量", int(pd.to_numeric(filtered["海外仓数量"], errors="coerce").fillna(0).sum()))
        m3.metric("剩余数量", int(pd.to_numeric(filtered["剩余数量"], errors="coerce").fillna(0).sum()))
        overdue = filtered[
            filtered["预估送达日期"].notna()
            & filtered["实际送达日期"].isna()
            & (filtered["预估送达日期"].dt.date < base_today)
        ]
        m4.metric("已超预计送达", len(overdue))

        display_df = filtered.copy()
        for col in DATE_COLUMNS:
            display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")

        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            height=620,
        )

        st.download_button(
            "下载当前筛选结果.xlsx",
            data=to_excel_bytes(display_df),
            file_name="供应链批次筛选结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ===================== SKU库存计划 =====================
with tab_sku:
    st.subheader("📋 SKU库存与返单计划")
    if df_calc.empty:
        st.info("暂无SKU库存计划数据。")
    else:
        show_columns = [
            "SKU编码", "产品名称", "深圳仓库存", "海外仓库存(批次合并)", "Full仓库存",
            "采购途中(批次合并)", "海运途中(批次合并)", "入Full途中(批次合并)",
            "可售现货库存", "全链路总库存(含在途)", "日均销量",
            "预计售罄日期", "返单最晚截止日期", "距离必须返单剩余天数",
            "库存周转天数", "安全库存数量", "平台仓目标库存",
            "返单周转目标库存", "建议返单数量",
            "预警-即将断货", "预警-急需立即返单"
        ]
        display = df_calc[show_columns].copy()
        for col in ["预计售罄日期", "返单最晚截止日期"]:
            display[col] = display[col].apply(
                lambda x: x.strftime("%Y-%m-%d") if x is not None and not pd.isna(x) else ""
            )
        display["库存周转天数"] = pd.to_numeric(display["库存周转天数"], errors="coerce").round(1)

        st.dataframe(display, width="stretch", hide_index=True, height=620)

# ===================== 甘特图 =====================
with tab_gantt:
    st.subheader("📊 单SKU未来库存销售时间线")
    st.caption(
        "为避免切换SKU时立即执行大量日期与图表计算，"
        "现在改为先选择SKU，再点击按钮加载测算结果。"
    )

    if df_calc.empty:
        st.info("暂无SKU数据。")
    else:
        sku_options = (
            df_calc["SKU编码"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        sku_options = sku_options[sku_options != ""].drop_duplicates().tolist()

        if not sku_options:
            st.warning("没有有效SKU可供选择。")
        else:
            selected_sku = st.selectbox(
                "选择SKU",
                sku_options,
                key="gantt_selected_sku",
            )

            load_key = f"gantt_loaded_sku_{selected_sku}"
            if load_key not in st.session_state:
                st.session_state[load_key] = False

            action_col1, action_col2 = st.columns([1, 4])
            with action_col1:
                if st.button(
                    "加载该SKU测算",
                    key=f"load_gantt_{selected_sku}",
                    width="stretch",
                ):
                    st.session_state[load_key] = True
            with action_col2:
                if st.session_state[load_key]:
                    st.success(f"已加载 SKU：{selected_sku}")
                else:
                    st.info("选择SKU后，请点击“加载该SKU测算”。")

            if st.session_state[load_key]:
                try:
                    target_rows = df_calc[
                        df_calc["SKU编码"].fillna("").astype(str).str.strip()
                        == str(selected_sku).strip()
                    ]

                    if target_rows.empty:
                        st.error("所选SKU在计算结果中不存在，请刷新页面后重试。")
                    else:
                        target = target_rows.iloc[0]

                        sales_col, save_col = st.columns([4, 1])
                        with sales_col:
                            gantt_daily_sales = st.number_input(
                                "日均销量（修改后立即重新测算全部库存销售区间）",
                                min_value=0.0,
                                value=float(target["日均销量"]),
                                step=0.1,
                                key=f"gantt_daily_sales_{selected_sku}",
                            )
                        with save_col:
                            st.write("")
                            st.write("")
                            if st.button(
                                "💾 保存日均销量",
                                key=f"save_sales_{selected_sku}",
                                width="stretch",
                            ):
                                sku_update = init_sku_table()
                                sku_update.loc[
                                    sku_update["SKU编码"].astype(str) == str(selected_sku),
                                    "日均销量"
                                ] = gantt_daily_sales
                                save_sku_df(sku_update)
                                st.success("日均销量已保存")
                                st.rerun()

                        default_reorder_qty = float(
                            target["计划返单数量"]
                            if "计划返单数量" in target.index
                            and pd.notna(target["计划返单数量"])
                            and float(target["计划返单数量"]) > 0
                            else gantt_daily_sales
                            * (
                                global_safety_days
                                + platform_stock_days
                                + reorder_turnover_days
                            )
                        )

                        reorder_input_col, reorder_save_col = st.columns([4, 1])
                        with reorder_input_col:
                            manual_reorder_qty = st.number_input(
                                "计划返单库存数量",
                                min_value=0.0,
                                value=default_reorder_qty,
                                step=1.0,
                                key=f"manual_reorder_qty_{selected_sku}",
                            )
                        with reorder_save_col:
                            st.write("")
                            st.write("")
                            if st.button(
                                "💾 保存返单数量",
                                key=f"save_reorder_qty_{selected_sku}",
                                width="stretch",
                            ):
                                sku_update = init_sku_table()
                                sku_update.loc[
                                    sku_update["SKU编码"].astype(str) == str(selected_sku),
                                    "计划返单数量"
                                ] = manual_reorder_qty
                                save_sku_df(sku_update)
                                st.success("计划返单数量已保存")
                                st.rerun()

                        if gantt_daily_sales <= 0:
                            st.warning("请将日均销量设置为大于0。")
                        elif gantt_daily_sales < 0.01:
                            st.warning("日均销量过小，请至少设置为0.01。")
                        else:
                            current_available = float(target["可售现货库存"])
                            lead_days = (
                                global_purchase + global_ship + global_customs
                            )
                            timeline_rows = []

                            if current_available > 0:
                                current_days = current_available / gantt_daily_sales
                                timeline_rows.append({
                                    "库存位置": "当前可售现货",
                                    "状态": "当前现货",
                                    "批次标识": "当前现货合计",
                                    "库存数量": current_available,
                                    "预计可售日期": pd.Timestamp(base_today),
                                    "预计销售结束": safe_future_timestamp(
                                        base_today, current_days
                                    ),
                                    "可销售天数": current_days,
                                    "优先级": 0,
                                })

                            sku_batches = supply_df[
                                supply_df["公司SKU"]
                                .fillna("")
                                .astype(str)
                                .str.strip()
                                == str(selected_sku).strip()
                            ].copy()

                            valid_statuses = [
                                "采购途中", "海运途中", "入full途中"
                            ]
                            sku_batches = sku_batches[
                                sku_batches["是否入仓"].isin(valid_statuses)
                            ].copy()

                            status_priority = {
                                "入full途中": 1,
                                "海运途中": 2,
                                "采购途中": 3,
                            }

                            for idx, row in sku_batches.iterrows():
                                status = row["是否入仓"]
                                qty = float(batch_quantity(row))
                                if qty <= 0:
                                    continue

                                available_date = expected_available_date_for_batch(
                                    row=row,
                                    status=status,
                                    base_today=base_today,
                                    purchase_days=global_purchase,
                                    ship_days=global_ship,
                                    customs_days=global_customs,
                                )
                                sales_days = qty / gantt_daily_sales
                                shipment = (
                                    clean_text(row.get("Shipment ID"))
                                    or clean_text(row.get("入仓号/留仓号"))
                                    or f"批次{idx + 1}"
                                )

                                timeline_rows.append({
                                    "库存位置": f"{status}｜{shipment}",
                                    "状态": status,
                                    "批次标识": shipment,
                                    "库存数量": qty,
                                    "预计可售日期": pd.Timestamp(available_date),
                                    "预计销售结束": safe_future_timestamp(
                                        available_date, sales_days
                                    ),
                                    "可销售天数": sales_days,
                                    "优先级": status_priority.get(status, 9),
                                })

                            if not timeline_rows:
                                st.warning("该SKU没有可绘制的现货或在途批次。")
                            else:
                                timeline_df = pd.DataFrame(timeline_rows)
                                timeline_df = timeline_df.sort_values(
                                    [
                                        "预计可售日期",
                                        "优先级",
                                        "预计销售结束",
                                    ],
                                    kind="stable",
                                ).reset_index(drop=True)

                                source_candidates = timeline_df[
                                    timeline_df["状态"] == "采购途中"
                                ]
                                source_name = "最新采购在途"

                                if source_candidates.empty:
                                    source_candidates = timeline_df[
                                        timeline_df["状态"] == "海运途中"
                                    ]
                                    source_name = "最新海运在途"

                                if source_candidates.empty:
                                    source_candidates = timeline_df
                                    source_name = "当前最新库存"

                                connection_source = source_candidates.sort_values(
                                    ["预计可售日期", "预计销售结束"],
                                    kind="stable",
                                ).iloc[-1]

                                planned_arrival = (
                                    connection_source["预计销售结束"]
                                    - pd.Timedelta(days=global_safety_days)
                                )
                                reorder_deadline = (
                                    planned_arrival
                                    - pd.Timedelta(days=lead_days)
                                )
                                reorder_qty = float(manual_reorder_qty)
                                reorder_sales_days = (
                                    reorder_qty / gantt_daily_sales
                                    if reorder_qty > 0
                                    else 0
                                )
                                reorder_sales_end = safe_future_timestamp(
                                    planned_arrival,
                                    reorder_sales_days,
                                )

                                if reorder_qty > 0:
                                    timeline_df = pd.concat([
                                        timeline_df,
                                        pd.DataFrame([{
                                            "库存位置": "计划返单库存",
                                            "状态": "计划返单",
                                            "批次标识": "手动计划返单",
                                            "库存数量": reorder_qty,
                                            "预计可售日期": planned_arrival,
                                            "预计销售结束": reorder_sales_end,
                                            "可销售天数": reorder_sales_days,
                                            "优先级": 4,
                                        }])
                                    ], ignore_index=True)

                                timeline_df = timeline_df.sort_values(
                                    [
                                        "预计可售日期",
                                        "优先级",
                                        "预计销售结束",
                                    ],
                                    kind="stable",
                                ).reset_index(drop=True)

                                coverage_end = None
                                check_rows = []
                                gap_ranges = []
                                overlap_ranges = []

                                for _, row in timeline_df.iterrows():
                                    start_date = row["预计可售日期"]
                                    end_date = row["预计销售结束"]

                                    if coverage_end is None:
                                        relation = "起始库存"
                                    elif start_date <= coverage_end:
                                        overlap_days = max(
                                            0.0,
                                            (
                                                coverage_end - start_date
                                            ).total_seconds() / 86400,
                                        )
                                        relation = (
                                            f"🟢 重合 {overlap_days:.1f} 天"
                                        )
                                        if overlap_days > 0:
                                            overlap_ranges.append(
                                                (
                                                    start_date,
                                                    min(coverage_end, end_date),
                                                )
                                            )
                                    else:
                                        gap_days = (
                                            start_date - coverage_end
                                        ).total_seconds() / 86400
                                        relation = (
                                            f"🔴 断货 {gap_days:.1f} 天"
                                        )
                                        gap_ranges.append(
                                            (coverage_end, start_date)
                                        )

                                    coverage_end = (
                                        end_date
                                        if coverage_end is None
                                        else max(coverage_end, end_date)
                                    )

                                    check_rows.append({
                                        "库存位置/批次": row["库存位置"],
                                        "状态": row["状态"],
                                        "库存数量": round(
                                            float(row["库存数量"]), 1
                                        ),
                                        "预计可售日期": start_date.strftime(
                                            "%Y-%m-%d"
                                        ),
                                        "预计销售结束": end_date.strftime(
                                            "%Y-%m-%d"
                                        ),
                                        "可销售天数": round(
                                            float(row["可销售天数"]), 1
                                        ),
                                        "与前序库存关系": relation,
                                    })

                                stockout_days = sum(
                                    (b - a).total_seconds() / 86400
                                    for a, b in gap_ranges
                                )
                                overlap_days_total = sum(
                                    (b - a).total_seconds() / 86400
                                    for a, b in overlap_ranges
                                )

                                st.markdown("#### 📦 计划返单库存测算")
                                rq1, rq2, rq3, rq4 = st.columns(4)
                                rq1.metric(
                                    "计划返单数量",
                                    round(reorder_qty, 1),
                                )
                                rq2.metric(
                                    "返单库存可售天数",
                                    round(reorder_sales_days, 1),
                                )
                                rq3.metric(
                                    "计划到货日期",
                                    planned_arrival.strftime("%Y-%m-%d"),
                                )
                                rq4.metric(
                                    "返单库存销售结束",
                                    reorder_sales_end.strftime("%Y-%m-%d")
                                    if reorder_qty > 0
                                    else "无",
                                )

                                m1, m2, m3, m4, m5 = st.columns(5)
                                m1.metric(
                                    "日均销量",
                                    round(gantt_daily_sales, 2),
                                )
                                m2.metric("返单衔接依据", source_name)
                                m3.metric(
                                    "最晚返单日期",
                                    reorder_deadline.strftime("%Y-%m-%d"),
                                )
                                m4.metric(
                                    "预计断货合计",
                                    f"{stockout_days:.1f} 天",
                                )
                                m5.metric(
                                    "库存重合合计",
                                    f"{overlap_days_total:.1f} 天",
                                )

                                st.markdown("#### 🔍 库存衔接检查")
                                st.dataframe(
                                    pd.DataFrame(check_rows),
                                    width="stretch",
                                    hide_index=True,
                                )

                                show_chart = st.toggle(
                                    "显示图形",
                                    value=False,
                                    key=f"show_chart_{selected_sku}",
                                )

                                if show_chart:
                                    plot_df = timeline_df.copy()
                                    plot_df["标签"] = plot_df.apply(
                                        lambda r: (
                                            f"{r['库存位置']}｜"
                                            f"{r['库存数量']:.0f}件｜"
                                            f"{r['可销售天数']:.1f}天"
                                        ),
                                        axis=1,
                                    )

                                    fig = px.timeline(
                                        plot_df,
                                        x_start="预计可售日期",
                                        x_end="预计销售结束",
                                        y="标签",
                                        color="状态",
                                    )
                                    fig.update_yaxes(
                                        autorange="reversed",
                                        title=None,
                                    )
                                    fig.update_xaxes(
                                        title="未来日期",
                                        type="date",
                                        tickformat="%Y-%m-%d",
                                    )
                                    fig.update_layout(
                                        height=max(
                                            520,
                                            80 + len(plot_df) * 55,
                                        ),
                                        legend_title_text="库存状态",
                                        margin=dict(
                                            l=20, r=20, t=40, b=20
                                        ),
                                    )

                                    try:
                                        st.plotly_chart(
                                            fig,
                                            width="stretch",
                                            key=f"timeline_chart_{selected_sku}",
                                        )
                                    except Exception as chart_error:
                                        st.error(
                                            "该SKU图表无法渲染，"
                                            "但测算结果仍然有效。"
                                        )
                                        st.exception(chart_error)

                except Exception as gantt_error:
                    st.error(
                        "该SKU测算发生异常。页面不会再黑屏，"
                        "请展开下方错误信息。"
                    )
                    st.exception(gantt_error)

# ===================== 导入导出 =====================
with tab_export:
    st.subheader("📥 数据模板与备份")

    sku_template = pd.DataFrame(columns=SKU_COLUMNS)
    supply_template = pd.DataFrame(columns=SUPPLY_COLUMNS)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "下载SKU导入模板",
            data=to_excel_bytes(sku_template),
            file_name="SKU库存导入模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.download_button(
            "导出SKU库存计划",
            data=to_excel_bytes(df_calc),
            file_name="跨境SKU库存返单测算.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    with d2:
        st.download_button(
            "下载供应链批次模板",
            data=to_excel_bytes(supply_template),
            file_name="海运及转运货物明细模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.download_button(
            "导出全部供应链批次",
            data=to_excel_bytes(supply_df),
            file_name="供应链批次完整备份.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
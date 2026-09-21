# -*- coding: utf-8 -*-
"""
=================================================================================
电商 BI 数据分析项目 —— Python 数据处理主脚本
数据集：Kaggle "Brazilian E-commerce Public Dataset by Olist"
（巴西 Olist 电商平台 2016-09 ~ 2018-10 的真实订单数据，约 10 万订单）

输出（全部位于 ./output/，UTF-8-sig 编码，Excel/Tableau/Power BI 均可直接读取）：
  1. overview.csv               全局核心指标（单行）
  2. daily_sales.csv            每日销售趋势
  3. monthly_sales.csv          每月销售趋势（YYYY-MM）
  4. rfm.csv                    用户 RFM 打分与分层（customer_unique_id 粒度）
  5. category_sales.csv         商品品类销售与评价分析（按 GMV 降序）
  6. review_analysis.csv        品类差评率排行（按差评率降序）
  7. rating_distribution.csv    评价星级 1~5 分布（看板“评价分布”页）
  8. payment_distribution.csv   支付方式笔数/金额分布
  9. order_funnel.csv           订单状态粗略漏斗（创建→支付→交付）

---------------------------------------------------------------------------------
【口径与规则总说明（面试可直接讲）】

1. 统一分析口径：已交付订单（order_status == 'delivered'）
   - 原需求中 GMV 取“已支付订单”、订单数取“已交付订单”，分子分母口径不一致
     （canceled/shipped 等未交付订单也有约 58.7 万雷亚尔支付记录），会使客单价虚高。
   - 经确认，全部分析统一为 delivered 口径；趋势/RFM/品类分析均使用同一口径，
     保证所有数字之间可交叉验证。
   - delivered 订单中有 1 个订单无支付记录：保留在订单数中，其支付额按 0 处理。

2. 复购率定义
   - 复购率 = 购买次数 >= 2 的客户数 / 有购买的客户总数（按 customer_unique_id 去重）。
   - 为什么用 >=2 次：复购的业务含义是“同一客户第二次回来购买”，1 次是首购基线，
     从第 2 次起才算复购；这是电商行业最通用、最容易向业务方解释的口径。
   - 注意：Olist 的 customer_id 是“按订单生成”的，同一人每下一单 customer_id 都不同，
     必须用 customer_unique_id 才能识别同一个人（面试高频考点）。

3. RFM 分箱规则（基准日 = 数据中最后一个下单时间 2018-10-17）
   - R(Recency)  ：最近一次购买距基准日的天数。R 越小分越高（最近购买=4分，
                   最久未买=1分），用四分位数 qcut 分 4 档，标签 [4,3,2,1] 反转。
   - F(Frequency)：购买次数（去重 order_id）。本数据集 97.0% 的客户只买过 1 次，
                   F 的 25/50/75 分位数全部=1，四分位数分箱会退化为“所有人 1 分”，
                   高价值/潜力分层将全部为 0。因此 F 改用业务规则分箱：
                   1次=1分，2次=2分，3次=3分，>=4次=4分。
   - M(Monetary) ：累计支付金额，越大分越高，四分位数 qcut 分 4 档，标签 [1,2,3,4]。
   - 为什么 R/M 用分位数而不是等宽：R/M 是连续且右偏严重的变量（少数人金额极高），
     等宽分箱会把绝大多数客户压在同一档；分位数保证每档人数约 25%，分层区分度最好。
     F 是离散且极端倾斜的计数变量，分位数退化，故改用可解释的业务阈值。
     （“先检查变量分布，再决定分箱方法”本身就是面试加分项。）
   - rfm_total_score = R_score + F_score + M_score（3~12 分）
   - 分层（按优先级从上到下判定，互斥）：
       高价值用户  ：总分 >= 10
       潜力用户    ：F_score>=3 且 M_score>=3 且 R_score<=2（高消费但已沉睡，可唤回）
       流失风险用户：R_score == 1（最久未购买的 1/4 客户）
       一般用户    ：其余

4. 数据清洗做了什么（每步在代码中均有注释和删除行数打印）
   - orders ：日期字段转 datetime；只保留 delivered；过滤“下单晚于送达”等异常时间戳；
              delivered 但缺送达时间戳的 8 单保留（销售按“下单时间”聚合，不受影响）。
   - items  ：丢弃 price<=0、freight_value<0 的行（本数据集实际为 0 行）；
              丢弃不属于 delivered 订单的明细；丢弃 orders 中不存在的孤立 order_id（实际 0 行）。
              【重要】不按 99 分位截断高价：经核查 >99 分位(890 雷亚尔)的 1117 行全部是
              电脑/家电/乐器/游戏机/手表等真实高价成交商品，截断会扭曲品类 GMV 排名。
   - products：品类名缺失填 'unknown'；葡语品类名关联翻译表转英文；
              2 个品类翻译表未覆盖（pc_gamer、portateis_cozinha_e_preparadores_de_alimentos），
              按约定保留原名。
   - payments：聚合到订单级（支付总额、支付方式数）；与 delivered 订单左连接（1 单无支付记 0）。
   - reviews ：547 个订单有多条评价，按订单取平均评分；订单均分 <=2 标记为差评。

5. 两种 GMV 口径（务必区分，面试容易被追问）
   - 全局/趋势/RFM 的 gmv = 订单 payment_value 合计（客户实付，含运费）。
   - 品类表 total_gmv = 该品类商品 price 合计（不含运费）：运费是订单级费用，一单多品类时
     无法合理分摊，故品类分析只统计商品金额；两张口径的差额≈全平台运费，属正常现象。

6. 数据局限（简历/面试如实说明）
   - Olist 是平台聚合脱敏数据，无浏览/点击/加购数据，无法做完整营销漏斗，
     只能计算粗略“成交率 = delivered 订单 / 全部创建订单”。
   - 2016 年仅有零星订单、2018-10 月数据不完整（月趋势末端断崖是数据截断，非业务下跌）。
   - 复购率低（约 3%）与平台模式有关：Olist 连接的是众多独立小卖家，客户跨店复购
     本就不记为同店复购，解读时需说明。
=================================================================================
"""

import os
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# 全局配置：全部使用相对路径，不硬编码绝对路径
# ----------------------------------------------------------------------
DATA_DIR = "./data"
OUTPUT_DIR = "./output"

# Olist 9 张表的文件名
FILES = {
    "orders": "olist_orders_dataset.csv",
    "items": "olist_order_items_dataset.csv",
    "products": "olist_products_dataset.csv",
    "customers": "olist_customers_dataset.csv",
    "sellers": "olist_sellers_dataset.csv",
    "payments": "olist_order_payments_dataset.csv",
    "reviews": "olist_order_reviews_dataset.csv",
    "geolocation": "olist_geolocation_dataset.csv",
    "translation": "product_category_name_translation.csv",
}

# orders 表中需要解析为 datetime 的列（注意 purchase 列以 _timestamp 结尾）
ORDER_DATE_COLS = [
    "order_purchase_timestamp",
    "order_approved_at",
    "order_delivered_carrier_date",
    "order_delivered_customer_date",
    "order_estimated_delivery_date",
]


# ======================================================================
# 1. 数据加载与探索
# ======================================================================
def load_data(data_dir=DATA_DIR):
    """读取 data_dir 下全部 CSV，返回 {表名: DataFrame} 字典。"""
    dfs = {}
    for name, fn in FILES.items():
        path = os.path.join(data_dir, fn)
        dfs[name] = pd.read_csv(path)
    return dfs


def explore_data(dfs):
    """打印每张表的行数、列数、缺失值比例、数据类型、唯一值数量。"""
    print("=" * 90)
    print("【数据概览】各表规模 / 数据类型 / 缺失率")
    print("=" * 90)
    for name, df in dfs.items():
        print(f"\n### {name:12s}  行数={len(df):,}  列数={df.shape[1]}")
        miss_ratio = (df.isna().mean() * 100).round(2)
        for col in df.columns:
            print(f"    {col:38s} dtype={str(df[col].dtype):12s} "
                  f"缺失率={miss_ratio[col]:5.2f}%  唯一值={df[col].nunique():,}")


def check_key_integrity(dfs):
    """
    检查表间关联键完整性：从表的外键是否都能在主表中找到（孤立键计数）。
    返回检查结果字典，并打印结论。
    """
    orders, items = dfs["orders"], dfs["items"]
    products, customers = dfs["products"], dfs["customers"]
    sellers, payments = dfs["sellers"], dfs["payments"]
    reviews = dfs["reviews"]

    order_ids = set(orders["order_id"])
    checks = {
        "orders.order_id 是否重复": int(orders["order_id"].duplicated().sum()),
        "items 孤立 order_id 行数": int((~items["order_id"].isin(order_ids)).sum()),
        "items 孤立 product_id 行数": int((~items["product_id"].isin(set(products["product_id"]))).sum()),
        "items 孤立 seller_id 行数": int((~items["seller_id"].isin(set(sellers["seller_id"]))).sum()),
        "orders 孤立 customer_id 行数": int((~orders["customer_id"].isin(set(customers["customer_id"]))).sum()),
        "payments 孤立 order_id 行数": int((~payments["order_id"].isin(order_ids)).sum()),
        "reviews 孤立 order_id 行数": int((~reviews["order_id"].isin(order_ids)).sum()),
    }

    print("\n" + "=" * 90)
    print("【关联键完整性检查】（孤立键 = 从表外键在主表中找不到对应主键的行数）")
    print("=" * 90)
    for k, v in checks.items():
        flag = "OK" if v == 0 else f"存在 {v} 行问题"
        print(f"    {k:28s}: {v:,}  [{flag}]")
    return checks


# ======================================================================
# 2. 数据清洗
# ======================================================================
def clean_orders(orders):
    """
    清洗 orders 表：
      - 5 个时间字段解析为 datetime（无法解析的变 NaT）；
      - 只保留 order_status == 'delivered'（统一分析口径）；
      - 过滤时间逻辑异常：下单时间晚于客户签收时间；
      - 过滤签收时间超出数据合理范围（比预计送达晚 365 天以上，本数据实际 0 行）。
    返回清洗后的 delivered 订单 DataFrame。
    """
    df = orders.copy()
    n0 = len(df)

    for col in ORDER_DATE_COLS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # 只保留已交付订单
    df = df[df["order_status"] == "delivered"].copy()
    n_after_status = len(df)

    # 异常时间戳 1：下单晚于签收（逻辑不可能）
    bad_purchase = df["order_purchase_timestamp"] > df["order_delivered_customer_date"]
    # 异常时间戳 2：实际签收比预计送达晚 365 天以上（疑似录入错误/时间戳溢出）
    bad_range = df["order_delivered_customer_date"] > (
        df["order_estimated_delivery_date"] + pd.Timedelta(days=365)
    )
    bad_mask = bad_purchase | bad_range
    n_bad = int(bad_mask.sum())
    df = df[~bad_mask].copy()

    print("\n" + "-" * 90)
    print("【清洗 orders】")
    print(f"    原始行数 {n0:,} -> delivered {n_after_status:,} "
          f"（剔除其他状态 {n0 - n_after_status:,} 行）")
    print(f"    剔除时间戳逻辑异常订单 {n_bad} 行；最终保留 {len(df):,} 行")
    print(f"    提示：{int(df['order_delivered_customer_date'].isna().sum())} 个 delivered 订单缺签收时间，"
          f"因销售按下单时间聚合，予以保留")
    return df


def clean_items(items, valid_order_ids):
    """
    清洗 order_items 表：
      - 丢弃 price <= 0 的行（价格必须为正，本数据实际 0 行）；
      - 丢弃 freight_value < 0 的行（运费不能为负，本数据实际 0 行）；
      - 丢弃不属于有效 delivered 订单的明细（含孤立 order_id，本数据孤立键为 0）。
    【刻意不做】按 99 分位截断高价：经核查高价行均为电脑/家电等真实成交，截断会失真。
    """
    df = items.copy()
    n0 = len(df)

    bad_price = df["price"] <= 0
    bad_freight = df["freight_value"] < 0
    bad_order = ~df["order_id"].isin(valid_order_ids)

    print("\n" + "-" * 90)
    print("【清洗 order_items】")
    print(f"    price<=0 行数 {int(bad_price.sum())}；freight<0 行数 {int(bad_freight.sum())}；"
          f"不属于 delivered 订单的明细 {int(bad_order.sum())} 行（一并剔除孤立键）")

    df = df[~(bad_price | bad_freight | bad_order)].copy()
    print(f"    原始 {n0:,} 行 -> 保留 {len(df):,} 行")
    return df


def clean_products(products, translation):
    """
    清洗 products 表并翻译品类名：
      - product_category_name 缺失填 'unknown'；
      - 左连接翻译表得到英文名；
      - 翻译表未覆盖的 2 个品类保留葡萄牙语原名。
    返回带 product_category_name_english 列的商品表。
    """
    df = products.merge(translation, on="product_category_name", how="left")
    n_miss = int(df["product_category_name"].isna().sum())

    # 缺失品类 -> unknown；已翻译用英文名；未翻译（english 为空）保留葡语原名
    df["category_name"] = df["product_category_name_english"]
    df.loc[df["product_category_name"].isna(), "category_name"] = "unknown"
    df["category_name"] = df["category_name"].fillna(df["product_category_name"])

    uncovered = sorted(
        df.loc[df["product_category_name_english"].isna()
               & df["product_category_name"].notna(), "product_category_name"].unique()
    )
    print("\n" + "-" * 90)
    print("【清洗 products / 品类翻译】")
    print(f"    品类名缺失商品 {n_miss} 个，已填为 unknown；翻译表未覆盖、保留原名的品类: {uncovered}")
    return df


def aggregate_payments(payments, valid_order_ids):
    """
    将 payments 聚合到订单级别：
      - payment_value：订单总支付额（同一订单分期/券抵扣会有多行，需求和）；
      - payment_type_count：该订单使用的支付方式种数。
    只保留 delivered 订单（无支付记录的订单后续左连接补 0）。
    同时打印支付方式分布，作为面试素材。
    """
    pay = payments[payments["order_id"].isin(valid_order_ids)].copy()
    pay_order = (
        pay.groupby("order_id")
        .agg(total_payment_value=("payment_value", "sum"),
             payment_type_count=("payment_type", "nunique"))
        .reset_index()
    )

    print("\n" + "-" * 90)
    print("【payments 聚合到订单级】")
    print("    delivered 订单支付方式行数分布：")
    for ptype, cnt in pay["payment_type"].value_counts().items():
        print(f"        {ptype:12s}: {cnt:,} 行")
    print(f"    聚合后订单数：{len(pay_order):,}（delivered 订单中有 1 单无支付记录，后续补 0）")
    return pay_order


def aggregate_reviews(reviews, valid_order_ids):
    """
    将 reviews 聚合到订单级别：
      - 547 个订单存在多条评价，按订单取平均评分；
      - avg_review_score：订单平均评分（1~5，可能为小数）；
      - is_bad_review：订单均分 <= 2 标记为差评（1=差评，0=非差评）。
    """
    rv = reviews[reviews["order_id"].isin(valid_order_ids)].copy()
    rv_order = (
        rv.groupby("order_id")["review_score"]
        .mean()
        .rename("avg_review_score")
        .reset_index()
    )
    rv_order["is_bad_review"] = (rv_order["avg_review_score"] <= 2).astype(int)

    print("\n" + "-" * 90)
    print("【reviews 聚合到订单级】")
    print(f"    评价明细 {len(rv):,} 行 -> 覆盖订单 {len(rv_order):,} 个"
          f"（其中 {int((rv.groupby('order_id').size() > 1).sum())} 个订单有多条评价，已取均值）")
    print(f"    差评订单（均分<=2）{int(rv_order['is_bad_review'].sum()):,} 个，"
          f"差评率 {rv_order['is_bad_review'].mean() * 100:.2f}%")
    return rv_order


# ======================================================================
# 3. 组装宽表
# ======================================================================
def build_order_wide(orders_clean, customers, pay_order, rv_order):
    """
    构建订单级分析宽表：delivered 订单 + 客户 unique_id + 支付额 + 评分。
    后续 overview / 趋势 / RFM 全部基于这张表，保证口径统一。
    """
    df = (
        orders_clean
        .merge(customers[["customer_id", "customer_unique_id", "customer_state"]],
               on="customer_id", how="left")
        .merge(pay_order, on="order_id", how="left")
        .merge(rv_order, on="order_id", how="left")
    )
    # 1 个 delivered 订单无支付记录：支付额补 0、支付方式数补 0
    df["total_payment_value"] = df["total_payment_value"].fillna(0.0)
    df["payment_type_count"] = df["payment_type_count"].fillna(0).astype(int)
    # 无评价订单的评分保持 NaN（聚合时自然排除），差评标记不填 1
    return df


def build_item_wide(items_clean, products_clean, rv_order):
    """
    构建商品明细级宽表：明细 + 英文品类 + 订单平均评分。
    用于品类销量/GMV 统计，以及“订单×品类”口径的评价统计。
    """
    df = (
        items_clean
        .merge(products_clean[["product_id", "category_name"]], on="product_id", how="left")
        .merge(rv_order, on="order_id", how="left")
    )
    df["category_name"] = df["category_name"].fillna("unknown")
    return df


# ======================================================================
# 4. 指标计算
# ======================================================================
def build_overview(order_wide):
    """
    全局总览指标（单行）：
      total_gmv       ：delivered 订单实付总额（payment_value 合计，含运费）
      total_orders    ：delivered 订单总数
      avg_order_value ：total_gmv / total_orders（客单价 AOV）
      total_customers ：去重 customer_unique_id 数量
      repurchase_rate ：购买>=2 次客户数 / 总购买客户数
    """
    total_gmv = float(order_wide["total_payment_value"].sum())
    total_orders = int(order_wide["order_id"].nunique())
    avg_order_value = total_gmv / total_orders

    freq = order_wide.groupby("customer_unique_id")["order_id"].nunique()
    total_customers = int(freq.shape[0])
    repurchase_rate = float((freq >= 2).sum() / total_customers)

    overview = pd.DataFrame([{
        "total_gmv": round(total_gmv, 2),
        "total_orders": total_orders,
        "avg_order_value": round(avg_order_value, 2),
        "total_customers": total_customers,
        "repurchase_rate": round(repurchase_rate, 4),
    }])

    print("\n" + "=" * 90)
    print("【总览指标 overview】")
    print(f"    GMV={total_gmv:,.2f} | 订单数={total_orders:,} | AOV={avg_order_value:,.2f}")
    print(f"    客户数={total_customers:,} | 复购率={repurchase_rate * 100:.2f}%")
    return overview


def build_daily_sales(order_wide):
    """
    每日销售趋势（按订单创建日期 order_purchase_timestamp 聚合）：
      date             ：日期（YYYY-MM-DD）
      order_count      ：订单数（去重 order_id）
      gmv              ：当日实付总额
      aov              ：gmv / order_count（客单价）
      paying_customers ：当日有支付的去重客户数
    注：不强制补全无订单的日历日期，避免人为制造 0 值；在 BI 工具中按需处理。
    """
    df = order_wide.copy()
    df["date"] = df["order_purchase_timestamp"].dt.date.astype(str)

    paid = df[df["total_payment_value"] > 0]
    daily = (
        df.groupby("date")
        .agg(order_count=("order_id", "nunique"),
             gmv=("total_payment_value", "sum"))
        .reset_index()
    )
    daily_pay = (
        paid.groupby("date")["customer_unique_id"].nunique()
        .rename("paying_customers").reset_index()
    )
    daily = daily.merge(daily_pay, on="date", how="left")
    daily["paying_customers"] = daily["paying_customers"].fillna(0).astype(int)
    daily["gmv"] = daily["gmv"].round(2)
    daily["aov"] = (daily["gmv"] / daily["order_count"]).round(2)
    daily = daily[["date", "order_count", "gmv", "aov", "paying_customers"]]
    daily = daily.sort_values("date").reset_index(drop=True)
    return daily


def build_monthly_sales(order_wide):
    """
    每月销售趋势（按订单创建月份聚合，date 格式 YYYY-MM），字段同每日趋势。
    注意：2016 年订单极少、2018-10 为不完整月份，看板解读时需说明。
    """
    df = order_wide.copy()
    df["date"] = df["order_purchase_timestamp"].dt.to_period("M").astype(str)

    paid = df[df["total_payment_value"] > 0]
    monthly = (
        df.groupby("date")
        .agg(order_count=("order_id", "nunique"),
             gmv=("total_payment_value", "sum"))
        .reset_index()
    )
    monthly_pay = (
        paid.groupby("date")["customer_unique_id"].nunique()
        .rename("paying_customers").reset_index()
    )
    monthly = monthly.merge(monthly_pay, on="date", how="left")
    monthly["paying_customers"] = monthly["paying_customers"].fillna(0).astype(int)
    monthly["gmv"] = monthly["gmv"].round(2)
    monthly["aov"] = (monthly["gmv"] / monthly["order_count"]).round(2)
    monthly = monthly[["date", "order_count", "gmv", "aov", "paying_customers"]]
    monthly = monthly.sort_values("date").reset_index(drop=True)
    return monthly


def build_rfm(order_wide, snapshot_date):
    """
    RFM 用户分层（customer_unique_id 粒度）。
    分箱规则见文件头部说明：R/M 用四分位数（R 反转），F 用业务规则分箱。
    snapshot_date：分析基准日（数据中最后一个订单的下单时间，取全量 orders 口径）。
    输出列：customer_unique_id, R, F, M, R_score, F_score, M_score,
            rfm_total_score, segment
    """
    rfm = (
        order_wide.groupby("customer_unique_id")
        .agg(
            R=("order_purchase_timestamp", lambda s: int((snapshot_date - s.max()).days)),
            F=("order_id", "nunique"),
            M=("total_payment_value", "sum"),
        )
        .reset_index()
    )
    rfm["M"] = rfm["M"].round(2)

    # R：天数越少分越高 -> qcut 标签反转 [4,3,2,1]
    rfm["R_score"] = pd.qcut(rfm["R"], 4, labels=[4, 3, 2, 1]).astype(int)
    # M：金额越大分越高
    rfm["M_score"] = pd.qcut(rfm["M"], 4, labels=[1, 2, 3, 4]).astype(int)
    # F：四分位数在本数据退化（75 分位仍为 1），改用业务规则分箱
    rfm["F_score"] = rfm["F"].map(
        lambda v: 1 if v == 1 else 2 if v == 2 else 3 if v == 3 else 4
    ).astype(int)

    rfm["rfm_total_score"] = rfm[["R_score", "F_score", "M_score"]].sum(axis=1)

    # 分层（np.select 按顺序判定，先命中先得，保证互斥）
    cond_high = rfm["rfm_total_score"] >= 10
    cond_potential = (
        (rfm["F_score"] >= 3) & (rfm["M_score"] >= 3) & (rfm["R_score"] <= 2)
    )
    cond_churn = rfm["R_score"] == 1
    rfm["segment"] = np.select(
        [cond_high, cond_potential, cond_churn],
        ["高价值用户", "潜力用户", "流失风险用户"],
        default="一般用户",
    )

    print("\n" + "=" * 90)
    print(f"【RFM 分层】基准日 = {snapshot_date}")
    seg_stat = rfm["segment"].value_counts()
    seg_ratio = (rfm["segment"].value_counts(normalize=True) * 100).round(2)
    for seg_name in ["高价值用户", "潜力用户", "流失风险用户", "一般用户"]:
        print(f"    {seg_name:6s}: {seg_stat.get(seg_name, 0):,} 人 "
              f"({seg_ratio.get(seg_name, 0):.2f}%)")
    # 面试素材：高价值用户的人均消费与 GMV 占比（如实呈现，不套用二八法则）
    hv = rfm[rfm["segment"] == "高价值用户"]
    hv_gmv_ratio = hv["M"].sum() / rfm["M"].sum() * 100
    hv_avg_m = hv["M"].mean()
    all_avg_m = rfm["M"].mean()
    print(f"    >> 业务发现：高价值用户 {len(hv):,} 人（占 {len(hv)/len(rfm)*100:.2f}%），"
          f"贡献 GMV 占比 {hv_gmv_ratio:.2f}%；")
    print(f"       人均累计消费 {hv_avg_m:,.2f}，是全体人均 {all_avg_m:,.2f} 的 "
          f"{hv_avg_m / all_avg_m:.1f} 倍。该平台复购率低，二八效应不典型，"
          f"说明用户关系运营仍有较大空间（面试如实讲，比硬套二八法则更稳）。")

    return rfm[["customer_unique_id", "R", "F", "M", "R_score", "F_score",
                "M_score", "rfm_total_score", "segment"]]


def build_category_sales(item_wide):
    """
    品类销售分析（品类 total_gmv 用商品 price 口径，不含运费，理由见文件头）：
      category_name    ：英文品类名（未翻译保留原名，缺失为 unknown）
      sales_count      ：销量（商品件数 = 明细行数）
      total_gmv        ：该品类商品销售额合计（price 求和）
      avg_price        ：件均价
      avg_review_score ：该品类订单平均评分（订单×品类去重，避免一单多件重复计评）
      bad_review_rate  ：差评率（订单均分<=2 的订单×品类数 / 有评价的订单×品类数）
    按 total_gmv 降序。
    """
    # --- 销售部分（明细粒度） ---
    sales = (
        item_wide.groupby("category_name")
        .agg(sales_count=("order_id", "size"),
             total_gmv=("price", "sum"),
             avg_price=("price", "mean"))
        .reset_index()
    )

    # --- 评价部分（订单×品类去重，同一订单买同品类多件只计一次评价） ---
    review_tbl = _category_review_table(item_wide).rename(
        columns={"avg_score": "avg_review_score"})

    df = sales.merge(review_tbl, on="category_name", how="left")
    df["total_gmv"] = df["total_gmv"].round(2)
    df["avg_price"] = df["avg_price"].round(2)
    df["avg_review_score"] = df["avg_review_score"].round(3)
    df["bad_review_rate"] = df["bad_review_rate"].round(4)
    df = df.sort_values("total_gmv", ascending=False).reset_index(drop=True)

    # 面试素材：GMV TOP10 品类中差评率偏高的（高销售+低口碑，运营改进机会）
    print("\n" + "=" * 90)
    print("【品类分析】GMV TOP10 品类及其差评率（关注“卖得多但差评率高”的品类）")
    top10 = df.head(10)
    for _, r in top10.iterrows():
        print(f"    {r['category_name']:38s} GMV={r['total_gmv']:>12,.2f} "
              f"评分={r['avg_review_score'] if pd.notna(r['avg_review_score']) else 'NA':>5} "
              f"差评率={r['bad_review_rate'] if pd.notna(r['bad_review_rate']) else 'NA'}")
    return df


def _category_review_table(item_wide):
    """
    生成品类级评价汇总（category_sales 与 review_analysis 共用，保证两表口径一致）：
      review_count       ：有评分的“订单×品类”去重数
      avg_score          ：平均订单评分
      bad_review_count   ：差评（订单均分<=2）的订单×品类数
      bad_review_rate    ：差评数 / 评价数
    """
    oc = item_wide[["order_id", "category_name", "avg_review_score", "is_bad_review"]].dropna(
        subset=["avg_review_score"]
    ).drop_duplicates(subset=["order_id", "category_name"])

    tbl = (
        oc.groupby("category_name")
        .agg(review_count=("order_id", "size"),
             avg_score=("avg_review_score", "mean"),
             bad_review_count=("is_bad_review", "sum"))
        .reset_index()
    )
    tbl["bad_review_count"] = tbl["bad_review_count"].astype(int)
    tbl["bad_review_rate"] = tbl["bad_review_count"] / tbl["review_count"]
    tbl["avg_score"] = tbl["avg_score"].round(3)
    tbl["bad_review_rate"] = tbl["bad_review_rate"].round(4)
    return tbl


def build_review_analysis(item_wide):
    """
    品类评价分析：在共用评价汇总表基础上按 bad_review_rate 降序排列，
    用于快速定位差评率最高、最需要质量/物流改进的品类。
    输出列：category_name, review_count, avg_score, bad_review_count, bad_review_rate
    """
    tbl = _category_review_table(item_wide)
    # 评价样本太少（<30）的品类差评率波动大，单独打标记列便于看板筛选，但不删除
    tbl["low_sample_flag"] = (tbl["review_count"] < 30).astype(int)
    tbl = tbl.sort_values(["bad_review_rate", "review_count"],
                          ascending=[False, False]).reset_index(drop=True)

    print("\n" + "=" * 90)
    print("【差评率 TOP10 品类】（review_count>=30 时结论更稳健，括号内为评价数）")
    shown = 0
    for _, r in tbl.iterrows():
        if r["review_count"] < 30:
            continue
        print(f"    {r['category_name']:38s} 差评率={r['bad_review_rate']:.2%} "
              f"均分={r['avg_score']:.2f} (评价数 {int(r['review_count']):,})")
        shown += 1
        if shown >= 10:
            break
    return tbl


# ======================================================================
# 5. 看板辅助分布表（Power BI / Tableau 常用，支撑“评价分布/支付方式/漏斗”页）
# ======================================================================
def build_rating_distribution(reviews, valid_order_ids):
    """
    评价星级分布（delivered 口径，按评价明细行统计 1~5 星各多少条、占比）。
    用于看板“评价分布”页（饼图/柱状图）。
    """
    rv = reviews[reviews["order_id"].isin(valid_order_ids)]
    tbl = (
        rv.groupby("review_score").size()
        .rename("review_count").reset_index()
    )
    tbl["ratio"] = (tbl["review_count"] / tbl["review_count"].sum()).round(4)
    tbl = tbl.sort_values("review_score").reset_index(drop=True)
    return tbl


def build_payment_distribution(payments, valid_order_ids):
    """
    支付方式分布（delivered 口径，按支付明细行统计）。
    注意：一个订单可有多条支付行（如代金券 voucher + 信用卡组合支付），
    因此本表是“支付笔数/金额”口径，不是订单数口径，看板中需注明。
    """
    pay = payments[payments["order_id"].isin(valid_order_ids)]
    tbl = (
        pay.groupby("payment_type")
        .agg(payment_count=("payment_value", "size"),
             payment_value=("payment_value", "sum"))
        .reset_index()
    )
    tbl["payment_value"] = tbl["payment_value"].round(2)
    tbl["count_ratio"] = (tbl["payment_count"] / tbl["payment_count"].sum()).round(4)
    tbl = tbl.sort_values("payment_count", ascending=False).reset_index(drop=True)
    return tbl


def build_order_funnel(dfs, orders_clean):
    """
    订单状态粗略漏斗：订单创建 -> 有支付记录 -> 成功交付。
    数据无浏览/点击/加购日志，无法做完整营销漏斗（已在 README 标注局限）。
      stage      ：阶段名（带序号，便于看板排序）
      order_count：该阶段订单数
      stage_rate ：相对“订单创建”的转化率
    """
    total_created = len(dfs["orders"])
    total_paid = int(dfs["payments"]["order_id"].nunique())
    total_delivered = len(orders_clean)
    tbl = pd.DataFrame([
        {"stage": "1_订单创建", "order_count": total_created, "stage_rate": 1.0},
        {"stage": "2_有支付记录", "order_count": total_paid,
         "stage_rate": round(total_paid / total_created, 4)},
        {"stage": "3_成功交付", "order_count": total_delivered,
         "stage_rate": round(total_delivered / total_created, 4)},
    ])
    return tbl


def print_funnel(dfs, orders_clean):
    """
    粗略转化漏斗（数据局限说明：Olist 无浏览/点击/加购行为数据，无法做完整营销漏斗）。
    仅能计算订单状态层面的“成交率 = delivered 订单 / 全部创建订单”和支付覆盖率。
    """
    total_created = len(dfs["orders"])
    total_delivered = len(orders_clean)
    total_paid = dfs["payments"]["order_id"].nunique()
    print("\n" + "=" * 90)
    print("【粗略转化（受数据限制，非完整漏斗）】")
    print(f"    全部创建订单 : {total_created:,}")
    print(f"    有支付记录   : {total_paid:,}  支付覆盖率 {total_paid / total_created:.4%}")
    print(f"    成功交付     : {total_delivered:,}  成交率(delivered/创建) "
          f"{total_delivered / total_created:.4%}")
    print("    说明：数据集无浏览/加购/点击数据，无法还原“曝光-点击-加购-下单”完整漏斗，")
    print("          简历中应如实标注该局限，看板中不呈现虚构漏斗。")


# ======================================================================
# 主流程
# ======================================================================
def main():
    """主流程：加载 -> 探索 -> 清洗 -> 组装宽表 -> 指标计算 -> 输出 CSV。"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1) 加载与探索
    dfs = load_data()
    explore_data(dfs)
    check_key_integrity(dfs)

    # 2) 清洗各表
    orders_clean = clean_orders(dfs["orders"])
    valid_order_ids = set(orders_clean["order_id"])
    items_clean = clean_items(dfs["items"], valid_order_ids)
    products_clean = clean_products(dfs["products"], dfs["translation"])
    pay_order = aggregate_payments(dfs["payments"], valid_order_ids)
    rv_order = aggregate_reviews(dfs["reviews"], valid_order_ids)

    # 3) 组装宽表
    order_wide = build_order_wide(orders_clean, dfs["customers"], pay_order, rv_order)
    item_wide = build_item_wide(items_clean, products_clean, rv_order)

    # 4) 指标计算
    overview = build_overview(order_wide)
    daily_sales = build_daily_sales(order_wide)
    monthly_sales = build_monthly_sales(order_wide)
    # 基准日取“数据中最后一个订单日期”（全量 orders 口径，含未交付订单），与需求一致
    snapshot_date = pd.to_datetime(dfs["orders"]["order_purchase_timestamp"]).max()
    rfm = build_rfm(order_wide, snapshot_date)
    category_sales = build_category_sales(item_wide)
    review_analysis = build_review_analysis(item_wide)
    rating_dist = build_rating_distribution(dfs["reviews"], valid_order_ids)
    payment_dist = build_payment_distribution(dfs["payments"], valid_order_ids)
    funnel = build_order_funnel(dfs, orders_clean)
    print_funnel(dfs, orders_clean)

    # 打印两张分布表，便于控制台留痕
    print("\n" + "=" * 90)
    print("【评价星级分布】")
    for _, r in rating_dist.iterrows():
        print(f"    {int(r['review_score'])} 星: {int(r['review_count']):,} 条 ({r['ratio']:.2%})")
    print("【支付方式分布（支付笔数口径）】")
    for _, r in payment_dist.iterrows():
        print(f"    {r['payment_type']:12s}: {int(r['payment_count']):,} 笔 "
              f"({r['count_ratio']:.2%}), 金额 {r['payment_value']:,.2f}")

    # 5) 输出 CSV（utf-8-sig：带 BOM 的 UTF-8，Excel 直接打开中文不乱码，
    #    Tableau Public / Power BI / FineBI 均可正常识别）
    outputs = {
        "overview.csv": overview,
        "daily_sales.csv": daily_sales,
        "monthly_sales.csv": monthly_sales,
        "rfm.csv": rfm,
        "category_sales.csv": category_sales,
        "review_analysis.csv": review_analysis,
        "rating_distribution.csv": rating_dist,
        "payment_distribution.csv": payment_dist,
        "order_funnel.csv": funnel,
    }
    print("\n" + "=" * 90)
    print("【输出 CSV】")
    for fn, df in outputs.items():
        path = os.path.join(OUTPUT_DIR, fn)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"    {path:32s} {len(df):,} 行 × {df.shape[1]} 列")

    print("\n全部处理完成。")


if __name__ == "__main__":
    main()

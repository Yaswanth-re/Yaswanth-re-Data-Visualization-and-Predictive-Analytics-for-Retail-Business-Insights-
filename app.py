import streamlit as st

# Must remain the first Streamlit command.
st.set_page_config(page_title="Data Visualization and Predictive Analytics for Retail Business Insights", layout="wide")

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from prophet import Prophet
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from sklearn.ensemble import IsolationForest
from sklearn.cluster import KMeans
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error
from statsmodels.tsa.arima.model import ARIMA


def login() -> bool:
    st.sidebar.title("Login")
    username = st.sidebar.text_input("Username")
    password = st.sidebar.text_input("Password", type="password")

    if username == "admin" and password == "1234":
        return True
    if username and password:
        st.sidebar.error("Invalid credentials")
    return False


if not login():
    st.stop()

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&family=Source+Sans+3:wght@400;500;600&display=swap');

    :root {
        --bg-soft: #f6f1e8;
        --paper: #fffdf8;
        --ink: #1d2a38;
        --muted: #62707e;
        --accent: #0f766e;
        --accent-2: #c97b2d;
        --line: rgba(29, 42, 56, 0.12);
        --shadow: 0 20px 45px rgba(29, 42, 56, 0.08);
    }

    .stApp {
        background:
            radial-gradient(circle at top right, rgba(201, 123, 45, 0.10), transparent 24%),
            radial-gradient(circle at top left, rgba(15, 118, 110, 0.12), transparent 28%),
            linear-gradient(180deg, #f8f4ec 0%, #f3efe6 45%, #fbfaf7 100%);
    }

    html, body, [class*="css"] {
        font-family: "Source Sans 3", sans-serif;
        color: var(--ink);
    }

    h1, h2, h3, h4 {
        font-family: "Manrope", sans-serif !important;
        letter-spacing: -0.02em;
    }

    .block-container {
        padding-top: 1rem;
        padding-bottom: 2.4rem;
        max-width: 1480px;
    }

    .hero-shell {
        position: relative;
        overflow: hidden;
        background: linear-gradient(135deg, rgba(255,253,248,0.95), rgba(245,239,229,0.9));
        border: 1px solid var(--line);
        border-radius: 28px;
        padding: 30px 30px 24px 30px;
        box-shadow: var(--shadow);
        animation: fadeUp 560ms ease-out;
        margin-bottom: 1rem;
    }

    .hero-shell::after {
        content: "";
        position: absolute;
        right: -40px;
        top: -40px;
        width: 180px;
        height: 180px;
        border-radius: 999px;
        background: radial-gradient(circle, rgba(15,118,110,0.16), rgba(15,118,110,0.02) 68%, transparent 72%);
        animation: drift 11s ease-in-out infinite;
    }

    .hero-kicker {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 999px;
        background: rgba(15, 118, 110, 0.10);
        color: var(--accent);
        font-weight: 700;
        font-size: 0.88rem;
        margin-bottom: 14px;
    }

    .hero-title {
        font-family: "Manrope", sans-serif;
        font-size: clamp(2rem, 4vw, 3rem);
        font-weight: 800;
        line-height: 1.04;
        margin: 0 0 12px 0;
        color: var(--ink);
        max-width: 900px;
    }

    .hero-copy {
        max-width: 840px;
        color: var(--muted);
        font-size: 1.03rem;
        line-height: 1.6;
        margin-bottom: 18px;
    }

    .hero-guides {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px;
        margin-top: 12px;
    }

    .guide-chip {
        background: rgba(255,255,255,0.76);
        border: 1px solid var(--line);
        border-radius: 16px;
        padding: 12px 14px;
        box-shadow: 0 8px 24px rgba(29, 42, 56, 0.05);
        animation: fadeUp 650ms ease-out;
    }

    .guide-chip strong {
        display: block;
        color: var(--ink);
        font-family: "Manrope", sans-serif;
        margin-bottom: 2px;
    }

    .guide-chip span {
        color: var(--muted);
        font-size: 0.96rem;
    }

    [data-testid="stMetric"] {
        background: rgba(255, 253, 248, 0.86);
        border: 1px solid var(--line);
        border-radius: 20px;
        padding: 14px 16px;
        box-shadow: 0 12px 28px rgba(29, 42, 56, 0.05);
        backdrop-filter: blur(8px);
        min-height: 122px;
    }

    [data-testid="stMetricLabel"] {
        font-weight: 700;
        color: var(--muted);
    }

    [data-testid="stMetricValue"] {
        font-family: "Manrope", sans-serif;
        color: var(--ink);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: rgba(255,255,255,0.55);
        padding: 8px;
        border: 1px solid var(--line);
        border-radius: 18px;
        box-shadow: 0 12px 24px rgba(29, 42, 56, 0.05);
    }

    .stTabs [data-baseweb="tab"] {
        height: 44px;
        border-radius: 12px;
        padding-left: 16px;
        padding-right: 16px;
        color: var(--muted);
        font-weight: 700;
        transition: all 180ms ease;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, rgba(15,118,110,0.12), rgba(201,123,45,0.12));
        color: var(--ink) !important;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(255,253,248,0.98), rgba(244,239,230,0.98));
        border-right: 1px solid var(--line);
    }

    .stAlert, [data-testid="stExpander"] {
        border-radius: 18px;
    }

    [data-testid="stDataFrame"], .js-plotly-plot, [data-testid="stTable"] {
        background: rgba(255, 253, 248, 0.72);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 6px;
        box-shadow: 0 10px 26px rgba(29, 42, 56, 0.04);
    }

    .stMarkdown p {
        font-size: 1rem;
        line-height: 1.62;
    }

    @keyframes fadeUp {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
    }

    @keyframes drift {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(8px); }
    }

    @media (max-width: 900px) {
        .block-container {
            padding-top: 0.8rem;
        }

        .hero-shell {
            padding: 22px 18px 18px 18px;
            border-radius: 22px;
        }

        .hero-copy {
            font-size: 0.98rem;
        }

        [data-testid="stMetric"] {
            min-height: auto;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <section class="hero-shell">
        <div class="hero-kicker">Retail Analytics Dashboard</div>
        <h1 class="hero-title">Data Visualization and Predictive Analytics for Retail Business Insights</h1>
        <p class="hero-copy">
            This dashboard explains retail business performance in a way that both technical and non-technical viewers can understand.
            It shows what happened, what is changing, what may happen next, and what action the business should consider.
        </p>
        <div class="hero-guides">
            <div class="guide-chip">
                <strong>1. Understand</strong>
                <span>See sales, profit, trends, and top business areas.</span>
            </div>
            <div class="guide-chip">
                <strong>2. Predict</strong>
                <span>Compare forecasting models and estimate future performance.</span>
            </div>
            <div class="guide-chip">
                <strong>3. Decide</strong>
                <span>Use alerts, simulations, and insights for clearer action.</span>
            </div>
        </div>
    </section>
    """,
    unsafe_allow_html=True,
)

with st.expander("How to read this dashboard"):
    st.write("1. Use the filters in the left sidebar to choose the data you want to study.")
    st.write("2. Start with the Overview tab to understand sales or profit at a glance.")
    st.write("3. Open Forecasting to see what may happen next and which model performed best.")
    st.write("4. Open Operations and Insights to understand recommendations in plain language.")
    st.write("5. Use Export to download the filtered data and reports.")


@st.cache_data
def build_demo_dataset() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    dates = pd.date_range("2020-01-01", "2025-12-31", freq="D")
    categories = ["Furniture", "Office Supplies", "Technology", "Clothing", "Appliances", "Food & Beverages"]
    regions = ["West", "East", "Central", "South"]
    segments = ["Consumer", "Corporate", "Home Office", "Small Business", "Online", "Wholesale"]
    sub_categories = {
        "Furniture": ["Chairs", "Tables", "Bookcases"],
        "Office Supplies": ["Binders", "Paper", "Storage"],
        "Technology": ["Phones", "Accessories", "Machines"],
        "Clothing": ["Shirts", "Footwear", "Jackets"],
        "Appliances": ["Mixers", "Microwaves", "Refrigerators"],
        "Food & Beverages": ["Snacks", "Beverages", "Packaged Food"],
    }

    rows = []
    order_id = 10000
    customer_ids = [f"CUST-{idx:04d}" for idx in range(1, 601)]
    cities = ["New York", "Chicago", "Los Angeles", "Houston", "Seattle", "Boston", "Phoenix", "Atlanta", "Dallas", "San Diego"]
    states = ["NY", "IL", "CA", "TX", "WA", "MA", "AZ", "GA", "TX", "CA"]
    review_pool = [
        "Fast delivery and excellent experience",
        "Very happy with product quality",
        "Average experience but acceptable value",
        "Packaging was damaged and service was slow",
        "Great price and easy to recommend",
        "Support response was poor and disappointing",
    ]
    for current_date in dates:
        order_count = int(rng.integers(10, 22))
        seasonal_factor = 1 + 0.18 * np.sin((current_date.dayofyear / 365) * 2 * np.pi)
        for _ in range(order_count):
            category = rng.choice(categories, p=[0.18, 0.20, 0.18, 0.16, 0.14, 0.14])
            sub_category = rng.choice(sub_categories[category])
            region = rng.choice(regions)
            segment = rng.choice(segments)
            customer_id = rng.choice(customer_ids)
            city_index = int(rng.integers(0, len(cities)))
            quantity = int(rng.integers(1, 8))
            discount = float(rng.choice([0.0, 0.1, 0.2, 0.3], p=[0.45, 0.25, 0.2, 0.1]))
            base_price = {
                "Furniture": 220,
                "Office Supplies": 55,
                "Technology": 330,
                "Clothing": 95,
                "Appliances": 410,
                "Food & Beverages": 35,
            }[category]
            sales = max(10, rng.normal(base_price, base_price * 0.35)) * quantity * seasonal_factor * (1 - discount)
            profit = sales * rng.uniform(0.08, 0.24) - discount * sales * rng.uniform(0.05, 0.12)
            rows.append(
                {
                    "Order ID": f"CA-{order_id}",
                    "Customer ID": customer_id,
                    "Order Date": current_date,
                    "Category": category,
                    "Sub-Category": sub_category,
                    "Region": region,
                    "City": cities[city_index],
                    "State": states[city_index],
                    "Store": f"{region} Hub {city_index + 1}",
                    "Segment": segment,
                    "Sales": round(sales, 2),
                    "Profit": round(profit, 2),
                    "Quantity": quantity,
                    "Discount": discount,
                    "Review": rng.choice(review_pool),
                }
            )
            order_id += 1

    return pd.DataFrame(rows)


@st.cache_data
def load_data(uploaded_file) -> pd.DataFrame:
    if uploaded_file is not None:
        return pd.read_csv(uploaded_file, encoding="latin1")

    default_path = "Sample - Superstore.csv"
    try:
        return pd.read_csv(default_path, encoding="latin1")
    except FileNotFoundError:
        return build_demo_dataset()


def safe_mape(actual: np.ndarray, predicted: np.ndarray) -> float:
    actual = np.array(actual, dtype=float)
    predicted = np.array(predicted, dtype=float)
    mask = actual != 0
    if not mask.any():
        return 0.0
    return float(np.mean(np.abs((actual[mask] - predicted[mask]) / actual[mask])) * 100)


def score_sentiment(text: str) -> str:
    positive_words = {"great", "excellent", "happy", "fast", "good", "recommend", "easy", "quality"}
    negative_words = {"poor", "slow", "damaged", "disappointing", "bad", "late", "issue", "problem"}
    text_lower = str(text).lower()
    positive_score = sum(word in text_lower for word in positive_words)
    negative_score = sum(word in text_lower for word in negative_words)
    if positive_score > negative_score:
        return "Positive"
    if negative_score > positive_score:
        return "Negative"
    return "Neutral"


def ensure_supporting_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Segment" not in df.columns:
        df["Segment"] = "General"
    if "Quantity" not in df.columns:
        df["Quantity"] = 1
    if "Discount" not in df.columns:
        df["Discount"] = 0.0
    if "Customer ID" not in df.columns:
        df["Customer ID"] = [f"CUST-{idx % 300:03d}" for idx in range(len(df))]
    if "Order ID" not in df.columns:
        df["Order ID"] = [f"ORD-{idx:05d}" for idx in range(len(df))]
    if "Store" not in df.columns:
        if "City" in df.columns:
            df["Store"] = df["City"].astype(str) + " Store"
        else:
            df["Store"] = df["Region"].astype(str) + " Store"
    if "Review" not in df.columns:
        df["Review"] = "Average experience"
    if "Sub-Category" not in df.columns:
        df["Sub-Category"] = df["Category"]
    return df


def build_customer_segments(filtered: pd.DataFrame) -> pd.DataFrame:
    customer_summary = (
        filtered.groupby("Customer ID")
        .agg(
            Monetary=("Sales", "sum"),
            Frequency=("Order ID", "nunique"),
            LastOrder=("Order Date", "max"),
        )
        .reset_index()
    )
    customer_summary["RecencyDays"] = (filtered["Order Date"].max() - customer_summary["LastOrder"]).dt.days
    feature_frame = customer_summary[["Monetary", "Frequency", "RecencyDays"]].copy()
    if len(customer_summary) >= 3:
        model = KMeans(n_clusters=3, n_init=10, random_state=42)
        customer_summary["SegmentId"] = model.fit_predict(feature_frame)
    else:
        customer_summary["SegmentId"] = 0
    segment_order = customer_summary.groupby("SegmentId")["Monetary"].mean().sort_values().index.tolist()
    segment_labels = {}
    label_names = ["At Risk", "Regular", "High Value"]
    for idx, segment_id in enumerate(segment_order):
        segment_labels[segment_id] = label_names[min(idx, len(label_names) - 1)]
    customer_summary["Customer Segment"] = customer_summary["SegmentId"].map(segment_labels)
    return customer_summary


def build_basket_pairs(filtered: pd.DataFrame) -> pd.DataFrame:
    if "Order ID" not in filtered.columns or "Sub-Category" not in filtered.columns:
        return pd.DataFrame(columns=["Product Pair", "Orders Together"])
    pair_counts = {}
    grouped = filtered.groupby("Order ID")["Sub-Category"].apply(lambda s: sorted(set(s.dropna().astype(str))))
    for items in grouped:
        for index, left_item in enumerate(items):
            for right_item in items[index + 1:]:
                pair = f"{left_item} + {right_item}"
                pair_counts[pair] = pair_counts.get(pair, 0) + 1
    if not pair_counts:
        return pd.DataFrame(columns=["Product Pair", "Orders Together"])
    return (
        pd.DataFrame({"Product Pair": list(pair_counts.keys()), "Orders Together": list(pair_counts.values())})
        .sort_values("Orders Together", ascending=False)
        .head(10)
    )


def build_export_package(
    performance_df: pd.DataFrame,
    best_forecast: pd.DataFrame,
    anomalies: pd.DataFrame,
    customer_segments: pd.DataFrame,
) -> bytes:
    import io

    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        performance_df.to_excel(writer, sheet_name="Model Performance", index=False)
        best_forecast.to_excel(writer, sheet_name="Best Forecast", index=False)
        anomalies.to_excel(writer, sheet_name="Anomalies", index=False)
        customer_segments.to_excel(writer, sheet_name="Customer Segments", index=False)
    output.seek(0)
    return output.getvalue()


def generate_business_insights(
    monthly: pd.DataFrame,
    region_ranking: pd.DataFrame,
    anomalies: pd.DataFrame,
    target: str,
    growth: float,
    volatility: float,
    projected_change: float,
) -> list[str]:
    insights = []
    recent_values = monthly["y"].tail(4).tolist()
    if len(recent_values) >= 4 and recent_values[-1] < recent_values[-2] < recent_values[-3]:
        insights.append(f"{target} has been declining for the last 3 months, which signals a short-term slowdown.")
    elif len(recent_values) >= 4 and recent_values[-1] > recent_values[-2] > recent_values[-3]:
        insights.append(f"{target} has increased for the last 3 months, showing positive momentum.")

    if growth > 0:
        insights.append(f"Month-over-month {target.lower()} growth is {growth:.2f}%, indicating current expansion.")
    elif growth < 0:
        insights.append(f"Month-over-month {target.lower()} growth is {growth:.2f}%, which indicates contraction.")

    if projected_change > 0:
        insights.append(f"The forecast suggests a projected {target.lower()} increase of {projected_change:.2f}% over the recent average.")
    elif projected_change < 0:
        insights.append(f"The forecast suggests a projected {target.lower()} decrease of {abs(projected_change):.2f}% over the recent average.")

    if not region_ranking.empty:
        top_region = region_ranking.iloc[0]
        insights.append(f"{top_region['Region']} region currently contributes the highest {target.lower()} at {top_region[target]:,.2f}.")

    if volatility >= 20:
        insights.append(f"{target} volatility is high at {volatility:.2f}%, so planning should include larger safety buffers.")
    elif volatility <= 8:
        insights.append(f"{target} volatility is low at {volatility:.2f}%, suggesting stable recent performance.")

    if len(anomalies) > 0:
        latest_anomaly = anomalies["ds"].max().date()
        insights.append(f"{len(anomalies)} anomaly month(s) were detected; the latest anomaly occurred in {latest_anomaly}.")
    else:
        insights.append("No anomaly months were detected in the current filtered view.")

    return insights


def build_pdf_report(
    title: str,
    target: str,
    kpis: list[list[str]],
    monthly: pd.DataFrame,
    region_ranking: pd.DataFrame,
    performance_df: pd.DataFrame,
    insights: list[str],
) -> bytes:
    import io

    output = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    story = [
        Paragraph(title, styles["Title"]),
        Spacer(1, 0.15 * inch),
        Paragraph("Retail Analytics PDF Report", styles["Heading2"]),
        Spacer(1, 0.12 * inch),
    ]

    kpi_table = Table([["KPI", "Value"]] + kpis, colWidths=[2.8 * inch, 2.6 * inch])
    kpi_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3c88")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#eef3ff")]),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend([Paragraph("Key KPIs", styles["Heading2"]), kpi_table, Spacer(1, 0.18 * inch)])

    line_drawing = Drawing(500, 190)
    line_chart = HorizontalLineChart()
    line_chart.x = 40
    line_chart.y = 30
    line_chart.height = 120
    line_chart.width = 420
    line_chart.data = [tuple(monthly["y"].tail(8).tolist())]
    line_chart.categoryAxis.categoryNames = [d.strftime("%b %Y") for d in monthly["ds"].tail(8)]
    line_chart.categoryAxis.labels.angle = 30
    line_chart.valueAxis.valueMin = 0
    line_chart.lines[0].strokeColor = colors.HexColor("#2a6f97")
    line_chart.lines[0].strokeWidth = 2
    line_drawing.add(String(40, 165, f"Recent {target} Trend", fontName="Helvetica-Bold", fontSize=12))
    line_drawing.add(line_chart)
    story.extend([Paragraph("Charts", styles["Heading2"]), line_drawing, Spacer(1, 0.18 * inch)])

    bar_drawing = Drawing(500, 220)
    bar_chart = VerticalBarChart()
    bar_chart.x = 45
    bar_chart.y = 35
    bar_chart.height = 130
    bar_chart.width = 400
    top_regions = region_ranking.head(4)
    bar_chart.data = [tuple(top_regions[target].tolist())] if not top_regions.empty else [(0,)]
    bar_chart.categoryAxis.categoryNames = top_regions["Region"].tolist() if not top_regions.empty else ["N/A"]
    bar_chart.valueAxis.valueMin = 0
    bar_chart.bars[0].fillColor = colors.HexColor("#3a7d44")
    bar_drawing.add(String(45, 180, f"Top Region {target}", fontName="Helvetica-Bold", fontSize=12))
    bar_drawing.add(bar_chart)
    story.extend([bar_drawing, Spacer(1, 0.18 * inch)])

    performance_table = Table(
        [["Model", "MAE", "RMSE", "MAPE %"]]
        + performance_df.round(2).astype({"Model": str}).values.tolist(),
        colWidths=[1.6 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch],
    )
    performance_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#5c677d")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f4f6")]),
                ("PADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.extend([Paragraph("Model Performance", styles["Heading2"]), performance_table, Spacer(1, 0.18 * inch)])

    story.append(Paragraph("Automated Insights", styles["Heading2"]))
    for insight in insights:
        story.append(Paragraph(f"- {insight}", styles["BodyText"]))
        story.append(Spacer(1, 0.08 * inch))

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def compute_regional_growth(filtered: pd.DataFrame, target: str) -> pd.DataFrame:
    regional_monthly = (
        filtered.groupby(["Region", pd.Grouper(key="Order Date", freq="ME")])[target]
        .sum()
        .reset_index()
        .sort_values(["Region", "Order Date"])
    )
    if regional_monthly.empty:
        return pd.DataFrame(columns=["Region", "Previous", "Current", "Growth %"])
    regional_latest = regional_monthly.groupby("Region").tail(2).copy()
    growth_rows = []
    for region_name, group in regional_latest.groupby("Region"):
        values = group[target].tolist()
        if len(values) == 1:
            previous, current = 0.0, values[0]
        else:
            previous, current = values[-2], values[-1]
        growth_pct = ((current - previous) / previous * 100) if previous else 0.0
        growth_rows.append(
            {
                "Region": region_name,
                "Previous": previous,
                "Current": current,
                "Growth %": growth_pct,
            }
        )
    return pd.DataFrame(growth_rows).sort_values("Growth %", ascending=False).reset_index(drop=True)


def prophet_model(train_df: pd.DataFrame, forecast_period: int) -> pd.DataFrame:
    model = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False)
    model.fit(train_df)
    future = model.make_future_dataframe(periods=forecast_period, freq="ME")
    forecast = model.predict(future)
    return forecast[["ds", "yhat", "yhat_lower", "yhat_upper"]]


def linear_model(train_df: pd.DataFrame, forecast_period: int) -> pd.DataFrame:
    train_copy = train_df.copy()
    train_copy["time"] = np.arange(len(train_copy))

    model = LinearRegression()
    model.fit(train_copy[["time"]].values, train_copy["y"].values)

    future_time = np.arange(len(train_copy), len(train_copy) + forecast_period).reshape(-1, 1)
    preds = model.predict(future_time)
    future_dates = pd.date_range(start=train_copy["ds"].iloc[-1], periods=forecast_period + 1, freq="ME")[1:]

    return pd.DataFrame(
        {
            "ds": future_dates,
            "yhat": preds,
            "yhat_lower": preds,
            "yhat_upper": preds,
        }
    )


def arima_model(train_df: pd.DataFrame, forecast_period: int) -> pd.DataFrame:
    model = ARIMA(train_df["y"], order=(1, 1, 1))
    fit = model.fit()
    forecast = fit.forecast(steps=forecast_period)

    future_dates = pd.date_range(start=train_df["ds"].iloc[-1], periods=forecast_period + 1, freq="ME")[1:]
    return pd.DataFrame(
        {
            "ds": future_dates,
            "yhat": forecast.values,
            "yhat_lower": forecast.values,
            "yhat_upper": forecast.values,
        }
    )


def train_and_score_models(monthly: pd.DataFrame, forecast_period: int) -> tuple[dict, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    train = monthly.iloc[:-forecast_period].copy()
    test = monthly.iloc[-forecast_period:].copy()

    prophet_forecast = prophet_model(train, forecast_period)
    linear_forecast = linear_model(train, forecast_period)
    arima_forecast = arima_model(train, forecast_period)

    scored_models = {
        "Prophet": prophet_forecast.tail(forecast_period).reset_index(drop=True),
        "Linear": linear_forecast.reset_index(drop=True),
        "ARIMA": arima_forecast.reset_index(drop=True),
    }

    metrics = []
    for model_name, forecast_df in scored_models.items():
        actual = test["y"].values
        predicted = forecast_df["yhat"].values
        metrics.append(
            {
                "Model": model_name,
                "MAE": mean_absolute_error(actual, predicted),
                "RMSE": float(np.sqrt(mean_squared_error(actual, predicted))),
                "MAPE %": safe_mape(actual, predicted),
            }
        )

    performance_df = pd.DataFrame(metrics).sort_values(by="MAE").reset_index(drop=True)
    return scored_models, performance_df, train, test


uploaded_file = st.sidebar.file_uploader("Upload retail CSV", type=["csv"])
df = load_data(uploaded_file)
df["Order Date"] = pd.to_datetime(df["Order Date"])
df = ensure_supporting_columns(df)

required_columns = {"Order Date", "Category", "Region", "Sales", "Profit"}
missing_columns = required_columns - set(df.columns)
if missing_columns:
    st.error(f"Dataset is missing required columns: {', '.join(sorted(missing_columns))}")
    st.stop()

df = df.sort_values("Order Date").copy()

st.sidebar.header("Filters")
st.sidebar.caption("Choose the business view you want to understand.")
min_date = df["Order Date"].min().date()
max_date = df["Order Date"].max().date()
date_range = st.sidebar.date_input("Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
if not isinstance(date_range, tuple) or len(date_range) != 2:
    st.warning("Please select a start and end date.")
    st.stop()

category_options = ["All"] + sorted(df["Category"].dropna().unique().tolist())
region_options = ["All"] + sorted(df["Region"].dropna().unique().tolist())
segment_options = ["All"] + sorted(df["Segment"].dropna().unique().tolist()) if "Segment" in df.columns else ["All"]

category = st.sidebar.selectbox("Category", category_options)
region = st.sidebar.selectbox("Region", region_options)
segment = st.sidebar.selectbox("Segment", segment_options)
target = st.sidebar.selectbox("Target Variable", ["Sales", "Profit"])
forecast_period = st.sidebar.slider("Forecast Months", 1, 12, 6)
selected_models = st.sidebar.multiselect("Models to Display", ["Prophet", "Linear", "ARIMA"], default=["Prophet", "Linear", "ARIMA"])
monthly_target_value = st.sidebar.number_input("Monthly Target", min_value=0.0, value=50000.0, step=5000.0)
price_increase_pct = st.sidebar.slider("Price Increase %", -20, 30, 0)
discount_change_pct = st.sidebar.slider("What-If Discount Change %", -20, 20, 0)
demand_change_pct = st.sidebar.slider("What-If Demand Change %", -30, 30, 0)
stock_buffer_pct = st.sidebar.slider("Stock Buffer %", 0, 50, 20)

with st.sidebar.expander("Meaning of these controls"):
    st.write("Target Variable: choose whether you want to study Sales or Profit.")
    st.write("Forecast Months: how many future months the dashboard should predict.")
    st.write("Price Increase %: tests how revenue may change if prices go up or down.")
    st.write("Demand Change %: tests how revenue may change if customer demand rises or falls.")
    st.write("Stock Buffer %: adds extra safety stock for uncertain demand.")

start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
filtered = df[(df["Order Date"] >= start_date) & (df["Order Date"] <= end_date)].copy()
if category != "All":
    filtered = filtered[filtered["Category"] == category]
if region != "All":
    filtered = filtered[filtered["Region"] == region]
if segment != "All" and "Segment" in filtered.columns:
    filtered = filtered[filtered["Segment"] == segment]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

monthly = (
    filtered.groupby(pd.Grouper(key="Order Date", freq="ME"))[target]
    .sum()
    .reset_index()
    .rename(columns={"Order Date": "ds", target: "y"})
)
monthly = monthly[monthly["y"].notna()].copy()

if len(monthly) < max(12, forecast_period + 3):
    st.warning("Not enough monthly data for reliable forecasting. Try widening the filters or date range.")
    st.stop()

model_outputs, performance_df, train, test = train_and_score_models(monthly, forecast_period)
best_model = performance_df.iloc[0]["Model"]
best_forecast = model_outputs[best_model]
best_forecast_full = best_forecast.copy()

growth = 0.0
if len(monthly) > 1 and monthly["y"].iloc[-2] != 0:
    growth = ((monthly["y"].iloc[-1] - monthly["y"].iloc[-2]) / monthly["y"].iloc[-2]) * 100

volatility = monthly["y"].pct_change().dropna().std() * 100 if len(monthly) > 2 else 0.0
recent_avg = monthly["y"].tail(3).mean()
forecast_avg = best_forecast["yhat"].mean()
projected_change = ((forecast_avg - recent_avg) / recent_avg) * 100 if recent_avg else 0.0
seasonality_view = (
    filtered.assign(MonthName=filtered["Order Date"].dt.strftime("%b"))
    .groupby("MonthName")[target]
    .sum()
    .reindex(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    .fillna(0)
    .reset_index()
)

customer_segments = build_customer_segments(filtered)
basket_pairs = build_basket_pairs(filtered)
filtered["Sentiment"] = filtered["Review"].apply(score_sentiment)

scenario_multiplier = (1 + demand_change_pct / 100) * (1 + price_increase_pct / 100) * (1 - discount_change_pct / 100)
scenario_forecast = best_forecast_full.copy()
scenario_forecast["Scenario Forecast"] = scenario_forecast["yhat"] * scenario_multiplier
current_revenue = float(monthly["y"].sum())
simulated_revenue = float(current_revenue * scenario_multiplier)
scenario_comparison = pd.DataFrame(
    {
        "Scenario": ["Current Revenue", "Simulated Revenue"],
        "Revenue": [current_revenue, simulated_revenue],
    }
)
region_ranking = (
    filtered.groupby("Region")[[target]]
    .sum()
    .sort_values(by=target, ascending=False)
    .reset_index()
)
regional_growth = compute_regional_growth(filtered, target)
store_metric_columns = list(dict.fromkeys([target, "Profit", "Sales"]))
store_ranking = filtered.groupby("Store")[store_metric_columns].sum().sort_values(by=target, ascending=False).reset_index()
category_summary = (
    filtered.groupby("Category")[["Sales", "Profit", "Quantity"]]
    .sum()
    .reset_index()
    .sort_values(by=target, ascending=False)
)
segment_summary_detail = (
    filtered.groupby("Segment")[["Sales", "Profit", "Quantity"]]
    .sum()
    .reset_index()
    .sort_values(by=target, ascending=False)
)
profitability_matrix = (
    filtered.assign(Month=filtered["Order Date"].dt.strftime("%b"))
    .pivot_table(index="Region", columns="Month", values="Profit", aggfunc="sum", fill_value=0)
)
profitability_matrix = profitability_matrix.reindex(
    columns=["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    fill_value=0,
)
discount_analysis = (
    filtered.assign(DiscountBand=pd.cut(filtered["Discount"], bins=[-0.01, 0.0, 0.1, 0.2, 0.3, 1.0], labels=["0%", "1-10%", "11-20%", "21-30%", "30%+"]))
    .groupby("DiscountBand", observed=False)[["Sales", "Profit", "Quantity"]]
    .mean()
    .reset_index()
)
target_vs_actual = monthly.copy()
target_vs_actual["Target"] = monthly_target_value
target_vs_actual["Variance"] = target_vs_actual["y"] - target_vs_actual["Target"]
target_vs_actual["Status"] = np.where(target_vs_actual["Variance"] >= 0, "Above Target", "Below Target")

customer_segments["Churn Risk"] = np.where(
    customer_segments["RecencyDays"] > customer_segments["RecencyDays"].median() * 1.5,
    "High",
    "Low",
)
high_risk_customers = customer_segments.sort_values(["Churn Risk", "RecencyDays", "Monetary"], ascending=[False, False, False]).head(10)

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Average Monthly", f"${monthly['y'].mean():,.0f}")
col2.metric("Peak Monthly", f"${monthly['y'].max():,.0f}")
col3.metric("Growth %", f"{growth:.2f}%")
col4.metric("Best Model", best_model)
col5.metric("Volatility", f"{volatility:.2f}%")

with st.expander("Simple meaning of the top numbers"):
    st.write("Average Monthly: the average value per month in the selected data.")
    st.write("Peak Monthly: the highest month in the selected data.")
    st.write("Growth %: how much the latest month changed compared with the previous month.")
    st.write("Best Model: the forecasting method with the lowest prediction error.")
    st.write("Volatility: how unstable or changeable the monthly performance is.")

iso = IsolationForest(contamination=0.08, random_state=42)
anomaly_frame = monthly.copy()
anomaly_frame["anomaly"] = iso.fit_predict(anomaly_frame[["y"]])
anomaly_frame["label"] = np.where(anomaly_frame["anomaly"] == -1, "Anomaly", "Normal")
anomalies = anomaly_frame[anomaly_frame["anomaly"] == -1].copy()

best_mae = performance_df.iloc[0]["MAE"]
safety_multiplier = (1.10 if volatility < 10 else 1.20 if volatility < 20 else 1.30) * (1 + stock_buffer_pct / 100)
recommended_stock = max(0, forecast_avg * safety_multiplier)
stock_alert = "Low Risk"
if projected_change > 15 or volatility > 25:
    stock_alert = "High Reorder Risk"
elif projected_change > 5 or volatility > 15:
    stock_alert = "Medium Reorder Risk"

forecast_export = best_forecast.copy()
forecast_export["Selected Target"] = target
forecast_export["Best Model"] = best_model
forecast_export["Category Filter"] = category
forecast_export["Region Filter"] = region
forecast_export["Segment Filter"] = segment
forecast_export["Scenario Forecast"] = scenario_forecast["Scenario Forecast"].values

automated_insights = generate_business_insights(
    monthly=monthly,
    region_ranking=region_ranking,
    anomalies=anomalies,
    target=target,
    growth=growth,
    volatility=volatility,
    projected_change=projected_change,
)
if not regional_growth.empty:
    fastest_region = regional_growth.iloc[0]
    automated_insights.append(
        f"{fastest_region['Region']} region shows the highest recent growth at {fastest_region['Growth %']:.2f}%."
    )
automated_insights.append(
    f"A {price_increase_pct}% price change and {demand_change_pct}% demand change would move projected revenue from {current_revenue:,.2f} to {simulated_revenue:,.2f}."
)

excel_package = build_export_package(performance_df, forecast_export, anomalies, customer_segments)
text_report = "\n".join(
    [
        "Retail Analytics Executive Summary",
        f"Target variable: {target}",
        f"Best model: {best_model}",
        f"Forecast average: {forecast_avg:,.2f}",
        f"Projected change: {projected_change:.2f}%",
        f"Inventory recommendation: {recommended_stock:,.0f} units",
        f"Stock alert: {stock_alert}",
        "Automated insights:",
        *[f"- {insight}" for insight in automated_insights],
    ]
)
pdf_kpis = [
    ["Average Monthly", f"${monthly['y'].mean():,.0f}"],
    ["Peak Monthly", f"${monthly['y'].max():,.0f}"],
    ["Growth %", f"{growth:.2f}%"],
    ["Best Model", best_model],
    ["Volatility", f"{volatility:.2f}%"],
    ["Recommended Stock", f"{recommended_stock:,.0f} units"],
    ["Current Revenue", f"${current_revenue:,.0f}"],
    ["Simulated Revenue", f"${simulated_revenue:,.0f}"],
]
pdf_report = build_pdf_report(
    title="Data Visualization and Predictive Analytics for Retail Business Insights",
    target=target,
    kpis=pdf_kpis,
    monthly=monthly,
    region_ranking=region_ranking,
    performance_df=performance_df,
    insights=automated_insights,
)

overview_tab, customer_tab, product_tab, category_tab, operations_tab, insights_tab = st.tabs(
    ["Overview", "Customers", "Products", "Category & Segment", "Business Actions", "Smart Insights"]
)

with overview_tab:
    st.caption("Start here. This section explains the current business situation in the simplest way.")
    top_left, top_right = st.columns([2, 1])

    with top_left:
        st.subheader("Monthly Trend")
        st.write("This chart shows how the selected business measure changed month by month.")
        trend_fig = go.Figure()
        trend_fig.add_trace(go.Scatter(x=monthly["ds"], y=monthly["y"], mode="lines+markers", name="Monthly Actual"))
        trend_fig.add_trace(go.Scatter(x=target_vs_actual["ds"], y=target_vs_actual["Target"], mode="lines", name="Target"))
        trend_fig.update_layout(xaxis_title="Month", yaxis_title=target, height=420)
        st.plotly_chart(trend_fig, use_container_width=True)

    with top_right:
        st.subheader("Business Snapshot")
        st.write("These quick notes summarize the current condition of the business view you selected.")
        st.write(f"Projected average change vs last 3 months: **{projected_change:.2f}%**")
        st.write(f"Forecast horizon: **{forecast_period} months**")
        st.write(f"Records in current view: **{len(filtered):,}**")
        st.write(f"Date coverage: **{monthly['ds'].min().date()} to {monthly['ds'].max().date()}**")
        st.write(f"Inventory alert level: **{stock_alert}**")

    detail_left, detail_right = st.columns(2)

    with detail_left:
        subcategory_summary = (
            filtered.groupby("Sub-Category")[[target]]
            .sum()
            .sort_values(by=target, ascending=False)
            .head(10)
        )
        st.subheader("Top Sub-Categories")
        st.write("These are the strongest product groups inside the current filtered data.")
        st.dataframe(subcategory_summary, use_container_width=True)

    with detail_right:
        regional_summary = (
            filtered.groupby("Region")[[target]]
            .sum()
            .sort_values(by=target, ascending=False)
            .reset_index()
        )
        region_fig = go.Figure(data=[go.Bar(x=regional_summary["Region"], y=regional_summary[target], marker_color="#1f77b4")])
        region_fig.update_layout(title="Regional Contribution", yaxis_title=target, height=350)
        st.plotly_chart(region_fig, use_container_width=True)

    st.subheader("Seasonal Trend Analysis")
    st.write("This shows which months are usually stronger or weaker, helping with planning and promotion timing.")
    seasonal_fig = go.Figure(data=[go.Bar(x=seasonality_view["MonthName"], y=seasonality_view[target], marker_color="#2ca02c")])
    seasonal_fig.update_layout(height=320, yaxis_title=target)
    st.plotly_chart(seasonal_fig, use_container_width=True)

    st.subheader("Filtered Data Preview")
    st.write(
        f"This sample table shows the actual records behind the charts. "
        f"The current filtered view contains **{len(filtered):,} rows**, and the table below shows the first 100 rows."
    )
    st.dataframe(filtered.head(100), use_container_width=True)

with customer_tab:
    st.caption("This section explains customer behavior, customer quality, and customer risk in a simple way.")
    st.subheader("Customer Segmentation")
    st.write("Customers are grouped by buying value, buying frequency, and how recently they ordered.")
    segment_summary = (
        customer_segments.groupby("Customer Segment")[["Monetary", "Frequency", "RecencyDays"]]
        .mean()
        .round(2)
        .reset_index()
    )
    st.dataframe(segment_summary, use_container_width=True)

    segment_fig = go.Figure(
        data=[
            go.Scatter(
                x=customer_segments["Frequency"],
                y=customer_segments["Monetary"],
                mode="markers",
                marker=dict(
                    size=np.clip(customer_segments["RecencyDays"], 8, 24),
                    color=customer_segments["SegmentId"],
                    colorscale="Viridis",
                    showscale=True,
                ),
                text=customer_segments["Customer Segment"],
            )
        ]
    )
    segment_fig.update_layout(height=380, xaxis_title="Frequency", yaxis_title="Monetary")
    st.plotly_chart(segment_fig, use_container_width=True)

    st.subheader("Customer Churn Prediction")
    st.write("These are customers who may be at risk of not ordering again soon.")
    st.dataframe(
        high_risk_customers[["Customer ID", "Customer Segment", "RecencyDays", "Monetary", "Churn Risk"]],
        use_container_width=True,
    )

    sentiment_summary = filtered.groupby("Sentiment").size().reset_index(name="Count")
    st.subheader("Customer Sentiment / Review Analysis")
    st.write("This summarizes whether customer review language looks mostly positive, neutral, or negative.")
    sentiment_fig = go.Figure(data=[go.Pie(labels=sentiment_summary["Sentiment"], values=sentiment_summary["Count"], hole=0.45)])
    sentiment_fig.update_layout(height=320)
    st.plotly_chart(sentiment_fig, use_container_width=True)

with product_tab:
    st.caption("This section focuses on products, bundles, and the effect of discounts.")
    st.subheader("Product Recommendation Insights")
    st.write("These product pairs can help with cross-selling and bundle planning.")
    if basket_pairs.empty:
        st.info("Not enough multi-item order data is available to calculate product pair recommendations.")
    else:
        st.dataframe(basket_pairs, use_container_width=True)

    st.subheader("Discount Impact Analysis")
    st.write("This compares how different discount ranges affect average sales and average profit.")
    discount_fig = go.Figure()
    discount_fig.add_trace(go.Bar(x=discount_analysis["DiscountBand"], y=discount_analysis["Sales"], name="Avg Sales"))
    discount_fig.add_trace(go.Bar(x=discount_analysis["DiscountBand"], y=discount_analysis["Profit"], name="Avg Profit"))
    discount_fig.update_layout(barmode="group", height=360)
    st.plotly_chart(discount_fig, use_container_width=True)
    st.dataframe(discount_analysis, use_container_width=True)

with category_tab:
    st.caption("This section helps non-technical viewers compare which categories and customer segments are strongest.")
    st.subheader("Category Performance")
    st.write("Use this to compare sales and profit across major product categories.")
    category_fig = go.Figure()
    category_fig.add_trace(go.Bar(x=category_summary["Category"], y=category_summary["Sales"], name="Sales"))
    category_fig.add_trace(go.Bar(x=category_summary["Category"], y=category_summary["Profit"], name="Profit"))
    category_fig.update_layout(barmode="group", height=360)
    st.plotly_chart(category_fig, use_container_width=True)
    st.dataframe(category_summary, use_container_width=True)

    st.subheader("Segment Performance")
    st.write("Use this to compare performance across customer groups such as Consumer or Corporate.")
    segment_fig = go.Figure()
    segment_fig.add_trace(go.Bar(x=segment_summary_detail["Segment"], y=segment_summary_detail["Sales"], name="Sales"))
    segment_fig.add_trace(go.Bar(x=segment_summary_detail["Segment"], y=segment_summary_detail["Profit"], name="Profit"))
    segment_fig.update_layout(barmode="group", height=360)
    st.plotly_chart(segment_fig, use_container_width=True)
    st.dataframe(segment_summary_detail, use_container_width=True)

    st.subheader("Category vs Segment Pivot")
    st.write("This table shows which segment contributes most inside each category.")
    category_segment_pivot = filtered.pivot_table(
        index="Category",
        columns="Segment",
        values=target,
        aggfunc="sum",
        fill_value=0,
    )
    st.dataframe(category_segment_pivot, use_container_width=True)

with operations_tab:
    st.caption("This section turns analytics into practical business actions.")
    st.subheader("Inventory and Action Recommendations")
    insight_col1, insight_col2 = st.columns(2)
    with insight_col1:
        st.metric("Recommended Monthly Stock", f"{recommended_stock:,.0f} units")
        st.metric("Expected Forecast Average", f"${forecast_avg:,.0f}")
        st.metric("Best Model MAE", f"{best_mae:,.2f}")
        st.metric("Stock Alert", stock_alert)
        st.metric("Simulated Revenue", f"${simulated_revenue:,.0f}")

    with insight_col2:
        recommendation_lines = [
            f"Use **{best_model}** as the lead planning model for this filtered view.",
            f"Maintain a safety stock multiplier of **{safety_multiplier:.2f}x** including the selected stock buffer.",
            "Investigate anomaly months before locking procurement decisions.",
            "Use customer churn and sentiment sections to guide retention offers.",
        ]
        for line in recommendation_lines:
            st.write(f"- {line}")

    st.subheader("Store and Region Ranking")
    st.write("This helps identify the best-performing and weakest business locations quickly.")
    rank_left, rank_right = st.columns(2)
    with rank_left:
        st.dataframe(region_ranking, use_container_width=True)
    with rank_right:
        st.dataframe(store_ranking.head(10), use_container_width=True)

    st.subheader("What-If Analysis")
    st.write("Change price, demand, and discount assumptions to see how projected revenue may change.")
    scenario_summary = pd.DataFrame(
        {
            "Metric": ["Price Increase %", "Demand Change %", "Discount Change %", "Scenario Multiplier", "Current Revenue", "Scenario Avg Forecast"],
            "Value": [
                f"{price_increase_pct}%",
                f"{demand_change_pct}%",
                f"{discount_change_pct}%",
                f"{scenario_multiplier:.2f}x",
                f"${current_revenue:,.2f}",
                f"${scenario_forecast['Scenario Forecast'].mean():,.2f}",
            ],
        }
    )
    st.table(scenario_summary)

    comparison_left, comparison_right = st.columns(2)
    with comparison_left:
        st.metric("Current Revenue", f"${current_revenue:,.0f}")
        delta_pct = ((simulated_revenue - current_revenue) / current_revenue * 100) if current_revenue else 0.0
        st.metric("Simulated Revenue", f"${simulated_revenue:,.0f}", f"{delta_pct:.2f}%")
    with comparison_right:
        scenario_bar = go.Figure(
            data=[go.Bar(x=scenario_comparison["Scenario"], y=scenario_comparison["Revenue"], marker_color=["#1f77b4", "#ff7f0e"])]
        )
        scenario_bar.update_layout(height=320, yaxis_title="Revenue", title="Current vs Simulated Revenue")
        st.plotly_chart(scenario_bar, use_container_width=True)

    if {"Quantity", "Discount"}.issubset(filtered.columns):
        st.subheader("Operational Summary")
        summary_df = pd.DataFrame(
            {
                "Metric": ["Total Orders", "Total Quantity", "Average Discount", "Average Order Value"],
                "Value": [
                    f"{filtered['Order ID'].nunique():,}",
                    f"{filtered['Quantity'].sum():,.0f}",
                    f"{filtered['Discount'].mean() * 100:.2f}%",
                    f"${filtered[target].sum() / max(len(filtered), 1):,.2f}",
                ],
            }
        )
        st.table(summary_df)

with insights_tab:
    st.caption("This section converts numbers into plain-language business findings for a general audience.")
    st.subheader("Automated Business Insights")
    for index, insight in enumerate(automated_insights, start=1):
        st.write(f"{index}. {insight}")

    st.subheader("Insight Summary")
    st.write("This table groups the important findings into simple business themes.")
    insight_summary_df = pd.DataFrame(
        {
            "Theme": ["Trend", "Growth", "Volatility", "Top Region", "Anomalies"],
            "Observation": [
                automated_insights[0] if len(automated_insights) > 0 else "No trend insight generated.",
                automated_insights[1] if len(automated_insights) > 1 else "No growth insight generated.",
                next((item for item in automated_insights if "volatility" in item.lower()), "No volatility insight generated."),
                next((item for item in automated_insights if "region" in item.lower()), "No region insight generated."),
                next((item for item in automated_insights if "anomaly" in item.lower()), "No anomaly insight generated."),
            ],
        }
    )
    st.dataframe(insight_summary_df, use_container_width=True)

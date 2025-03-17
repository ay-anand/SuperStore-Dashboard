import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# ----- PAGE CONFIG -----
st.set_page_config(
    page_title="SuperStore KPI Dashboard",
    layout="wide",
    page_icon=":bar_chart:"
)

# ----- LOAD EXTERNAL FONTAWESOME FOR ICONS -----
st.markdown(
    """
    <link rel="stylesheet" 
          href="https://use.fontawesome.com/releases/v5.15.4/css/all.css" 
          integrity="sha384-DyZ88mC6kzNeFjsV12o4F2X0p7mUp72mmfj/wzLA16pJo1sw8a4z9ShS4rA6m1wR" 
          crossorigin="anonymous">
    """,
    unsafe_allow_html=True
)

# ----- CUSTOM CSS (including tooltip and hover zoom styles) -----
st.markdown(
    """
    <style>
    /* Hide Streamlit default elements for a cleaner look */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Custom scrollbar styling */
    ::-webkit-scrollbar {
        width: 8px;
    }
    ::-webkit-scrollbar-track {
        background: #333;
    }
    ::-webkit-scrollbar-thumb {
        background: #444;
    }

    /* KPI box styling */
    .kpi-box {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid #444;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        position: relative;
        cursor: default;
        transition: transform 0.3s ease-in-out; /* Added transition for smooth zoom */
    }
    .kpi-box:hover {
        transform: scale(1.05); /* Zoom effect on hover */
    }
    .kpi-title {
        font-weight: 600;
        color: #FFFFFF;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    .icon {
        font-size: 24px;
        margin-right: 8px;
        color: #1E90FF;
    }
    /* Tooltip styling */
    .tooltip-text {
        visibility: hidden;
        width: 220px;
        background-color: #333;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 8px;
        position: absolute;
        z-index: 1;
        bottom: 110%;
        left: 50%;
        margin-left: -110px;
        opacity: 0;
        transition: opacity 0.3s;
        font-size: 14px;
    }
    .tooltip-text::after {
        content: "";
        position: absolute;
        top: 100%;
        left: 50%;
        margin-left: -5px;
        border-width: 5px;
        border-style: solid;
        border-color: #333 transparent transparent transparent;
    }
    .kpi-box:hover .tooltip-text {
        visibility: visible;
        opacity: 1;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ----- LOAD DATA WITH CACHING -----
@st.cache_data
def load_data():
    df = pd.read_excel(
        "C:/Users/ayush/Downloads/ALY 6040/Module 4/Sample - Superstore-1.xlsx",
        engine="openpyxl"
    )
    # Convert "Order Date" to datetime if needed
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ----- TITLE -----
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>SuperStore KPI Dashboard</h1>", unsafe_allow_html=True)

# ----- SIDEBAR FILTERS -----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_regions = st.sidebar.multiselect("Region(s)", options=all_regions, default=all_regions)
df_filtered = df_original[df_original["Region"].isin(selected_regions)] if selected_regions else df_original

# State Filter
all_states = sorted(df_filtered["State"].dropna().unique())
selected_states = st.sidebar.multiselect("State(s)", options=all_states, default=all_states)
df_filtered = df_filtered[df_filtered["State"].isin(selected_states)] if selected_states else df_filtered

# Category Filter
all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_categories = st.sidebar.multiselect("Category(ies)", options=all_categories, default=all_categories)
df_filtered = df_filtered[df_filtered["Category"].isin(selected_categories)] if selected_categories else df_filtered

# Sub-Category Filter
all_subcategories = sorted(df_filtered["Sub-Category"].dropna().unique())
selected_subcategories = st.sidebar.multiselect("Sub-Category(ies)", options=all_subcategories, default=all_subcategories)
df_filtered = df_filtered[df_filtered["Sub-Category"].isin(selected_subcategories)] if selected_subcategories else df_filtered

# Date Range Filter
if not df_filtered.empty:
    min_date = df_filtered["Order Date"].min()
    max_date = df_filtered["Order Date"].max()
else:
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()

st.sidebar.subheader("Date Range")
from_date = st.sidebar.date_input("From", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To", value=max_date, min_value=min_date, max_value=max_date)

if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

df = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
].copy()

# ----- KPI CALCULATIONS -----
if df.empty:
    total_sales = total_quantity = total_profit = margin_rate = avg_discount = 0
else:
    total_sales = df["Sales"].sum()
    total_quantity = df["Quantity"].sum()
    total_profit = df["Profit"].sum()
    margin_rate = total_profit / total_sales if total_sales else 0
    avg_discount = df["Discount"].mean() if "Discount" in df.columns else 0

margin_color = "#FF4B4B" if margin_rate < 0.15 else "#1E90FF"

# ----- KPI DISPLAY -----
kpi_cols = st.columns(5)
kpi_data = [
    ("Sales", f"${total_sales:,.2f}", "Detailed Analysis: Total revenue generated."),
    ("Quantity Sold", f"{total_quantity:,.0f}", "Detailed Analysis: Total units sold."),
    ("Profit", f"${total_profit:,.2f}", "Detailed Analysis: Net profit after costs."),
    ("Margin Rate", f"{(margin_rate * 100):.2f}%", "Detailed Analysis: Profit margin percentage.", margin_color),
    ("Avg Discount", f"{(avg_discount * 100):.2f}%", "Detailed Analysis: Average discount applied.")
]

for i, item in enumerate(kpi_data):
    title, value, tooltip = item[0], item[1], item[2]
    extra_style = f"style='color:{item[3]}'" if len(item) > 3 else ""
    kpi_cols[i].markdown(
        f"""
        <div class='kpi-box' title='{tooltip}'>
            <div class='kpi-title'><i class="icon fa fa-{ 'dollar-sign' if title=='Sales' else 
                                                       'boxes' if title=='Quantity Sold' else 
                                                       'money-bill' if title=='Profit' else 
                                                       'percent' if title=='Margin Rate' else 
                                                       'tag' }"></i>{title}</div>
            <div class='kpi-value' {extra_style}>{value}</div>
            <span class="tooltip-text">{tooltip}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# ----- CHARTS -----
if df.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate", "Avg Discount"]
    selected_kpi = st.radio("Select KPI:", options=kpi_options, horizontal=True)

    # 1. Prepare daily data
    daily_grouped = df.groupby("Order Date", as_index=False).agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum",
        "Discount": "mean"
    })
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)
    daily_grouped["Avg Discount"] = daily_grouped["Discount"]

    # 2. Convert daily to monthly
    df_monthly = daily_grouped.copy()
    df_monthly.set_index("Order Date", inplace=True)
    monthly_agg = {
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum",
        "Discount": "mean"
    }
    df_monthly = df_monthly.resample("M").agg(monthly_agg)
    df_monthly["Margin Rate"] = df_monthly["Profit"] / df_monthly["Sales"].replace(0, 1)
    df_monthly["Avg Discount"] = df_monthly["Discount"]
    df_monthly.reset_index(inplace=True)

    # 3. Add a 3-month rolling average for the selected KPI
    df_monthly["3M_Rolling_Avg"] = df_monthly[selected_kpi].rolling(window=3).mean()

    # 4. Create an Area Chart for monthly data with rolling average
    fig_area = px.area(
        df_monthly,
        x="Order Date",
        y=selected_kpi,
        title=f"{selected_kpi} (Monthly) Over Time",
        template="plotly_dark",
        color_discrete_sequence=["#1E90FF"]
    )
    fig_area.add_scatter(
        x=df_monthly["Order Date"],
        y=df_monthly["3M_Rolling_Avg"],
        mode="lines+markers",
        name="3M Rolling Avg",
        line=dict(color="orange")
    )
    fig_area.data[0].update(
        name=f"{selected_kpi} Monthly",
        hovertemplate=f"Date: %{{x}}<br>{selected_kpi}: %{{y:,.2f}}"
    )
    fig_area.data[1].update(
        hovertemplate=f"Date: %{{x}}<br>3M Rolling Avg: %{{y:,.2f}}"
    )
    fig_area.update_layout(
        hovermode="x unified",
        legend=dict(yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    # 5. Top 10 Products Chart
    product_grouped = df.groupby("Product Name", as_index=False).agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum",
        "Discount": "mean"
    })
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped["Avg Discount"] = product_grouped["Discount"]
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    fig_bar = px.bar(
        top_10,
        x=selected_kpi,
        y="Product Name",
        orientation="h",
        title=f"Top 10 Products by {selected_kpi}",
        template="plotly_dark",
        color=selected_kpi,
        color_continuous_scale="Blues"
    )
    fig_bar.update_layout(yaxis={"categoryorder": "total ascending"})

    # 6. Layout: left col = monthly area chart, right col = top 10 bar chart
    col_left, col_right = st.columns(2)
    with col_left:
        st.plotly_chart(fig_area, use_container_width=True)
    with col_right:
        st.plotly_chart(fig_bar, use_container_width=True)

    # 7. Sales by Region Chart & Data Download
    col_region, col_csv = st.columns(2)
    with col_region:
        st.markdown("## Sales by Region")
        region_grouped = df.groupby("Region", as_index=False)["Sales"].sum()
        fig_region = px.bar(
            region_grouped,
            x="Region",
            y="Sales",
            title="Sales by Region",
            template="plotly_dark",
            color="Sales",
            color_continuous_scale="Blues"
        )
        fig_region.update_layout(xaxis={"categoryorder": "total descending"})
        st.plotly_chart(fig_region, use_container_width=True)
    with col_csv:
        st.markdown("## Download Data")
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name="filtered_superstore_data.csv",
            mime="text/csv"
        )
        st.dataframe(df.reset_index(drop=True), height=300, use_container_width=True)

    # ----- ADDITIONAL INSIGHTS -----
    st.markdown("## Additional Insights")

    # Sales & Profit Over Time: Dual-line chart using the monthly data
    fig_sales_profit = px.line(
        df_monthly,
        x="Order Date",
        y=["Sales", "Profit"],
        title="Sales and Profit Over Time",
        template="plotly_dark"
    )
    fig_sales_profit.update_layout(hovermode="x unified")

    # Profit by Category: Bar chart of total Profit by Category
    category_profit = df.groupby("Category", as_index=False)["Profit"].sum().sort_values(by="Profit", ascending=False)
    fig_profit_category = px.bar(
        category_profit,
        x="Category",
        y="Profit",
        title="Profit by Category",
        template="plotly_dark",
        color="Profit",
        color_continuous_scale="Blues"
    )

    # Discount vs Profit Margin: Scatter plot (create Profit Margin for each order)
    df["Profit Margin"] = df.apply(lambda row: row["Profit"] / row["Sales"] if row["Sales"] != 0 else 0, axis=1)
    fig_discount_profit = px.scatter(
        df,
        x="Discount",
        y="Profit Margin",
        color="Category",
        title="Discount vs Profit Margin",
        template="plotly_dark",
        hover_data=["Sales", "Profit"]
    )

    # Sales Distribution by Sub-Category: Treemap of Sales by Sub-Category
    subcat_sales = df.groupby("Sub-Category", as_index=False)["Sales"].sum()
    fig_treemap = px.treemap(
        subcat_sales,
        path=['Sub-Category'],
        values="Sales",
        title="Sales Distribution by Sub-Category",
        template="plotly_dark",
        color="Sales",
        color_continuous_scale="Blues"
    )

    # Layout the additional charts in two rows
    col_insight1, col_insight2 = st.columns(2)
    with col_insight1:
        st.plotly_chart(fig_sales_profit, use_container_width=True)
    with col_insight2:
        st.plotly_chart(fig_profit_category, use_container_width=True)

    col_insight3, col_insight4 = st.columns(2)
    with col_insight3:
        st.plotly_chart(fig_discount_profit, use_container_width=True)
    with col_insight4:
        st.plotly_chart(fig_treemap, use_container_width=True)

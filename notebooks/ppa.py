# %%
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from plotly.subplots import make_subplots

pio.templates

# set path to root directory
root = Path(__file__).parent.parent
file_path = root / "data" / "PPA.xlsx"

# load excel file
xls = pd.ExcelFile(file_path)

# %%
# Load the "PPA Data" sheet
ppa_data = pd.read_excel(xls, sheet_name="PPA Data")
df = ppa_data.copy()

# %%
# Generate time features
df["Date Time"] = pd.to_datetime(df["Date Time"])
df["Quarter"] = df["Date Time"].dt.quarter
df["Month"] = df["Date Time"].dt.month
df["Hour"] = df["Date Time"].dt.hour
df["Year"] = df["Date Time"].dt.year
df["Date"] = df["Date Time"].dt.date

FLOOR_PRICE = 0
FIXED_RATE = 52.55

# Float Rate = MAX(RRP, Floor Price)
df["Float_Rate"] = df["RRP"].clip(lower=FLOOR_PRICE)

# Settlement = Net Energy × (Float Rate - Fixed Rate)
df["Settlement"] = df["Net Energy (Loss Factor Adjusted)"] * (
    df["Float_Rate"] - FIXED_RATE
)

# Calculate cumulative P&L
df["Cumulative_PL"] = df["Settlement"].cumsum()

# Quarterly P&L summary
quarterly_pl = df.groupby("Quarter")["Settlement"].sum().round(2)
quarterly_stats = (
    df.groupby("Quarter")
    .agg(
        {
            "RRP": ["mean", "min", "max"],
            "Net Energy (Loss Factor Adjusted)": "sum",
            "Settlement": "sum",
        }
    )
    .round(2)
)

# %%
# Set the theme
pio.templates.default = "plotly"

# Quarterly P&L summary
quarterly_pl = df.groupby("Quarter")["Settlement"].sum().round(2)
quarterly_stats = (
    df.groupby("Quarter")
    .agg(
        {
            "RRP": ["mean", "min", "max"],
            "Net Energy (Loss Factor Adjusted)": "sum",
            "Settlement": "sum",
        }
    )
    .round(2)
)

# Create the dashboard
fig = make_subplots(
    rows=5,
    cols=2,
    subplot_titles=(
        "Quarterly P&L Summary",
        "Daily Energy Volume and RRP Trends",
        "Regional Reference Price Distribution",
        "Average Price by Hour",
        "Quarterly P&L",
        "Cumulative P&L Over Time",
        "Key Metrics Summary",
    ),
    specs=[
        [{"type": "table", "colspan": 2}, None],
        [{"type": "xy", "secondary_y": True, "colspan": 2}, None],
        [{"type": "histogram"}, {"type": "scatter"}],
        [{"type": "bar"}, {"type": "scatter"}],
        [{"type": "table", "colspan": 2}, None],
    ],
    vertical_spacing=0.05,
    horizontal_spacing=0.08,
    row_heights=[0.12, 0.25, 0.25, 0.25, 0.15],
)

# 1. Quarterly P&L Table
fig.add_trace(
    go.Table(
        header=dict(
            values=[
                "Quarter",
                "Avg RRP ($/MWh)",
                "Min RRP",
                "Max RRP",
                "Total Energy (MWh)",
                "Total P&L ($)",
            ]
        ),
        cells=dict(
            values=[
                quarterly_stats.index.astype(str),
                quarterly_stats[("RRP", "mean")],
                quarterly_stats[("RRP", "min")],
                quarterly_stats[("RRP", "max")],
                quarterly_stats[("Net Energy (Loss Factor Adjusted)", "sum")],
                quarterly_stats[("Settlement", "sum")],
            ]
        ),
    ),
    row=1,
    col=1,
)

# 2. Daily Volume and RRP Trends
daily_stats = (
    df.groupby("Date")
    .agg({"Net Energy (Loss Factor Adjusted)": "sum", "RRP": "mean"})
    .reset_index()
)

fig.add_trace(
    go.Scatter(
        x=daily_stats["Date"],
        y=daily_stats["Net Energy (Loss Factor Adjusted)"],
        mode="lines",
        name="Daily Energy Volume",
    ),
    row=2,
    col=1,
    secondary_y=False,
)

fig.add_trace(
    go.Scatter(
        x=daily_stats["Date"],
        y=daily_stats["RRP"],
        mode="lines",
        name="Average Daily RRP",
    ),
    row=2,
    col=1,
    secondary_y=True,
)

# 3. Price Distribution
fig.add_trace(
    go.Histogram(x=df["RRP"], nbinsx=50, name="RRP Distribution"), row=3, col=1
)

# 4. Average Hourly Price
hourly_price = df.groupby("Hour")["RRP"].mean().reset_index()
fig.add_trace(
    go.Scatter(
        x=hourly_price["Hour"],
        y=hourly_price["RRP"],
        mode="lines+markers",
        name="Avg Hourly Price",
    ),
    row=3,
    col=2,
)

# 5. Quarterly P&L
fig.add_trace(
    go.Bar(
        x=[f"Q{q}" for q in quarterly_pl.index],
        y=quarterly_pl.values,
        name="Quarterly P&L",
    ),
    row=4,
    col=1,
)

# 6. Cumulative P&L
fig.add_trace(
    go.Scatter(
        x=df["Date Time"], y=df["Cumulative_PL"], mode="lines", name="Cumulative P&L"
    ),
    row=4,
    col=2,
)


# Calculate key metrics
def calculate_key_metrics(df):
    # Calculate RRP mean and standard deviation
    rrp_mean = df["RRP"].mean()
    rrp_std = df["RRP"].std()

    # Define outlier thresholds (3 standard deviations)
    high_price_threshold = rrp_mean + (3 * rrp_std)
    low_price_threshold = rrp_mean - (3 * rrp_std)

    metrics = {
        "Price_Analysis": {
            "Time_Weighted_Average": rrp_mean,
            "Volume_Weighted_Average": (
                df["RRP"] * df["Net Energy (Loss Factor Adjusted)"]
            ).sum()
            / df["Net Energy (Loss Factor Adjusted)"].sum(),
            "Price_Volatility": rrp_std,
        },
        "Market_Events": {
            "High_Price_Events": len(df[df["RRP"] > high_price_threshold]),
            "Low_Price_Events": len(df[df["RRP"] < low_price_threshold]),
            "Negative_Price_Events": len(df[df["RRP"] < 0]),
        },
    }
    return metrics


metrics = calculate_key_metrics(df)

# 7. Key Metrics Table
fig.add_trace(
    go.Table(
        header=dict(values=["Metric", "Value"]),
        cells=dict(
            values=[
                [
                    "Time Weighted Avg Price",
                    "Volume Weighted Avg Price",
                    "Price Volatility (Standard Deviations)",
                    "High Price Events",
                    "Low_Price_Events",
                    "Negative Price Events",
                ],
                [
                    f"{metrics['Price_Analysis']['Time_Weighted_Average']:.2f}",
                    f"{metrics['Price_Analysis']['Volume_Weighted_Average']:.2f}",
                    f"{metrics['Price_Analysis']['Price_Volatility']:.2f}",
                    metrics["Market_Events"]["High_Price_Events"],
                    metrics["Market_Events"]["Low_Price_Events"],
                    metrics["Market_Events"]["Negative_Price_Events"],
                ],
            ]
        ),
    ),
    row=5,
    col=1,
)

# Formatting
fig.update_layout(
    height=1600,
    title_text="PPA Analysis Dashboard",
    title_x=0.5,
    title_font_size=24,
    showlegend=True,
    legend=dict(orientation="h", yanchor="bottom", xanchor="center", x=0.5),
    autosize=True,
    margin=dict(l=50, r=50, t=100, b=50),
)

# Update axes labels
fig.update_xaxes(title_text="Date", row=2, col=1)
fig.update_yaxes(title_text="Volume (MWh)", secondary_y=False, row=2, col=1)
fig.update_yaxes(title_text="RRP ($/MWh)", secondary_y=True, row=2, col=1)

fig.update_xaxes(title_text="Price ($/MWh)", row=3, col=1)
fig.update_yaxes(title_text="Frequency", row=3, col=1)

fig.update_xaxes(title_text="Hour of Day", row=3, col=2)
fig.update_yaxes(title_text="Average Price ($/MWh)", row=3, col=2)

fig.update_xaxes(title_text="Quarter", row=4, col=1)
fig.update_yaxes(title_text="P&L ($)", row=4, col=1)

fig.update_xaxes(title_text="Date", row=4, col=2)
fig.update_yaxes(title_text="Cumulative P&L ($)", row=4, col=2)

fig.show()

# Save dashboard to html
fig.write_html(
    root / "dashboard.html",
    config={
        "responsive": True,
        "displayModeBar": True,
        "displaylogo": False,
    },
    full_html=True,
    include_plotlyjs=True,
)
# %%

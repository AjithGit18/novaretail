import pandas as pd
import matplotlib.pyplot as plt

from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR

# --------------------------------------------------
# Project folder paths
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_FILE = (
    BASE_DIR
    / "data"
    / "novaretail_dataset.xlsx"
)

TEMPLATE_FILE = (
    BASE_DIR
    / "template"
    / "novaretail_template.pptx"
)

CHARTS_DIR = BASE_DIR / "charts"
OUTPUT_DIR = BASE_DIR / "output"
ICONS_DIR = BASE_DIR / "icons"


# --------------------------------------------------
# Tool 1: Analyze NovaRetail data
# --------------------------------------------------

def analyze_data() -> dict:
    """
    Read the NovaRetail dataset and return
    verified Q2 2026 business metrics.
    """

    performance = pd.read_excel(
        DATA_FILE,
        sheet_name="Regional_Performance"
    )

    customer = pd.read_excel(
        DATA_FILE,
        sheet_name="Customer_Experience"
    )

    risk = pd.read_excel(
        DATA_FILE,
        sheet_name="Operational_Risk"
    )

    q1_performance = performance[
        performance["Quarter"] == "2026-Q1"
    ]

    q2_performance = performance[
        performance["Quarter"] == "2026-Q2"
    ]

    q2_customer = customer[
        customer["Quarter"] == "2026-Q2"
    ]

    q2_risk = risk[
        risk["Quarter"] == "2026-Q2"
    ]

    # Total Q2 revenue
    total_revenue = float(
        q2_performance["Revenue_USD_M"].sum()
    )

    # Largest revenue region
    largest_region_row = q2_performance.loc[
        q2_performance["Revenue_USD_M"].idxmax()
    ]

    largest_region = str(
        largest_region_row["Region"]
    )

    largest_region_revenue = float(
        largest_region_row["Revenue_USD_M"]
    )

    # Calculate regional QoQ growth
    growth_results = []

    for region in q2_performance["Region"]:

        q1_revenue = float(
            q1_performance.loc[
                q1_performance["Region"] == region,
                "Revenue_USD_M"
            ].iloc[0]
        )

        q2_revenue = float(
            q2_performance.loc[
                q2_performance["Region"] == region,
                "Revenue_USD_M"
            ].iloc[0]
        )

        growth_percentage = (
            (q2_revenue - q1_revenue)
            / q1_revenue
        ) * 100

        growth_results.append(
            {
                "region": str(region),
                "growth_pct": float(growth_percentage)
            }
        )

    fastest_growth_result = max(
        growth_results,
        key=lambda item: item["growth_pct"]
    )

    fastest_growth_region = (
        fastest_growth_result["region"]
    )

    fastest_growth_pct = round(
        fastest_growth_result["growth_pct"],
        2
    )

    # Highest NPS region
    highest_nps_row = q2_customer.loc[
        q2_customer["NPS"].idxmax()
    ]

    highest_nps_region = str(
        highest_nps_row["Region"]
    )

    highest_nps = int(
        highest_nps_row["NPS"]
    )

    highest_nps_csat = float(
        highest_nps_row["CSAT_pct"]
    )

    highest_nps_delivery_days = float(
        highest_nps_row["Avg_Delivery_Days"]
    )

    # Lowest NPS region
    lowest_nps_row = q2_customer.loc[
        q2_customer["NPS"].idxmin()
    ]

    lowest_nps_region = str(
        lowest_nps_row["Region"]
    )

    lowest_nps = int(
        lowest_nps_row["NPS"]
    )

    lowest_nps_csat = float(
        lowest_nps_row["CSAT_pct"]
    )

    lowest_nps_delivery_days = float(
        lowest_nps_row["Avg_Delivery_Days"]
    )

    lowest_nps_support_tickets = float(
        lowest_nps_row["Support_Tickets_K"]
    )

    # Highest operational-risk region
    highest_risk_row = q2_risk.loc[
        q2_risk["Stockout_Rate_pct"].idxmax()
    ]

    highest_risk_region = str(
        highest_risk_row["Region"]
    )

    stockout_rate = float(
        highest_risk_row["Stockout_Rate_pct"]
    )

    fraud_incidents = int(
        highest_risk_row["Fraud_Incidents"]
    )

    return_processing_days = float(
        highest_risk_row["Return_Processing_Days"]
    )

    # Highest warehouse capacity utilization
    highest_capacity_row = q2_risk.loc[
        q2_risk[
            "Warehouse_Capacity_Util_pct"
        ].idxmax()
    ]

    highest_capacity_region = str(
        highest_capacity_row["Region"]
    )

    highest_capacity_utilization = float(
        highest_capacity_row[
            "Warehouse_Capacity_Util_pct"
        ]
    )

    return {
        "reporting_period": "2026-Q2",

        "total_revenue_usd_m": total_revenue,

        "total_revenue_formatted": (
            f"${total_revenue / 1000:.2f}B"
        ),

        "largest_region": largest_region,

        "largest_region_revenue_usd_m": (
            largest_region_revenue
        ),

        "fastest_growth_region": (
            fastest_growth_region
        ),

        "fastest_growth_qoq_pct": (
            fastest_growth_pct
        ),

        "highest_nps_region": (
            highest_nps_region
        ),

        "highest_nps": highest_nps,

        "highest_nps_csat_pct": (
            highest_nps_csat
        ),

        "highest_nps_delivery_days": (
            highest_nps_delivery_days
        ),

        "lowest_nps_region": (
            lowest_nps_region
        ),

        "lowest_nps": lowest_nps,

        "lowest_nps_csat_pct": (
            lowest_nps_csat
        ),

        "lowest_nps_delivery_days": (
            lowest_nps_delivery_days
        ),

        "lowest_nps_support_tickets_k": (
            lowest_nps_support_tickets
        ),

        "highest_risk_region": (
            highest_risk_region
        ),

        "highest_risk_stockout_pct": (
            stockout_rate
        ),

        "highest_risk_fraud_incidents": (
            fraud_incidents
        ),

        "highest_risk_return_processing_days": (
            return_processing_days
        ),

        "highest_capacity_region": (
            highest_capacity_region
        ),

        "highest_capacity_utilization_pct": (
            highest_capacity_utilization
        )
    }


# --------------------------------------------------
# Tool 2: Generate charts
# --------------------------------------------------

def generate_charts() -> dict:
    """
    Generate chart images for Slides 2, 3, and 4.
    """

    performance = pd.read_excel(
        DATA_FILE,
        sheet_name="Regional_Performance"
    )

    customer = pd.read_excel(
        DATA_FILE,
        sheet_name="Customer_Experience"
    )

    risk = pd.read_excel(
        DATA_FILE,
        sheet_name="Operational_Risk"
    )

    CHARTS_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    quarter_order = [
        "2025-Q3",
        "2025-Q4",
        "2026-Q1",
        "2026-Q2"
    ]

    region_order = [
        "North America",
        "Europe",
        "APAC",
        "Latin America"
    ]

    # --------------------------------------------------
    # Slide 2: Grouped revenue bar chart
    # --------------------------------------------------

    revenue = performance.pivot(
        index="Quarter",
        columns="Region",
        values="Revenue_USD_M"
    )

    revenue = revenue.reindex(
        index=quarter_order,
        columns=region_order
    )

    ax = revenue.plot(
    kind="bar",
    figsize=(10, 5),
    width=0.75
)

    plt.legend(
    loc="upper center",
    bbox_to_anchor=(0.5, -0.18),
    ncol=2,
    fontsize=8
    )

    ax.tick_params(labelsize=8)

    ax.set_title(
        "Revenue by Region Across Quarters"
    )

    ax.set_xlabel("Quarter")
    ax.set_ylabel("Revenue (USD M)")

    plt.xticks(rotation=0)
    plt.tight_layout()

    revenue_file = (
        CHARTS_DIR / "revenue_chart.png"
    )

    plt.savefig(
        revenue_file,
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()

   # --------------------------------------------------
    # Slide 3: NPS line chart
    # --------------------------------------------------

    nps = customer.pivot(
        index="Quarter",
        columns="Region",
        values="NPS"
    )

    nps = nps.reindex(
        index=quarter_order,
        columns=region_order
    )

    nps.plot(
    figsize=(10, 5),
    marker="o",
    linewidth=2
)

    plt.legend(
    loc="upper center",
    bbox_to_anchor=(0.5, -0.18),
    ncol=2,
    fontsize=8
)
    ax.tick_params(labelsize=8)

    plt.title("NPS by Region")
    plt.xlabel("Quarter")
    plt.ylabel("Net Promoter Score")
    plt.tight_layout()

    nps_file = (
        CHARTS_DIR / "nps_chart.png"
    )

    plt.savefig(
        nps_file,
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()

    # --------------------------------------------------
    # Slide 4: Stockout-risk line chart
    # --------------------------------------------------

    stockout = risk.pivot(
        index="Quarter",
        columns="Region",
        values="Stockout_Rate_pct"
    )

    stockout = stockout.reindex(
        index=quarter_order,
        columns=region_order
    )

    stockout.plot(
    figsize=(10, 5),
    marker="o",
    linewidth=2
)

    plt.legend(
    loc="upper center",
    bbox_to_anchor=(0.5, -0.18),
    ncol=2,
    fontsize=8
)
    ax.tick_params(labelsize=8)

    plt.title("Stockout Rate by Region")
    plt.xlabel("Quarter")
    plt.ylabel("Stockout Rate (%)")
    plt.tight_layout()

    risk_file = (
        CHARTS_DIR / "risk_chart.png"
    )

    plt.savefig(
        risk_file,
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()

    return {
        "status": "success",
        "revenue_chart": str(revenue_file),
        "nps_chart": str(nps_file),
        "risk_chart": str(risk_file)
    }


# --------------------------------------------------
# PowerPoint helper functions
# --------------------------------------------------

def remove_shape(shape) -> None:
    """
    Remove a PowerPoint shape from its slide.
    """

    shape_element = shape._element

    shape_element.getparent().remove(
        shape_element
    )


def replace_icon(
    slide,
    icon_placeholder,
    icon_label_placeholder,
    icon_path
):
    """
    Replace an icon placeholder with a PNG icon.
    """

    icon_box = None
    icon_label = None

    for shape in slide.shapes:

        if shape.name == icon_placeholder:
            icon_box = shape

        elif shape.name == icon_label_placeholder:
            icon_label = shape

    if icon_box is None:
        return

    left = icon_box.left
    top = icon_box.top
    width = icon_box.width
    height = icon_box.height

    remove_shape(icon_box)

    if icon_label is not None:
        remove_shape(icon_label)

    slide.shapes.add_picture(
        str(icon_path),
        left,
        top,
        width,
        height
    )


def format_text(
    shape,
    color=(255, 255, 255),
    size=14,
    bold=False
):
    """
    Format every text run inside a PowerPoint shape.
    """

    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:

        for run in paragraph.runs:

            run.font.color.rgb = RGBColor(
                color[0],
                color[1],
                color[2]
            )

            run.font.size = Pt(size)
            run.font.bold = bold


def center_text(shape):
    """
    Center text horizontally and vertically
    inside a PowerPoint text shape.
    """

    if not shape.has_text_frame:
        return

    shape.text_frame.vertical_anchor = (
        MSO_ANCHOR.MIDDLE
    )

    for paragraph in shape.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER


# --------------------------------------------------
# Tool 3: Build final PowerPoint
# --------------------------------------------------

def build_presentation(
    slide2_title: str,
    slide2_insights: str,
    slide3_title: str,
    slide3_insights: str,
    slide4_title: str,
    slide4_insights: str,
    slide5_recommendations: str,
    slide5_closing_statement: str
) -> dict:
    """
    Build the complete five-slide NovaRetail presentation.
    """

    metrics = analyze_data()
    chart_files = generate_charts()

    prs = Presentation(
        str(TEMPLATE_FILE)
    )

# --------------------------------------------------
# Slide 1: Executive Summary
# --------------------------------------------------

    slide1 = prs.slides[0]

    for shape in slide1.shapes:

        if shape.name == "subtitle_1":

            shape.text = (
                "Q2 2026 Executive Review"
            )

            format_text(
                shape,
                color=(202, 220, 252),
                size=14
            )

        elif shape.name == "kpi1_value":

            shape.text = metrics[
                "total_revenue_formatted"
            ]

            format_text(
                shape,
                color=(255, 255, 255),
                size=32,
                bold=True
            )
            center_text(shape)

        elif shape.name == "kpi1_label":

            shape.text = "Total Revenue"

            format_text(
                shape,
                color=(202, 220, 252),
                size=12,
                bold=True
            )

            center_text(shape)

        elif shape.name == "kpi2_value":

            shape.text = (
                f"{metrics['fastest_growth_qoq_pct']:.1f}%"
            )

            format_text(
                shape,
                color=(255, 255, 255),
                size=32,
                bold=True
            )
            center_text(shape)

        elif shape.name == "kpi2_label":

            shape.text = (
                "Fastest Regional QoQ Growth"
            )

            format_text(
                shape,
                color=(202, 220, 252),
                size=12,
                bold=True
            )
            center_text(shape)

        elif shape.name == "kpi3_value":

            shape.text = str(
                metrics["highest_nps"]
            )

            format_text(
                shape,
                color=(255, 255, 255),
                size=32,
                bold=True
            )

            center_text(shape)

        elif shape.name == "kpi3_label":

            shape.text = (
                "Highest Regional NPS"
            )

            format_text(
                shape,
                color=(202, 220, 252),
                size=12,
                bold=True
            )
            center_text(shape)

        elif shape.name == "summary_text":

            shape.text = (
                f"NovaRetail delivered strong Q2 2026 performance with "
                f"{metrics['total_revenue_formatted']} in revenue. "
                f"{metrics['largest_region']} remained the largest revenue region at "
                f"${metrics['largest_region_revenue_usd_m']:.0f}M. "
                f"{metrics['fastest_growth_region']} achieved the fastest "
                f"quarter-over-quarter growth of "
                f"{metrics['fastest_growth_qoq_pct']:.1f}% and recorded the "
                f"highest NPS of {metrics['highest_nps']}. "
                f"{metrics['highest_risk_region']} remains the most significant "
                f"operational risk area and requires focused leadership attention."
            )

            format_text(
                shape,
                color=(255, 255, 255),
                size=15
            )

    #---- remove_template_elements

    def remove_template_elements(slide):

        shapes_to_remove = [
        "footer_1", "footer_2", "footer_3",
        "footer_4", "footer_5",
        "page_num_1", "page_num_2", "page_num_3",
        "page_num_4", "page_num_5"
    ]

        for shape in list(slide.shapes):

            if shape.name in shapes_to_remove:
                remove_shape(shape)


    # --------------------------------------------------
    # Slide 2: Regional Performance
    # --------------------------------------------------

    slide2 = prs.slides[1]

    performance_chart_box = None
    performance_chart_label = None

    for shape in slide2.shapes:

        if shape.name == "title_2":
            shape.text = slide2_title

        elif shape.name == "insights_performance":
            shape.text = slide2_insights

        elif shape.name == "chart_performance":
            performance_chart_box = shape

        elif shape.name == "chart_performance_label":
            performance_chart_label = shape

    if performance_chart_box is not None:

        chart_left = performance_chart_box.left
        chart_top = performance_chart_box.top
        chart_width = performance_chart_box.width
        chart_height = performance_chart_box.height

        remove_shape(
            performance_chart_box
        )

        if performance_chart_label is not None:
            remove_shape(
                performance_chart_label
            )

        slide2.shapes.add_picture(
            chart_files["revenue_chart"],
            chart_left,
            chart_top,
            chart_width,
            chart_height
        )

    # --------------------------------------------------
    # Slide 3: Customer Experience
    # --------------------------------------------------

    slide3 = prs.slides[2]

    customer_chart_box = None
    customer_chart_label = None

    for shape in slide3.shapes:

        if shape.name == "title_3":
            shape.text = slide3_title

        elif shape.name == "insights_customer":
            shape.text = slide3_insights

        elif shape.name == "chart_customer":
            customer_chart_box = shape

        elif shape.name == "chart_customer_label":
            customer_chart_label = shape

    if customer_chart_box is not None:

        chart_left = customer_chart_box.left
        chart_top = customer_chart_box.top
        chart_width = customer_chart_box.width
        chart_height = customer_chart_box.height

        remove_shape(
            customer_chart_box
        )

        if customer_chart_label is not None:
            remove_shape(
                customer_chart_label
            )

        slide3.shapes.add_picture(
            chart_files["nps_chart"],
            chart_left,
            chart_top,
            chart_width,
            chart_height
        )

    # --------------------------------------------------
    # Slide 4: Operational Risk
    # --------------------------------------------------

    slide4 = prs.slides[3]

    risk_chart_box = None
    risk_chart_label = None

    for shape in slide4.shapes:

        if shape.name == "title_4":
            shape.text = slide4_title

        elif shape.name == "insights_risk":
            shape.text = slide4_insights

        elif shape.name == "chart_risk":
            risk_chart_box = shape

        elif shape.name == "chart_risk_label":
            risk_chart_label = shape

    if risk_chart_box is not None:

        chart_left = risk_chart_box.left
        chart_top = risk_chart_box.top
        chart_width = risk_chart_box.width
        chart_height = risk_chart_box.height

        remove_shape(
            risk_chart_box
        )

        if risk_chart_label is not None:
            remove_shape(
                risk_chart_label
            )

        slide4.shapes.add_picture(
            chart_files["risk_chart"],
            chart_left,
            chart_top,
            chart_width,
            chart_height
        )

    # --------------------------------------------------
    # Slide 5: Recommendations
    # --------------------------------------------------

    slide5 = prs.slides[4]

    for shape in slide5.shapes:

        if shape.name == "recommendations_text":

            shape.text = slide5_recommendations

            format_text(
                shape,
                color=(255, 255, 255),
                size=14
            )

        elif shape.name == "closing_statement":

            shape.text = (
                slide5_closing_statement
            )

            format_text(
                shape,
                color=(202, 220, 252),
                size=12,
                bold=True
            )

    # --------------------------------------------------
    # Replace icons
    # --------------------------------------------------

    replace_icon(
        prs.slides[0],
        "icon_1",
        "icon_1_label",
        ICONS_DIR / "globe_region.png"
    )

    replace_icon(
        prs.slides[1],
        "icon_2",
        "icon_2_label",
        ICONS_DIR / "growth_chart.png"
    )

    replace_icon(
        prs.slides[2],
        "icon_3",
        "icon_3_label",
        ICONS_DIR / "customer_star.png"
    )

    replace_icon(
        prs.slides[3],
        "icon_4",
        "icon_4_label",
        ICONS_DIR / "risk_alert.png"
    )

    replace_icon(
        prs.slides[4],
        "icon_5",
        "icon_5_label",
        ICONS_DIR / "lightbulb_action.png"
    )

    # --------------------------------------------------
    # Remove template footer and page numbers
    # --------------------------------------------------

    for slide in prs.slides:
        remove_template_elements(slide)

    # --------------------------------------------------
    # Save final presentation
    # --------------------------------------------------

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    output_file = (
        OUTPUT_DIR
        / f"NovaRetail_QBR_{timestamp}.pptx"
    )

    prs.save(
        str(output_file)
    )

    return {
        "status": "success",
        "message": (
            "The complete AI-generated NovaRetail "
            "presentation was created successfully."
        ),
        "output_file": str(output_file),
        "slides_created": [
            1,
            2,
            3,
            4,
            5
        ]
    }
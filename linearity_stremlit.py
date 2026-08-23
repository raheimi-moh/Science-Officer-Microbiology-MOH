import io
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(page_title="Linearity Analysis Tool", layout="wide")

st.title("🧪 Linearity Analysis & Validation Tool")
st.markdown(
    "Upload your Excel dataset containing concentration proportions and replicate measurements to execute the linearity study."
)

# Sidebar - File Upload & Options
st.sidebar.header("Data Input Options")
uploaded_file = st.sidebar.file_uploader(
    "Upload Linearity Excel File (.xlsx)", type=["xlsx", "xls"]
)

# Custom ADL Percentage Setting
adl_pct = (
    st.sidebar.number_input(
        "Allowable Deviation Limit (% ADL)",
        min_value=0.1,
        max_value=20.0,
        value=2.0,
        step=0.1,
    )
    / 100.0
)

if uploaded_file is not None:
    # Read uploaded file
    df = pd.read_excel(uploaded_file)
    original_filename = uploaded_file.name
    output_filename = f"{original_filename.rsplit('.', 1)[0]}_analyzed.xlsx"

    # Automatically Detect Columns
    prop_cols = [
        c
        for c in df.columns
        if any(k in c.lower() for k in ["prop", "pool", "conc", "level"])
    ]
    prop_col = prop_cols[0] if prop_cols else df.columns[0]
    replicate_cols = [c for c in df.columns if c.lower().startswith("rep")]

    st.sidebar.success(f"Detected Proportion Column: **{prop_col}**")
    st.sidebar.info(
        f"Detected Replicates ({len(replicate_cols)}): **{', '.join(replicate_cols)}**"
    )

    # ----------------------------------------------------
    # Step 1: Compute Mean, SD, %CV
    # ----------------------------------------------------
    df["Measured Mean"] = df[replicate_cols].mean(axis=1)
    df["SD"] = df[replicate_cols].std(axis=1, ddof=1)
    df["% CV"] = (df["SD"] / df["Measured Mean"]) * 100

    # ----------------------------------------------------
    # Step 2: Precision Profile Model (Force Intercept = 0)
    # Model: SD = A * Measured_Mean
    # ----------------------------------------------------
    df["XY (Product 1)"] = df["Measured Mean"] * df["SD"]
    df["X^2 (Product 2)"] = df["Measured Mean"] ** 2

    product_1 = np.sum(df["XY (Product 1)"])
    product_2 = np.sum(df["X^2 (Product 2)"])
    slope_A = product_1 / product_2 if product_2 != 0 else 0

    # ----------------------------------------------------
    # Step 3: Sigma Value and Weight Calculation
    # ----------------------------------------------------
    df["Sigma Value"] = slope_A * df["Measured Mean"]

    weights = []
    for sig in df["Sigma Value"]:
        if sig == 0 or np.isnan(sig):
            weights.append(2.0)
        else:
            weights.append(1.0 / (sig**2))
    df["Weight"] = weights

    # ----------------------------------------------------
    # Step 4: Weighted Least Squares (WLS) Linear Regression
    # ----------------------------------------------------
    df["WX"] = df["Weight"] * df[prop_col]
    df["WY"] = df["Weight"] * df["Measured Mean"]
    df["WX^2"] = df["Weight"] * (df[prop_col] ** 2)
    df["WXY"] = df["Weight"] * df[prop_col] * df["Measured Mean"]

    sum_W = np.sum(df["Weight"])
    sum_WX = np.sum(df["WX"])
    sum_WY = np.sum(df["WY"])
    sum_WX2 = np.sum(df["WX^2"])
    sum_WXY = np.sum(df["WXY"])

    denom = (sum_W * sum_WX2) - (sum_WX**2)
    if denom != 0:
        wls_slope = ((sum_W * sum_WXY) - (sum_WX * sum_WY)) / denom
        wls_intercept = (sum_WY - (wls_slope * sum_WX)) / sum_W
    else:
        wls_slope = 0
        wls_intercept = 0

    # ----------------------------------------------------
    # Step 5 & 6: Predicted Value & Deviation from Linearity
    # ----------------------------------------------------
    df["Predicted Value of Mean"] = (wls_slope * df[prop_col]) + wls_intercept
    df["Deviation from Linearity"] = (
        df["Measured Mean"] - df["Predicted Value of Mean"]
    )

    # ----------------------------------------------------
    # Step 7: Confidence Interval (CI) Calculation
    # ----------------------------------------------------
    num_replicates = len(replicate_cols)
    num_levels = len(df)
    z_factor = 2.38 if num_levels == 6 else 2.31

    df["Upper Limit CI"] = df["Deviation from Linearity"] + (
        z_factor * (df["Sigma Value"] / np.sqrt(num_replicates))
    )
    df["Lower Limit CI"] = df["Deviation from Linearity"] - (
        z_factor * (df["Sigma Value"] / np.sqrt(num_replicates))
    )
    df["+ ADL"] = adl_pct * df["Predicted Value of Mean"]

    # ----------------------------------------------------
    # Step 8: Evaluation Summary & Overall Conclusion
    # ----------------------------------------------------
    evaluation = []
    for idx, row in df.iterrows():
        dev = abs(row["Deviation from Linearity"])
        adl = abs(row["+ ADL"])
        if dev <= adl:
            evaluation.append("Within ADL")
        elif (row["Lower Limit CI"] <= adl) or (row["Upper Limit CI"] >= -adl):
            evaluation.append("CI overlaps ADL")
        else:
            evaluation.append("Exceeds ADL")

    df["Linearity Assessment"] = evaluation

    exceed_adl_pools = df[df["Linearity Assessment"] == "Exceeds ADL"][
        prop_col
    ].tolist()
    if not exceed_adl_pools:
        overall_conclusion = (
            "Linearity is ACCEPTABLE. All concentration pools fall within ADL "
            "or have confidence limits overlapping ADL."
        )
        st.success(f"**Conclusion:** {overall_conclusion}")
    else:
        overall_conclusion = (
            f"Linearity is NOT ACCEPTABLE. Pools exceeding ADL: {exceed_adl_pools}"
        )
        st.error(f"**Conclusion:** {overall_conclusion}")

    export_df = df.drop(
        columns=["XY (Product 1)", "X^2 (Product 2)", "WX", "WY", "WX^2", "WXY"]
    )
    export_df["Overall Conclusion"] = overall_conclusion

    # ----------------------------------------------------
    # DISPLAY INTERACTIVE TABLES IN TABS
    # ----------------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📊 Step 1-3: Descriptive Statistics",
            "🧮 Step 4: Regression Terms",
            "📋 Step 5-8: Linearity Summary",
            "📉 Visualizations",
        ]
    )

    with tab1:
        st.subheader("Step 1–3: Means, SD, Precision Profile & Weights")
        st.dataframe(
            df[
                [
                    prop_col,
                    "Measured Mean",
                    "SD",
                    "% CV",
                    "Sigma Value",
                    "Weight",
                ]
            ],
            use_container_width=True,
        )
        st.info(f"**Precision Profile Model:** $SD = {slope_A:.6f} \\times Mean$")

    with tab2:
        st.subheader("Step 4: Weighted Least Squares (WLS) Parameters")
        st.dataframe(
            df[[prop_col, "Measured Mean", "Weight", "WX", "WY", "WX^2", "WXY"]],
            use_container_width=True,
        )
        st.info(
            f"**WLS Linear Model Equation:** $Y = {wls_slope:.6f} \\times X + {wls_intercept:.6f}$"
        )

    with tab3:
        st.subheader("Step 5–8: Linearity & Confidence Limit Summary")
        summary_cols = [
            prop_col,
            "Measured Mean",
            "Predicted Value of Mean",
            "Deviation from Linearity",
            "Upper Limit CI",
            "Lower Limit CI",
            "+ ADL",
            "Linearity Assessment",
        ]
        st.dataframe(df[summary_cols], use_container_width=True)

    with tab4:
        st.subheader("Plots & Diagnostic Charts")
        fig, axes = plt.subplots(1, 2, figsize=(14, 5))

        # Precision Profile Chart
        axes[0].scatter(
            df["Measured Mean"],
            df["SD"],
            color="blue",
            s=50,
            zorder=3,
            label="Observed SD",
        )
        x_line = np.linspace(0, df["Measured Mean"].max(), 100)
        axes[0].plot(
            x_line,
            slope_A * x_line,
            color="red",
            linestyle="--",
            label=f"Precision Profile (A={slope_A:.5f})",
        )
        axes[0].set_title("Precision Profile Model", fontweight="bold")
        axes[0].set_xlabel("Measured Mean")
        axes[0].set_ylabel("Standard Deviation")
        axes[0].grid(True, linestyle=":", alpha=0.6)
        axes[0].legend()

        # Linearity Deviation Chart
        df_sorted = df.sort_values(prop_col)
        axes[1].plot(
            df_sorted[prop_col],
            df_sorted["Deviation from Linearity"],
            marker="o",
            color="black",
            label="Deviation",
            linewidth=2,
        )
        axes[1].plot(
            df_sorted[prop_col],
            df_sorted["Upper Limit CI"],
            color="blue",
            linestyle="--",
            label="CI Limits",
        )
        axes[1].plot(
            df_sorted[prop_col],
            df_sorted["Lower Limit CI"],
            color="blue",
            linestyle="--",
        )
        axes[1].plot(
            df_sorted[prop_col],
            df_sorted["+ ADL"],
            color="green",
            linestyle=":",
            linewidth=2,
            label=f"+/- ADL ({adl_pct*100:.1f}%)",
        )
        axes[1].plot(
            df_sorted[prop_col],
            -df_sorted["+ ADL"],
            color="green",
            linestyle=":",
            linewidth=2,
        )
        axes[1].axhline(0, color="gray", alpha=0.5, linestyle="-")
        axes[1].set_title("Linearity Deviation Analysis", fontweight="bold")
        axes[1].set_xlabel(prop_col)
        axes[1].set_ylabel("Deviation Value")
        axes[1].grid(True, linestyle=":", alpha=0.6)
        axes[1].legend()

        st.pyplot(fig)

    # ----------------------------------------------------
    # FILE DOWNLOAD BUTTON
    # ----------------------------------------------------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Linearity Analysis")

    st.sidebar.markdown("---")
    st.sidebar.download_button(
        label="📥 Download Excel Results",
        data=buffer.getvalue(),
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info(
        "👋 Please upload an Excel file using the sidebar to begin the linearity evaluation."
    )

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

st.set_page_config(page_title="Power Query & CFO Financial Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS Styling for KPI Cards
st.markdown("""
<style>
    .metric-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 16px;
        border: 1px solid #e2e8f0;
        border-left: 6px solid #2563eb;
        box-shadow: 0 4px 6px rgba(0,0,0,0.04);
    }
    .metric-title { font-size: 12px; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-value { font-size: 22px; font-weight: 800; color: #0f172a; margin-top: 4px; }
    .metric-sub { font-size: 12px; font-weight: 600; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

st.title("⚡ Power Query & CFO Financial Operating System")

# Initialize Session State
if "custom_df" not in st.session_state:
    st.session_state.custom_df = None

# Sidebar - Global Controls
st.sidebar.header("📁 File & Data Controls")
uploaded_files = st.sidebar.file_uploader("Upload Financial Report(s)", type=["xlsx", "xls"], accept_multiple_files=True)

# Auto-load default file if available in folder and no file uploaded
default_excel = "Financial_Year_Report.xlsx"
combined_sheets = {}

if uploaded_files:
    data_dict = {}
    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=sheet)
            df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]
            if sheet not in data_dict:
                data_dict[sheet] = []
            data_dict[sheet].append(df_temp)
    combined_sheets = {sheet: pd.concat(dfs, ignore_index=True) for sheet, dfs in data_dict.items()}
elif os.path.exists(default_excel):
    xls = pd.ExcelFile(default_excel)
    combined_sheets = {sheet: pd.read_excel(xls, sheet_name=sheet).loc[:, lambda d: ~d.columns.duplicated()] for sheet in xls.sheet_names}

if combined_sheets:
    selected_sheet = st.sidebar.selectbox("Select View Module", list(combined_sheets.keys()))
    df = combined_sheets[selected_sheet]

    if "INV. DATE" in df.columns or "BILL AMOUNT" in df.columns:
        if "INV. DATE" in df.columns:
            df["INV. DATE"] = pd.to_datetime(df["INV. DATE"], errors="coerce")
            df["Financial Year"] = df["INV. DATE"].apply(
                lambda dt: f"FY {dt.year}-{str(dt.year+1)[-2:]}" if pd.notnull(dt) and dt.month >= 4 else (f"FY {dt.year-1}-{str(dt.year)[-2:]}" if pd.notnull(dt) else "Unknown")
            )
            
        if "Gross Service Revenue" in df.columns and "BILL AMOUNT" in df.columns:
            df["Gross Profit %"] = (df["Gross Service Revenue"] / df["BILL AMOUNT"].replace(0, 1)) * 100

        if st.session_state.custom_df is None or len(st.session_state.custom_df) != len(df):
            st.session_state.custom_df = df.copy()

        # Global Filters
        if "Financial Year" in df.columns:
            years = ["All Years"] + sorted([y for y in df["Financial Year"].unique() if y != "Unknown"], reverse=True)
            selected_year = st.sidebar.selectbox("Filter Financial Year", years)
            working_df = st.session_state.custom_df[st.session_state.custom_df["Financial Year"] == selected_year] if selected_year != "All Years" else st.session_state.custom_df.copy()
        else:
            working_df = st.session_state.custom_df.copy()

        if "CUSTOMER NAME" in df.columns:
            customers = ["All Customers"] + sorted(df["CUSTOMER NAME"].dropna().astype(str).unique().tolist())
            selected_customer = st.sidebar.selectbox("Filter Customer", customers)
            if selected_customer != "All Customers":
                working_df = working_df[working_df["CUSTOMER NAME"] == selected_customer]

        tab_dash, tab_pq, tab_data = st.tabs(["📊 Executive Dashboard", "⚡ Power Query Studio", "📝 Raw Data & Grid"])

        # TAB 1: EXECUTIVE DASHBOARD
        with tab_dash:
            st.subheader("📌 Key Financial & Profit Margin Metrics")
            
            tot_rev = working_df["TOTAL BILL AMOUNT"].sum() if "TOTAL BILL AMOUNT" in working_df.columns else 0
            tot_bill = working_df["BILL AMOUNT"].sum() if "BILL AMOUNT" in working_df.columns else 0
            tot_gross_profit = working_df["Gross Service Revenue"].sum() if "Gross Service Revenue" in working_df.columns else 0
            tot_rec = working_df["TOTAL AMOUNT RECEIVED"].sum() if "TOTAL AMOUNT RECEIVED" in working_df.columns else 0
            tot_bal = working_df["BALANCE AMOUNT TO BE RECEIVED"].sum() if "BALANCE AMOUNT TO BE RECEIVED" in working_df.columns else 0

            gp_percentage_base = (tot_gross_profit / tot_bill * 100) if tot_bill > 0 else 0
            gp_percentage_total = (tot_gross_profit / tot_rev * 100) if tot_rev > 0 else 0

            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            with kpi1:
                st.markdown(f'<div class="metric-card" style="border-left-color: #2563eb;"><div class="metric-title">TOTAL REVENUE (INC TAX)</div><div class="metric-value">₹{tot_rev:,.2f}</div></div>', unsafe_allow_html=True)
            with kpi2:
                st.markdown(f'<div class="metric-card" style="border-left-color: #059669;"><div class="metric-title">GROSS SERVICE PROFIT</div><div class="metric-value" style="color:#059669;">₹{tot_gross_profit:,.2f}</div></div>', unsafe_allow_html=True)
            with kpi3:
                st.markdown(f'<div class="metric-card" style="border-left-color: #0d9488;"><div class="metric-title">GROSS PROFIT MARGIN (%)</div><div class="metric-value" style="color:#0d9488;">{gp_percentage_base:.2f}%</div><div class="metric-sub" style="color:#0d9488;">({gp_percentage_total:.2f}% on Total Revenue)</div></div>', unsafe_allow_html=True)
            with kpi4:
                st.markdown(f'<div class="metric-card" style="border-left-color: #dc2626;"><div class="metric-title">OUTSTANDING BALANCE</div><div class="metric-value" style="color:#dc2626;">₹{tot_bal:,.2f}</div></div>', unsafe_allow_html=True)
            with kpi5:
                st.markdown(f'<div class="metric-card" style="border-left-color: #16a34a;"><div class="metric-title">COLLECTED REVENUE</div><div class="metric-value">₹{tot_rec:,.2f}</div></div>', unsafe_allow_html=True)

            st.divider()

            # LEADERBOARD SECTION
            st.subheader("🏆 Client Ranking & Dynamic Analysis")
            c_ctrl1, c_ctrl2 = st.columns([1, 2])
            with c_ctrl1:
                top_n = st.slider("Select Top N Clients", min_value=5, max_value=25, value=10)
            with c_ctrl2:
                view_mode = st.radio("Select Leaderboard View Mode:", ["Side-by-Side Comparison", "Total Revenue (Inc Tax)", "Base Bill Amount", "Gross Profit (₹)", "Gross Margin (%)"], horizontal=True)

            if view_mode == "Side-by-Side Comparison":
                col_c1, col_c2, col_c3 = st.columns(3)
                with col_c1:
                    st.markdown("##### 💰 Top Clients by Total Revenue (Inc Tax)")
                    if "CUSTOMER NAME" in working_df.columns and "TOTAL BILL AMOUNT" in working_df.columns:
                        df_rev = working_df.groupby("CUSTOMER NAME")["TOTAL BILL AMOUNT"].sum().nlargest(top_n).reset_index()
                        fig = px.bar(df_rev, x="TOTAL BILL AMOUNT", y="CUSTOMER NAME", orientation="h", color="TOTAL BILL AMOUNT", color_continuous_scale="Blues")
                        fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False, margin=dict(l=0, r=0, t=20, b=0))
                        st.plotly_chart(fig, use_container_width=True)

                with col_c2:
                    st.markdown("##### 📄 Top Clients by Base Bill Amount")
                    if "CUSTOMER NAME" in working_df.columns and "BILL AMOUNT" in working_df.columns:
                        df_bill = working_df.groupby("CUSTOMER NAME")["BILL AMOUNT"].sum().nlargest(top_n).reset_index()
                        fig = px.bar(df_bill, x="BILL AMOUNT", y="CUSTOMER NAME", orientation="h", color="BILL AMOUNT", color_continuous_scale="Purples")
                        fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False, margin=dict(l=0, r=0, t=20, b=0))
                        st.plotly_chart(fig, use_container_width=True)

                with col_c3:
                    st.markdown("##### 📈 Top Clients by Gross Profit (₹)")
                    if "CUSTOMER NAME" in working_df.columns and "Gross Service Revenue" in working_df.columns:
                        df_profit = working_df.groupby("CUSTOMER NAME")["Gross Service Revenue"].sum().nlargest(top_n).reset_index()
                        fig = px.bar(df_profit, x="Gross Service Revenue", y="CUSTOMER NAME", orientation="h", color="Gross Service Revenue", color_continuous_scale="Greens")
                        fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False, margin=dict(l=0, r=0, t=20, b=0))
                        st.plotly_chart(fig, use_container_width=True)

            elif view_mode == "Total Revenue (Inc Tax)":
                df_rev = working_df.groupby("CUSTOMER NAME")["TOTAL BILL AMOUNT"].sum().nlargest(top_n).reset_index()
                fig = px.bar(df_rev, x="TOTAL BILL AMOUNT", y="CUSTOMER NAME", orientation="h", color="TOTAL BILL AMOUNT", color_continuous_scale="Blues", text_auto=',.2f')
                fig.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)

            elif view_mode == "Base Bill Amount":
                df_bill = working_df.groupby("CUSTOMER NAME")["BILL AMOUNT"].sum().nlargest(top_n).reset_index()
                fig = px.bar(df_bill, x="BILL AMOUNT", y="CUSTOMER NAME", orientation="h", color="BILL AMOUNT", color_continuous_scale="Purples", text_auto=',.2f')
                fig.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)

            elif view_mode == "Gross Profit (₹)":
                df_profit = working_df.groupby("CUSTOMER NAME")["Gross Service Revenue"].sum().nlargest(top_n).reset_index()
                fig = px.bar(df_profit, x="Gross Service Revenue", y="CUSTOMER NAME", orientation="h", color="Gross Service Revenue", color_continuous_scale="Greens", text_auto=',.2f')
                fig.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)

            elif view_mode == "Gross Margin (%)":
                cust_pct = working_df.groupby("CUSTOMER NAME")[["Gross Service Revenue", "BILL AMOUNT"]].sum().reset_index()
                cust_pct["Margin %"] = (cust_pct["Gross Service Revenue"] / cust_pct["BILL AMOUNT"].replace(0, 1)) * 100
                top_pct_cust = cust_pct.nlargest(top_n, "Margin %")
                fig = px.bar(top_pct_cust, x="Margin %", y="CUSTOMER NAME", orientation="h", color="Margin %", color_continuous_scale="Teal", text_auto='.2f')
                fig.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)

            st.divider()

            # Monthly Trend Chart
            st.subheader("📅 Monthly Financial & Profit Trend")
            if "INV RAISED MONTH" in working_df.columns:
                trend_df = working_df.groupby("INV RAISED MONTH")[["TOTAL BILL AMOUNT", "BILL AMOUNT", "Gross Service Revenue"]].sum().reset_index()
                fig_trend = px.line(
                    trend_df, x="INV RAISED MONTH", y=["TOTAL BILL AMOUNT", "BILL AMOUNT", "Gross Service Revenue"],
                    markers=True, labels={"value": "Amount (₹)", "INV RAISED MONTH": "Month", "variable": "Financial Metric"}
                )
                st.plotly_chart(fig_trend, use_container_width=True)

        # TAB 2: POWER QUERY STUDIO
        with tab_pq:
            st.subheader("🛠️ Power Query Engine")
            st.info("Build custom formulas, transform columns, or create pivot aggregations interactively without writing code.")

            pq_col1, pq_col2 = st.columns(2)
            with pq_col1:
                st.markdown("### 🧮 1. Add Custom Calculated Column")
                num_cols = [c for c in working_df.columns if pd.api.types.is_numeric_dtype(working_df[c])]
                
                if len(num_cols) >= 2:
                    col_a = st.selectbox("Select Column A", num_cols, key="pq_a")
                    op = st.selectbox("Operation", ["-", "+", "*", "/", "% (A as % of B)"], key="pq_op")
                    col_b = st.selectbox("Select Column B", num_cols, key="pq_b")
                    new_col_name = st.text_input("New Field Name", value="Custom Calculated Field", key="pq_name")

                    if st.button("➕ Create Calculated Column"):
                        if op == "+":
                            st.session_state.custom_df[new_col_name] = st.session_state.custom_df[col_a] + st.session_state.custom_df[col_b]
                        elif op == "-":
                            st.session_state.custom_df[new_col_name] = st.session_state.custom_df[col_a] - st.session_state.custom_df[col_b]
                        elif op == "*":
                            st.session_state.custom_df[new_col_name] = st.session_state.custom_df[col_a] * st.session_state.custom_df[col_b]
                        elif op == "/":
                            st.session_state.custom_df[new_col_name] = st.session_state.custom_df[col_a] / st.session_state.custom_df[col_b].replace(0, 1)
                        elif op == "% (A as % of B)":
                            st.session_state.custom_df[new_col_name] = (st.session_state.custom_df[col_a] / st.session_state.custom_df[col_b].replace(0, 1)) * 100

                        st.success(f"Column '{new_col_name}' created successfully!")
                        st.rerun()

            with pq_col2:
                st.markdown("### 📊 2. Dynamic Group By / Pivot Aggregation")
                all_cols = working_df.columns.tolist()
                group_col = st.selectbox("Group By Dimension", all_cols, index=0, key="grp_col")
                valid_metrics = [c for c in all_cols if c != group_col and pd.api.types.is_numeric_dtype(working_df[c])]
                
                if valid_metrics:
                    metric_col = st.selectbox("Metric Column to Aggregate", valid_metrics, key="grp_metric")
                    agg_func = st.selectbox("Aggregation Function", ["SUM", "AVERAGE (MEAN)", "COUNT", "MIN", "MAX"], key="grp_func")

                    if group_col and metric_col:
                        try:
                            grouped = working_df.groupby(group_col, as_index=False)[metric_col]
                            if agg_func == "SUM":
                                pivot_df = grouped.sum()
                            elif agg_func == "AVERAGE (MEAN)":
                                pivot_df = grouped.mean()
                            elif agg_func == "COUNT":
                                pivot_df = grouped.count()
                            elif agg_func == "MIN":
                                pivot_df = grouped.min()
                            elif agg_func == "MAX":
                                pivot_df = grouped.max()

                            st.dataframe(pivot_df, use_container_width=True)
                        except Exception as e:
                            st.error(f"Could not group by selected columns: {e}")

            st.divider()
            if st.button("🔄 Reset Power Query Transformations"):
                st.session_state.custom_df = df.copy()
                st.success("Transformed data reset back to original upload!")
                st.rerun()

        # TAB 3: LIVE GRID
        with tab_data:
            st.subheader("📝 Interactive Grid (Edit, Append or Modify Records)")
            edited_df = st.data_editor(working_df, num_rows="dynamic", use_container_width=True)
            csv = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Export Current Transformed Dataset to CSV", data=csv, file_name="Power_Query_Transformed_Financials.csv", mime="text/csv")

    else:
        st.subheader(f"Sheet View: {selected_sheet}")
        st.data_editor(df, num_rows="dynamic", use_container_width=True)
else:
    st.info("👈 Please upload your financial Excel report in the sidebar or place 'Financial_Year_Report.xlsx' in the folder to get started.")
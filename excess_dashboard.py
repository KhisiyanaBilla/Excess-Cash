import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime, timedelta
import re

# -----------------------------
# Page Setup
# -----------------------------
st.set_page_config(page_title="Excess Cash Monitoring", layout="wide")
PASSWORD = "jabalpur123"

# -----------------------------
# Login
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    password_input = st.text_input("Enter Password:", type="password")
    if st.button("Login"):
        if password_input == PASSWORD:
            st.session_state.authenticated = True
            st.success("Access Granted! Press Login again to enter.")
else:
    st.title("Excess Cash Monitoring – Jabalpur Region")

    # -----------------------------
    # Tabs
    # -----------------------------
    tab1, tab2 = st.tabs(
        ["🚨 VERY HIGH RISK OFFICES", "💸 REMITTANCE MONITORING"]
    )

    # ================================
    # TAB 1: Very High Risk Offices
    # ================================
    with tab1:

        st.markdown("""
        <div style="
            background-color:#7a1025;
            padding:14px 18px;
            border-radius:10px;
            margin-bottom:14px;
        ">
        <h2 style="color:white; margin:0;">
        🚨 VERY HIGH RISK OFFICES
        </h2>
        <p style="color:#f5b73b; margin:4px 0 0 0; font-size:0.9em;">
        Persistent excess cash beyond permissible limits
        </p>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Select Excel File for High Risk Analysis",
            type=["xlsx"],
            key="tab1_upload"
        )

        if uploaded_file:
            df = pd.read_excel(uploaded_file)

            required_columns = [
                "Date", "Division", "Office Type", "Office Name",
                "Office ID", "Max Amount", "Excess Amount", "Closing Balance"
            ]

            missing_cols = [c for c in required_columns if c not in df.columns]
            if missing_cols:
                st.error(f"Missing columns: {missing_cols}")
            else:
                df['Date'] = pd.to_datetime(
                    df['Date'], format='%d%m%Y', errors='coerce'
                )
                df = df.dropna(subset=['Date'])
                df = df[df['Date'].dt.day_name() != 'Sunday']

                working_days_count = df['Date'].nunique()
                from_date = df['Date'].min()
                to_date = df['Date'].max()

                total_branch_offices = 4466
                total_sub_offices = 411

                branch_count = df[df['Office Type'] == 'BPO']['Office Name'].nunique()
                sub_count = df[df['Office Type'] == 'SPO']['Office Name'].nunique()

                col1, col2, col3 = st.columns(3)
                col1.metric("Working Days", working_days_count)
                col2.metric(
                    "Branch Offices with excess cash",
                    f"{branch_count} ({round(branch_count/total_branch_offices*100,2)}%)"
                )
                col3.metric(
                    "Sub Offices with excess cash",
                    f"{sub_count} ({round(sub_count/total_sub_offices*100,2)}%)"
                )

                st.subheader("Very High Risk Offices")

                risk_tables = {}

                for office in ['BPO', 'SPO']:
                    df_office = df[df['Office Type'] == office]
                    threshold = 100000 if office == 'BPO' else 500000

                    office_group = df_office.groupby(
                        ['Office Name', 'Division'],
                        as_index=False
                    ).agg(
                        Days_Exceeding_Threshold=('Excess Amount', lambda x: (x > threshold).sum()),
                        Avg_Excess_Above_Threshold=('Excess Amount', lambda x: x[x > threshold].mean())
                    )

                    office_group['Avg_Excess_Above_Threshold'] = office_group[
                        'Avg_Excess_Above_Threshold'
                    ].apply(
                        lambda x: f"{round(x/1e5,2)} L" if pd.notnull(x) else "0 L"
                    )

                    min_days = 0.9 * working_days_count
                    high_risk = office_group[
                        office_group['Days_Exceeding_Threshold'] >= min_days
                    ].copy()

                    high_risk['Office Type'] = office
                    high_risk['_avg'] = high_risk[
                        'Avg_Excess_Above_Threshold'
                    ].str.replace(' L', '').astype(float)

                    high_risk = high_risk.sort_values(
                        ['Days_Exceeding_Threshold', '_avg'],
                        ascending=[False, False]
                    ).drop(columns='_avg')

                    heading = (
                        "Very High Risk Branch Offices"
                        if office == 'BPO'
                        else "Very High Risk Sub Offices"
                    )

                    risk_tables[heading] = high_risk

                    with st.expander(f"{heading} ({len(high_risk)})"):
                        st.dataframe(
                            high_risk if not high_risk.empty
                            else pd.DataFrame({"Info": ["No offices found"]}),
                            use_container_width=True
                        )

                if risk_tables:
                    combined_df = pd.concat(risk_tables.values(), ignore_index=True)
                    combined_df['Remark'] = "Pending"

                    # From-To rows
                    from_to_df = pd.DataFrame({
                        'Office Name': [
                            f"From Date: {from_date.strftime('%d-%m-%Y')}",
                            f"To Date: {to_date.strftime('%d-%m-%Y')}"
                        ],
                        'Division': [None, None],
                        'Days_Exceeding_Threshold': [None, None],
                        'Avg_Excess_Above_Threshold': [None, None],
                        'Office Type': [None, None],
                        'Remark': [None, None]
                    })

                    # Last Updated IST
                    ist_time = datetime.utcnow() + timedelta(hours=5, minutes=30)
                    last_updated_df = pd.DataFrame({
                        'Office Name': [f"Last Updated (IST): {ist_time.strftime('%d-%m-%Y %H:%M:%S')}"],
                        'Division': [None],
                        'Days_Exceeding_Threshold': [None],
                        'Avg_Excess_Above_Threshold': [None],
                        'Office Type': [None],
                        'Remark': [None]
                    })

                    combined_export = pd.concat(
                        [combined_df, from_to_df, last_updated_df],
                        ignore_index=True
                    )

                    output_tab1 = BytesIO()
                    with pd.ExcelWriter(output_tab1, engine='xlsxwriter') as writer:
                        combined_export.to_excel(
                            writer, index=False, sheet_name="High_Risk_Offices"
                        )

                    file_name_tab1 = (
                        f"High_Risk_Offices_"
                        f"{from_date.strftime('%d%m%Y')}_to_"
                        f"{to_date.strftime('%d%m%Y')}.xlsx"
                    )

                for heading, table in risk_tables.items():
                    if not table.empty:
                        table['Office_Label'] = (
                            table['Office Name'] + " (" + table['Division'] + ")"
                        )
                        fig = px.bar(
                            table,
                            x='Office_Label',
                            y=table['Avg_Excess_Above_Threshold']
                              .str.replace(' L', '').astype(float),
                            color='Days_Exceeding_Threshold',
                            color_continuous_scale=['#f5b73b', '#5b1025'],
                            text='Avg_Excess_Above_Threshold',
                            title=f"{heading} – Avg Excess Cash (L)"
                        )
                        fig.update_layout(xaxis_tickangle=-45)
                        st.plotly_chart(fig, use_container_width=True)

                if risk_tables:
                    st.markdown("---")
                    st.download_button(
                        "📥 Download Very High Risk Offices as Excel",
                        data=output_tab1.getvalue(),
                        file_name=file_name_tab1,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    # ================================
    # TAB 2: Remittance Monitoring
    # ================================
    with tab2:

        st.markdown("""
        <div style="
            background-color:#14532d;
            padding:14px 18px;
            border-radius:10px;
            margin-bottom:14px;
        ">
        <h2 style="color:white; margin:0;">
        💸 REMITTANCE MONITORING
        </h2>
        <p style="color:#b7f0c2; margin:4px 0 0 0; font-size:0.9em;">
        Office-wise follow-up and remittance status update
        </p>
        </div>
        """ , unsafe_allow_html=True)

        st.subheader("Remittance Monitoring for High Risk Offices")

        if "play_sound" not in st.session_state:
            st.session_state.play_sound = False

        uploaded_remit = st.file_uploader(
            "Upload Excel exported from Tab 1",
            type=["xlsx"],
            key="tab2_upload"
        )

        if uploaded_remit:
            df = pd.read_excel(uploaded_remit)

            # Remove footer rows
            df = df[~df['Office Name'].astype(str).str.startswith(("From Date", "To Date", "Last Updated"))]
            df = df.reset_index(drop=True)

            df['Days_Exceeding_Threshold'] = df['Days_Exceeding_Threshold'].fillna(0).astype(int)
            df['Remark'] = df.get('Remark', 'Pending')

            branch_df = df[df['Office Type'] == 'BPO'].reset_index(drop=True)
            sub_df = df[df['Office Type'] == 'SPO'].reset_index(drop=True)

            remark_options = [
                "Pending",
                "Cash Remitted",
                "Balance lowered but cash not remitted"
            ]

            remark_colors = {
                "Pending": "#6b1f2b",
                "Cash Remitted": "#1f6b3b",
                "Balance lowered but cash not remitted": "#6b5a1f"
            }

            if "branch_remark" not in st.session_state:
                st.session_state.branch_remark = branch_df['Remark'].tolist()

            if "sub_remark" not in st.session_state:
                st.session_state.sub_remark = sub_df['Remark'].tolist()

            # ------------------------
            # SOUND
            # ------------------------
            if st.session_state.play_sound:
                st.components.v1.html("""
                <script>
                const audio = new Audio("https://actions.google.com/sounds/v1/alarms/beep_short.ogg");
                audio.play();
                </script>
                """, height=0)
                st.session_state.play_sound = False

            # ------------------------
            # STATUS TABLE
            # ------------------------
            def render_status_table(df, remark_key, title):
                st.markdown(f"### {title}")
                show = df.copy()
                show['Remark'] = st.session_state[remark_key]

                def color_rows(row):
                    bg = remark_colors.get(row['Remark'], "#2b2b2b")
                    return [f"background-color: {bg}; color: white"] * len(row)

                st.dataframe(show.style.apply(color_rows, axis=1), use_container_width=True)

            # ------------------------
            # OPTION 2 — COLUMN-ALIGNED CARDS
            # ------------------------
            def render_cards(df, remark_key, title):
                st.markdown(f"### {title}")

                for i in range(len(df)):
                    remark = st.session_state[remark_key][i]
                    bg = remark_colors.get(remark, "#2b2b2b")

                    st.markdown(
                        f"""
                        <div style="
                            border: 2px solid {bg};
                            border-radius: 10px;
                            padding: 15px;
                            margin-bottom: 6px;
                            background-color: rgba(255,255,255,0.02);
                        ">
                            <b>{df.loc[i, 'Office Name']}</b><br>
                            <span style="color: #cccccc; font-size: 0.85em;">
                                {df.loc[i, 'Division']}
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    new_val = st.selectbox(
                        "Update Remark",
                        remark_options,
                        index=remark_options.index(remark),
                        key=f"{remark_key}_{i}"
                    )

                    st.markdown(
                        """
                        <hr style='border: 1.5px solid #444444; margin: 10px 0 4px 0;'>
                        <hr style='border: 1.5px solid #444444; margin: 0 0 14px 0;'>
                        """,
                        unsafe_allow_html=True
                    )

                    if new_val != remark:
                        st.session_state[remark_key][i] = new_val
                        st.session_state.play_sound = True
                        st.rerun()

            render_status_table(branch_df, "branch_remark", "Branch Offices – Current Status")
            render_cards(branch_df, "branch_remark", "Update Remarks – Branch Offices")
            st.markdown("---")
            render_status_table(sub_df, "sub_remark", "Sub Offices – Current Status")
            render_cards(sub_df, "sub_remark", "Update Remarks – Sub Offices")

            # ------------------------
            # EXPORT UPDATED FILE
            # ------------------------
            final_df = pd.concat([
                branch_df.assign(Remark=st.session_state.branch_remark),
                sub_df.assign(Remark=st.session_state.sub_remark)
            ], ignore_index=True)

            # -----------------------------
            # Read From/To dates from uploaded file
            # -----------------------------
            from_date = None
            to_date = None
            for val in uploaded_remit:
                # Already read into df, check Office Name column
                for name in df['Office Name'].astype(str):
                    if name.startswith("From Date:"):
                        from_date = datetime.strptime(name.replace("From Date:", "").strip(), "%d-%m-%Y")
                    elif name.startswith("To Date:"):
                        to_date = datetime.strptime(name.replace("To Date:", "").strip(), "%d-%m-%Y")
            if from_date is None:
                from_date = datetime.utcnow() + timedelta(hours=5, minutes=30)
            if to_date is None:
                to_date = datetime.utcnow() + timedelta(hours=5, minutes=30)

            from_to_df = pd.DataFrame({
                'Office Name': [
                    f"From Date: {from_date.strftime('%d-%m-%Y')}",
                    f"To Date: {to_date.strftime('%d-%m-%Y')}"
                ],
                'Division': [None, None],
                'Days_Exceeding_Threshold': [None, None],
                'Avg_Excess_Above_Threshold': [None, None],
                'Office Type': [None, None],
                'Remark': [None, None]
            })

            ist_time = datetime.utcnow() + timedelta(hours=5, minutes=30)
            last_updated_df = pd.DataFrame({
                'Office Name': [f"Last Updated (IST): {ist_time.strftime('%d-%m-%Y %H:%M:%S')}"],
                'Division': [None],
                'Days_Exceeding_Threshold': [None],
                'Avg_Excess_Above_Threshold': [None],
                'Office Type': [None],
                'Remark': [None]
            })

            final_export = pd.concat(
                [final_df, from_to_df, last_updated_df],
                ignore_index=True
            )

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                final_export.to_excel(writer, index=False, sheet_name="Updated")

            st.download_button(
                "📥 Download Updated Remarks",
                data=output.getvalue(),
                file_name="High_Risk_Updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

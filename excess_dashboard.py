import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
import random

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
            st.success("Access Granted! Press Login button again to enter.")
else:
    st.title("Excess Cash Monitoring – Jabalpur Region")

    # -----------------------------
    # Tabs
    # -----------------------------
    tab1, tab2 = st.tabs(["Very High Risk Offices", "Remittance Monitoring"])

    # ================================
    # TAB 1: Very High Risk Offices
    # ================================
    with tab1:
        uploaded_file = st.file_uploader(
            "Select Excel File for High Risk Analysis",
            type=["xlsx"],
            key="tab1_upload"
        )

        if uploaded_file:
            df = pd.read_excel(uploaded_file)

            # Required columns
            required_columns = [
                "Date", "Division", "Office Type", "Office Name",
                "Office ID", "Max Amount", "Excess Amount", "Closing Balance"
            ]
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                st.error(f"Uploaded file is missing columns: {missing_cols}")
            else:
                # Process Date & Remove Sundays
                df['Date'] = pd.to_datetime(df['Date'], format='%d%m%Y', errors='coerce')
                df = df.dropna(subset=['Date'])
                df['Day_of_Week'] = df['Date'].dt.day_name()
                df = df[df['Day_of_Week'] != 'Sunday']
                working_days_count = df['Date'].nunique()
                from_date = df['Date'].min()
                to_date = df['Date'].max()

                # Summary Metrics
                total_branch_offices = 4466
                total_sub_offices = 411
                branch_count = len(df[df['Office Type']=='BPO']['Office Name'].unique())
                sub_count = len(df[df['Office Type']=='SPO']['Office Name'].unique())
                branch_percentage = round((branch_count / total_branch_offices) * 100, 2)
                sub_percentage = round((sub_count / total_sub_offices) * 100, 2)
                col1, col2, col3 = st.columns(3)
                col1.metric("Working Days", working_days_count)
                col2.metric("Branch Offices with excess cash", f"{branch_count} ({branch_percentage}%)")
                col3.metric("Sub Offices with excess cash", f"{sub_count} ({sub_percentage}%)")

                # Identify Very High Risk Offices
                st.subheader("Very High Risk Offices")
                risk_tables = {}
                for office in ['BPO','SPO']:
                    df_office = df[df['Office Type']==office]
                    threshold = 100000 if office=='BPO' else 500000
                    office_group = df_office.groupby(['Office Name','Division'], as_index=False).agg(
                        Days_Exceeding_Threshold=('Excess Amount', lambda x: (x>threshold).sum()),
                        Avg_Excess_Above_Threshold=('Excess Amount', lambda x: x[x>threshold].mean())
                    )
                    office_group['Avg_Excess_Above_Threshold'] = office_group['Avg_Excess_Above_Threshold'].apply(
                        lambda x: f"{round(x/1e5,2)} L" if pd.notnull(x) else "0 L"
                    )
                    min_days = 0.9 * working_days_count
                    high_risk = office_group[office_group['Days_Exceeding_Threshold'] >= min_days]
                    high_risk['Office Type'] = office
                    high_risk['_Avg'] = high_risk['Avg_Excess_Above_Threshold'].str.replace(' L','').astype(float)
                    high_risk = high_risk.sort_values(['Days_Exceeding_Threshold','_Avg'], ascending=[False,False]).reset_index(drop=True)
                    high_risk.drop(columns=['_Avg'], inplace=True)
                    high_risk = high_risk[['Office Name','Division','Days_Exceeding_Threshold','Avg_Excess_Above_Threshold','Office Type']]
                    heading = "Very High Risk Branch Offices" if office=='BPO' else "Very High Risk Sub Offices"
                    risk_tables[heading] = high_risk
                    with st.expander(f"{heading} ({len(high_risk)})"):
                        st.dataframe(high_risk if not high_risk.empty else pd.DataFrame({"Info":["No offices found"]}))

                # Export single sheet
                if risk_tables:
                    combined_df = pd.concat(risk_tables.values(), ignore_index=True)
                    combined_df['Remark'] = "Pending"

                    # From/To dates at bottom
                    from_to_df = pd.DataFrame({
                        'Office Name':[f"From Date: {from_date.strftime('%d-%m-%Y')}"],
                        'Division':[f"To Date: {to_date.strftime('%d-%m-%Y')}"],
                        'Days_Exceeding_Threshold':[None],
                        'Avg_Excess_Above_Threshold':[None],
                        'Office Type':[None],
                        'Remark':[None]
                    })

                    # Last updated
                    now = datetime.now()
                    last_updated_str = now.strftime("%d-%m-%Y %H:%M:%S")
                    last_updated_df = pd.DataFrame({
                        'Office Name':[f"Last Updated: {last_updated_str}"],
                        'Division':[None],
                        'Days_Exceeding_Threshold':[None],
                        'Avg_Excess_Above_Threshold':[None],
                        'Office Type':[None],
                        'Remark':[None]
                    })

                    combined_export = pd.concat([combined_df, from_to_df, last_updated_df], ignore_index=True)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        combined_export.to_excel(writer, sheet_name="High_Risk_Offices", index=False)

                    file_name_tab1 = f"High_Risk_Offices_{from_date.strftime('%d%m%Y')}_to_{to_date.strftime('%d%m%Y')}.xlsx"
                    st.download_button(
                        "📥 Download Very High Risk Offices as Excel",
                        data=output.getvalue(),
                        file_name=file_name_tab1,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Bar charts
                for heading, table in risk_tables.items():
                    if not table.empty:
                        table['Office_Label'] = table['Office Name'] + " (" + table['Division'] + ")"
                        fig = px.bar(
                            table,
                            x='Office_Label',
                            y=table['Avg_Excess_Above_Threshold'].str.replace(' L','').astype(float),
                            color='Days_Exceeding_Threshold',
                            color_continuous_scale=['#f5b73b','#5b1025'],
                            text='Avg_Excess_Above_Threshold',
                            title=f"{heading} - Avg Excess Cash in L"
                        )
                        fig.update_layout(xaxis_tickangle=-45)
                        st.plotly_chart(fig, use_container_width=True)

    # ================================
    # TAB 2: Remittance Monitoring
    # ================================
    with tab2:
        st.subheader("Remittance Monitoring for High Risk Offices")
        uploaded_remit = st.file_uploader("Upload Excel exported from Tab 1", type=["xlsx"], key="tab2_upload")
        if uploaded_remit:
            remit_df = pd.read_excel(uploaded_remit)

            # Remove From/To/LastUpdated rows
            if 'From Date' in str(remit_df.iloc[-2,0]):
                remit_df = remit_df.iloc[:-2].reset_index(drop=True)

            if 'Remark' not in remit_df.columns:
                remit_df['Remark'] = "Pending"

            remark_options = ["Pending","Cash Remitted","Balance lowered but cash not remitted"]

            # Function to color rows
            def color_rows(row):
                if row["Remark"] == "Pending":
                    return ["background-color: #800000; color: white"]*len(row)
                elif row["Remark"] == "Cash Remitted":
                    return ["background-color: #008000; color: white"]*len(row)
                elif row["Remark"] == "Balance lowered but cash not remitted":
                    return ["background-color: #FFD700; color: black"]*len(row)
                else:
                    return [""]*len(row)

            # Editable table using st.data_editor
            edited_df = st.data_editor(
                remit_df,
                column_config=None,
                disabled=False,
                hide_index=True,
                use_container_width=True,
                row_config={"style": color_rows}
            )

            # Play sound if any remark changed
            if not edited_df['Remark'].equals(remit_df['Remark']):
                rand_suffix = random.randint(1,100000)
                st.components.v1.html(f"""
                <audio autoplay>
                    <source src="https://actions.google.com/sounds/v1/alarms/beep_short.ogg?{rand_suffix}" type="audio/ogg">
                </audio>
                """, height=0)

            # Append From/To and LastUpdated
            from_date = datetime.today()
            to_date = datetime.today()
            from_to_df = pd.DataFrame({
                'Office Name':[f"From Date: {from_date.strftime('%d-%m-%Y')}"],
                'Division':[f"To Date: {to_date.strftime('%d-%m-%Y')}"],
                'Days_Exceeding_Threshold':[None],
                'Avg_Excess_Above_Threshold':[None],
                'Office Type':[None],
                'Remark':[None]
            })
            now = datetime.now()
            last_updated_str = now.strftime("%d-%m-%Y %H:%M:%S")
            last_updated_df = pd.DataFrame({
                'Office Name':[f"Last Updated: {last_updated_str}"],
                'Division':[None],
                'Days_Exceeding_Threshold':[None],
                'Avg_Excess_Above_Threshold':[None],
                'Office Type':[None],
                'Remark':[None]
            })

            combined_export = pd.concat([edited_df, from_to_df, last_updated_df], ignore_index=True)
            output2 = BytesIO()
            with pd.ExcelWriter(output2, engine='xlsxwriter') as writer:
                combined_export.to_excel(writer, sheet_name="High_Risk_Updated", index=False)

            file_name_tab2 = f"High_Risk_Updated_{from_date.strftime('%d%m%Y')}_to_{to_date.strftime('%d%m%Y')}.xlsx"
            st.download_button(
                "📥 Download Updated High Risk Offices with Remarks",
                data=output2.getvalue(),
                file_name=file_name_tab2,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

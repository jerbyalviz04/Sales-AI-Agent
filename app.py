import streamlit as st
import pandas as pd
import speech_recognition as sr
from openai import OpenAI
import matplotlib.pyplot as plt

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent Dashboard", layout="wide")
st.title("📊 AI-Powered Sales Performance Dashboard")

st.markdown("### Hi Jerby! How can I help you today? 😄")

file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

st.sidebar.header("Search for Outlet")
search_input = st.sidebar.text_input("Enter Outlet ID or Name")

if st.sidebar.button("🎤 Start Voice Input"):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        st.sidebar.info("Listening... Speak now")
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=15)
            text = recognizer.recognize_google(audio)
            search_input = text.strip()
            st.sidebar.success(f"Captured Search: {search_input}")
        except sr.WaitTimeoutError:
            st.sidebar.warning("Listening timed out. Please try again.")
        except sr.UnknownValueError:
            st.sidebar.warning("Could not understand the audio.")
        except sr.RequestError as e:
            st.sidebar.error(f"Speech recognition error: {e}")

if search_input:
    found = False
    for sheet_name, df in sheets.items():
        if 'Outlet ID' in df.columns:
            matched_rows = df[
                (df['Outlet ID'].astype(str).str.strip() == search_input.strip()) |
                (df['Outlet Name'].astype(str).str.strip().str.lower() == search_input.strip().lower())
            ]
            if not matched_rows.empty:
                found = True
                for _, outlet_row in matched_rows.iterrows():
                    st.subheader(f"Performance Dashboard for: {outlet_row['Outlet Name']} ({outlet_row['Outlet ID']})")

                    st.markdown(f"**Head Office Name:** {outlet_row.get('Head Office Name', 'N/A')}")
                    head_office = outlet_row.get('Head Office Name', None)
                    if head_office:
                        total_branches = df[df['Head Office Name'].astype(str).str.strip() == str(head_office).strip()].shape[0]
                        st.markdown(f"**Total Branches under this Head Office:** {total_branches}")

                    st.markdown(f"**Channel:** {outlet_row.get('Customer Channel', 'N/A')}")
                    st.markdown(f"**Segment:** {outlet_row.get('Customer Segment', 'N/A')}")
                    st.markdown(f"**Status:** {outlet_row.get('Customer Status', 'N/A')}")
                    st.markdown(f"**Warehouse:** {outlet_row.get('Warehouse', 'N/A')}")

                    sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
                    sales_data = outlet_row[sales_cols].fillna(0)
                    sales_df = pd.DataFrame({
                        'Month': sales_cols,
                        'Sales': sales_data.values
                    })
                    sales_df['Sales'] = sales_df['Sales'].map('{:,.0f}'.format)

                    st.subheader("📈 Monthly Sales Trend")
                    st.dataframe(sales_df)

                    # LRB Sales monthly trend chart if sheet exists
                    if "LRB Sales" in sheets:
                        lrb_df = sheets["LRB Sales"]
                        lrb_match = lrb_df[
                            (lrb_df['Outlet ID'].astype(str).str.strip() == str(outlet_row['Outlet ID']).strip())
                        ]
                        if not lrb_match.empty:
                            st.subheader("📊 LRB Sales Monthly Trend")
                            lrb_sales_cols = sorted([col for col in lrb_df.columns if col[:4].isdigit() and len(col) == 7])
                            lrb_sales_data = lrb_match.iloc[0][lrb_sales_cols].fillna(0)

                            fig, ax = plt.subplots(figsize=(10, 4))
                            ax.plot(lrb_sales_cols, lrb_sales_data.values, marker='o')
                            ax.set_title("LRB Sales Trend")
                            ax.set_xlabel("Month")
                            ax.yaxis.set_visible(False)
                            ax.tick_params(axis='x', rotation=45)

                            for x, y in zip(lrb_sales_cols, lrb_sales_data.values):
                                ax.text(x, y, f"{y:,.0f}", ha='center', va='bottom', fontsize=9)

                            st.pyplot(fig)

                    question = st.text_input("Ask a sales-related question about this outlet:")
                    if st.button("Get AI Answer"):
                        context = f"Outlet: {outlet_row.get('Outlet Name', 'N/A')} | Channel: {outlet_row.get('Customer Channel', 'N/A')} | Segment: {outlet_row.get('Customer Segment', 'N/A')} | Status: {outlet_row.get('Customer Status', 'N/A')} | Warehouse: {outlet_row.get('Warehouse', 'N/A')}"
                        top_sales = sales_df.to_csv(index=False)

                        prompt = (
                            f"You are a sales data analyst.\n"
                            f"Context: {context}\n"
                            f"Here are the monthly sales data:\n{top_sales}\n"
                            f"Question: {question}\n"
                            f"Provide a detailed and precise answer based on the data."
                        )
                        try:
                            response = client.chat.completions.create(
                                model="gpt-3.5-turbo",
                                messages=[{"role": "user", "content": prompt}],
                                temperature=0
                            )
                            st.markdown("### 🧠 AI Agent Answer")
                            st.write(response.choices[0].message.content.strip())
                        except Exception as e:
                            st.error(f"OpenAI API error: {e}")
                break
    if not found:
        st.warning(f"No outlet matching '{search_input}' found in any sheet.")

    # Show Category Summary
    st.subheader("📦 Category Summary")

    current_month = "2025-06"
    previous_year = "2024"
    current_year = "2025"

    summary_data = []

    for sheet_name, df in sheets.items():
        if 'Outlet ID' not in df.columns:
            continue

        sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
        current_month_col = current_month
        last_year_month_col = current_month.replace(current_year, previous_year)

        if current_month_col not in df.columns or last_year_month_col not in df.columns:
            continue

        df_outlet = df[
            (df['Outlet ID'].astype(str).str.strip() == search_input.strip()) |
            (df['Outlet Name'].astype(str).str.strip().str.lower() == search_input.strip().lower())
        ]

        if df_outlet.empty:
            continue

        current_month_sales = df_outlet[current_month_col].values[0]
        last_year_month_sales = df_outlet[last_year_month_col].values[0]
        if last_year_month_sales == 0:
            month_growth = 0.0
        else:
            month_growth = (current_month_sales / last_year_month_sales - 1) * 100

        ytd_months_current = [f"{current_year}-{str(m).zfill(2)}" for m in range(1, int(current_month[-2:]) + 1)]
        ytd_months_previous = [f"{previous_year}-{str(m).zfill(2)}" for m in range(1, int(current_month[-2:]) + 1)]

        ytd_months_current = [m for m in ytd_months_current if m in df_outlet.columns]
        ytd_months_previous = [m for m in ytd_months_previous if m in df_outlet.columns]

        ytd_current = df_outlet[ytd_months_current].sum(axis=1).values[0]
        ytd_previous = df_outlet[ytd_months_previous].sum(axis=1).values[0]

        if ytd_previous == 0:
            ytd_growth = 0.0
        else:
            ytd_growth = (ytd_current / ytd_previous - 1) * 100

        zero_sales_outlets = 0
        for idx, row in df.iterrows():
            if row.get(current_month_col, 0) == 0:
                prev_sales = []
                for offset in range(1, 4):
                    month_int = int(current_month[-2:]) - offset
                    if month_int > 0:
                        prev_month = f"{current_year}-{str(month_int).zfill(2)}"
                        if prev_month in df.columns:
                            prev_sales.append(row.get(prev_month, 0) > 0)
                if sum(prev_sales) >= 2:
                    zero_sales_outlets += 1

        summary_data.append({
            "Category": sheet_name,
            f"{current_month} Growth % vs LY": f"{month_growth:.1f}%",
            "YTD Growth %": f"{ytd_growth:.1f}%",
            "Zero Sales Outlets": zero_sales_outlets
        })

    if summary_data:
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df)

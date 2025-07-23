import streamlit as st
import pandas as pd
import speech_recognition as sr
from openai import OpenAI
import matplotlib.pyplot as plt

# Use OpenAI API key from Streamlit secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent", layout="wide")
st.title("💬 Sales AI Agent Chat")

# Load the Excel file from the repo directory
file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

with st.chat_message("assistant"):
    st.markdown("Hi Jerby! 👋 How can I help you today?")

user_query = st.chat_input("Ask a question about a specific outlet (ID or name)...")

if user_query:
    st.chat_message("user").markdown(user_query)

    found = False
    for sheet_name, df in sheets.items():
        if 'Outlet ID' in df.columns:
            matched_rows = df[
                (df['Outlet ID'].astype(str).str.strip() == user_query.strip()) |
                (df['Outlet Name'].astype(str).str.strip().str.lower() == user_query.strip().lower())
            ]
            if not matched_rows.empty:
                found = True
                for _, outlet_row in matched_rows.iterrows():
                    context = f"Outlet: {outlet_row.get('Outlet Name', 'N/A')} | Channel: {outlet_row.get('Customer Channel', 'N/A')} | Segment: {outlet_row.get('Customer Segment', 'N/A')} | Status: {outlet_row.get('Customer Status', 'N/A')} | Warehouse: {outlet_row.get('Warehouse', 'N/A')}"

                    st.markdown(f"**{context}**")

                    category_summaries = []

                    for category in ['LRB Sales', 'CSD Sales', 'Water Sales']:
                        if category in sheets:
                            cat_df = sheets[category]
                            cat_row = cat_df[cat_df['Outlet ID'].astype(str).str.strip() == outlet_row['Outlet ID']]
                            if not cat_row.empty:
                                cat_row = cat_row.iloc[0]
                                current_year = '2025'
                                previous_year = '2024'

                                june_2025 = cat_row.get('2025-06', 0)
                                june_2024 = cat_row.get('2024-06', 0)
                                ytd_2025 = cat_row[[f"2025-{str(m).zfill(2)}" for m in range(1, 7)]].sum()
                                ytd_2024 = cat_row[[f"2024-{str(m).zfill(2)}" for m in range(1, 7)]].sum()

                                june_growth = ((june_2025 - june_2024) / june_2024) * 100 if june_2024 else 0
                                ytd_growth = ((ytd_2025 - ytd_2024) / ytd_2024) * 100 if ytd_2024 else 0

                                last_3_months = [f"2025-{str(m).zfill(2)}" for m in [4, 5, 6] if f"2025-{str(m).zfill(2)}" in cat_row.index]
                                zero_sales = cat_row.get("2025-06", 0) == 0
                                recent_sales = cat_row[last_3_months[:-1]].fillna(0) > 0
                                zero_flag = "Yes" if zero_sales and recent_sales.sum() >= 2 else "No"

                                summary_df = pd.DataFrame({
                                    ' ': ['June 2025', 'June 2024', 'June Growth %', 'YTD Growth %', 'Zero Sales Flag'],
                                    category: [f"{june_2025:,.0f}", f"{june_2024:,.0f}", f"{june_growth:.1f}%", f"{ytd_growth:.1f}%", zero_flag]
                                })
                                st.subheader(f"📊 {category} Category Summary")
                                st.table(summary_df.set_index(' '))

                                sales_cols = [col for col in cat_df.columns if '-' in col and cat_row.get(col) is not None]
                                trend_df = pd.DataFrame({
                                    'Month': sales_cols,
                                    'Sales': [cat_row[col] for col in sales_cols]
                                })
                                trend_df['Sales'] = pd.to_numeric(trend_df['Sales'], errors='coerce')
                                trend_df.dropna(inplace=True)

                                st.markdown(f"### Monthly {category} Sales Trend")
                                st.dataframe(trend_df)

                                fig, ax = plt.subplots(figsize=(10, 3))
                                ax.plot(trend_df['Month'], trend_df['Sales'], marker='o')
                                for i, val in enumerate(trend_df['Sales']):
                                    ax.text(i, val, f"{val:,.0f}", ha='center', va='bottom', fontsize=8)
                                ax.set_xticks(range(len(trend_df['Month'])))
                                ax.set_xticklabels(trend_df['Month'], rotation=45)
                                ax.set_ylabel("Sales")
                                ax.set_title(f"Monthly {category} Sales Trend")
                                ax.spines['left'].set_visible(False)
                                st.pyplot(fig)

                    combined_sales = []
                    for sheet_name, df in sheets.items():
                        if 'Outlet ID' in df.columns and outlet_row['Outlet ID'] in df['Outlet ID'].astype(str).values:
                            row = df[df['Outlet ID'].astype(str) == str(outlet_row['Outlet ID'])].iloc[0]
                            sales_cols = [col for col in df.columns if '-' in col]
                            row_sales = pd.DataFrame({
                                'Month': sales_cols,
                                'Sales': row[sales_cols].values
                            })
                            row_sales['Sales'] = pd.to_numeric(row_sales['Sales'], errors='coerce')
                            row_sales.dropna(inplace=True)
                            row_sales['Sales'] = row_sales['Sales'].apply(lambda x: f"{x:,.0f}")
                            combined_sales.append(row_sales)

                    if combined_sales:
                        merged = pd.concat(combined_sales)
                        merged_summary = merged.to_csv(index=False)

                        prompt = (
                            f"You are a sales data analyst.\n"
                            f"Context: {context}\n"
                            f"Here are the monthly sales data:\n{merged_summary}\n"
                            f"Provide a summary of sales performance for this outlet."
                        )

                        try:
                            response = client.chat.completions.create(
                                model="gpt-3.5-turbo",
                                messages=[{"role": "user", "content": prompt}],
                                temperature=0
                            )
                            ai_answer = response.choices[0].message.content.strip()
                            st.chat_message("assistant").markdown(ai_answer)
                        except Exception as e:
                            st.chat_message("assistant").markdown(f"⚠️ OpenAI API error: {e}")
                break

    if not found:
        st.chat_message("assistant").markdown(f"🔍 No outlet matching '{user_query}' found in any sheet.")

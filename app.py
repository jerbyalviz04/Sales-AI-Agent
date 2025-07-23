import streamlit as st
import pandas as pd
import speech_recognition as sr
from openai import OpenAI

client = OpenAI(api_key="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")

st.set_page_config(page_title="Sales AI Agent", layout="wide")
st.title("💬 Sales AI Agent Chat")

file_path = r"C:\\Users\\41419\\Desktop\\Insights\\AI Test Folder\\MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None)
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

                    sales_cols = [col for col in df.columns if '-' in col]
                    sales_data = outlet_row[sales_cols]
                    sales_df = pd.DataFrame({
                        'Month': sales_cols,
                        'Sales': sales_data.values
                    })

                    top_sales = sales_df.to_csv(index=False)

                    prompt = (
                        f"You are a sales data analyst.\n"
                        f"Context: {context}\n"
                        f"Here are the monthly sales data:\n{top_sales}\n"
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

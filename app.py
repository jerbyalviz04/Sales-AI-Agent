import streamlit as st
import pandas as pd
from openai import OpenAI
import matplotlib.pyplot as plt

# Secure OpenAI API access from secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent", layout="wide")
st.title("🤖 Sales AI Agent Dashboard")

# Load Excel data
file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"❌ Failed to load data: {e}")
    st.stop()

# Initialize session state for chat history
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# Assistant greeting
with st.chat_message("assistant"):
    st.markdown("Hi Jerby! 👋 How can I help you today?")

# User input
user_query = st.chat_input("Ask a question about a specific outlet (ID or name)...")

if user_query:
    st.chat_message("user").markdown(user_query)
    found = False

    for sheet_name, df in sheets.items():
        if 'Outlet ID' in df.columns:
            matched = df[
                (df['Outlet ID'].astype(str).str.strip() == user_query.strip()) |
                (df['Outlet Name'].astype(str).str.strip().str.lower() == user_query.strip().lower())
            ]
            if not matched.empty:
                found = True
                for _, row in matched.iterrows():
                    context = (
                        f"Outlet: {row.get('Outlet Name', 'N/A')} | "
                        f"Channel: {row.get('Customer Channel', 'N/A')} | "
                        f"Segment: {row.get('Customer Segment', 'N/A')} | "
                        f"Status: {row.get('Customer Status', 'N/A')} | "
                        f"Warehouse: {row.get('Warehouse', 'N/A')}"
                    )

                    sales_cols = [col for col in df.columns if '-' in col]
                    sales_data = row[sales_cols]
                    sales_df = pd.DataFrame({'Month': sales_cols, 'Sales': sales_data.values})
                    sales_df['Sales'] = pd.to_numeric(sales_df['Sales'], errors='coerce')

                    # Format values for prompt
                    prompt_df = sales_df.copy()
                    prompt_df['Sales'] = prompt_df['Sales'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "0")
                    csv_sales = prompt_df.to_csv(index=False)

                    # Build prompt for OpenAI
                    prompt = (
                        f"You are a sales analyst.\n"
                        f"Context: {context}\n"
                        f"Here is the sales data:\n{csv_sales}\n"
                        f"Summarize the sales performance of this outlet."
                    )

                    # Get AI response
                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.2
                        )
                        reply = response.choices[0].message.content.strip()
                        st.chat_message("assistant").markdown(reply)

                        # Plot chart
                        clean_df = sales_df.dropna()
                        fig, ax = plt.subplots()
                        ax.plot(clean_df['Month'], clean_df['Sales'], marker='o')
                        for i, val in enumerate(clean_df['Sales']):
                            ax.text(i, val, f"{val:,.0f}", ha='center', va='bottom', fontsize=8)
                        ax.set_xticks(range(len(clean_df['Month'])))
                        ax.set_xticklabels(clean_df['Month'], rotation=45)
                        ax.set_title("📈 Monthly LRB Sales Trend")
                        ax.spines['left'].set_visible(False)
                        ax.set_ylabel("")  # Optional: remove axis label
                        st.pyplot(fig)

                    except Exception as e:
                        st.chat_message("assistant").markdown(f"⚠️ OpenAI API error: `{e}`")
                break

    if not found:
        st.chat_message("assistant").markdown(f"🚫 No outlet found for: `{user_query}`")

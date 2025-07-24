import streamlit as st
import pandas as pd
import speech_recognition as sr
from openai import OpenAI
import matplotlib.pyplot as plt

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent Chat", layout="wide")
st.title("💬 Sales AI Agent Chat")

with st.chat_message("assistant"):
    st.markdown("Hi Jerby! 👋 How can I help you today?")

file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

# Chat Input Section
search_input = st.chat_input("Enter Outlet ID or Name or Ask a Question")

# Voice Input Button
if st.button("🎤 Use Voice Input"):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        st.info("🎙 Listening... Speak now")
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=15)
            text = recognizer.recognize_google(audio)
            search_input = text.strip()
            st.success(f"🎧 You said: {search_input}")
        except sr.WaitTimeoutError:
            st.warning("⌛ Listening timed out. Please try again.")
        except sr.UnknownValueError:
            st.warning("🤷 Could not understand the audio.")
        except sr.RequestError as e:
            st.error(f"🔌 Speech recognition error: {e}")

if search_input:
    with st.chat_message("user"):
        st.markdown(search_input)

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
                    with st.chat_message("assistant"):
                        st.subheader(f"📍 Outlet: {outlet_row['Outlet Name']} ({outlet_row['Outlet ID']})")
                        st.markdown(f"**Head Office:** {outlet_row.get('Head Office Name', 'N/A')}")

                        head_office = outlet_row.get('Head Office Name', None)
                        if head_office:
                            total_branches = df[df['Head Office Name'].astype(str).str.strip() == str(head_office).strip()].shape[0]
                            st.markdown(f"**Total Branches:** {total_branches}")

                        st.markdown(f"**Channel:** {outlet_row.get('Customer Channel', 'N/A')}  ")
                        st.markdown(f"**Segment:** {outlet_row.get('Customer Segment', 'N/A')}  ")
                        st.markdown(f"**Status:** {outlet_row.get('Customer Status', 'N/A')}  ")
                        st.markdown(f"**Warehouse:** {outlet_row.get('Warehouse', 'N/A')}  ")

                        sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
                        sales_data = outlet_row[sales_cols].fillna(0)
                        sales_df = pd.DataFrame({
                            'Month': sales_cols,
                            'Sales': sales_data.values
                        })
                        st.subheader("📈 Monthly Sales")
                        st.dataframe(sales_df.style.format({"Sales": "{:,.0f}"}))

                        if "LRB Sales" in sheets:
                            lrb_df = sheets["LRB Sales"]
                            lrb_match = lrb_df[(lrb_df['Outlet ID'].astype(str).str.strip() == str(outlet_row['Outlet ID']).strip())]
                            if not lrb_match.empty:
                                st.subheader("📊 LRB Sales Trend")
                                lrb_sales_cols = sorted([col for col in lrb_df.columns if col[:4].isdigit() and len(col) == 7])
                                lrb_sales_data = lrb_match.iloc[0][lrb_sales_cols].fillna(0)
                                fig, ax = plt.subplots(figsize=(10, 3))
                                ax.plot(lrb_sales_cols, lrb_sales_data.values, marker='o')
                                ax.set_title("LRB Sales")
                                ax.tick_params(axis='x', rotation=45)
                                ax.yaxis.set_visible(False)
                                for x, y in zip(lrb_sales_cols, lrb_sales_data.values):
                                    ax.text(x, y, f"{y:,.0f}", ha='center', va='bottom', fontsize=8)
                                st.pyplot(fig)

                        question = st.text_input("Ask something about this outlet")
                        if st.button("Get AI Answer"):
                            context = f"Outlet: {outlet_row.get('Outlet Name')} | Channel: {outlet_row.get('Customer Channel')} | Segment: {outlet_row.get('Customer Segment')}"
                            top_sales = sales_df.to_csv(index=False)
                            prompt = (
                                f"You are a sales analyst.\nContext: {context}\nSales Data:\n{top_sales}\nQuestion: {question}\nAnswer with insights."
                            )
                            try:
                                response = client.chat.completions.create(
                                    model="gpt-3.5-turbo",
                                    messages=[{"role": "user", "content": prompt}],
                                    temperature=0
                                )
                                st.markdown("### 💡 AI Agent Says:")
                                st.write(response.choices[0].message.content.strip())
                            except Exception as e:
                                st.error(f"OpenAI API Error: {e}")
                break
    if not found:
        st.warning(f"No outlet found for '{search_input}'")

    # Category Summary
    with st.chat_message("assistant"):
        st.subheader("📦 LRB Sales Category Summary")

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

            current_sales = df_outlet[current_month_col].values[0]
            last_year_sales = df_outlet[last_year_month_col].values[0]
            month_growth = ((current_sales / last_year_sales - 1) * 100) if last_year_sales != 0 else 0.0

            ytd_current = df_outlet[[m for m in sales_cols if m.startswith(current_year)]].sum(axis=1).values[0]
            ytd_previous = df_outlet[[m for m in sales_cols if m.startswith(previous_year)]].sum(axis=1).values[0]
            ytd_growth = ((ytd_current / ytd_previous - 1) * 100) if ytd_previous != 0 else 0.0

            # Zero sales logic
            zero_sales_outlets = 0
            for idx, row in df.iterrows():
                if row.get(current_month_col, 0) == 0:
                    prev_sales = [
                        row.get(f"{current_year}-{str(m).zfill(2)}", 0) > 0 for m in range(int(current_month[-2:]) - 3, int(current_month[-2:]))
                    ]
                    if sum(prev_sales) >= 2:
                        zero_sales_outlets += 1

            summary_data.append({
                "Category": sheet_name,
                f"{current_month} vs LY": f"{month_growth:.1f}%",
                "YTD Growth": f"{ytd_growth:.1f}%",
                "Zero Sales Outlets": zero_sales_outlets
            })

        if summary_data:
            st.dataframe(pd.DataFrame(summary_data))

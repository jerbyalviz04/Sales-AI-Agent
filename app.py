import streamlit as st
import pandas as pd
from openai import OpenAI
import matplotlib.pyplot as plt

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent", layout="wide")

st.markdown("## 💬 Sales AI Agent Chat")
st.markdown("Hi Jerby! 👋 How can I help you today?")
st.markdown("Enter Outlet ID or Name (use voice input via your keyboard mic)")

file_path = "MT Sales Raw Data.xlsx"

try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

user_input = st.text_input("")

def normalize(text):
    return str(text).strip().lower().replace(" ", "")

if user_input:
    norm_input = normalize(user_input)
    found = False

    for sheet_name, df in sheets.items():
        if "Outlet ID" not in df.columns or "Outlet Name" not in df.columns:
            continue

        df["__id_norm"] = df["Outlet ID"].astype(str).apply(normalize)
        df["__name_norm"] = df["Outlet Name"].astype(str).apply(normalize)

        matched = df[(df["__id_norm"] == norm_input) | (df["__name_norm"] == norm_input)]

        if not matched.empty:
            found = True
            for _, row in matched.iterrows():
                st.markdown(f"### 🏪 Outlet: {row['Outlet Name']} ({row['Outlet ID']})")

                st.markdown(f"**Head Office:** {row.get('Head Office Name', 'N/A')}")
                ho_name = row.get('Head Office Name', None)
                if ho_name:
                    total_branches = df[df['Head Office Name'].astype(str).str.strip() == str(ho_name).strip()].shape[0]
                    st.markdown(f"**Total Branches under Head Office:** {total_branches}")
                st.markdown(f"**Channel:** {row.get('Customer Channel', 'N/A')}")
                st.markdown(f"**Segment:** {row.get('Customer Segment', 'N/A')}")
                st.markdown(f"**Status:** {row.get('Customer Status', 'N/A')}")
                st.markdown(f"**Warehouse:** {row.get('Warehouse', 'N/A')}")

                sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
                sales_data = row[sales_cols].fillna(0)
                sales_df = pd.DataFrame({"Month": sales_cols, "Sales": sales_data.values})
                sales_df["Sales"] = sales_df["Sales"].map("{:,.0f}".format)

                st.markdown("#### 📈 Monthly Sales Trend")
                st.dataframe(sales_df)

                # LRB Sales Chart
                if "LRB Sales" in sheets:
                    lrb_df = sheets["LRB Sales"]
                    lrb_df["__id_norm"] = lrb_df["Outlet ID"].astype(str).apply(normalize)
                    lrb_match = lrb_df[lrb_df["__id_norm"] == normalize(row["Outlet ID"])]
                    if not lrb_match.empty:
                        lrb_cols = sorted([col for col in lrb_df.columns if col[:4].isdigit()])
                        lrb_sales = lrb_match.iloc[0][lrb_cols].fillna(0)

                        fig, ax = plt.subplots(figsize=(10, 3))
                        ax.plot(lrb_cols, lrb_sales.values, marker="o")
                        ax.set_title("LRB Monthly Sales Trend")
                        ax.set_xlabel("Month")
                        ax.tick_params(axis='x', rotation=45)
                        ax.yaxis.set_visible(False)

                        for x, y in zip(lrb_cols, lrb_sales.values):
                            ax.text(x, y, f"{y:,.0f}", ha="center", va="bottom", fontsize=8)
                        st.pyplot(fig)

                question = st.text_input("❓ Ask a question about this outlet:")
                if st.button("💡 Get AI Answer"):
                    context = (
                        f"Outlet: {row.get('Outlet Name')} | Channel: {row.get('Customer Channel')} | "
                        f"Segment: {row.get('Customer Segment')} | Warehouse: {row.get('Warehouse')}"
                    )
                    sales_csv = sales_df.to_csv(index=False)

                    prompt = (
                        f"You are a smart sales data analyst.\n"
                        f"Outlet Details: {context}\n"
                        f"Sales Data:\n{sales_csv}\n"
                        f"Question: {question}\n"
                        f"Answer in clear, structured format."
                    )
                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.2
                        )
                        st.markdown("### 🧠 AI Response")
                        st.write(response.choices[0].message.content.strip())
                    except Exception as e:
                        st.error(f"OpenAI API Error: {e}")
            break

    if not found:
        st.warning(f"No outlet matching '{user_input}' found.")

    # Category Summary
    st.markdown("### 📦 Category Summary")
    month = "2025-06"
    year = "2025"
    last_year = "2024"

    summary = []
    for name, df in sheets.items():
        if "Outlet ID" not in df.columns or month not in df.columns:
            continue

        df["__id_norm"] = df["Outlet ID"].astype(str).apply(normalize)
        df["__name_norm"] = df["Outlet Name"].astype(str).apply(normalize)
        outlet_df = df[(df["__id_norm"] == norm_input) | (df["__name_norm"] == norm_input)]

        if outlet_df.empty:
            continue

        current = outlet_df[month].values[0]
        prev = outlet_df[month.replace(year, last_year)].values[0] if month.replace(year, last_year) in df.columns else 0
        growth = (current / prev - 1) * 100 if prev != 0 else 0

        ytd_months = [f"{year}-{str(i).zfill(2)}" for i in range(1, 7)]
        ytd_last = [f"{last_year}-{str(i).zfill(2)}" for i in range(1, 7)]
        ytd_cur = outlet_df[ytd_months].sum(axis=1).values[0] if all(m in df.columns for m in ytd_months) else 0
        ytd_old = outlet_df[ytd_last].sum(axis=1).values[0] if all(m in df.columns for m in ytd_last) else 0
        ytd_growth = (ytd_cur / ytd_old - 1) * 100 if ytd_old != 0 else 0

        zero = 0
        for _, r in df.iterrows():
            if r.get(month, 0) == 0:
                count = sum([r.get(f"{year}-{str(m).zfill(2)}", 0) > 0 for m in range(3, 6)])
                if count >= 2:
                    zero += 1

        summary.append({
            "Category": name,
            f"{month} Growth % vs LY": f"{growth:.1f}%",
            "YTD Growth %": f"{ytd_growth:.1f}%",
            "Zero Sales Outlets": zero
        })

    if summary:
        st.dataframe(pd.DataFrame(summary))

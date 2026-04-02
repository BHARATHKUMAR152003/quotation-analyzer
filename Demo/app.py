import streamlit as st
import pandas as pd

st.title("📊 Quotation Conversion Analyzer")

# File uploads
input_file = st.file_uploader("Upload Input Excel", type=["xlsx"])
remove_file = st.file_uploader("Upload Remove CSV", type=["csv"])

if input_file and remove_file:

    df = pd.read_excel(input_file)
    remove_df = pd.read_csv(remove_file)

    # Clean data
    df['CustomerName'] = df['CustomerName'].astype(str).str.strip().str.lower()
    remove_df['CustomerName'] = remove_df['CustomerName'].astype(str).str.strip().str.lower()

    # Remove matching customers
    df = df[~df['CustomerName'].isin(remove_df['CustomerName'])]

    # Clean Converted column
    df['Converted?'] = df['Converted?'].astype(str).str.strip().str.capitalize()

    # Grouping
    summary = df.groupby('QuotationNumber')['Converted?'].value_counts().unstack(fill_value=0)

    for col in ['Yes', 'No']:
        if col not in summary.columns:
            summary[col] = 0

    summary = summary.reset_index()
    summary = summary.rename(columns={
        'QuotationNumber': 'QuotationNumber',
        'Yes': 'Converted Yes',
        'No': 'Converted No'
    })

    summary['Total'] = summary['Converted Yes'] + summary['Converted No']
    summary['Conversion %'] = (summary['Converted Yes'] / summary['Total']) * 100

    # KPIs
    yes_count = (summary['Converted Yes'] > 0).sum()
    no_count = (summary['Converted No'] > 0).sum()

    st.subheader("📌 Key Metrics")
    st.write(f"✔ Converted Yes Count: {yes_count}")
    st.write(f"✔ Converted No Count: {no_count}")

    st.subheader("📊 Summary Table")
    st.dataframe(summary)

    # 🤖 AI Insights
    st.subheader("🤖 AI Insights")

    conversion_rate = (summary['Converted Yes'].sum() / summary['Total'].sum()) * 100

    if conversion_rate > 70:
        st.success("High conversion rate! Sales strategy is effective.")
    elif conversion_rate > 40:
        st.warning("Moderate conversion rate. Improvement possible.")
    else:
        st.error("Low conversion rate. Needs immediate attention.")

    if no_count > yes_count:
        st.warning("More failures than success. Investigate lost opportunities.")
    else:
        st.success("Conversions are performing well!")

    # Download button
    csv = summary.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download Output", csv, "output.csv", "text/csv")
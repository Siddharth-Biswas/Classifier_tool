import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Product Classifier", layout="wide")
st.title("🧠 Product Classifier Tool for Flywheel")
st.markdown("Upload your product details and rules to auto-classify using AI-style logic.")

product_file = st.file_uploader("📦 Upload Product Details (Excel)", type=["xlsx"])
rules_file = st.file_uploader("📋 Upload Classification Rules (Excel)", type=["xlsx"])

def clean_or_split(text):
    """Always splits by OR logic (used for Exclude)."""
    if pd.isna(text) or not isinstance(text, str):
        return []
    text = text.replace(' and ', ',').replace(' or ', ',')
    return [t.strip().lower() for t in text.split(',') if t.strip()]

def parse_include(text):
    """Splits Include text respecting AND and OR logic."""
    if pd.isna(text) or not isinstance(text, str):
        return []

    text = text.lower()

    if ' and ' in text:
        and_parts = text.split(' and ')
        return [clean_or_split(part) for part in and_parts]
    else:
        return [clean_or_split(text)]

def preprocess_rules(rules_df):
    parsed_rules = []
    for _, row in rules_df.iterrows():
        include = parse_include(row['Include'])
        exclude = clean_or_split(row['Exclude'])
        label = row['Rule']
        parsed_rules.append((include, exclude, label))
    return parsed_rules

def matches_rule(title, include, exclude):
    for and_block in include:
        if not any(word in title for word in and_block):
            return False
    if any(word in title for word in exclude):
        return False
    return True

def classify_products(product_df, parsed_rules):
    results = []
    titles = product_df['TITLE'].str.lower().fillna('')
    progress = st.progress(0)

    for idx, title in enumerate(titles):
        matches = []
        for include, exclude, label in parsed_rules:
            if matches_rule(title, include, exclude):
                matches.append(label)
        results.append(', '.join(matches))
        if idx % 100 == 0 or idx == len(titles) - 1:
            progress.progress((idx + 1) / len(titles))

    product_df['mapped_classifications'] = results
    return product_df

def create_excel_download(df, filename="output.xlsx", sheet_name="Sheet1"):
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)
        return output.read()
    except Exception as e:
        st.error(f"Error creating Excel file: {e}")
        return None

if product_file and rules_file:
    try:
        product_df = pd.read_excel(product_file, sheet_name=0, engine='openpyxl')  # dynamic: first sheet
        rules_df = pd.read_excel(rules_file, sheet_name=0, engine='openpyxl')      # dynamic: first sheet

        if "TITLE" not in product_df.columns:
            st.error("Product file must contain a 'TITLE' column.")
            st.stop()
        if not all(col in rules_df.columns for col in ['Rule', 'Include', 'Exclude']):
            st.error("Rules file must contain 'Rule', 'Include', and 'Exclude' columns.")
            st.stop()

    except Exception as e:
        st.error(f"Error reading Excel files: {e}")
        st.stop()

    st.success("✅ Files uploaded successfully!")

    parsed_rules = preprocess_rules(rules_df)
    output_df = classify_products(product_df, parsed_rules)

    st.subheader("🔍 Preview of Classified Products")
    st.dataframe(output_df, use_container_width=True)

    csv = output_df.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download Classified CSV", csv, "classified_products.csv", "text/csv")

    excel_output = create_excel_download(output_df, "classified_products.xlsx", "Classified_Products")
    if excel_output:
        st.download_button(
            label="⬇️ Download Classified Excel",
            data=excel_output,
            file_name="classified_products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif product_file or rules_file:
    st.warning("⚠️ Please upload both the Product and Rules files to proceed.")

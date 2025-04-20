import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from functools import lru_cache

st.set_page_config(page_title="Product Classifier", layout="wide")
st.title("🧠 Product Classifier Tool for Flywheel")

st.markdown("Upload your product details and rules to auto-classify using AI-style logic.")

product_file = st.file_uploader("📦 Upload Product Details (Excel)", type=["xlsx"])
rules_file = st.file_uploader("📋 Upload Classification Rules (Excel)", type=["xlsx"])

def clean_and_split(text):
    """
    Cleans and splits a text string by commas, handling missing values and "and" conditions.
    Converts all words to lowercase and strips whitespace.
    NOTE: This function forces OR logic by replacing "and" with "," for both includes and excludes.
    """
    if pd.isna(text):
        return []
    if not isinstance(text, str):  # Handle non-string inputs
        return []
    text = text.replace(' and ', ',')
    return [t.strip().lower() for t in text.split(',') if t.strip()]

def parse_include(include_text):
    """Parses the 'Include' column, handling 'and' conditions for AND logic."""
    if pd.isna(include_text):
        return []
    include_text = include_text.lower()
    if ' and ' in include_text:
        return [clean_and_split(part) for part in include_text.split(' and ')]
    else:
        return [clean_and_split(include_text)]

def preprocess_rules(rules_df):
    """Preprocesses the rules DataFrame to extract and parse rule components."""
    parsed_rules = []
    for _, rule in rules_df.iterrows():
        include = parse_include(rule['Include'])
        exclude = clean_and_split(rule['Exclude'])
        label = rule['Rule']
        parsed_rules.append((include, exclude, label))
    return parsed_rules

@lru_cache(maxsize=None)
def matches_rule(title, include, exclude):
    """
    Checks if a title matches a rule's include and exclude criteria.
    NOTE: Exclude uses OR logic due to the clean_and_split function.
    """
    for and_block in include:
        if not any(word in title for word in and_block):
            return False
    if any(word in title for word in exclude):
        return False
    return True

def classify_products(product_df, parsed_rules):
    """
    Classifies products based on preprocessed rules.
    Returns the product DataFrame with an added 'mapped_classifications' column.
    """
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
    """
    Creates an in-memory Excel file from a DataFrame and returns its bytes for download.
    Handles potential errors during Excel writing.
    """
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)
        return output.read()  # Return the bytes
    except Exception as e:
        st.error(f"Error creating Excel file: {e}")
        return None

if product_file and rules_file:
    try:
        product_df = pd.read_excel(product_file, sheet_name='Product_details', engine='openpyxl')
        rules_df = pd.read_excel(rules_file, sheet_name='Rules', engine='openpyxl')

        # Check for required columns
        if "TITLE" not in product_df.columns:
            st.error("Product file must contain a 'TITLE' column.")
            st.stop()
        if "Rule" not in rules_df.columns or "Include" not in rules_df.columns or "Exclude" not in rules_df.columns:
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

    # Download as CSV
    csv = output_df.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download Classified CSV", csv, "classified_products.csv", "text/csv")

    # Download as Excel
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

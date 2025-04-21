# 🧠 Product Classifier Tool for Flywheel

This Streamlit app helps classify product titles using predefined rules and simple AI-style logic. It's built for non-coders who want to optimize catalog management workflows using rule-based automation.

## 🔧 Features

- Upload **Product Details** and **Classification Rules** in Excel format
- Supports complex matching logic using `AND`/`OR` conditions
- Applies inclusion and exclusion rules to classify products
- Progress bar for large datasets
- Preview results and download as CSV or Excel

## 📁 Input File Format

### Product File (Excel)
Must include a column titled `TITLE`.

### Rules File (Excel)
Must include the following columns:
- `Rule`: The label to assign
- `Include`: Text to look for in the title (supports `and`, `or`)
- `Exclude`: Terms that should not be present

#### Example Rule

| Rule         | Include              | Exclude       |
|--------------|----------------------|----------------|
| Electronics  | phone and charger    | refurbished    |
| Clothing     | shirt or jeans       | used           |

## 🚀 How to Run

You can run this app locally:

```bash
pip install streamlit pandas openpyxl
streamlit run your_script.py

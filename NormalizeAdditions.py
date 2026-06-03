import pandas as pd
import streamlit as st

st.set_page_config(layout="wide")
st.title("Cleaned Additions")
uploaded_file = st.file_uploader("Upload CSV")

if uploaded_file is None or not uploaded_file.name.endswith(".csv"):
    st.error("Please upload a CSV file")
    st.stop()

# convert csv into pandas DF
original_df = pd.read_csv(uploaded_file)
additions_df = original_df.copy()

# Yes / No normalization
responses_dict = {
    "yes": "Y",
    "Yes": "Y",
    "y": "Y",
    "No": "N",
    "no": "N",
    "n": "N",
}
additions_df["yesNoQuestionColumnName"] = additions_df["yesNoQuestionColumnName"].replace(responses_dict)

# Manufacturer Column
manufacturers_dict = {}

def normalizeManufacturerName(manufacturer):
    string_size = 12
    if pd.isna(manufacturer):
        return "UNKNOWN"
    if len(manufacturer) <= string_size:
        return manufacturer
    if manufacturer in manufacturers_dict:
        return manufacturers_dict[manufacturer]
    return manufacturer[0:string_size]
additions_df["Manufacturer"] = additions_df["Manufacturer"].apply(normalizeManufacturerName)

# NA check for Dates
default_value = 1900
additions_df["MFG"] = additions_df["MFG"].fillna(default_value)

# Side-by-side display
def zebra(df):
    return df.style.apply(
        lambda x: ["background-color: #f2f2f2" if i % 2 == 0 else "background-color: #ffffff" for i in range(len(x))],
        axis=0
    )

col_left, col_right = st.columns(2)
with col_left:
    st.caption("Original")
    st.dataframe(zebra(original_df), use_container_width=True)
with col_right:
    st.caption("Cleaned")
    st.dataframe(zebra(additions_df), use_container_width=True)

# download instead of local file write
st.download_button(
    label="Download Cleaned CSV",
    data=additions_df.to_csv(index=False).encode("utf-8"),
    file_name="cleaned_additions.csv",
    mime="text/csv",
)

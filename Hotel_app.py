import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz
import io

# ============================================================
# PAGE HEADER
# ============================================================
st.set_page_config(page_title="Hotel Comparison Tool", layout="wide")
st.title("🏨 Hotel Market Value Comparison Tool")
st.write("Upload your Excel file, process matching logic, and download results.")

# ============================================================
# CONFIG — TOLERANCE SETTINGS
# ============================================================
MV_TOLERANCE = 0.20   # 20%

# ============================================================
# SAFE VALUE FOR EXCEL
# ============================================================
def safe_excel_value(val):
    try:
        if pd.isna(val) or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
            return ""
        return val
    except:
        return ""

# ============================================================
# FUZZY MATCH FUNCTION
# ============================================================
def fuzzy_match(val, query, threshold=90):
    if pd.isna(val):
        return False
    return fuzz.partial_ratio(str(val).lower(), str(query).lower()) >= threshold

# ============================================================
# STATE TAX RATES
# ============================================================
state_tax_rates = {
    'Alabama': 0.0039, 'Arkansas': 0.0062, 'Arizona': 0.0066, 'California': 0.0076,
    'Colorado': 0.0051, 'Connecticut': 0.0214, 'Florida': 0.0089, 'Georgia': 0.0083,
    'Iowa': 0.0157, 'Idaho': 0.0069, 'Illinois': 0.0210, 'Indiana': 0.0085, 'Kansas': 0.0133,
    'Kentucky': 0.0080, 'Louisiana': 0.0000, 'Massachusetts': 0.0112, 'Maryland': 0.0109,
    'Michigan': 0.0154, 'Missouri': 0.0097, 'Mississippi': 0.0075, 'Montana': 0.0084,
    'North Carolina': 0.0077, 'Nebraska': 0.0173, 'New Jersey': 0.0249, 'New Mexico': 0.0080,
    'Nevada': 0.0060, 'Newyork': 0.0172, 'Ohio': 0.0157, 'Oklahoma': 0.0090, 'Oregon': 0.0097,
    'Pennsylvania': 0.0158, 'South Carolina': 0.0057, 'Tennessee': 0.0071, 'Texas': 0.0250,
    'Utah': 0.0057, 'Virginia': 0.0082, 'Washington': 0.0098
}

def get_state_tax_rate(state):
    return state_tax_rates.get(state, 0)

# ============================================================
# FILE UPLOAD  
# ============================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [c.strip() for c in df.columns]

    # Numeric conversions
    for col in ['No. of Rooms', 'Market Value-2024', '2024 VPR']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['No. of Rooms', 'Market Value-2024', '2024 VPR'])

    hotel_class_map = {
        "Budget (Low End)": 1, "Economy (Name Brand)": 2, "Midscale": 3,
        "Upper Midscale": 4, "Upscale": 5, "Upper Upscale First Class": 6,
        "Luxury Class": 7, "Independent Hotel": 8
    }

    df["Hotel Class Order"] = df["Hotel Class"].map(hotel_class_map)
    df = df.dropna(subset=["Hotel Class Order"])
    df["Hotel Class Order"] = df["Hotel Class Order"].astype(int)

    st.success("File uploaded successfully!")

    # ========================================================
    # PROCESS DATA
    # ========================================================
    if st.button("Process File"):
        st.info("Processing... please wait...")

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_out = []

            for i in range(len(df)):
                base = df.iloc[i]
                mv = base['Market Value-2024']
                vpr = base['2024 VPR']
                rooms = base["No. of Rooms"]

                subset = df[df.index != i]

                # Allowed hotel class logic
                allowed = {
                    1:[1,2,3],2:[1,2,3,4],3:[2,3,4,5],4:[3,4,5,6],
                    5:[4,5,6,7],6:[5,6,7,8],7:[6,7,8],8:[7,8]
                }.get(base["Hotel Class Order"], [])

                # Tolerance range
                mv_min = mv * (1 - MV_TOLERANCE)
                mv_max = mv * (1 + MV_TOLERANCE)

                mask = (
                    (subset['State'] == base['State']) &
                    (subset['Property County'] == base['Property County']) &
                    (subset['No. of Rooms'] < rooms) &
                    (subset['Market Value-2024'].between(mv_min, mv_max)) &
                    (subset['2024 VPR'] < vpr) &
                    (subset['Hotel Class Order'].isin(allowed))
                )

                matches = subset[mask]

                if not matches.empty:
                    selected = matches.sort_values("2024 VPR").head(5)
                else:
                    selected = pd.DataFrame()

                df_out.append((base, selected))

            # Flatten results
            result_rows = []
            for base, selected in df_out:
                row_data = base.to_dict()
                for j in range(5):
                    if j < len(selected):
                        for col in selected.columns:
                            row_data[f"Result{j+1}_{col}"] = safe_excel_value(selected.iloc[j][col])
                    else:
                        for col in df.columns:
                            row_data[f"Result{j+1}_{col}"] = ""
                result_rows.append(row_data)

            final_df = pd.DataFrame(result_rows)
            final_df.to_excel(writer, sheet_name="Comparison Results", index=False)

        st.success("Processing completed!")

        # ========================================================
        # FILE DOWNLOAD BUTTON
        # ========================================================
        st.download_button(
            label="📥 Download Processed Excel",
            data=output.getvalue(),
            file_name="comparison_results_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

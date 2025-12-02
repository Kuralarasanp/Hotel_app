import streamlit as st
import pandas as pd
import numpy as np
import io
from rapidfuzz import fuzz

# ============================================================
# CONFIG
# ============================================================
MV_TOLERANCE = 0.20   # 20% Market Value tolerance

st.set_page_config(page_title="Hotel Comparison Engine", layout="wide")
st.title("🏨 Hotel Market Value Comparison Tool")
st.write("Upload your Excel file, process matching logic, and download results.")

# ============================================================
# HOTEL CLASS MAPPING
# ============================================================
hotel_class_map = {
    "Budget (Low End)": 1,
    "Economy (Name Brand)": 2,
    "Midscale": 3,
    "Upper Midscale": 4,
    "Upscale": 5,
    "Upper Upscale First Class": 6,
    "Luxury Class": 7,
    "Independent Hotel": 8
}

def map_hotel_class(value):
    return hotel_class_map.get(str(value).strip(), "")

# ============================================================
# STATE TAX RATES
# ============================================================
state_tax_rates = {
    'Alabama': 0.0039, 'Arkansas': 0.0062, 'Arizona': 0.0066, 'California': 0.0076, 'Colorado': 0.0051,
    'Connecticut': 0.0214, 'Florida': 0.0089, 'Georgia': 0.0083, 'Iowa': 0.0157, 'Idaho': 0.0069,
    'Illinois': 0.0210, 'Indiana': 0.0085, 'Kansas': 0.0133, 'Kentucky': 0.0080, 'Louisiana': 0.0000,
    'Massachusetts': 0.0112, 'Maryland': 0.0109, 'Michigan': 0.0154, 'Missouri': 0.0097, 'Mississippi': 0.0075,
    'Montana': 0.0084, 'North Carolina': 0.0077, 'Nebraska': 0.0173, 'New Jersey': 0.0249, 'New Mexico': 0.0080,
    'Nevada': 0.0060, 'Newyork': 0.0172, 'Ohio': 0.0157, 'Oklahoma': 0.0090, 'Oregon': 0.0097,
    'Pennsylvania': 0.0158, 'South Carolina': 0.0057, 'Tennessee': 0.0071, 'Texas': 0.0250, 'Utah': 0.0057,
    'Virginia': 0.0082, 'Washington': 0.0098
}

def get_state_tax_rate(state):
    return state_tax_rates.get(state, 0)

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
# FILE UPLOAD
# ============================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [col.strip() for col in df.columns]

    # Convert numeric columns
    for col in ['No. of Rooms', 'Market Value-2024', '2024 VPR']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['No. of Rooms', 'Market Value-2024', '2024 VPR'])

    # Map hotel class
    df["Hotel Class Mapped"] = df["Hotel Class"].apply(map_hotel_class)
    df = df.dropna(subset=["Hotel Class Mapped"])
    df["Hotel Class Mapped"] = df["Hotel Class Mapped"].astype(int)

    st.success("File uploaded successfully!")

    if st.button("Process File"):
        st.info("Processing... please wait...")

        match_columns = [
            'Property Address', 'State', 'Property County', 'Project / Hotel Name',
            'Owner Name/ LLC Name', 'No. of Rooms', 'Market Value-2024',
            '2024 VPR', 'Hotel Class'
        ]

        df_out_rows = []

        for i in range(len(df)):
            base = df.iloc[i]
            mv = base['Market Value-2024']
            vpr = base['2024 VPR']
            rooms = base["No. of Rooms"]
            hotel_class_order = base["Hotel Class Mapped"]

            subset = df[df.index != i]

            allowed = {
                1:[1,2,3],2:[1,2,3,4],3:[2,3,4,5],4:[3,4,5,6],
                5:[4,5,6,7],6:[5,6,7,8],7:[6,7,8],8:[7,8]
            }.get(hotel_class_order, [])

            mv_min = mv * (1 - MV_TOLERANCE)
            mv_max = mv * (1 + MV_TOLERANCE)

            mask = (
                (subset['State'] == base['State']) &
                (subset['Property County'] == base['Property County']) &
                (subset['No. of Rooms'] < rooms) &
                (subset['Market Value-2024'].between(mv_min, mv_max)) &
                (subset['2024 VPR'] < vpr) &
                (subset['Hotel Class Mapped'].isin(allowed))
            )

            matches = subset[mask]

            # Select top 5 matches
            if not matches.empty:
                nearest = matches.sort_values("2024 VPR").head(3)
                remaining = matches.drop(nearest.index)
                least = remaining.sort_values("2024 VPR").head(1)
                remaining = remaining.drop(least.index)
                top = remaining.sort_values("2024 VPR", ascending=False).head(1)

                selected = pd.concat([nearest, least, top]).head(5)
            else:
                selected = pd.DataFrame()

            # Build output row
            row_data = {}
            for col in match_columns:
                row_data[col] = safe_excel_value(base[col])

            # Matching status
            if not matches.empty:
                row_data["Matching Results Count / Status"] = f"Total: {len(matches)} | Selected: {len(selected)}"
            else:
                row_data["Matching Results Count / Status"] = "No_Match_Case"

            # OverPaid calculation
            if not matches.empty:
                median_vpr = selected["2024 VPR"].head(3).median()
                state_rate = get_state_tax_rate(base["State"])
                assessed = median_vpr * rooms * state_rate
                subject_tax = mv * state_rate
                overpaid = subject_tax - assessed
            else:
                overpaid = ""

            row_data["OverPaid"] = safe_excel_value(overpaid)

            # Add top 5 results with Hotel Class
            for idx in range(5):
                if idx < len(selected):
                    result_row = selected.iloc[idx]
                    for col in match_columns:
                        if col == "Hotel Class":
                            row_data[f"Result{idx+1}_{col}"] = safe_excel_value(result_row[col])
                        else:
                            row_data[f"Result{idx+1}_{col}"] = safe_excel_value(result_row[col])
                else:
                    for col in match_columns:
                        row_data[f"Result{idx+1}_{col}"] = ""

            df_out_rows.append(row_data)

        # Convert to dataframe
        final_df = pd.DataFrame(df_out_rows)

        # Save to in-memory Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            final_df.to_excel(writer, sheet_name="Comparison Results", index=False)

        st.success("Processing completed!")

        st.download_button(
            label="📥 Download Final Comparison Excel",
            data=output.getvalue(),
            file_name="comparison_results_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

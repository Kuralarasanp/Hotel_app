import streamlit as st
import pandas as pd
import numpy as np
import os
from fuzzywuzzy import fuzz
from io import BytesIO

# ============================================================
# STREAMLIT APP TITLE
# ============================================================
st.title("Hotel Market Value & VPR Comparison Tool")

# ============================================================
# FILE UPLOAD
# ============================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)
    df.columns = [col.strip() for col in df.columns]

    # ============================================================
    # CONFIG — TOLERANCE SETTINGS
    # ============================================================
    MV_TOLERANCE = 0.20

    # ============================================================
    # SAFE VALUE
    # ============================================================
    def safe_excel_value(val):
        try:
            if pd.isna(val) or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
                return ""
            return val
        except:
            return ""

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
    # PROCESS DATA
    # ============================================================
    for col in ['No. of Rooms', 'Market Value-2024', '2024 VPR']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=['No. of Rooms', 'Market Value-2024', '2024 VPR'])

    hotel_class_map = {
        "Budget (Low End)": 1, "Economy (Name Brand)": 2, "Midscale": 3, "Upper Midscale": 4,
        "Upscale": 5, "Upper Upscale First Class": 6, "Luxury Class": 7, "Independent Hotel": 8
    }

    df["Hotel Class Order"] = df["Hotel Class"].map(hotel_class_map)
    df = df.dropna(subset=["Hotel Class Order"])
    df["Hotel Class Order"] = df["Hotel Class Order"].astype(int)

    # ============================================================
    # HELPER FUNCTIONS
    # ============================================================
    def get_nearest_three(df, mv, vpr):
        df = df.copy()
        df["dist"] = ((df["Market Value-2024"] - mv)**2 + (df["2024 VPR"] - vpr)**2)**0.5
        return df.sort_values("dist").head(3).drop(columns="dist")

    def get_least_one(df):
        return df.sort_values(["Market Value-2024","2024 VPR"], ascending=[True,True]).head(1)

    def get_top_one(df):
        return df.sort_values(["Market Value-2024","2024 VPR"], ascending=[False,False]).head(1)

    # ============================================================
    # MAIN PROCESSING
    # ============================================================
    match_columns = [
        'Property Address', 'State', 'Property County', 'Project / Hotel Name',
        'Owner Name/ LLC Name', 'No. of Rooms', 'Market Value-2024',
        '2024 VPR', 'Hotel Class', 'Hotel Class Number'
    ]
    all_columns = list(df.columns)
    max_results_per_row = 5

    results = []

    for i in range(len(df)):
        base = df.iloc[i]
        mv = base['Market Value-2024']
        vpr = base['2024 VPR']
        rooms = base["No. of Rooms"]

        subset = df[df.index != i]

        allowed = {
            1:[1,2,3],2:[1,2,3,4],3:[2,3,4,5],4:[3,4,5,6],
            5:[4,5,6,7],6:[5,6,7,8],7:[6,7,8],8:[7,8]
        }.get(base["Hotel Class Order"], [])

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

        matches = subset[mask].drop_duplicates(
            subset=['Project / Hotel Name','Property Address','Owner Name/ LLC Name']
        )

        row_data = {col: safe_excel_value(base[col]) for col in match_columns[:-1]}
        row_data["Hotel Class Number"] = base["Hotel Class Order"]

        if not matches.empty:
            nearest = get_nearest_three(matches, mv, vpr)
            rem = matches.drop(nearest.index)
            least = get_least_one(rem)
            rem = rem.drop(least.index)
            top = get_top_one(rem)
            selected = pd.concat([nearest, least, top]).head(5).reset_index(drop=True)

            row_data["Matching Results Count / Status"] = f"Total: {len(matches)} | Selected: {len(selected)}"

            median_vpr = selected["2024 VPR"].head(3).median()
            state_rate = get_state_tax_rate(base["State"])
            assessed = median_vpr * rooms * state_rate
            subject_tax = mv * state_rate
            row_data["OverPaid"] = subject_tax - assessed
        else:
            row_data["Matching Results Count / Status"] = "No_Match_Case"
            row_data["OverPaid"] = ""

        results.append(row_data)

    output_df = pd.DataFrame(results)

    # ============================================================
    # DOWNLOAD BUTTON
    # ============================================================
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name="Comparison Results")
    output_buffer.seek(0)

    st.download_button(
        label="Download Processed Excel",
        data=output_buffer,
        file_name="comparison_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import streamlit as st
import pandas as pd
from collections import Counter
import io

st.set_page_config(page_title="Logistics Quote Generator", layout="wide")

st.title("📦 Logistics Quote Pipeline")

# --- SIDEBAR: MANUAL INPUTS ---
with st.sidebar:
    st.header("Shipment Details")
    destination = st.text_input("Destination", value="UK - Radial FAO Monat, Middleton Oldham OL9 9XA")
    service = st.selectbox("Service", ["LCL", "LTL", "FCL", "Air", "Courier"])
    commodity = st.text_input("Commodity", value="Finished goods / Haircare / Skincare")
    cargo_value = st.text_input("Value of Cargo", value="USD$ 33,650.35")
    incoterms = st.selectbox("Incoterms", ["-", "EXW", "FOB", "DDP", "DAP", "CIF"])

# --- MAIN: FILE UPLOAD ---
packing_file = st.file_uploader("Upload Outbound Packing List (.xlsx)", type=['xlsx'])

if packing_file:
    # Read the whole sheet as strings first to find keywords easily
    df_raw = pd.read_excel(packing_file, header=None).astype(str)
    
    # Helper to find a value based on a keyword
    def get_val(keyword, row_off=0, col_off=0):
        for r in range(len(df_raw)):
            for c in range(len(df_raw.columns)):
                if keyword.lower() in df_raw.iloc[r, c].lower():
                    # Return the cell at the offset
                    res = df_raw.iloc[r + row_off, c + col_off]
                    return res
        return "0"

    # 1. GRAB TOTALS FROM FOOTER (Looking above the labels as per your image)
    pallet_count = get_val("Pallets", row_off=-1)
    total_units = get_val("Units", row_off=-1)
    total_weight_lbs = get_val("Gross Weight", row_off=-1)

    # Clean the numbers (remove commas or symbols)
    def clean_num(val):
        clean = "".join(c for c in str(val) if c.isdigit() or c == '.')
        return float(clean) if clean else 0.0

    pallets_final = int(clean_num(pallet_count))
    units_final = int(clean_num(total_units))
    lbs_final = clean_num(total_weight_lbs)
    kgs_final = lbs_final * 0.453592

    # 2. DYNAMIC DIMENSION FINDER
    # Instead of 'Dim / Pallet', we find whichever column contains "Dim" and "Pallet"
    dim_list = []
    dim_col_idx = -1
    
    # Find which column index has the dimensions
    for c in range(len(df_raw.columns)):
        column_data = df_raw.iloc[:, c]
        if any("dim" in str(val).lower() and "pallet" in str(val).lower() for val in column_data):
            dim_col_idx = c
            break
    
    if dim_col_idx != -1:
        # Get all values in that column, skip headers, and filter out noise
        potential_dims = df_raw.iloc[3:, dim_col_idx].tolist()
        dim_list = [d.strip() for d in potential_dims if "x" in d.lower() and len(d) > 3]

    dim_counts = Counter(dim_list)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    st.success(f"✅ Data Found: {pallets_final} Pallets | {units_final:,} Units | {lbs_final:,.2f} LBS")

    # --- GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 3. CONSTRUCT THE OUTPUT
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", f"{units_final:,}"],
            ["PALLETS", pallets_final],
            ["DIMENSIONS", formatted_dims[0] if formatted_dims else ""],
        ]
        
        if len(formatted_dims) > 1:
            for extra_dim in formatted_dims[1:]:
                quote_data.append(["", extra_dim])

        quote_data.extend([
            ["", ""],
            ["TOTAL WEIGHT", f"{lbs_final:,.2f} LBS | {kgs_final:,.2f} KGS"],
            ["COMMODITY", commodity],
            ["INCOTERMS", incoterms],
            ["VALUE OF CARGO", cargo_value]
        ])
        
        df_output = pd.DataFrame(quote_data)
        
        # Excel Save
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_output.to_excel(writer, index=False, header=False)

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Download Excel", data=buf.getvalue(), file_name="Quote_Request.xlsx")
            st.table(df_output)
        with c2:
            email = f"Hi,\n\nQuote for {destination}:\n- {pallets_final} Pallets\n- {lbs_final} LBS\n\nThanks!"
            st.text_area("Email:", value=email, height=300)
else:
    st.info("Upload your Packing List to begin.")

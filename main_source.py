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
    # Read the whole sheet as strings to find keywords
    df_raw = pd.read_excel(packing_file, header=None).astype(str)
    
    # Helper to find a value based on a keyword
    def get_val(keyword, row_off=0, col_off=0):
        for r in range(len(df_raw)):
            for c in range(len(df_raw.columns)):
                if keyword.lower() in df_raw.iloc[r, c].lower():
                    try:
                        return df_raw.iloc[r + row_off, c + col_off]
                    except:
                        return "0"
        return "0"

    # NEW: Robust number cleaner that handles commas like "9,132"
    def clean_num(val):
        if pd.isna(val) or str(val).lower() == 'nan':
            return 0.0
        # Remove commas and any other non-numeric chars except the decimal point
        clean = "".join(c for c in str(val) if c.isdigit() or c == '.')
        try:
            return float(clean)
        except:
            return 0.0

    # 1. GRAB TOTALS FROM FOOTER (Looking above the labels)
    pallet_raw = get_val("Pallets", row_off=-1)
    units_raw = get_val("Units", row_off=-1)
    weight_raw = get_val("Gross Weight", row_off=-1)

    pallets_final = int(clean_num(pallet_raw))
    units_final = int(clean_num(units_raw))
    lbs_final = clean_num(weight_raw)
    kgs_final = lbs_final * 0.453592

    # 2. DYNAMIC DIMENSION FINDER
    dim_list = []
    dim_col_idx = -1
    
    # Find column that contains "Dim" and "Pallet" (usually column N or O)
    for c in range(len(df_raw.columns)):
        if any("dim" in str(val).lower() and "pallet" in str(val).lower() for val in df_raw.iloc[:, c]):
            dim_col_idx = c
            break
    
    if dim_col_idx != -1:
        # Start looking after the header rows
        potential_dims = df_raw.iloc[3:, dim_col_idx].tolist()
        # Only keep strings that look like "47 X 31 X 52"
        dim_list = [d.strip() for d in potential_dims if "x" in d.lower() and len(d) > 5]

    dim_counts = Counter(dim_list)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    st.success(f"✅ Data Extracted: {pallets_final} Pallets | {units_final:,} Units | {lbs_final:,.2f} LBS")

    # --- GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 3. CONSTRUCT THE OUTPUT TABLE
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", f"{units_final:,}"],
            ["PALLETS", pallets_final],
        ]
        
        # Dimensions
        if formatted_dims:
            quote_data.append(["DIMENSIONS", formatted_dims[0]])
            if len(formatted_dims) > 1:
                for extra_dim in formatted_dims[1:]:
                    quote_data.append(["", extra_dim])
        else:
            quote_data.append(["DIMENSIONS", "Pending"])

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
            st.subheader("1. Final Form")
            st.download_button("📥 Download Excel", data=buf.getvalue(), file_name="Quote_Request.xlsx")
            st.table(df_output)
        with c2:
            st.subheader("2. Email Draft")
            email = f"Hi Team,\n\nPlease quote {service} to {destination}:\n- {pallets_final} Pallets\n- {lbs_final:,.2f} LBS\n- {units_final:,} Units\n\nThanks!"
            st.text_area("Copy Email:", value=email, height=300)
else:
    st.info("Please upload the Outbound Packing List to proceed.")

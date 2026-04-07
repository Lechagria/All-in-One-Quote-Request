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
    # Read raw data without headers to search freely
    df_raw = pd.read_excel(packing_file, header=None)
    
    # helper function to find numbers near keywords
    def find_value_near(keyword, offset_row=-1, offset_col=0):
        try:
            # Find the row and column index of the keyword
            mask = df_raw.astype(str).apply(lambda x: x.str.contains(keyword, case=False, na=False))
            pos = mask.values.nonzero()
            if len(pos[0]) > 0:
                row, col = pos[0][0], pos[1][0]
                val = df_raw.iloc[row + offset_row, col + offset_col]
                return pd.to_numeric(val, errors='coerce')
            return 0
        except:
            return 0

    # 1. GRAB TOTALS FROM THE SUMMARY FOOTER (As requested)
    pallet_count = int(find_value_near("Pallets", offset_row=-1)) 
    total_units = int(find_value_near("Units", offset_row=-1))
    total_weight_lbs = find_value_near("Gross Weight", offset_row=-1)
    
    # Calculate KG based on the LBS found in the footer
    total_weight_kgs = total_weight_lbs * 0.453592

    # 2. DIMENSIONS (Still needs to look at the main table)
    df_table = pd.read_excel(packing_file, header=2)
    raw_dims = df_table['Dim / Pallet'].dropna().astype(str).tolist()
    raw_dims = [d for d in raw_dims if "Dim" not in d and d.strip() != ""]
    dim_counts = Counter(raw_dims)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    st.success(f"✅ Found Footer Totals: {pallet_count} Pallets | {total_units:,} Units | {total_weight_lbs:,.2f} LBS")

    # --- THE GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 3. CREATE THE DATA TABLE FOR EXCEL
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", f"{total_units:,}"],
            ["PALLETS", pallet_count],
            ["DIMENSIONS", formatted_dims[0] if formatted_dims else ""],
        ]
        
        if len(formatted_dims) > 1:
            for extra_dim in formatted_dims[1:]:
                quote_data.append(["", extra_dim])

        quote_data.extend([
            ["", ""],
            ["TOTAL WEIGHT", f"{total_weight_lbs:,.2f} LBS | {total_weight_kgs:,.2f} KGS"],
            ["COMMODITY", commodity],
            ["INCOTERMS", incoterms],
            ["VALUE OF CARGO", cargo_value]
        ])
        
        df_quote = pd.DataFrame(quote_data)

        # 4. EXCEL GENERATION
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_quote.to_excel(writer, index=False, header=False, sheet_name='Quote Request')
        
        # 5. EMAIL DRAFT
        email_body = f"Hi Team,\n\nQuote Request for {destination}:\n- {pallet_count} Pallets\n- {total_weight_lbs:,.2f} LBS\n- {total_units:,} Units\n\nAttached: Packing List."

        st.divider()
        col_dl, col_em = st.columns(2)
        
        with col_dl:
            st.download_button("📥 Download Quote Request.xlsx", data=buffer.getvalue(), file_name=f"Quote_{pallet_count}PLTS.xlsx")
            st.table(df_quote)

        with col_em:
            st.text_area("Email Draft:", value=email_body, height=300)

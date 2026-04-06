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
    # Read the file skipping the header branding
    df = pd.read_excel(packing_file, header=2)
    
    # 1. CLEAN THE PALLET COLUMN (The Fix for your crash)
    # This turns "Pallets" into a blank (NaN) so the math works
    df['PALLET QTY'] = pd.to_numeric(df['PALLET QTY'], errors='coerce')
    df['PALLET QTY'] = df['PALLET QTY'].ffill()
    
    # 2. CLEAN THE UNITS AND WEIGHT COLUMNS
    df['Total Units'] = pd.to_numeric(df['Total Units'], errors='coerce')
    df['Weight / Pallet'] = pd.to_numeric(df['Weight / Pallet'], errors='coerce')
    
    # 3. FILTER TO ITEM ROWS ONLY (Ignores the footer rows with "Pallets" text)
    df_items = df.dropna(subset=['P.O.'])
    
    # 4. CALCULATIONS
    total_units = int(df_items['Total Units'].sum())
    total_weight_lbs = df_items['Weight / Pallet'].sum()
    total_weight_kgs = total_weight_lbs * 0.453592 
    
    # Dimensions Grouping
    raw_dims = df['Dim / Pallet'].dropna().astype(str).tolist()
    raw_dims = [d for d in raw_dims if "Dim" not in d and d.strip() != ""]
    dim_counts = Counter(raw_dims)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]
    
    # Safe Max Pallet Count
    pallet_count = int(df['PALLET QTY'].max()) if not df['PALLET QTY'].dropna().empty else 0

    st.success(f"✅ Data extracted: {pallet_count} Pallets and {total_units:,} Units found.")

    if st.button("🚀 Generate Template"):
        # Create Data for Excel
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", total_units],
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
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_quote.to_excel(writer, index=False, header=False, sheet_name='Quote Request')
        
        email_body = f"Quote Request for {destination}\nPallets: {pallet_count}\nWeight: {total_weight_lbs:,.2f} LBS"

        st.divider()
        col_dl, col_em = st.columns(2)
        with col_dl:
            st.download_button("📥 Download Excel", data=buffer.getvalue(), file_name="Quote_Request.xlsx")
            st.table(df_quote) 
        with col_em:
            st.text_area("Email Draft:", value=email_body, height=300)

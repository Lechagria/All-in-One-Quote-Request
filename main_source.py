import streamlit as st
import pandas as pd
from collections import Counter
import io

# Set page configuration
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
    # 1. READ AND PROCESS DATA
    # Skipping first 2 rows to reach headers: [PALLET QTY, P.O., SKU...]
    df = pd.read_excel(packing_file, header=2)
    
    # FIX 1: Fill down the Pallet QTY so every row belongs to a pallet number
    df['PALLET QTY'] = df['PALLET QTY'].ffill()
    
    # NEW FIX: Convert numeric columns and ignore text errors (like "Units" or "LBS" in rows)
    df['Total Units'] = pd.to_numeric(df['Total Units'], errors='coerce')
    df['Weight / Pallet'] = pd.to_numeric(df['Weight / Pallet'], errors='coerce')
    
    # Filter to get only rows that have an actual Purchase Order (ignores summary rows at the bottom)
    df_items = df.dropna(subset=['P.O.'])
    
    # FIX 2: Perform Calculations
    total_units = int(df_items['Total Units'].sum())
    # Sum only the 'Weight / Pallet' column values
    total_weight_lbs = df['Weight / Pallet'].sum()
    total_weight_kgs = total_weight_lbs * 0.453592 
    
    # FIX 3: Group Dimensions (e.g., "47 X 31 X 52 (x2)")
    # Extract unique dimensions for each pallet
    raw_dims = df['Dim / Pallet'].dropna().astype(str).tolist()
    # Filter out any header-like text that might have been picked up
    raw_dims = [d for d in raw_dims if "Dim" not in d]
    
    dim_counts = Counter(raw_dims)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]
    
    # Correct pallet count based on the max value found in the filled column
    pallet_count = int(df['PALLET QTY'].max())

    st.success(f"✅ Data extracted: {pallet_count} Pallets and {total_units:,} Units found.")

    # --- THE GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 2. CREATE THE EXCEL QUOTE STRUCTURE
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", total_units],
            ["PALLETS", pallet_count],
            ["DIMENSIONS", formatted_dims[0] if formatted_dims else ""],
        ]
        
        # Add extra dimension rows if they exist
        if len(formatted_dims) > 1:
            for extra_dim in formatted_dims[1:]:
                quote_data.append(["", extra_dim])

        # Add the remaining footer data
        quote_data.extend([
            ["", ""],
            ["TOTAL WEIGHT", f"{total_weight_lbs:,.2f} LBS | {total_weight_kgs:,.2f} KGS"],
            ["COMMODITY", commodity],
            ["INCOTERMS", incoterms],
            ["VALUE OF CARGO", cargo_value]
        ])
        
        df_quote = pd.DataFrame(quote_data)

        # Buffer for the Excel file
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_quote.to_excel(writer, index=False, header=False, sheet_name='Quote Request')
        
        # 3. GENERATE EMAIL CONTENT
        email_body = f"""
Hi Team,

Please provide a quote for the following outbound shipment:
- Destination: {destination}
- Service: {service}
- Total Units: {total_units:,}
- Total Pallets: {pallet_count}
- Weight: {total_weight_lbs:,.2f} LBS ({total_weight_kgs:,.2f} KGS)
- Commodity: {commodity}
- Value: {cargo_value}

Please find the formal Quote Request attached. Thanks!
        """

        # --- DISPLAY RESULTS ---
        st.divider()
        col_dl, col_em = st.columns(2)
        
        with col_dl:
            st.subheader("1. Download Document")
            st.download_button(
                label="📥 Download Quote Request.xlsx",
                data=buffer.getvalue(),
                file_name=f"Quote_Request_{pallet_count}PLTS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.write("**Preview:**")
            st.table(df_quote) 

        with col_em:
            st.subheader("2. Copy Email")
            st.text_area("Copy into your email draft:", value=email_body, height=350)

else:
    st.info("Please upload the Outbound Packing List to begin.")

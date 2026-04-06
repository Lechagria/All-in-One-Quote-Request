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
    # We skip the first 2 rows to reach the headers [cite: 1]
    df = pd.read_excel(packing_file, header=2)
    
    # FIX 1: "Fill Down" the Pallet QTY so every row has its pallet number 
    df['PALLET QTY'] = df['PALLET QTY'].ffill()
    
    # Filter to get only item rows (rows with an SKU or Units) [cite: 1]
    df_items = df.dropna(subset=['Total Units'])
    
    # FIX 2: Correct Total Units and Weight
    total_units = int(df_items['Total Units'].sum()) [cite: 4]
    # Weight / Pallet only appears on specific rows, so we sum that specifically [cite: 1]
    total_weight_lbs = df['Weight / Pallet'].sum() [cite: 4]
    total_weight_kgs = total_weight_lbs * 0.453592 
    
    # FIX 3: Group Dimensions (e.g., "47 x 31 x 52 (x2)")
    # We grab the dimension listed for each pallet [cite: 1]
    raw_dims = df['Dim / Pallet'].dropna().astype(str).tolist()
    dim_counts = Counter(raw_dims)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]
    
    pallet_count = int(df['PALLET QTY'].max()) [cite: 4]

    st.success(f"✅ Data extracted: {pallet_count} Pallets and {total_units} Units found.")

    if st.button("🚀 Generate Template"):
        # ... (Rest of the generation and email code) ...
        
        # In the quote_data section, use the new formatted_dims:
        ["DIMENSIONS", " | ".join(formatted_dims)],
        
        # 2. CREATE THE EXCEL QUOTE (MATCHING YOUR TARGET )
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", total_units],
            ["PALLETS", pallet_count],
            ["DIMENSIONS", " | ".join(unique_dims)],
            ["TOTAL WEIGHT", f"{total_weight_lbs:,.2f} LBS | {total_weight_kgs:,.2f} KGS"],
            ["COMMODITY", commodity],
            ["INCOTERMS", incoterms],
            ["VALUE OF CARGO", cargo_value]
        ]
        
        df_quote = pd.DataFrame(quote_data)

        # Buffer for the Excel file
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_quote.to_excel(writer, index=False, header=False, sheet_name='Quote Request')
        
        # 3. GENERATE EMAIL CONTENT
        email_body = f"""
Hi Team,

Please provide a quote for the following:
- Destination: {destination}
- Service: {service}
- Pallets: {pallet_count}
- Weight: {total_weight_lbs:,.2f} LBS ({total_weight_kgs:,.2f} KGS)
- Commodity: {commodity}

Form attached. Thanks!
        """

        # --- DISPLAY RESULTS ---
        st.divider()
        col_dl, col_em = st.columns(2)
        
        with col_dl:
            st.subheader("1. Download Document")
            st.download_button(
                label="📥 Download Quote Request.xlsx",
                data=buffer.getvalue(),
                file_name="Generated_Quote_Request.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.table(df_quote) # Show preview of the generated data

        with col_em:
            st.subheader("2. Copy Email")
            st.text_area("Copy into your email draft:", value=email_body, height=300)

else:
    st.info("Waiting for Packing List upload...")

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
    # 1. READ DATA (Skip first 2 rows for header)
    df = pd.read_excel(packing_file, header=2)
    
    # 2. CLEAN COLUMNS (Force everything to be a number, turning text into "Empty")
    # This prevents the crash when it hits words like "Pallets" or "Units"
    df['PALLET QTY'] = pd.to_numeric(df['PALLET QTY'], errors='coerce')
    df['Total Units'] = pd.to_numeric(df['Total Units'], errors='coerce')
    df['Weight / Pallet'] = pd.to_numeric(df['Weight / Pallet'], errors='coerce')
    
    # 3. FILL PALLET NUMBERS (Forward fill the empty gaps)
    df['PALLET QTY'] = df['PALLET QTY'].ffill()
    
    # 4. CALCULATIONS (Ignore the "Empty" text rows)
    pallet_count = int(df['PALLET QTY'].max()) if not df['PALLET QTY'].dropna().empty else 0
    total_units = int(df['Total Units'].sum())
    total_weight_lbs = df['Weight / Pallet'].sum()
    total_weight_kgs = total_weight_lbs * 0.453592

    # 5. DIMENSIONS GROUPING
    raw_dims = df['Dim / Pallet'].dropna().astype(str).tolist()
    # Filter out header titles if they got picked up
    raw_dims = [d for d in raw_dims if "Dim" not in d and d.strip() != ""]
    dim_counts = Counter(raw_dims)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    st.success(f"✅ Extracted: {pallet_count} Pallets | {total_units:,} Units | {total_weight_lbs:,.2f} LBS")

    # --- THE GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 6. CREATE THE DATA FOR THE TABLE (Exactly as your screenshot)
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", f"{total_units:,}"],
            ["PALLETS", pallet_count],
            ["DIMENSIONS", formatted_dims[0] if formatted_dims else ""],
        ]
        
        # Add additional dimension rows if there are multiple sizes
        if len(formatted_dims) > 1:
            for extra_dim in formatted_dims[1:]:
                quote_data.append(["", extra_dim])

        # Add the footer info
        quote_data.extend([
            ["", ""],
            ["TOTAL WEIGHT", f"{total_weight_lbs:,.2f} LBS | {total_weight_kgs:,.2f} KGS"],
            ["COMMODITY", commodity],
            ["INCOTERMS", incoterms],
            ["VALUE OF CARGO", cargo_value]
        ])
        
        df_quote = pd.DataFrame(quote_data)

        # 7. EXCEL GENERATION
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_quote.to_excel(writer, index=False, header=False, sheet_name='Quote Request')
        
        # 8. EMAIL DRAFT
        email_body = f"""
Hi Team,

Please provide a quote for the following shipment:
- POs: Found in Packing List
- Pallets: {pallet_count}
- Dimensions: {', '.join(formatted_dims)}
- Weight: {total_weight_lbs:,.2f} LBS ({total_weight_kgs:,.2f} KGS)
- Destination: {destination}

Quote Request and Packing List attached.
        """

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
            st.table(df_quote) # This will now show the clean LBS | KG line

        with col_em:
            st.subheader("2. Copy Email")
            st.text_area("Email Draft:", value=email_body, height=350)

else:
    st.info("Upload the Outbound Packing List to get started.")
else:
    st.info("Upload the Outbound Packing List to get started.")

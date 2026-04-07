import streamlit as st
import pandas as pd
from collections import Counter
import io
import re

st.set_page_config(page_title="Logistics Quote Generator", layout="wide")

st.title("📦 Logistics Quote Pipeline")

# --- CUSTOM LISTS ---
destinations = [
    "UK - Radial FAO Monat, 26, 26 Broadgate, Chadderton, Middleton Oldham OL9 9XA",
    "POLAND - Radial Poland Sp. z o.o. Moszna Parcela 29, Budynek C3 05-840 Brwinów",
    "AUSTRALIA - FDM WAREHOUSING C/O Landmark Global 7 Eucalyptus Place",
    "MONAT Global Canada — 135 SPARKS AVE NORTH YORK ON M2H 2S5 Canada",
    "FENIX FWD INC. - 417 LOGISTIC LAREDO, TEXAS 78045",
    "OTHER (Type Manually below)"
]

services = [
    "40\" REEFER",
    "40\" DRY",
    "20\" DRY",
    "HAZMAT LCL",
    "LCL Ocean",
    "LTL Road",
    "Air Freight",
    "Courier"
]

# --- SIDEBAR: MANUAL INPUTS ---
with st.sidebar:
    st.header("Shipment Details")
    
    selected_dest = st.selectbox("Select Destination", destinations)
    destination = st.text_input("Manual Destination Entry", value=selected_dest) if selected_dest == "OTHER (Type Manually below)" else selected_dest

    service = st.selectbox("Service", services)
    commodity = st.text_input("Commodity", value="Finished goods / Haircare / Skincare")
    cargo_value = st.text_input("Value of Cargo", value="USD$ ")
    incoterms = st.selectbox("Incoterms", ["-", "EXW", "FOB", "DDP", "DAP", "CIF"])

# --- MAIN: FILE UPLOAD ---
packing_file = st.file_uploader("Upload Outbound Packing List (.xlsx)", type=['xlsx'])

if packing_file:
    df_raw = pd.read_excel(packing_file, header=None).astype(str)
    
    def get_val(keyword, row_off=0, col_off=0):
        for r in range(len(df_raw)-1, -1, -1):
            for c in range(len(df_raw.columns)):
                cell_val = str(df_raw.iloc[r, c]).lower().strip()
                if keyword.lower() == cell_val:
                    try:
                        return df_raw.iloc[r + row_off, c + col_off]
                    except:
                        return "0"
        return "0"

    def clean_num(val):
        if pd.isna(val) or str(val).lower() == 'nan':
            return 0.0
        clean = re.sub(r'[^\d.]', '', str(val))
        try:
            return float(clean)
        except:
            return 0.0

    # Extracting core data
    pallets_final = int(clean_num(get_val("Pallets", row_off=-1)))
    units_final = int(clean_num(get_val("Units", row_off=-1)))
    lbs_final = clean_num(get_val("Gross Weight", row_off=-1))
    kgs_final = lbs_final * 0.453592

    # Dimension logic
    dim_list = []
    for c in range(len(df_raw.columns)):
        if any("dim" in str(val).lower() and "pallet" in str(val).lower() for val in df_raw.iloc[:5, c]):
            potential_dims = df_raw.iloc[3:, c].tolist()
            dim_list = [d.strip() for d in potential_dims if "x" in str(d).lower() and len(str(d)) > 5]
            break

    dim_counts = Counter(dim_list)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    if pallets_final == 0 and units_final == 0:
        st.error("⚠️ Summary not found. Please check labels in footer.")
    else:
        st.success(f"✅ Data Extracted: **{pallets_final}** Pallets | **{units_final:,}** Units")

    if st.button("🚀 Generate Template"):
        # Excel Generation
        quote_data = [
            ["QUOTE REQUEST", ""],
            ["DESTINATION", destination],
            ["SERVICE", service],
            ["UNITS", f"{units_final:,}"],
            ["PALLETS", pallets_final],
        ]
        
        if formatted_dims:
            quote_data.append(["DIMENSIONS", formatted_dims[0]])
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
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_output.to_excel(writer, index=False, header=False)

        # Email Generation with Bold Topics
        dim_string = "".join([f"\n- **Dimensions**: {d}" for d in formatted_dims])
        email_body = f"""Hi Team,

Hope you are having a great week! 

Please find the details below for a new {service} shipment quote:

- **Destination**: {destination}
- **Service**: {service}
- **Total Units**: {units_final:,}
- **Pallets**: {pallets_final}{dim_string}
- **Total Weight**: {lbs_final:,.2f} LBS | {kgs_final:,.2f} KGS
- **Commodity**: {commodity}
- **Value**: {cargo_value}
- **Incoterms**: {incoterms}

Please let us know the best rates and estimated transit times for this. 

Attached are the Quote Request and Packing List.

Thanks!"""

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("1. Download Document")
            st.download_button("📥 Download Excel", data=buf.getvalue(), file_name=f"Quote_{pallets_final}PLTS.xlsx")
            st.table(df_output)
        with col2:
            st.subheader("2. Email Draft")
            # Fallback copy method: code block with built-in copy icon
            st.code(email_body, language="markdown")
            st.info("💡 Click the 'Copy' icon in the top-right of the box above.")
else:
    st.info("Please upload the Outbound Packing List to begin.")

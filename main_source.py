import streamlit as st
import pandas as pd
from collections import Counter
import io
import re

st.set_page_config(page_title="Logistics Quote Generator", layout="wide")

st.title("📦 Logistics Quote Pipeline")

# --- CUSTOM LISTS ---
destinations = [
    "UK - Radial FAO Monat, Middleton Oldham OL9 9XA",
    "POLAND - Panattoni Park Pruszków IV, ul. Przejazdowa 25, 05-800 Pruszków",
    "SPAIN - MONAT Spain c/o Logifashion, Calle de la Malvasia 2, 19200 Azuqueca de Henares, Guadalajara",
    "LITHUANIA - Girteka Logistics, Metalistų g. 6, LT-78107 Šiauliai",
    "IRELAND - Monat Ireland c/o Ace Express Freight, Unit 3, Blakes Cross Business Park, Lusk, Co. Dublin",
    "OTHER (Type Manually below)"
]

services = [
    "LCL Ocean",
    "LTL Road",
    "FCL Ocean",
    "Air Freight",
    "Courier"
]

# --- SIDEBAR: MANUAL INPUTS ---
with st.sidebar:
    st.header("Shipment Details")
    
    selected_dest = st.selectbox("Select Destination", destinations)
    if selected_dest == "OTHER (Type Manually below)":
        destination = st.text_input("Enter Manual Destination", value="")
    else:
        destination = selected_dest

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

    # 2. GRAB TOTALS FROM FOOTER
    pallet_raw = get_val("Pallets", row_off=-1)
    units_raw = get_val("Units", row_off=-1)
    weight_raw = get_val("Gross Weight", row_off=-1)

    pallets_final = int(clean_num(pallet_raw))
    units_final = int(clean_num(units_raw))
    lbs_final = clean_num(weight_raw)
    kgs_final = lbs_final * 0.453592

    # 3. DYNAMIC DIMENSION FINDER
    dim_list = []
    for c in range(len(df_raw.columns)):
        if any("dim" in str(val).lower() and "pallet" in str(val).lower() for val in df_raw.iloc[:5, c]):
            potential_dims = df_raw.iloc[3:, c].tolist()
            dim_list = [d.strip() for d in potential_dims if "x" in str(d).lower() and len(str(d)) > 5]
            break

    dim_counts = Counter(dim_list)
    formatted_dims = [f"{d} (x{count})" if count > 1 else d for d, count in dim_counts.items()]

    if pallets_final == 0 and units_final == 0:
        st.error("⚠️ Summary not found. Please check 'Pallets' and 'Units' labels in footer.")
    else:
        st.success(f"✅ Data Extracted: {pallets_final} Pallets | {units_final:,} Units")

    # --- GENERATE BUTTON ---
    if st.button("🚀 Generate Template"):
        
        # 4. CONSTRUCT THE EXCEL OUTPUT
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
        else:
            quote_data.append(["DIMENSIONS", "Refer to Packing List"])

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

        # 5. GENERATE PROFESSIONAL EMAIL CONTENT (Table Format)
        # We define dim_rows FIRST so we don't get a NameError
        dim_rows = "\n".join([f"| Dimensions | {d} |" for d in formatted_dims])
        
        email_body = f"""Hi Team,

Hope you are having a great week! 

Please find the details below for a new {service} shipment quote:

- **Destination:** {destination}
- Service: {service}
- Total Units: {units_final:,}
- Pallets: {pallets_final}
{dim_string}
- Total Weight: {lbs_final:,.2f} LBS | {kgs_final:,.2f} KGS
- Commodity: {commodity}
- Value: {cargo_value}
- Incoterms: {incoterms}

Please let us know the best rates and estimated transit times for this. 

Attached are the Quote Request and Packing List.

Thanks!"""

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("1. Download Document")
            st.download_button("📥 Download Excel", data=buf.getvalue(), file_name=f"Quote_{pallets_final}PLTS.xlsx")
            st.table(df_output)
        with c2:
            st.subheader("2. Email Draft")
            # NEW: COPY TO CLIPBOARD BUTTON
            st.copy_to_clipboard(email_body)
            st.info("👆 Click the button above to copy the email.")
            st.markdown(email_body) # Preview of how it looks
else:
    st.info("Please upload the Outbound Packing List to begin.")

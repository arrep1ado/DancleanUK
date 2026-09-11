import streamlit as st
import pandas as pd
import requests
import io
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="DanCleanUK Optimizer", layout="centered")
st.title("Daily Route & Profit Optimizer")

# --- BLOCK MOBILE PULL-TO-REFRESH ---
st.markdown(
    """
    <style>
    body { overscroll-behavior-y: none; }
    html { overscroll-behavior-y: none; }
    </style>
    """,
    unsafe_allow_html=True
)

# --- RESET LOGIC ---
if st.sidebar.button("Start New Day / Reset"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# --- SIDEBAR SETTINGS ---
st.sidebar.title("Settings")
DEPOT_POSTCODE = st.sidebar.text_input("Depot Postcode (Grantham)", value="NG31 9RA")
FUEL_PRICE = st.sidebar.number_input("Fuel Price (£/liter)", value=1.50, step=0.01)
MPG = st.sidebar.number_input("Vehicle MPG", value=30.0, step=0.1)
TAX_RATE = st.sidebar.slider("Tax Deduction (%)", 0, 50, 20) / 100

# --- EXPORT LOGIC ---
st.sidebar.markdown("---")
st.sidebar.title("Export Monthly Records")
if 'master_df' in st.session_state and not st.session_state.master_df.empty:
    export_df = st.session_state.master_df.copy()
    if 'Status' in export_df.columns:
        export_df = export_df[export_df['Status'] != 'depot'].copy()
    for col in ['latitude', 'longitude', 'Status']:
        if col in export_df.columns:
            export_df = export_df.drop(columns=[col])
            
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"
    header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for r in dataframe_to_rows(export_df, index=False, header=True): 
        ws.append(r)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center")
    wb.save(output)
    st.sidebar.download_button("Download Professional Report (XLSX)", output.getvalue(), "DanCleanUK_Records.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.sidebar.info("Upload a file to begin.")
    
# --- APP LOGIC ---
if "API_KEY" not in st.secrets:
    st.error("API_KEY missing in Streamlit secrets. Please add your OpenRouteService API key.")
    st.stop()

API_KEY = st.secrets["API_KEY"]
uploaded_file = st.file_uploader("Upload your Day's File (Headers required: Postcode, Price, Phone)", type=["csv", "xlsx"])

if uploaded_file and 'master_df' not in st.session_state:
    try:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=['Postcode', 'Price', 'Phone'])
        df['Postcode'] = df['Postcode'].astype(str).str.upper().str.strip()
        df['Status'] = 'pending'
        df['Payment'] = 'waiting'
        st.session_state.master_df = df
        st.rerun()
    except Exception as e:
        st.error(f"Error loading file: {e}")

# --- FOOLPROOF MULTI-SOURCE GEOCODER (POSTCODES.IO ACTIVE + TERMINATED + NOMINATIM) ---
def get_coords(postcode):
    cleaned_pc = postcode.upper().strip()
    pc_no_space = cleaned_pc.replace(" ", "")
    
    # 1. Try postcodes.io (Active postcodes)
    try:
        res = requests.get(f"https://api.postcodes.io/postcodes/{pc_no_space}", timeout=4)
        if res.status_code == 200:
            data = res.json().get("result", {})
            lat, lon = data.get("latitude"), data.get("longitude")
            if lat is not None and lon is not None:
                return lat, lon
    except Exception:
        pass

    # 2. Try postcodes.io (Terminated/historical postcodes)
    try:
        res = requests.get(f"https://api.postcodes.io/terminated_postcodes/{pc_no_space}", timeout=4)
        if res.status_code == 200:
            data = res.json().get("result", {})
            lat, lon = data.get("latitude"), data.get("longitude")
            if lat is not None and lon is not None:
                return lat, lon
    except Exception:
        pass

    # 3. Fallback to OpenStreetMap Nominatim API
    url_nom = "https://nominatim.openstreetmap.org/search"
    headers = {'User-Agent': 'DanCleanUKOptimizer/1.0'}
    params = {'q': f"{cleaned_pc}, United Kingdom", 'format': 'json', 'limit': 1}
    try:
        res = requests.get(url_nom, params=params, headers=headers, timeout=5)
        if res.status_code == 200:
            results = res.json()
            if results:
                return float(results[0]['lat']), float(results[0]['lon'])
    except Exception:
        pass

    return None, f"Postcode not found or invalid ({postcode})."

# --- HELPER TO CLEAN VALUES (e.g. remove trailing .0 from Excel numbers) ---
def clean_val(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        try:
            s = str(int(float(s)))
        except ValueError:
            pass
    return s

# --- ROUTING & OPTIMIZATION ---
if 'master_df' in st.session_state:
    if st.button("Optimize Route"):
        all_postcodes = [DEPOT_POSTCODE.upper().strip()] + st.session_state.master_df['Postcode'].tolist()
        
        extra_cols = [col for col in st.session_state.master_df.columns if col not in ['Postcode', 'Price', 'Phone', 'Status', 'Payment', 'latitude', 'longitude']]
        
        routing_data = []
        error_occurred = False
        
        prices = [0.0] + st.session_state.master_df['Price'].tolist()
        phones = [''] + st.session_state.master_df['Phone'].tolist()
        
        extra_data_lists = {}
        for col in extra_cols:
            extra_data_lists[col] = [''] + st.session_state.master_df[col].tolist()

        with st.spinner("Geocoding UK postcodes..."):
            for i, pc in enumerate(all_postcodes):
                lat, lon_or_err = get_coords(pc)
                if lat is None:
                    st.error(f"Could not find coordinates for postcode '{pc}'. Reason: {lon_or_err}")
                    error_occurred = True
                    break
                
                row_dict = {
                    'Postcode': pc,
                    'Price': prices[i],
                    'Phone': phones[i],
                    'latitude': lat,
                    'longitude': lon_or_err
                }
                for col in extra_cols:
                    row_dict[col] = extra_data_lists[col][i]
                
                routing_data.append(row_dict)
                time.sleep(0.2)
        
        if not error_occurred:
            df_routing = pd.DataFrame(routing_data)
            locations = [[float(row['longitude']), float(row['latitude'])] for _, row in df_routing.iterrows()]
            
            with st.spinner("Calculating optimal route matrix..."):
                body = {"locations": locations, "metrics": ["distance"], "units": "km"}
                response = requests.post('https://api.openrouteservice.org/v2/matrix/driving-car', json=body, headers={'Authorization': API_KEY, 'Content-Type': 'application/json'}, timeout=15)
                
                if response.status_code == 200:
                    dist_matrix = response.json()['distances']
                    
                    unvisited = set(range(1, len(locations)))
                    current_node = 0  
                    route_indices = [0]
                    total_meters = 0
                    
                    while unvisited:
                        next_node = min(unvisited, key=lambda j: dist_matrix[current_node][j])
                        total_meters += dist_matrix[current_node][next_node]
                        route_indices.append(next_node)
                        current_node = next_node
                        unvisited.remove(next_node)
                    
                    total_meters += dist_matrix[current_node][0]
                    route_indices.append(0)  
                    
                    df_resolved = df_routing.iloc[route_indices].reset_index(drop=True)
                    
                    start_depot = df_resolved.iloc[[0]].copy()
                    active_jobs = df_resolved[df_resolved.index > 0].iloc[:-1].copy()
                    return_depot = df_resolved.iloc[[-1]].copy()
                    
                    new_master = pd.concat([start_depot, active_jobs, return_depot]).reset_index(drop=True)
                    new_master['Status'] = 'pending'
                    new_master['Payment'] = 'waiting'
                    
                    new_master.loc[0, 'Status'] = 'depot'
                    new_master.loc[len(new_master) - 1, 'Status'] = 'depot'
                    
                    st.session_state.master_df = new_master
                    
                    total_km = total_meters / 1000
                    total_miles = total_km * 0.621371
                    total_fuel_cost = ((total_miles / MPG) * 4.54609) * FUEL_PRICE
                    locked_profit = (st.session_state.master_df['Price'].sum() - total_fuel_cost) * (1 - TAX_RATE)
                    st.session_state.route_data = {"initial_miles": total_miles, "locked_profit": locked_profit}
                    st.rerun()
                else:
                    st.error(f"Matrix Routing API Error {response.status_code}: {response.text}")

    # --- DASHBOARD DISPLAY ---
    if 'route_data' in st.session_state:
        st.write(f"### Planned Daily Take-Home Profit: £{st.session_state.route_data.get('locked_profit', 0):.2f}")
        st.write(f"### Estimated Total Distance: {st.session_state.route_data.get('initial_miles', 0):.2f} miles")
    
    # --- HELPER FUNCTION TO BUILD MAP URL WITH CLEANED DOOR/ADDRESS ---
    def get_map_destination_string(row_data):
        parts = []
        for col_name in st.session_state.master_df.columns:
            if col_name.lower() in ['door', 'door no', 'door number', 'address', 'street', 'name']:
                val = clean_val(row_data.get(col_name))
                if val != '':
                    parts.append(val)
        parts.append(str(row_data['Postcode']))
        return ", ".join(parts)

    # --- OPTIMIZED SIDEBAR NAVIGATION (NEXT STOP EXCLUDING DEPOT) ---
    st.sidebar.markdown("---")
    st.sidebar.title("Route Navigation")
    
    pending_df = st.session_state.master_df[
        (st.session_state.master_df['Status'].str.lower() == 'pending') & 
        (st.session_state.master_df['Status'].str.lower() != 'depot')
    ]
    
    if not pending_df.empty:
        next_row = pending_df.iloc[0]
        next_dest = get_map_destination_string(next_row)
        gmaps_url = f"https://www.google.com/maps/dir/?api=1&destination={next_dest}&travelmode=driving"
        st.sidebar.link_button("🚗 Navigate to Next Stop", gmaps_url)
        st.sidebar.caption(f"Next in sequence: {next_dest} ({len(pending_df)} stops remaining)")
    else:
        st.sidebar.success("All customer stops completed for today!")

    # --- MAIN DISPLAY & ADDRESS CARDS WITH EMBEDDED NAVIGATION ---
    for idx, row in st.session_state.master_df.iterrows():
        postcode = str(row['Postcode'])
        status = str(row['Status'])
        price = row.get('Price', 0)
        phone = row.get('Phone', '')
        payment = str(row.get('Payment', 'waiting'))
        
        # Build extra info text for display cards (excluding dates/times)
        extra_info_parts = []
        for col_name in st.session_state.master_df.columns:
            if col_name.lower() in ['postcode', 'price', 'phone', 'status', 'payment', 'latitude', 'longitude']:
                continue
            if 'date' in col_name.lower() or 'time' in col_name.lower():
                continue  # Skip displaying date and time columns
            
            val = clean_val(row.get(col_name))
            if val != '':
                extra_info_parts.append(f"**{col_name}:** {val}")
        
        extra_text = " | ".join(extra_info_parts)
        if extra_text:
            extra_text = f" | {extra_text}"
        
        is_start_depot = (idx == 0)
        is_return_depot = (idx == len(st.session_state.master_df) - 1)
        
        dest_string = get_map_destination_string(row)
        map_url = f"https://www.google.com/maps/dir/?api=1&destination={dest_string}&travelmode=driving"

        if is_start_depot:
            with st.container(border=True):
                st.write(f"**📍 Start Depot (Grantham)** ({postcode})")
        elif is_return_depot:
            with st.container(border=True):
                st.write(f"**🏁 Return to Depot (Grantham)** ({postcode})")
                st.link_button("🚗 Navigate Here", map_url, key=f"nav_{idx}")
        else:
            status_icon = "✅" if status.lower() == 'completed' else "⏳"
            status_text = "Completed" if status.lower() == 'completed' else "Pending"
            payment_display = f" | **Payment:** {payment}"
            
            with st.container(border=True):
                st.write(f"**{status_icon} Postcode: {postcode}**{extra_text} | **Price:** £{price} | **Status:** {status_text}{payment_display}")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if status.lower() == 'pending':
                        if st.button("Mark Complete", key=f"complete_{idx}"):
                            st.session_state.master_df.at[idx, 'Status'] = 'completed'
                            st.rerun()
                    else:
                        st.success("Completed")
                        if phone:
                            msg = f"Hi from DanCleanUK! Your service is complete today. Total: £{price}. Please pay via bank transfer to Mettle - Sort Code: 04-03-33 | Account: 72515806. Thank you!"
                            wa_url = f"https://wa.me/{phone}?text={msg.replace(' ', '%20')}"
                            st.link_button("Send WhatsApp", wa_url, key=f"wa_{idx}")
                    
                    c_col1, c_col2 = st.columns(2)
                    with c_col1:
                        if st.button("💵 Cash Paid", key=f"cash_{idx}"):
                            st.session_state.master_df.at[idx, 'Payment'] = 'Cash'
                            st.rerun()
                    with c_col2:
                        if st.button("❌ Not Paid", key=f"not_paid_{idx}"):
                            st.session_state.master_df.at[idx, 'Payment'] = 'Not Paid'
                            st.rerun()
                        
                with col2:
                    st.link_button("🚗 Navigate Here", map_url, key=f"nav_{idx}")
else:
    st.info("Upload your day's file to begin.")

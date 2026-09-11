import streamlit as st
import pandas as pd
import requests
import io
import time
import math
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
DEPOT_FULL_ADDRESS = "192 Queensway, Grantham NG31 9RA"
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
    for col in ['latitude', 'longitude', 'Status', 'geo_query']:
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
        
        # Completely drop blank/empty rows imported from Excel trailing cells
        df = df.dropna(how='all')
        
        # Strict validation: Drop rows where Postcode is missing, blank, or NaN
        df = df.dropna(subset=['Postcode'])
        df['Postcode'] = df['Postcode'].astype(str).str.upper().str.strip()
        df = df[(df['Postcode'] != '') & (df['Postcode'] != 'NAN') & (df['Postcode'] != 'NAT')]
        
        # Strict validation: Drop rows without a valid positive Price
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
        df = df.dropna(subset=['Price'])
        df = df[df['Price'] > 0]
        
        # Strict validation: Drop rows without a valid Phone number
        df['Phone'] = df['Phone'].astype(str).str.strip()
        df = df[(df['Phone'] != '') & (df['Phone'].lower() != 'nan') & (df['Phone'].lower() != 'nat')]
        
        df['Status'] = 'pending'
        df['Payment'] = 'waiting'
        st.session_state.master_df = df.reset_index(drop=True)
        st.rerun()
    except Exception as e:
        st.error(f"Error loading file: {e}")

# --- BULLETPROOF MULTI-TIER GEOCODER ---
def get_coords(query_string, postcode_fallback):
    headers = {'User-Agent': 'DanCleanUKOptimizer/1.0'}
    
    query = str(query_string).strip()
    if query and query.lower() != 'nan':
        full_q = query if ("uk" in query.lower() or "united kingdom" in query.lower()) else f"{query}, United Kingdom"
        try:
            res = requests.get("https://nominatim.openstreetmap.org/search", params={'q': full_q, 'format': 'json', 'limit': 1}, headers=headers, timeout=4)
            if res.status_code == 200 and res.json():
                return float(res.json()[0]['lat']), float(res.json()[0]['lon'])
        except Exception:
            pass

    pc_clean = str(postcode_fallback).upper().strip()
    if pc_clean:
        try:
            res = requests.get("https://nominatim.openstreetmap.org/search", params={'q': f"{pc_clean}, United Kingdom", 'format': 'json', 'limit': 1}, headers=headers, timeout=4)
            if res.status_code == 200 and res.json():
                return float(res.json()[0]['lat']), float(res.json()[0]['lon'])
        except Exception:
            pass

        pc_no_space = pc_clean.replace(" ", "")
        try:
            res = requests.get(f"https://api.postcodes.io/postcodes/{pc_no_space}", timeout=3)
            if res.status_code == 200:
                data = res.json().get("result", {})
                lat, lon = data.get("latitude"), data.get("longitude")
                if lat is not None and lon is not None:
                    return lat, lon
        except Exception:
            pass

        try:
            res = requests.get(f"https://api.postcodes.io/terminated_postcodes/{pc_no_space}", timeout=3)
            if res.status_code == 200:
                data = res.json().get("result", {})
                lat, lon = data.get("latitude"), data.get("longitude")
                if lat is not None and lon is not None:
                    return lat, lon
        except Exception:
            pass

    return 52.9141, -0.6414

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

def calculate_haversine_matrix(locations):
    n = len(locations)
    matrix = [[0.0 * n for _ in range(n)] for _ in range(n)]
    R = 6371.0 
    ROAD_FACTOR = 1.3 
    
    for i in range(n):
        lon1, lat1 = locations[i]
        for j in range(n):
            lon2, lat2 = locations[j]
            if i == j:
                matrix[i][j] = 0.0
            else:
                dlat = math.radians(lat2 - lat1)
                dlon = math.radians(lon2 - lon1)
                a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
                c = 2 * math.asin(math.sqrt(a))
                matrix[i][j] = R * c * 1000 * ROAD_FACTOR
    return matrix

def optimize_route_2opt(route_indices, dist_matrix):
    best_route = route_indices[:]
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best_route) - 2):
            for j in range(i + 1, len(best_route) - 1):
                a, b = best_route[i-1], best_route[i]
                c, d = best_route[j], best_route[j+1]
                
                current_cost = dist_matrix[a][b] + dist_matrix[c][d]
                new_cost = dist_matrix[a][c] + dist_matrix[b][d]
                
                if new_cost < current_cost:
                    best_route[i:j+1] = reversed(best_route[i:j+1])
                    improved = True
    return best_route

def build_geo_query(row_data, default_postcode):
    parts = []
    for col_name in row_data.index:
        if col_name.lower() in ['address', 'street', 'location', 'name', 'house']:
            val = clean_val(row_data[col_name])
            if val != '':
                parts.append(val)
    pc = str(row_data.get('Postcode', '')).strip()
    if pc and pc.upper() != 'NAN':
        parts.append(pc)
    else:
        parts.append(default_postcode)
    return ", ".join(parts)

# --- ROUTING & OPTIMIZATION ---
if 'master_df' in st.session_state:
    if st.button("Optimize Route"):
        depot_query = DEPOT_FULL_ADDRESS
        row_queries = [build_geo_query(row, DEPOT_POSTCODE) for _, row in st.session_state.master_df.iterrows()]
        all_queries = [depot_query] + row_queries
        
        all_postcodes = [DEPOT_POSTCODE.upper().strip()] + st.session_state.master_df['Postcode'].tolist()
        
        extra_cols = [col for col in st.session_state.master_df.columns if col not in ['Postcode', 'Price', 'Phone', 'Status', 'Payment', 'latitude', 'longitude', 'geo_query']]
        
        routing_data = []
        prices = [0.0] + st.session_state.master_df['Price'].tolist()
        phones = [''] + st.session_state.master_df['Phone'].tolist()
        
        extra_data_lists = {}
        for col in extra_cols:
            extra_data_lists[col] = [''] + st.session_state.master_df[col].tolist()

        with st.spinner("Geocoding addresses & postcodes safely..."):
            for i, q in enumerate(all_queries):
                pc_fallback = all_postcodes[i]
                lat, lon = get_coords(q, pc_fallback)
                
                row_dict = {
                    'Postcode': pc_fallback,
                    'Price': prices[i],
                    'Phone': phones[i],
                    'geo_query': q,
                    'latitude': lat,
                    'longitude': lon
                }
                for col in extra_cols:
                    row_dict[col] = extra_data_lists[col][i]
                
                routing_data.append(row_dict)
                time.sleep(0.1)
        
        df_routing = pd.DataFrame(routing_data)
        locations = [[float(row['longitude']), float(row['latitude'])] for _, row in df_routing.iterrows()]
        
        dist_matrix = None
        with st.spinner("Calculating optimal route matrix..."):
            try:
                body = {"locations": locations, "metrics": ["distance"], "units": "km"}
                response = requests.post('https://api.openrouteservice.org/v2/matrix/driving-car', json=body, headers={'Authorization': API_KEY, 'Content-Type': 'application/json'}, timeout=10)
                
                if response.status_code == 200:
                    dist_matrix = response.json()['distances']
                else:
                    st.warning("Routing API limit reached. Using offline smart routing fallback...")
            except Exception:
                pass
            
            if dist_matrix is None:
                dist_matrix = calculate_haversine_matrix(locations)
        
        if dist_matrix is not None:
            # TRUE NEAREST-NEIGHBOR CHAINING: Always pick the absolute closest next stop from current location
            unvisited = set(range(1, len(locations)))
            current_node = 0  
            route_indices = [0]
            
            while unvisited:
                next_node = min(unvisited, key=lambda j: dist_matrix[current_node][j])
                route_indices.append(next_node)
                current_node = next_node
                unvisited.remove(next_node)
            
            route_indices.append(0)  
            
            # Polish route using 2-opt
            route_indices = optimize_route_2opt(route_indices, dist_matrix)
            
            total_meters = sum(dist_matrix[route_indices[k]][route_indices[k+1]] for k in range(len(route_indices) - 1))
            
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

    # --- DASHBOARD DISPLAY ---
    if 'route_data' in st.session_state:
        st.write(f"### Planned Daily Take-Home Profit: £{st.session_state.route_data.get('locked_profit', 0):.2f}")
        st.write(f"### Estimated Total Distance: {st.session_state.route_data.get('initial_miles', 0):.2f} miles")
    
    def get_map_destination_string(row_data, is_depot=False):
        if is_depot:
            return DEPOT_FULL_ADDRESS
        if 'geo_query' in row_data and pd.notna(row_data['geo_query']) and str(row_data['geo_query']).strip() != '':
            return str(row_data['geo_query'])
        
        parts = []
        for col_name in st.session_state.master_df.columns:
            if col_name.lower() in ['address', 'street', 'location', 'name']:
                val = clean_val(row_data.get(col_name))
                if val != '':
                    parts.append(val)
        parts.append(str(row_data['Postcode']))
        return ", ".join(parts)

    # --- OPTIMIZED SIDEBAR NAVIGATION ---
    if 'route_data' in st.session_state:
        st.sidebar.markdown("---")
        st.sidebar.title("Route Navigation")
        
        max_idx = len(st.session_state.master_df) - 1
        pending_df = st.session_state.master_df.loc[
            (st.session_state.master_df['Status'].str.lower() == 'pending') & 
            (st.session_state.master_df.index > 0) & 
            (st.session_state.master_df.index < max_idx)
        ]
        
        if not pending_df.empty:
            next_row = pending_df.iloc[0]
            next_dest = get_map_destination_string(next_row, is_depot=False)
            gmaps_url = f"https://www.google.com/maps/dir/?api=1&destination={next_dest}&travelmode=driving"
            st.sidebar.link_button("🚗 Navigate to Next Stop", gmaps_url)
            st.sidebar.caption(f"Next in sequence: {next_dest} ({len(pending_df)} stops remaining)")
        else:
            st.sidebar.success("All customer stops completed for today!")

    # --- MAIN DISPLAY & ADDRESS CARDS ---
    for idx, row in st.session_state.master_df.iterrows():
        postcode = str(row['Postcode'])
        status = str(row['Status'])
        price = row.get('Price', 0)
        phone = row.get('Phone', '')
        payment = str(row.get('Payment', 'waiting'))
        
        is_start_depot = (idx == 0)
        is_return_depot = (idx == len(st.session_state.master_df) - 1)
        
        if is_start_depot:
            with st.container(border=True):
                st.write(f"**📍 Start Depot:** {DEPOT_FULL_ADDRESS}")
            continue
        elif is_return_depot:
            dest_string = get_map_destination_string(row, is_depot=True)
            map_url = f"https://www.google.com/maps/dir/?api=1&destination={dest_string}&travelmode=driving"
            with st.container(border=True):
                st.write(f"**🏁 Return to Depot:** {DEPOT_FULL_ADDRESS}")
                st.link_button("🚗 Navigate Here", map_url, key=f"nav_{idx}")
            continue

        # Regular Customer Stops Only Below
        extra_info_parts = []
        for col_name in st.session_state.master_df.columns:
            if col_name.lower() in ['postcode', 'price', 'phone', 'status', 'payment', 'latitude', 'longitude', 'geo_query']:
                continue
            if 'date' in col_name.lower() or 'time' in col_name.lower():
                continue  
            
            val = clean_val(row.get(col_name))
            if val != '':
                extra_info_parts.append(f"**{col_name}:** {val}")
        
        extra_text = " | ".join(extra_info_parts)
        if extra_text:
            extra_text = f" | {extra_text}"
        
        dest_string = get_map_destination_string(row, is_depot=False)
        map_url = f"https://www.google.com/maps/dir/?api=1&destination={dest_string}&travelmode=driving"

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

import json
import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify

app = Flask(__name__)
app.secret_key = 'sinani_secret'

# —————————————————————————————
# STATIC LOOKUPS
# —————————————————————————————
CATEGORIES = [
    "Airports", "Bus Stations", "Convention Centre", "Data Centres", "Employee Residence",
    "Farms", "Government Building", "Health and Safety", "Hospital", "Hotels And Lodges",
    "Logistics", "Manufacturing/Production/Packaging Plants/Warehousing", "Medicals",
    "Mining", "Office Parks", "Residential", "Retail Park", "Service Provider",
    "Shopping Malls", "Sinani Offices", "Site Accommodation", "Supplier", "Training",
    "Vehicle and Trailer Service Centre"
]

SUBCATEGORIES = [
    "Branding", "Capstone", "Citrus Farm", "Client", "Crown Foods", "Dairy Farm",
    "EMIRA", "Equipment Supplier", "External Storage", "Gauteng", "Groceries",
    "Growthpoint", "Isuzu", "Jonsson", "Just Milk", "Macsteel", "Management Accommodation",
    "Medicals", "N2P", "Nissan", "PACE", "Panelbeaters", "Piggery", "PPE Supplier",
    "Recycling", "Redefine Properties", "Sanral", "Shoprite", "Sibanye",
    "Site Team Accommodation", "Stationery", "Stock and Tool Supplier", "Stock Supplier",
    "Teraco", "Tool Supplier", "Tool Supplier, Repair and Service", "Toyota",
    "Traffic Department", "Trailer Rental", "Twizza", "Tyres and Alignment",
    "Unifrutti", "Vehicle Accessories", "Waste Management"
]

TRIP_TYPES = [
    "Site", "Operations", "Maintenance", "Repairs", "Expense", "Outreach"
]

USERS = {
    "admin":       ("admin123", "admin"),
    "Jesse":       ("Jesse123", "admin"),
    "Cameron":     ("Cameron123", "user"),
    "Christopher": ("ImGay123",  "admin"),
}

# —————————————————————————————
# 1) LOAD EXTERNAL DATA HELPERS
# —————————————————————————————
def load_site_data():
    path = os.path.join(app.root_path, "GeofenceTable.xlsx")
    try:
        df = pd.read_excel(path)
        return df.fillna("")
    except Exception as e:
        print(f"[WARN] load_site_data failed: {e}")
        return pd.DataFrame({
            "Category":       ["Farm", "Accommodation"],
            "Sub-Category":   ["Dairy", "Team Housing"],
            "Geofence Name":  ["Galliers", "Komani Base"],
            "Address":        ["Farm Road 1", "61 Livingstone St"],
            "Coordinates":    ["-29.27,29.93", "-31.89,26.88"],
        })

def load_trip_logs_data():
    path = os.path.join(app.root_path, "TripLogs.xlsx")
    cols = [
        "Registration", "Model",
        "Name and Surname", "Department", "HOD",
        "Role", "Number", "Email",
        "Date", "Start Location", "End Location",
        "Trip Type", "Notes"
    ]
    try:
        df = pd.read_excel(path)
        return df.fillna("")
    except FileNotFoundError:
        return pd.DataFrame(columns=cols)

def load_driver_data():
    path = os.path.join(app.root_path, "DriverTable.xlsx")
    try:
        df = pd.read_excel(path)
    except FileNotFoundError:
        return pd.DataFrame(columns=[
            "Name and Surname", "Department", "HOD", "Role", "Email"
        ])

    df = df[df["Status"] == "Active"]
    df["Name and Surname"] = df["Name"].str.strip() + " " + df["Surname"].str.strip()
    return df[[
        "Name and Surname",
        "Department",
        "HOD",
        "Role",
        "Email"
    ]]

def load_vehicle_data():
    path = os.path.join(app.root_path, "Vehicle Info.xlsx")
    try:
        df = pd.read_excel(path)
    except FileNotFoundError:
        return pd.DataFrame(columns=["Registration", "Model"])

    if "Staus" in df.columns and "Status" not in df.columns:
        df = df.rename(columns={"Staus": "Status"})

    if "Status" not in df.columns:
        return pd.DataFrame(columns=["Registration", "Model"])

    df = df[df["Status"] == "Active"]
    return df[["Registration", "Model"]]

# —————————————————————————————
# 2) ROUTES
# —————————————————————————————
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = request.form['username']
        p = request.form['password']
        user = USERS.get(u)
        if user and user[0] == p:
            session['username'] = u
            session['role']     = user[1]
            return redirect(url_for('home'))
        flash('Invalid credentials', 'danger')
    return render_template('login.html')

@app.route('/home')
def home():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('home.html',
        username=session['username'],
        role=session['role']
    )

@app.route('/vehicle')
def vehicle():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('vehicle.html', role=session['role'])

@app.route('/trip_logs', methods=['GET', 'POST'])
def trip_logs():
    if 'username' not in session:
        return redirect(url_for('login'))

    df       = load_trip_logs_data()
    vehicles = load_vehicle_data()
    drivers  = load_driver_data()
    sites    = load_site_data()["Geofence Name"].dropna().unique().tolist()

    vehicle_regs = vehicles["Registration"].tolist()
    model_map    = vehicles.set_index("Registration")["Model"].to_dict()

    driver_names = drivers["Name and Surname"].tolist()
    dept_map     = drivers.set_index("Name and Surname")["Department"].to_dict()
    hod_map      = drivers.set_index("Name and Surname")["HOD"].to_dict()
    role_map     = drivers.set_index("Name and Surname")["Role"].to_dict()
    email_map    = drivers.set_index("Name and Surname")["Email"].to_dict()

    if request.method == 'POST':
        new = { col: request.form.get(col, "") for col in df.columns }
        df  = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
        df.to_excel(os.path.join(app.root_path, "TripLogs.xlsx"), index=False)
        return redirect(url_for('trip_logs'))

    return render_template('trip_logs.html',
        data          = df.to_dict(orient='records'),
        columns       = df.columns.tolist(),
        role          = session['role'],
        dept_map      = dept_map,
        hod_map       = hod_map,
        role_map      = role_map,
        email_map     = email_map,
        vehicle_regs  = vehicle_regs,
        driver_names  = driver_names,
        site_names    = sites,
        trip_types    = TRIP_TYPES,
        model_map     = model_map
    )

@app.route('/trip_logs_table')
def trip_logs_table():
    if 'username' not in session:
        return redirect(url_for('login'))

    df = load_trip_logs_data()
    q = request.args.get('q', '').strip()

    if q:
        mask = df['Name and Surname'].astype(str).str.contains(q, case=False, na=False)
        df = df.loc[mask]

    return render_template(
        'trip_logs_table.html',
        data=df.to_dict(orient='records'),
        columns=df.columns.tolist(),
        q=q
    )

@app.route('/delete_trip_log', methods=['POST'])
def delete_trip_log():
    idx = int(request.form['idx'])
    df  = load_trip_logs_data().drop(index=idx).reset_index(drop=True)
    df.to_excel(os.path.join(app.root_path, "TripLogs.xlsx"), index=False)
    return redirect(url_for('trip_logs'))

@app.route('/edit_trip_log', methods=['POST'])
def edit_trip_log():
    data = request.form.to_dict()
    idx  = data.pop('idx')
    return render_template('edit_trip_log.html', data=data, idx=idx)

@app.route('/update_trip_log', methods=['POST'])
def update_trip_log():
    form = request.form.to_dict()
    idx  = int(form.pop('idx'))
    df   = load_trip_logs_data()
    for k, v in form.items():
        df.at[idx, k] = v
    df.to_excel(os.path.join(app.root_path, "TripLogs.xlsx"), index=False)
    return redirect(url_for('trip_logs'))

@app.route('/sites', methods=['GET', 'POST'])
def sites():
    if 'username' not in session:
        return redirect(url_for('login'))

    df         = load_site_data()
    excel_path = os.path.join(app.root_path, "GeofenceTable.xlsx")

    if request.method == 'POST' and request.form.get('Category') is not None:
        new_row = {
            "Category":      request.form['Category'],
            "Sub-Category":  request.form['Sub-Category'],
            "Geofence Name": request.form['Geofence Name'],
            "Address":       request.form['Address'],
            "Coordinates":   request.form['Coordinates'],
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(excel_path, index=False)

    q = request.args.get('q', '').strip()
    if q:
        mask = (
            df['Geofence Name'].astype(str).str.contains(q, case=False, na=False) |
            df['Category'].astype(str).str.contains(q, case=False, na=False) |
            df['Sub-Category'].astype(str).str.contains(q, case=False, na=False) |
            df['Address'].astype(str).str.contains(q, case=False, na=False)
        )
        df = df.loc[mask].copy()

    return render_template('site_manager.html',
        data          = df.to_dict(orient='records'),
        role          = session['role'],
        categories    = CATEGORIES,
        subcategories = SUBCATEGORIES,
        q             = q,
        total_count   = len(df)
    )

@app.route('/api/site_lookup')
def api_site_lookup():
    if 'username' not in session:
        return jsonify([])

    term = request.args.get('term', '').strip()
    df = load_site_data()

    if term:
        mask = df['Geofence Name'].astype(str).str.contains(term, case=False, na=False)
        df = df.loc[mask]

    names = df['Geofence Name'].dropna().astype(str).drop_duplicates().head(15).tolist()
    return jsonify(names)

@app.route('/delete_site', methods=['POST'])
def delete_site():
    geofence = request.form['Geofence Name']
    df        = load_site_data()
    df        = df[df['Geofence Name'] != geofence]
    df.to_excel(os.path.join(app.root_path, "GeofenceTable.xlsx"), index=False)
    return redirect(url_for('sites'))

@app.route('/edit_site', methods=['POST'])
def edit_site():
    data = request.form.to_dict()
    return render_template('edit_site.html', data=data)

@app.route('/update_site', methods=['POST'])
def update_site():
    updated  = request.form.to_dict()
    original = updated.pop('original_name')
    df       = load_site_data()
    df.loc[df['Geofence Name'] == original, list(updated.keys())] = list(updated.values())
    df.to_excel(os.path.join(app.root_path, "GeofenceTable.xlsx"), index=False)
    return redirect(url_for('sites'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)

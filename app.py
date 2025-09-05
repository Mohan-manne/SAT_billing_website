from datetime import datetime
import os
import datetime
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, session, flash
from xhtml2pdf import pisa
from threading import Lock
from num2words import num2words
from io import BytesIO
import os
from flask import Flask, render_template, url_for, send_from_directory, send_file, session, flash, redirect
from flask import Flask, render_template, request, jsonify
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json


app = Flask(__name__)
# Path where your Word file is stored
DOCS_DIR = os.path.join(app.root_path, "static", "docs")
DOCX_FILENAME = "new_10_bill.docx"   # Make sure this matches your file name exactly

app.secret_key = os.environ.get('FLASK_SECRET_KEY', os.urandom(24))
excel_lock = Lock()

def generate_bill_no():
    now = datetime.now()
    date_str = now.strftime("%d%m%y")
    file_path = "bill_counter.txt"
    with excel_lock:
        if os.path.exists(file_path):
            with open(file_path, "r") as f:
                counter = int(f.read().strip()) + 1
        else:
            counter = 1
        with open(file_path, "w") as f:
            f.write(str(counter).zfill(3))
    return f"{date_str}{counter}"



USERS_FILE = "users.json"

# Helper to load users
def load_users():
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, "r") as f:
        return json.load(f)

# Helper to save users
def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f)

# -------- LOGIN ROUTE --------
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        users = load_users()
        
        if username in users and users[username] == password:
            session["user"] = username
            flash("✅ Login successful!", "success")
            return redirect("/menu")
        else:
            flash("❌ Enter the correct details", "error")
            return redirect("/")
    return render_template("welcome.html")

# -------- SIGNUP ROUTE --------
@app.route("/signup", methods=["GET", "POST"])
def signup():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        users = load_users()
        
        if username in users:
            flash("❌ Username already exists!", "error")
        else:
            users[username] = password
            save_users(users)
            flash("✅ Account created! You can login now.", "success")
            return redirect("/")
    return render_template("signup.html")




@app.route("/menu")
def menu():
    if "user" not in session:
        flash("⚠️ Please log in first.", "warning")
        return redirect("/")
    return render_template("menu.html", user=session["user"])

# Inside your existing app.py, update only the sale_bill route like this:

@app.route("/sale-bill", methods=["GET", "POST"])
def sale_bill():
    if "user" not in session:
        flash("⚠️ Please log in to continue.", "warning")
        return redirect("/")

    if request.method == "POST":
        try:
            data = request.form
            now = datetime.now()
            bill_no = generate_bill_no()
            date = now.strftime("%d-%m-%Y")

            mill_name = data["mill_name"]
            mill_code = data["mill_code"]
            farmer_name = data["farmer_name"]  # Added
            bags = int(data["bags"])
            ntwt = float(data["ntwt"])
            price = float(data["price"])
            calc_type = data["calc_type"]

            stwt = 0
            sut_rate = 0

            if calc_type == "1":
                net_bags = ntwt / 77
            elif calc_type == "2":
                sut_rate = float(data.get("sut_rate", 0))
                stwt = bags * sut_rate
                net_bags = (ntwt - stwt) / 75
            else:
                net_bags = 0

            amount = net_bags * price

            commission = amount / 100 if data.get("commission") == "yes" else 0
            hamali_rate = float(data.get("hamali_rate", 0)) if data.get("hamali") == "yes" else 0
            hamali = bags * hamali_rate

            gunny_rate = float(data.get("gunny_rate", 0)) if data.get("gunny_bags") == "yes" else 0
            gunny = bags * gunny_rate

            advance = float(data.get("advance", 0))
            grand_total = amount + commission + hamali + gunny + advance

            lorry_no = data["lorry_no"]
            mobile_no = data["mobile_no"]

            bill_data = {
                "bill_type": "Sale",
                "bill_no": bill_no,
                "date": date,
                "mill_name": mill_name,
                "mill_code": mill_code,
                "farmer_name": farmer_name,  # Added for Excel
                "bags": bags,
                "ntwt": ntwt,
                "stwt": round(stwt, 2),
                "sut_rate": sut_rate,
                "price": price,
                "net_bags": round(net_bags, 2),
                "amount": round(amount, 2),
                "commission": round(commission, 2),
                "hamali": round(hamali, 2),
                "gunny": round(gunny, 2),
                "advance": round(advance, 2),
                "grand_total": round(grand_total, 2),
                "lorry_no": lorry_no,
                "mobile_no": mobile_no
            }

            df = pd.DataFrame([bill_data])
            file_path = "sale_bills.xlsx"
            with excel_lock:
                if os.path.exists(file_path):
                    df_existing = pd.read_excel(file_path, engine='openpyxl')
                    df = pd.concat([df_existing, df], ignore_index=True)
                df.to_excel(file_path, index=False, engine='openpyxl')
            
            # Separate data for PDF: exclude farmer_name
            bill_data_pdf = bill_data.copy()
            bill_data_pdf.pop("farmer_name")  # Exclude from PDF

            pdf_file = f"generated_pdfs/{bill_no}.pdf"
            os.makedirs("generated_pdfs", exist_ok=True)
            with open(pdf_file, "wb") as f:
                html = render_template("bill_template.html", **bill_data)
                pisa.CreatePDF(html, dest=f)

            return send_file(pdf_file, as_attachment=True)
        except Exception as e:
            flash("⚠️ Hey Buddy Please enter the data before proceeding.", "error")
            return redirect("/sale-bill")

    return render_template("sale_bill.html")

from datetime import datetime   # make sure this is at the top

@app.route("/purchase-bill", methods=["GET", "POST"])
def purchase_bill():
    if "user" not in session:
        flash("⚠️ Please log in to continue.", "warning")
        return redirect("/")

    if request.method == "POST":
        try:
            data = request.form
            now = datetime.now()   # ✅ fixed
            bill_no = generate_bill_no()
            date = now.strftime("%d-%m-%Y")

            farmer_name = data["farmer_name"]
            village_name = data["village_name"]
            mill_name = data["mill_name"]
            bags = int(data["bags"])
            ntwt = float(data["ntwt"])
            sut_rate = float(request.form['sut_rate'])
            stwt = bags * sut_rate
            total_ntwt = (ntwt - stwt) / 75
            rate = float(data["rate"])
            amount = total_ntwt * rate
            hamali_rate = float(data["hamali_rate"])
            hamali = hamali_rate * bags
            weigh_bridge = float(data["weigh_bridge"])
            grand_total = amount - hamali - weigh_bridge
            lorry_no = data["lorry_no"]

            bill_data_excel = {
                "bill_type": "Purchase",
                "bill_no": bill_no,
                "date": date,
                "farmer_name": farmer_name,
                "village_name": village_name,
                "mill_name": mill_name,
                "bags": bags,
                "ntwt": ntwt,
                "sut_rate": sut_rate,
                "stwt": stwt,
                "total_ntwt": round(total_ntwt, 2),
                "rate": rate,
                "amount": round(amount, 2),
                "hamali": round(hamali, 2),
                "weigh_bridge": round(weigh_bridge, 2),
                "grand_total": round(grand_total, 2),
                "lorry_no": lorry_no
            }

            bill_data_pdf = bill_data_excel.copy()
            bill_data_pdf.pop("mill_name")

            df = pd.DataFrame([bill_data_excel])
            file_path = "purchase_bills.xlsx"
            with excel_lock:
                if os.path.exists(file_path):
                    df_existing = pd.read_excel(file_path, engine='openpyxl')
                    df = pd.concat([df_existing, df], ignore_index=True)
                df.to_excel(file_path, index=False, engine='openpyxl')

            pdf_file = f"generated_pdfs/{bill_no}.pdf"
            os.makedirs("generated_pdfs", exist_ok=True)
            with open(pdf_file, "wb") as f:
                html = render_template("purchase_bill_template.html", **bill_data_pdf)
                pisa.CreatePDF(html, dest=f)

            return send_file(pdf_file, as_attachment=True)
        except Exception as e:
            flash("⚠️ Hey Buddy Please enter the data before proceeding.", "error")
            return redirect("/purchase-bill")

    return render_template("purchase_bill.html")



@app.route("/transportation-bill", methods=["GET", "POST"])
def transportation_bill():
    if "user" not in session:
        flash("⚠️ Please log in to continue.", "warning")
        return redirect("/")

    if request.method == "POST":
        try:
            data = request.form
            bill_no = generate_bill_no()
            date = datetime.now().strftime("%d-%m-%Y")  # Auto-generate date

            bags = int(data.get("bags", 0))
            kgs = float(data.get("kgs", 0))
            lorry_freight = float(data.get("lorry_freight", 0))
            advance = float(data.get("advance", 0))

           # Convert freight amount to words in Indian Rupees
            whole_part = int(lorry_freight)
            decimal_part = int((lorry_freight % 1) * 100)
            rupees_in_words = num2words(whole_part, lang='en_IN').replace('-', ' ').title()
            paise_in_words = num2words(decimal_part, lang='en_IN').replace('-', ' ').title() if decimal_part > 0 else "Zero"
            freight_in_words = f"{rupees_in_words} Rupees and {paise_in_words} Paise Only"

            bill_data = {
                "bill_type": "Transportation",
                "bill_no": bill_no,
                "date": date,
                "ref": bill_no,
                "ms": data.get("ms", ""),
                "from_location": data.get("from_location", ""),
                "to_location": data.get("to_location", ""),
                "bags": bags,
                "kgs": kgs,
                "rice_type": data.get("rice_type", ""),
                "lorry_no": data.get("lorry_no", ""),
                "lorry_freight": lorry_freight,
                "advance": advance,
                "mobile_no": data.get("mobile_no", ""),
                "freight_in_words": freight_in_words
            }

            df = pd.DataFrame([bill_data])
            file_path = "bills.xlsx"
            with excel_lock:
                if os.path.exists(file_path):
                    df_existing = pd.read_excel(file_path, engine='openpyxl')
                    df = pd.concat([df_existing, df], ignore_index=True)
                df.to_excel(file_path, index=False, engine='openpyxl')

            pdf_file = f"generated_pdfs/{bill_no}.pdf"
            os.makedirs("generated_pdfs", exist_ok=True)
            with open(pdf_file, "wb") as f:
                html = render_template("transportation_bill_template.html", **bill_data)
                pisa.CreatePDF(html, dest=f)

            flash(f"✅ Transportation Bill {bill_no} created!", "success")
            return send_file(pdf_file, as_attachment=True)
        except Exception as e:
            flash("⚠️ Hey Buddy Please enter the data before proceeding.", "error")
            return redirect("/transportation-bill")

    return render_template("transportation_bill.html")


@app.route("/view-bills", methods=["GET", "POST"])
def view_bills():
    if "user" not in session:
        return redirect("/")

    file_path = "sale_bills.xlsx"

    # --- Handle deletion ---
    if request.method == "POST":
        selected_bills = request.form.getlist("delete_ids")
        if os.path.exists(file_path) and selected_bills:
            df = pd.read_excel(file_path, engine='openpyxl')
            df = df[~df['bill_no'].astype(str).isin(selected_bills)]
            df.to_excel(file_path, index=False, engine='openpyxl')
            flash(f"✅ {len(selected_bills)} Sale Bill(s) deleted.", "success")
            return redirect("/view-bills")

    # --- Load bills ---
    bills = []
    if os.path.exists(file_path):
        with excel_lock:
            df = pd.read_excel(file_path, engine='openpyxl')
            df = df.drop_duplicates(subset='bill_no', keep='last')

            # --- FILTER LOGIC ---
            date = request.args.get("date")
            bill_no = request.args.get("bill_no")
            mill_name = request.args.get("mill_name")
            farmer_name = request.args.get("farmer_name")

            if date:
                df = df[df['date'] == date]
            if bill_no:
                df = df[df['bill_no'].astype(str) == bill_no]
            if mill_name:
                df = df[df['mill_name'] == mill_name]
            if farmer_name:
                df = df[df['farmer_name'] == farmer_name]

            # --- For dropdowns ---
            unique_bill_nos = df['bill_no'].dropna().astype(str).unique()
            unique_mill_names = df['mill_name'].dropna().unique()
            unique_farmer_names = df['farmer_name'].dropna().unique()

            bills = df.to_dict(orient='records')

    return render_template(
        "view_bills_sale.html",
        bills=bills,
        unique_bill_nos=unique_bill_nos if 'unique_bill_nos' in locals() else [],
        unique_mill_names=unique_mill_names if 'unique_mill_names' in locals() else [],
        unique_farmer_names=unique_farmer_names if 'unique_farmer_names' in locals() else []
    )



@app.route("/view-purchase-bills", methods=["GET", "POST"])
def view_purchase_bills():
    if "user" not in session:
        return redirect("/")

    file_path = "purchase_bills.xlsx"

    # --- Handle deletion ---
    if request.method == "POST":
        selected_bills = request.form.getlist("delete_ids")
        if os.path.exists(file_path) and selected_bills:
            df = pd.read_excel(file_path, engine='openpyxl')
            df = df[~df['bill_no'].astype(str).isin(selected_bills)]
            df.to_excel(file_path, index=False, engine='openpyxl')
            flash(f"✅ {len(selected_bills)} Purchase Bill(s) deleted.", "success")
            return redirect("/view-purchase-bills")

    # --- Load bills ---
    bills = []
    if os.path.exists(file_path):
        with excel_lock:
            df = pd.read_excel(file_path, engine='openpyxl')
            df = df.drop_duplicates(subset='bill_no', keep='last')

            # --- FILTER LOGIC ---
            date = request.args.get("date")
            bill_no = request.args.get("bill_no")
            mill_name = request.args.get("mill_name")
            farmer_name = request.args.get("farmer_name")

            if date:
                df = df[df['date'] == date]
            if bill_no:
                df = df[df['bill_no'].astype(str) == bill_no]
            if mill_name:
                df = df[df['mill_name'] == mill_name]
            if farmer_name:
                df = df[df['farmer_name'] == farmer_name]

            # --- For dropdowns ---
            unique_bill_nos = df['bill_no'].dropna().astype(str).unique()
            unique_mill_names = df['mill_name'].dropna().unique()
            unique_farmer_names = df['farmer_name'].dropna().unique()

            bills = df.to_dict(orient='records')

    return render_template(
        "view_bills_purchase.html",
        bills=bills,
        unique_bill_nos=unique_bill_nos if 'unique_bill_nos' in locals() else [],
        unique_mill_names=unique_mill_names if 'unique_mill_names' in locals() else [],
        unique_farmer_names=unique_farmer_names if 'unique_farmer_names' in locals() else []
    )


# --- 10 Form (Word .docx) routes ---
@app.route("/10form", methods=["GET"])
def ten_form_page():
    if "user" not in session:
        flash("⚠️ Please log in first.", "warning")
        return redirect("/")

    full_path = os.path.join(DOCS_DIR, DOCX_FILENAME)
    if not os.path.exists(full_path):
        flash("❌ 10 Form file not found in static/docs.", "error")
        return redirect("/menu")

    docx_http_url = url_for("ten_form_download", _external=True)
    ms_word_url = f"ms-word:ofe|u|{docx_http_url}"

    return render_template("10form_choice.html", ms_word_url=ms_word_url)


@app.route("/10form/download", methods=["GET"])
def ten_form_download():
    if "user" not in session:
        flash("⚠️ Please log in first.", "warning")
        return redirect("/")
    return send_from_directory(
        DOCS_DIR,
        DOCX_FILENAME,
        as_attachment=False,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
# --- end 10 Form routes ---


@app.route("/download/<billtype>/<filetype>")
def download_file(billtype, filetype):
    if "user" not in session:
        flash("⚠️ Please log in first.", "warning")
        return redirect("/")

    if billtype.lower() == "sale":
        file_path = "sale_bills.xlsx"
    elif billtype.lower() == "purchase":
        file_path = "purchase_bills.xlsx"
    else:
        flash("❌ Invalid bill type.", "error")
        return redirect("/menu")

    if not os.path.exists(file_path):
        flash(f"❌ {file_path} not found.", "error")
        return redirect("/menu")

    df = pd.read_excel(file_path, engine='openpyxl')

    if filetype == "csv":
        path = f"{billtype}_bills.csv"
        df.to_csv(path, index=False)
        return send_file(path, as_attachment=True)
    elif filetype == "excel":
        path = f"{billtype}_bills_download.xlsx"
        df.to_excel(path, index=False, engine='openpyxl')
        return send_file(path, as_attachment=True)

    flash("❌ Invalid file type.", "error")
    return redirect("/menu")

@app.route("/clear-<billtype>-bills")
def clear_specific_bills(billtype):
    if "user" not in session:
        return redirect("/")
    try:
        file_path = "bills.xlsx"
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, engine='openpyxl')
            df = df[df['bill_type'].str.lower() != billtype.lower()]
            df.to_excel(file_path, index=False, engine='openpyxl')
        flash(f"✅ {billtype.title()} Bills cleared successfully.", "success")
    except Exception as e:
        flash(f"❌ Failed to clear bills: {str(e)}", "error")
    return redirect("/view-bills" if billtype == "sale" else "/view-purchase-bills")

@app.route("/download-excel")
def download_excel():
    if "user" not in session:
        flash("⚠️ Please log in first.", "warning")
        return redirect("/")

    file_path = "bills.xlsx"
    output_path = "sale_purchase_bills_only.xlsx"

    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine='openpyxl')

        sale_cols = [
            'bill_type', 'bill_no', 'date', 'mill_name', 'mill_code','farmer_name',
            'bags', 'ntwt', 'stwt', 'sut_rate', 'price', 'net_bags',
            'amount', 'commission', 'hamali', 'gunny', 'advance',
            'grand_total', 'lorry_no', 'mobile_no'
        ]

        purchase_cols = [
            'bill_type', 'bill_no', 'date', 'farmer_name', 'village_name','mill_name',
            'bags', 'ntwt', 'sut_rate', 'stwt', 'total_ntwt', 'rate',
            'amount', 'hamali', 'weigh_bridge', 'grand_total','lorry_no'
        ]

        sale_bills = df[df['bill_type'] == 'Sale'][sale_cols]
        purchase_bills = df[df['bill_type'] == 'Purchase'][purchase_cols]

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            if not sale_bills.empty:
                sale_bills.to_excel(writer, sheet_name='Sale Bills', index=False)
            if not purchase_bills.empty:
                purchase_bills.to_excel(writer, sheet_name='Purchase Bills', index=False)

        return send_file(output_path, as_attachment=True)

    flash("❌ bills.xlsx file not found.", "error")
    return redirect("/menu")


@app.route("/logout")
def logout():
    session.pop("user", None)
    flash("🔒 Logged out successfully.", "info")
    return redirect("/")

import os
import pandas as pd
from flask import Flask, render_template, request



SALE_FILE = "sale_bills.xlsx"
PURCHASE_FILE = "purchase_bills.xlsx"

# ----------------- Helpers -----------------
def safe_read_excel(path):
    if os.path.exists(path):
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def prep_df(df):
    """Standardize types and column presence so math never breaks."""
    if df.empty:
        return df

    # Normalize column names we rely on
    needed = [
        "date", "mill_name", "village_name", "farmer_name",
        "lorry_no", "bags", "ntwt", "net_bags", "amount"
    ]
    for col in needed:
        if col not in df.columns:
            df[col] = pd.NA

    # Parse date safely
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Coerce numerics
    for col in ["bags", "ntwt", "net_bags", "amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Clean strings
    for col in ["mill_name", "village_name", "farmer_name", "lorry_no"]:
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col].isin(["", "nan", "NaN", "None"]), col] = ""

    return df

def apply_common_filters(df, from_date, to_date, mill, village, farmer, lorry):
    if df.empty:
        return df
    if from_date:
        fdt = pd.to_datetime(from_date, errors="coerce")
        if pd.notnull(fdt):
            df = df[df["date"] >= fdt]
    if to_date:
        tdt = pd.to_datetime(to_date, errors="coerce")
        if pd.notnull(tdt):
            df = df[df["date"] <= tdt]
    if mill:
        df = df[df["mill_name"] == mill]
    if village:
        df = df[df["village_name"] == village]
    if farmer:
        df = df[df["farmer_name"] == farmer]
    if lorry:
        df = df[df["lorry_no"] == lorry]
    return df

def series_to_aligned_lists(s1, s2):
    """Union the indices of two Series and return aligned lists for charting."""
    idx = s1.index.union(s2.index)
    idx = sorted(idx)
    lab = [str(i) for i in idx]
    a = [float(s1.get(i, 0)) for i in idx]
    b = [float(s2.get(i, 0)) for i in idx]
    return lab, a, b

# ----------------- Route -----------------
@app.route("/analytics")
def analytics():
    # Load & prep
    sales_df = prep_df(safe_read_excel(SALE_FILE))
    purchase_df = prep_df(safe_read_excel(PURCHASE_FILE))

    # ---- Get filters ----
    from_date = request.args.get("from_date", "")
    to_date = request.args.get("to_date", "")
    mill_filter = request.args.get("mill", "")
    village_filter = request.args.get("village", "")
    farmer_filter = request.args.get("farmer", "")
    lorry_filter = request.args.get("lorry", "")

    # ---- Apply filters ----
    sales_df = apply_common_filters(sales_df, from_date, to_date, mill_filter, village_filter, farmer_filter, lorry_filter)
    purchase_df = apply_common_filters(purchase_df, from_date, to_date, mill_filter, village_filter, farmer_filter, lorry_filter)

    # ---- KPIs ----
    sales_kpi = {
        "total_bags": float(sales_df["bags"].sum()) if not sales_df.empty else 0,
        "total_ntwt": float(sales_df["ntwt"].sum()) if not sales_df.empty else 0,
        "total_net_bags": float(sales_df["net_bags"].sum()) if not sales_df.empty else 0,
        "total_count": int(len(sales_df)),
        "total_sales": float(sales_df["amount"].sum()) if not sales_df.empty else 0,
    }

    purchase_kpi = {
        "total_bags": float(purchase_df["bags"].sum()) if not purchase_df.empty else 0,
        "total_ntwt": float(purchase_df["ntwt"].sum()) if not purchase_df.empty else 0,
        "total_net_bags": float(purchase_df["net_bags"].sum()) if not purchase_df.empty else 0,
        "total_count": int(len(purchase_df)),
        "total_purchase": float(purchase_df["amount"].sum()) if not purchase_df.empty else 0,
    }

    # ---- Dropdowns (from both files) ----
    combined = pd.concat([sales_df, purchase_df], ignore_index=True) if not sales_df.empty or not purchase_df.empty else pd.DataFrame(columns=["mill_name","village_name","farmer_name","lorry_no"])
    mills = sorted([m for m in combined.get("mill_name", pd.Series()).dropna().unique().tolist() if m])
    villages = sorted([v for v in combined.get("village_name", pd.Series()).dropna().unique().tolist() if v])
    farmers = sorted([f for f in combined.get("farmer_name", pd.Series()).dropna().unique().tolist() if f])
    lorries = sorted([l for l in combined.get("lorry_no", pd.Series()).dropna().unique().tolist() if l])

    # ---- 1) Daily bar: totals per day ----
    if not sales_df.empty:
        sales_daily = sales_df.dropna(subset=["date"]).groupby(sales_df["date"].dt.date)["amount"].sum()
    else:
        sales_daily = pd.Series(dtype=float)
    if not purchase_df.empty:
        purchase_daily = purchase_df.dropna(subset=["date"]).groupby(purchase_df["date"].dt.date)["amount"].sum()
    else:
        purchase_daily = pd.Series(dtype=float)
    daily_labels, daily_sales, daily_purchase = series_to_aligned_lists(sales_daily, purchase_daily)

    # ---- 2) Weekly bar: totals per ISO week (use week start date) ----
    if not sales_df.empty:
        s_week = sales_df.dropna(subset=["date"]).groupby(sales_df["date"].dt.to_period("W").apply(lambda p: p.start_time.date()))["amount"].sum()
    else:
        s_week = pd.Series(dtype=float)
    if not purchase_df.empty:
        p_week = purchase_df.dropna(subset=["date"]).groupby(purchase_df["date"].dt.to_period("W").apply(lambda p: p.start_time.date()))["amount"].sum()
    else:
        p_week = pd.Series(dtype=float)
    weekly_labels, weekly_sales, weekly_purchase = series_to_aligned_lists(s_week, p_week)

    # ---- 3) Monthly trend: sums per month ----
    if not sales_df.empty:
        s_month = sales_df.dropna(subset=["date"]).groupby(sales_df["date"].dt.to_period("M"))["amount"].sum()
    else:
        s_month = pd.Series(dtype=float)
    if not purchase_df.empty:
        p_month = purchase_df.dropna(subset=["date"]).groupby(purchase_df["date"].dt.to_period("M"))["amount"].sum()
    else:
        p_month = pd.Series(dtype=float)
    # align months
    month_idx = s_month.index.union(p_month.index)
    month_idx = sorted(month_idx)
    trend_labels = [str(m) for m in month_idx]
    trend_sales = [float(s_month.get(m, 0)) for m in month_idx]
    trend_purchase = [float(p_month.get(m, 0)) for m in month_idx]

    # ---- 6) Monthly difference: (sales - purchase) per month ----
    monthly_diff = [float((s_month.get(m, 0) - p_month.get(m, 0))) for m in month_idx]

    # ---- 4) Donut: total bags (sales vs purchase) ----
    donut_bags = [sales_kpi["total_bags"], purchase_kpi["total_bags"]]

    # ---- 5) Pie: total amount (sales vs purchase) ----
    pie_amounts = [sales_kpi["total_sales"], purchase_kpi["total_purchase"]]

    # ---- 7) Top 5 farmers by purchase amount ----
    if not purchase_df.empty and "farmer_name" in purchase_df.columns:
        top_farmers_df = (purchase_df[purchase_df["farmer_name"] != ""]
                          .groupby("farmer_name")["amount"].sum()
                          .sort_values(ascending=False).head(5))
    else:
        top_farmers_df = pd.Series(dtype=float)
    top_farmers_labels = top_farmers_df.index.tolist()
    top_farmers_values = [float(v) for v in top_farmers_df.values.tolist()]

    # ---- 8) Top 5 mills by sales amount ----
    if not sales_df.empty and "mill_name" in sales_df.columns:
        top_mills_df = (sales_df[sales_df["mill_name"] != ""]
                        .groupby("mill_name")["amount"].sum()
                        .sort_values(ascending=False).head(5))
    else:
        top_mills_df = pd.Series(dtype=float)
    top_mills_labels = top_mills_df.index.tolist()
    top_mills_values = [float(v) for v in top_mills_df.values.tolist()]

    # ---- 9) Top 5 villages by total bags purchased ----
    if not purchase_df.empty and "village_name" in purchase_df.columns:
        top_villages_df = (purchase_df[purchase_df["village_name"] != ""]
                           .groupby("village_name")["bags"].sum()
                           .sort_values(ascending=False).head(5))
    else:
        top_villages_df = pd.Series(dtype=float)
    top_villages_labels = top_villages_df.index.tolist()
    top_villages_values = [float(v) for v in top_villages_df.values.tolist()]

    # ---- 10) Top 5 trucks by total bags (combined sales+purchase) ----
    bags_cols = ["lorry_no", "bags"]
    sp_bags = []
    if not sales_df.empty:
        sp_bags.append(sales_df[bags_cols].copy())
    if not purchase_df.empty:
        sp_bags.append(purchase_df[bags_cols].copy())
    if sp_bags:
        both_bags = pd.concat(sp_bags, ignore_index=True)
        top_trucks_df = (both_bags[both_bags["lorry_no"] != ""]
                         .groupby("lorry_no")["bags"].sum()
                         .sort_values(ascending=False).head(5))
    else:
        top_trucks_df = pd.Series(dtype=float)
    top_trucks_labels = top_trucks_df.index.tolist()
    top_trucks_values = [float(v) for v in top_trucks_df.values.tolist()]

    # ---- Records tables (unchanged) ----
    sales_records, purchase_records = [], []
    if not sales_df.empty:
        for _, row in sales_df.iterrows():
            sales_records.append({
                "date": row.get("date", ""),
                "mill": row.get("mill_name", ""),
                "farmer": row.get("farmer_name", ""),
                "lorry": row.get("lorry_no", ""),
                "bags": row.get("bags", 0),
                "ntwt": row.get("ntwt", 0),
                "amount": row.get("amount", 0)
            })
    if not purchase_df.empty:
        for _, row in purchase_df.iterrows():
            purchase_records.append({
                "date": row.get("date", ""),
                "mill": row.get("mill_name", ""),
                "farmer": row.get("farmer_name", ""),
                "village": row.get("village_name", ""),
                "lorry": row.get("lorry_no", ""),
                "bags": row.get("bags", 0),
                "ntwt": row.get("ntwt", 0),
                "amount": row.get("amount", 0)
            })

    return render_template(
        "analytics.html",
        # Filters
        mills=mills, villages=villages, farmers=farmers, lorries=lorries,
        # KPIs
        sales=sales_kpi, purchase=purchase_kpi,
        # 1) Daily
        daily_labels=daily_labels, daily_sales=daily_sales, daily_purchase=daily_purchase,
        # 2) Weekly
        weekly_labels=weekly_labels, weekly_sales=weekly_sales, weekly_purchase=weekly_purchase,
        # 3) Monthly trend
        trend_labels=trend_labels, trend_sales=trend_sales, trend_purchase=trend_purchase,
        # 6) Monthly difference
        diff_labels=trend_labels, monthly_diff=monthly_diff,
        # 4) Donut bags
        donut_bags=donut_bags,
        # 5) Pie amounts
        pie_amounts=pie_amounts,
        # 7) Top farmers
        top_farmers_labels=top_farmers_labels, top_farmers_values=top_farmers_values,
        # 8) Top mills
        top_mills_labels=top_mills_labels, top_mills_values=top_mills_values,
        # 9) Top villages
        top_villages_labels=top_villages_labels, top_villages_values=top_villages_values,
        # 10) Top trucks
        top_trucks_labels=top_trucks_labels, top_trucks_values=top_trucks_values,
        # Tables
        sales_records=sales_records, purchase_records=purchase_records
    )




# run if this file is executed directly (for testing)
if __name__ == '__main__':
    app.run(debug=True)



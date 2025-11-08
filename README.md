🧾 SRI ANJANEYA TRADERS Billing System

A complete Flask-based billing web application designed for SRI ANJANEYA TRADERS, automating billing operations for Sale, Purchase, Transport, and IO Forms — now enhanced with Business Analytics & Insights Dashboard for tracking performance and profitability.

🚀 Features
🧾 Billing Modules

Sale Bill

Auto-generated Bill No. (DDMMYYNNN)

Auto-filled date, item, and calculation fields

Handles Commission, Hamali, Gunny Bags, Advance, and Lorry details

PDF generation and WhatsApp sharing

Option to save draft entries

Mobile-friendly responsive layout

Purchase Bill

Formula-based stwt = bags × sut_rate

Auto-calculates STWT, Total NTWT, Amount, and Grand Total

Excel export, view, and delete options

Separate view page for all Purchase Bills

Transport Bill

Includes broker cash, lorry charges, hamali, commission, and delivery info

Auto-calculated totals with clean printable PDF

IO Form (Form X / Way Bill)

Matches government form layout with boxes, borders, and exact formatting

Printable and downloadable as PDF

📊 Business Analytics Module

The Analytics Dashboard gives clear insights into your business performance using interactive graphs and tables.

📈 Features:

Total Sales & Purchases Overview (Monthly / Yearly)

Profit & Expense Analysis

Tracks commissions, hamali, and additional charges

Mill-wise Summary Reports

Shows top-performing mills and purchase trends

Dynamic Charts

Visualize trends using bar and line charts (powered by Chart.js / Recharts)

Date Range Filter

Select and analyze specific periods

Export Reports to Excel / PDF

Analytics automatically fetches data from sale_bills.xlsx and purchase_bills.xlsx.

💾 Data Management
File	Purpose
sale_bills.xlsx	Stores all Sale Bills
purchase_bills.xlsx	Stores all Purchase Bills
analytics_cache.xlsx (optional)	Used for caching summary reports

Prevents duplicates using unique Bill No.

View, Download (Excel), or Delete selected bills directly from web UI

🧮 Calculation Logic
Sale Bill:
Option 1: Net Bags = ntwt / 77
Option 2: Net Bags = (ntwt - stwt) / 75
stwt = bags × sut_value
Amount = Net Bags × Price
Commission = Amount / 100 (if applicable)
Hamali = Bags × rate (if applicable)
Gunny Bags = Bags × rate (if applicable)
Grand Total = Amount + Commission + Hamali + Gunny Bags + Advance

⚙️ Technologies Used
Category	Technology
Backend	Python, Flask
Frontend	HTML, CSS, JavaScript
Database	Excel (via Pandas)
PDF Engine	xhtml2pdf
Charts	Chart.js / Recharts
Authentication	Flask Flash Messages
File Handling	Pandas, OS Module
📂 Project Structure
SRI_ANJANEYA_TRADERS/
│
├── app.py
├── analytics.py                  # Analytics logic & summary calculations
│
├── templates/
│   ├── welcome.html
│   ├── menu.html
│   ├── sale_bill.html
│   ├── sale_bill_template.html
│   ├── purchase_bill.html
│   ├── purchase_bill_template.html
│   ├── transport_bill.html
│   ├── transportation_bill_template.html
│   ├── 10form.html
│   ├── 10form_template.html
│   ├── view_sale_bills.html
│   ├── view_purchase_bills.html
│   ├── analytics.html             # New Analytics Dashboard
│   └── bill_template.html
│
├── static/                        # CSS, JS, Images
├── sale_bills.xlsx
├── purchase_bills.xlsx
├── analytics_cache.xlsx
└── README.md

💻 Setup Instructions
1️⃣ Clone the Repository
git clone https://github.com/yourusername/sri-anjaneya-traders.git
cd sri-anjaneya-traders

2️⃣ Install Dependencies
pip install flask pandas xhtml2pdf openpyxl num2words

3️⃣ Run the Application
python app.py

4️⃣ Open in Browser
http://127.0.0.1:5000/

📱 Highlights

✅ Auto-generated Bill Numbers
✅ Easy navigation through menu page
✅ Export to Excel and PDF
✅ Delete selected bills using checkboxes
✅ Toast popups for success messages
✅ Real-time analytics dashboard

🧩 Future Enhancements

Multi-user Login System

Cloud Data Storage (MySQL / Firebase)

Auto WhatsApp PDF Sending

Voice-based Bill Entry using Speech-to-Text

Integration with Mobile App

👨‍💼 Developer Information

Developer: Mohan M
Role: Paddy Commission Agent
Organization: SRI ANJANEYA TRADERS
Location: Yadgir, Karnataka, India
Tech Stack: Flask • Pandas • HTML • JS • xhtml2pdf • Chart.js

import subprocess
import webbrowser
import time
import os

# Go to project folder
os.chdir(r"C:\Users\mohan\OneDrive\Desktop\sri_anjaneya_traders")

# Start Flask app silently using pythonw
subprocess.Popen([r"venv\Scripts\pythonw.exe", "app.py"])

# Wait for Flask to start
time.sleep(3)

# Open default browser to the local website
webbrowser.open("http://127.0.0.1:5000")

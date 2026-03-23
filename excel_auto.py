# === IMPORT LIBRARY ===
import win32com.client as win32
import os
import smtplib
from email.message import EmailMessage
import ssl
from datetime import datetime

# === CONFIG ===
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # True kalau mau lihat prosesnya
excel.DisplayAlerts = False

file_path = r"C:\Document\pythonCode\Automation-Email-Excel\excel_auto.py"
output_folder = r"C:\Document\pythonCode\Automation-Email-Excel"

# Format tanggal untuk nama file
today = datetime.now().strftime("%d-%m-%Y")  

# === EXPORT EXCEL KE PDF ===
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Dashboard")  # ganti sesuai nama sheet


output_file = os.path.join(output_folder, f"DailyReport_{today}.pdf")

ws.ExportAsFixedFormat(0, output_file)

# Hapus filter
ws.AutoFilterMode = False

wb.Close(False)
excel.Quit()
print("Export PDF berhasil")


# === CONFIG EMAIL ===
sender_email = "data@yamahabismagroup.com"
password = "databisma1*"
receiver_email = ["desakiintan25@gmail.com", "praandikayoga@gmail.com"]

subject = "Daily Report"
body = f"""
Salam Semakin Didepan,

Berikut kami sampaikan laporan harian per tanggal {today}.
Silakan cek file terlampir.

Note : email ini dikirim secara otomatis, mohon untuk tidak membalas email ini.

---------------------------------------------------
Best Regards.
Divisi Data -- CRM Bisma -- Group     
Jl. Teuku Umar Barat No 100X Malboro Denpasar
"""

# === BUAT EMAIL ===
msg = EmailMessage()
msg["From"] = sender_email
msg["To"] = ", ".join(receiver_email)
msg["Subject"] = subject
msg.set_content(body)

# Attach PDF
with open(output_file, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="pdf",
        filename=f"DailyReport_{today}.pdf"
    )

# === KIRIM EMAIL (SSL) ===
context = ssl.create_default_context()

with smtplib.SMTP_SSL("mail.yamahabismagroup.com", 465, context=context) as server:
    server.login(sender_email, password)
    server.send_message(msg)

print("Email berhasil dikirim via SMTP")
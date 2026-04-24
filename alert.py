import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- SETTINGS ---
FILE_NAME = "starhealth.xlsx"

SENDER_EMAIL = "aditya.padale@modgensolutions.com"
SENDER_PASSWORD = "nbxscnmfzkjxhqvk"
RECEIVER_EMAIL = "aditya.padale@modgensolutions.com"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# --- EMAIL FUNCTION ---
def send_summary_email(expiring_policies, expired_policies):

    subject = "🩺 StarHealth Policy Expiry Alert"

    html = f"""
    <html>
    <body style="font-family: Arial;">

    <h2>⚠️ Policies Expiring Within 30 Days</h2>

    <table border="1" cellspacing="0" cellpadding="6">
        <tr style="background-color:#f2f2f2;">
            <th>Customer Name</th>
            <th>Coverage</th>
            <th>Expiry Date</th>
            <th>Days Left</th>
        </tr>
        {''.join(expiring_policies) if expiring_policies else '<tr><td colspan="4">No policies</td></tr>'}
    </table>

    <br><br>

    <h2>❌ Already Expired Policies</h2>

    <table border="1" cellspacing="0" cellpadding="6">
        <tr style="background-color:#f2f2f2;">
            <th>Customer Name</th>
            <th>Coverage</th>
            <th>Expiry Date</th>
            <th>Status</th>
        </tr>
        {''.join(expired_policies) if expired_policies else '<tr><td colspan="4">No expired policies</td></tr>'}
    </table>

    <br>
    <p>Regards,<br>Policy Alert System</p>

    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL

    msg.attach(MIMEText(html, "html"))

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()

        print("📧 Email sent successfully")

    except Exception as e:
        print(f"❌ Email failed: {e}")

# --- MAIN ---
print("--- 🚀 Policy Alert Script Started ---")

df = pd.read_excel(FILE_NAME)

today = datetime.now().date()

expiring_list = []
expired_list = []

for _, row in df.iterrows():

    name = row['Customer Name']
    coverage = row['Coverage Amount']
    expiry_date = pd.to_datetime(row['Policy Expiry Date']).date()

    days_left = (expiry_date - today).days

    # Ignore garbage
    if days_left > 3650 or days_left < -3650:
        continue

    # --- EXPIRING ---
    if 0 <= days_left <= 30:

        # 10–30 → every 5 days
        if days_left > 10:
            if days_left % 5 != 0:
                continue

        # Color logic
        if days_left <= 3:
            color = "red"
        elif days_left <= 10:
            color = "orange"
        else:
            color = "goldenrod"

        expiring_list.append((days_left, f"""
        <tr>
            <td>{name}</td>
            <td>{coverage}</td>
            <td>{expiry_date}</td>
            <td style="color:{color}; font-weight:bold;">{days_left} days</td>
        </tr>
        """))

    # --- EXPIRED ---
    elif days_left < 0:
        expired_list.append((days_left, f"""
        <tr>
            <td>{name}</td>
            <td>{coverage}</td>
            <td>{expiry_date}</td>
            <td style="color:red;">Expired {abs(days_left)} days ago</td>
        </tr>
        """))

# --- SORT ---
expiring_list.sort(key=lambda x: x[0])
expired_list.sort(key=lambda x: x[0])

expiring_policies = [x[1] for x in expiring_list]
expired_policies = [x[1] for x in expired_list]

# --- FINAL CONDITION ---
if len(expiring_policies) > 0:
    send_summary_email(expiring_policies, expired_policies)
else:
    print("✅ No alerts to send")

print("--- ✅ Process Completed ---")

import requests
import base64
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

# --- KONFIGURATION (Wird über GitHub Secrets geladen) ---
EPO_KEY = os.getenv("EPO_CONSUMER_KEY")
EPO_SECRET = os.getenv("EPO_CONSUMER_SECRET")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") # App-Passwort!
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
SMTP_SERVER = "smtp.gmail.com" # Oder dein Anbieter
SMTP_PORT = 587

def get_last_wednesday():
    today = datetime.now()
    offset = (today.weekday() - 2) % 7
    return (today - timedelta(days=offset)).strftime("%Y%m%d")

def get_token():
    url = "https://ops.epo.org/3.2/auth/accesstoken"
    auth_string = f"{EPO_KEY}:{EPO_SECRET}"
    encoded = base64.b64encode(auth_string.encode()).decode()
    headers = {"Authorization": f"Basic {encoded}", "Content-Type": "application/x-www-form-urlencoded"}
    res = requests.post(url, headers=headers, data={"grant_type": "client_credentials"})
    return res.json().get("access_token")

def fetch_data(token, date_str):
    query = f"pd={date_str} and ic=B23"
    url = "https://ops.epo.org/3.2/rest-services/published-data/search/biblio"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/xml", "X-OPS-Range": "1-100"}
    res = requests.get(url, headers=headers, params={'q': query})
    return res.text if res.status_code == 200 else None

def parse_xml(xml_data):
    if not xml_data: return pd.DataFrame()
    ns = {'ops': 'http://ops.epo.org', 'exchange': 'http://www.epo.org/exchange'}
    root = ET.fromstring(xml_data)
    results = []
    for biblio in root.findall(".//exchange:bibliographic-data", ns):
        pub_ref = biblio.find(".//exchange:publication-reference", ns)
        country = pub_ref.find(".//exchange:country", ns).text if pub_ref is not None else "EP"
        doc_num = pub_ref.find(".//exchange:doc-number", ns).text if pub_ref is not None else ""
        kind = pub_ref.find(".//exchange:kind", ns).text if pub_ref is not None else ""
        
        title = "N/A"
        titles = biblio.findall(".//exchange:invention-title", ns)
        for t in titles:
            title = t.text
            if t.get('lang') == 'de': break
            
        results.append({
            "Land": country,
            "Patentnummer": doc_num,
            "Dokumentart": kind,
            "Titel": title,
            "Link zu ESPACENET": f"https://worldwide.espacenet.com/patent/search?q={country}{doc_num}"
        })
    return pd.DataFrame(results)

def send_mail(df, date_str):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = f"Wöchentliches Patent-Update B23 ({date_str})"
    
    body = f"Anbei findest du die neuen Patentveröffentlichungen der Klasse B23 vom {date_str}."
    msg.attach(MIMEText(body, 'plain'))
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(output.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename=B23_Update_{date_str}.xlsx")
    msg.attach(part)
    
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)

if __name__ == "__main__":
    date_str = get_last_wednesday()
    print(f"Starte Abfrage für {date_str}...")
    token = get_token()
    if token:
        xml = fetch_data(token, date_str)
        df = parse_xml(xml)
        if not df.empty:
            send_mail(df, date_str)
            print("E-Mail erfolgreich versendet.")
        else:
            print("Keine neuen Patente gefunden.")

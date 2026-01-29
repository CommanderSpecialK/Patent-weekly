import requests
import base64
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import pandas as pd
import smtplib
import os
import sys
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

# --- KONFIGURATION ---
# Wir rufen die Variablen ab und prüfen direkt, ob sie existieren
EPO_KEY = os.getenv("EPO_CONSUMER_KEY")
EPO_SECRET = os.getenv("EPO_CONSUMER_SECRET")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

def get_last_wednesday():
    today = datetime.now()
    offset = (today.weekday() - 2) % 7
    return (today - timedelta(days=offset)).strftime("%Y%m%d")

def get_token():
    url = "https://ops.epo.org/3.2/auth/accesstoken"
    
    # Detaillierte Prüfung der Secrets
    missing_secrets = []
    if not EPO_KEY: missing_secrets.append("EPO_CONSUMER_KEY")
    if not EPO_SECRET: missing_secrets.append("EPO_CONSUMER_SECRET")
    
    if missing_secrets:
        print(f"FEHLER: Folgende Secrets fehlen in der Umgebung: {', '.join(missing_secrets)}")
        print("Hinweis: Prüfe deine GitHub Repository Settings -> Secrets -> Actions.")
        return None
    
    auth_string = f"{EPO_KEY}:{EPO_SECRET}"
    encoded = base64.b64encode(auth_string.encode()).decode()
    headers = {
        "Authorization": f"Basic {encoded}", 
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    try:
        print("Sende Token-Anfrage an EPO...")
        res = requests.post(url, headers=headers, data={"grant_type": "client_credentials"}, timeout=20)
        
        if res.status_code == 200:
            try:
                token_data = res.json()
                return token_data.get("access_token")
            except json.JSONDecodeError:
                print(f"FEHLER: Antwort von EPO ist kein gültiges JSON. Antwortinhalt: {res.text[:200]}")
                return None
        else:
            print(f"FEHLER Authentifizierung: Status {res.status_code}")
            print(f"Antwort vom Server: {res.text}")
            return None
    except Exception as e:
        print(f"EXCEPTION bei Token-Abfrage: {e}")
        return None

def fetch_data(token, date_str):
    query = f"pd={date_str} and ic=B23"
    url = "https://ops.epo.org/3.2/rest-services/published-data/search/biblio"
    headers = {
        "Authorization": f"Bearer {token}", 
        "Accept": "application/xml", 
        "X-OPS-Range": "1-100"
    }
    
    try:
        print(f"Suche Patente für {date_str}...")
        res = requests.get(url, headers=headers, params={'q': query}, timeout=30)
        if res.status_code == 200:
            return res.text
        else:
            print(f"FEHLER API Abfrage: Status {res.status_code}, Antwort: {res.text}")
            return None
    except Exception as e:
        print(f"EXCEPTION bei Daten-Abfrage: {e}")
        return None

def parse_xml(xml_data):
    if not xml_data: return pd.DataFrame()
    ns = {'ops': 'http://ops.epo.org', 'exchange': 'http://www.epo.org/exchange'}
    try:
        root = ET.fromstring(xml_data)
        results = []
        for biblio in root.findall(".//exchange:bibliographic-data", ns):
            pub_ref = biblio.find(".//exchange:publication-reference", ns)
            
            country = "EP"
            doc_num = "N/A"
            kind = ""
            
            if pub_ref is not None:
                c_node = pub_ref.find(".//exchange:country", ns)
                n_node = pub_ref.find(".//exchange:doc-number", ns)
                k_node = pub_ref.find(".//exchange:kind", ns)
                if c_node is not None: country = c_node.text
                if n_node is not None: doc_num = n_node.text
                if k_node is not None: kind = k_node.text
            
            title = "N/A"
            titles = biblio.findall(".//exchange:invention-title", ns)
            for t in titles:
                if t.text:
                    title = t.text
                    if t.get('lang') == 'de': break
                
            results.append({
                "Land": country,
                "Patentnummer": doc_num,
                "Dokumentart": kind,
                "Titel": title,
                "Link zu ESPACENET": f"https://worldwide.espacenet.com/patent/search?q={country}{doc_num}{kind}"
            })
        return pd.DataFrame(results)
    except Exception as e:
        print(f"FEHLER beim XML-Parsing: {e}")
        return pd.DataFrame()

def send_mail(df, date_str):
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        print("FEHLER: E-Mail Zugangsdaten (SENDER oder PASSWORD) fehlen.")
        return
        
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = f"Wöchentliches Patent-Update B23 ({date_str})"
    
    body = f"Anbei findest du die neuen Patentveröffentlichungen der Klasse B23 vom {date_str}.\nAnzahl der Treffer: {len(df)}"
    msg.attach(MIMEText(body, 'plain'))
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(output.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename=B23_Update_{date_str}.xlsx")
    msg.attach(part)
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        print("E-Mail erfolgreich versendet.")
    except Exception as e:
        print(f"FEHLER beim E-Mail Versand: {e}")
        sys.exit(1)

if __name__ == "__main__":
    date_str = get_last_wednesday()
    print(f"--- Prozess gestartet für Publikationsdatum: {date_str} ---")
    
    token = get_token()
    if not token:
        print("Abbruch: Authentifizierung beim EPO fehlgeschlagen.")
        sys.exit(1)
        
    xml = fetch_data(token, date_str)
    if not xml:
        print("Abbruch: Keine Daten von der API erhalten.")
        sys.exit(1)
        
    df = parse_xml(xml)
    if not df.empty:
        print(f"{len(df)} Patente gefunden. Bereite Versand vor...")
        send_mail(df, date_str)
    else:
        print("Keine neuen Patente für diesen Zeitraum in Klasse B23 gefunden.")

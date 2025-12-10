import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import json  # <--- ახალი ბიბლიოთეკა

# --- კონფიგურაცია ---
st.set_page_config(page_title="NeuroCRM", page_icon="🧠", layout="wide")

# პაროლები და უსაფრთხოება
ADMIN_PASSWORD = "Neuro2025"
SENDER_EMAIL = "your_email@gmail.com" # <--- ჩაწერე შენი მეილი!
SENDER_PASSWORD = "xxxx xxxx xxxx xxxx" # <--- ჩაწერე შენი App Password!

# --- DATABASE CONNECTION (ROBUST CLOUD FIX) ---
@st.cache_resource
def connect_db():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive",
        ]

        # 🎯 იძულებით ვხმარობთ creds.json-ს
        creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open("NeuroCRM_DB")
        return spreadsheet

    except Exception as e:
        st.error(f"კავშირის კრიტიკული შეცდომა: {e}")
        return None

# --- EMAIL SYSTEM (UTF-8 FIX INCLUDED) ---
def send_confirmation_email(to_email, patient_name, slot_time):
    if "your_email" in SENDER_EMAIL: return False

    subject_text = "ჯავშნის დადასტურება - ფსიქოთერაპია"
    body = f"""
    გამარჯობა {patient_name},
    
    თქვენი ვიზიტი წარმატებით დაიჯავშნა.
    დრო: {slot_time}
    
    მისამართი: თბილისი...
    
    პატივისცემით,
    NeuroCRM System.
    """

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = to_email
    msg['Subject'] = Header(subject_text, 'utf-8') 
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"მეილის ერორი: {e}")
        return False

# --- HELPER FUNCTIONS ---
def get_free_slots(spreadsheet):
    try:
        ws = spreadsheet.worksheet("Slots")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if df.empty: return []
        df['Display'] = df['Date'] + " | " + df['Time']
        return sorted(df[df['Status'] == 'Open']['Display'].tolist())
    except: return []

def mark_slot_booked(spreadsheet, slot_display):
    try:
        ws = spreadsheet.worksheet("Slots")
        date_part, time_part = slot_display.split(" | ")
        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if len(row) >= 3 and row[0] == date_part and row[1] == time_part and row[2] == 'Open':
                ws.update_cell(i+1, 3, 'Booked') 
                return True
        return False
    except: return False

def add_new_slot(spreadsheet, date, time):
    try:
        ws = spreadsheet.worksheet("Slots")
        ws.append_row([str(date), str(time), "Open"])
        return True
    except: return False

def delete_slot(spreadsheet, date, time):
    try:
        ws = spreadsheet.worksheet("Slots")
        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if row[0] == str(date) and row[1] == str(time):
                ws.delete_rows(i+1)
                return True
        return False
    except: return False

def get_all_slots(spreadsheet):
    try:
        ws = spreadsheet.worksheet("Slots")
        return pd.DataFrame(ws.get_all_records())
    except: return pd.DataFrame()

def generate_time_options():
    times = []
    for h in range(9, 22):
        times.append(f"{h:02d}:00")
        times.append(f"{h:02d}:30")
    return times

# --- MAIN APP ---
def main():
    st.title("🧠 NeuroCRM v2.3 (Stable)")
    
    db = connect_db()
    if not db: 
        st.warning("მონაცემთა ბაზა მიუწვდომელია. შეამოწმეთ Secrets.")
        return

    st.sidebar.title("ნავიგაცია")
    menu = st.sidebar.radio("მენიუ:", ["პაციენტის რეგისტრაცია", "ადმინ პანელი"])

    if menu == "პაციენტის რეგისტრაცია":
        st.header("ვიზიტის დაჯავშნა")
        free_slots = get_free_slots(db)
        
        with st.form("booking"):
            c1, c2 = st.columns(2)
            name = c1.text_input("სახელი და გვარი")
            phone = c1.text_input("ტელეფონი")
            email = c2.text_input("ელ-ფოსტა")
            slot = c2.selectbox("დრო", free_slots if free_slots else ["ადგილები არ არის"])
            
            if st.form_submit_button("დაჯავშნა"):
                if name and phone and email and slot != "ადგილები არ არის":
                    if mark_slot_booked(db, slot):
                        ws = db.worksheet("Sheet1")
                        ws.append_row([str(datetime.now().timestamp()), name, phone, email, "Active", "", str(datetime.now().date()), slot])
                        send_confirmation_email(email, name, slot)
                        st.success("✅ დაჯავშნილია!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("❌ უკვე დაკავებულია.")

    elif menu == "ადმინ პანელი":
        pwd = st.sidebar.text_input("პაროლი", type="password")
        if pwd == ADMIN_PASSWORD:
            t1, t2, t3 = st.tabs(["დამატება", "წაშლა", "ბაზა"])
            
            with t1:
                d = st.date_input("თარიღი")
                t = st.selectbox("საათი", generate_time_options())
                if st.button("დამატება"):
                    if add_new_slot(db, d, t): st.success("OK"); st.cache_data.clear()
            
            with t2:
                df = get_all_slots(db)
                if not df.empty:
                    df['S'] = df['Date'] + " | " + df['Time'] + " (" + df['Status'] + ")"
                    to_del = st.selectbox("აირჩიე", df['S'])
                    if st.button("წაშლა"):
                        raw = to_del.split(" (")[0]
                        dd, tt = raw.split(" | ")
                        if delete_slot(db, dd, tt): st.success("Deleted"); st.cache_data.clear(); st.rerun()
            
            with t3:
                ws = db.worksheet("Sheet1")
                st.dataframe(pd.DataFrame(ws.get_all_records()))

if __name__ == "__main__":
    main()

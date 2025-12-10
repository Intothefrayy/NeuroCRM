import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header # <--- ახალი ბიბლიოთეკა სათაურის ფიქსისთვის

# --- კონფიგურაცია ---
st.set_page_config(page_title="NeuroCRM", page_icon="🧠", layout="wide")

# 🔒 უსაფრთხოება (შეავსე შენი მონაცემებით)
ADMIN_PASSWORD = "Neuro2025"
SENDER_EMAIL = "fkhomasuridze@gmail.com"  # <--- ჩაწერე მეილი
SENDER_PASSWORD = "dadn fycl vfvo ahmr" # <--- ჩაწერე პაროლი

# Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = 'creds.json'
SHEET_NAME = 'NeuroCRM_DB'

# --- BACKEND LOGIC ---

@st.cache_resource
def connect_db():
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        return spreadsheet
    except Exception as e:
        st.error(f"კავშირის შეცდომა: {e}")
        return None

def send_confirmation_email(to_email, patient_name, slot_time):
    """აგზავნის დასტურის მეილს (Full UTF-8 Fix)"""
    if "your_email" in SENDER_EMAIL:
        return False

    # სათაურის ტექსტი (სუფთა UTF-8)
    subject_text = "ჯავშნის დადასტურება - ფსიქოთერაპია"
    
    body = f"""
    გამარჯობა {patient_name},
    
    თქვენი ვიზიტი წარმატებით დაიჯავშნა.
    დრო: {slot_time}
    
    მისამართი: (შენი მისამართი)
    
    პატივისცემით,
    თქვენი თერაპევტი.
    """

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = to_email
    # ვახდენთ სათაურის იძულებით კოდირებას, რომ \xa0 სიმბოლომ არ გაჭედოს
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
        st.error(f"მეილის გაგზავნის შეცდომა: {e}")
        return False

# --- SLOT MANAGEMENT ---

def get_all_slots(spreadsheet):
    """აბრუნებს ყველა სლოტს (ადმინისთვის)"""
    try:
        ws = spreadsheet.worksheet("Slots")
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except:
        return pd.DataFrame()

def get_free_slots(spreadsheet):
    """აბრუნებს მხოლოდ თავისუფალ დროებს (პაციენტისთვის)"""
    try:
        ws = spreadsheet.worksheet("Slots")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if df.empty: return []
        
        df['Display'] = df['Date'] + " | " + df['Time']
        free_slots = df[df['Status'] == 'Open']['Display'].tolist()
        return sorted(free_slots)
    except:
        return []

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
    except Exception as e:
        st.error(f"შეცდომა: {e}")
        return False

def add_new_slot(spreadsheet, date, time):
    try:
        ws = spreadsheet.worksheet("Slots")
        # ვამოწმებთ უკვე ხომ არ არის ასეთი დრო
        cell = ws.find(str(time)) # მარტივი შემოწმება (გასაუმჯობესებელია)
        ws.append_row([str(date), str(time), "Open"])
        return True
    except Exception as e:
        st.error(f"დამატების შეცდომა: {e}")
        return False

def delete_slot(spreadsheet, date, time):
    """შლის სლოტს ბაზიდან (Delete Row)"""
    try:
        ws = spreadsheet.worksheet("Slots")
        all_values = ws.get_all_values()
        
        # ვეძებთ შესაბამის რიგს
        for i, row in enumerate(all_values):
            # row[0] არის Date, row[1] არის Time. i არის ინდექსი (0-დან იწყება)
            if row[0] == str(date) and row[1] == str(time):
                # gspread-ში რიგები იწყება 1-დან, ამიტომ i+1
                ws.delete_rows(i+1)
                return True
        return False
    except Exception as e:
        st.error(f"წაშლის შეცდომა: {e}")
        return False

def generate_time_options():
    times = []
    for h in range(9, 22):
        times.append(f"{h:02d}:00")
        times.append(f"{h:02d}:30")
    return times

# --- MAIN APP ---

def main():
    st.title("🧠 NeuroCRM v2.2")
    
    db = connect_db()
    if not db: return

    st.sidebar.title("ნავიგაცია")
    menu = st.sidebar.radio("აირჩიეთ განყოფილება", ["პაციენტის რეგისტრაცია", "ადმინ პანელი"])

    # --- 1. პაციენტის რეგისტრაცია ---
    if menu == "პაციენტის რეგისტრაცია":
        st.header("ვიზიტის დაჯავშნა")
        free_slots = get_free_slots(db)
        
        with st.form("booking_form"):
            col1, col2 = st.columns(2)
            with col1:
                name = st.text_input("სახელი და გვარი")
                phone = st.text_input("ტელეფონი")
            with col2:
                email = st.text_input("ელ-ფოსტა (დასტურის მისაღებად)")
                if free_slots:
                    chosen_slot = st.selectbox("აირჩიეთ თავისუფალი დრო", free_slots)
                else:
                    chosen_slot = st.selectbox("დრო", ["ადგილები არ არის"])
                notes = st.text_area("დამატებითი ინფორმაცია")

            submit = st.form_submit_button("დაჯავშნა")

            if submit:
                if name and phone and email and chosen_slot != "ადგილები არ არის":
                    slot_booked = mark_slot_booked(db, chosen_slot)
                    
                    if slot_booked:
                        ws_patients = db.worksheet("Sheet1")
                        new_id = int(datetime.now().timestamp())
                        date_booked = datetime.now().strftime("%Y-%m-%d")
                        ws_patients.append_row([new_id, name, phone, email, "Active", notes, date_booked, chosen_slot])
                        
                        with st.spinner('იგზავნება დასტური...'):
                            email_sent = send_confirmation_email(email, name, chosen_slot)
                        
                        st.success(f"✅ ჯავშანი დადასტურებულია: {chosen_slot}")
                        if email_sent:
                            st.toast("📧 მეილი გაიგზავნა წარმატებით!")
                        
                        st.cache_data.clear()
                    else:
                        st.error("❌ ეს დრო უკვე დაკავებულია.")
                else:
                    st.warning("გთხოვთ შეავსოთ სავალდებულო ველები (*)")

    # --- 2. ადმინ პანელი ---
    elif menu == "ადმინ პანელი":
        st.sidebar.markdown("---")
        password_input = st.sidebar.text_input("შეიყვანეთ ადმინის პაროლი", type="password")

        if password_input == ADMIN_PASSWORD:
            st.header("🛠 ადმინისტრირება")
            
            tab1, tab2, tab3 = st.tabs(["სლოტის დამატება", "სლოტის წაშლა", "ბაზის ნახვა"])
            
            # TAB 1: დამატება
            with tab1:
                st.subheader("ახალი დროის დამატება")
                c1, c2 = st.columns(2)
                d_in = c1.date_input("თარიღი", min_value=datetime.today(), key="add_date")
                t_in = c2.selectbox("საათი", generate_time_options(), key="add_time")
                
                if st.button("➕ დამატება"):
                    if add_new_slot(db, d_in, t_in):
                        st.success(f"დაემატა: {d_in} | {t_in}")
                        st.cache_data.clear()

            # TAB 2: წაშლა (ახალი ფუნქცია)
            with tab2:
                st.subheader("არსებული სლოტის წაშლა")
                st.warning("ფრთხილად: წაშლა შეუქცევადია!")
                
                # ვიღებთ ყველა სლოტს (დაკავებულსაც და თავისუფალსაც)
                df_slots = get_all_slots(db)
                
                if not df_slots.empty:
                    # ვაკეთებთ სიას არჩევისთვის
                    df_slots['Select'] = df_slots['Date'] + " | " + df_slots['Time'] + " (" + df_slots['Status'] + ")"
                    slot_to_delete = st.selectbox("აირჩიე სლოტი გასაუქმებლად", df_slots['Select'].tolist())
                    
                    if st.button("🗑️ არჩეული სლოტის წაშლა"):
                        # ვშლით ტექსტს რომ მივიღოთ სუფთა თარიღი და დრო
                        raw_info = slot_to_delete.split(" (")[0] # ვაშორებთ სტატუსს
                        del_date, del_time = raw_info.split(" | ")
                        
                        if delete_slot(db, del_date, del_time):
                            st.success(f"სლოტი {del_date} {del_time} წაიშალა!")
                            st.cache_data.clear()
                            st.rerun() # გვერდის გადატვირთვა რომ განახლდეს სია
                        else:
                            st.error("წაშლა ვერ მოხერხდა.")
                else:
                    st.info("სლოტები ცარიელია")

            # TAB 3: ბაზა
            with tab3:
                st.subheader("მონაცემები")
                if st.button("განახლება"): st.cache_data.clear()
                ws_patients = db.worksheet("Sheet1")
                st.dataframe(pd.DataFrame(ws_patients.get_all_records()), use_container_width=True)

        elif password_input:
            st.error("პაროლი არასწორია!")

if __name__ == "__main__":
    main()

import pygsheets
import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import json
import tempfile
import os

# --- CONFIGURATION FROM SECRETS ---
try:
    SENDER_EMAIL = st.secrets.get("email", {}).get("sender_email", "")
    SENDER_PASSWORD = st.secrets.get("email", {}).get("sender_password", "")
    ADMIN_PASSWORD = st.secrets.get("admin_password", "admin123")
except Exception as e:
    SENDER_EMAIL = ""
    SENDER_PASSWORD = ""
    ADMIN_PASSWORD = "admin123"


@st.cache_resource
def connect_db():
    try:
        # ვიღებთ service account-ის ინფორმაციას Streamlit Secrets-დან
        if "gcp_service_account" not in st.secrets:
            raise Exception("gcp_service_account არ არის Streamlit Secrets-ში!")

        service_account_info = dict(st.secrets["gcp_service_account"])

        # ვამოწმებთ, რომ ყველა საჭირო ველი არის
        required_fields = ["type", "project_id", "private_key_id", "private_key",
                          "client_email", "client_id", "auth_uri", "token_uri"]
        missing = [f for f in required_fields if f not in service_account_info]
        if missing:
            raise Exception(f"ამ ველები აკლია: {missing}")

        # pygsheets-ის ავტორიზაცია service account-ით
        gc = pygsheets.authorize(service_account_info=service_account_info)

        # ვხსნით ცხრილს
        sh = gc.open("NeuroCRM_DB")
        return sh

    except Exception as e:
        st.error(f"❌ Google Sheets-თან კავშირის შეცდომა: {e}")
        st.error("გთხოვ, დარწმუნდი რომ:")
        st.error("1. Streamlit Cloud-ზე secrets-ში გაქვს gcp_service_account")
        st.error("2. Google Sheets-ში NeuroCRM_DB ფაილი გაზიარებულია service account-ის email-ზე")
        return None


# --- EMAIL SYSTEM ---
def send_confirmation_email(to_email: str, patient_name: str, slot_time: str) -> bool:
    """
    უგზავნის პაციენტს ჯავშნის დადასტურების მეილს.
    """
    # თუ შენი მეილი ჯერ არ ჩაგიწერია, ფუნქცია არაფერს აკეთებს
    if "your_email" in SENDER_EMAIL:
        return False

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
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = Header(subject_text, "utf-8")
    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"მეილის გაგზავნის ერორი: {e}")
        return False


# --- HELPER FUNCTIONS (pygsheets ვერსია) ---
def get_slots_ws(spreadsheet):
    return spreadsheet.worksheet_by_title("Slots")


def get_patients_ws(spreadsheet):
    return spreadsheet.worksheet_by_title("Sheet1")


def get_free_slots(spreadsheet):
    """აბრუნებს თავისუფალი სლოტების list-ს 'Date | Time' ფორმატში."""
    try:
        ws = get_slots_ws(spreadsheet)
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            return []

        df["Display"] = df["Date"] + " | " + df["Time"]
        free = df[df["Status"] == "Open"]["Display"].tolist()
        return sorted(free)
    except Exception as e:
        st.error(f"სლოტების წაკითხვის ერორი: {e}")
        return []


def mark_slot_booked(spreadsheet, slot_display: str) -> bool:
    """მითითებულ სლოტს 'Booked'-ად აქცევს."""
    try:
        ws = get_slots_ws(spreadsheet)
        date_part, time_part = slot_display.split(" | ")

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            # row index i → sheet row i+1
            if len(row) >= 3 and row[0] == date_part and row[1] == time_part and row[2] == "Open":
                ws.update_value((i + 1, 3), "Booked")
                return True
        return False
    except Exception as e:
        st.error(f"სლოტის დაჯავშნის ერორი: {e}")
        return False


def add_new_slot(spreadsheet, date, time) -> bool:
    """Slots ფურცელზე ამატებს ახალ Open სლოტს."""
    try:
        ws = get_slots_ws(spreadsheet)
        # append_table ბოლოს ამატებს ახალ სტრიქონს
        ws.append_table(values=[str(date), str(time), "Open"], dimension="ROWS")
        return True
    except Exception as e:
        st.error(f"ახალი სლოტის დამატების ერორი: {e}")
        return False


def delete_slot(spreadsheet, date, time) -> bool:
    """წაშლის კონკრეტულ სლოტს Slots ფურცლიდან."""
    try:
        ws = get_slots_ws(spreadsheet)
        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if row and len(row) >= 2 and row[0] == str(date) and row[1] == str(time):
                ws.delete_rows(i + 1)
                return True
        return False
    except Exception as e:
        st.error(f"სლოტის წაშლის ერორი: {e}")
        return False


def get_all_slots(spreadsheet) -> pd.DataFrame:
    """აბრუნებს Slots ფურცლის მთელ ცხრილს DataFrame-ის სახით."""
    try:
        ws = get_slots_ws(spreadsheet)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"სლოტების DataFrame-ის ერორი: {e}")
        return pd.DataFrame()


def generate_time_options():
    times = []
    for h in range(9, 22):
        times.append(f"{h:02d}:00")
        times.append(f"{h:02d}:30")
    return times


# --- MAIN APP ---
def main():
    st.title("🧠 NeuroCRM v3.0 (pygsheets)")

    db = connect_db()
    if not db:
        st.warning("მონაცემთა ბაზა მიუწვდომელია. გადაამოწმე Streamlit secrets და Google Sheets-ის დაშვებები.")
        return

    st.sidebar.title("ნავიგაცია")
    menu = st.sidebar.radio("მენიუ:", ["პაციენტის რეგისტრაცია", "ადმინ პანელი"])

    # --- პაციენტის რეგისტრაცია ---
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
                if not (name and phone and email):
                    st.error("გთხოვ, შეავსე ყველა ველი.")
                elif slot == "ადგილები არ არის":
                    st.error("ამ ეტაპზე თავისუფალი დრო არ არის.")
                else:
                    if mark_slot_booked(db, slot):
                        try:
                            ws = get_patients_ws(db)
                            ws.append_table(
                                values=[
                                    str(datetime.now().timestamp()),
                                    name,
                                    phone,
                                    email,
                                    "Active",
                                    "",
                                    str(datetime.now().date()),
                                    slot,
                                ],
                                dimension="ROWS",
                            )
                        except Exception as e:
                            st.error(f"პაციენტის მონაცემების შენახვის ერორი: {e}")
                        else:
                            send_confirmation_email(email, name, slot)
                            st.success("✅ დაჯავშნილია!")
                            st.rerun()
                    else:
                        st.error("❌ სლოტი უკვე დაკავებულია ან ვერ მოიძებნა.")

    # --- ადმინ პანელი ---
    elif menu == "ადმინ პანელი":
        pwd = st.sidebar.text_input("პაროლი", type="password")
        if pwd != ADMIN_PASSWORD:
            st.warning("ადმინ პანელზე წვდომისთვის შეიყვანე სწორი პაროლი.")
            return

        t1, t2, t3 = st.tabs(["დამატება", "წაშლა", "ბაზა"])

        # --- Slots დამატება ---
        with t1:
            st.subheader("ახალი სლოტის დამატება")
            d = st.date_input("თარიღი")
            t = st.selectbox("საათი", generate_time_options())
            if st.button("სლოტის დამატება"):
                if add_new_slot(db, d, t):
                    st.success("✅ სლოტი დამატებულია.")
                else:
                    st.error("❌ სლოტის დამატება ვერ მოხერხდა.")

        # --- Slots წაშლა ---
        with t2:
            st.subheader("სლოტის წაშლა")
            df = get_all_slots(db)
            if df.empty:
                st.info("სლოტები არ არის დამატებული.")
            else:
                df["S"] = df["Date"] + " | " + df["Time"] + " (" + df["Status"] + ")"
                to_del = st.selectbox("აირჩიე სლოტი", df["S"])
                if st.button("სლოტის წაშლა"):
                    raw = to_del.split(" (")[0]
                    dd, tt = raw.split(" | ")
                    if delete_slot(db, dd, tt):
                        st.success("✅ სლოტი წაიშალა.")
                        st.rerun()
                    else:
                        st.error("❌ სლოტის წაშლა ვერ მოხერხდა.")

        # --- ბაზის ნახვა ---
        with t3:
            st.subheader("პაციენტების ბაზა")
            try:
                ws = get_patients_ws(db)
                data = ws.get_all_records()
                st.dataframe(pd.DataFrame(data))
            except Exception as e:
                st.error(f"ბაზის წამოღების ერორი: {e}")


if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import pygsheets
from datetime import datetime
import smtplib
from email.mime_text import MIMEText
from email.mime_multipart import MIMEMultipart
from email.header import Header

# --- კონფიგურაცია ---
st.set_page_config(page_title="NeuroCRM", page_icon="🧠", layout="wide")

# პაროლები და უსაფრთხოება
ADMIN_PASSWORD = "Neuro2025"
SENDER_EMAIL = "your_email@gmail.com"          # ჩაწერე შენი მეილი
SENDER_PASSWORD = "xxxx xxxx xxxx xxxx"        # Gmail App Password (არა ჩვეულებრივი პაროლი!)

# --- DATABASE CONNECTION (pygsheets ვერსია) ---
@st.cache_resource
def connect_db():
    """
    უკავშირდება Google Sheets-ს pygsheets-ის საშუალებით.
    ჯერ ცდილობს Streamlit Secrets-დან წაკითხვას,
    თუ ვერ მოახერხა – creds.json-დან (ლოკალური ტესტისთვის).
    """
    try:
        # სცენარი 1: ვკითულობთ Streamlit secrets-დან
        if "gcp_service_account" in st.secrets:
            service_account_info = dict(st.secrets["gcp_service_account"])
            gc = pygsheets.authorize(service_account_info=service_account_info)
        else:
            # სცენარი 2: fallback – ლოკალური ფაილით
            gc = pygsheets.authorize(service_file="creds.json")

        sh = gc.open("NeuroCRM_DB")
        return sh

    except Exception as e:
        st.error(f"კავშირის კრიტიკული შეცდომა (DB): {e}")
        return None


# --- EMAIL SYSTEM (UTF-8 FIX INCLUDED) ---
def send_confirmation_email(to_email: str, patient_name: str, slot_time: str) -> bool:
    """გაგზავნის დამადასტურებელ მეილს პაციენტს."""
    if "your_email" in SENDER_EMAIL:
        # თუ შენს რეალურ მეილს არ ჩაწერ, არაფერი გააკეთოს
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


# --- HELPER FUNCTIONS (pygsheets) ---
def get_free_slots(spreadsheet):
    """იღებს ყველა თავისუფალ (Open) სლოტს Slots worksheet-დან."""
    try:
        ws = spreadsheet.worksheet_by_title("Slots")
        data = ws.get_all_records()
        df = pd.DataFrame(data)

        if df.empty:
            return []

        df["Display"] = df["Date"] + " | " + df["Time"]
        free_df = df[df["Status"] == "Open"]

        return sorted(free_df["Display"].tolist())
    except Exception as e:
        st.error(f"სლოტების წაკითხვის ერორი: {e}")
        return []


def mark_slot_booked(spreadsheet, slot_display: str) -> bool:
    """
    მონიშნავს სლოტს როგორც Booked Slots worksheet-ში.
    slot_display ფორმატია: 'YYYY-MM-DD | HH:MM'
    """
    try:
        ws = spreadsheet.worksheet_by_title("Slots")
        date_part, time_part = slot_display.split(" | ")

        all_values = ws.get_all_values(
            include_tailing_empty=False,
            include_tailing_empty_rows=False,
        )

        # ვეძებთ სტრიქონს, სადაც Date, Time, Status == Open
        for i, row in enumerate(all_values):
            if len(row) >= 3 and row[0] == date_part and row[1] == time_part and row[2] == "Open":
                # pygsheets: update_value((row, col), value)
                ws.update_value((i + 1, 3), "Booked")
                return True

        return False
    except Exception as e:
        st.error(f"სლოტის დაჯავშნის ერორი: {e}")
        return False


def add_new_slot(spreadsheet, date, time) -> bool:
    """ძAdds new Open slot in Slots worksheet."""
    try:
        ws = spreadsheet.worksheet_by_title("Slots")
        ws.append_table(
            values=[str(date), str(time), "Open"],
            dimension="ROWS",
        )
        return True
    except Exception as e:
        st.error(f"ახალი სლოტის დამატების ერორი: {e}")
        return False


def delete_slot(spreadsheet, date, time) -> bool:
    """შლის კონკრეტულ სლოტს Slots worksheet-იდან."""
    try:
        ws = spreadsheet.worksheet_by_title("Slots")
        all_values = ws.get_all_values(
            include_tailing_empty=False,
            include_tailing_empty_rows=False,
        )

        for i, row in enumerate(all_values):
            if len(row) >= 2 and row[0] == str(date) and row[1] == str(time):
                ws.delete_rows(i + 1)
                return True

        return False
    except Exception as e:
        st.error(f"სლოტის წაშლის ერორი: {e}")
        return False


def get_all_slots(spreadsheet) -> pd.DataFrame:
    """აბრუნებს ყველა სლოტს DataFrame-ს სახით."""
    try:
        ws = spreadsheet.worksheet_by_title("Slots")
        return pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.error(f"ყველა სლოტის წაკითხვის ერორი: {e}")
        return pd.DataFrame()


def generate_time_options():
    """აგენერირებს დროის სლოტებს 09:00–21:30, ნახევარ-ნახევარ საათზე."""
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
        st.warning("მონაცემთა ბაზა მიუწვდომელია. შეამოწმეთ Secrets ან creds.json.")
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
                    st.error("გთხოვ, შეავსო ყველა ველი.")
                elif slot == "ადგილები არ არის":
                    st.error("თავისუფალი ადგილები ამ ეტაპზე არაა.")
                else:
                    # ვცდილობთ სლოტის დაჯავშნას
                    if mark_slot_booked(db, slot):
                        try:
                            ws = db.worksheet_by_title("Sheet1")
                        except Exception as e:
                            st.error(f"Sheet1-ის გახსნის ერორი: {e}")
                        else:
                            try:
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
                                send_confirmation_email(email, name, slot)
                                st.success("✅ დაჯავშნილია!")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"პაციენტის ჩანაწერის ჩაწერის ერორი: {e}")
                    else:
                        st.error("❌ ეს დრო უკვე დაკავებულია ან ვერ მოიძებნა.")

    # --- ადმინ პანელი ---
    elif menu == "ადმინ პანელი":
        pwd = st.sidebar.text_input("პაროლი", type="password")
        if pwd != ADMIN_PASSWORD:
            st.warning("არასწორი პაროლი.")
            return

        t1, t2, t3 = st.tabs(["დამატება", "წაშლა", "ბაზა"])

        # ახალი სლოტის დამატება
        with t1:
            d = st.date_input("თარიღი")
            t = st.selectbox("საათი", generate_time_options())
            if st.button("დამატება"):
                if add_new_slot(db, d, t):
                    st.success("სლოტი დამატებულია.")
                    st.cache_data.clear()
                else:
                    st.error("სლოტის დამატება ვერ მოხერხდა.")

        # სლოტის წაშლა
        with t2:
            df = get_all_slots(db)
            if df.empty:
                st.info("სლოტები ვერ მოიძებნა.")
            else:
                df["S"] = df["Date"] + " | " + df["Time"] + " (" + df["Status"] + ")"
                to_del = st.selectbox("აირჩიე სლოტი წასაშლელად", df["S"])
                if st.button("წაშლა"):
                    raw = to_del.split(" (")[0]
                    dd, tt = raw.split(" | ")
                    if delete_slot(db, dd, tt):
                        st.success("სლოტი წაიშალა.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("სლოტის წაშლა ვერ მოხერხდა.")

        # ძირითადი ბაზა (Sheet1)
        with t3:
            try:
                ws = db.worksheet_by_title("Sheet1")
                st.dataframe(pd.DataFrame(ws.get_all_records()))
            except Exception as e:
                st.error(f"Sheet1-ის წაკითხვის ერორი: {e}")


if __name__ == "__main__":
    main()

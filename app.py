import pygsheets
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import io

# --- CONFIGURATION FROM SECRETS ---
def get_email_config():
    """იღებს მეილის კონფიგურაციას Streamlit Secrets-დან."""
    try:
        sender_email = st.secrets.get("SENDER_EMAIL", "")
        sender_password = st.secrets.get("SENDER_PASSWORD", "")
        return sender_email, sender_password
    except Exception:
        return "", ""


def get_admin_password():
    """იღებს ადმინის პაროლს Streamlit Secrets-დან."""
    try:
        return st.secrets.get("ADMIN_PASSWORD", "admin123")
    except Exception:
        return "admin123"


@st.cache_resource
def connect_db():
    try:
        service_account_info = dict(st.secrets["gcp_service_account"])
        gc = pygsheets.authorize(service_account_info=service_account_info)
        sh = gc.open("NeuroCRM_DB")
        return sh
    except Exception as e:
        st.error(f"Google Sheets-თან კავშირის შეცდომა: {e}")
        return None


# --- EMAIL SYSTEM ---
def send_confirmation_email(to_email: str, patient_name: str, slot_time: str) -> bool:
    """უგზავნის პაციენტს ჯავშნის დადასტურების მეილს."""
    sender_email, sender_password = get_email_config()

    if not sender_email or not sender_password:
        st.warning("მეილის კონფიგურაცია არ არის დაყენებული. დაამატე SENDER_EMAIL და SENDER_PASSWORD secrets-ში.")
        return False

    subject_text = "ჯავშნის დადასტურება - ფსიქოთერაპია"
    body = f"""
გამარჯობა {patient_name},

თქვენი ვიზიტი წარმატებით დაიჯავშნა.
დრო: {slot_time}

მისამართი: თბილისი...

გთხოვთ, მობრძანდეთ დანიშნულ დროს.

პატივისცემით,
NeuroCRM System.
"""

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = Header(subject_text, "utf-8")
    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"მეილის გაგზავნის შეცდომა: {e}")
        return False


# --- HELPER FUNCTIONS ---
def get_slots_ws(spreadsheet):
    return spreadsheet.worksheet_by_title("Slots")


def get_patients_ws(spreadsheet):
    return spreadsheet.worksheet_by_title("Sheet1")


def get_sessions_ws(spreadsheet):
    """იღებს ან ქმნის Sessions ფურცელს ჩატარებული სესიებისთვის."""
    try:
        return spreadsheet.worksheet_by_title("Sessions")
    except pygsheets.WorksheetNotFound:
        # თუ არ არსებობს, ვქმნით ახალს
        ws = spreadsheet.add_worksheet("Sessions")
        ws.update_row(1, ["ID", "PatientName", "Phone", "Email", "SlotTime", "CompletedDate", "Notes"])
        return ws


def get_free_slots(spreadsheet):
    """აბრუნებს თავისუფალი სლოტების list-ს."""
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
        st.error(f"სლოტების წაკითხვის შეცდომა: {e}")
        return []


def mark_slot_booked(spreadsheet, slot_display: str) -> bool:
    """მითითებულ სლოტს 'Booked'-ად აქცევს."""
    try:
        ws = get_slots_ws(spreadsheet)
        date_part, time_part = slot_display.split(" | ")

        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if len(row) >= 3 and row[0] == date_part and row[1] == time_part and row[2] == "Open":
                ws.update_value((i + 1, 3), "Booked")
                return True
        return False
    except Exception as e:
        st.error(f"სლოტის დაჯავშნის შეცდომა: {e}")
        return False


def add_new_slot(spreadsheet, date, time) -> bool:
    """Slots ფურცელზე ამატებს ახალ Open სლოტს."""
    try:
        ws = get_slots_ws(spreadsheet)
        ws.append_table(values=[str(date), str(time), "Open"], dimension="ROWS")
        return True
    except Exception as e:
        st.error(f"ახალი სლოტის დამატების შეცდომა: {e}")
        return False


def delete_slot(spreadsheet, date, time) -> bool:
    """წაშლის კონკრეტულ სლოტს."""
    try:
        ws = get_slots_ws(spreadsheet)
        all_values = ws.get_all_values()
        for i, row in enumerate(all_values):
            if row and len(row) >= 2 and row[0] == str(date) and row[1] == str(time):
                ws.delete_rows(i + 1)
                return True
        return False
    except Exception as e:
        st.error(f"სლოტის წაშლის შეცდომა: {e}")
        return False


def get_all_slots(spreadsheet) -> pd.DataFrame:
    """აბრუნებს Slots ფურცლის მთელ ცხრილს."""
    try:
        ws = get_slots_ws(spreadsheet)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"სლოტების DataFrame-ის შეცდომა: {e}")
        return pd.DataFrame()


def get_all_patients(spreadsheet) -> pd.DataFrame:
    """აბრუნებს პაციენტების ბაზას."""
    try:
        ws = get_patients_ws(spreadsheet)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"პაციენტების წამოღების შეცდომა: {e}")
        return pd.DataFrame()


def get_all_sessions(spreadsheet) -> pd.DataFrame:
    """აბრუნებს ჩატარებული სესიების ბაზას."""
    try:
        ws = get_sessions_ws(spreadsheet)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"სესიების წამოღების შეცდომა: {e}")
        return pd.DataFrame()


def mark_session_completed(spreadsheet, patient_data: dict, notes: str = "") -> bool:
    """მონიშნავს სესიას როგორც ჩატარებულს."""
    try:
        ws = get_sessions_ws(spreadsheet)
        ws.append_table(
            values=[
                str(datetime.now().timestamp()),
                patient_data.get("Name", ""),
                patient_data.get("Phone", ""),
                patient_data.get("Email", ""),
                patient_data.get("Slot", ""),
                str(datetime.now().date()),
                notes
            ],
            dimension="ROWS"
        )

        # პაციენტის სტატუსს ვცვლით "Completed"-ზე
        patients_ws = get_patients_ws(spreadsheet)
        all_values = patients_ws.get_all_values()
        for i, row in enumerate(all_values):
            if len(row) >= 8 and row[1] == patient_data.get("Name") and row[7] == patient_data.get("Slot"):
                patients_ws.update_value((i + 1, 5), "Completed")
                break

        return True
    except Exception as e:
        st.error(f"სესიის მონიშვნის შეცდომა: {e}")
        return False


def get_statistics(spreadsheet) -> dict:
    """ითვლის სტატისტიკას: დღის, კვირის, თვის სესიები."""
    try:
        sessions_df = get_all_sessions(spreadsheet)
        if sessions_df.empty:
            return {"today": 0, "week": 0, "month": 0, "total": 0}

        today = datetime.now().date()
        week_ago = today - timedelta(days=7)
        month_ago = today - timedelta(days=30)

        sessions_df["CompletedDate"] = pd.to_datetime(sessions_df["CompletedDate"], errors="coerce").dt.date

        today_count = len(sessions_df[sessions_df["CompletedDate"] == today])
        week_count = len(sessions_df[sessions_df["CompletedDate"] >= week_ago])
        month_count = len(sessions_df[sessions_df["CompletedDate"] >= month_ago])
        total_count = len(sessions_df)

        return {
            "today": today_count,
            "week": week_count,
            "month": month_count,
            "total": total_count
        }
    except Exception as e:
        st.error(f"სტატისტიკის გამოთვლის შეცდომა: {e}")
        return {"today": 0, "week": 0, "month": 0, "total": 0}


def generate_time_options():
    times = []
    for h in range(9, 22):
        times.append(f"{h:02d}:00")
        times.append(f"{h:02d}:30")
    return times


def convert_df_to_csv(df):
    """DataFrame-ს გარდაქმნის CSV-ში."""
    return df.to_csv(index=False).encode('utf-8')


def convert_df_to_excel(df):
    """DataFrame-ს გარდაქმნის Excel-ში."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()


# --- MAIN APP ---
def main():
    st.set_page_config(page_title="NeuroCRM", page_icon="🧠", layout="wide")
    st.title("🧠 NeuroCRM v4.0")

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
            slot = c2.selectbox("აირჩიეთ თავისუფალი დრო", free_slots if free_slots else ["ადგილები არ არის"])

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
                            st.error(f"პაციენტის მონაცემების შენახვის შეცდომა: {e}")
                        else:
                            email_sent = send_confirmation_email(email, name, slot)
                            if email_sent:
                                st.success("✅ დაჯავშნილია! დადასტურების მეილი გაიგზავნა.")
                            else:
                                st.success("✅ დაჯავშნილია!")
                            st.rerun()
                    else:
                        st.error("სლოტი უკვე დაკავებულია ან ვერ მოიძებნა.")

    # --- ადმინ პანელი ---
    elif menu == "ადმინ პანელი":
        pwd = st.sidebar.text_input("პაროლი", type="password")
        if pwd != get_admin_password():
            st.warning("ადმინ პანელზე წვდომისთვის შეიყვანე სწორი პაროლი.")
            return

        # --- სტატისტიკა ---
        st.header("📊 სტატისტიკა")
        stats = get_statistics(db)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("დღეს", stats["today"])
        col2.metric("კვირაში", stats["week"])
        col3.metric("თვეში", stats["month"])
        col4.metric("სულ", stats["total"])

        st.divider()

        # --- ტაბები ---
        t1, t2, t3, t4, t5, t6 = st.tabs([
            "📅 კალენდარი",
            "✅ სესიის მონიშვნა",
            "➕ სლოტის დამატება",
            "🗑️ სლოტის წაშლა",
            "👥 პაციენტები",
            "📥 ექსპორტი"
        ])

        # --- კალენდარი ---
        with t1:
            st.subheader("დაჯავშნილი ვიზიტები")
            patients_df = get_all_patients(db)
            if patients_df.empty:
                st.info("ჯერ არცერთი ჯავშანი არ არის.")
            else:
                # აქტიური ჯავშნები
                active_patients = patients_df[patients_df.get("Status", pd.Series()) == "Active"] if "Status" in patients_df.columns else patients_df

                if not active_patients.empty:
                    # კალენდარის ფორმატში გადაყვანა
                    calendar_data = []
                    for _, row in active_patients.iterrows():
                        slot = row.get("Slot", "") if "Slot" in row else ""
                        if slot and " | " in slot:
                            date_part, time_part = slot.split(" | ")
                            calendar_data.append({
                                "თარიღი": date_part,
                                "დრო": time_part,
                                "პაციენტი": row.get("Name", "") if "Name" in row else "",
                                "ტელეფონი": row.get("Phone", "") if "Phone" in row else "",
                                "სტატუსი": row.get("Status", "") if "Status" in row else ""
                            })

                    if calendar_data:
                        calendar_df = pd.DataFrame(calendar_data)
                        calendar_df = calendar_df.sort_values("თარიღი")
                        st.dataframe(calendar_df, use_container_width=True)
                    else:
                        st.info("აქტიური ჯავშნები არ არის.")
                else:
                    st.info("აქტიური ჯავშნები არ არის.")

        # --- სესიის მონიშვნა ---
        with t2:
            st.subheader("სესიის ჩატარების მონიშვნა")
            patients_df = get_all_patients(db)

            if patients_df.empty:
                st.info("პაციენტები არ არის.")
            else:
                # მხოლოდ აქტიური პაციენტები
                if "Status" in patients_df.columns:
                    active = patients_df[patients_df["Status"] == "Active"]
                else:
                    active = patients_df

                if active.empty:
                    st.info("აქტიური ჯავშნები არ არის.")
                else:
                    # სელექტორისთვის ტექსტის ფორმირება
                    options = []
                    for _, row in active.iterrows():
                        name = row.get("Name", "უცნობი")
                        slot = row.get("Slot", "")
                        options.append(f"{name} - {slot}")

                    selected = st.selectbox("აირჩიე პაციენტი", options)
                    notes = st.text_area("შენიშვნები (არასავალდებულო)")

                    if st.button("✅ მონიშნე როგორც ჩატარებული"):
                        # ვპოულობთ შესაბამის პაციენტს
                        idx = options.index(selected)
                        patient_row = active.iloc[idx]

                        patient_data = {
                            "Name": patient_row.get("Name", ""),
                            "Phone": patient_row.get("Phone", ""),
                            "Email": patient_row.get("Email", ""),
                            "Slot": patient_row.get("Slot", "")
                        }

                        if mark_session_completed(db, patient_data, notes):
                            st.success("✅ სესია მონიშნულია როგორც ჩატარებული!")
                            st.rerun()
                        else:
                            st.error("შეცდომა სესიის მონიშვნისას.")

        # --- სლოტის დამატება ---
        with t3:
            st.subheader("ახალი სლოტის დამატება")
            d = st.date_input("თარიღი")
            t = st.selectbox("საათი", generate_time_options())
            if st.button("სლოტის დამატება"):
                if add_new_slot(db, d, t):
                    st.success("✅ სლოტი დამატებულია.")
                    st.rerun()
                else:
                    st.error("სლოტის დამატება ვერ მოხერხდა.")

        # --- სლოტის წაშლა ---
        with t4:
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
                        st.error("სლოტის წაშლა ვერ მოხერხდა.")

        # --- პაციენტების ბაზა ---
        with t5:
            st.subheader("პაციენტების ბაზა")
            patients_df = get_all_patients(db)
            if patients_df.empty:
                st.info("პაციენტები არ არის.")
            else:
                st.dataframe(patients_df, use_container_width=True)

            st.divider()
            st.subheader("ჩატარებული სესიები")
            sessions_df = get_all_sessions(db)
            if sessions_df.empty:
                st.info("ჩატარებული სესიები არ არის.")
            else:
                st.dataframe(sessions_df, use_container_width=True)

        # --- ექსპორტი ---
        with t6:
            st.subheader("მონაცემების ექსპორტი")

            export_type = st.radio("რა მონაცემები გსურთ?", ["პაციენტები", "სესიები", "სლოტები"])

            if export_type == "პაციენტები":
                df = get_all_patients(db)
            elif export_type == "სესიები":
                df = get_all_sessions(db)
            else:
                df = get_all_slots(db)

            if df.empty:
                st.info("მონაცემები არ არის.")
            else:
                st.dataframe(df, use_container_width=True)

                col1, col2 = st.columns(2)

                with col1:
                    csv_data = convert_df_to_csv(df)
                    st.download_button(
                        label="📥 ჩამოტვირთე CSV",
                        data=csv_data,
                        file_name=f"{export_type}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )

                with col2:
                    try:
                        excel_data = convert_df_to_excel(df)
                        st.download_button(
                            label="📥 ჩამოტვირთე Excel",
                            data=excel_data,
                            file_name=f"{export_type}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception:
                        st.info("Excel-ისთვის საჭიროა openpyxl ბიბლიოთეკა.")


if __name__ == "__main__":
    main()

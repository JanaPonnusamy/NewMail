import streamlit as st
import shutil
import smtplib
from email.message import EmailMessage
import os

# ---------- Email Sending Logic ----------
def send_email(to_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = 'janaponnusamy@gmail.com'  # ✅ Your Gmail
    msg['To'] = to_email
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='octet-stream',
            filename=os.path.basename(attachment_path)
        )

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('janaponnusamy@gmail.com', 'phku opdk dzwt teui')  # ✅ Gmail App Password
        smtp.send_message(msg)

# ---------- Streamlit UI ----------
st.title("📤 Excel File Sender")

uploaded_files = st.file_uploader("📁 Select Excel files", type=['xls', 'xlsx'], accept_multiple_files=True)
recipient = st.text_input("📧 Recipient Email")

if st.button("📨 Send Email"):
    if uploaded_files and recipient:
        # Create temp folder
        temp_folder = "temp_excel_files"
        os.makedirs(temp_folder, exist_ok=True)

        # Save uploaded files
        for file in uploaded_files:
            with open(os.path.join(temp_folder, file.name), 'wb') as f:
                f.write(file.getbuffer())

        # Create ZIP archive
        zip_name = "excel_files.zip"
        shutil.make_archive(zip_name.replace('.zip', ''), 'zip', temp_folder)

        # Send the email
        try:
            send_email(recipient, "Excel Files", "Please find attached the Excel files.", zip_name)
            st.success("✅ Email sent successfully!")
        except Exception as e:
            st.error(f"❌ Failed to send email: {e}")

        # Clean up files
        try:
            shutil.rmtree(temp_folder)
            os.remove(zip_name)
        except Exception as e:
            st.warning(f"⚠️ Cleanup warning: {e}")
    else:
        st.warning("⚠️ Please upload at least one Excel file and enter a recipient email.")

import pandas as pd
from docx import Document
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)

    for p in doc.paragraphs:
        for key, value in data.items():
            if key in p.text:
                for run in p.runs:
                    run.text = run.text.replace(key, value)
    doc_content = "\n".join([p.text for p in doc.paragraphs])

    return doc_content


def generate_invitation(csv_path, template_path):
    df = pd.read_csv(csv_path)

    for _, row in df.iterrows():
        full_name = f"{row['first_name']}_{row['last_name']}"
        data = {
            "[Salutation]": row["salutation"],
            "[First Name]": row["first_name"],
            "[Last Name]": row["last_name"],
            "[Last Contacted]": row["last_contacted"],
            "[Company Name]": row["company"],
        }
        output_path = f"{full_name} invitation.docx"
        yield fill_invitation(template_path, output_path, data)


def send_email(csv_path, subject, username, password, body):
    msg = MIMEText(body, "plain", "utf-8")
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(username, password)
    df = pd.read_csv(csv_path)
    for idx, row in df.iterrows():
        email = row["email"]
        server.sendmail(username, email, msg.as_string(), subject)
    server.quit()


if __name__ == "__main__":
    csv_path = "contacts.csv"
    template_path = "template.docx"
    subject = "Invite to Event in 2023"
    username = "abdallahfetyani2@gmail.com"
    password = "vawr zjoj epoa jntx"
    for inv in generate_invitation(csv_path, template_path):
        send_email(csv_path, subject, username, password, body=inv)

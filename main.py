import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment
import re

email_user = 'YOUREMAIL'
email_password = 'YOURPASSWORD'

# Gmail setup
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(email_user, email_password)

# Select mailbox
mail.select("inbox")

# Filter emails(here we're listing just UNSEEN emails)
status, messages = mail.search(None, "UNSEEN")
messages_ids = messages[0].split()

# Creating xlsx
wb = Workbook()
ws = wb.active
ws.title = "Values and dates"
ws.append(["Date", "Value"])

value_pattern = re.compile(r"R\$\s?(\d+,\d{2})")
date_pattern = re.compile(r"(\d{1,2}\sde\s[a-zA-Z]+\sde\s\d{4})")

# Iterate over emails
for mail_id in messages_ids:
    print("==== New E-mail ====")

    status, msg_data = mail.fetch(mail_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8")

    from_, encoding = decode_header(msg.get("From"))[0]
    if isinstance(from_, bytes):
        from_ = from_.decode(encoding or "utf-8")

    print(f"Assunto: {subject}")
    print(f"De: {from_}")

    body = ""
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            try:
                body = part.get_payload(decode=True).decode("utf-8")
            except UnicodeDecodeError:
                body = part.get_payload(decode=True).decode("latin-1", errors="ignore")

    print("Corpo do e-mail (HTML):")
    print(body)
    soup = BeautifulSoup(body, "html.parser")

    valores_match = re.search(value_pattern, soup.get_text())
    datas = re.findall(date_pattern, soup.get_text())

    if valores_match:
        valor = valores_match.group(1).replace(",", ".")
        print(f"Valor encontrado: {valor}")

        for data in datas:
            ws.append([data, float(valor)])

for column in ws.columns:
    max_length = 0
    column = [cell for cell in column]
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column[0].column_letter].width = adjusted_width
    for cell in column:
        cell.alignment = Alignment(horizontal='center')

wb.save("valores_e_datas.xlsx")

mail.logout()

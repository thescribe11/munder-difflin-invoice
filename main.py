#! python3

import openpyxl as excel
from openpyxl.styles import Font
import shutil
from datetime import timedelta, date
import smtplib, ssl, email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


bold_font = Font(bold=True)

MESSAGE_BODY = """
<html>
    <body>
    <p>Thank you for your purchase at Munder Difflin!</p>
    <p>Attached is the transaction invoice.</p>
    <p>If you have any problems, questions, or complaints, just reply to this email.</p>
    <p>Yours truly,<br>Mr. Dwight<br>Regional Manager, Munder Difflin</p>
    </body>
</html>
"""


def send_email(receiver_email, name):
    subject = "Here's the invoice for your purchase at Munder Difflin"
    sender_email = "youremail"  # You'll need to set up your own account in here.
    password = "yourpassword"   # Google may complain about this being a "less-secure app"; you'll have to go in your
    message = MIMEMultipart()   # account settings to fix that.
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    message.attach(MIMEText(MESSAGE_BODY, "html"))

    with open(f'Invoice for {name}.xlsx', 'rb') as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        "attachment; filename= Invoice.xlsx",
    )
    message.attach(part)

    text = message.as_string()
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)


def get_invoice_number():
    with open('invoice_number.txt', 'r') as f:
        numba = int(f.read()) + 1
    with open('invoice_number.txt', 'w') as f:
        f.write(str(numba))
    return numba


def generate_invoice(name, email_address, address, transactions: list) -> None:
    shutil.copy('invoice_template.xlsx', f'Invoice for {name}.xlsx')  # This is the invoice proper.

    wb = excel.load_workbook(f'Invoice for {name}.xlsx')
    sheet = wb['Sheet1']

    # Add customer information
    sheet['B3'].value = name
    sheet['B4'].value = email_address
    sheet['B5'].value = address

    # Determine and increment invoice number
    invoice_number = get_invoice_number()
    sheet['E3'].value = invoice_number

    # Add payment due date
    sheet['E4'].value = date.today() + timedelta(days=14)

    totals = []
    current_row = 9
    for i in transactions:
        print(i)
        sheet[f'B{current_row}'].value = i[0]  # Description
        sheet[f'C{current_row}'].value = i[1]  # Amount
        sheet[f'D{current_row}'].value = i[2]  # Cost per item
        spec_tot = i[2] * float((lambda x: x[1:len(x)])(i[1]))  # Amount of items * cost per item
        sheet[f'E{current_row}'].value = f'${spec_tot}'
        totals.append(spec_tot)
        current_row += 1

    total = sum(totals)
    current_row += 1
    sheet[f'D{current_row}'].font = bold_font
    sheet[f'D{current_row}'].value = "Total:"
    sheet[f'E{current_row}'].font = bold_font
    sheet[f'E{current_row}'].value = f'${total}'

    wb.save(f'Invoice for {name}.xlsx')


def convert_to_int(x):
    try:
        x = int(x)
    except ValueError:
        raise ValueError("Invalid input!")

    return x


def get_input():
    desc = input("Item description: ")
    cost = (lambda stuff: "$" + stuff if stuff[0] != '$' else stuff)(str(float(input("Cost per Unit:    "))))
    amnt = convert_to_int(input("Amount of Items:  "))

    return [desc, cost, amnt]


def main():
    entry_completed = 'N'
    transactions = []

    client_name = input("\nClient's name: ")
    client_email = input("Client's email address: ")
    client_addr = input("Client's actual address: ")

    while entry_completed == 'N':
        print()

        transactions.append(get_input())

        entry_completed = 'Y' if input("\nAdd another item (Y/n)? ").upper()[0] == 'N' else 'N'

    generate_invoice(client_name, client_email, client_addr, transactions)

    send_email(client_email, client_name)

    print("\n\nInvoice created.")


if __name__ == '__main__':
    main()

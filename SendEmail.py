import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import VAR

def SendEmail(filename, rowsCount, resultStr):
    # определяем сколнение по кол-ву строк
    if rowsCount>=11 and rowsCount<=19:
        countName = 'строк'
    else:
        i = rowsCount % 10
        if i == 1:
            countName = 'строка'
        elif i in [2,3,4]:
            countName = 'строки'
        else:
            countName = 'строк'

    # создаем письмо 
    subject = "Python task"
    body = resultStr + chr(10) + "Было выгружено " + str(rowsCount) + " " + countName
    sender_email = VAR.senderEmail
    receiver_email = VAR.receiverEmail
    password = VAR.password

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # добавляем тело письма
    message.attach(MIMEText(body, "plain"))

    # добавляем вложение
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # кодируем символы
    encoders.encode_base64(part)

    # Добавить заголовок в часть вложения
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Добавить вложение в сообщение и преобразовать сообщение в строку
    message.attach(part)
    text = message.as_string()

    # проходим аутентификацию и отправляем письмо
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(VAR.host, VAR.port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

def SendEmailError(resultStr):

    # создаем письмо 
    subject = "Python task - error"
    body = resultStr
    sender_email = VAR.senderEmail
    receiver_email = VAR.receiverEmail
    password = VAR.password

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # добавляем тело письма
    message.attach(MIMEText(body, "plain"))

    text = message.as_string()

    # проходим аутентификацию и отправляем письмо
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(VAR.host, VAR.port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(recipients, subject, text):
    sender = 'p.j.schietvereniging@gmail.com'

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject
    msg.attach(MIMEText(text, 'plain'))

    try:
        mail = smtplib.SMTP(host='smtp.gmail.com', port=587)
        mail.ehlo()
        mail.starttls()
        mail.login(sender, '!Password123')
        mail.sendmail(sender, recipients, msg.as_string())
        mail.close()
        print('Sent email successfully')
        return True
    except smtplib.SMTPException as e:
        print('Error sending email: ' + str(e))
        return False

import smtplib
import oracledb
import os
import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)


class SqlAndMail:

    lib_dir = "C:\\app\\product\\instantclient_12_2"
    oracledb.init_oracle_client(lib_dir=lib_dir)
    smtp_server = '172.23.100.20'
    smtp_port = 465
    email_sender = 'DispetcherGOR@polymetal.ru'

    @classmethod
    def send_emails(cls, receiver_email, body, subject, file_name):
        for mail in receiver_email:
            with smtplib.SMTP(cls.smtp_server, cls.smtp_port) as server:
                server.starttls()
                server.login(os.getenv("LOGIN"), os.getenv("PASSWORD_ML"), initial_response_ok=True)
                body = f'{body}'
                message = MIMEMultipart()
                message["Subject"] = f'{subject}'
                message["From"] = cls.email_sender
                message["To"] = mail

                if file_name:
                    file_name = f'{file_name}'
                    message_text = MIMEText(body, 'plain', 'utf-8')
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(open(file_name, 'rb').read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment; filename="%s"' % file_name)
                    message.attach(message_text)
                    message.attach(part)
                else:
                    message_text = MIMEText(body, 'plain', 'utf-8')
                    message.attach(message_text)

                server.ehlo()  # Может быть опущено
                server.sendmail(cls.email_sender, receiver_email, message.as_string())
                server.quit()

    @classmethod
    def cursor_data(cls, query, date, shift=None):
        connection = oracledb.connect(user=os.getenv('USER'), password=os.getenv('PASSWORD_BD'), dsn=os.getenv('DNS'))
        cursor = connection.cursor()
        cursor_data = cursor.var(oracledb.CURSOR)
        if shift:
            cursor.callproc(query, [date, shift, cursor_data])
        else:
            cursor.callproc(query, [date, cursor_data])
        cursor.close()
        return cursor_data

    @classmethod
    def p_date(cls, days):
        return (datetime.datetime.now() - datetime.timedelta(days=days)).strftime('%d.%m.%Y')

    @classmethod
    def p_date_sep(cls, days):
        return cls.p_date(days).split('.')

    @classmethod
    def now(cls):
        return datetime.datetime.now().strftime("%d.%m.%Y")

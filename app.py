import os
import pandas as pd
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message

ARCHIVE_FOLDER = './archive'
ACCOUNT = os.environ['mikado_exchange_account']
PASSWORD = os.environ['mikado_exchange_password']
MAILBOX = os.environ['mikado_exchange_mailbox']


creds = Credentials(
username=ACCOUNT,
password=PASSWORD)

config = Configuration(server='outlook.office365.com', credentials=creds)

account = Account(
primary_smtp_address=MAILBOX,
autodiscover=False,
config = config,
access_type=DELEGATE)


root_folder = account.root
dachser = root_folder.glob('**/3 Dachser Duitsland')

def init():
    if not os.path.exists(ARCHIVE_FOLDER):
        os.makedirs(ARCHIVE_FOLDER)

def get_latest_file():
    p = None
    for item in dachser.all().order_by('-datetime_received')[:1]:
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                filename, file_extension = os.path.splitext(attachment.name)
                if file_extension.lower() == '.csv':
                    p = os.path.join(ARCHIVE_FOLDER, attachment.name)
                    with open(p, 'wb') as f:
                        f.write(attachment.content)
    return p

def convert_mikado():
    fp = get_latest_file()
    if os.path.isfile(fp):
        df = pd.read_csv(fp, sep=';', skiprows=16, encoding='latin-1')
        df.to_csv('./mikado_report.csv')


if __name__ == '__main__':
    init()
    convert_mikado()

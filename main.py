import os
import pandas as pd
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message
from dotenv import load_dotenv
import logging
from mail import mail

from shutil import move

import pandas as pd
import math
import lxml.etree as etree

ARCHIVE_FOLDER = './archive'
OUTPUT_FILE = './mikado_report.csv'

MAGAZIJN_ID="030"
SYNC_SCRIPT_PATH="D:\King\Scripts\DachserSync.bat"
XML_FILE = 'dachser.xml'
KING_FILE = 'D:\\King\\dachser.xml'


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

def convert_mikado(OUTPUT_FILE):
    fp = get_latest_file()
    if os.path.isfile(fp):
        df = pd.read_csv(fp, sep=';', skiprows=16, encoding='latin-1')
        df['Article no.'] = df['Article no.'].str.replace("'", "")
        df['BBD'] = df['BBD'].str.replace(".", "/")
        df.to_csv(OUTPUT_FILE, sep=';')

def add_xml(node, art_id, amount):
    # Ignore VP articles for now
    if 'VP' in art_id:
        return
    item = etree.SubElement(node, 'VOORRAADCORRECTIEREGEL')
    etree.SubElement(item, 'VCR_ARTIKEL').text = art_id
    etree.SubElement(item, 'VCR_AANTAL').text = amount
    etree.SubElement(item, 'VCR_MAGAZIJN').text = MAGAZIJN_ID
    etree.SubElement(item, 'VCR_LOCATIE').text = '(Standaard)'

    # if 'VP' in art_id:
    #     print(art_id)
    #     etree.SubElement(item, 'VCR_PARTIJ').text = '(Standaard)'

def convert_csv(file):
    logging.info("Converting CSV")

    df = pd.read_csv(file, sep=';', decimal=',')
    root = etree.parse('king_voorraadcorrectie.xml').getroot()
    regels = root.find("VOORRAADCORRECTIES/VOORRAADCORRECTIE/VOORRAADCORRECTIE_REGELS")
    for i, row in df.iterrows():
        if math.isnan(row['Current SHU qty']):
            continue
        art_id = str(row['Article no.']).strip()
        amount = str(row['Current SHU qty']).strip()
        add_xml(regels, art_id, amount)

    with open(XML_FILE, 'wb') as f:
        f.write(etree.tostring(root, pretty_print=True))
    move(XML_FILE, KING_FILE)
    os.remove(file)

def sync_king():
    logging.info("Synchronizing KING")
    os.system(SYNC_SCRIPT_PATH)

if __name__ == '__main__':
    load_dotenv(verbose=True)
    logging.basicConfig(format='%(asctime)s %(message)s', filename='dachser.log', level=logging.INFO)


    ACCOUNT = os.getenv('mikado_exchange_account')
    PASSWORD = os.getenv('mikado_exchange_password')
    MAILBOX = os.getenv('mikado_exchange_mailbox')

    SMTP_SERVER = os.getenv('SMTP_SERVER')
    RECEIVER_EMAIL = os.getenv('RECEIVER_EMAIL')
    SENDER_EMAIL = os.getenv('SENDER_EMAIL')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')

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
    dachser = root_folder.glob('**/0000 Voorraadlijst - Dachser')

try:
    init()
    convert_mikado(OUTPUT_FILE)
    convert_csv(OUTPUT_FILE)
    sync_king()
except Exception as e:
    text = 'An import from the Dachser (Germany) system failed.\n\nPlease contact IT to review the following error:\n\n{}'.format(e)
    logging.error(e)
    mail(SMTP_SERVER, SENDER_EMAIL, RECEIVER_EMAIL, SENDER_PASSWORD, text)

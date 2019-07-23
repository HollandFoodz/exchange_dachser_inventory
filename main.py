import os
import pandas as pd
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message
from dotenv import load_dotenv
import logging
from mail import mail

from shutil import move

import pandas as pd
import math
import pyodbc
import lxml.etree as etree

from utils import *
from article import Article
from reset_inventory import reset_inventory


ARCHIVE_FOLDER = './archive'
CSV_FILE = './mikado_report.csv'

MAGAZIJN_ID="030"
SYNC_SCRIPT_PATH="D:\King\Scripts\DachserSync.bat"
XML_FILE = 'dachser.xml'
KING_FILE = 'D:\\King\\dachser.xml'
KING_XML_TEMPLATE = 'king_voorraadcorrectie.xml'



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

def convert_mikado(CSV_FILE):
    fp = get_latest_file()
    if os.path.isfile(fp):
        df = pd.read_csv(fp, sep=';', skiprows=16, encoding='latin-1')
        df['Article no.'] = df['Article no.'].str.replace("'", "")
        df['BBD'] = df['BBD'].str.replace(".", "/")
        df.to_csv(CSV_FILE, sep=';')


def convert_csv(cursor, mag_id, input_file, xml_file):
    logging.info("Converting CSV")

    df = pd.read_csv(input_file, sep=';', decimal=',')
    root, regels = get_xml_file_insert(xml_file)

    art_col = 'Article no.'
    articles = ', '.join('\'{}\''.format(str(row[art_col]).strip()) for _, row in df.iterrows())
    sql_query = 'SELECT ArtCode, ArtIsPartijRegistreren, KingSystem.tabArtikelPartij.ArtPartijNummer as ArtPartijNummer \
        from KingSystem.tabArtikel LEFT JOIN KingSystem.tabArtikelPartij \
        ON KingSystem.tabArtikel.ArtGid=KingSystem.tabArtikelPartij.ArtPartijArtGid \
        WHERE (KingSystem.tabArtikelPartij.ArtPartijIsGeblokkeerdVoorVerkoop = 0 OR KingSystem.tabArtikel.ArtIsPartijRegistreren = 0) AND \
        KingSystem.tabArtikel.ArtCode in ({})'.format(articles)

    # Fetch partij information
    cursor.execute(sql_query)
    rows = cursor.fetchall()
    articles = {}
    for row in rows:
        articles[row.ArtCode] = Article(row.ArtCode, row.ArtPartijNummer, row.ArtIsPartijRegistreren)

    for i, row in df.iterrows():
        if math.isnan(row['Current SHU qty']):
            continue

        art_id = str(row['Article no.']).strip()
        amount = str(row['Current SHU qty']).strip()
        article = articles[art_id]

        if article.partijregistratie:
            add_xml(regels, art_id, str(amount).strip(), mag_id, article.partijnummer)
        else:
            add_xml(regels, art_id, str(amount).strip(), mag_id)

    write_xml(xml_file, root)
    os.remove(input_file)

def sync_king(output_file):
    logging.info("Synchronizing KING")
    if os.path.exists(KING_FILE):
        os.remove(KING_FILE)
    move(output_file, KING_FILE)
    return os.system(SYNC_SCRIPT_PATH)

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

    ODBC_SOURCE = os.getenv('ODBC_SOURCE')
    ODBC_UID = os.getenv('ODBC_UID')
    ODBC_PWD = os.getenv('ODBC_PWD')

    conn_str = 'DSN={};UID={};PWD={}'.format(ODBC_SOURCE, ODBC_UID, ODBC_PWD)

    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

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

    exit_code = 0
    exit_msg = ''
    try:
        init()
        reset_inventory(cursor, MAGAZIJN_ID, KING_XML_TEMPLATE, XML_FILE)
        convert_mikado(CSV_FILE)
        convert_csv(cursor, MAGAZIJN_ID, CSV_FILE, XML_FILE)

        exit_code = sync_king(XML_FILE)

        if exit_code:
            error_msg = 'King could not import new inventory (Job failed)'
    except Exception as e:
        logging.error(e)
        exit_code = 1
        error_msg = e

    if exit_code:
        error_msg = 'An import from the Dachser (Germany) system failed.\n\nPlease contact IT to review the following error:\n\n{}'.format(error_msg)
        mail(SMTP_SERVER, SENDER_EMAIL, RECEIVER_EMAIL, SENDER_PASSWORD, error_msg)
    else:
        logging.info("Completed successfully")

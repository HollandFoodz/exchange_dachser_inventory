import lxml.etree as etree
from utils import *

# Reindex database?
ignored_articles = ['TSDFDF', 'Kok', 'Deegbereiding', 'Rollen', \
    'Flowpacken', 'chocolateren', 'Deegbreiding likz', 'Inleg lolly', 'Insteek lolly', 'Inpak likz', 'Inpak Nougat', \
    'Inpak stokken', 'Inpak wichtgoed', 'Magazijn', 'Nougat snijden', 'Spinnen lolly', 'TD']

def reset_inventory(cursor, mag_id, xml_file, output_file):
    root, regels = get_xml_file_insert(xml_file)
    sql_query = 'SELECT * from KingSystem.tabArtikel LEFT JOIN KingSystem.tabArtikelPartij \
        ON KingSystem.tabArtikel.ArtGid=KingSystem.tabArtikelPartij.ArtPartijArtGid \
        WHERE KingSystem.tabArtikelPartij.ArtPartijIsGeblokkeerdVoorVerkoop = 0 OR KingSystem.tabArtikel.ArtIsPartijRegistreren = 0'

    cursor.execute(sql_query)
    rows = cursor.fetchall()
    amount = 0
    for row in rows:
        art_id = str(row.ArtCode)

        # skip all the 000, 001, 099, etc.
        if len(art_id) == 3:
            if art_id[0].isdigit() and art_id[1].isdigit() and art_id[2].isdigit():
                continue

        if art_id in ignored_articles:
            continue

        if row.ArtIsPartijRegistreren:
            partij = str(row.ArtPartijNummer)
            add_xml(regels, art_id, str(amount), mag_id, partij)
        else:
            add_xml(regels, art_id, str(amount), mag_id)

    write_xml(output_file, root)

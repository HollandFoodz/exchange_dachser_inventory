import lxml.etree as etree
import pandas as pd
import os

def add_xml(node, art_id, amount, mag_id, partij=None):

    item = etree.SubElement(node, 'VOORRAADCORRECTIEREGEL')
    etree.SubElement(item, 'VCR_ARTIKEL').text = art_id
    etree.SubElement(item, 'VCR_AANTAL').text = amount
    etree.SubElement(item, 'VCR_MAGAZIJN').text = mag_id
    etree.SubElement(item, 'VCR_LOCATIE').text = '(Standaard)'

    if partij:
        etree.SubElement(item, 'VCR_PARTIJ').text = partij

def get_latest_file(directory):
    files = os.listdir(directory)
    paths = [os.path.join(directory, basename) for basename in files]
    latest_file = max(paths, key=os.path.getctime)
    return latest_file

def get_xml_file_insert(xml_file):
    root = etree.parse(xml_file).getroot()
    regels = root.find("VOORRAADCORRECTIES/VOORRAADCORRECTIE/VOORRAADCORRECTIE_REGELS")
    return root, regels

def write_xml(output_file, root):
    with open(output_file, 'wb') as f:
        f.write(etree.tostring(root, pretty_print=True))

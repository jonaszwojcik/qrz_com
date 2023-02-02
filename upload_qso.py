import os
import logging
import requests
import argparse
import pandas as pd
from datetime import datetime

# script for upload QSO from excel to QRZ.com

# documentation:
# https://www.qrz.com/docs/logbook/QRZLogbookAPI.html
# https://www.qrz.com/docs/logbook30/start
# https://www.qrz.com/docs/logbook30/adif-standard
# http://adif.org.uk/314/ADIF_314.htm#Fields

def main(qso_import_file):
    logging.basicConfig(filename=os.path.join(os.path.dirname(__file__), 'log', f'qrz_upload_log_{ts()}'), level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')
    qrz_api_key = get_api_key()
    adif_qso_list = import_from_excel(qso_import_file)
    upload_qso(qrz_api_key, adif_qso_list)

def upload_qso(qrz_api_key, adif_qso_list):
    upload_count = 0
    count = 0
    total = len(adif_qso_list)
    for adif_qso in adif_qso_list:
        count += 1
        logging.info(f'{upload_count}/{total}')
        logging.info(adif_qso)
        print(f'uploading:{upload_count}/{total} [{upload_count}]     ', end="\r")
        upload_status = qrz_logbook_upload_qso(qrz_api_key, adif_qso)
        logging.info(upload_status)
        if upload_status.find('RESULT=OK') > -1:
            upload_count += 1
    print(f'upladed:{upload_count}/{total} failed:{total-upload_count}  ')
        
def import_from_excel(file_path):
    qso_adif_list = list()
    qso_df = pd.read_excel(file_path)
    for qso_dict in qso_df.to_dict('records'):
        qso_adif_list.append(dict2adif(qso_dict))
    return qso_adif_list

def dict2adif(qso_dict):
    adif = ''
    for k in sorted(qso_dict.keys()):
        adif += f"<{k.lower()}:{len(str(qso_dict[k]))}>{str(qso_dict[k]).replace(',','.')}"
    adif += '<eor>'
    return adif

def qrz_logbook_upload_qso(qrz_api_key, qso_adif_data):
    upload_status = ''
    url = 'https://logbook.qrz.com/api'
    data = {'KEY':qrz_api_key, 'ACTION':'INSERT', 'ADIF':qso_adif_data}
    x = requests.post(url, data = data)
    if x.status_code != 200:
        print(x.status_code)
        print(x.text)
        upload_status += str(x.status_code)
    upload_status += x.text
    return upload_status

def get_api_key():
    api_key_file_name = os.path.join(os.path.dirname(__file__), 'api_key.txt')
    qrz_api_key = ''

    if os.path.exists(api_key_file_name):
        with open(api_key_file_name) as api_key_file:
            qrz_api_key = api_key_file.read()
    else:
        print('no api key file. please save your api key to api_key.txt file')
        exit(1)
    return qrz_api_key

def ts():
    now = datetime.now()
    ts_string = now.strftime("%Y%m%d%H%M")
    #ts_string = now.strftime("%Y%m%d")
    return ts_string

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('-i', '-import', required=True, help='qso import file name')
    parsed = ap.parse_args()
    main(parsed.i)

import os
from googleapiclient.http import MediaIoBaseDownload
import io
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from agricolletto.models.agricollet_models import Agricolletto


def agricolletto():
    agri = Agricolletto()
    agri.base_dir = r'C:\Users\ooaka\Downloads'
    files = agri.download_file()
    for file in files:
        print(file['name'], file['id'])
        request = agri.service.files().get_media(fileId=file['id'])
        agri.download_path = os.path.join(agri.base_dir, file['name'])
        print(agri.download_path)
        fh = io.FileIO(agri.download_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print ("Download %d%%." % int(status.progress() * 100))
        agri.convert_pdf()
        agri.shaping_pdf()
        agri.reflect_spreadsheet()
        agri.service.files().delete(fileId=file['id']).execute()

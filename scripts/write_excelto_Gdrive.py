from __future__ import print_function
#googleapiclient.discovery
from apiclient.discovery  import build
from httplib2 import Http
from oauth2client import file, client, tools
from oauth2client.contrib import gce
from apiclient.http import MediaFileUpload
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import os
import pathlib
from pathlib import Path
import io
import glob
import itertools
import shutil
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import (
    LineChart,
    Reference,
    Series
)

import random
import time
from time import sleep
import datetime
import csv


CLIENT_SECRET = "../secret/client_secret.json"
FILE_MASTERLIST = '../data/covid-19_IT.xlsx'
FILE_SUMMARY = '../data/generated_summary.xlsx'
OUTPUT_DIRECTORY = 'Output/'


# --------------------------------
# GDrive API: GDrive Authorization
# --------------------------------

SCOPES='https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
    creds = tools.run_flow(flow, store)
SERVICE = build('drive', 'v3', http=creds.authorize(Http()))
SS_SERVICE = build('sheets', 'v4', http=creds.authorize(Http()))


PARENT_FOLDER = '1MtOHnLPRJfWJM28DyFjhNN7ef8VUmTnn'


# ------------------------------------
# GDrive API: Check if Filename exists
# ------------------------------------
def fileInGDrive(filename):
    results = SERVICE.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and name='"+filename+"' and trashed = false and parents in '"+PARENT_FOLDER+"'",fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if items:
        return True
    else:
        return False


# ---------------------------------------
# GDrive API: Upload files to Google Drive
# ---------------------------------------
def writeToGDrive(filename,source,folder_id):
    file_metadata = {'name': filename,'parents': [folder_id],
    'mimeType': 'application/vnd.google-apps.spreadsheet'}
    media = MediaFileUpload(source,
                            mimetype='application/vnd.ms-excel')

    if fileInGDrive(filename) is False:
        file = SERVICE.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        print('Upload Success!')
        print('File ID:', file.get('id'))
        return file.get('id')

    else:
        print('File already exists as', filename)


def main():
    writeToGDrive('covid-19_IT.xls', FILE_MASTERLIST, PARENT_FOLDER)
    writeToGDrive('generated_summary.xlsx', FILE_SUMMARY, PARENT_FOLDER)


if __name__ == '__main__':
    main()
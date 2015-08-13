# -*- coding: utf-8 -*-
"""
Test the "backup" function, which saves sheet data to file.
"""
import sheetsync
import time, os

CLIENT_ID = os.environ['SHEETSYNC_CLIENT_ID']  
CLIENT_SECRET = os.environ['SHEETSYNC_CLIENT_SECRET']

TESTS_FOLDER_KEY = os.environ.get("SHEETSYNC_FOLDER_KEY")
SHEET_TO_BE_BACKED_UP = "1-HpLBDvGS5V8pIXR9GSqseeciWyy41I1uyLhzyAjPq4"

def test_backup():
    print ('setup_function: Retrieve OAuth2.0 credentials.')
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')


    print ('Open a spreadsheet, back it up to named file.')
    target = sheetsync.Sheet(credentials=creds,
                             document_key = SHEET_TO_BE_BACKED_UP,
                             worksheet_name = 'Simpsons',
                             key_column_headers = ['Character'],
                             header_row_ix=1)

    backup_name = 'backup test: %s' % int(time.time())
    backup_key = target.backup(backup_name, folder_name="sheetsync backups")

    backup_sheet = sheetsync.Sheet(credentials=creds,
                                   document_key = backup_key,
                                   worksheet_name = 'Simpsons',
                                   key_column_headers = ['Character'],
                                   header_row_ix=1)

    backup_data = backup_sheet.data() 
    assert "Bart Simpson" in backup_data
    assert backup_data["Bart Simpson"]["Voice actor"] == "Nancy Cartwright"

    print ('teardown_function Delete test spreadsheet')
    backup_sheet.drive_service.files().delete(fileId=backup_sheet.document_key).execute()

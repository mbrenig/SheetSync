# -*- coding: utf-8 -*-
"""
Test the "backup" function, which saves sheet data to file.
"""
import sheetsync
import time, os

GOOGLE_U = os.environ.get("SHEETSYNC_GOOGLE_ACCOUNT")
GOOGLE_P = os.environ.get("SHEETSYNC_GOOGLE_PASSWORD")

TESTS_FOLDER_KEY = os.environ.get("SHEETSYNC_FOLDER_KEY")
SHEET_TO_BE_BACKED_UP = "1-HpLBDvGS5V8pIXR9GSqseeciWyy41I1uyLhzyAjPq4"

def test_backup():
    print ('Open a spreadsheet, back it up to named file.')
    target = sheetsync.Sheet(GOOGLE_U,
                             GOOGLE_P,
                             document_key = SHEET_TO_BE_BACKED_UP,
                             sheet_name = 'Simpsons',
                             key_column_headers = ['Character'],
                             header_row_ix=1)

    backup_name = 'backup test: %s' % int(time.time())
    backup_key = target.backup(backup_name, folder_name="sheetsync backups")

    backup_sheet = sheetsync.Sheet(GOOGLE_U, GOOGLE_P,
                                   document_key = backup_key,
                                   sheet_name = 'Simpsons',
                                   key_column_headers = ['Character'],
                                   header_row_ix=1)

    backup_data = backup_sheet.data() 
    assert "Bart Simpson" in backup_data
    assert backup_data["Bart Simpson"]["Voice actor"] == "Nancy Cartwright"

    bckp_rsrc = backup_sheet.docs_client.get_resource_by_id(backup_sheet.document_key)
    backup_sheet.docs_client.Delete(bckp_rsrc)


# -*- coding: utf-8 -*-
"""
Test advanced CRUD features.
"""
import sheetsync
import time, os

# TODO: Use this: http://stackoverflow.com/questions/22574109/running-tests-with-api-authentication-in-travis-ci-without-exposing-api-password

CLIENT_ID = os.environ['SHEETSYNC_CLIENT_ID']  
CLIENT_SECRET = os.environ['SHEETSYNC_CLIENT_SECRET']
# Optional folder_key that all spreadsheets, and folders, will be created in.
TESTS_FOLDER = os.environ.get("SHEETSYNC_FOLDER_KEY")

# Template hosted by a dedicated "sheetsync" account that Mark set up.
TEMPLATE_DOC = "1-jjFDO11zLo6i6vpL7LbdhNMUJTAnjYVhUT7ZHMMtKQ"

target = None

def setup_function(function):
    global target
    print ('setup_function: Retrieve OAuth2.0 credentials.')
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')


    print ('setup_function: Create test spreadsheet.')
    # Copy the template spreadsheet into the prescribed folder.
    new_doc_name = '%s %s' % (__name__, int(time.time()))
    target = sheetsync.Sheet(credentials=creds,
                             document_name = new_doc_name,
                             worksheet_name = "Arsenal",
                             folder_key = TESTS_FOLDER,
                             template_key = TEMPLATE_DOC,
                             key_column_headers = ["No."],
                             header_row_ix=2,
                             formula_ref_row_ix=1,
                             protected_fields=["Apps","Goals"])


def teardown_function(function):
    print ('teardown_function Delete test spreadsheet')
    target.drive_service.files().delete(fileId=target.document_key).execute()


def test_protected_fields():
    print ('Update/Insert into a sheet w protected fields. check no overwrites.')
    updates = {"1" : {"Name" : "Jens Lehmann",
                       "Apps" : "0",
                       "Goals" : "0"},
                "4" : {"Name" : "Patrick V",
                       "Apps" : "100",
                       "Pos." : "??"},
                }

    def test_callback_fn(key_tuple, wks_row, raw_row, changed_fields):
        assert "Apps" not in changed_fields
        assert "Goals" not in changed_fields

    target.inject(updates, test_callback_fn)

    new_data = target.data()
    
    assert new_data["4"]["Name"] == "Patrick V"
    assert new_data["1"]["Apps"] == "54"

"""
def test_post_soft_delete():
    print ('See how .data and .sync handles soft deleted rows.')
    assert False

def test_logging_large_changes():
    print ('Test the truncate function.')
    assert False
"""

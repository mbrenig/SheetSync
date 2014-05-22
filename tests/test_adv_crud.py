# -*- coding: utf-8 -*-
"""
Test advanced CRUD features.
"""
import sheetsync
import time, os

# TODO: Use this: http://stackoverflow.com/questions/22574109/running-tests-with-api-authentication-in-travis-ci-without-exposing-api-password

GOOGLE_U = os.environ.get("SHEETSYNC_GOOGLE_ACCOUNT")
GOOGLE_P = os.environ.get("SHEETSYNC_GOOGLE_PASSWORD")
# Optional folder_key that all spreadsheets, and folders, will be created in.
TESTS_FOLDER = os.environ.get("SHEETSYNC_FOLDER_KEY")

# Template hosted by a dedicated "sheetsync" account that Mark set up.
TEMPLATE_DOC = "0AsrRHMfAlOZrdFlLLWlzM2dhZ0tyS1k5RUxmVGU3cEE"

target = None

def setup_function(function):
    global target
    print ('setup_function: Create test spreadsheet.')
    # Copy the template spreadsheet into the prescribed folder.
    new_doc_name = '%s %s' % (__name__, int(time.time()))
    target = sheetsync.Sheet(GOOGLE_U,
                             GOOGLE_P,
                             document_name = new_doc_name,
                             sheet_name = "Arsenal",
                             folder_key = TESTS_FOLDER,
                             template_key = TEMPLATE_DOC,
                             key_column_headers = ["No."],
                             header_row_ix=2,
                             formula_ref_row_ix=1,
                             protected_fields=["Apps","Goals"])


def teardown_function(function):
    print ('teardown_function Delete test spreadsheet')
    gdc = target._doc_client_pool[GOOGLE_U]
    target_rsrc = gdc.get_resource_by_id(target.document_key)
    gdc.Delete(target_rsrc)


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

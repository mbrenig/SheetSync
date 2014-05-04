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
TEMPLATE_DOC = ""

target = None

"""
TODO: Write these tests...

def setup_function(function):
    global target
    print ('setup_function: Create test spreadsheet.')
    # Copy the template spreadsheet into the prescribed folder.
    target = sheetsync.Sheet(GOOGLE_U,
                             GOOGLE_P,
                             title = ("test_%s" % int(time.time())),
                             folder_key = TESTS_FOLDER,
                             template_key = TEMPLATE_DOC,
                             sheet_name = "Arsenal",
                             header_row_ix=2,
                             key_field_headers = ["No."],
                             formula_ref_row_ix=1)


def teardown_function(function):
    print ('teardown_function Delete test spreadsheet')
    gdc = target._doc_client_pool[GOOGLE_U]
    target_rsrc = gdc.get_resource_by_id(target.document_key)
    gdc.Delete(target_rsrc)

def test_protected_fields():
    print ('Update/Insert into a sheet w protected fields. check no overwrites.')
    assert False

def test_post_soft_delete():
    print ('See how .data and .sync handles soft deleted rows.')
    assert False

def test_logging_large_changes():
    print ('Test the truncate function.')
    assert False
"""

# -*- coding: utf-8 -*-
"""
Various tests that cover CRUD with data that has keys that are not simple
strings. For instance:
    - Dates as keys (because Google can transpose them to date objects)
    - Integers as keys
    - Tuples as keys
    - Integers that start with 0, and some that don't. E.g. '01' to '20'
    - Not specifying the key_column_headers
"""

import sheetsync
import time, os

CLIENT_ID = os.environ['SHEETSYNC_CLIENT_ID']  
CLIENT_SECRET = os.environ['SHEETSYNC_CLIENT_SECRET']

TEMPLATE_DOC = "0AsrRHMfAlOZrdFlLLWlzM2dhZ0tyS1k5RUxmVGU3cEE"
TESTS_FOLDER_KEY = os.environ.get("SHEETSYNC_FOLDER_KEY")

target = None

"""
def setup_function(function):
    global target
    print ('setup_function: Create test spreadsheet.')
    # Copy the template spreadsheet into the prescribed folder.
    target = sheetsync.Sheet(GOOGLE_U,
                             GOOGLE_P,
                             title = ("test_%s" % int(time.time())),
                             folder_key = TESTS_FOLDER,
                             template_key = TEMPLATE_DOC,
                             sheet_name = "TEST",
                             header_row_ix=2,
                             key_column_headers = ["Initials"],
                             formula_ref_row_ix=1)


def teardown_function(function):
    print ('teardown_function Delete test spreadsheet')
    gdc = target._doc_client_pool[GOOGLE_U]
    target_rsrc = gdc.get_resource_by_id(target.document_key)
    gdc.Delete(target_rsrc)

def test_date_keys():
    print ('TODO: Test dates as keys.')
    assert True

def test_tuple_keys():
    print ('TODO: Test dates as keys.')
    assert True

def test_integers_keys():
    print ('TODO: Test dates as keys.')
    assert True

def test_tuple_mix_keys():
    print ('TODO: Test dates as keys.')
    assert True
"""

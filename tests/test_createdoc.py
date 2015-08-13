# -*- coding: utf-8 -*-
"""
Test document creation:
    - Create new blank spreadsheets.
    - Creating from a template copy.
    - Creating in an existing folder (and removing it from the root container!)
"""

import sheetsync
import time, os, sys

CLIENT_ID = os.environ['SHEETSYNC_CLIENT_ID']  
CLIENT_SECRET = os.environ['SHEETSYNC_CLIENT_SECRET']

TEMPLATE_DOC_KEY = "TODO"
TEMPLATE_DOC_NAME = "Interesting edge case?"
TESTS_FOLDER_KEY = os.environ.get("SHEETSYNC_FOLDER_KEY")
TESTS_FOLDER_NAME = "sheetsync testruns"

"""
def test_create_from_copy_template_key():
    print ('TODO: Create by copying template sheet.')
    assert False


def test_create_from_copy_template_name():
    print ('TODO: Create by copying template sheet.')
    assert False

"""

def test_move_to_folder_by_key():
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')

    new_doc_name = '%s-%s-%s' % (__name__, sys._getframe().f_code.co_name, int(time.time()))
    target = sheetsync.Sheet(credentials=creds,
                             document_name = new_doc_name,
                             worksheet_name = 'Sheet1',
                             folder_key = TESTS_FOLDER_KEY)
    # Delete the doc
    target.drive_service.files().delete(fileId=target.document_key).execute()


def test_move_to_folder_by_name():
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')

    new_doc_name = '%s-%s-%s' % (__name__, sys._getframe().f_code.co_name, int(time.time()))
    target = sheetsync.Sheet(credentials=creds,
                             document_name = new_doc_name,
                             worksheet_name = 'Sheet1',
                             folder_name = TESTS_FOLDER_NAME)
    # Delete the doc
    target.drive_service.files().delete(fileId=target.document_key).execute()


def test_move_to_new_folder_by_name():
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')

    new_doc_name = '%s-%s-%s' % (__name__, sys._getframe().f_code.co_name, int(time.time()))
    new_folder_name = 'sheetsync testrun %s' % int(time.time())
    target = sheetsync.Sheet(credentials=creds,
                             document_name = new_doc_name,
                             worksheet_name = 'Sheet1',
                             folder_name = new_folder_name)
    # Delete the doc
    target.drive_service.files().delete(fileId=target.document_key).execute()

    # Delete the new folder too..
    assert new_folder_name == target.folder['title']
    target.drive_service.files().delete(fileId=target.folder['id']).execute()


def test_the_kartik_test():
    # The most basic usage of creating a new sheet and adding data to it.
    # From April, Google defaults to using new-style sheets which requires
    # workarounds right now.
    creds = sheetsync.ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')

    new_doc_name = '%s-%s-%s' % (__name__, sys._getframe().f_code.co_name, int(time.time()))
    target = sheetsync.Sheet(credentials=creds,
                             document_name = new_doc_name)
    # Check we can sync data to the newly created sheet.
    data = {"1" : {"name" : "Gordon", "color" : "Green"},
            "2" : {"name" : "Thomas", "color" : "Blue" } }
    target.sync(data)
    
    retrieved_data = target.data()
    assert "1" in retrieved_data
    assert retrieved_data["1"]["name"] == "Gordon"
    assert "2" in retrieved_data
    assert retrieved_data["2"]["color"] == "Blue"
    assert retrieved_data["2"]["Key"] == "2"

    # Try opening the doc with a new instance (thereby guessing key columns)
    test_read = sheetsync.Sheet(credentials=creds,
                                document_name = new_doc_name)
    retrieved_data_2 = test_read.data()
    assert "1" in retrieved_data_2
    assert "2" in retrieved_data_2
    assert retrieved_data["2"]["color"] == "Blue"

    # Delete the doc
    target.drive_service.files().delete(fileId=target.document_key).execute()

# -*- coding: utf-8 -*-
"""
CRUD tests for row maniupulation.
"""
import sheetsync
import time, os

# TODO: Use this: http://stackoverflow.com/questions/22574109/running-tests-with-api-authentication-in-travis-ci-without-exposing-api-password

# TODO: Remove the need for the tests_folder key by automatically creating the
#       folder

GOOGLE_U = os.environ.get("SHEETSYNC_GOOGLE_ACCOUNT")
GOOGLE_P = os.environ.get("SHEETSYNC_GOOGLE_PASSWORD")
# Optional folder_key that all spreadsheets, and folders, will be created in.
TESTS_FOLDER = os.environ.get("SHEETSYNC_FOLDER_KEY")

# Template hosted by a dedicated "sheetsync" account that Mark set up.
# This contains half the 03/04 Arsenal squad. We will add to it, delete from it
# and more.
TEMPLATE_DOC = "0AsrRHMfAlOZrdFlLLWlzM2dhZ0tyS1k5RUxmVGU3cEE"

target = None

ARSENAL_0304 = {'1': {'Apps': '54',
       'Goals': '0',
       'Name': 'Jens Lehmann',
       'Nat.': ' GER',
       'Pos.': 'GK'},
 '10': {'Apps': '38',
        'Goals': '5',
        'Name': 'Dennis Bergkamp',
        'Nat.': ' NED',
        'Pos.': 'FW'},
 '11': {'Apps': '22',
        'Goals': '4',
        'Name': 'Sylvain Wiltord',
        'Nat.': ' FRA',
        'Pos.': 'FW'},
 '12': {'Apps': '47',
        'Goals': '0',
        'Name': 'Lauren',
        'Nat.': ' CMR',
        'Pos.': 'DF'},
 '14': {'Apps': '39',
        'Goals': '39',
        'Name': 'Thierry Henry',
        'Nat.': ' FRA',
        'Pos.': 'DF'},
 '15': {'Apps': '38',
        'Goals': '0',
        'Name': 'Ray Parlour',
        'Nat.': ' ENG',
        'Pos.': 'MF'},
 '16': {'Apps': '11',
        'Goals': '0',
        'Name': 'Giovanni van Bronckhorst',
        'Nat.': ' NED',
        'Pos.': 'MF'},
 '17': {'Apps': '48',
        'Goals': '7',
        'Name': 'Edu',
        'Nat.': ' BRA',
        'Pos.': 'MF'},
 '18': {'Apps': '24',
        'Goals': '0',
        'Name': 'Pascal Cygan',
        'Nat.': ' FRA',
        'Pos.': 'DF'},
 '19': {'Apps': '46',
        'Goals': '4',
        'Name': 'Gilberto Silva',
        'Nat.': ' BRA',
        'Pos.': 'MF'},
 '22': {'Apps': '22',
        'Goals': '0',
        'Name': 'Ga\xc3\xabl Clichy',
        'Nat.': ' FRA',
        'Pos.': 'DF'},
 '23': {'Apps': '50',
        'Goals': '1',
        'Name': 'Sol Campbell',
        'Nat.': ' ENG',
        'Pos.': 'DF'},
 '25': {'Apps': '24',
        'Goals': '3',
        'Name': 'Nwankwo Kanu',
        'Nat.': ' NGR',
        'Pos.': 'FW'},
 '27': {'Apps': '3',
        'Goals': '0',
        'Name': 'Efstathios Tavlaridis',
        'Nat.': ' GRE',
        'Pos.': 'DF'},
 '28': {'Apps': '55',
        'Goals': '3',
        'Name': 'Kolo Tour\xc3\xa9',
        'Nat.': ' CIV',
        'Pos.': 'DF'},
 '3': {'Apps': '47',
       'Goals': '1',
       'Name': 'Ashley Cole',
       'Nat.': ' ENG',
       'Pos.': 'DF'},
 '30': {'Apps': '15',
        'Goals': '4',
        'Name': 'J\xc3\xa9r\xc3\xa9mie Aliadi\xc3\xa8re',
        'Nat.': ' FRA',
        'Pos.': 'FW'},
 '32': {'Apps': '1',
        'Goals': '0',
        'Name': 'Michal Papadopulos',
        'Nat.': ' CZE',
        'Pos.': 'FW'},
 '33': {'Apps': '5',
        'Goals': '0',
        'Name': 'Graham Stack',
        'Nat.': ' IRL',
        'Pos.': 'GK'},
 '39': {'Apps': '8',
        'Goals': '1',
        'Name': 'David Bentley',
        'Nat.': ' ENG',
        'Pos.': 'MF'},
 '4': {'Apps': '44',
       'Goals': '3',
       'Name': 'Patrick Vieira',
       'Nat.': ' FRA',
       'Pos.': 'MF'},
 '45': {'Apps': '3',
        'Goals': '0',
        'Name': 'Justin Hoyte',
        'Nat.': ' ENG',
        'Pos.': 'DF'},
 '5': {'Apps': '15',
       'Goals': '0',
       'Name': 'Martin Keown',
       'Nat.': ' ENG',
       'Pos.': 'DF'},
 '51': {'Apps': '1',
        'Goals': '0',
        'Name': 'Frank Simek',
        'Nat.': ' USA',
        'Pos.': 'DF'},
 '52': {'Apps': '1',
        'Goals': '0',
        'Name': 'John Spicer',
        'Nat.': ' ENG',
        'Pos.': 'FW'},
 '53': {'Apps': '3',
        'Goals': '0',
        'Name': 'Jerome Thomas',
        'Nat.': ' ENG',
        'Pos.': 'MF'},
 '54': {'Apps': '3',
        'Goals': '0',
        'Name': 'Quincy Owusu-Abeyie',
        'Nat.': ' GHA',
        'Pos.': 'FW'},
 '55': {'Apps': '1',
        'Goals': '0',
        'Name': '\xc3\x93lafur Ingi Sk\xc3\xbalason',
        'Nat.': ' ISL',
        'Pos.': 'MF'},
 '56': {'Apps': '3',
        'Goals': '0',
        'Name': 'Ryan Smith',
        'Nat.': ' ENG',
        'Pos.': 'FW'},
 '57': {'Apps': '3',
        'Goals': '1',
        'Name': 'Cesc F\xc3\xa0bregas',
        'Nat.': ' ESP',
        'Pos.': 'MF'},
 '7': {'Apps': '51',
       'Goals': '19',
       'Name': 'Robert Pir\xc3\xa8s',
       'Nat.': ' FRA',
       'Pos.': 'MF'},
 '8': {'Apps': '44',
       'Goals': '10',
       'Name': 'Fredrik Ljungberg',
       'Nat.': ' SWE',
       'Pos.': 'MF'},
 '9': {'Apps': '21',
       'Goals': '5',
       'Name': 'Jos\xc3\xa9 Antonio Reyes',
       'Nat.': ' FRA',
       'Pos.': 'FW'}}

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
                             key_field_headers = ["No."],
                             header_row_ix=2,
                             formula_ref_row_ix=1)


def teardown_function(function):
    print ('teardown_function Delete test spreadsheet')
    gdc = target._doc_client_pool[GOOGLE_U]
    target_rsrc = gdc.get_resource_by_id(target.document_key)
    gdc.Delete(target_rsrc)

def test_insert_row():
    print ('Add enough rows to spreadsheet that it needs extending')
    target.sync(ARSENAL_0304)
    # Now get the data and check certain rows were added.
    raw_data = target.data()
    assert '57' in raw_data
    assert '10' in raw_data
    assert 'Goals per 100 apps' in raw_data['57']

def test_soft_delete():
    print ('Soft delete rows from the spreadsheet')
    raw_data = target.data()
    # Delete Giovanni
    assert "16" in raw_data
    del raw_data["16"]
    target.sync(raw_data)
    # Check new data has "16 (DELETED)" in it. 
    new_raw_data = target.data()
    assert "1" in new_raw_data
    assert "16" not in new_raw_data
    assert "16 (DELETED)" in new_raw_data
 
def test_full_delete():
    print ('Test hard deletes of rows from the spreadsheet.')
    raw_data = target.data()
    target.flag_delete_mode = False
    # Turn on full deletion.
    assert "16" in raw_data
    del raw_data["16"]
    target.sync(raw_data)
    # Check we deleted "16" for real.
    new_raw_data = target.data()
    assert "1" in new_raw_data
    assert "16" not in new_raw_data
    assert "16 (DELETED)" not in new_raw_data

def test_change_row():
    print ('Test changing multiple rows.')
    correct_data = {}
    correct_data.update(ARSENAL_0304)
    correct_data["14"]["Apps"] = 51
    correct_data["14"]["Pos."] = "FW"
    correct_data["9"]["Nat."] = "ENG"
    target.sync(correct_data)
    new_raw_data = target.data()
    assert new_raw_data["14"]["Apps"] == '51'
    assert new_raw_data["14"]["Pos."] == "FW"
    assert new_raw_data["9"]["Nat."] == "ENG"

def test_extend_header():
    print ('Add additional columns, include expanding the spreadsheet.')
    full_data = {}
    full_data.update(ARSENAL_0304)
    # Add a column.
    for row in full_data.itervalues():
        row["Club"] = "Arsenal"
    target.sync(full_data)
    new_raw_data = target.data()
    assert "Club" in target.header
    assert new_raw_data["1"]["Club"] == "Arsenal"

def test_inject_only():
    print ('Test injecting partial rows, check no deleting.')
    extra_players = { "57" : {}, "32" : {} }
    extra_players["57"].update( ARSENAL_0304["57"] )
    extra_players["32"].update( ARSENAL_0304["32"] )
    target.inject(extra_players)
    new_raw_data = target.data()
    assert "1" in new_raw_data
    assert "19" in new_raw_data
    assert "57" in new_raw_data
    assert "32" in new_raw_data
    assert new_raw_data["57"]["Name"] == 'Cesc F\xc3\xa0bregas'


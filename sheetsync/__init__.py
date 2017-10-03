# -*- coding: utf-8 -*-
"""
    sheetsync
    ~~~~~~~~~

    A library to synchronize data with a google spreadsheet, with support for:
        - Creating new spreadsheets. Including by copying template sheets.
        - Call-back functions when rows are added/updated/deleted.
        - Protected columns.
        - Extending formulas to new rows.

    :copyright: (c) 2014 by Mark Brenig-Jones.
    :license: MIT, see LICENSE.txt for more details.
"""

from version import __version__

import logging
import httplib2 # pip install httplib2
from datetime import datetime
import json

# import latest google api python client.
import apiclient.errors # pip install --upgrade google-api-python-client
import apiclient.discovery 
from oauth2client.client import OAuth2WebServerFlow, OAuth2Credentials, AccessTokenRefreshError

# import the excellent gspread library.
import gspread # pip install --upgrade gspread
from gspread import SpreadsheetNotFound, WorksheetNotFound

import dateutil.parser # pip install python-dateutil

logger = logging.getLogger('sheetsync')

MAX_BATCH_LEN = 500   # Google's limit is 1MB or 1000 batch entries.
DELETE_ME_FLAG = ' (DELETED)'
DEFAULT_WORKSHEET_NAME = 'Sheet1'

def ia_credentials_helper(client_id, client_secret, 
                          credentials_cache_file="credentials.json",
                          cache_key="default"):
    """Helper function to manage a credentials cache during testing.

    This function attempts to load and refresh a credentials object from a
    json cache file, using the cache_key and client_id as a lookup.

    If this isn't found then it starts an OAuth2 authentication flow, using
    the client_id and client_secret and if successful, saves those to the local
    cache. See :ref:`helper`.
  
    Args:
        client_id (str): Google Drive API client id string for an installed app
        client_secret (str): The corresponding client secret.
        credentials_cache_file (str): Filepath to the json credentials cache file
        cache_key (str): Optional string to allow multiple credentials for a client
           to be stored in the cache.

    Returns:
        OAuth2Credentials: A google api credentials object. As described here:
        https://developers.google.com/api-client-library/python/guide/aaa_oauth

    """
    def _load_credentials(key):
        with open(credentials_cache_file, 'rb') as inf:
            cache = json.load(inf)
        cred_json = cache[key]
        return OAuth2Credentials.from_json(cred_json)

    def _save_credentials(key, credentials):
        cache = {}
        try:
            with open(credentials_cache_file, 'rb') as inf:
                cache = json.load(inf)
        except (IOError, ValueError), e:
            pass
        cache[key] = credentials.to_json()
        with open(credentials_cache_file, 'wb') as ouf:
            json.dump(cache, ouf)

    credentials_key = "%s/%s/%s" % (client_id, client_secret, cache_key)
    try:
        credentials = _load_credentials(credentials_key)
        if credentials.access_token_expired:
            http = httplib2.Http()
            credentials.refresh(http)
    except (IOError, 
            ValueError, 
            KeyError, 
            AccessTokenRefreshError), e:
        # Check https://developers.google.com/drive/scopes for all available scopes
        OAUTH_SCOPE = ('https://www.googleapis.com/auth/drive '+
                       'https://spreadsheets.google.com/feeds')
        # Redirect URI for installed apps
        REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

        # Run through the OAuth flow and retrieve credentials
        flow = OAuth2WebServerFlow(client_id, client_secret, OAUTH_SCOPE,
                                   redirect_uri=REDIRECT_URI)
        authorize_url = flow.step1_get_authorize_url()
        print('Go to the following link in your browser:\n' + authorize_url)
        code = raw_input('Enter verification code: ').strip()
        credentials = flow.step2_exchange(code)

    _save_credentials(credentials_key, credentials)
    return credentials



def _is_google_fmt_date(text):
    frags = text.split('/')
    if len(frags) != 3:
        return False
    if not all(frag.isdigit() for frag in frags):
        return False
    frags[0] = frags[0].zfill(2)
    frags[1] = frags[1].zfill(2)
    try:
        date = datetime.strptime('-'.join(frags), '%m-%d-%Y')
        return date
    except:
        return False

def google_equivalent(text1, text2):
    # Google spreadsheets modify some characters, and anything that looks like
    # a date. So this function will return true if text1 would equal text2 if 
    # both were input into a google cell.
    lines1 = [l.replace('\t',' ').strip() for l in text1.splitlines()]
    lines2 = [l.replace('\t',' ').strip() for l in text2.splitlines()]
    if len(lines1) != len(lines2):
        return False
    if len(lines1) == 0 and len(lines2) == 0:
        return True
    if len(lines1) > 1:
        for a,b in zip(lines1, lines2):
            if a != b:
                return False
        # Multiline string matches on every line.
        return True
    elif lines1[0] == lines2[0]:
        # Single line string.. that matches.
        return True

    # Might be dates.
    text1 = lines1[0]
    text2 = lines2[0]
    if _is_google_fmt_date(text1) or _is_google_fmt_date(text2):
        try:
            date1 = dateutil.parser.parse(text1)
            date2 = dateutil.parser.parse(text2)
            if date1 == date2:
                return True
        except ValueError:
            # Couldn't parse one of the dates.
            pass
    return False

class MissingSheet(Exception):
    pass

class CorruptHeader(Exception):
    pass

class BadDataFormat(Exception):
    pass

class DuplicateRows(Exception):
    pass

class UpdateResults(object):
    """ A lightweight counter object that holds statistics about number of
    updates made after using the 'sync' or 'inject' method. 
    
    Attributes:
      added (int): Number of rows added
      changed (int): Number of rows changed                
      nochange (int): Number of rows that were not modified.
      deleted (int): Number of rows deleted (which will always be 0 when using
          the 'inject' function)
    """
    def __init__(self):
        self.added = 0
        self.changed = 0
        self.deleted = 0
        self.nochange = 0

    def __str__(self):
        r = 'Added: %s Changed: %s Deleted: %s No Change: %s' % (
                    self.added, self.changed, self.deleted, self.nochange)
        return r

class Row(dict):
    def __init__(self, row_num):
        self.row_num = row_num
        self.db = {}
        dict.__init__(self)

    def __setitem__(self, key, cell):
        dict.__setitem__(self, key, cell)
        self.db[key] = cell.value

    def cell_list(self):
        for cell in self.itervalues():
            yield cell

    def is_empty(self):
        return all((val is None or val == '') for val in self.db.itervalues())

class Header(object):
    def __init__(self):
        self.col_to_header = {}
        self.header_to_col = {}

    def reset(self):
        self.col_to_header = {}
        self.header_to_col = {}

    def col_lookup(self, col_ix):
        return self.col_to_header.get(col_ix)

    def header_lookup(self, header):
        return self.header_to_col.get(header)

    def set(self, col, header):
        if header in self.header_to_col and self.header_to_col[header] != col:
            raise CorruptHeader("'%s' was found twice in header row" % header)
        if col in self.col_to_header and self.col_to_header[col] != header:
            raise CorruptHeader("Header for column '%s' changed while running" % col)
        self.col_to_header[col] = header
        self.header_to_col[header] = col

    @property
    def headers_in_order(self):
        col_header_list = self.col_to_header.items()
        col_header_list.sort(lambda x,y: cmp(x[0],y[0]))
        return [header for col,header in col_header_list]

    @property
    def last_column(self):
        if self.col_to_header:
            return max(self.col_to_header.keys())
        return 0

    @property
    def first_column(self):
        if self.col_to_header:
            return min(self.col_to_header.keys())
        return 0

    @property
    def columns(self):
        return self.col_to_header.keys()

    def __contains__(self, header):
        return (header in self.header_to_col)



class Sheet(object):
    """ Represents a single worksheet within a google spreadsheet.
    
    This class tracks the google connection, the reference to the worksheet, as
    well as options controlling the structure of the data in the worksheet.. for
    example:
        * Which row is used as the table header
        * What header names should be used for the key column(s)
        * Whether some columns are protected from overwriting
   
    Attributes:
       document_key (str): The spreadsheet's document key assigned by google 
           drive. If you are using sheetsync to create a spreadsheet then use 
           this attribute to saved the document_key, and make sure you pass 
           it as a parameter in subsequent calls to __init__
       document_name (str): The title of the google spreadsheet document
       document_href (str): The HTML href for the google spreadsheet document
    """
    def __init__(self, 
                 credentials=None,
                 document_key=None, document_name=None,
                 worksheet_name=None,
                 # Behavior modifiers
                 key_column_headers=None, 
                 header_row_ix=1,
                 formula_ref_row_ix=None,
                 flag_deletes=True,
                 protected_fields=None,
                 # Document creation behavior
                 template_key=None, template_name=None,
                 folder_key=None, folder_name=None):
        """Creates a worksheet object (also creating a new Google sheet doc if required)

        Args:
            credentials (OAuth2Credentials): Credentials object returned by the
                google authorization server. Described in detail in this article:
                https://developers.google.com/api-client-library/python/guide/aaa_oauth
                For testing and development consider using the `ia_credentials_helper`
                helper function
            document_key (Optional) (str): Document key for the existing spreadsheet to
                sync data to. More info here:
                https://productforums.google.com/forum/#!topic/docs/XPOR9bTTS50
                If this is not provided sheetsync will use document_name to try and
                find the correct spreadsheet.
            document_name (Optional) (str): The name of the spreadsheet document to 
                access. If this is not found it will be created. If you know
                the document_key then using that is faster and more reliable.
            worksheet_name (str): The name of the worksheet inside the spreadsheet
                that data will be synced to. If omitted then the default name
                "Sheet1" will be used, and a matching worksheet created if
                necessary.
            key_column_headers (Optional) (list of str): Data in the key column(s) uniquely
                identifies a row in your data. So, for example, if your data is 
                indexed by a single username string, that you want to store in a
                column with the header 'Username', you would pass this:
                    key_column_headers=['Username']
                However, sheetsync also supports component keys. Python dictionaries can
                use tuples as keys, for example if you had a tuple key like
                this:
                    ('Tesla', 'Model-S', '2013')
                You can make the column meanings clear by passing in a list of
                three key_column_headers:
                    ['Make', 'Model', 'Year']
                If no value is given, then the default behavior is to name the
                column "Key"; or "Key-1", "Key-2", ... if your data dictionaries 
                keys are tuples.
            header_row_ix (Optional) (int): The row number we expect to see column headers
                in. Defaults to 1 (the very top row).
            formula_ref_row_ix (Optional) (int): If you want formulas to be added to some
                cells when inserting new rows then use a formula reference row. 
                See :ref:`formulas` for an example use.
            flag_deletes (Optional) (bool): Specify if deleted rows should only be flagged
                for deletion. By default sheetsync does not delete rows of data, it
                just marks that they are deleted by appending the string 
                " (DELETED)" to key values. If you pass in the value "False" then
                rows of data will be deleted by the sync method if they are not
                found in the input data. Note, use the inject method if you only want
                to add or modify data to in a worksheet.
            protected_fields (Optional) (list of str): An list of fields (column 
                headers) that contain protected data. sheetsync will only write to 
                cells in these columns if they are blank. This can be useful if you
                are expecting users of the spreadsheet to colaborate on the document
                and edit values in certain columns (e.g. modifying a "Test result" 
                column from "PENDING" to "PASSED") and don't want to overwrite
                their edits.
            template_key (Optional) (str): This optional key references the spreadsheet 
                that will be copied if a new spreadsheet needs to be created. 
                This is useful for copying over formatting, a specific header 
                order, or apps-script functions. See :ref:`templates`.
            template_name (Optional) (str): As with template_key but the name of the
                template spreadsheet. If known, using the template_key will be
                faster.
            folder_key (Optional) (str): This optional key references the folder that a new
                spreadsheet will be moved to if a new spreadsheet needs to be
                created.
            folder_name (Optional) (str): Like folder_key this parameter specifies the
                optional folder that a spreadsheet will be created in (if
                required). If a folder matching the name cannot be found, sheetsync
                will attempt to create it.
 
        """

        # Record connection settings, and create a connection.
        self.credentials = credentials
        self._drive_service = None
        self._gspread_client = None
        self._sheet = None              # Gspread sheet instance.
        self._worksheet = None          # Gspread worksheet instance.

        # Find or create the Google spreadsheet document
        document = self._find_document(document_key, document_name)
        if document is None:
            if document_name is None:
                raise ValueError("Must specify a document_name")
            # We need to create the document
            template = self._find_document(template_key, template_name) 
            if template is None and template_name is not None:
                raise ValueError("Could not find template: %s" % template_name)
            self.folder = self._find_or_create_folder(folder_key, folder_name)
            document = self._create_new_or_copy(source_doc=template, 
                                                target_name=document_name, 
                                                folder=self.folder)
            if not document:
                raise Exception("Could not create doc '%s'." % document_name)
  
        # Create attribute for document key
        self.document_key = document['id']
        self.document_name = document['title']
        self.document_href = document['alternateLink']

        # Find or create the worksheet
        if worksheet_name is None:
            logger.info("Using the default worksheet name")
            worksheet_name = DEFAULT_WORKSHEET_NAME
        self.worksheet_name = worksheet_name

        # Store off behavioural settings for interacting with the worksheet.
        if key_column_headers is None:
            logger.info("No key column names. Will use 'Key'; or 'Key-1', 'Key-2' etc..")
            key_column_headers = []

        self.key_column_headers = key_column_headers
        self.header_row_ix = header_row_ix
        self.formula_ref_row_ix = formula_ref_row_ix
        self.flag_delete_mode = flag_deletes
        self.protected_fields = (protected_fields or [])
 
        # Cache batch operations to write efficiently
        self._batch_request = None
        self._batch_href = None

        # Track headers and reference formulas
        self.header = Header()
        self._get_or_create_headers()
        self.header_to_ref_formula = {}
        self.read_ref_formulas()


    @property
    def sheet(self):
        # Finds and returns a gspread.Spreadsheet object
        if self._sheet:
            return self._sheet
        self._sheet = self.gspread_client.open_by_key(self.document_key)
        return self._sheet

    @property
    def worksheet(self):
        # Finds (or creates) then returns a gspread.Worksheet object 
        if self._worksheet:
            return self._worksheet
        try:
            try:
                self._worksheet = self.sheet.worksheet(self.worksheet_name)
            except WorksheetNotFound:
                logger.info("Not found. Creating worksheet '%s'", self.worksheet_name)
                self._worksheet = self.sheet.add_worksheet(title=self.worksheet_name, 
                                                           rows=20, cols=10)
        except Exception, e:
            logger.exception("Failed to find or create worksheet: %s. %s", 
                                                      self.worksheet_name, e)
            raise e
        return self._worksheet

    @property
    def gspread_client(self):
        if self._gspread_client:
            return self._gspread_client
        self._gspread_client = gspread.authorize(self.credentials)
        self._gspread_client.login()
        return self._gspread_client 

    @property
    def drive_service(self):
        if self._drive_service:
            return self._drive_service

        http = httplib2.Http()
        if self.credentials.access_token_expired:
            logger.info('Refreshing expired credentials')
            self.credentials.refresh(http)

        logger.info('Creating drive service')
        http = self.credentials.authorize(http)
        drive_service = apiclient.discovery.build('drive', 'v2', http=http)
        # Cache the drive_service object for future calls. 
        self._drive_service = drive_service 
        return drive_service


    def _create_new_or_copy(self, 
                            target_name=None,
                            source_doc=None, 
                            folder=None,
                            sheet_description="new"):
        if target_name is None:
            raise KeyError("Must specify a name for the new document")

        body = {'title': target_name }
        if folder:
            body['parents'] = [{'kind' : 'drive#parentReference',
                                'id' : folder['id'],
                                'isRoot' : False }]

        drive_service = self.drive_service
        if source_doc is not None:
            logger.info("Copying spreadsheet.")
            try:
                print body
                print source_doc['id']
                new_document = drive_service.files().copy(fileId=source_doc['id'], body=body).execute()
            except Exception, e:
                logger.exception("gdata API error. %s", e)
                raise e

        else:
            # Create new blank spreadsheet.
            logger.info("Creating blank spreadsheet.")
            body['mimeType'] = 'application/vnd.google-apps.spreadsheet'
            try:
                new_document = drive_service.files().insert(body=body).execute()
            except Exception, e:
                logger.exception("gdata API error. %s", e)
                raise e

        logger.info("Created %s spreadsheet with ID '%s'", 
                sheet_description,
                new_document.get('id'))

        return new_document

    def _find_or_create_folder(self, folder_key=None, folder_name=None):
        drive_service = self.drive_service
        # Search by folder key.. raise Exception if not found.
        if folder_key is not None:
            try:
                folder_rsrc = drive_service.files().get(fileId=folder_key).execute()
            except apiclient.errors.HttpError, e:
                # XXX: WRONG... probably returns 404 if not found,.. which is not an error.
                logger.exception("Google API error: %s", e)
                raise e

            if not folder_rsrc:
                raise KeyError("Folder with key %s was not found." % folder_key)
            return folder_rsrc

        if not folder_name:
            return None

        # Search by folder name.
        try:
            name_query = drive_service.files().list(
                q=("title='%s' and trashed=false and "
                   "mimeType='application/vnd.google-apps.folder'") % 
                        folder_name.replace("'","\\'")
                        ).execute()
            items = name_query['items']
            if len(items) == 1:
                return items[0]
            elif len(items) > 1:
                raise KeyError("%s folders found named: %s" % (len(items), folder_name))
        except Exception, e:
            logger.exception("Google API error. %s", e)
            raise e

        logger.info("Creating a new folder named: '%s'", folder_name)
        try:
            new_folder_rsrc = drive_service.files().insert(
                body={ 'mimeType' : 'application/vnd.google-apps.folder',
                       'title' : folder_name }).execute()
        except Exception, e:
            logger.exception("Google API error. %s", e)
            raise e

        return new_folder_rsrc


    def _find_document(self, doc_key=None, doc_name=None):
        # Find the document by key and raise "KeyError" if not found.
        # Otherwise search by document_name
        drive_service = self.drive_service
        if doc_key is not None:
            logger.debug("Finding document by key.")
            try:
                doc_rsrc = drive_service.files().get(fileId=doc_key).execute()
            except Exception, e:
                logger.exception("gdata API error. %s", e)
                raise e

            if doc_rsrc is None:
                raise KeyError("Could not find document with key: %s" % doc_key)
            return doc_rsrc

        if doc_name is None:
            return None

        try:
            name_query = drive_service.files().list(
                q=("title='%s' and trashed=false and "
                   "mimeType='application/vnd.google-apps.spreadsheet'") %
                        doc_name.replace("'","\\'")
                        ).execute()
            matches = name_query['items']
        except Exception, e:
            logger.exception("gdata API error. %s", e)
            raise e

        if len(matches) == 1:
            return matches[0]

        if len(matches) > 1:
            raise KeyError("Too many matches for doc named '%s'" % doc_name)
        
        return None


    def _extends(self, rows=None, columns=None):
        # Resizes the sheet if needed, to match the given
        # number of rows and/or columns
        new_rows, new_cols = None, None
        if rows is not None and rows > self.worksheet.row_count:
            # Need to add new rows to the spreadsheet.
            new_rows = rows
        if columns is not None and columns > self.worksheet.col_count:
            new_cols = columns

        if new_rows or new_cols:
            try:
                self.worksheet.resize(rows=new_rows, cols=new_cols)
            except Exception, e:
                logger.exception("Error resizing worksheet. %s", e)
                raise e


    def _write_cell(self, cell):
        # Creates a batch_update if required, and adds the passed cell
        # to it. Then tests if a flush_writes call is required (when the
        # batch write might be close to the 1MB limit)
        if not self._batch_request: 
            self._batch_request = []

        logger.debug("_write_cell: Adding batch update")
        self._batch_request.append(cell)

        if len(self._batch_request) > MAX_BATCH_LEN:
            self._flush_writes()


    def _flush_writes(self):
        # Write current batch_updates to google sheet.
        if self._batch_request:
            logger.info("_flush_writes: Writing %s cell writes",
                                        len(self._batch_request))
            try:
                self.worksheet.update_cells(self._batch_request)
            except Exception, e:
                logger.exception("gdata API error. %s", e)
                raise e

            # Now check the response code. 
            #for entry in resp.entry:
            #    if entry.batch_status.code != '200':
            #        error = "gdata API error. %s - %s" % (entry.batch_status.code,
            #                                entry.batch_status.reason)
            #        logger.error("Batch update failed: %s", error)
            #        raise Exception(error)

            self._batch_request = []


    def _cell_feed(self, row=None, max_row=None, further_rows=False,        # XXX: REFACTOR
                         col=None, max_col=None, further_cols=False,
                         return_empty=False):

        # Fetches cell data for a given row, and all following rows if 
        # further_rows is True. If no row is given, all cells are returned.
        params = {}
        if row is not None:
            params['min-row'] = str(row)
            if max_row is not None:
                params['max-row'] = str(max_row)
            elif not further_rows:
                params['max-row'] = str(row)

        if col is not None:
            params['min-col'] = str(col)
            if max_col is not None:
                params['max-col'] = str(max_col)
            elif not further_cols:
                params['max-col'] = str(col)

            if (params['min-col'] == params['max-col'] and
                params['min-col'] == '0'):
                return []

        if return_empty:
            params['return-empty'] = "true"

        logger.info("getting cell feed")
        try:
            feed = self.gspread_client.get_cells_feed(self.worksheet, params=params)
            # Bit of a hack to rip out Gspread's xml parsing.
            cfeed = [gspread.Cell(self, elem) for elem in
                                        feed.findall(gspread.client._ns('entry'))]
        except Exception, e:
            logger.exception("gspread error. %s", e)
            raise e

        return cfeed

    def read_ref_formulas(self):
        self.header_to_ref_formula = {}

        if self.formula_ref_row_ix:
            for cell in self._cell_feed(row=self.formula_ref_row_ix):
                ref_formula = cell.input_value
                header = self.header.col_lookup(cell.col)
                if header and ref_formula.startswith("="):
                    self.header_to_ref_formula[header] = ref_formula


    def _get_or_create_headers(self, required_headers=[]):
        # Reads the header row, adds missing headers if required.
        self.header.reset()

        for cell in self._cell_feed(row=self.header_row_ix):
            self.header.set(cell.col, cell.value)
 
        headers_to_add = []
        for key_field in self.key_column_headers:
            if key_field not in self.header:
                headers_to_add.append(key_field)

        # Write new headers in alphabetical order.
        sorted_required_headers = list(required_headers)
        sorted_required_headers.sort()
        for header in sorted_required_headers:
            if header not in self.header:
                headers_to_add.append(header)

        if not headers_to_add:
            return 

        target_cols = self.header.last_column + len(headers_to_add)
        self._extends(columns=target_cols)

        cells_list = self._cell_feed(row=self.header_row_ix, 
                                     return_empty=True)
        for cell in cells_list:
            if not headers_to_add:
                break
            if not cell.value:
                header = headers_to_add.pop(0)
                cell.value = header
                self.header.set(cell.col, header)
                self._write_cell(cell)

        if headers_to_add:
            raise CorruptHeader("Error adding headers")

        self._flush_writes()

    def backup(self, backup_name, folder_key=None, folder_name=None):
        """Copies the google spreadsheet to the backup_name and folder specified. 
        
        Args:
          backup_name (str): The name of the backup document to create.

          folder_key (Optional) (str): The key of a folder that the new copy will
            be moved to.
         
          folder_name (Optional) (str): Like folder_key, references the folder to move a
            backup to. If the folder can't be found, sheetsync will create it.

        """

        folder = self._find_or_create_folder(folder_key, folder_name)
        drive_service = self.drive_service
        try:
            source_rsrc = drive_service.files().get(fileId=self.document_key).execute()
        except Exception, e:
            logger.exception("Google API error. %s", e)
            raise e

        backup = self._create_new_or_copy(source_doc=source_rsrc, 
                                        target_name=backup_name, 
                                        folder=folder,
                                        sheet_description="backup")

        backup_key = backup['id']
        return backup_key
         
    def _yield_rows(self, cells_feed):
        cur_row = None
        for cell in cells_feed:
            if cell.row <= self.header_row_ix:
                # Never yield the header from this function to avoid overwrites
                continue
            if self.formula_ref_row_ix and cell.row == self.formula_ref_row_ix:
                # Never yield the formula ref row to avoid overwrites
                continue
            if cur_row is None or cur_row.row_num != cell.row:
                if cur_row is not None:
                    yield cur_row
                # Make a new row.
                cur_row = Row(cell.row)
            if cell.col in self.header.columns:
                cur_row[self.header.col_lookup(cell.col)] = cell

        if cur_row is not None:
            yield cur_row

    def data(self, as_cells=False):
        """ Reads the worksheet and returns an indexed dictionary of the
        row objects.
        
        For example:

        >>>print sheet.data()

        {'Miss Piggy': {'Color': 'Pink', 'Performer': 'Frank Oz'}, 'Kermit': {'Color': 'Green', 'Performer': 'Jim Henson'}} 
        
        """
        sheet_data = {}
        self.max_row = max(self.header_row_ix, self.formula_ref_row_ix)
        all_cells = self._cell_feed(row=self.max_row+1,
                                    further_rows=True,
                                    col=self.header.first_column,
                                    max_col=self.header.last_column,
                                    return_empty=True)

        for wks_row in self._yield_rows(all_cells):
            if wks_row.row_num not in sheet_data and not wks_row.is_empty():
                sheet_data[wks_row.row_num] = wks_row

        all_rows = sheet_data.keys()
        if all_rows:
            self.max_row = max(all_rows)

        # Now index by key_tuple
        indexed_sheet_data = {}
        for row, wks_row in sheet_data.iteritems():
            # Make the key tuple
            if len(self.key_column_headers) == 0:
                # Are there any default key column headers?
                if "Key" in wks_row:
                    logger.info("Assumed key column's header is 'Key'")
                    self.key_column_headers = ['Key']
                elif "Key-1" in wks_row:
                    self.key_column_headers = [h for h in wks_row.keys() 
                        if h.startswith("Key-") and h.split("-")[1].isdigit()]
                    logger.info("Assumed key column headers were: %s",
                                self.key_column_headers)
                else:
                    raise Exception("Unable to read spreadsheet. Specify"
                        "key_column_headers when initializing Sheet object.")

            key_list = []
            for key_hdr in self.key_column_headers:
                key_val = wks_row.db.get(key_hdr,"")
                if key_val.startswith("'"):
                    key_val = key_val[1:]
                key_list.append(key_val)
            key_tuple = tuple(key_list)
            if all(k == "" for k in key_tuple):
                continue

            if as_cells:
                indexed_sheet_data[key_tuple] = wks_row
            else:
                if len(key_tuple) == 1:
                    key_tuple = key_tuple[0]
                indexed_sheet_data[key_tuple] = wks_row.db

        return indexed_sheet_data

    @property
    def key_length(self):
        return len(self.key_column_headers)

    #--------------------------------------------------------------------------
    # Update the worksheet to match the raw_data, calling
    # the row_change_callback for any adds/deletes/fieldchanges.
    #
    # Read the data to build a list of required headers and 
    # check the keys are valid tuples.
    # sync and update.
    #--------------------------------------------------------------------------
    def sync(self, raw_data, row_change_callback=None):
        """ Equivalent to the inject method but will delete rows from the
        google spreadsheet if their key is not found in the input (raw_data) 
        dictionary.
    
        Args:
            raw_data (dict): See inject method
            row_change_callback (Optional) (func): See inject method

        Returns:
            UpdateResults (object): See inject method
        """
        return self._update(raw_data, row_change_callback, delete_rows=True)

    def inject(self, raw_data, row_change_callback=None):
        """ Use this function to add rows or update existing rows in the
        spreadsheet.
    
        Args: 
          raw_data (dict): A dictionary of dictionaries. Where the keys of the
             outer dictionary uniquely identify each row of data, and the inner
             dictionaries represent the field,value pairs for a row of data.
   
          row_change_callback (Optional) (func): A callback function that you
             can use to track changes to rows on the spreadsheet. The
             row_change_callback function must take four parameters like so:

             change_callback(row_key, 
                             row_dict_before, 
                             row_dict_after, 
                             list_of_changed_keys)

        Returns:
          UpdateResults (object): A simple counter object providing statistics
            about the changes made by sheetsync.
        """
        return self._update(raw_data, row_change_callback, delete_rows=False)

    def _update(self, raw_data, row_change_callback=None, delete_rows=False):
        required_headers = set()
        logger.debug("In _update. Checking for bad keys and missing headers")
        fixed_data = {}
        missing_raw_keys = set()
        for key, row_data in raw_data.iteritems():
            if not isinstance(key, tuple):
                key = (str(key),)
            else:
                key = tuple([str(k) for k in key])

            if len(self.key_column_headers) == 0:
                # Pick default key_column_headers.
                if len(key) == 1:
                    self.key_column_headers = ["Key"]
                else:
                    self.key_column_headers = ["Key-%s" % i for i in range(1,len(key)+1)]

            # Cast row_data values to unicode strings.
            fixed_data[key] = dict([(k,unicode(v)) for (k,v) in row_data.items()])

            missing_raw_keys.add(key)
            if len(key) != self.key_length:
                raise BadDataFormat("Key %s does not match key field headers %s" % (key,
                                                    self.key_column_headers))
            required_headers.update( set(row_data.keys()) )

        self._get_or_create_headers(required_headers)

        results = UpdateResults()

        # Check for changes and deletes.
        for key_tuple, wks_row in self.data(as_cells=True).iteritems():
            if key_tuple in fixed_data:
                # This worksheet row is in the fixed_data, might be a change or no-change.
                raw_row = fixed_data[key_tuple]
                different_fields = []
                for header, raw_value in raw_row.iteritems():
                    sheet_val = wks_row.db.get(header, "")
                    if not google_equivalent(raw_value, sheet_val):
                        logger.debug("Identified different field '%s' on %s: %s != %s", header, key_tuple, sheet_val, raw_value)
                        different_fields.append(header)

                if different_fields:
                    if self._change_row(key_tuple, 
                                        wks_row, 
                                        raw_row, 
                                        different_fields,
                                        row_change_callback):
                        results.changed += 1
                else:
                    results.nochange += 1
                missing_raw_keys.remove( key_tuple )
            elif delete_rows:
                # This worksheet row is not in fixed_data and needs deleting.
                if self.flag_delete_mode:
                    # Just mark the row as deleted somehow (strikethrough)
                    if not self._is_flagged_delete(key_tuple, wks_row):
                        logger.debug("Flagging row %s for deletion (key %s)", 
                                                   wks_row.row_num, key_tuple)
                        if row_change_callback:
                            row_change_callback(key_tuple, wks_row.db, 
                                        None, self.key_column_headers[:])
                        self._delete_flag_row(key_tuple, wks_row)
                        results.deleted += 1
                else:
                    # Hard delete. Actually delete the row's data.
                    logger.debug("Deleting row: %s for key %s", 
                                                    wks_row.row_num, key_tuple)
                    if row_change_callback:
                        row_change_callback(key_tuple, wks_row.db, 
                                            None, wks_row.db.keys())
                    self._log_change(key_tuple, "Deleted entry.")
                    self._delete_row(key_tuple, wks_row)
                    results.deleted += 1

        if missing_raw_keys:
            # Add missing key in raw
            self._extends(rows=(self.max_row+len(missing_raw_keys)))
            
            empty_cells_list = self._cell_feed(row=self.max_row+1,
                                               col=self.header.first_column, 
                                               max_col=self.header.last_column, 
                                               further_rows=True,
                                               return_empty=True)

            iter_empty_rows = self._yield_rows(empty_cells_list)
            while missing_raw_keys:
                wks_row = iter_empty_rows.next()
                key_tuple = missing_raw_keys.pop()
                logger.debug("Adding new row: %s", str(key_tuple))
                raw_row = fixed_data[key_tuple]
                results.added += 1
                if row_change_callback:
                    row_change_callback(key_tuple, None, 
                                        raw_row, raw_row.keys())
                self._insert_row(key_tuple, wks_row, raw_row)

        self._flush_writes()
        return results

    def _log_change(self, key_tuple, description, old_val="", new_val=""):

        def truncate(text, length=18):
            if len(text) <= length:
                return text
            return text[:(length-3)] + "..."

        if old_val or new_val:
            logger.debug("%s: %s '%s'-->'%s'", ".".join(key_tuple), 
                        description, truncate(old_val), truncate(new_val))
        else:
            logger.debug("%s: %s", ".".join(key_tuple), description)

    def _is_flagged_delete(self, key_tuple, wks_row):
        # Ideally we'd use strikethrough to indicate deletes - but Google api
        # doesn't allow access to get or set formatting.
        for key in key_tuple:
            if key.endswith(DELETE_ME_FLAG):
                return True
        return False

    def _delete_flag_row(self, key_tuple, wks_row):
        for cell in wks_row.cell_list():
            if self.header.col_lookup(cell.col) in self.key_column_headers:
                # Append the DELETE_ME_FLAG
                cell.value = "%s%s" % (cell.value,DELETE_ME_FLAG)
                self._write_cell(cell)

        self._log_change(key_tuple, "Deleted entry")

    def _delete_row(self, key_tuple, wks_row):
        for cell in wks_row.cell_list():
            cell.value = ''
            self._write_cell(cell)

    def _get_value_for_column(self, key_tuple, raw_row, col):
        # Given a column, and a row dictionary.. returns the value
        # of the field corresponding with that column.
        try:
            header = self.header.col_lookup(col)
        except KeyError:
            logger.error("Unexpected: column %s has no header", col)
            return ""

        if header in self.key_column_headers:
            key_dict = dict(zip(self.key_column_headers, key_tuple))
            key_val = key_dict[header]
            if key_val.isdigit() and not key_val.startswith('0'):
                # Do not prefix integers so that the key column can be sorted 
                # numerically.
                return key_val
            return "'%s" % key_val
        elif header in raw_row:
            return raw_row[header]
        elif header in self.header_to_ref_formula:
            return self.header_to_ref_formula[header]

        return ""

    def _insert_row(self, key_tuple, wks_row, raw_row):

        for cell in wks_row.cell_list():
            if cell.col in self.header.columns:
                value = self._get_value_for_column(key_tuple, raw_row, cell.col)
                logger.debug("Batching write of %s", value[:50])
                cell.value = value
                self._write_cell(cell)

        logger.debug("Inserting row %s with batch operation.", wks_row.row_num)

        self._log_change(key_tuple, "Added entry")
        self.max_row += 1

 
    def _change_row(self, key_tuple, wks_row, 
                    raw_row, different_fields,
                    row_change_callback):

        changed_fields = []
        for cell in wks_row.cell_list():
            if cell.col not in self.header.columns:
                continue

            header = self.header.col_lookup(cell.col)
            if header in different_fields:
                raw_val = raw_row[header]
                sheet_val = wks_row.db.get(header,"")
                if (header in self.protected_fields) and sheet_val != "":
                    # Do not overwrite this protected field.
                    continue

                cell.value = raw_val
                self._write_cell(cell)
                changed_fields.append(header)
                self._log_change(key_tuple, ("Updated %s" % header), 
                                 old_val=sheet_val, new_val=raw_val)

        if row_change_callback:
            row_change_callback(key_tuple, wks_row.db, raw_row, changed_fields)

        return changed_fields

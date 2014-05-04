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

__version__ = '0.1'

import logging
from datetime import datetime

logger = logging.getLogger('sheetsync')
change_log = logging.getLogger('sheetsync.change')
gdata_log = logging.getLogger('gdata')

import dateutil.parser 
import gdata.spreadsheet.service, gdata.spreadsheet, gdata.docs.client

MAX_BATCH_LEN = 40960   # Google's limit is 1MB or 1000 batch entries.
DELETE_ME_FLAG = ' (DELETED)'
DEFAULT_SHEET_NAME = 'sheetsync data'
SOURCE_STRING = ('sheetsync.py version:%s' % __version__)
ROOT_FOLDER_URL = 'https://docs.google.com/feeds/default/private/full/folder%3Aroot/contents/' 

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
    text1 = text1.replace('\r',' ').replace('\t',' ')
    text2 = text2.replace('\r',' ').replace('\t',' ')
    if text1 == text2:
        return True
    # Might be dates.
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

    def __setitem__(self, key, cell_elem):
        dict.__setitem__(self, key, cell_elem)
        text_val = cell_elem.cell.text
        if text_val is None:
            text_val = ''
        self.db[key] = text_val

    def cell_list(self):
        for cell_elem in self.itervalues():
            yield (int(cell_elem.cell.row), int(cell_elem.cell.col), cell_elem)

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
            raise CorruptHeader("'%s' was found twice in row %s of %s" % (header, 
                                        self.header_row_ix,  self.sheet_name))
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
    """Represents a single worksheet within a google spreadsheet.
    
    This class tracks the google connection, the reference to the worksheet, as
    well as options governing the structure of the data in the worksheet.. for
    instance:
        * Where the header row is found
        * What header should be used for the key column(s)
        * Whether some columns are protected
    see the __init__ function.for all the options.

    Attributes:
        document_key: The key of the spreadsheet. If you are relying on
          sheetsync to automatically create a spreadsheet then use this
          attribute to access the document_key
        service: The connected instance of the gdata SpreadsheetService. Use at
          your own risk!
    """

    # Global _pool object holds mapping to google connections. 
    _ss_pool = {}
    _doc_client_pool = {}

    def __init__(self, 
                 # Google connection information:
                 username,
                 password,
                 # Spreadsheet and worksheet name or key: 
                 document_key=None, 
                 document_name=None,
                 sheet_name=None,
                 # Operational modifiers:
                 key_field_headers=None, 
                 header_row_ix=1,
                 formula_ref_row_ix=None,
                 flag_deletes=True,
                 protected_fields=None,
                 # Advanced document creation behavior:
                 template_key=None,
                 template_name=None,
                 folder_key=None,
                 folder_name=None):
        """Creates a worksheet object (creating the Google doc if required)

        Args:
          username (str): Google drive account username (an email address).

          password (str): Password for account. To avoid capcha challenges or
            two-factor authentication limits we recommend that you set up an 
            application specific password for the script. For an explaination 
            of this see:
            https://support.google.com/accounts/answer/185833?hl=en

          document_key (str): Document key for the existing spreadsheet to
            sync data to. This is a long hex string and can be found in the 
            URL of the spreadsheet. See this article for more:
            https://productforums.google.com/forum/#!topic/docs/XPOR9bTTS50

          document_name (str): The name of the spreadsheet to sync to, if
            this is not found - it will be created. If known, identifying the 
            document by the document_key is faster.

          sheet_name (str): The name of the worksheet inside the spreadsheet
            that data will be synced to. If omitted then the default name
            "sheetsync data" will be used, and a worksheet created if
            necessary.

         key_field_headers (list of str): A list of the column headers to use 
            for key values. Typically your data's keys are strings, so this is
            a list of length one. For example if your data is indexed by 
            unique username keys you pass in 
                key_field_headers=["Username"] 
            to identify the column that usernames are written to. If your data 
            is indexed by keys which are tuples, for example you have keys like:
                ("student_id","class_number")
            then you would pass in a list of two column headers, e.g.:
                ["Student ID", "Class Number"]
            If no value is given, then the default behavior is to name the
            column "Key"; or "Key1", "Key2", and so forth in the case of tuple
            keys.

          header_row_ix (int): The row number we expect to see column headers
            in. Defaults to 1 (the very top row).
 
          formula_ref_row_ix (int): If you want formulas to be added to some
            cells when inserting new rows then use a formula reference row. We
            recommend you use row 1 for formula references (then hide that row)
            and start the header on row 2.
            As an example, imagine you have a column A containing product 
            indexes and want to have a column B containing hyperlinks
            to a page describing the product. You could populate cell B1 with
            the formula:
              =Hyperlink(Concatenate("http://widgets.com/product/",RC[-1]),RC[-1])
            When adding rows to the spreadsheet, sheetsync will copy this
            formula into the B column.

          flag_deletes (bool): Specify if deleted rows should only be flagged
            for deletion. By default sheetsync does not delete rows of data, it
            just marks that they are deleted by appending the string 
            " (DELETED)" to key values. If you pass in the value "False" then
            rows of data will be deleted by the sync() method if they are not
            found in the input data. Use the inject() method if you only want
            to add data to a worksheet.

          protected_fields (list of str): An optional list of fields (column 
            headers) that contain protected data. sheetsync will only write to 
            cells in these columns if they are blank. This can be useful if you
            are expecting users of the spreadsheet to colaborate on the document
            and edit values in certain columns (e.g. modifying a "Test result" 
            column from "PENDING" to "PASSED") and don't want to overwrite
            their edits.

          template_key (str): This optional key references the spreadsheet 
            that will be copied if a new spreadsheet needs to be created. 
            This is useful for copying over formatting, a specific header 
            order, or apps-script functions.

          template_name (str): As with template_key but the name of the
            template spreadsheet. If known, using the template_key will be
            faster.
 
          folder_key (str): This optional key references the folder that a new
            spreadsheet will be moved to if a new spreadsheet needs to be
            created.

          folder_name (str): Like folder_key this parameter specifies the
            optional folder that a spreadsheet will be created in (if
            required). If a folder matching the name cannot be found, sheetsync
            will attempt to create it.
 
        """

        # Record connection settings, and create a connection.
        self.username = username
        self.password = password
        self.service = self._create_spreadsheet_service()
        self.docs_client = self._create_docs_client()

        # Find or create the spreadsheet document
        document = self._find_document(document_key, document_name)
        if document is None:
            if document_name is None:
                raise Exception("Must specify a document_name")
            # We need to create the document
            template = self._find_document(template_key, template_name)
            if template is None and template_name is not None:
                raise KeyError("Could not find template: %s" % template_name)
            self.folder = self._find_or_create_folder(folder_key, folder_name)
            document = self._create_new_or_copy(source_doc=template, 
                                                target_name=document_name, 
                                                folder=self.folder)
            if document is None:
                raise Exception("Could not create doc '%s'." % document_name)
  
        # Create attribute for document key
        self.document_key = document.GetId().split('%3A')[1]

        # Find or create the worksheet
        if sheet_name is None:
            logger.info("Using the default worksheet name")
            sheet_name = DEFAULT_SHEET_NAME
        self._find_or_create_worksheet(sheet_name)

        # Store off behavioural settings for interacting with the worksheet.
        if key_field_headers is None:
            logger.info("No key fields passed! Will use 'key1', 'key2', etc..")
            key_field_headers = []

        self.key_field_headers = key_field_headers
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

    def _create_spreadsheet_service(self):
        if self.username in self._ss_pool:
            return self._ss_pool[self.username]

        gc = gdata.spreadsheet.service.SpreadsheetsService()
        logger.info("Creating new connection for user '%s'", self.username)
        gc.email = self.username
        gc.password = self.password
        gc.source = SOURCE_STRING

        try:
            gdata_log.info("Logging in")
            gc.ProgrammaticLogin()
        except gdata.service.CaptchaRequired:
            gdata_log.info("Captcha login required")
            captcha_token = gc._GetCaptchaToken()
            url = gc._GetCaptchaURL()
            print "Google login error. Please go to this URL:"
            print "  " + url
            captcha_response = raw_input("Type the captcha image here: ")
            gdata_log.info("Logging in with captcha")
            gc.ProgrammaticLogin(captcha_token, captcha_response)

        # Successful connection, save to the self._ss_pool
        self._ss_pool[self.username] = gc
        return gc

    def _create_docs_client(self):
        if self.username in self._doc_client_pool:
            return self._doc_client_pool[self.username]

        g_doc_cl = gdata.docs.client.DocsClient()
        logger.info("Creating google docs client for user '%s'", self.username)
        g_doc_cl.client_login(self.username, self.password, SOURCE_STRING)

        # Successful connection, save to the connection pool.
        self._doc_client_pool[self.username] = g_doc_cl
        return g_doc_cl


    def _create_new_or_copy(self, 
                            target_name=None,
                            source_doc=None, 
                            folder=None):
        if target_name is None:
            raise KeyError("Must specify a name for the new document")

        if source_doc is not None:
            logger.info("Copying spreadsheet from template.")
            self.new_document = self.docs_client.copy_resource(source_doc,
                                                         title=target_name)
        else:
            # Create new blank spreadsheet.
            logger.info("Creating blank spreadsheet.")
            rsrc = gdata.docs.data.Resource(type='spreadsheet',
                                            title=target_name)
            self.new_document = self.docs_client.create_resource(rsrc)

        logger.info("Created new spreadsheet with Id: %s", 
                self.new_document.GetId())

        if folder is not None:
            self.docs_client.move_resource(self.new_document, 
                                           folder, 
                                           keep_in_collections=False)
            self.docs_client.Delete(ROOT_FOLDER_URL + self.new_document.resource_id.text, 
                                    force=True)
            logger.info("Moved resource to folder.")

        return self.new_document

    def _find_or_create_folder(self, folder_key=None, folder_name=None):
        if folder_key is not None:
            folder_rsrc = self.docs_client.get_resource_by_id(folder_key)
            if folder_rsrc is None:
                raise KeyError("Folder with key %s was not found." % folder_key)
            return folder_rsrc

        if not folder_name:
            return None

        name_query = gdata.docs.client.DocsQuery(title=folder_name,
                                                 title_exact='true',
                                                 show_collections='true')
        matches = self.docs_client.GetResources(q=name_query)
        for doc_rsrc in matches.entry:
            if doc_rsrc.get_resource_type() == 'folder':
                return doc_rsrc

        logger.info("Creating the new folder: %s", folder_name)
        new_folder_rsrc = gdata.docs.data.Resource(type='folder',
                                                   title=folder_name)
        new_folder_rsrc = self.docs_client.create_resource(new_folder_rsrc)
        return new_folder_rsrc


    def _find_document(self, doc_key=None, doc_name=None):
        # Find the document by key and raise "KeyError" if not found.
        # Otherwise search by document_name
        if doc_key is not None:
            logger.debug("Finding document by key.")
            doc_rsrc = self.docs_client.get_resource_by_id(doc_key)
            if doc_rsrc is None:
                raise KeyError("Could not find document with key: %s" % doc_key)
            return doc_rsrc

        if doc_name is None:
            return None

        name_query = gdata.docs.client.DocsQuery(title=doc_name,
                                                 title_exact='true')
        matches = self.docs_client.GetResources(q=name_query)
        if len(matches.entry) == 1:
            return matches.entry[0]

        if len(matches.entry) > 1:
            raise KeyError("Too many matches for doc named: %s" % doc_name)
        
        return None


    def _find_or_create_worksheet(self, sheet_name):
        logger.debug("Finding worksheet '%s'", sheet_name)
        docquery = gdata.spreadsheet.service.DocumentQuery()
        docquery.title = sheet_name
        docquery.title_exact = 'true'
        gdata_log.info("Getting worksheet feed")
        wks_feed = self.service.GetWorksheetsFeed(self.document_key, query=docquery)
        for entry in wks_feed.entry:
            if entry.title.text == sheet_name:
                self.sheet_obj = entry
                self.sheet_id = self.sheet_obj.id.text.rsplit("/",1)[1]
                return self.sheet_obj

        logger.info("Adding a new worksheet named: %s" % sheet_name)
        new_sheet = self.service.AddWorksheet(sheet_name, '20', '10',
                                              self.document_key)
        if new_sheet is not None:
            self.sheet_obj = new_sheet
            self.sheet_id = self.sheet_obj.id.text.rsplit('/',1)[1]
            return new_sheet

        raise Exception("Couldn't add a new worksheet")

    def _extends(self, rows=None, columns=None):
        # Resizes the sheet if needed, to match the given
        # number of rows and/or columns
        self.sheet_number_total_rows = int(self.sheet_obj.row_count.text)
        self.sheet_number_total_cols = int(self.sheet_obj.col_count.text)

        update_rows, update_cols = False, False
        if rows is not None and rows > self.sheet_number_total_rows:
            # Need to add new rows to the spreadsheet.
            update_rows = True
        if columns is not None and columns > self.sheet_number_total_cols:
            update_cols = True

        if update_rows or update_cols:
            gdata_log.info("_extends: fetching worksheet feed")
            self.sheet_obj = self.service.GetWorksheetsFeed(self.document_key, 
                                                  wksht_id=self.sheet_id)
            if update_cols:
                self.sheet_obj.col_count.text = str(columns)
            if update_rows:
                self.sheet_obj.row_count.text = str(rows)

            gdata_log.info("_extends: updating worksheet")
            self.sheet_obj = self.service.UpdateWorksheet(self.sheet_obj)
            self.sheet_number_total_rows = int(self.sheet_obj.row_count.text)
            self.sheet_number_total_cols = int(self.sheet_obj.col_count.text)


    def _write_cell(self, cell_elem):
        # Creates a batch_update if required, and adds the passed cell_elem
        # to it. Then tests if a flush_writes call is required (when the
        # batch write might be close to the 1MB limit)
        if not self._batch_request:
            self._batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()

        gdata_log.debug("_write_cell: Adding batch update")
        self._batch_request.AddUpdate(cell_elem)

        if len(self._batch_request.ToString()) > MAX_BATCH_LEN:
            gdata_log.info("_write_cells: Writing %s cell writes",
                                        len(self._batch_request.entry))
            self.service.ExecuteBatch(self._batch_request, self._batch_href)
            self._batch_request = None


    def _flush_writes(self):
        # Write current batch_updates to google sheet.
        if self._batch_request:
            gdata_log.info("_flush_writes: Writing %s cell writes",
                                        len(self._batch_request.entry))
            self.service.ExecuteBatch(self._batch_request, self._batch_href)
            self._batch_request = None


    def _cell_feed(self, row=None, max_row=None, further_rows=False, 
                         col=None, max_col=None, further_cols=False,
                         return_empty=False):

        # Fetches cell data for a given row, and all following rows if 
        # further_rows is True. If no row is given, all cells are returned.
        cellquery = None
        if row is not None or return_empty or col is not None:
            cellquery = gdata.spreadsheet.service.CellQuery()
        if row is not None:
            cellquery.min_row = str(row)
            if max_row is not None:
                cellquery.max_row = str(max_row)
            elif not further_rows:
                cellquery.max_row = str(row)

        if col is not None:
            cellquery.min_col = str(col)
            if max_col is not None:
                cellquery.max_col = str(max_col)
            elif not further_cols:
                cellquery.max_col = str(col)

        if return_empty:
            cellquery.return_empty = "true"

        gdata_log.info("getting cell feed")
        rfeed = self.service.GetCellsFeed(key=self.document_key,
                                     wksht_id=self.sheet_id,
                                     query=cellquery)

        cells_list = []
        for cell_elem in rfeed.entry:
            row, col = int(cell_elem.cell.row), int(cell_elem.cell.col)
            cells_list.append( (row, col, cell_elem) )

        self._batch_href = rfeed.GetBatchLink().href

        return cells_list

    def read_ref_formulas(self):
        self.header_to_ref_formula = {}

        if self.formula_ref_row_ix:
            for row, col, cell_elem in self._cell_feed(row=self.formula_ref_row_ix):
                ref_formula = cell_elem.cell.inputValue
                header = self.header.col_lookup(col)
                if header and ref_formula.startswith("="):
                    self.header_to_ref_formula[header] = ref_formula


    def _get_or_create_headers(self, required_headers=[]):
        # Reads the header row, adds missing headers if required.
        self.header.reset()

        for row, col, cell_elem in self._cell_feed(row=self.header_row_ix):
            self.header.set(col, cell_elem.cell.text)
 
        headers_to_add = []
        for key_field in self.key_field_headers:
            if key_field not in self.header:
                headers_to_add.append(key_field)
        for header in required_headers:
            if header not in self.header:
                headers_to_add.append(header)

        if not headers_to_add:
            return 

        target_cols = self.header.last_column + len(headers_to_add)
        self._extends(columns=target_cols)

        cells_list = self._cell_feed(row=self.header_row_ix, 
                                     return_empty=True)
        for row, col, cell_elem in cells_list:
            if not headers_to_add:
                break
            if not cell_elem.cell.text:
                header = headers_to_add.pop()
                cell_elem.cell.inputValue = header
                self.header.set(col, header)
                self._write_cell(cell_elem)

        if headers_to_add:
            raise CorruptHeader("Error adding headers")

        self._flush_writes()

    def backup(self, backup_name, folder_key=None, folder_name=None):
        """Makes a copy of the spreadsheet with the name and folder specified. 
        
        Args:
          backup_name (str): The name of the backup document to create.

          folder_key (str): The optional key of a folder that the backup will
            be moved to.
         
          folder_name (str): Like folder_key, references the folder to move a
            backup to. If the folder can't be found, sheetsync will create it.

        """

        folder = self._find_or_create_folder(folder_key, folder_name)
        source_rsrc = self.docs_client.get_resource_by_id(self.document_key)
        backup = self._create_new_or_copy(source_doc=source_rsrc, 
                                        target_name=backup_name, 
                                        folder=folder)

        backup_key = backup.GetId().rsplit('%3A',1)[1]
        return backup_key
         
    def _yield_rows(self, cells_feed):
        cur_row = None
        for row, col, cell_elem in cells_feed:
            if row <= self.header_row_ix:
                # Never yield the header from this function to avoid overwrites
                continue
            if self.formula_ref_row_ix and row == self.formula_ref_row_ix:
                # Never yield the formula ref row to avoid overwrites
                continue
            if cur_row is None or cur_row.row_num != row:
                if cur_row is not None:
                    yield cur_row
                # Make a new row.
                cur_row = Row(row)
            if col in self.header.columns:
                cur_row[self.header.col_lookup(col)] = cell_elem

        yield cur_row

    def data(self, as_cells=False):
        # Reads the worksheet and returns an indexed dictionary of the
        # row objects.
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
            key_list = []
            for key_hdr in self.key_field_headers:
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
        return len(self.key_field_headers)

    #--------------------------------------------------------------------------
    # Update the worksheet to match the raw_data, calling
    # the row_change_callback for any adds/deletes/fieldchanges.
    #
    # Read the data to build a list of required headers and 
    # check the keys are valid tuples.
    # sync and update.
    #--------------------------------------------------------------------------
    def sync(self, raw_data, row_change_callback=None):
        return self._update(raw_data, row_change_callback, delete_rows=True)

    def inject(self, raw_data, row_change_callback=None):
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

            # Cast row_data values to strings.
            fixed_row_data = dict((key, str(val)) for key,val in row_data.iteritems())
            fixed_data[key] = fixed_row_data

            missing_raw_keys.add(key)
            if len(key) != self.key_length:
                raise BadDataFormat("Key %s does not match key field headers %s" % (key,
                                                    self.key_field_headers))
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
                    if row_change_callback:
                        # We expose a callback function in case the driving script needs
                        # to do anything else when fields on the spreadsheet change.
                        row_change_callback(key_tuple, wks_row.db, raw_row, different_fields)
                    if self._change_row(key_tuple, 
                                        wks_row, 
                                        raw_row, 
                                        different_fields):
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
                            row_change_callback(key_tuple, wks_row.db, None, [])
                        self._delete_flag_row(key_tuple, wks_row)
                        results.deleted += 1
                else:
                    # Hard delete. Actually delete the row's data.
                    logger.debug("Deleting row: %s for key %s", 
                                                    wks_row.row_num, key_tuple)
                    if row_change_callback:
                        row_change_callback(key_tuple, wks_row.db, None, [])
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
                    row_change_callback(key_tuple, None, raw_row, [])
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
        # Ideally we'd use strikethrough to indicate deletes - but gdata
        # doesn't allow access to get or set formatting.
        for key in key_tuple:
            if key.endswith(DELETE_ME_FLAG):
                return True
        return False

    def _delete_flag_row(self, key_tuple, wks_row):
        for row, col, cell_elem in wks_row.cell_list():
            if self.header.col_lookup(col) in self.key_field_headers:
                # Append the DELETE_ME_FLAG
                cell_elem.cell.inputValue = cell_elem.cell.text+DELETE_ME_FLAG
                self._write_cell(cell_elem)

        self._log_change(key_tuple, "Deleted entry")

    def _delete_row(self, key_tuple, wks_row):
        for row, col, cell_elem in wks_row.cell_list():
            cell_elem.cell.inputValue = ''
            self._write_cell(cell_elem)

    def _get_value_for_column(self, key_tuple, raw_row, col):
        # Given a column, and a row dictionary.. returns the value
        # of the field corresponding with that column.
        try:
            header = self.header.col_lookup(col)
        except KeyError:
            logger.error("Unexpected: column %s has no header", col)
            return ""

        if header in self.key_field_headers:
            key_dict = dict(zip(self.key_field_headers, key_tuple))
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

        for row, col, cell_elem in wks_row.cell_list():
            if col in self.header.columns:
                value = str(self._get_value_for_column(key_tuple, raw_row, col))
                logger.debug("Batching write of %s", value[:50])
                cell_elem.cell.inputValue = value
                self._write_cell(cell_elem)

        logger.debug("Inserting row %s with batch operation.", wks_row.row_num)

        self._log_change(key_tuple, "Added entry")
        self.max_row += 1

 
    def _change_row(self, key_tuple, wks_row, raw_row, different_fields):

        changed_fields = []
        for row, col, cell_elem in wks_row.cell_list():
            if col not in self.header.columns:
                continue

            header = self.header.col_lookup(col)
            if header in different_fields:
                raw_val = raw_row[header]
                sheet_val = wks_row.db.get(header,"")
                if (header in self.protected_fields) and sheet_val != "":
                    # Do not overwrite this protected field.
                    continue

                cell_elem.cell.inputValue = raw_val
                self._write_cell(cell_elem)
                changed_fields.append(header)
                self._log_change(key_tuple, ("Updated %s" % header), 
                                 old_val=sheet_val, new_val=raw_val)

        return changed_fields

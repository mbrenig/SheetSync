Welcome to SheetSync!
=====================
blurb to come... this is all pre-release experimentation.

Installing
----------
Install **SheetSync** with::

  $ pip install sheetsync

Basic usage
-----------
SheetSync works with data in a dictionary of dictionaries. I.e. each row is
represented by a dictionary, and these are themselves stored in a dictionary
indexed by a row-specific key. E.g.:

.. code-block:: python
    # TODO: Change to a real data object example.
    data = { key1 : { "field" : "value", "field2" : "value2", ... },
            key2 : { "field" : "val",   "field3" : "value3", ... },
            ... }

To synchronize this data (add rows, update rows, delete rows) with a target
spreadsheet you do this:

.. code-block:: python

    import sheetsync
    worksheet = sheetsync.Sheet(username="yourusername@domain.com", 
                                password="app-specific-password",
                                title="My auto-generated spreadsheet",
                                sheet_name="data")
    worksheet.sync(data)

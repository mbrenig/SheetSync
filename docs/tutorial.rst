Tutorial
========
Let's extend the example from Getting Started, and use more of sheetsync's features. 
(With apologies in advance to the Muppets involved).

Customizing the spreadsheet
---------------------------

Key Column Headers
~~~~~~~~~~~~~~~~~~
The first thing we'll fix is that top-left cell with the value 'Key'. The keys
for our data are Names and the column header should reflect that. This is easy
enough to do with the key_column_headers field:

.. code-block:: python
   :emphasize-lines: 3

    target = sheetsync.Sheet(credentials=creds,
                             document_name="Muppet Show Tonight",
                             key_column_headers=["Name"])

.. _templates:

Templates for Formatting
~~~~~~~~~~~~~~~~~~~~~~~~
Google's spreadsheet API doesn't currently allow control over 
cell formatting, but you can specify a template spreadsheet that has the 
formatting you want - and use sheetsync to add data to a copy of the template.
Here's a template spreadsheet created to keep my list of Muppets:

.. image:: Template01.png

https://docs.google.com/spreadsheets/d/1J__SpvQvI9S4bW-BkA0PmPykH8VVT9bdoWZ-AW7V_0U/edit#gid=0

The template's document key is ``1J__SpvQvI9S4bW-BkA0PmPykH8VVT9bdoWZ-AW7V_0U`` we can instruct
sheetsync to use this as a basis for the new spreadsheet it creates as follows:

.. code-block:: python
   :linenos:
   :emphasize-lines: 4

    target = sheetsync.Sheet(credentials=creds,
                             document_name="Muppet Show Tonight",
                             worksheet_name="Muppets",
                             template_key="1J__SpvQvI9S4bW-BkA0PmPykH8VVT9bdoWZ-AW7V_0U",
                             key_column_headers=["Name"])

Note that I've also specified the worksheet name in that example with the 
'worksheet_name' parameter.

Folders
~~~~~~~
If you use folders to organize your Google drive, you can specify the folder a
new spreadsheet will be created in. Use either the 'folder_name' or 'folder_key' parameters. 
Here for example I have a folder with the key ``0B8rRHMfAlOZrWUw4LUhZejk4c0E``:

.. image:: FolderURL.png

and instruct sheetsync to move the new spreadsheet into that folder with this
code:

.. code-block:: python
   :linenos:
   :emphasize-lines: 6

   target = sheetsync.Sheet(credentials=creds,
                            document_name="Muppet Show Tonight",
                            worksheet_name="Muppets",
                            key_column_headers=["Name"],
                            template_key="1J__SpvQvI9S4bW-BkA0PmPykH8VVT9bdoWZ-AW7V_0U",
                            folder_key="0B8rRHMfAlOZrWUw4LUhZejk4c0E")

.. _formulas:

Formulas
~~~~~~~~
Often you'll need some columns to contain formulas that depend on data in other columns, and when new rows are inserted by sheetsync, ideally you'd want those formulas to be added too.
When initializing the spreadsheet you can specify a row (typically above the
header row) that contains reference formulas. Best illustrated by this example

.. image:: MuppetsFormulas.png

https://docs.google.com/spreadsheets/d/1tn-lGqGHDrVbnW2PRvwie4LMmC9ZgYHWlbyTjCvwru8/edit#gid=0

Here row 2 contains formulas (Written out in row 1 for readability) that
reference hidden columns. Row 3 contains the headers. 

When new rows are added to this spreadsheet the 'Photo' and 'Muppet' columns will be populated with a formula similar to the reference row. Here are the parameters to set this up:

.. code-block:: python
   :emphasize-lines: 5,6

    target = sheetsync.Sheet(credentials=creds,
                             document_key="1tn-lGqGHDrVbnW2PRvwie4LMmC9ZgYHWlbyTjCvwru8",
                             worksheet_name="Muppets",
                             key_column_headers=["Name"],
                             header_row_ix=3,
                             formula_ref_row_ix=2)

    animal =  {'Animal': {'Color': 'Red',
                          'Image URL': 'http://upload.wikimedia.org/wikipedia/en/e/e7/Animal_%28Muppet%29.jpg',
                          'Performer': 'Frank Oz',
                          'Wikipedia': 'http://en.wikipedia.org/wiki/Animal_(Muppet)'} }

    target.inject(animal)

Synchronizing data
~~~~~~~~~~~~~~~~~~
Until now all examples have used the 'inject' method to add data into a spreadsheet or
update existing rows. As the name suggests, sheetsync also has a 'sync' method which
will make sure the rows in the spreadsheet match the rows passed to the
function. This might require that rows are deleted from the spreadsheet.

The default behavior is to not actually delete rows, but instead flag them for
deletion with the text "(DELETED)" being appended to the values of the Key columns on rows to delete. This is to help recovery from accidental deletions. Full row deletion can be enabled by passing the flag_deletes argument as follows:

.. code-block:: python
   :emphasize-lines: 11

    target = sheetsync.Sheet(credentials=creds,
                             document_key="1J__SABCD1234bW-ABCD1234kH8VABCD1234-AW7V_0U",
                             worksheet_name="Muppets",
                             key_column_headers=["Name"],
                             flag_deletes=False)

    new_list = { 'Kermit' : { 'Color' : 'Green',
                              'Performer' : 'Jim Henson' },
                 'Fozzie Bear' : {'Color' : 'Orange' } }
                                
    target.sync(new_list)

With rows for Miss Piggy and Kermit already in the spreadsheet, the sync
function (in the example above) would remove Miss Piggy and add Fozzie Bear.

Taking backups
--------------
.. warning::
   The sync function could delete a lot of data from your worksheet if the Key
   values get corrupted somehow. You should use the backup function to protect
   yourself from errors like this.

Some simple mistakes can cause bad results. For instance, if the key column headers on the spreadsheet don't match those passed to the Sheet constructor the sync method will delete all the existing rows and add new ones! You could protect rows and ranges to guard against this, but perhaps the simplest way to mitigate the risk is by creating a backup of your spreadsheet before syncing data. Here's an example:

.. code-block:: python

    target.backup("Backup of my important sheet. 16th June",
                  folder_name = "sheetsync Backups.")

This code would take a copy of the entire spreadsheet that the Sheet instance 'target'
belongs to, name it "Backup of my important sheet. 16th June", and move it to a
folder named "sheetsync Backups.".

Debugging 
---------
sheetsync uses the standard python logging module, the easiest way to find
out what's going on under the covers is to turn on all logging:

.. code-block:: python

    import sheetsync
    import logging
    # Set all loggers to DEBUG level..
    logging.getLogger('').setLevel(logging.DEBUG)
    # Register the default log handler to send logs to console..
    logging.basicConfig()

If you find issues please raise them on `github
<http://github.com/mbrenig/sheetsync/issues>`_, and if you have fixes please
submit pull requests. Thanks!

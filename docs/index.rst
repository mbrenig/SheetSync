Welcome to SheetSync!
=====================
A python library to synchronize rows of data with a google spreadsheet.

**WARNING!** This library is in a hidden development phase. APIs might change, function might
break, bad things could happen. View the code and issues on `github
<http://github.com/mbrenig/SheetSync>`_. 

Installation
------------

Install from PyPi using `pip <http://www.pip-installer.org/en/latest/>`__, a
package manager for Python.::

  pip install sheetsync

Or to develop this library further, you can clone the git repo and install::

  git clone git@github.com:mbrenig/SheetSync.git SheetSyncRepo
  pip install SheetSyncRepo

Note, you may need to run the commands above with ``sudo``.

Getting Started
---------------
SheetSync works with data in a dictionary of dictionaries. Each row is
represented by a dictionary, and these are themselves stored in a dictionary
indexed by a row-specific key. For example:

.. code-block:: python

    data = { "Kermit": {"Color" : "Green", "Performer" : "Jim Henson"},
             "Miss Piggy" : {"Color" : "Pink", "Performer" : "Frank Oz"}
            }

To synchronize this data (insert, modify or delete rows) with a target
sheet in a google spreadsheet document you do this:

.. code-block:: python

    import sheetsync
    # Get or create a spreadsheet...
    target = sheetsync.Sheet(username="googledriveuser@domain.com", 
                             password="app-specific-password",
                             document_name="Let's try out SheetSync")
    # Add data to the spreadsheet...
    target.sync(data)
    print "Review the new spreadsheet created here: %s" % target.document_href

This creates a new spreadsheet document in your google drive and then inserts the data.

Advanced topics
---------------
TODO: Will cover:
 - document_key instead of name
 - folder key
 - template key
 - header index
 - formula ref

The (DELETED) flag and backups
------------------------------
Couldn't the sync function delete all my rows? **Yes it could!**

You might have noticed by now that the default deletion behavior for
the sync method is to append the values in Key fields with "(DELETED)" and not actually 
delete the rows. This is to avoid accidental bulk deletion while a new user
experiments with the library. It can be turned off by passing the argument: 'flag_deletes = False' into the sheetsync.Sheet constructor. In that case rows that need to be deleted
from the target spreadsheet will be erased.

Some simple mistakes can cause very bad results. For instance, if the key column headers on the spreadsheet don't match those passed to the Sheet constructor the sync method will delete all the existing rows and add new ones! You can mitigate the risks by taking backups of your spreadsheet before sync'ing data. Here's an example:

.. code-block:: python

    target.backup("Backup of my important sheet. 15th May",
                  folder_name = "SheetSync Backups.")

This code would take a copy of the entire spreadsheet that the Sheet 'target'
belongs to, name it "Backup of my important sheet. 15th May", and move it to a
folder named "SheetSync Backups.".

Debugging 
---------
SheetSync uses the standard python logging module, the simplest way to find
out what it's doing under the covers is to turn on all logging:

.. code-block:: python

    import sheetsync
    import logging
    # Set all loggers to DEBUG level..
    logging.getLogger('').setLevel(logging.DEBUG)
    # Register the default log handler to send logs to console..
    logging.basicConfig()

If you find issues please raise them on `github
<http://github.com/mbrenig/SheetSync/issues>`_, and if you have fixes please
submit pull requests. Thanks!

SheetSync
=========

|Build Status|

A python library to create, update and delete rows of data within a google spreadsheet. `Click here to read the full documentation.
<http://sheetsync.readthedocs.org/>`__

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

To insert this data (add or update rows) with a target
sheet in a google spreadsheet document you do this:

.. code-block:: python

    import sheetsync
    # Get or create a spreadsheet...
    target = sheetsync.Sheet(username="googledriveuser@domain.com", 
                             password="app-specific-password",
                             document_name="Let's try out SheetSync")
    # Insert or update rows on the spreadsheet...
    target.inject(data)
    print "Review the new spreadsheet created here: %s" % target.document_href

This creates a new spreadsheet document in your google drive and then inserts the data like so:

.. image:: Sheet1.png

Later on you'll probably want to access this data, to do that note the
spreadsheet's document key from the URL:

.. image:: URL.png

and access the data as follows:

.. code-block:: python

    source = sheetsync.Sheet(username="googledriveuser@domain.com",
                             password="app-specific-password",
                             document_key="1bnieREGAyXZ2TnhXgYrIacCIY09Q2lfGXNZbjsvJ82M")
    print source.data()

The 'inject' method only adds or updates rows. If you want to delete rows from the spreadsheet to keep it in sync with the input data then use the 'sync' method described in the 'Synchronizing data' section below.


Debugging 
~~~~~~~~~
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

Digging Deeper
~~~~~~~~~~~~~~
The `full documentation <http://sheetsync.readthedocs.org/>`__ covers extra features such as:

-  `Using templates when creating new spreadsheets <http://sheetsync.readthedocs.org>`__
-  `Using a row for formula references <http://sheetsync.readthedocs.org>`__
-  `The on_change_callback function [TODO]<http://sheetsync.readthedocs.org>`__
-  `Live examples [TODO]<http://sheetsync.readthedocs.org>`__


.. |Build Status| image:: https://travis-ci.org/mbrenig/SheetSync.svg?branch=master
   :target: https://travis-ci.org/mbrenig/SheetSync

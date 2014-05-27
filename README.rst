SheetSync
=========

|Build Status|

A Python library for creating and updating spreadsheets that contain rows of
data. `Click here to read the full documentation.
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

The `full documentation <http://sheetsync.readthedocs.org/>`__
is in the process of being written but will soon cover features such as:

-  `Using templates to create new spreadsheets <http://sheetsync.readthedocs.org>`__
-  `Using a row for formula references <http://sheetsync.readthedocs.org>`__
-  `The on_change_callback function <http://sheetsync.readthedocs.org>`__
-  `Live examples <http://sheetsync.readthedocs.org>`__


.. |Build Status| image:: https://travis-ci.org/mbrenig/SheetSync.svg?branch=master
   :target: https://travis-ci.org/mbrenig/SheetSync

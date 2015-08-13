import logging, os

from sheetsync import Sheet, ia_credentials_helper

logging.getLogger('sheetsync').setLevel(logging.DEBUG)
logging.basicConfig()

CLIENT_ID = os.environ['SHEETSYNC_CLIENT_ID']  
CLIENT_SECRET = os.environ['SHEETSYNC_CLIENT_SECRET']  

creds = ia_credentials_helper(CLIENT_ID, CLIENT_SECRET, 
                    credentials_cache_file='credentials.json',
                    cache_key='default')

data = { "Kermit": {"Color" : "Green", "Performer" : "Jim Henson"},
         "Miss Piggy" : {"Color" : "Pink", "Performer" : "Frank Oz"}
        }

target = Sheet(credentials=creds, document_name="sheetsync quicktest")

target.inject(data)

print "Spreadsheet created here: %s" % target.document_href

import sheetsync, os

data = { "Kermit": {"Color" : "Green", "Performer" : "Jim Henson"},
         "Miss Piggy" : {"Color" : "Pink", "Performer" : "Frank Oz"}
        }

GOOGLE_U = os.environ['SHEETSYNC_GOOGLE_ACCOUNT']
GOOGLE_P = os.environ['SHEETSYNC_GOOGLE_PASSWORD']

print "GOOGLE_U=%s" % GOOGLE_U
print "GOOGLE_P=%s" % GOOGLE_P
# Get or create a spreadsheet...
target = sheetsync.Sheet(username=GOOGLE_U,
                         password=GOOGLE_P,
                         document_name="sheetsync quicktest")
# Insert or update rows on the spreadsheet...
target.inject(data)
print "Review the new spreadsheet created here: %s" % target.document_href

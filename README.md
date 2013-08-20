Cell Merge Work Around
===================

This tool provides a work around for the bug in Google Apps Script (Issue 618 -- http://code.google.com/p/google-apps-script-issues/issues/detail?id=618) that it provides you with no way to determine if cells are merged. 

Usage
------

	A tool for getting the row and column spans of a single sheet in a Google
	Drive Spreadsheet. (Note that for command line access, Google requires client
	authentication *even if* a document is "Public to anyone on the web").

	optional arguments:
	  -h, --help            show this help message and exit
	  -t GOOGLE_SERVICE, --service_type GOOGLE_SERVICE
	                        If you intend to work with something other than a
	                        spreadsheet you'll need to change this. (Default :
	                        'wise')
	  -k WORKBOOK_KEY, --spreadsheet_key WORKBOOK_KEY
	                        The key parameter taken from URL of the spreadsheet.
	                        (Required!)
	  -s SHEET_ID, --sheet_id SHEET_ID
	                        The identifier of the single sheet to be accessed.
	                        (Default : 0)
	  -a GAUTH, --service_authentication GAUTH
	                        The 'Auth' parameter returned by Google Client Login.
	                        (Can't be used with 'user_id')
	  -u GOOGLE_UID, --user_id GOOGLE_UID
	                        Thi is the Google user's email address. (Can't be used
	                        with 'service_authentication')
	  -p GOOGLE_PWD, --user_passwd GOOGLE_PWD
	                        The Google user's password
	
You will need to provide, at least, the url of the spreadsheet and your Google Auth key for spreadsheets.

An example command to execute would be :

	./getSpans.py -k "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" -s 3 -a "yourGoogleAuthKey" > spans.csv

If you do not have a Google Auth key you can get one using the command :

	./getGauth.py wise yourUserName@gmail.com yourPassword

Note:  *wise* is the "service name" Google gives to the "Spreadsheets Data API" (more listed here -- https://developers.google.com/gdata/faq#clientlogin)


You can also execute getSpans like this : 

	./getSpans.py -k "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" -s 3 -u "yourUID@gmail.com" -p "yourPWD" > spans.csv
	
In which case the Auth key is generated, and used, internally.

The output is a CSV list of cells that have colspan or rowspan (or both) attributes greater than one.

For testing you can use the key "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" it is "Public to anyone on the web."

For sheet #0 you should get :

	"row", "col", "cols", "rows"
	"2", "1", "1", "6"
	"5", "2", "3", "5"
	"4", "3", "3", "1"

For sheet #2 you should get :

	"row", "col", "cols", "rows"
	"1", "2", "1", "6"
	"8", "1", "2", "2"
	"5", "4", "3", "5"

You can then load these back into your spreadsheet as a new sheet and use it as a reference for the location of merged cells for what ever (static) purpose, you might need.  Obviously you will have to run it before any sheet copying actions in case something has changed since the last time.




Cell Merge Work Around
===================

This tool provides a work around for the bug in Google Apps Script (Issue 618 -- http://code.google.com/p/google-apps-script-issues/issues/detail?id=618) that you cannot not detect if cells are merged. 

Usage
------

	A tool for getting the row and column spans of a single sheet in a Google
	Drive Spreadsheet. With Google requires authentication *even if* a document is
	"Public to anyone on the web".
	
	optional arguments:
	  -h, --help            show this help message and exit
	  -t GOOGLE_SERVICE, --service_type GOOGLE_SERVICE
	                        If you intend to work with something other than a
	                        spreadsheet you'll need to change this.
	  -k WORKBOOK_KEY, --spreadsheet_key WORKBOOK_KEY
	                        The key parameter taken from URL of the spreadsheet.
	                        (Required!)
	  -s SHEET_ID, --sheet_id SHEET_ID
	                        The identifier of the single sheet to be accessed.
	  -a GAUTH, --service_authentication GAUTH
	                        The 'Auth' parameter returned by Google Client Login.
	                        (Can't be used with 'user_id')
	  -u GOOGLE_UID, --user_id GOOGLE_UID
	                        Thi is the Google user's email address. (Can't be used
	                        with 'service_authentication')
	  -p GOOGLE_PWD, --user_passwd GOOGLE_PWD
	                        The Google user's password

You will need to have available the url of the spreadsheet and your Google Auth key for spreadsheets.

An example command to execute would be:

	./getSpans.py -k "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" -s 3 -a "yourGoogleAuthKey" > spans.csv




Cell Merge Work Around
===================

This tool provides a work around for two bugs in Google Apps Script.
 - Issue 618 -- http://code.google.com/p/google-apps-script-issues/issues/detail?id=618  -- that it provides no way to determine if cells are merged.
 - Unreported issue - No way to determine cell border settings.

What it does
-------------

This tool grabs a HTML copy of the sheets you specify and generates a CSV file that holds the row span, column span & / or border style data of each affected cell.

Supply it with ...

- your access credentials
- a Spreadsheet unique key
- a comma separated list of sheet index numbers (eg 3,6,2  OR just  7)
 
... and get back (via stdout): a CSV listing of all merged cell areas on the specified sheet.

Developed as a Python 2.7 shell script running out of /usr/bin/python

Tested only in Ubuntu 12.04 LTS

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
      -s SHEET_IDS, --sheet_ids SHEET_IDS
                            A comma separated list of sheet identifiers of the
                            single sheet to be accessed, eg 1,8,4,3. (Default : 0)
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

	./generate_CSV_File.py -k "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" -s 2,0 -a "yourGoogleAuthKey" > spans.csv

If you do not have a Google Auth key you can get one using this other command :

	./getGauth.py wise yourUserName@gmail.com yourPassword

Note:  *wise* is the "service name" Google gives to the "Spreadsheets Data API" (more listed here -- https://developers.google.com/gdata/faq#clientlogin)


You can also execute generate_CSV_File like this : 

	./generate_CSV_File.py -k "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" -s 2,0 -u "yourUID@gmail.com" -p "yourPWD" > spans.csv
	
In which case the Auth key is generated, and used, internally.

The output is a CSV list of cells that have colspan or rowspan (or both) attributes greater than one.

For testing you can use the key "0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE" it is "Public to anyone on the web." The full workbook is here : https://docs.google.com/spreadsheet/ccc?key=0AhgdNB3-bSxAdDFBQWJ3YTAzd015UFJTZ3FwZlc1TlE#gid=2

You should get a result like this :

        "sheet", "row", "col", "rows", "cols", "Tstyle", "Tcolor", "Tthick", "Lstyle", "Lcolor", "Lthick", "Bstyle", "Bcolor", "Bthick", "Rstyle", "Rcolor", "Rthick"
        "2", "22", "3", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "dashed", "#1155cc", "1"
        "2", "19", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "17", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "20", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "20", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "22", "4", "1", "5", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "2", "18", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "0", "4", "3", "1", "3", "", "", "", "", "", "", "", "", "", "", "", ""
        "2", "18", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "14", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "0", "3", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "0", "3", "3", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "3", "6", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "5", "2", "5", "3", "", "", "", "", "", "", "", "", "", "", "", ""
        "0", "3", "5", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "2", "6", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "2", "5", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "2", "4", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "2", "3", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "0", "2", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "0", "2", "1", "6", "1", "", "", "", "", "", "", "", "", "", "", "", ""
        "2", "17", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "13", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "16", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "15", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "21", "6", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "21", "5", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "21", "4", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "21", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "14", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "22", "8", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "2", "12", "3", "11", "1", "", "", "", "", "", "", "solid", "#990000", "1", "dashed", "#1155cc", "1"
        "2", "12", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "12", "4", "10", "4", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "12", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "5", "4", "3", "3", "", "", "", "", "", "", "", "", "", "", "", ""
        "2", "21", "8", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "4", "7", "", "", "", "", "", "", "", "", "dashed", "#1155cc", "1", "", "", ""
        "2", "4", "6", "", "", "", "", "", "", "", "", "", "", "", "dashed", "#1155cc", "1"
        "2", "13", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "10", "3", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "", "", ""
        "2", "21", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "16", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "10", "7", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "", "", ""
        "2", "10", "6", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "", "", ""
        "2", "10", "5", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "", "", ""
        "2", "10", "4", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "", "", ""
        "2", "1", "2", "6", "1", "", "", "", "", "", "", "", "", "", "", "", ""
        "2", "11", "3", "1", "5", "", "", "", "", "", "", "", "", "", "", "", ""
        "2", "11", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "0", "3", "4", "", "", "", "", "", "", "", "", "solid", "#990000", "1", "solid", "#990000", "1"
        "2", "15", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "11", "8", "11", "1", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"
        "2", "3", "7", "", "", "", "", "", "", "", "", "dashed", "#1155cc", "1", "", "", ""
        "2", "19", "7", "", "", "", "", "", "", "", "", "dashed", "#1c4587", "1", "dashed", "#1c4587", "1"
        "2", "22", "2", "", "", "", "", "", "", "", "", "", "", "", "solid", "#990000", "1"

You can then load these back into your spreadsheet as a new sheet and use it as a reference for the location of merged cells for what ever (static) purpose, you might need.  Obviously you will have to run it before any sheet copying actions in case something has changed since the last time.




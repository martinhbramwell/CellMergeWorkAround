Cell Merge Work Around
===================

This tool provides a work around for two bugs in Google Apps Script.
 - Issue 618 -- http://code.google.com/p/google-apps-script-issues/issues/detail?id=618  -- that it provides no way to determine if cells are merged.
 - Unreported issue - No way to determine cell border settings.

What it does
-------------

This tool grabs a HTML copy of the sheets you specify and generates a CSV file that holds the row span, column span & / or border style data of each affected cell.

Optionally, it pushes the CSV data back into the same workbook as a new sheet.

Supply it with ...

- your OAuth2 credentials
- a Spreadsheet unique key
- a comma separated list of sheet index numbers (eg 3,6,2  OR just  7)
 
... and get back (via stdout): a CSV listing of all merged cell areas on the specified sheet.

Developed as a Python 2.7 shell script running out of /usr/bin/python

Tested only in Ubuntu 12.04 LTS

Learning Python?
-------

This program is fairly straight forward.  Experts would see many ways to do it better, but -- it works!

You can use it to learn about screen scraping other people's spreadsheets :-)  :

 - urllib2 : reading a web page from a remote server and writing it on your local machine
 - TinyCss : analyzing CSS style info
 - argparse : getting command line parameters
 - BeautifulSoup : analyzing HTML to get at data you want
 - gspread : writing data from a Python program into a Google Spreadsheet

Preparation
------

Let's get the dependencies out of the way.  You will need the Python Package Installer "pip", so do:

		sudo apt-get install -y python-pip
		
 1. sudo pip install beautifulsoup4
 1. sudo pip install tinycss
 

Obtain this project [  https://github.com/martinhbramwell/CellMergeWorkAround  ] and also this one [  https://github.com/martinhbramwell/gspread  ]

You will have two directories like this :

    /-+
      + - CellMergeWorkAround
      + - gspread
           + - docs
           + - gspread
                + - googauth.py
                + - etc, etc
           + - logs
           + - tests

CellMergeWorkAround will not work without parts of gspread, so you need to copy the gspread/gspread subdirectory into CellMergeWorkAround, so as to end up with this:

    /-+
      + - CellMergeWorkAround
           + - gspread
                + - googauth.py
                + - etc, etc
      + - gspread
           + - docs
           + - gspread
                + - googauth.py
                + - etc, etc
         + - logs
         + - tests

Also, you may want to be able to email your remote user to get them to tell Google your access attempts are authorized anmd should be permitted.
Run the CellMergeWorkAround/gspread/prepSMTP.py and retrieve the values from the foot of the file test_parms.py


Usage
------

	usage: generate_CSV_File.py [-h] [-il IMMEDIATE_LOAD] -ssk WORKBOOK_KEY -si
		                    SHEET_IDS -ue USER_EMAIL -ci CLIENT_PROJECT
		                    [-ce CLIENT_EMAIL] [-st GOOGLE_SERVICE]
		                    [-sn SHEET_NAME] [-sat SMTP_ACCESS_TOKEN]
		                    [-cs CLIENT_SECRET] [-am AUTH_METHOD]
		                    [-srt SMTP_REFRESH_TOKEN] [-ru REDIRECT_URI]

	A tool for getting the row span, column span & / or border style data of one
	or more sheets in a Google Drive Spreadsheet. (Note that for command line
	access, Google requires client authentication *even if* a document is "Public
	to anyone on the web"). If OAuth2 arguments are configured and
	"-i/--immediate_load" is indicated, then the CSV file will be pushed into the
	same spreadsheet from which the data was first read.

	optional arguments:
	  -h, --help            show this help message and exit
	  -il IMMEDIATE_LOAD, --immediate_load IMMEDIATE_LOAD
		                The name for the sheet in the target workbook to be
		                created/replaced with the generated CSV. (Default :
		                Null)
	  -ssk WORKBOOK_KEY, --spreadsheet_key WORKBOOK_KEY
		                The key parameter taken from URL of the spreadsheet.
		                (Required.)
	  -si SHEET_IDS, --sheet_ids SHEET_IDS
		                A comma separated list of sheet identifiers of the
		                single sheet to be accessed, eg 0,1,8,4,3.
	  -ue USER_EMAIL, --user_email USER_EMAIL
		                This is the GMail account of the spreadsheet owner.
	  -ci CLIENT_PROJECT, --client_id CLIENT_PROJECT
		                This is the Client ID you get from Google's API
		                console.
	  -ce CLIENT_EMAIL, --client_email CLIENT_EMAIL
		                This is the GMail account that will be used for SMTP
		                communication with spreadsheet owners.
	  -st GOOGLE_SERVICE, --service_type GOOGLE_SERVICE
		                If you intend to work with something other than a
		                spreadsheet you'll need to change this. (Default :
		                'wise')
	  -sn SHEET_NAME, --sheet_name SHEET_NAME
		                The name of the sheet to create/replace. (Required.
		                Default : meta_patches)
	  -sat SMTP_ACCESS_TOKEN, --smtp_access_token SMTP_ACCESS_TOKEN
		                The token you get from Google that allows you to use
		                your own GMail account as an SMTP server, without
		                needing to use your username and password.
	  -cs CLIENT_SECRET, --client_secret CLIENT_SECRET
		                The Client Secret you get from Google's API console.
	  -am AUTH_METHOD, --auth_method AUTH_METHOD
		                Can be either 'ForDevices', 'InstalledApp' or
		                'ToBeDropped' (to edit storage to drop the connection
		                between project and user)
	  -srt SMTP_REFRESH_TOKEN, --smtp_refresh_token SMTP_REFRESH_TOKEN
		                This refreshes the access token that refreshes your
		                authorization to use your GMail account as an SMTP
		                server without having to ask you again.
	  -ru REDIRECT_URI, --args.redirect_uri REDIRECT_URI
		                This is the Redirect URI you get from Google's API
		                console.



	
You will need to provide, at least, the url of the spreadsheet and your Google OAuth2 credentials for spreadsheets.

An example command to execute would be :

	 ./generate_CSV_File.py -ssk 0Asxy-TdDgbjidxxxxxxxxxxxxxxxxlpHdkw0ek1Md2c -si 6 -ue dude.awap@gmail.com -cs 'Zis38xxxxxxxxxfMRthW-' -ci 2046189xxxxxxxxxxxxxxxxxxxx579r0l.apps.googleusercontent.com -am ForDevices -ce alicia.factorepo@gmail.com  -il meta_patches -sat 'ya29.AHES6Zxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx6wX4c1w' -srt '1/hO_TrQqKJ063FGxxxxxxxxxxxxxxxxxxxxxZUg9ZvE'

Note:  *wise* is the "service name" Google gives to the "Spreadsheets Data API" (more listed here -- https://developers.google.com/gdata/faq#clientlogin)

The output is a CSV list of:
 - cells that span rows &/or columns
 - have borders.

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

If you specify "-il a_sheet_name" , the program will then load these lines back into the spreadsheet as a new sheet of that name.  You can then use it as a reference for the location of merged cells for what ever (static) purpose, you might need.  Obviously, you will always have to run it after every change to the indicated sheets before any sheet copying actions in case spans or borders have changed since the last time.





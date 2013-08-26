#!/usr/bin/python
#

import argparse, urllib2
from BeautifulSoup import BeautifulSoup
import getGauth
import getSpans
import getBorders

import os
import math

def downloadChunks(url, filename):
    """Helper to download large files
        the only arg is a url
       this file will go to a temp directory
       the file will also be downloaded
       in chunks and print out how much remains
    """

    baseFile = os.path.basename(filename)

    #move the file to a more uniq path
    os.umask(0002)
    temp_path = "/tmp/"
    try:
    
        file = os.path.join(temp_path, baseFile)

        req = urllib2.urlopen(url)
        
        downloaded = 0
        CHUNK = 2**15
        with open(file, 'wb') as fp:
            while True:
                chunk = req.read(CHUNK)
                downloaded += len(chunk)
                if not chunk: break
                fp.write(chunk)
    except urllib2.HTTPError, e:
        print "HTTP Error:",e.code , url
        return False
    except urllib2.URLError, e:
        print "URL Error:",e.reason , url
        return False

    return open(file)
    
#use it like this
#downloadChunks("http://localhost/a.zip")


def getConnectionHeader(gauth, google_UID, google_PWD, google_service):

    if gauth is not None:
        google_authorization_key = gauth
    else:
        google_authorization_key = getGauth.getGoogleAuthorizationKey(
                 google_service
               , google_UID
               , google_PWD
            )
            
    return {"Authorization": "GoogleLogin auth=" + google_authorization_key}


def main():

    desc = 'A tool for getting the row span, column span & / or border style data of one or more sheets in a Google Drive Spreadsheet. (Note that for command line access, Google requires client authentication *even if* a document is "Public to anyone on the web").'
    usage = "usage: %prog [options] arg"
    parser = argparse.ArgumentParser(description=desc)
    
    parser.add_argument("-t", "--service_type", dest="google_service", default="wise"
                      , help="If you intend to work with something other than a spreadsheet you'll need to change this. (Default : 'wise')")

    parser.add_argument("-k", "--spreadsheet_key", dest="workbook_key", required=True
                      , help="The key parameter taken from URL of the spreadsheet. (Required!)")
                      
    parser.add_argument("-s", "--sheet_ids", dest="sheet_ids"
                      , help="A comma separated list of sheet identifiers of the single sheet to be accessed, eg 1,8,4,3. (Default : 0)")
                      
    group1 = parser.add_mutually_exclusive_group(required=True)
    group1.add_argument("-a", "--service_authentication", dest="gauth"
                      , help="The 'Auth' parameter returned by Google Client Login. (Can't be used with 'user_id')")
    
    group1.add_argument("-u", "--user_id", dest="google_UID", help="Thi is the Google user's email address. (Can't be used with 'service_authentication')")

    parser.add_argument("-p", "--user_passwd", dest="google_PWD", help="The Google user's password")

    args = parser.parse_args()

    if args.sheet_ids is None:
        return "Found no sheet ID to use."

    
    header = getConnectionHeader(args.gauth, args.google_UID, args.google_PWD, args.google_service)

    url_part  = "https://docs.google.com/feeds/download/spreadsheets/Export?exportFormat=html"
    url_part += "&key=" + args.workbook_key

    patchRows = '"sheet", "row", "col", "rows", "cols", "Tstyle", "Tcolor", "Tthick", "Lstyle", "Lcolor", "Lthick", "Bstyle", "Bcolor", "Bthick", "Rstyle", "Rcolor", "Rthick"'
    
    dictCells = {}
    for sheet_id in args.sheet_ids.split(','):

        request = urllib2.Request(url_part + "&gid=" + str(sheet_id), headers=header)

        html_file = downloadChunks(request, "temporary.html")

        dictCells = getBorders.getSpreadsheetCellBorders(dictCells, html_file, sheet_id)  #  BORDERS !!!
    
        dictCells = getSpans.getSpreadsheetSpannedCells(dictCells, html_file, sheet_id)  #  SPANS!!!

    no_spans = '"", ""'
    no_bords = '"", "", "", "", "", "", "", "", "", "", "", ""'

    for cell, contents in dictCells.items() :
    
        frag = '"{}", "{}", "{}", '.format(cell.split(':')[0], cell.split(':')[1], cell.split(':')[2])
        if 'span' in contents :
            frag += contents['span']
        else :
            frag += no_spans

        frag += ', ' 

        if 'brdr' in contents :
            frag += contents['brdr']
        else :
            frag += no_bords
            
        patchRows += '\n' + frag
            
    return patchRows


if __name__ == "__main__":

    result = main()
    print result
	

# Example :  ./getSpans.py -k "0AhGdN..S.P.R.E.A.D.S.H.E.E.T...K.E.Y..c1TlE" -s 3 -u "yourUserName@gmail.com" -p "yourPassword" > spans.csv
#    OR
# Example :  ./getSpans.py -k "0AhGdN..S.P.R.E.A.D.S.H.E.E.T...K.E.Y..c1TlE" -s 3 -a "yourGoogleAuthKey" > spans.csv



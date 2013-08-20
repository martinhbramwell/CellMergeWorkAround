#!/usr/bin/python
#

import argparse, urllib2
from BeautifulSoup import BeautifulSoup
import getGauth

def getSpreadsheetSpannedCells(workbook_key, gauth, google_UID, google_PWD, sheet_id=0, google_service = "wise"):

    sheet_number=str(sheet_id)    
    url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + workbook_key + "&exportFormat=html&gid=" + sheet_number
    
    if gauth is not None:
        google_authorization_key = gauth
    else:
        google_authorization_key = getGauth.getGoogleAuthorizationKey(
                 google_service
               , google_UID
               , google_PWD
            )

    request = urllib2.Request(url, headers={"Authorization": "GoogleLogin auth=" + google_authorization_key})

    page = urllib2.urlopen(request).read()
#    return url
#    return page
    
    soup = BeautifulSoup(page)

    table = soup.body.table
            
    idxRow = 0
    spanned_cells = {}
    for row in table:
        idxCell = 0
        col = {}
        foundInRow = False
        for cell in row:
            idxAttr = 1
            cellSpec = {'c':'1', 'r':'1'}
            found = False
            for attr, value in cell.attrs:
                if attr == "colspan":
                    cellSpec['c'] = value
                    found = True
                if attr == "rowspan":
                    cellSpec['r'] = value
                    found = True
                if found:
                    foundInRow = True
                    col[str(idxCell)] = cellSpec
                idxAttr += 1
            if foundInRow:
                spanned_cells[str(idxRow)] = col
            idxCell += 1
        idxRow += 1

    result = '"row", "col", "cols", "rows"'
    for item in spanned_cells:
        row = spanned_cells[item]
        for coord in row:
            cell = row[coord]
            result += '\n"{}", "{}", "{}", "{}"'.format(item, coord, cell['c'], cell['r'])

    return result



def main():

    getSpans = 'A tool for getting the row and column spans of a single sheet in a Google Drive Spreadsheet. With Google requires authentication *even if* a document is "Public to anyone on the web".'
    usage = "usage: %prog [options] arg"
    parser = argparse.ArgumentParser(description=getSpans)
    
    parser.add_argument("-t", "--service_type", dest="google_service", default="wise"
                      , help="If you intend to work with something other than a spreadsheet you'll need to change this.")

    parser.add_argument("-k", "--spreadsheet_key", dest="workbook_key", required=True
                      , help="The key parameter taken from URL of the spreadsheet. (Required!)")
                      
    parser.add_argument("-s", "--sheet_id", dest="sheet_id"
                      , help="The identifier of the single sheet to be accessed.")
                      
    group1 = parser.add_mutually_exclusive_group(required=True)
    group1.add_argument("-a", "--service_authentication", dest="gauth"
                      , help="The 'Auth' parameter returned by Google Client Login. (Can't be used with 'user_id')")
    
    group1.add_argument("-u", "--user_id", dest="google_UID", help="Thi is the Google user's email address. (Can't be used with 'service_authentication')")

    parser.add_argument("-p", "--user_passwd", dest="google_PWD", help="The Google user's password")

    args = parser.parse_args()
    
    return getSpreadsheetSpannedCells(args.workbook_key, args.gauth, args.google_UID, args.google_PWD, args.sheet_id, args.google_service)


if __name__ == "__main__":
    result = main()
    print result
	

# Example :  ./getSpans.py -k "0AhGdN..S.P.R.E.A.D.S.H.E.E.T...K.E.Y..c1TlE" -u "yourUserName@gmail.com" -p "yourPassword" > spans.csv


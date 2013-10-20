#!/usr/bin/python
#

import argparse
import urllib

try :

    from goauth2_helper import GeneratePermissionUrl, AuthorizeTokens, RefreshToken, GenerateOAuth2String
    
except ImportError :

    # Get Google oauth2 helper file
    webFile = urllib.urlopen('http://google-mail-oauth2-tools.googlecode.com/svn/trunk/python/oauth2.py')
    localFile = open('goauth2_helper.py', 'w')
    localFile.write(webFile.read())
    webFile.close()
    localFile.close()
    
    from goauth2_helper import GeneratePermissionUrl, AuthorizeTokens, RefreshToken, GenerateOAuth2String

try :
    import gspread
    from gspread import googoauth
    
except ImportError :

    github_url = 'https://github.com/martinhbramwell/gspread/tree/master/gspread'
    print 'You need a sub-directory called "gspread", containing all the files listed here:\n     %s' % github_url
    exit(-1)

import getGauth
import getSpans
import getBorders

import os
import math

import urllib2

def pushBackToSpreadsheet(credentials, workbook_key, sheet_name, patchData) :

    # Hook up to Google Spreadsheet stream.
    conn = gspread.connect(credentials)

    # Get the destination workbook
    wkbk = conn.open_by_key(workbook_key)

    # Get the target sheet in that workbook 
    try :
        wkbk.del_worksheet(wkbk.worksheet(sheet_name))
    except gspread.WorksheetNotFound as wsnf :
        pass
    sht = wkbk.add_worksheet(sheet_name, patchData['h'], patchData['w'])


    # Create a lisat of all the target cells
    rangeSpec = 'A1:{}'.format(sht.get_addr_int(patchData['h'], patchData['w']))
    allCells = sht.range(rangeSpec)

    w = patchData['w']
    h = patchData['h']

    # Fill each of the target cells with 
    cnt = 0
    for cell in allCells :
        r = cell.row
        c = cell.col
        idx = (r-1) * w + (c-1)
#        print 'At #{}({}) -- Cell R{}C{} has'.format(idx, cnt, r, c)
#        print '- -  {}'.format(patchData['data'][r-1][c-1])
        allCells[idx].value = patchData['data'][r-1][c-1]
        cnt += 1

    sht.update_cells(allCells)
    
    conn.disconnect()

    return


def generateTable(dictCells) :

    titleRow = ["sheet", "row", "col", "rows", "cols", "Tstyle", "Tcolor", "Tthick", "Lstyle", "Lcolor", "Lthick", "Bstyle", "Bcolor", "Bthick", "Rstyle", "Rcolor", "Rthick"]
    
    no_spans = ["", ""]
    no_bords = ["", "", "", "", "", "", "", "", "", "", "", ""]
    
    table = []
    table.append(titleRow)

    idx = 2
    for cell, contents in dictCells.items() :
#        print 'Doing #{} : Cell = {}, Content = {}'.format(idx, cell, contents)
        idx += 1
        row = cell.split(':')
        
        if 'span' in contents :
            row.extend(contents['span'][1:-1].split('", "'))
        else :
            row.extend(no_spans)
        
        if 'brdr' in contents :
            row.extend(contents['brdr'][1:-1].split('", "'))
        else :
            row.extend(no_bords)
            
        table.append(row)

    return {'data': table, 'h' :len(table), 'w' : len(titleRow)}        



def generateCSV(dictCells) :
    '''
    Assemble a CSV string of the data. Rows separated by '\n'
    
                This is out of use.  Might not work.
    '''
    no_spans = '"", ""'
    no_bords = '"", "", "", "", "", "", "", "", "", "", "", ""'
    patchRows = '"sheet", "row", "col", "rows", "cols", "Tstyle", "Tcolor", "Tthick", "Lstyle", "Lcolor", "Lthick", "Bstyle", "Bcolor", "Bthick", "Rstyle", "Rcolor", "Rthick"'
    

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



def downloadToFile(url, filename):

    """
    Helper to download possibly large files
    
        the only arg is a url
       this file will go to a temp directory
       the file will also be downloaded
       in chunks and print out how much remains
       
        use it like this :
            downloadToFile("http://localhost/a.zip")
       
    """

    baseFile = os.path.basename(filename)

    # move the file to a more uniq path
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
    


def getConnectionHeader(credentials, url):

    oauth_tokens = googoauth.get_auth_tokens(credentials)
    
    triesLimit = 5
    tries = triesLimit
    while tries > 0 :
        header = {"Authorization": 'Bearer %s' % oauth_tokens.access_token}
        print 'Headr will be : %s.' % header
        try :
            something = urllib2.urlopen(urllib2.Request(url, headers=header)).read()
            googoauth.reopen_store()
            return header
            
        except :
            googoauth.erase_access_token(credentials)
            oauth_tokens = googoauth.refreshToken()

        if tries < triesLimit :
            time.sleep(6)
            print 'Trying again to refresh.  {} tries remain.'.format(tries)
        tries -= 1

    return header


def main():

    from manage_arguments import getArgs
    a = getArgs()
    args = a.args
    credentials = a.credentials

    if args.sheet_ids is None:
        return "Found no sheet ID to use."

    
    url_part  = "https://docs.google.com/feeds/download/spreadsheets/Export?exportFormat=html"
    url_part += "&key=" + args.workbook_key
    url_part += "&gid="

    
    header = getConnectionHeader(credentials, url_part + "0")

    dictCells = {}
    for sheet_id in args.sheet_ids.split(','):

        the_url = url_part + str(sheet_id)

        print 'url : %s' % the_url
        print 'header : %s' % header

        request = urllib2.Request(the_url, headers=header)

        html_file = downloadToFile(request, "temporary.html")

        dictCells = getBorders.getSpreadsheetCellBorders(dictCells, html_file, sheet_id)  #  BORDERS !!!
    
        dictCells = getSpans.getSpreadsheetSpannedCells(dictCells, html_file, sheet_id)  #  SPANS!!!
        

    patchRows = generateTable(dictCells)
    
    if args.immediate_load :

        pushBackToSpreadsheet(credentials, args.workbook_key, args.immediate_load, patchRows)
        print 'Span and Border specifications have been pushed to the sheet "%s" in the workbook "%s".' % (args.immediate_load, args.workbook_key)
    
    return patchRows


if __name__ == "__main__":

    result = main()
	
'''	

Example :

 ./generate_CSV_File.py -ssk 0Asxy-TdDgbjidEpmYVBXaHRUclRtalpHdkw0ek1Md2c -si 6 -ue doowa.diddee@gmail.com -cs 'Zis38NZ_wyBII2Q9xfMRthW-' -il meta_patches -ci 204618981389-fod7457tdhtfvmglt287dg7p30579r0l.apps.googleusercontent.com -am ForDevices -ce alicia.factorepo@gmail.com -sat 'ya29.AHES6ZQUvWvxxxxxxxxxxxxxXWe_0x-aNn6Bxxel9LNT1gv6wX4c1w' -srt '1/hO_TrQqxxxxxxxxxxxxxxxxxxxxxFnSJeVbhZUg9ZvE'

'''

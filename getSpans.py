#!/usr/bin/python
#
from BeautifulSoup import BeautifulSoup

spans_key = 'span'
borders_key = 'brdr'

def extendEdgesOfSpan(key, dictCells, cellSpec):

    dictCells[key][spans_key] = '"{}", "{}"'.format(cellSpec['r'], cellSpec['c'])

    top_left_borders = dictCells[key][borders_key]

    coords = key.split(":")
    sheet = int(coords[0])
    col = int(coords[2])
    row = int(coords[1])
    colSpan = int(cellSpec['c'])
    rowSpan = int(cellSpec['r'])

    if colSpan > 1 :
        for ixRow in range(rowSpan) :
            newRow = row + ixRow
            newCol = col + colSpan - 1
            newKey = '{}:{}:{}'.format(sheet, newRow, newCol)
            '''
            print 'Extend column {}{} ({}) to {}{} ({})'.format( 
                  str(unichr(64 + col))
                , row
                , key
                , str(unichr(64 + newCol))
                , newRow
                , newKey)
            '''
            if newKey not in dictCells :
                dictCells[newKey] = {}
            dictCells[newKey][borders_key] = top_left_borders

    if rowSpan > 1 :
        for ixCol in range(colSpan) :
            newRow = row + rowSpan - 1
            newCol = col + ixCol
            newKey = '{}:{}:{}'.format(sheet, newRow, newCol)
            '''
            print 'Extend row {}{} ({}) to {}{} ({})'.format( 
                str(unichr(64 + col))
                , row
                , key
                , str(unichr(64 + newCol))
                , newRow
                , newKey)
            '''
            if newKey not in dictCells :
                dictCells[newKey] = {}
            dictCells[newKey][borders_key] = top_left_borders

    return dictCells

def getSheetSpannedCells(dictCells, page, sheet):

    page.seek(0)
    soup = BeautifulSoup(page)

    
    table = soup.body.table
            
    idxRow = 0

    for row in table:
        idxCell = 0
        col = {}
        foundInRow = False
        for cell in row:
        
            key = '{}:{}:{}'.format(sheet, idxRow, idxCell)

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
                idxAttr += 1
                
            if found:
                if not key in dictCells :
                    dictCells[key] = {}
                    dictCells[key][spans_key] = '"{}", "{}"'.format(cellSpec['r'], cellSpec['c'])
                else:
                    dictCells = extendEdgesOfSpan(key, dictCells, cellSpec)

            idxCell += 1
        idxRow += 1

    return dictCells


def getSpreadsheetSpannedCells(dictCells, page, sheet):

    return getSheetSpannedCells(dictCells, page, sheet)
    
    result = ''
    for row in spanned_cells:
        line = spanned_cells[row]
        for col in line:
            cell = line[col]
            result += '\n"{}", "{}", "{}", "span", "{}", "{}", "", "", "", ""'.format(sheet, row, col, cell['r'], cell['c'])


    return result

if __name__ == "__main__":
    result = getSpreadsheetSpannedCells("0Asxy-TdDgbjidG5ycTlRYTg0R0RyZzdJbE81MWtWcmc", "6")
    print result

# Example :  ./getSpans.py -k "0AhGdN..S.P.R.E.A.D.S.H.E.E.T...K.E.Y..c1TlE" -s 3 -u "yourUserName@gmail.com" -p "yourPassword" > spans.csv
#    OR
# Example :  ./getSpans.py -k "0AhGdN..S.P.R.E.A.D.S.H.E.E.T...K.E.Y..c1TlE" -s 3 -a "yourGoogleAuthKey" > spans.csv



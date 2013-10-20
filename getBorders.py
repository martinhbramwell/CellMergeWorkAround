#!/usr/bin/python
#
# from BeautifulSoup import BeautifulSoup
from bs4 import BeautifulSoup

import tinycss

borders_key = 'brdr'

def getSheetCellStyles(page):

    parser = tinycss.make_parser('page3')

    stylesheet = parser.parse_stylesheet_file(page)

    container_tokens = {}
    rules = {}
    edge_code = {'top':'T', 'right':'R', 'bottom':'B', 'left':'L'}
    for rule in stylesheet.rules:
        ident_tokens = [ token for token in rule.selector if token.type == 'IDENT' ]
        if ident_tokens[0].value == u'tblGenFixed':
#            print 'Ident {} {}'.format(ident_tokens[0].value, ident_tokens[2].value)
            got = False
            rule_id = ident_tokens[2].value
            temp = {'T':'', 'R':'', 'B':'', 'L':''}
            for declaration in rule.declarations:
                if declaration.name.startswith('border'):
                
                    style = [ attribute.value for attribute in declaration.value if attribute.type == 'IDENT' ][0]
                    color = [ attribute.value for attribute in declaration.value if attribute.type == 'HASH' ][0]
                    thickness = [ attribute.value for attribute in declaration.value if attribute.type == 'DIMENSION' ][0]
                    
#                     print 'Style {} is {} {} {} {}'.format(rule_id, declaration.name, thickness, style, color)
                    
                    if style != 'solid' or color != '#CCC':
                        got = True
                        edge = edge_code[declaration.name[7:]]
                        temp[edge] = {}
                        temp[edge]['type'] = style
                        temp[edge]['color'] = color
                        temp[edge]['thickness'] = str(thickness)
                    
            if got:
                rules[rule_id] = temp
                
#     for key, rule in rules.items() :
#         print 'Rule {} is {}'.format(key, rule)
    return rules

def getSheetCellBorders(dictCells, page, rules, sheet):

    page.seek(0)
    soup = BeautifulSoup(page)

    table = soup.body.table

    idxRow = 0
    cells = {}
    for row in table:
        idxCell = 0
        col = {}
        foundInRow = False            
        for cell in row:
            sep = ''
            for attr, value in cell.attrs:
                if attr == "class":
                    rule = None
                    try :
                        rule = rules[value]
                    except :
                        pass

                    if rule is not None :                       
                        '''
                        dictCells += '\n"{}", "{}", "{}", "border", "", "", "{}", "{}", "{}"'.format(sheet, row, col, )
                        '''
                        key = '{}:{}:{}'.format(sheet, idxRow, idxCell)
                        value = ''
                        for edge in ['T', 'L', 'B', 'R']:

                            if rule[edge] != '' :
                                value += '{}"{}"'.format(sep, rule[edge]['type'])
                                value += ', "{}"'.format(rule[edge]['color'])
                                value += ', "{}"'.format(rule[edge]['thickness'])
                            else :
                                value += '{}"", "", ""'.format(sep)
                            sep = ', '

                        if not key in dictCells :
                            dictCells[key] = {}    
                        dictCells[key][borders_key] = value
                        
            idxCell += 1
        idxRow += 1
    return dictCells
    
    
def getSpreadsheetCellBorders(dictCells, html_file, sheet):

    rules = getSheetCellStyles(html_file)
    dictCells = getSheetCellBorders(dictCells, html_file, rules, sheet)
    
    return dictCells


if __name__ == "__main__":

    result = getSpreadsheetCellBorders({}, "0Asxy-TdDgbjidG5ycTlRYTg0R0RyZzdJbE81MWtWcmc", "6")
    
    print result



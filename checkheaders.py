#!/usr/bin/env python

import sys
import xlrd
import openpyxl


class HeaderChecker(object):
    
    def __init__(self):
                """
        Check the headers for a particular CSHARE File
        """
        self._files = []
        
    def check_file(self, filename):
        if filename.endswith('.xls'):
            # old school xls file
            
            try:
                workbook = xlrd.open_workbook(filename)
                self._files.append(filename)
            except Exception, e:
                print 'Failed to open %s : %s' % (filename, e)
                return
            print 'Opened ' + filename
            print 'Sheets'
            for sheet_name in workbook.sheet_names():
                print '%s' % sheet_name
            print 'Headers'
            for sheet_name in workbook.sheet_names():
                if ('Rules' in sheet_name or 
                    'Decision' in sheet_name or 
                    'Terminology' in sheet_name):
                    # Don't look for headings in pages that we don't care about
                    continue
                
                sheet = workbook.sheet_by_name(sheet_name)
                for row in range(sheet.nrows):
                    content = sheet.row_values(row)
                    if content[0].strip().lower() == 'variable name':
                        # Found the row with the column headings
                        print '%s,%s' % (sheet_name, ','.join([x.strip() for x in content]))
                        
                        
        else:
            # Office 2003 or later xml format 
            try:
                workbook = openpyxl.reader.excel.load_workbook(filename)
                self._files.append(filename)
            except Exception, e:
                import traceback, sys
                
                print 'Failed to open %s : %s' % (filename, e)
                traceback.print_tb(sys.exc_info()[2])
                return
            print 'Opened ' + filename
            print 'Sheets'
            for sheet in workbook.get_sheet_names():
                print '%s' % sheet
            print 'Headers'
            for sheet_name in workbook.get_sheet_names():
                
                if ('Rules' in sheet_name or 
                    'Decision' in sheet_name or
                    'Terminology' in sheet_name):
                    # Don't look for headings in pages we don't care about
                    continue
                sheet = workbook.get_sheet_by_name(sheet_name)
                for row in sheet.rows:
                    if not row[0].value:
                        # blank rows
                        continue

                    if row[0].value.strip().lower() == 'variable name':
                        # Found the row with the column headings                                                         
                        print '%s,%s' % (sheet_name.strip(), ','.join([x.value.strip() for x in row]))

class ContentTemplate(object):
    pass

if __name__  == "__main__":
    import optparse
    parser = optparse.OptionParser('Read the headers out of the tabs in a SHARE spreadsheet and dump to cmd')
    (opts, args) = parser.parse_args()
    ch = HeaderChecker()
    for filename in args:
        ch.check_file(filename)


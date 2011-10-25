#!/usr/bin/env python

import os, sys
"""
XML Spreadsheet format
"""
import openpyxl
"""
BIFF Spreadsheet format
"""
import xlrd

class CodeLoader(object):
    
    def __init__(self, options):
        if options.template:
            self.terminology = self._load_template(options.template)
        else:
            self.terminology = self._load_reference(options.reference)
        
    def _load_template(self, filename):
        """
        Template has terminology in the 'Terms' sheet
        """
        terminology = {}
        print 'Loading %s from Content Template' % filename
        try:
            workbook = xlrd.open_workbook(filename)
                
        except xlrd.biffh.XLRDError, e:
            # use the xml reader
            terminology = self._load_template_xml(filename)
            return terminology
        except IOError, e:
            print 'Failed to open %s - %s' % (filename, e)
            sys.exit(1)
            
        sheet = workbook.sheet_by_name('Terms')
        print 'Loading %s' % sheet_name
        for row in range(1, sheet.nrows):
            content = sheet.row_values(row)
            terminology[content[0].value] = content[1].value
        print 'Loaded %s terms' % len(terminology)
        return terminology

    def _load_template_xml(self, filename):
        """
        use a different reader
        """
        try:
            workbook = openpyxl.reader.excel.load_workbook(filename)
        except Exception, e:
            import traceback, sys
            print 'Failed to open %s : %s' % (filename, e)
            traceback.print_tb(sys.exc_info()[2])
            return
        terminology = {}
        sheet = workbook.get_sheet_by_name('Terms')
        for (idx, row) in enumerate(sheet.rows):
            if idx == 0:
                # ignore header
                continue
            if not row[0].value:
                # blank rows
                continue
            terminology[row[0].value] = {None : ''}.get(row[1].value, row[1].value)
        print 'Loaded %s terms' % len(terminology)
        return terminology
        
    def _load_reference(self, filename):
        terminology = {}
        # Office 2003 or later xml format 
        print 'Loading %s as a reference' % filename
        try:
            workbook = openpyxl.reader.excel.load_workbook(filename)
        except Exception, e:
            import traceback, sys
            
            print 'Failed to open %s : %s' % (filename, e)
            traceback.print_tb(sys.exc_info()[2])
            return
        for sheet_name in workbook.get_sheet_names():
            
            if (not 'Generic' in sheet_name):
                # Only look for headings in pages we don't care about
                continue
            sheet = workbook.get_sheet_by_name(sheet_name)
            for (idx, row) in enumerate(sheet.rows):
                if not row[0].value:
                    # blank rows
                    continue
                
                if row[0].value.strip().lower() == 'variable name':
                    # Found the row with the column headings                                                         
                    for this_row in sheet.rows[idx + 1:]:
                        if this_row[0].value == '' or this_row[0].value is None:
                            continue
                        if this_row[1].value is None:
                            print 'C-code required for %s' % this_row[0].value.strip()
                            terminology[str(this_row[0].value).strip()] = ''
                        else:
                            if str(this_row[1].value).strip() in terminology.items():
                                print 'Duplicated C-code: %s' % str(this_row[1].value).strip()
                            terminology[str(this_row[0].value).strip()] = str(this_row[1].value).strip()
                    else:
                        return terminology

    def dump_map(self, filename):
        if os.path.splitext(filename)[1] == '.xls':
            try:
                workbook = xlrd.open_workbook(filename)
                
            except Exception, e:
                print 'Failed to open %s - %s' % (filename, e)
                return
            
            for sheet_name in workbook.sheet_names():
                if ('Rules' in sheet_name or 
                    'Decision' in sheet_name or 
                    'Terminology' in sheet_name):
                    # Don't look for headings in pages that we don't care about
                    continue

                sheet = workbook.sheet_by_name(sheet_name)
                print 'Loading %s' % sheet_name
                for row in range(sheet.nrows):
                    
                    content = sheet.row_values(row)
                    if str(content[0]).strip().lower() == 'variable name':
                        # Found the row with the column headings
                        for this_row in range(row, sheet.nrows):
                            content = sheet.row_values(this_row)
                            if content[0] is None:
                                print ''
                            print self.terminology.get(content[0].strip(), '')
        else:

            # Office 2003 or later xml format 
            try:
                workbook = openpyxl.reader.excel.load_workbook(filename)
            except Exception, e:
                import traceback, sys
            
                print 'Failed to open %s : %s' % (filename, e)
                traceback.print_tb(sys.exc_info()[2])
                return
            for sheet_name in workbook.get_sheet_names():
            
                if not (sheet_name.startswith('General')
                        or sheet_name.startswith('Generic')):
                    # Only look for headings in pages we don't care about
                    continue
                sheet = workbook.get_sheet_by_name(sheet_name)
                for row in sheet.rows:
                    if not row[0].value:
                        # blank rows
                        continue
                
                    if row[0].value.strip().lower() == 'variable name':
                        # Found the row with the column headings                                                         
                        for this_row in sheet.rows:
                            if (this_row[0].value == '' or this_row[0].value is None):
                                print ''
                                continue
                            print self.terminology.get(this_row[0].value.strip(), '')
                        else:
                            # only process the first sheet
                            return
                        
                        
                        
if __name__ == "__main__":
    import optparse
    parser = optparse.OptionParser()
    parser.add_option('-r',
                      metavar='FILE',
                      action='store',
                      dest='reference',
                      default='')
    
    parser.add_option('-t',
                      metavar='FILE',
                      action='store',
                      dest='template',
                      default='')
    
    (opts, args) = parser.parse_args()
    if not ((opts.reference != '') ^ (opts.template != '') ):
        sys.exit()
    tk = CodeLoader(opts)
    for arg in args:
        import os
        if not os.path.splitext(arg)[1] in ['.xls', '.xlsx']:
            print 'Skipping %s' % arg
            
        print 'Generating %s' % arg
        tk.dump_map(arg)


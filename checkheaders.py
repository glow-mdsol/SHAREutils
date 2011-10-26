#!/usr/bin/env python

import os
import sys
import time

import xlrd
import xlwt
import openpyxl


class HeaderChecker(object):
    
    def __init__(self):
        """
        Check the headers for a particular CSHARE File
        """
        self._files = []
        self._file_obs = {}
        
    def check_files(self):
        import glob
        for filename in glob.glob('*.xls*'):
            if not filename.startswith('20'):
                continue
            self.check_file(filename)
        print 'Loaded %s files' % len(self._files)

    def export_results(self):
        boldtype = xlwt.easyxf("font: bold on; align: wrap on, vert centre, horiz center")
        book = xlwt.Workbook()
        header_sheet = book.add_sheet('Headers by Book')
        index = 0
        print 'Creating Header Export Sheet'
        for template in self._files:
            ctemp = self._file_obs.get(template)
            header_sheet.write(index, 0, template, boldtype)
            index += 1
            for (idx, tab) in enumerate(ctemp.exported()):
                header_sheet.write(index + idx, 0, tab[0], boldtype)
                for c_idx in range(1, len(tab)):
                    header_sheet.write(index + idx, c_idx, tab[c_idx])
            else:
                index += (idx + 1) 
        print 'Creating Header Definition Sheet'
        uniq = book.add_sheet('Unique Headers')
        _uniq = []
        for template in self._file_obs.values():
            for header in template.unique_headers:
                if not header in _uniq:
                    _uniq.append(header)
        uniq.write(0, 0, 'Header', boldtype)
        uniq.write(0, 1, 'Definition', boldtype)
        uniq.write(0, 2, 'Comments', boldtype)
        for (idx, header) in enumerate(sorted(_uniq)):
            uniq.write(idx + 1, 0, header)
        book.save('SHARE Header Report %s.xls' % time.strftime('%Y%m%d'))
        
    def check_file(self, filename):
        if os.path.splitext(filename)[1] == '.xls':
            # old school xls file
            
            try:
                workbook = xlrd.open_workbook(filename)
                self._files.append(filename)
            except Exception, e:
                print 'Failed to open %s : %s' % (filename, e)
                return
            for sheet_name in workbook.sheet_names():
                if ('Rules' in sheet_name or 
                    'Decision' in sheet_name or
                    'Terminology' in sheet_name or
                    'Null' in sheet_name):
                    # Don't look for headings in pages that we don't care about
                    continue
                
                sheet = workbook.sheet_by_name(sheet_name)
                for row in range(sheet.nrows):
                    content = [si(x) for x in sheet.row_values(row)]
                    if si(content[0]).lower() == 'variable name':
                        tab = self._file_obs.setdefault(filename, ContentTemplate(filename))
                        tab.add_headers(sheet_name, content)
                         
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
            for sheet_name in workbook.get_sheet_names():
                
                if ('Rules' in sheet_name or 
                    'Decision' in sheet_name or
                    'Terminology' in sheet_name or
                    'Null' in sheet_name):
                    # Don't look for headings in pages we don't care about
                    continue
                sheet = workbook.get_sheet_by_name(sheet_name)
                for row in sheet.rows:
                    if not si(row[0]):
                        # blank rows
                        continue
                    if si(row[0]).lower() == 'variable name':
                        # Found the row with the column headings
                        content = [si(x) for x in row]
                        tab = self._file_obs.setdefault(filename, ContentTemplate(filename))
                        tab.add_headers(sheet_name, content)                        

def si(content):
    """
    Sanitise Input
    If it gets a None, return '', else return cleaned string
    """
    import types
    if isinstance(content, openpyxl.cell.Cell):
        if isinstance(content.value, types.NoneType):
            return unicode('')
        else:
            return unicode(content.value).strip()
    elif isinstance(content, xlrd.sheet.Cell):
        if content.ctype == xlrd.sheet.XL_CELL_EMPTY:
            return unicode('')
        else:
            return unicode(content.value).strip()
    return unicode(content).strip()

class ContentTemplate(object):

    def __init__(self, filename):
        self.tabs = []
        self.tab_headers = {}

    def add_headers(self, tabname, headers):
        """
        Add headers into the dictionary
        """
        self.tabs.append(tabname)
        self.tab_headers[tabname] = headers

    def exported(self):
        """
        Return the content
        """
        content = []
        for tab in self.tabs:
            _t = [tab]
            _t.extend(self.tab_headers[tab])
            content.append(_t)
        return content
    
    @property
    def unique_headers(self):
        """
        Return the unique headers for a template sheet
        """
        _uniq = []
        for tab_list in self.tab_headers.values():
            for tab_header in tab_list:
                if not tab_header in _uniq:
                    _uniq.append(tab_header)
        return sorted(_uniq)
    
if __name__  == "__main__":
    import optparse
    parser = optparse.OptionParser('Read the headers out of the tabs in a SHARE spreadsheet and dump to cmd')
    (opts, args) = parser.parse_args()
    ch = HeaderChecker()
    if len(args) > 0:
        for filename in args:
            ch.check_file(filename)
    else:
        ch.check_files()

    ch.export_results()

#!/usr/bin/env python

import os, sys

import openpyxl
import xlrd
import xlwt

class ContentReader(object):

    def __init__(self):
        self.files = []
        self.variable_names = {}

    def load_spreadsheet(self, filename):
        if os.path.splitext(filename)[1] not in ['.xlsx', '.xls']:
            print 'Ignoring %s' % filename
            return
        if not filename.startswith('2011'):
            return
        print 'Opening %s' % filename
        self.files.append(filename)
        if os.path.splitext(filename)[1] == '.xlsx':
            self._xlsx_loader(filename)
        else:
            self._xls_loader(filename)
        print 'Loaded %s variables' % len(self.variable_names)
        
    def export_content(self):
        boldtype = xlwt.easyxf("font: bold on; align: wrap on, vert centre, horiz center")
        book = xlwt.Workbook()
        info = book.add_sheet('Information')
        info.write(0, 0, 'Variable Name', boldtype)
        for (f_idx, filename) in enumerate(sorted(self.files)):
            info.write(0, f_idx + 1, filename, boldtype)
        for (idx, term) in enumerate(sorted(self.variable_names.values())):
            info.write(idx+1, 0, term.name)
            for filename in sorted(term.files):
                info.write(idx+1, 1 + sorted(self.files).index(filename), 'X', boldtype)
        
        sheet = book.add_sheet('Terms')
        sheet.write(0, 0, 'Variable Name', boldtype)
        sheet.write(0, 1, 'Code', boldtype) 
        sheet.write(0, 2, 'Definition', boldtype)
        for (idx, term) in enumerate(sorted(self.variable_names.values())):
            sheet.write(idx+1, 0, term.name)
            sheet.write(idx+1, 1, term.code)
            sheet.write(idx+1, 2, term.definition)

        sheet_code = book.add_sheet('Codes')            
        sheet_code.write(0, 0, 'Variable Name', boldtype)
        for (f_idx, filename) in enumerate(sorted(self.files)):
            sheet_code.write(0, f_idx + 1, filename, boldtype)
        for (idx, term) in enumerate(sorted(self.variable_names.values())):
            sheet_code.write(idx+1, 0, term.name)
            for (filename, code) in term.codes.items():
                offset = sorted(self.files).index(filename)
                sheet_code.write(idx+1, offset + 1, code)
        
        sheet_def = book.add_sheet('Definitions')            
        sheet_def.write(0, 0, 'Variable Name', boldtype)
        for (f_idx, filename) in enumerate(sorted(self.files)):
            sheet_def.write(0, f_idx + 1, filename, boldtype)
        for (idx, term) in enumerate(sorted(self.variable_names.values())):
            sheet_def.write(idx+1, 0, term.name)
            for (filename, definition) in term.definitions.items():
                offset = sorted(self.files).index(filename)
                sheet_def.write(idx+1, offset + 1, definition)
        import time
        book.save('SHARE_Combined_Variables_%s.xls' % time.strftime('%Y%m%d'))

    def _xlsx_loader(self, filename):
        """
        Using pyopenxl
        """
        try:
            workbook = openpyxl.reader.excel.load_workbook(filename)
        except Exception, e:
            import traceback, sys
            
            print 'Failed to open %s : %s' % (filename, e)
            traceback.print_tb(sys.exc_info()[2])
            return
        for sheet_name in workbook.get_sheet_names():
            if ('Rules' in sheet_name or 
                'Decision' in sheet_name or 
                'Terminology' in sheet_name or
                'Flavors' in sheet_name):
                # Don't look for headings in pages that we don't care about
                continue
            sheet = workbook.get_sheet_by_name(sheet_name)
            for (idx, row) in enumerate(sheet.rows):
                if not row[0].value:
                    # blank rows
                    continue
                
                if si(row[0]).lower() == 'variable name':
                    headers = [si(x) for x in row]
                    #print 'Headers: %s' % headers
                    # Found the row with the column headings                                                         
                    for this_row in sheet.rows[idx + 1:]:
                        if this_row[0] == '' or this_row[0] is None:
                            continue
                        _content = dict(zip(headers, [si(x) for x in this_row] ))  
                        _name = _content.get('Variable Name')
                        if not _name:
                            continue
                        vble = self.variable_names.setdefault(_name, VariableName(_name))
                        vble.load_row(_content, filename)


    def _xls_loader(self, filename):
        """
        using xlrd
        """
        # old school xls file
        
        try:
            workbook = xlrd.open_workbook(filename)
        except Exception, e:
            print 'Failed to open %s; %s' % (filename, e)
            return
        for sheet_name in workbook.sheet_names():
            if ('Rules' in sheet_name or 
                'Decision' in sheet_name or 
                'Terminology' in sheet_name or
                'Flavors' in sheet_name):
                # Don't look for headings in pages that we don't care about
                continue
            sheet = workbook.sheet_by_name(sheet_name)
            for row_idx in range(sheet.nrows):
                row = sheet.row(row_idx)
                if not si(row[0]):
                    # blank rows
                    continue
                
                if si(row[0]).lower() == 'variable name':
                    headers = [si(x) for x in row]
                    # Found the row with the column headings                                                         
                    for this_row_idx in range(row_idx + 1, sheet.nrows):
                        this_row = sheet.row(this_row_idx)
                        if si(this_row[0]) == '':
                            continue
                        _content = dict(zip(headers, [si(x) for x in this_row] ))  
                        _name = _content.get('Variable Name')
                        if not _name:
                            continue
                        vble = self.variable_names.setdefault(_name, VariableName(_name))
                        vble.load_row(_content, filename)


def si(content):
    """
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
    
class VariableName(object):

    def __init__(self, name):
        # pass in the row as a dictionary agin the headers
        self.name = name
        self.definitions = {}
        self.codes = {}
        self.files = []

    @property
    def code(self):
        return ','.join(dict().fromkeys(self.codes.values()).keys())
        
    @property
    def definition(self):
        return ','.join(dict().fromkeys(self.definitions.values()).keys())
        
    def __cmp__(self, other):
        if self.name.startswith('--') or other.name.startswith('--'):
            # should order --XXXX, then VSXXXX, etc
            if self.name[2:] == other.name[2:]:
                # same fragment
                if self.name.startswith('--'):
                    return -1
                else:
                    return 1
        return cmp(self.name, other.name)
    
    def load_row(self, row, filename):
        # keep track of where things come from
        if not filename in self.files:
            self.files.append(filename)
        # pattern match on the keys for the definitions and c-codes (variable names)
        keys = row.keys()
        code = self._has_codes(keys)

        if code:
            _code = row.get(code)
            self.codes[filename] = _code
        definition = self._has_definition(keys)
        if definition:
            _definition = row.get(definition)
            self.definitions[filename] = _definition

    def _has_codes(self, keys):
        """
        Keys are the dict keys 
        """
        for key in keys:
            if (key.strip().lower().startswith('variable name') and 'c-code' in key.strip().lower()):
                return key
        return ''


    def _has_definition(self, keys):
        """
        Keys are the dict keys 
        """
        for key in keys:
            if 'definition' in key.lower():
                return key
        return ''
    
if __name__ == "__main__":
    import glob
    vmap = ContentReader()
    for excel in glob.glob("*.xls*"):
        vmap.load_spreadsheet(excel)
    vmap.export_content()

    

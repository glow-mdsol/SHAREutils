import os, sys
"""
XML Spreadsheet format
"""
import openpyxl
"""
BIFF Spreadsheet format
"""
import xlrd

"""
Extract content from an unknown state
"""

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

# General Dictionary of codes to 
MAPPING_CODES = {'Mapping to BRIDG Defined Class' : 'BRIDG Defined Class C-Code',
                 'Mapping to BRIDG Defined Class Attribute' : 'BRIDG Defined Class Attribute C-Code',
                 'Mapping to BRIDG Performed Class' : 	'BRIDG Performed Class C-Code',
                 'Mapping to BRIDG Performed Class Attribute' : 'BRIDG Performed Class Attribute C-Code',
                 'Mapping to BRIDG Non-defined/Non-performed Class' : 'BRIDG Non-defined/Non-performed Class C-Code',
                 'Mapping to BRIDG Non-defined/Non-performed Class Attribute' : 'BRIDG Non-defined/Non-performed Class Attribute C-Code',
                 'Variable Name' : 'Variable Name C-Code',
                 'ISO 21090 Datatype' :	'ISO 21090 Datatype C-Code',
                 'ISO 21090 Datatype Constraint' : 'ISO 21090 Datatype Constraint C-Code',
                 'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES' : 'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES C-Code'}

MAPPING_ORDER = ['Variable Name', 'Mapping to BRIDG Defined Class', 'Mapping to BRIDG Defined Class Attribute',
                 'Mapping to BRIDG Performed Class', 'Mapping to BRIDG Performed Class Attribute',
                 'Mapping to BRIDG Non-defined/Non-performed Class', 'Mapping to BRIDG Non-defined/Non-performed Class Attribute',
                 'ISO 21090 Datatype', 'ISO 21090 Datatype Constraint', 'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES']

class ContentSheetChecker(object):

    def __init__(self):
        self.sheets = {}
        self.content = {}

    def load(self, contentsheet):
        pass

    def run_checks(self):
        for check in self.__dict__.keys():
            if check.startswith('_run'):
                pass

    def _run_check_coded_items(self):
        """
        Checks that all coded items have been coded
        """
        pass

    def _run_check_miscoded_items(self):
        """
        Checks that all cases where something could be coded and it is NA or missing, then no code is present
        """
        pass

    def _run_check_instantiation_variables(self):
        """
        check that all variables in the instance sheets (as opposed to generics)
        are present in generic
        """
        pass
    
if __name__ == "__main__":
    pass

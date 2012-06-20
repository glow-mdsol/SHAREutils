#!/usr/bin/env python

import os
import sys
import glob

# XML Spreadsheet format
import openpyxl
# BIFF Spreadsheet format
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

COLUMNS = {"GENERIC" : [u'Variable Name',
                        u'Variable Name C-Code',
                        u'Variable Label',
                        u'SHARE Generic Definition',
                        u'SDTM IG 3.1.2',
                        u'SEND 3.0',
                        u'CDASH V1.1',
                        u'CDASH V1.1 Conceptual Datatype',
                        u'SDTM IG 3.1.2 Datatype',
                        u'Codelist Master',
                        u'Set of Valid Values',
                        u'Assigned Value',
                        u'Mapping to BRIDG Defined Class',
                        u'Mapping to BRIDG Defined Class Attribute',
                        u'BRIDG Defined Class C-Code',
                        u'BRIDG Defined Class Attribute C-Code',
                        u'Mapping to BRIDG Performed Class',
                        u'Mapping to BRIDG Performed Class Attribute',
                        u'BRIDG Performed Class C-Code',
                        u'BRIDG Performed Class Attribute C-Code',
                        u'Mapping to BRIDG Non-defined/Non-performed Class',
                        u'Mapping to BRIDG Non-defined/Non-performed Class Attribute',
                        u'BRIDG Non-defined/Non-performed Class C-Code',
                        u'BRIDG Non-defined/Non-performed Class Attribute C-Code',
                        u'Mapping to BRIDG Planned Class',
                        u'Mapping to BRIDG Planned Class Attribute',
                        u'BRIDG Planned Class C-Code',
                        u'BRIDG Planned Class Attribute C-Code',
                        u'ISO 21090 Datatype', u'ISO 21090 Datatype C-Code',
                        u'ISO 21090 Datatype Component',
                        u'AsCollectedIndicator',
                        u'Observation, ObservationResult, Activity, Relationship',
                        u'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES',
                        u'Description of Observation, ObservationResult or Activity or Relationship - C-Codes',
                        u'Description of Observation, ObservationResult or Activity or Relationship - NON-CODED VALUES',
                        u'NOTES'],
        "TEMPLATE" : [u'Variable Name',
                      u'Variable Label',
                      u'Codelist Master',
                      u'Set of Valid Values',
                      u'Assigned Value',
                      u'Null Flavors',
                      u'Boolean Mapping',
                      u'ISO 21090 Datatype Constraint',
                      u'ISO 21090 Datatype Constraint C-Code',
                      u'ISO 21090 Datatype Constraint Attribute',
                      u'Observation or Activity']
                      }

"""
Rules per the Team
Within header BRIDG version and Domain are mandatory. See below for column rules:
A – populated, no blanks
B – populated, no blanks
C – populated, no blanks
D - populated, no blanks
E - populated, no blanks
F - should be blank
G – populated, no blanks
H – If G = Y, should be populated
I – If E = Y, should be populated
J – okay to be blank
K - okay to be blank
L - does not have to populated
M – should be populated with NA, if not the C-code should be present
N - should be populated with NA or data, no blanks
O – AB may have data or be blank depending on BRIDG entries
If AC is populated, AE should contain DT component
AG, AH, AJ should all be populated, if not with a real value, then NA
AI- C-code should be present (linked with AH)
AK – optional, okay to be blank
o Should be one set of planned and defined for each concept, but some are
implementation specific like SEQ
o Can one variable have both defined and performed? Per Diane, answer = yes
o If any BRIDG class is populated than the DT and C-code fields should be populated
"""

# Columns that must be set
MUSTSET = {u'Variable Name' : True,
           u'Variable Name C-Code' : True,
           u'Variable Label' : True,
           u'SHARE Generic Definition' : True,
           u'SDTM IG 3.1.2' : True,
           u'SEND 3.0' : False,
           u'CDASH V1.1' : True,
           u'CDASH V1.1 Conceptual Datatype' : {u'CDASH V1.1' : "Y"},
           u'SDTM IG 3.1.2 Datatype' : {'SDTM IG 3.1.2' : "Y"},
           u'Codelist Master' : False,
           u'Set of Valid Values' : False,
           u'Assigned Value' : False,
           u'Mapping to BRIDG Defined Class' : False,
           u'Mapping to BRIDG Defined Class Attribute' : False,
           u'BRIDG Defined Class C-Code' : False,
           u'BRIDG Defined Class Attribute C-Code' : False,
           u'Mapping to BRIDG Performed Class': False,
           u'Mapping to BRIDG Performed Class Attribute' : False,
           u'BRIDG Performed Class C-Code' : False,
           u'BRIDG Performed Class Attribute C-Code' : False,
           u'Mapping to BRIDG Non-defined/Non-performed Class' : False,
           u'Mapping to BRIDG Non-defined/Non-performed Class Attribute' : False,
           u'BRIDG Non-defined/Non-performed Class C-Code' : False,
           u'BRIDG Non-defined/Non-performed Class Attribute C-Code' : False,
           u'Mapping to BRIDG Planned Class' : False,
           u'Mapping to BRIDG Planned Class Attribute' : False,
           u'BRIDG Planned Class C-Code' : False,
           u'BRIDG Planned Class Attribute C-Code' : False,
           u'ISO 21090 Datatype' : {u'Mapping to BRIDG Defined Class' : "SET",
                                    u'Mapping to BRIDG Performed Class': "SET",
                                    u'Mapping to BRIDG Non-defined/Non-performed Class': "SET",
                                    u'Mapping to BRIDG Planned Class' : "SET"},
           u'ISO 21090 Datatype C-Code' : False,
           u'ISO 21090 Datatype Component' : {u'ISO 21090 Datatype' : "SET"},
           u'AsCollectedIndicator' : False,
           u'Observation, ObservationResult, Activity, Relationship' : True,
           u'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES' : False,
           u'Description of Observation, ObservationResult or Activity or Relationship - C-Codes' : False,
           u'Description of Observation, ObservationResult or Activity or Relationship - NON-CODED VALUES' : False,
           u'NOTES' : False}

# Set of columns that must not be set
MUSTNOTSET = ["SEND 3.0"]

# Set of columns that must be set or NA
MUSTVALORNA = [u'Mapping to BRIDG Defined Class',
           u'Mapping to BRIDG Defined Class Attribute',
           u'Mapping to BRIDG Performed Class',
           u'Mapping to BRIDG Performed Class Attribute',
           u'Mapping to BRIDG Non-defined/Non-performed Class',
           u'Mapping to BRIDG Non-defined/Non-performed Class Attribute',
           u'Mapping to BRIDG Planned Class',
           u'Mapping to BRIDG Planned Class Attribute',
           u'Observation, ObservationResult, Activity, Relationship',
           u'Description of Observation, ObservationResult or Activity or Relationship - CODED VALUES',
           u'Description of Observation, ObservationResult or Activity or Relationship - NON-CODED VALUES'
           ]

def columnify(set_of_columns):
    """
    return a dict with the column indicies (per Excel)
    """
    mapped = {}
    for (idx, col) in enumerate(set_of_columns):
        prefix = {0 : ''}.get(idx/26, chr(65 + idx/26 - 1))
        mapped["%s%s" % (prefix, chr(65 + idx % 26))] = col
    return mapped


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

    def _null_check(self, row):
        """
        Checks that a column is populated when it should be, dependency rolls up any if X=Y, then Z must be populated
        """
        missing = []
        for (column, content) in row.iteritems():
            if MUSTSET.get(column) != False:
                if content == "":
                    # nothing set
                    if MUSTSET.get(column) is True:
                        missing.append(column)
                    else:
                        # check the dependency
                        for (dep, depval) in MUSTSET.get(column, {}).iteritems():
                            if row.get(dep) != depval or (row.get(dep) != "" and depval == "SET"):
                                missing.append(column)

    def _check_for_trolls(self, row):
        """
        Check for BRIDG assignment
        """
        pass

    def _check_coding_columns(self, row):
        """
        Check that all codable elements have been coded
        """
        pass
    
    def _check_columns(self, sheetname, headers):
        """
        Only expect Templates or Generic
        """
        _lower = sheetname.lower()
        if 'generic' in _lower:
            if not COLUMNS.get('GENERIC') == headers:
                return False
        else:
            if not COLUMNS.get('TEMPLATE') == headers:
                return False
        return True
            
    def _run_generic_checks(self, sheetname, headers):
        """
        Run Checks on the Generic Sheets
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

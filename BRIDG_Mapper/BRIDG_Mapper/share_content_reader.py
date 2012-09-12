#!/usr/bin/env python

import sys
import os
import glob
import re

# XML Spreadsheet format
import openpyxl
# BIFF Spreadsheet format
import xlrd

from common import si

COLUMNS = [u'Variable Name',
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
  u'NOTES']
                        
class SHAREReader(object):
  
  def __init__(self, options={}):
    self.options = options
    self.domains = {}
 
  def get_domain(self, name):
    re
  def load_files(self, path):
    # the SHARE content exists in a number of sheets
    if not os.path.exists(path):
      print "No such path %s" % path
      sys.exit()
    for document in glob.glob(os.path.join(path, "*.xlsx")):
      if not re.match("^[^~](.*) Template.xlsx$", os.path.basename(document)):
        continue
      print "Checking %s" % document
      try:
        workbook = openpyxl.reader.excel.load_workbook(os.path.join(path, document))
      except Exception, exc:
        print "Loading %s failed: %s" % (document, exc)
        sys.exit()
      for sheetname in workbook.get_sheet_names():
        if not re.match("^Generic ([A-Z\-\s]+) Template$", sheetname.strip()):
          continue
        sheet = workbook.get_sheet_by_name(sheetname)
        domain_or_class = None
        # set this to something large
        start_offset = 99
        for (row_idx, row) in enumerate(sheet.rows):
          if si(row[0]) == "Domain":
            for col in [si(x) for x in row[1:]]:
              if col != "":
                if "GENERIC" in col.upper():
                  domain_or_class = SDTMObservationClass(col)
                else:   
                  domain_or_class = SDTMDomain(col)
          elif si(row[0]) == "Variable Name":
            start_offset = row_idx
          elif (row_idx > start_offset):
            if si(row[0]) != "":
              domain_or_class.add_row(dict(zip(COLUMNS, [si(x) for x in row])))
        else:
          self.domains[domain_or_class.name] = domain_or_class
    print "Loaded %s Domains or Observation Classes" % len(self.domains)
    for domain in sorted(self.domains.values(), top_level_sort):
      if isinstance(domain, SDTMObservationClass):
        print "Observation Class %s with %s variables" % (domain.name, len(domain.variables))
      else:
        print "Domain %s with %s variables" % (domain.name, len(domain.variables))

def top_level_sort(x, y):
  if x.__class__.__name__ == y.__class__.__name__:
    return cmp(x.name, y.name)
  return -1 * cmp(x.__class__.__name__, y.__class__.__name__)

class SDTMObservationClass(object):
  
  def __init__(self, domain_name):
    (self.name,) = re.match("GENERIC ([A-Z]+)", domain_name).groups()
    self._variables = {}

  @property
  def variables(self):
    return [x.variable_name for x in self._variables.values()]
    
  def get_variable(self, name):
    self._variables.get(name, self._variables.get(name.replace(self.name, "--")))
    
  def add_row(self, row):
    self._variables.setdefault(row.get("Variable Name"), SDTMVariable(row))
  
class SDTMDomain(object):
  
  def __init__(self, domain_name):
    print domain_name
    (self.name, self.description) = re.match("^([A-Z]+)\s+\((.+)\)$", domain_name).groups()
    self._variables = {}
  
  @property
  def variables(self):
    return [x.qualified(self.name) for x in self._variables.values()]
    
  def get_variable(self, name):
    self._variables.get(name, self._variables.get(name.replace(self.name, "--")))
    
  def add_row(self, row):
    self._variables.setdefault(row.get("Variable Name"), SDTMVariable(row))

class SDTMVariable(object):
  
  def __init__(self, content):
    self._content = content
  
  @property
  def needs_substitution(self):
    return self._content.get(u"Variable Name").startswith("--")
    
  @property
  def variable_name(self):
    return self._content.get(u"Variable Name")
  
  @property
  def variable_label(self):
    return self._content.get(u"Variable Label")
    
  @property
  def bridg_class(self):
    cols = [x for x in COLUMNS if (x.startswith("Mapping to BRIDG") and x.endswith("Class"))]
      return " | ".join([self._content.get(x) for x in cols if self._content.get(x) != ""])   

  @property
  def bridg_attribute(self):
    cols = [x for x in COLUMNS if (x.startswith("Mapping to BRIDG") and x.endswith("Attribute"))]
    return " | ".join([self._content.get(x) for x in cols if self._content.get(x) != ""])  

  def qualified(self, domain):
    if not self.needs_substitution:
      return self.variable_name
    return self.variable_name.replace("--", domain)
    
if __name__ == "__main__":
  t = SHAREReader()
  t.load_files(sys.argv[1])
        
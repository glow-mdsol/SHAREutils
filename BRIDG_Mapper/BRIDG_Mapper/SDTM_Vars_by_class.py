#!/usr/bin/env python

import os
import sys
import openpyxl
from common import si


class SDTMVariablesByClassParser(object):
  
  def __init__(self, options={}):
    self.options = options
    self.observationclasses = {}
    
  def parse(self, filename):
    try:
      workbook = openpyxl.reader.excel.load_workbook(filename)
    except:
      print "Loading %s failed" % filename
      sys.exit()
    for sheetname in workbook.get_sheet_names():
      if sheetname == "Legend-Key":
        continue
      obsclass = self.observationclasses.setdefault(sheetname, SDTMObservationClass(sheetname))
      sheet = workbook.get_sheet_by_name(sheetname)
      for row in sheet.rows:
        # header row
        if si(row[0]).startswith("SDTM v3.1"):
          for (col_idx, col) in enumerate(row):
            if col_idx >= 4:
              obsclass.add_domain(col_idx, si(col))
        elif si(row[1]) == "":
          continue
        else:
          for (col_idx, col) in enumerate([si(x) for x in row]):
            if obsclass.domains.get(col_idx):
              if col.lower() != "not used":
                # add the column to the dataset
                obsclass.domains.get(col_idx).add_variable(row[0],
                                                            row[1],
                                                            row[2],
                                                            row[3], 
                                                            col)
      print "Loaded %s" % sheetname
    for obclass in self.observationclasses.itervalues():
      for domain in obclass.domains.itervalues():
        print "%s: %s (%s variables)" % (obclass.name, domain.name, len(domain.variables))


class SDTMObservationClass(object):
  # a tab
  
  def __init__(self, name):
    self.name = name
    self.domains = {}
    
  def add_domain(self, offset, domain):
    self.domains[offset] = SDTMDomain(domain)
    
class SDTMVariable(object):
  
  def __init__(self, name, label, datatype):
    self.name = name
    self.label = label
    self.datatype = datatype

class SDTMDomain(object):
  
  VARIABLES = {}
  
  def __init__(self, name):
    self.name = name
    self._variables = []
  
  @property
  def variables(self):
    return [x[1] for x in self._variables]
    
  @classmethod
  def get_variable(cls, name, label, datatype):
    if cls.VARIABLES.get(name) is None:
      cls.VARIABLES[name] = SDTMVariable(name, label, datatype)
    return cls.VARIABLES.get(name)
    
  def add_variable(self, index, name, label, datatype, permissibility):
    self._variables.append([index, 
                            self.get_variable(name, label, datatype), 
                            permissibility])

if __name__ == "__main__":
  parser = SDTMVariablesByClassParser()
  parser.parse(sys.argv[1])
#!/usr/bin/env python

"""
Super module to pull it all together and dump it out
"""
import openpxyl

import BRIDG_Mapper
import SDTM_Vars_by_class
import share_content_reader

# output columns
RAW_COLUMNS = [u'Mapped Group Name', u'Mapped Element Name', u'Element Type', 
               u'Data Type', u'Card.', u'Definition and Semantics', u'Custom', 
               u'Status ', u'Review by', u'Comments / Issues / Rationale', 
               u'Mapping Path / Derivation', u'Class Name', u'Element Name', 
               u'Element Type', u'Data Type', u'Card.', u'Definition & Usage', 
               u'Constraints', u'Revised name']
							 
class ConsolidatedIntoOneEasyMonthlyPayment(object):
  
  def __init__(self, options={}):
    self.options = options
    
  
  def consolidate(self):
    pass
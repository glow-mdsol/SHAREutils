
import win32com.client
import os
import sys

RAW_COLUMNS = [u'Mapped Group Name', u'Mapped Element Name', u'Element Type', 
               u'Data Type', u'Card.', u'Definition and Semantics', u'Custom', 
               u'Status ', u'Review by', u'Comments / Issues / Rationale', 
               u'Mapping Path / Derivation', u'Class Name', u'Element Name', 
               u'Element Type', u'Data Type', u'Card.', u'Definition & Usage', 
               u'Constraints', u'Revised name']
COLUMNS = []
for (idx, col) in enumerate(RAW_COLUMNS):
  if idx <= RAW_COLUMNS.index(u'Custom'):
    COLUMNS.append("Source %s" % col)
  elif idx >= RAW_COLUMNS.index(u'Class Name'):
    COLUMNS.append("Source %s" % col)
  else:
    COLUMNS.append(str(col))

class BRIDGMappingSheetLoader(object):
  
  def __init__(self, options={}):
    self.options = options
    self._app = None

  @property
  def xlApp(self):
    if self._app is None:
      self._app = win32com.client.Dispatch("Excel.Application")
    return self._app

#Workbook Open(
#	string Filename,
#	Object UpdateLinks,
#	Object ReadOnly,
#	Object Format,
#	Object Password,
#	Object WriteResPassword,
#	Object IgnoreReadOnlyRecommended,
#	Object Origin,
#	Object Delimiter,
#	Object Editable,
#	Object Notify,
#	Object Converter,
#	Object AddToMru,
#	Object CorruptLoad
#)
  def load_template(self, workbook, password="bridg"):
    if not os.path.exists(workbook):
      print "No such file %s" % workbook
      return
    try:
      book = self.xlApp.Workbooks.Open(os.path.join(os.getcwd(), workbook), 2, True, None, password)
    except Exception, exc:
       print "Loading workbook failed: " + str(exc)
       return
    bridg_mapping = BRIDGMapping(self.options)
    for worksheet in book.Worksheets:
      if worksheet.Name == "Mappings":
        for row_idx in range(3, worksheet.UsedRange.Rows.Count + 1):
          if worksheet.Cells(row_idx, 1).Value in [None, ""]:
            continue
          bridg_mapping.add_row(dict(zip(COLUMNS, [worksheet.Cells(row_idx, x).Value for x in range(1, len(COLUMNS))])))
        print "Loaded %s Rows" % len(bridg_mapping.rows)
        print "%s elements need mapping" % len(filter(lambda x: x.needs_mapping is True, bridg_mapping.rows))
        for todo in filter(lambda x: x.needs_mapping is True, bridg_mapping.rows):
          print "%s - %s" % (todo.domain, todo.element)
    book.Close(False)

class BRIDG_Source(object):
  
  def __init__(self):
    pass

  
class BRIDGMapping(object):

  def __init__(self, options={}):
    self.options = options
    self.rows = []

  def add_row(self, row):
    self.rows.append(MappingRow(row))

class MappingRow(object):

  def __init__(self, datarow={}):
    self.data = datarow

  def __eql__(self, other):
    return self.data == other.data

  @property
  def needs_mapping(self):
    return self.mapping_path in [None, ""]

  @property
  def domain(self):
    return self.data.get('Source Mapped Group Name', "")

  @property
  def element(self):
    return self.data.get('Source Mapped Element Name', "")
  
  @property
  def element_type(self):
    return self.data.get("Source Element Type", "")

  @property
  def data_type(self):
    return self.data.get("Source Data Type", "")

  @property
  def cardinality(self):
    return self.data.get("Source Cardinality", "")

  @property
  def definition_and_semantics(self):
    return self.data.get("Source Definition and Semantics", "")

  @property
  def custom(self):
    return self.data.get("Source Custom", "")

  @property
  def status(self):
    return self.data.get("Mapping Status", "")

  @property
  def mapping_path(self):
    return self.data.get("Mapping Path / Derivation", "")

  @property
  def bridg_class_name(self):
    return self.data.get("BRIDG Class Name", "")

  @property
  def bridg_element_name(self):
    return self.data.get("BRIDG Element Name", "")

  @property
  def bridg_element_type(self):
    return self.data.get("BRIDG Element Type", "")

  @property
  def bridg_data_type(self):
    return self.data.get("BRIDG Data Type", "")

  @property
  def bridg_cardinality(self):
    return self.data.get("BRIDG Cardinality", "")

  @property
  def bridg_definition_and_usage(self):
    return self.data.get("BRIDG Definiton & Usage", "")

  @property
  def bridg_constraints(self):
    return self.data.get("BRIDG Constraints", "")

  @property
  def bridg_revised_name(self):
    return self.data.get("BRIDG Revised Name", "")

class BRIDG_Super(object):

  def __init__(self, name, definition_notes, tags):
    self.name = name
    self.definition = definition_notes
    self.tags = self._derive_tags(tags)

  def _derive_tags(self, tags):
    content = tags.split(';')
    pass

class BRIDG_Class(BRIDG_Super):

  def __init__(self, name, element_type, data_type, cardinality, definition_notes, contstraints, tags):
    super(BRIDG_Class_Element, self).__init__(name, definition_notes, tags)
    self.element_type = element_type

class BRIDG_Class_Element(BRIDG_Super):

  def  __init__(self, name, element_type, data_type, cardinality, definition_notes, constraints, tags):
    super(BRIDG_Class_Element, self).__init__(name, definition_notes, tags)
    self.element_type = element_type
    self.data_type = data_type
    self.cardinality = cardinality
    self.constraints = self._derive_constraints(constraints)
    
  def _derive_constraints(self, contstraints):
    return []


if __name__ == "__main__":
  loader = BRIDGMappingSheetLoader()
  loader.load_template(sys.argv[1])


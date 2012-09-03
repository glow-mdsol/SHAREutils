
import win32com.client
import os
import sys
import time
import cPickle
import re

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
    print "Loading %s" % workbook
    if not os.path.exists(workbook):
      print "No such file %s" % workbook
      return
    try:
      book = self.xlApp.Workbooks.Open(os.path.join(os.getcwd(), workbook), 2, True, None, password)
    except Exception, exc:
       print "Loading workbook failed: " + str(exc)
       return
    for worksheet in book.Worksheets:
      print "Scanning %s" % worksheet.Name
      if worksheet.Name == "Mappings":
        bridg_mapping = BRIDGMapping(self.options)
        for row_idx in range(3, worksheet.UsedRange.Rows.Count + 1):
          if worksheet.Cells(row_idx, 1).Value == "Mapped Group Name":
            continue
          elif worksheet.Cells(row_idx, 1).Value in ["", None]:
            break
          bridg_mapping.add_row(dict(zip(COLUMNS, [worksheet.Cells(row_idx, x).Value for x in range(1, len(COLUMNS))])))
        print "BRIDG to SDTM"
        print "Loaded %s Rows" % len(bridg_mapping.rows)
        print "%s elements need mapping" % len(filter(lambda x: x.needs_mapping is True, bridg_mapping.rows))
        for todo in filter(lambda x: x.needs_mapping is True, bridg_mapping.rows):
          print "%s - %s" % (todo.domain, todo.element)
      elif worksheet.Name == "Mapped Specification source":
        mapped_spec = Mapped_Specification_source()
        print "Used: %s x %s" % (worksheet.UsedRange.Rows.Count, worksheet.UsedRange.Columns.Count)
        for row_idx in range(3, worksheet.UsedRange.Rows.Count + 1):
          if worksheet.Cells(row_idx, 1).Value in [None, ""]:
            break
          #print "Checking Row %s of %s" % (row_idx, worksheet.UsedRange.Rows.Count)
          mapped_spec.add_row([worksheet.Cells(row_idx, x).Value for x in range(1, worksheet.UsedRange.Columns.Count + 1)])
        print "Mapped Specification source"
        print "Loaded %s Rows" % len(mapped_spec.rows)
      
      elif worksheet.Name == "BRIDG source":
        bridg_source = BRIDG_Source()
        print "Used: %s x %s" % (worksheet.UsedRange.Rows.Count, worksheet.UsedRange.Columns.Count)
        for row_idx in range(3, worksheet.UsedRange.Rows.Count+1):
          if worksheet.Cells(row_idx, 1).Value in [None, ""]:
            break
          #print "Checking Row %s of %s" % (row_idx, worksheet.UsedRange.Rows.Count)
          bridg_source.add_row([worksheet.Cells(row_idx, x).Value for x in range(1, worksheet.UsedRange.Columns.Count + 1)])
        print "BRIDG source"
        print "Loaded %s classes" % len(bridg_source.classes)
        pass
    
    book.Close(False)
    # write to a file
    dataset = {}
    dataset["BRIDG Source"] = bridg_source
    dataset["Mapped Specification source"] = mapped_spec
    dataset["BRIDG Mapping"] = bridg_mapping
    bridg_pickle = open("BRIDG_MAPPING_%s.pkl" % time.strftime("%Y%m%d"), 'wb')
    cPickle.dump(dataset, bridg_pickle)
    bridg_pickle.close()


class Mapped_Specification_source(object):

  def __init__(self, options={}):
    self.options = options
    self.rows = []

  def add_row(self, row):
    self.rows.append(Mapped_Specification_row(row[0],
                      row[1],
                     row[2],
                     row[3],
                     row[4],
                     row[5],
                     row[6],
                     row[7]))

class Mapped_Specification_row(object):

  def __init__(self, mapped_group_name, mapped_element_name, 
               element_type, data_type, cardinality,
               definition_and_semantics, custom, 
               mapping_tag_value):
    self.mapped_group_name = mapped_group_name
    self.mapped_element_name = mapped_element_name
    self.element_type = element_type
    self.data_type = data_type
    self.cardinality = cardinality
    self.definition_and_semantics = definition_and_semantics
    self.custom = custom
    self.mapping_tag_value = mapping_tag_value
  
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

class BRIDG_Source(object):

  def __init__(self, options={}):
    self.options = options
    self.classes = {}

  def add_row(self, row):
    # all columns
    _COLUMNS=["Hidden", "BRIDG class name",	"BRIDG class element name",	"Element Type",
              "Data Type",	"Cardinality",	"Definition & Notes",	"Constraints",	"Tags"]
    # map it to a dictionary
    data = dict(zip(_COLUMNS, row))
    if data["Element Type"] == "Class":
      # we are a class
      self.classes[data["BRIDG class name"]] = BRIDG_Class(data["BRIDG class name"],
                                                           data["Element Type"],
                                                           data["Data Type"],
                                                           data["Cardinality"],
                                                           data["Definition & Notes"],
                                                           data["Constraints"],
                                                           data["Tags"])
    elif data["Element Type"] == "Attrib":
      try:
        self.classes.get(data["BRIDG class name"]).add_attribute(data)
      except AttributeError, error:
        print "Error with dataset: %s" % data
        print "Error message: %s" % error
        sys.exit()
    elif data["Element Type"] == "Assoc":
      try:
        self.classes.get(data["BRIDG class name"]).add_association(data)
      except AttributeError, error:
        print "Error with dataset: %s" % data
        print "Error message: %s" % error
        sys.exit()
    elif data["Element Type"] == "Gen":
      try:
        self.classes.get(data["BRIDG class name"]).add_generalization(data)
      except AttributeError, error:
        print "Error with dataset: %s" % data
        print "Error message: %s" % error
        sys.exit()


class BRIDG_Super(object):
  BRIDGClasses = {}

  def __init__(self, name, definition_notes, tags):
    self.name = name
    self.raw_tags = tags
    self.definition = definition_notes
    self.tags = self._derive_tags(tags)

  def _derive_tags(self, tags):
    _derived = {}
    if not tags is None:
      catch_re = re.compile("^(.+):(.+)=(.+)")
      for tag in tags.strip().split(';'):
        if catch_re.match(tag):
          (domain, attribute, value) = catch_re.match(tag).groups()
          _d = _derived.setdefault(domain, {})
          _d[attribute] = value
    return _derived


class BRIDG_Class(BRIDG_Super):
  
  # TODO: Generalisations and Associations - how to do the links?

  def __init__(self, name, element_type, data_type, cardinality, definition_notes, contstraints, tags):
    super(BRIDG_Class, self).__init__(name, definition_notes, tags)
    self.element_type = element_type
    self.attributes = {}
    self.specializations = []
    self.generalizations = []
    self.associations = []

  def add_attribute(self, data):
    self.attributes[data["BRIDG class element name"]] = BRIDG_Class_Element(data["BRIDG class element name"],
                                                                            data["Element Type"],
                                                                            data["Data Type"],
                                                                            data["Cardinality"],
                                                                            data["Definition & Notes"],
                                                                            data["Constraints"],
                                                                            data["Tags"])

  def add_association(self, data):
    # TBC
    #assoc_re = re.compile("^(.+)\((.+)\)$")
    #(attribute, classname) = assoc_re.match(data["BRIDG class element name"]).groups()

    self.associations.append(data)

  def add_generalization(self, data):
    # generalization of this instance
    self.generalizations.append(data)

  def add_specialization(self, bridg_class):
    # specialization of this instance
    self.specializations.append(bridg_class)

class BRIDG_Class_Element(BRIDG_Super):

  def  __init__(self, name, element_type, data_type, 
                cardinality, definition_notes, 
                constraints, tags):
    super(BRIDG_Class_Element, self).__init__(name, definition_notes, tags)
    self.element_type = element_type
    self.data_type = data_type
    self.cardinality = cardinality
    self.raw_constraints = constraints
    self.constraints = self._derive_constraints(constraints)
    self.associations = []
        
  def _derive_constraints(self, contstraints):
    return []


if __name__ == "__main__":
  loader = BRIDGMappingSheetLoader()
  loader.load_template(sys.argv[1])


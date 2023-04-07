import csv

from docx import Document
from docx.oxml import OxmlElement
from docx.table import Table
from docx.shared import Inches


# read csv for rule set data
def readRuleSetCSV():
    with open('./files/rules.csv', 'r') as f:
        reader = csv.reader(f)
        rules = list(reader)
    return rules

# define a class for rule set
class RuleSet:
    clarify = ''
    fieldName = ''
    fieldDescription = ''
    alternativeName = ''
    fieldType = ''
    cerLocation = ''
    locationType = ''
    sscpDestinationPartA = ''
    sscpDestinationPartB = ''
    languageConversionNeededForPartB = ''
    note = ''
    comment = ''

    def __init__(self, clarify, fieldName, fieldDescription, alternativeName, fieldType, cerLocation, locationType,
                 sscpDestinationPartA, sscpDestinationPartB, languageConversionNeededForPartB, note, comment):
        self.clarify = clarify
        self.fieldName = fieldName
        self.fieldDescription = fieldDescription
        self.alternativeName = alternativeName
        self.fieldType = fieldType
        self.cerLocation = cerLocation
        self.locationType = locationType
        self.sscpDestinationPartA = sscpDestinationPartA
        self.sscpDestinationPartB = sscpDestinationPartB
        self.languageConversionNeededForPartB = languageConversionNeededForPartB
        self.note = note
        self.comment = comment


# parse csv and create rule set object
def parseRuleSetCSV():
    rules = readRuleSetCSV()
    ruleSetList = []

    # ignore first row
    rules.pop(0)

    for rule in rules:
        ruleSet = RuleSet(rule[0], rule[1], rule[2], rule[3], rule[4], rule[5], rule[6], rule[7], rule[8], rule[9],
                          rule[10], rule[11])
        ruleSetList.append(ruleSet)
    return ruleSetList

# print all rule set
def printRuleSet():
    ruleSetList = parseRuleSetCSV()
    for ruleSet in ruleSetList:
        print(ruleSet.clarify)
        print(ruleSet.fieldName)
        print(ruleSet.fieldDescription)
        print(ruleSet.alternativeName)
        print(ruleSet.fieldType)
        print(ruleSet.cerLocation)
        print(ruleSet.locationType)
        print(ruleSet.sscpDestinationPartA)
        print(ruleSet.sscpDestinationPartB)
        print(ruleSet.languageConversionNeededForPartB)
        print(ruleSet.note)
        print(ruleSet.comment)

#printRuleSet()

# read input docx file
def readInputDocxFile():
    document = Document('./files/input.docx')
    return document

# print all heading 1 and heading 2
def printHeading(document):
    for para in document.paragraphs:
        if para.style.name == 'Heading 1':
            print(para.text)
        if para.style.name == 'Heading 2':
            print(para.text)

#printHeading(readInputDocxFile())

# read template docx file
def readTemplateDocxFile():
    document = Document('./files/SSCP_Template_Rev10.docx')
    return document

#printHeading(readTemplateDocxFile())

# save the updated template docx file
def saveDocxFile(document):
    document.save('./files/output.docx')


# for each rule set locate data in input docx file, update template docx file
def enforceRules():
    ruleSetList = parseRuleSetCSV()
    inputDocument = readInputDocxFile()
    templateDocument = readTemplateDocxFile()

    for ruleSet in ruleSetList:
        # if location type starts from "table(" then it is a table
        if ruleSet.locationType.startswith("Table ("):
            # get table name from location type
            # remove "Table (" and ")"
            tableName = ruleSet.locationType[7:-1].split(")")[0]
            # find table number by locating the table name
            tableNumber = 0

            inputDocument.add_paragraph(tableName)

            for para in inputDocument.paragraphs:
                # get caption of table
                # get reference 
                # get table number from caption reference 
                if para.style.name == 'Caption':
                    if tableName in para.text:
                        # read xml of para

                        # get reference
                        xmlObjRef = para._element.xml

                        # convert to xml object
                        xmlObj = OxmlElement(xmlObjRef)


                        tableNumber = int(para.text.split(" ")[-1])
                        break

            
            # get table from input docx file
            table = inputDocument.tables[tableNumber]

            # print table heading column 0
            for row in table.rows:
                print(row)


    #saveDocxFile(templateDocument)


enforceRules()
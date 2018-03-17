import traceback
import os
import subprocess
from pdftabextract.common import read_xml, parse_pages
from pprint import pprint
import string
import re
import csv

# Thanks to
# https://datascience.blog.wzb.eu/2017/02/16/data-mining-ocr-pdfs-using-pdftabextract-to-liberate-tabular-data-from-scanned-documents/

STATEMENT_DIR = "statements"
XML_DIR = "xml"
PRINTABLE = set(string.printable)

def getXMLPath(filepath):
	filename, file_extension = os.path.splitext(filepath)
	path, filename = os.path.split(filename)
	return os.path.join(XML_DIR, filename + ".xml")

def processNumber(s):
	if len(s) == 1:
		try:
			return int(s)
		except:
			pass
	try:
		asciiS = str(filter(lambda x: x in PRINTABLE, s))
		stripped = re.sub(r'\s+', '', asciiS).replace("*", "").replace("$", "")
		return float(stripped[:-2] + "." + stripped[-2:])
	except:
		return s

def processRow(r):
	return r[0:1] + map(lambda s: processNumber(s), r[1:])

def getDeductionRows(allRows):
	dr = []
	add = False
	for r in allRows:
		if r[0] == 'Statutory':
			add = True
		if r[0] == 'Net Pay':
			add = False
		if add:
			dr.append(r)
	return dr

def filterDeductionRows(dRows):
	newRows = []
	for r in dRows:
		if r[0] in ["Mcttax", "Hlth Plan Value", "Max Elig/Comp"]:
			continue
		if not isinstance(r[:2][-1], float):
			continue
		newRows.append(r)
	return newRows

def parseFile(filepath):
	xml_path = getXMLPath(filepath)
	if os.path.isfile(xml_path):
		# print "already xml'd"
		pass
	else:
		# Extract xml, don't need the images
		pdftohtml_command = "pdftohtml -c -hidden -xml '{}' '{}'".format(filepath, xml_path)
		subprocess.call(pdftohtml_command, shell=True)
		subprocess.call("rm xml/*.png", shell=True)

	# Load the XML that was generated with pdftohtml
	xmltree, xmlroot = read_xml(xml_path)

	# Parse it and generate a dict of pages
	pages = parse_pages(xmlroot)

	p = pages[1]
	textBoxes = p['texts']
	textBoxes.sort(key=lambda b: (b['top'], b['left']))

	# Group by row
	textBoxesByRow = {}
	for b in textBoxes:
		if b['top'] not in textBoxesByRow:
			textBoxesByRow[b['top']] = []
		textBoxesByRow[b['top']].append(b)

	allRows = []
	for t in sorted(textBoxesByRow):
		row = map(lambda b: b['value'], sorted(textBoxesByRow[t], key = lambda b: b['left']))
		# expand items with commas - ["asdf", "bla,oie", "fdsa"] -> ["asdf", "bla", "oie", "fdsa"]
		for i in xrange(len(row)):
			if ',' in row[i]:
				a = row[i].split(',')
				row[i:i+1] = a
		# print t, row
		row = map(lambda b: processNumber(b), row)
		allRows.append(row)
		# print t, row
	return allRows

def findValues(allRows):
	rowsByFirst = dict((r[0], r) for r in allRows)
	rowsBySecond = dict((r[1], r) for r in allRows if len(r) > 1)

	deductions = {}
	dRows = filterDeductionRows(getDeductionRows(allRows))
	for r in dRows:
		if len(r) >= 3:
			deductions[r[0]] = r[1]
		elif len(r) == 2:
			pass
		else:
			print r

	fedRow = rowsByFirst['Federal:']
	stateRow = rowsByFirst['NY:']
	localRow = rowsByFirst.get('New York Cit:', None)

	# print localRow

	getExtra = lambda r: r[2].replace("$","").replace(" Additional Tax", "") if len(r) == 3 else 0
	payRow = [None, None]
	if 'Pay Date:' in rowsByFirst:
		payRow = rowsByFirst['Pay Date:']
	elif 'Pay date:' in rowsByFirst:
		payRow = rowsByFirst['Pay date:']
	elif 'Pay Date:' in rowsBySecond:
		payRow = rowsBySecond['Pay Date:'][1:]

	statementRow = {
		"Pay Date": payRow[1],
		"Fed": int(fedRow[1]),
		"Fed+": getExtra(fedRow),
		"NY": int(stateRow[1]),
		"NY+": getExtra(stateRow),
		"NYC": int(localRow[1]) if localRow else None,
		"NYC+": getExtra(localRow) if localRow else None,
		"Gross Pay": rowsByFirst['Regular'][2],
		"Net Pay": rowsByFirst['Net Pay'][1],
		"Deductions": deductions
	}
	# pprint(statementRow)
	return statementRow


def doAll():
	allDeductionTypes = set()
	allStatementRows = []
	for filename in sorted(os.listdir(STATEMENT_DIR)):
		filepath = os.path.join(STATEMENT_DIR, filename)
		# print filepath
		allRows = parseFile(filepath)
		try:
			statementRow = findValues(allRows)
		except:
			for r in allRows:
				print r
			traceback.print_exc()
		allDeductionTypes.update(statementRow["Deductions"].keys())
		allStatementRows.append(statementRow)

	sortedDeductions = sorted(list(allDeductionTypes))
	headers = ["Pay Date", "F", "+", "S", "+", "L", "+", "Gross Pay"] + sortedDeductions + ["Net Pay"]
	csvfile = open('payStatements.csv', 'wb')
	csvwriter = csv.writer(csvfile, delimiter=',', quotechar='\'', quoting=csv.QUOTE_MINIMAL)
	csvwriter.writerow(headers)

	for row in allStatementRows:
		csvRow = [
			row["Pay Date"],
			row["Fed"],
			row["Fed+"],
			row["NY"],
			row["NY+"],
			row["NYC"] if row["NYC"] is not None else "",
			row["NYC+"] if row["NYC+"] is not None else "",
			row["Gross Pay"]
		]

		d = row["Deductions"]
		csvRow.extend([d[k] if k in d else "" for k in sortedDeductions])
		csvRow.extend([row["Net Pay"]])

		csvwriter.writerow(csvRow)

		# print statementRow["Pay Date"]
		# print "  F: {} ({}) S: {} ({}) L: {} ({})".format(
		# 	statementRow["Fed"],
		# 	statementRow["Fed+"],
		# 	statementRow["NY"],
		# 	statementRow["NY+"],
		# 	statementRow["NYC"] if statementRow["NYC"] is not None else "-",
		# 	statementRow["NYC+"] if statementRow["NYC+"] is not None else "-")
		# print "  Gross: {}, Net: {}".format(statementRow["Gross Pay"], statementRow["Net Pay"])
		# print "  " + ', '.join(['{0}: {1}'.format(k, v) for k,v in statementRow["Deductions"].iteritems()])


	# pprint(statementRow)

# testFile = "statements/PayStatement-2017-09-29.pdf"
# testFile = "statements/PayStatement-2015-04-17.pdf"
# testFile = "statements/PayStatement-2016-05-13.pdf"
# testFile = "statements/PayStatement-2017-08-18.pdf"
# a = parseFile(testFile)
# s = findValues(a)

doAll()
print "Done"
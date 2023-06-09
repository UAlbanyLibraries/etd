import os
import csv
from tqdm import tqdm
from lxml import etree

etd_root = "\\\\Lincoln\\Library\\ETDs\\Unzipped"


xml_list = []
for thing in os.listdir(etd_root):

	# exclude supplemental materials directories and pdfs themselves
	if thing.lower().endswith(".xml"):
		xml_list.append(thing)

count = 0

pubs = {}
embargos = {}
third = {}
unknown = {}

embargoDates = []

for etd in xml_list:
	# get the full path
	xml_file = os.path.join(etd_root, etd)
	xml_id = etd.split("_DATA")[0]

	count += 1

	tree = etree.parse(xml_file)
	# Get the root <DISS_submission> element
	root = tree.getroot()

	if root.attrib['publishing_option'] in pubs.keys():
		pubs[root.attrib['publishing_option']] += 1
	else:
		pubs[root.attrib['publishing_option']] = 1
	if root.attrib['third_party_search'] in third.keys():
		third[root.attrib['third_party_search']] += 1
	else:
		third[root.attrib['third_party_search']] = 1

	dates = root.xpath("//DISS_description/DISS_dates/DISS_comp_date")

	embargo = root.attrib['embargo_code']
	#if embargo != "0" and embargo != "4":
	if embargo in embargos.keys():
		embargos[embargo] += 1
	else:
		embargos[embargo] = 1
	if embargo == "3":
		print ("embargo --> " + root.xpath("//DISS_description/DISS_dates/DISS_comp_date")[0].text)
	elif embargo == "4":
		names = []
		if root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_fname")[0].text:
			names.append(root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_fname")[0].text)
		if root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_middle")[0].text:
			names.append(root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_middle")[0].text)
		if root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_surname")[0].text:
			names.append(root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_surname")[0].text)
		if root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_suffix")[0].text:
			names.append(root.xpath("//DISS_authorship/DISS_author/DISS_name/DISS_suffix")[0].text)

		embargoList = [" ".join(names), root.xpath("//DISS_description/DISS_dates/DISS_comp_date")[0].text, root.xpath("//DISS_description/DISS_title")[0].text]
		embargoDates.append(embargoList)
		if root.attrib['publishing_option'] in unknown.keys():
			unknown[root.attrib['publishing_option']] += 1
		else:
			unknown[root.attrib['publishing_option']] = 1
		if root.attrib['third_party_search'] in unknown.keys():
			unknown[root.attrib['third_party_search']] += 1
		else:
			unknown[root.attrib['third_party_search']] = 1

	contact = root.xpath("//DISS_description/DISS_author/DISS_contact")
	#if len(contact) == 0:
	#	print (dates[0].text)

	advisors = dates = root.xpath("//DISS_description/DISS_advisor")
	if len(advisors) > 2:
		print (advisors)


print (pubs)
print (embargos)
print (third)
print (unknown)


for e in embargoList:
	print (e)
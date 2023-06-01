import os
import openpyxl
from tqdm import tqdm
from lxml import etree
from datetime import datetime

embargoSheetFile = "\\\\Lincoln\\Library\\ETDs\\MMSIDs\\ETD Embargoed Submissions 1-2008 to 5-19-23.xlsx"
catalogEmbargosFile = "\\\\Lincoln\\Library\\ETDs\\MMSIDs\\MMSID_2008_2022EMBARGOED.xlsx"
etdDir = "\\\\Lincoln\\Library\\ETDs\\Unzipped"

compare_list = []

wb = openpyxl.load_workbook(filename = embargoSheetFile)

present = datetime.now().date()

grad_school_count = 0
sheet = wb.active

for row in sheet:
	title = row[2].value
	lname = row[5].value
	fname = row[6].value
	embargo = row[10].value
	degree_date = row[9].value


	if "-" in embargo:
		#print (embargo)
		embargo_date = datetime.strptime(embargo, "%Y-%m-%d").date()

		if embargo_date > present:
			grad_school_count += 1
			#print (embargo)

	compare_list.append([lname, fname])

print (f"Grad School data: {grad_school_count}")

for etd in os.listdir(etdDir):
	if etd.endswith(".xml"):
		xml_file = os.path.join(etdDir, etd)
		xml_id = etd.split("_DATA")[0]

		tree = etree.parse(xml_file)
		# Get the root <DISS_submission> element
		root = tree.getroot()

		embargo = root.attrib['embargo_code']

		for author in root.find("DISS_authorship"):

			name_count = 0
			fullname = []
			for name in author:
				if name.tag == "DISS_name":
					name_count += 1
					if name_count != 1:
						raise ValueError(f"{etd} has multiple names")
					# safely build the name
					if name.find("DISS_fname").text:
						first_name = name.find("DISS_fname").text
						fullname.append(name.find("DISS_fname").text.title())
					if name.find("DISS_middle").text:
						fullname.append(name.find("DISS_middle").text.title())
					if name.find("DISS_surname").text:
						last_name = name.find("DISS_surname").text
						fullname.append(name.find("DISS_surname").text.title())
					if name.find("DISS_suffix").text:
						fullname.append(name.find("DISS_suffix").text.title())
			name = " ".join(fullname)

		if embargo == "4":
			match = False
			for author in compare_list:
				if last_name.lower().strip() == author[0].lower().strip() and first_name.lower().strip() in author[1].lower().strip():
					match = True
			if match == False:
				print (f"Can't find {last_name}, {first_name}")
			




"""
for check in tqdm(compare_list):
	match = False
	for row in sheet:
		lname = row[5].value
		fname = row[6].value
		if check[0].lower().strip() == lname.lower().strip() and check[1].lower().strip() in fname.lower().strip():
			match = True
	if match == False:
		print (f"Can't find: {check[0]}, {check[2]}")

catalogData = openpyxl.load_workbook(filename = catalogEmbargosFile)
catalog = catalogData.active
"""
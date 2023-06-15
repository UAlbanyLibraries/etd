import os
import csv
import openpyxl
from tqdm import tqdm
from packages import ETD
from datetime import datetime

catalogEmbargosFile = "\\\\Lincoln\\Library\\ETDs\\MMSIDs\\MMSID_2008_2022EMBARGOED.xlsx"
outputFile = "\\\\Lincoln\\Library\\ETDs\\ETDWorkspace\\output.csv"
etd_root = "\\\\Lincoln\\Masters\\ETD-storage"

etd_list = []
embargo_list = []
with open(outputFile, newline='', encoding="utf-8") as csvfile:
	reader = csv.reader(csvfile)
	for row in reader:
		#print (row)
		etd_list.append([row[0], row[1], row[3]])
		if len(row[4]) > 0:
			embargo_list.append(row)

wb = openpyxl.load_workbook(filename = catalogEmbargosFile)

present = datetime.now().date()

catalog_embargo_count = 0
sheet = wb.active

outputList = [["Issue", "ETD ID", "Year", "Embargo Clear Date", "Author", "Title"]]
match_list = []

# For all the catalog embargos, are they active?
for row in sheet:
	catalog_embargo_count += 1
	if catalog_embargo_count > 1:
		if " / by " in row[3].value:
			title, author = row[3].value.split(" / by ")
		else:
			title, author = row[3].value.split(" / ")
		mms_id = row[24].value
		
		match = False
		for etd in etd_list:
			if mms_id == etd[0]:
				match = True
				etd_id = etd[1]
				year = etd[2]


		if match == False:
			#raise Exception(f"Cannot find {mms_id}")
			outputList.append(["No ETD", mms_id, "", "", author, title])
		else:
			etd_path = os.path.join(etd_root, year, etd_id)

			etd = ETD()
			etd.load(etd_path)
			if etd.bag.info["Embargo-Active"] == "True":
				#print ("embargo!")
				match_list.append(mms_id)
			elif etd.bag.info["Embargo"] == "True":
				print (f"Expired embargo for {mms_id}")
				outputList.append(["Expired", mms_id, etd.etd_id, etd.bag.info["Completion-Date"], author, title])
			else:
				print (f"No embargo for {mms_id}")
				outputList.append(["No embargo", mms_id, etd.etd_id, etd.bag.info["Completion-Date"], author, title])
	

print (f"Catalog count: {catalog_embargo_count}")

with open("\\\\Lincoln\\Library\\ETDs\\ETDWorkspace\\catalog_embargo_issues.csv", "w", newline="", encoding="utf-8") as f:
	writer = csv.writer(f)
	writer.writerows(outputList)

missing_embargos = []
for embargoed_etd in embargo_list:
	if embargoed_etd[0] in match_list:
		pass
	else:
		missing_embargos.append(embargoed_etd)

with open("\\\\Lincoln\\Library\\ETDs\\ETDWorkspace\\catalog_missing_embargos.csv", "w", newline="", encoding="utf-8") as f:
	writer = csv.writer(f)
	writer.writerows(missing_embargos)
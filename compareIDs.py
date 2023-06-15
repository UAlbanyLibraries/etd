import os
import re
import csv
import openpyxl
from tqdm import tqdm
from lxml import etree
from packages import ETD
from fuzzywuzzy import fuzz


# Lets get a big list of titles and mms_ids
# \ is an escape so you have to double them up for windows paths
catalog_root = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"
pre2008 = os.path.join(catalog_root, "MMSID_1967_2007.xlsx")
post2008 = os.path.join(catalog_root, "MMSID_2008_2022.xlsx")

# empty dict to add catalog data too
records = []
print ("reading catalog export...")
row_count = 0
wb = openpyxl.load_workbook(filename=post2008, read_only=True)
for sheet in wb.worksheets:
	for row in sheet.rows:
		row_count += 1
		#skip header
		if row_count > 1:
			title_text = row[3].value
			mms_id = row[24].value

			# why is there other junk in the title field?
			# This gets all the text before " / ". Dunno if this is safe
			if " / by " in title_text:
				title, author = title_text.split(" / by ")
			elif " / " in title_text:
				title, author = title_text.split(" / ")
			else:
				#print (title_text)
				title, author = title_text.split(" by ")

			# add to records as a list
			records.append([title, author, mms_id])



# Now lets loop though all the ETD packages
etd_root = "\\\\Lincoln\\Masters\\ETD-storage"

# Start of list of 
output = [["mms_id", "etd_id", "xml_id", "year", "active embargo", "author", "title"]]

print ("reading ETD packages...")
etd_list = []
for year in os.listdir(etd_root):
	for etd in os.listdir(os.path.join(etd_root, year)):
		etd_list.append([os.path.join(etd_root, year, etd), year])

count = 0
found = 0
missing = 0
missing_list = []

# Loop though the XML
for etd_dir in tqdm(etd_list):
	# get the full path
	#xml_file = os.path.join(etd_root, etd)
	#xml_id = etd.split("_DATA")[0]

	count += 1

	#if count == 959:
	if count > 0:

		etd = ETD()
		etd.load(etd_dir[0])

		title_text = etd.bag.info["Submitted-Title"]
		fullname = etd.bag.info["Author"]
		
		match_count = 0
		# First look for easy matches
		for record in records:
			if title_text.lower() in record[0].lower():
				match_count += 1
				catalog_id = record[2]
				# end loop
				break
			else:
				# then try removing punctuation and tokenizing
				xml_tokens = re.sub(r'[^a-zA-Z0-9]\s', ' ', title_text.lower()).split()
				catalog_tokens = re.sub(r'[^a-zA-Z0-9]\s', ' ', record[0].lower()).split()

				if " ".join(xml_tokens) in " ".join(catalog_tokens):
					match_count += 1
					catalog_id = record[2]


		# for others use fuzzy matching
		if match_count == 0:
			for record in records:
				compare_title = fuzz.ratio(title_text, record[0])
				compare_authors = fuzz.ratio(fullname, record[1])
				# 80 is probably good
				if compare_title > 80 or compare_authors > 80:
					match_count += 1
					catalog_id = record[2]
				elif compare_title > 50 and compare_authors > 50:
					match_count += 1
					catalog_id = record[2]
			if match_count == 0:
				# if still no matches, try compating the stripped lowercase
				for record in records:
					stipped_xml = re.sub(r'[^a-zA-Z0-9]\s', ' ', title_text.lower())
					stipped_catalog = re.sub(r'[^a-zA-Z0-9]\s', ' ', record[0].lower())
					compare_stripped = fuzz.ratio(stipped_xml, stipped_catalog)
					if compare_stripped > 80:
						match_count += 1
						catalog_id = record[2]

		if etd.bag.info["Embargo-Active"] == "True":
			embargo_string = etd.bag.info["Embargo-Date"]
		else:
			embargo_string = ""

		# make sure there is only one match
		if match_count == 1:
			found += 1
			output.append([
				catalog_id,
				etd.etd_id,
				etd.bag.info["XML-ID"],
				etd.bag.info["Completion-Date"],
				embargo_string,
				etd.bag.info["Author"],
				etd.bag.info["Submitted-Title"]
				])
		else:
			if match_count > 1:
				print (f"multiple {str(match_count)} matches --> {etd.etd_id}")
				out_str = "multiple"
			else:
				out_str = "missing"
			missing += 1
			missing_list.append(title_text)
			output.append([
				out_str,
				etd.etd_id,
				etd.bag.info["XML-ID"],
				etd.bag.info["Completion-Date"],
				embargo_string,
				etd.bag.info["Author"],
				etd.bag.info["Submitted-Title"]
				])

print (missing_list)
print (f"Found {str(found)} of {str(count)} ETDs. {str(missing)} missing.")		

# Write contacts to CSV
output_path = os.path.join(catalog_root, "output.csv")
with open(output_path, "w", newline="", encoding="utf-8") as f:
	writer = csv.writer(f)
	writer.writerows(output)

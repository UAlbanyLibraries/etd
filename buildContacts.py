import os
import csv
from tqdm import tqdm
from lxml import etree

etd_root = "\\\\Lincoln\\Library\\ETDs\\Unzipped"

out_list = [["xml_id", "name", "email", "complete_date", "accept_date"]]

xml_list = []
for thing in os.listdir(etd_root):

	# exclude supplemental materials directories and pdfs themselves
	if thing.lower().endswith(".xml"):
		xml_list.append(thing)

count = 0
for etd in tqdm(xml_list):
	# get the full path
	xml_file = os.path.join(etd_root, etd)
	xml_id = etd.split("_DATA")[0]

	out = [xml_id]

	count += 1

	tree = etree.parse(xml_file)
	# Get the root <DISS_submission> element
	root = tree.getroot()

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
					fullname.append(name.find("DISS_fname").text.title())
				if name.find("DISS_middle").text:
					fullname.append(name.find("DISS_middle").text.title())
				if name.find("DISS_surname").text:
					fullname.append(name.find("DISS_surname").text.title())
				if name.find("DISS_suffix").text:
					fullname.append(name.find("DISS_suffix").text.title())
		out.append(" ".join(fullname))

		current_email = author.xpath("//DISS_contact[@type='current']/DISS_email")
		future_email = author.xpath("//DISS_contact[@type='future']/DISS_email")
		if len(current_email) > 1 or len(future_email) > 1:
			raise ValueError(f"{etd} has weird contact info")
		else:
			if len(future_email) == 1:
				out.append(future_email[0].text)
			elif len(current_email) == 1:
				out.append(current_email[0].text)

	comp_date = root.xpath("//DISS_description/DISS_dates/DISS_comp_date")
	accept_date = root.xpath("//DISS_description/DISS_dates/DISS_accept_date")
	#if not accept_date[0].text.startswith("01/01/"):
	#	print (accept_date[0].text)
	#if accept_date[0].text.split("01/01/")[1] != comp_date[0].text:
	#	print (accept_date[0].text + " --> " + comp_date[0].text)

	out.append(accept_date[0].text)
	out.append(comp_date[0].text)

	if out[2].lower().endswith("@albany.edu") or out[2].lower().endswith("@uamail.albany.edu") and int(comp_date[0].text) < 2022:
		out_list.append(out)



# Write contacts to CSV
catalog_root = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"
out_path = os.path.join(catalog_root, "contacts_updated.csv")
with open(out_path, "w", newline="") as f:
	writer = csv.writer(f)
	writer.writerows(out_list)
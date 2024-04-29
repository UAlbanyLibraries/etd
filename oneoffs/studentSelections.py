import shutil
import zipfile
import tempfile
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from packages import etd_loop, etd_total
from tqdm import tqdm
import os
from lxml import etree
from datetime import datetime

total = 0
publishing_data = {"total": {"oa": 0, "traditional": 0, "copyright": 0}, "pre": {"oa": 0, "traditional": 0, "copyright": 0}, "post": {"oa": 0, "traditional": 0, "copyright": 0}}
switch_date = datetime.strptime("04/09/2020", "%m/%d/%Y")

dates = {}
"""
for etd in tqdm(etd_loop(), total=etd_total()):

	xml = etd.pq_xml()

	third_party_search = xml.attrib['third_party_search']
	#third_party_sales = xml.attrib['third_party_sales']

	embargo = xml.attrib['embargo_code']

	date_string = xml.find("DISS_description").find("DISS_dates").find("DISS_accept_date").text
	date = datetime.strptime(date_string, "%m/%d/%Y")

	if date.year in dates.keys():
		dates[date.year] += 1
	else:
		dates[date.year] = 1

	
	publishing = xml.attrib['publishing_option']
	#print (publishing)
	copyright = xml.find("DISS_description").attrib['apply_for_copyright']
	#print (copyright)

	if publishing == "1": 
		publishing_data["total"]["oa"] += 1
	elif publishing == "0":
		publishing_data["total"]["traditional"] += 1	
	else:
		raise ValueError(type(publishing))
	if copyright != "no":
		publishing_data["total"]["copyright"] += 1
	if date < switch_date:
		if publishing == "1": 
			publishing_data["pre"]["oa"] += 1
		elif publishing == "0":
			publishing_data["pre"]["traditional"] += 1	
		else:
			raise ValueError(type(publishing))
		if copyright == "yes":
			publishing_data["pre"]["copyright"] += 1
		elif copyright != "no":
			raise ValueError(copyright)
	else:
		if publishing == "1": 
			publishing_data["post"]["oa"] += 1
		elif publishing == "0":
			publishing_data["post"]["traditional"] += 1	
		else:
			raise ValueError(type(publishing))
		if copyright == "yes":
			publishing_data["post"]["copyright"] += 1
		elif copyright != "no":
			raise ValueError(copyright)
	
	#print (publishing_data)
"""

ETD_package_dir = "\\\\Lincoln\\Library\\ETDs"
count = 0
packages = []

for etd_package in os.listdir(ETD_package_dir):
	if etd_package.endswith(".zip"):
		package_path = os.path.join(ETD_package_dir, etd_package)
		packages.append(package_path)

for package_path in tqdm(packages):
	
	tempDir = tempfile.mkdtemp()
	with zipfile.ZipFile(package_path, "r") as zip_ref:
		zip_ref.extractall(tempDir)
	for xml_file in os.listdir(tempDir):
		if xml_file.lower().endswith(".xml"):
			xml_path = os.path.join(tempDir, xml_file)
			#print (xml_file)
			tree = etree.parse(xml_path)
			root = tree.getroot()

			date_string = root.find("DISS_description").find("DISS_dates").find("DISS_accept_date").text
			date = datetime.strptime(date_string, "%m/%d/%Y")
			if date.year in dates.keys():
				dates[date.year] += 1
			else:
				dates[date.year] = 1

			if root.attrib["embargo_code"] != "0":
				print (root.find("DISS_restriction").find("DISS_sales_restriction").attrib)

			"""
			publishing = root.attrib['publishing_option']
			#print (publishing)
			copyright = root.find("DISS_description").attrib['apply_for_copyright']
			#print (copyright)

			if publishing == "1": 
				publishing_data["total"]["oa"] += 1
			elif publishing == "0":
				publishing_data["total"]["traditional"] += 1	
			else:
				raise ValueError(type(publishing))
			if copyright != "no":
				publishing_data["total"]["copyright"] += 1
			if date < switch_date:
				if publishing == "1": 
					publishing_data["pre"]["oa"] += 1
				elif publishing == "0":
					publishing_data["pre"]["traditional"] += 1	
				else:
					raise ValueError(type(publishing))
				if copyright == "yes":
					publishing_data["pre"]["copyright"] += 1
				elif copyright != "no":
					raise ValueError(copyright)
			else:
				if publishing == "1": 
					publishing_data["post"]["oa"] += 1
				elif publishing == "0":
					publishing_data["post"]["traditional"] += 1	
				else:
					raise ValueError(type(publishing))
				if copyright == "yes":
					publishing_data["post"]["copyright"] += 1
				elif copyright != "no":
					raise ValueError(copyright)

			"""
	shutil.rmtree(tempDir)
	
print (publishing_data)


print (dates)
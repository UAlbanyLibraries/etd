from packages import etd_loop, etd_total
from tqdm import tqdm

dept_names = []
disciplines = {}
degrees = []
lang_codes = []

for etd in tqdm(etd_loop(), total=etd_total()):

	xml = etd.pq_xml()

	dept = xml.find("DISS_description/DISS_institution/DISS_inst_contact").text
	if dept not in dept_names:
		dept_names.append(dept)

	categories = xml.find("DISS_description/DISS_categorization")
	for catagory in categories:
		if catagory.tag == "DISS_category":
			discipline = catagory.find("DISS_cat_desc").text
			code = catagory.find("DISS_cat_code").text
			if code not in disciplines.keys():
				disciplines[code] = discipline

	degree = xml.find("DISS_description/DISS_degree").text
	if degree not in degrees:
		degrees.append(degree)

	lang = categories.find("DISS_language").text
	if lang not in lang_codes:
		lang_codes.append(lang)

print ("departments:")
for dept in dept_names:
	print (f"\t{dept}")

print ("disciplines:")
for code in disciplines.keys():
	print (f"\t{code}: {disciplines[code]}")

print (degrees)
print (lang_codes)
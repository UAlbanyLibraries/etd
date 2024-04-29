sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from packages import etd_loop, etd_total
from tqdm import tqdm

dept_names = []
disciplines = {}
degrees = []
lang_codes = []
none_count = 0
psych = {}
committee = {}
advisors = {}

for etd in tqdm(etd_loop(), total=etd_total()):

	xml = etd.pq_xml()

	#dept = xml.find("DISS_description/DISS_institution/DISS_inst_contact").text
	#if dept not in dept_names:
	#	dept_names.append(dept)
	"""
	if dept is None:
		print (f"blah")
	elif "spanish" in dept.lower():
		print (f"{etd.year}: {etd.etd_id}")
	"""
	"""
	categories = xml.find("DISS_description/DISS_categorization")
	catagory_count = 0
	none_switch = False
	for catagory in categories:
		if catagory.tag == "DISS_category":
			discipline = catagory.find("DISS_cat_desc").text
			code = catagory.find("DISS_cat_code").text
			if discipline:
				pass
				
				if "physical" in discipline.lower() or "therapy" in discipline.lower():
					print (discipline)
					if discipline in disciplines:
						disciplines[discipline].append(etd.year + "_" + etd.etd_id)
					else:
						disciplines[discipline] = [etd.year + "_" + etd.etd_id]
				
			else:
				none_count +=1
				print (etd.year + " " + etd.etd_id)

	
	"""
	committee_count = 0
	advisor_count = 0
	description = xml.find("DISS_description")
	for member in description:
		if member.tag == "DISS_cmte_member":
			if member.find("DISS_name/DISS_surname").text:
				committee_count += 1
	for advisor in description:
		if advisor.tag == "DISS_advisor":
			if advisor.find("DISS_name/DISS_surname").text:
				advisor_count += 1
	if committee_count in committee.keys():
		committee[committee_count] += 1
	else:
		committee[committee_count] = 1
	if advisor_count in advisors.keys():
		advisors[advisor_count] += 1
	else:
		advisors[advisor_count] = 1

	

	#degree = xml.find("DISS_description/DISS_degree").text
	#if degree not in degrees:
	#	degrees.append(degree)
	
	#lang = categories.find("DISS_language").text
	#if lang not in lang_codes:
	#	lang_codes.append(lang)

	#print ("departments:")
#for dept in dept_names:
#	print (f"\t{dept}")
#print (none_count)
#print ("disciplines:")
#for code in disciplines.keys():
#	print (f"\t{code}: {disciplines[code]}")

#print (degrees)
#print (lang_codes)
#print (degrees)
print ("Advisors:")

print (advisors)
print ("Committees:")
print (committee)
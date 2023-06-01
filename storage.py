import os
from tqdm import tqdm
from packages import ETD

testDir = "\\\\Lincoln\\Library\\ETDs\\Zipped"

#testFile = os.path.join(testDir, "etdadmin_upload_39129.zip")

#etd = ETD()
#etd.createFromZip(testFile)

zip_list = []
for file in os.listdir(testDir):
	if file.endswith(".zip"):
		zip_list.append(os.path.join(testDir, file))

for zip_file in tqdm(zip_list):
	etd = ETD()
	etd.createFromZip(zip_file)

import os
import time
import bagit
import shutil
import zipfile
import openpyxl
import shortuuid
import urllib.parse
from lxml import etree
from datetime import datetime
from dateutil.relativedelta import relativedelta

class ETD:

    def __init__(self):
        self.excludeList = ["thumbs.db", "desktop.ini", ".ds_store"]

        if os.name == "nt":
            #self.storage_path = "\\\\Lincoln\\Library\\ETDs\\mock-storage"
            self.storage_path ="\\\\Lincoln\\Masters\\ETD-storage"
            self.working_path = "\\\\Lincoln\\Library\\ETDs"
        else:
            self.storage_path = "/media/Masters/ETD-storage"
            self.working_path = "/media/Library/ETDs"
        self.temp_space = os.path.join(self.storage_path, ".tmp")
        
        
    def load(self, path):
        #print (path)
        if not os.path.isdir(path):
            raise Exception(f"ERROR: {path} does not exist.")
        
        self.bagDir = path
        self.etd_id = os.path.basename(path)
        self.bag = bagit.Bag(path)
        # Validate completeness only
        #if not self.bag.is_valid(True, True):
        #    raise Exception(f"ERROR: Bag {path} is not complete")
        self.year = self.bag.info["Completion-Date"]
        self.year_dir = os.path.join(self.storage_path, self.year)
        self.data = os.path.join(path, "data")
        xml_id = self.bag.info["XML-ID"]
        self.xml_file = os.path.join(self.data, xml_id + "_DATA.xml")
        self.pdf_file = os.path.join(self.data, xml_id + ".pdf")
        self.supplemental = False
        if "Supplemental-Materials" in self.bag.info.keys():
            if self.bag.info["Supplemental-Materials"].lower() == "true":
                self.supplemental = True
                self.supplemental_files = self.bag.info["Supplemental-Path"]


    def check_embargo(self, xml_id, last_name, first_name):
        embargo_file = os.path.join(self.working_path, "embargos.xlsx")
        if not os.path.isfile(embargo_file):
            raise Exception(f"ERROR: Embargo export sheet {embargo_file} not present")

        match_count = 0
        wb = openpyxl.load_workbook(filename = embargo_file)
        sheet = wb.active
        embargo_date = None
        
        for row in sheet:
            title = row[2].value.strip()
            lname = row[5].value.strip()
            fname = row[6].value.strip()
            embargo = row[10].value.strip()
            degree_date = row[9].value.strip()
            
            # the last integer of the xml_id is listed in the first column of the embargo spreadsheet
            if row[0].value == xml_id.rsplit('_', 1)[1] and last_name.lower().strip() in lname.lower():
                match_count += 1
                if "-" in embargo:
                    embargo_date = datetime.strptime(embargo, "%Y-%m-%d").date()
                # Returns None if a good date isn't found
        
        if match_count != 1:
            return None
        else:
            return embargo_date

        
    
    def createFromZip(self, zip_file):
        
        metadata = {\
        'Bag-Type': 'ETD', \
        'Bagging-Date': str(datetime.now().isoformat()), \
        'Posix-Date': str(time.time()), \
        'BagIt-Profile-Identifier': 'https://archives.albany.edu/static/bagitprofiles/etd-profile-v0.1.json', \
        }

        zip_filename = os.path.basename(zip_file)
        if not os.path.isdir(self.temp_space):
            os.mkdir(self.temp_space)
        shutil.copy2(zip_file, self.temp_space)
        tmp_zip = os.path.join(self.temp_space, zip_filename)
        self.zip_id = os.path.splitext(zip_filename)[0]
        metadata["Zip-ID"] = self.zip_id
        etd_tmp = os.path.join(self.temp_space, self.zip_id)
        os.mkdir(etd_tmp)

        with zipfile.ZipFile(tmp_zip, 'r') as zip_ref:
            zip_ref.extractall(etd_tmp)

        for thing in os.listdir(etd_tmp):
            if thing.endswith(".pdf"):
                self.xml_id = os.path.splitext(thing)[0]

        if self.xml_id is None:
            raise Exception(f"ERROR: No XML ID found in zip {zip_file}")
        metadata["XML-ID"] = self.xml_id

        # parse the XML
        xml_file = os.path.join(etd_tmp, self.xml_id + "_DATA.xml")
        if not os.path.isfile(xml_file):
            raise Exception(f"ERROR: Cannot find XML file {xml_file}")
        tree = etree.parse(xml_file)
        # Get the root <DISS_submission> element
        root = tree.getroot()

        # I guess using <DISS_comp_date> as the date
        comp_date = root.xpath("//DISS_description/DISS_dates/DISS_comp_date")[0].text
        self.year = comp_date
        self.year_dir = os.path.join(self.storage_path, comp_date)
        metadata["Completion-Date"] = comp_date
        accept_date_string = root.xpath("//DISS_description/DISS_dates/DISS_accept_date")[0].text
        accept_date = datetime.strptime(accept_date_string, "%m/%d/%Y").date()

        title = root.xpath("//DISS_description/DISS_title")[0].text.title()
        metadata["Submitted-Title"] = title
        dept = root.xpath("//DISS_description/DISS_institution/DISS_inst_contact")[0].text
        metadata["Department"] = dept

        # Parse author name
        fullname = []
        last_name = ""
        author_count = 0
        for author in root.find("DISS_authorship"):
            author_count += 1
            # Thay all have one author... for now
            if author_count != 1:
                raise Exception(f"{self.xml_id} has multiple authors")
            name_count = 0
            for name in author:
                if name.tag == "DISS_name":
                    name_count += 1
                    if name_count != 1:
                        raise Exception(f"ERROR: {self.xml_id} has multiple names")

                    # safely build the name
                    if name.find("DISS_fname").text:
                        first_name = name.find("DISS_fname").text.title()
                        fullname.append(first_name)
                        metadata["First-Name"] = first_name
                    if name.find("DISS_middle").text:
                        middle_name = name.find("DISS_middle").text.title()
                        fullname.append(middle_name)
                        metadata["Middle-Name"] = middle_name
                    if name.find("DISS_surname").text:
                        last_name = name.find("DISS_surname").text.title()
                        fullname.append(last_name)
                        metadata["Last-Name"] = last_name
                    else:
                        raise Exception(f"ERROR: {self.xml_id} has no surname in XML")
                    if name.find("DISS_suffix").text:
                        suffix = name.find("DISS_suffix").text.title()
                        fullname.append(suffix)
                        metadata["Suffix"] = suffix

        author_string = " ".join(fullname)
        metadata["Author"] = author_string
        self.etd_id = last_name.replace(" ", "_") + "-" + shortuuid.uuid()
        metadata["Bag-Identifier"] = self.etd_id

        # get contact email, prefer future email
        current_email = author.xpath("//DISS_contact[@type='current']/DISS_email")
        future_email = author.xpath("//DISS_contact[@type='future']/DISS_email")
        if len(current_email) > 1 or len(future_email) > 1:
            raise ValueError(f"ERROR: {self.etd_id} has weird contact info")
        if len(future_email) > 0:
            email = future_email[0].text
        elif len(current_email) > 0:
            email = current_email[0].text
        else:
            email = ""
        metadata["ProQuest-Email"] = email

        # Parse embargos
        embargo = root.attrib['embargo_code']
        present = datetime.now().date()
        if embargo == "0":
            metadata["Embargo"] = "False"
            metadata["Embargo-Date"] = "False"
            metadata["Embargo-Active"] = "False"
        else:
            metadata["Embargo"] = "True"
            if embargo == "1":
                metadata["Embargo"] = "True"
                metadata["Embargo-Type"] = "6 month"
                embargo_end = accept_date + relativedelta(months=+6)
            elif embargo == "2":
                metadata["Embargo-Type"] = "1 year"
                embargo_end = accept_date + relativedelta(years=+1)
            elif embargo == "3":
                metadata["Embargo-Type"] = "2 year"
                embargo_end = accept_date + relativedelta(years=+2)
            elif embargo == "4":
                metadata["Embargo-Type"] = "Specified Date"
                # Go read the spreadsheet to see if embargos are there
                embargo_date = self.check_embargo(self.xml_id, last_name, first_name)
                if embargo_date is None:
                    embargo_end = "Unknown"
                else:
                    embargo_end = embargo_date
            else:
                raise Exception(f"ERROR: Bad embargo code for {self.xml_id}") 
            if isinstance(embargo_end, str):
                metadata["Embargo-Date"] = embargo_end
                metadata["Embargo-Active"] = "True"
                raise Exception(f"Could not find embargo end date for {self.xml_id}")
            else:
                metadata["Embargo-Date"] = embargo_end.strftime("%Y-%m-%d")
                if embargo_end >= present:
                    metadata["Embargo-Active"] = "True"
                else:
                    metadata["Embargo-Active"] = "False"

        # Supplemental materials
        self.supplemental = False
        attachments = root.xpath("//DISS_content/DISS_attachment")
        if len(attachments) > 0:
            self.supplemental = True
            metadata["Supplemental-Materials"] = "True"


        # Publishing option
        pub = root.attrib['publishing_option']
        if pub == "0":
            metadata["ProQuest-Publishing"] = "Traditional"
        elif pub == "1":
            metadata["ProQuest-Publishing"] = "Open Access"
        else:
            raise Exception(f"Bad publishing code for {self.xml_id}") 

        # Create the year folder if it doesn't already exist
        if not os.path.isdir(self.year_dir):
            os.mkdir(self.year_dir)

        # Create the bag
        self.bagDir = os.path.join(self.year_dir, self.etd_id)
        os.mkdir(self.bagDir)
        self.bag = bagit.make_bag(self.bagDir, metadata)
        self.data = os.path.join(self.bagDir, "data")

        # Copy files into /data
        for thing in os.listdir(etd_tmp):
            thingPath = os.path.join(etd_tmp, thing)
            if os.path.isfile(thingPath):
                if not thing.lower() in self.excludeList:
                    shutil.copy2(thingPath, self.data)
            else:
                shutil.copytree(thingPath, os.path.join(self.data, thing))

        self.xml_file = os.path.join(self.data, self.xml_id + "_DATA.xml")
        self.pdf_file = os.path.join(self.data, self.xml_id + ".pdf")
        if not os.path.isfile(self.xml_file):
            raise Exception(f"ERROR: Package {self.xml_id} missing XML file {self.xml_file}")
        if not os.path.isfile(self.pdf_file):
            raise Exception(f"ERROR: Package {self.xml_id} missing PDF file {self.pdf_file}")
        #check if theres a supplemental materials folder with files
        there = False
        sup_folders = 0
        for thing in os.listdir(self.data):
            if os.path.isdir(os.path.join(self.data, thing)):
                sup_folders += 1
                for subthing in os.listdir(os.path.join(self.data, thing)):
                    if os.path.isfile(os.path.join(self.data, thing, subthing)):
                        there = True
                        self.supplemental_files = os.path.join(self.data, thing, subthing)
                        self.bag.info["Supplemental-Path"] = self.supplemental_files
        if self.supplemental:
            if not there or sup_folders != 1:
                raise Exception(f"ERROR: Package {self.xml_id} missing supplemental materials")
        elif there:
            raise Exception(f"ERROR: Package {self.xml_id} contains supplemental materials not listed in XML")

        # Save bag
        self.bag.save(manifests=True)

        # Delete temp contents after move
        os.remove(tmp_zip)
        for thing in os.listdir(etd_tmp):
            thingPath = os.path.join(etd_tmp, thing)
            if os.path.isfile(thingPath):
                os.remove(thingPath)
            else:
                shutil.rmtree(thingPath)
        if len(os.listdir(etd_tmp)) == 0:
            os.rmdir(etd_tmp)
        if len(os.listdir(self.temp_space)) == 0:
            os.rmdir(self.temp_space)
        
    def clean(self):
        for root, dirs, files in os.walk(self.data):
            for file in files:
                if file.lower() in self.excludeList:
                    filePath = os.path.join(root, file)
                    print ("removing " + filePath)
                    os.remove(filePath)

    def size(self):
        suffixes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
        bytes, fileCount = self.bag.info["Payload-Oxum"].split(".")
        dirSize = int(bytes)
        i = 0
        while dirSize >= 1024 and i < len(suffixes)-1:
            dirSize /= 1024.
            i += 1
        f = ('%.2f' % dirSize).rstrip('0').rstrip('.')
        return [f, suffixes[i], fileCount]

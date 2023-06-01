import os
import time
import bagit
import shutil
import zipfile
from datetime import datetime

class ETD:

	def __init__(self):
        self.excludeList = ["thumbs.db", "desktop.ini", ".ds_store"]
        self.storage_path = "/media/Masters/ETD-storage"
        self.working_path = "/media/Library/ETDs"
        self.temp_space = "/media/Masters/ETD-storage/.tmp"
        
    
    def load(self, path):
        if not os.path.isdir(path):
            raise Exception("ERROR: " + str(path) + " is not a valid SIP. You may want to create a SIP with .create().")

        self.bag = bagit.Bag(path)
        self.bagID = os.path.basename(path)
        self.colID = self.bagID.split("_")[0]
        self.data = os.path.join(path, "data")
        
    
    def createFromZip(self, zip_file):
        
        metadata = {\
        'Bag-Type': 'ETD', \
        'Bagging-Date': str(datetime.now().isoformat()), \
        'Posix-Date': str(time.time()), \
        'BagIt-Profile-Identifier': 'https://archives.albany.edu/static/bagitprofiles/sip-profile-v0.2.json', \
        'Collection-Identifier': colID \
        }

        zip_filename = os.path.basename(zip_file)
        if not os.path.isdir(self.temp_space):
            os.mkdir(self.temp_space)
        shutil.copy2(zip_file, self.temp_space)
        tmp_zip = os.path.join(self.temp_space, zip_filename)

        with zipfile.ZipFile(tmp_zip, 'r') as zip_ref:
            zip_ref.extractall(self.temp_space)
        
        """
        self.colID = colID
        self.bagID = colID + "_" + str(shortuuid.uuid())
        metadata["Bag-Identifier"] = self.bagID
        if not os.path.isdir(os.path.join(self.sipPath, colID)):
            os.mkdir(os.path.join(self.sipPath, colID))
            
        self.bagDir = os.path.join(self.sipPath, colID, self.bagID)
        os.mkdir(self.bagDir)

        self.bag = bagit.make_bag(self.bagDir, metadata)
        self.data = os.path.join(self.bagDir, "data")
        """
        
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
import os
import csv
import time
import bagit
import shutil
from datetime import datetime

class ArchivalInformationPackage:

    def __init__(self, colID, accession):
        aipPath= "/media/Masters/Archives/AIP"
        
        metadata = {\
        'Bag-Type': 'AIP', \
        'Bagging-Date': str(datetime.now().isoformat()), \
        'Posix-Date': str(time.time()), \
        'BagIt-Profile-Identifier': 'https://archives.albany.edu/static/bagitprofiles/aip-profile-v0.1.json', \
        }
        
        self.accession = accession
        if not os.path.isdir(os.path.join(aipPath, colID)):
            os.mkdir(os.path.join(aipPath, colID))
        self.bagDir = os.path.join(aipPath, colID, accession)
        os.mkdir(self.bagDir)
        metadata["Bag-Identifier"] = accession
        metadata["Collection-Identifier"] = colID
        
        self.bag = bagit.make_bag(self.bagDir, metadata)
        self.data = os.path.join(self.bagDir, "data")
        
    def addMetadata(self, hyraxData):
        headers = ["Type", "URIs", "File Paths", "Accession", "Collecting Area", "Collection Number", "Collection", \
        "ArchivesSpace ID", "Record Parents", "Title", "Description", "Date Created", "Resource Type", "License", \
        "Rights Statement", "Subjects", "Whole/Part", "Processing Activity", "Extent", "Language"]
        metadataPath = os.path.join(self.bagDir, "metadata")
        if not os.path.isdir(metadataPath):
            os.mkdir(metadataPath)
        metadataFile = os.path.join(metadataPath, accession + ".tsv")
        addHeaders = False
        if not os.path.isfile(metadataFile):
            addHeaders = True
        outfile = open(metadataFile, "a")
        writer = csv.writer(outfile, delimiter='\t', lineterminator='\n')
        if addHeaders == True:
            writer.writerow(headers)
        writer.writerow(hyraxData)
        outfile.close()
        
    def addFile(file):
        dataPath = os.path.join(self.bagDir, "data")
        if not os.path.isdir(metadataPath):
            os.mkdir(dataPath)
        shutil.copy2(file, dataPath)
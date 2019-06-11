import os
import csv
import time
import bagit
import shutil
from datetime import datetime

class ArchivalInformationPackage:

    def __init__(self):
        pass
        
    def load(self, path):
        if not os.path.isdir(path):
            raise Exception("ERROR: " + str(path) + " is not a valid AIP. You may want to create a AIP with .create().")

        self.bag = bagit.Bag(path)
        self.accession = os.path.basename(path)
        self.colID = self.accession.split("_")[0]
        self.data = os.path.join(path, "data")        
    
    def create(self, colID, accession):
        aipPath= "/media/Masters/Archives/AIP"
        
        metadata = {\
        'Bag-Type': 'AIP', \
        'Bagging-Date': str(datetime.now().isoformat()), \
        'Posix-Date': str(time.time()), \
        'BagIt-Profile-Identifier': 'https://archives.albany.edu/static/bagitprofiles/aip-profile-v0.1.json', \
        }
        
        self.accession = accession
        self.colID = accession.split("_")[0]
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
        
    def extentLog(self, logFile):
        import openpyxl
        if os.path.isfile(logFile):
            wb = openpyxl.load_workbook(filename=logFile, read_only=False)
            sheet = wb.active
            startRow = int(sheet.max_row) + 1
        else:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet["A1"] = "Date"
            sheet["B1"] = "Collection ID"
            sheet["C1"] = "Type"
            sheet["D1"] = "Package"
            sheet["E1"] = "Files"
            sheet["F1"] = "Extent"
            sheet["G1"] = "Extent Bytes"
            startRow = 2
            
        packageSize = self.size()
        sheet["A" + str(startRow)] = self.bag.info["Bagging-Date"]
        sheet["B" + str(startRow)] = self.bag.info["Collection-Identifier"]
        sheet["C" + str(startRow)] = self.bag.info["Bag-Type"]
        sheet["D" + str(startRow)] = self.bag.info["Bag-Identifier"]
        sheet["E" + str(startRow)] = packageSize[2]
        sheet["F" + str(startRow)] = str(packageSize[0]) + " " + str(packageSize[1])
        sheet["G" + str(startRow)] = self.bag.info["Payload-Oxum"]
            
        wb.save(filename=logFile)
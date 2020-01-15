import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD, GetFileAttributes
import csv
from difflib import SequenceMatcher

def GetScanDirs(fileshare_path):
    # write_to_csv("","ALL_Release_Folders",True)
    lastStr = ""
    currStr  = ""
    for root, dirs, files in os.walk(str(fileshare_path)):
        File_count = len(files)
        top_folder_pos = root.find(".Release")
        if  top_folder_pos> -1:
            data = root[top_folder_pos+9:]#+','+fileNm
            if data.find("\\") == -1 and data != "":
                currStr = root
                # print('Root :',root,",",data,",",similar(currStr,lastStr))
                #prnData = root+","+data+","+str(similar(currStr,lastStr))
                dist = similar(currStr,lastStr)
                if dist  < 0.80:
                    dataArr=[]
                    root = root.replace("\\"+data,"/"+data)
                    dataArr.append(root)
                    print(root,dist)
                    write_to_csv(dataArr,"scan",False)
                lastStr = currStr  
           

def similar(a,b):
    return SequenceMatcher(None,a,b).ratio()

def write_to_csv(rowdata, file_name, Is_Header):
        with open(file_name+'.csv', mode='a',newline='') as dll_file:
                fieldnames = ["Folder_name"]
                writer = csv.writer(dll_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                #writer.writeheader()
                if(not Is_Header):
                    writer.writerow(rowdata)
                if(Is_Header):
                    writer.writerow(fieldnames)

# fileshare_path = '\\\\iaai.com/tfs-builds/Buyer'

# GetScanDirs(fileshare_path)
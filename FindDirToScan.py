import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD, GetFileAttributes
import csv
from difflib import SequenceMatcher

def GetScanDirs(fileshare_path):
    # write_to_csv("","ALL_Release_Folders",True)
    lastStr = ""
    currStr  = ""
    latest_modified_time = 0.0
    latest_file =""
    data = ""
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
                # file_path = os.path.join(root, fileNm)
                #file_version = ".".join ([str (i) for i in get_version_number (file_path)])
                attrib = os.path.getsize(root)
                modified_time = os.path.getmtime(root)
                print("root",root,"latest_file =",latest_file , ", mod time =",modified_time)
                if modified_time>latest_modified_time:
                    latest_modified_time = modified_time
                    latest_file = currStr

                
                
                if dist  < 0.80:
                    latest_modified_time = modified_time
                    dataArr=[]
                    root = root.replace("\\"+data,"/"+data)
                    dataArr.append(root)
                    print("dist gt 0.80",root,latest_file)
                    write_to_csv(dataArr,"scan",False)
                lastStr = currStr  
    dataArr=[]
    print(data)
    latest_file = latest_file.replace("\\","/")
    dataArr.append(latest_file)
    write_to_csv(dataArr,"scan",False)       

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
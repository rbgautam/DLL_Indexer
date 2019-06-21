import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD, GetFileAttributes
import csv
from datetime import datetime

search_path_list= []
output_csv = 'DLL_Catalog'
def get_settings():
        output_csv_new = init_csv(output_csv)
        write_to_csv(output_csv_new,"","","","",True)
        with open('settings.csv',mode='r') as csv_file:
                csv_reader = csv.DictReader(csv_file)
                line_count = 0
                for row in csv_reader:
                        inspath = row["Install_path"]
                        search_path_list.append(inspath)
                        line_count = line_count +1
        for srcpath in search_path_list:
                print(srcpath)
                findDLLS(str(srcpath),output_csv_new)
       
        

def findDLLS(install_path,output_csv_new):
        try:
                #print(install_path)
                #print(output_csv_new)
                walkDirs(install_path,output_csv_new)
        except expression as identifier:
                print('Error')

def parseDirs(folderpath):
        folders = [f.path for f in os.scandir(folderpath) if f.is_dir()]
        for folder in folders:
                # get txt files from folder path
                files = [f.path for f in os.scandir(folder) if f.name.endswith(".dll")]
                for f_name in files:
                        print(f_name)

def walkDirs(install_path,output_csv_new):
        for root, dirs, files in os.walk(str(install_path)): #Get path from CSV file or JSON
                        for file in files:
                                if file.endswith(".dll") and not ignoreFile(file): #make the extension configurable (more than one allowed)
                                        file_path = os.path.join(root, file)
                                        file_version = ".".join ([str (i) for i in get_version_number (file_path)])
                                        print (file_version)
                                        print(file_path)
                                        write_to_csv(output_csv_new,file,file_version,file_path,install_path,False)
def init_csv(output_csv):
        now  = datetime.now()
        date_time = now.strftime("_%m%d%Y%H%M%S")
        output_version = output_csv + date_time+".csv"
        return output_version      

def ignoreFile(file):
        ignore_list =['IAAI','SYSTEM','MICROSOFT'] #get ignorelist from config or csv
        #ignore_flag =False
        for l in ignore_list:
                if file.upper().find(l.upper()) > -1:
                        return True
        return False


def get_version_number (filename):
   info = GetFileVersionInfo (filename, "\\")
   ms = info['FileVersionMS']
   ls = info['FileVersionLS']
   return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)

def write_to_csv(csv_fileName,dll_name, dll_version, dll_path,search_path, Is_Header ):
        with open(csv_fileName, mode='a',newline='') as dll_file:
                #fieldnames = ['DLL Name', 'Version', 'Path']
                #writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer = csv.writer(dll_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                #writer.writeheader()
                if(not Is_Header):
                        writer.writerow([dll_name, dll_version, dll_path,search_path])
                if(Is_Header):
                        writer.writerow(["DLL Name", "DLL Version", "DLL Path","Server Path"])

get_settings()
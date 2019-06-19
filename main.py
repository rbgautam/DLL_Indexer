import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import csv
from datetime import datetime

#dir_path = "c:\TFSDev\PAM\MainTrunk"
#nw_path = "\\\\iaai.com/EnterpriseServices/EVM/Staging/IHSData1"
install_path = "\\\\qevm-web02/EVM"
output_csv = 'dll_file'
def findDLLS():
        output_csv_new = init_csv(output_csv)
        write_to_csv(output_csv_new,"","","",True)
        for root, dirs, files in os.walk(install_path): #Get path from CSV file or JSON
                for file in files:
                        if file.endswith(".dll") and not ignoreFile(file): #make the extension configurable (more than one allowed)
                                file_path = os.path.join(root, file)
                                file_version = ".".join ([str (i) for i in get_version_number (file_path)])
                                print (file_version)
                                print(file_path)
                                write_to_csv(output_csv_new,file,file_version,file_path,False)
def init_csv(output_csv):
        now  = datetime.now()
        date_time = now.strftime("_%m%d%Y%H_%M_%S")
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
#'dll_file.csv'
def write_to_csv(csv_fileName,dll_name, dll_version, dll_path, Is_Header ):
        with open(csv_fileName, mode='a',newline='') as dll_file:
                #fieldnames = ['DLL Name', 'Version', 'Path']
                #writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer = csv.writer(dll_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                #writer.writeheader()
                if(not Is_Header):
                        writer.writerow([dll_name, dll_version, dll_path])
                if(Is_Header):
                        writer.writerow(["Dll Name", "Dll Version", "Dll Path"])

findDLLS()
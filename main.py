import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import csv

dir_path = "c:\TFSDev\PAM\MainTrunk"
nw_path = "\\\\iaai.com/EnterpriseServices/EVM/Staging/IHSData1"
install_path = "\\\\qevm-web02/EVM"
def findDLLS():
        for root, dirs, files in os.walk(install_path): #Get path from CSV file or JSON
                for file in files:
                        if file.endswith(".dll") and not ignoreFile(file): #make the extension configurable (more than one allowed)
                                file_path = os.path.join(root, file)
                                file_version = ".".join ([str (i) for i in get_version_number (file_path)])
                                print (file_version)
                                print(file_path)
                                write_to_csv(file,file_version,file_path)
        
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

def write_to_csv(dll_name, dll_version, dll_path ):
        with open('dll_file.csv', mode='a',newline='') as dll_file:
                #fieldnames = ['DLL Name', 'Version', 'Path']
                #writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer = csv.writer(dll_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                #writer.writeheader()
                writer.writerow([dll_name, dll_version, dll_path])

findDLLS()
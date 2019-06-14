import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD

dir_path = "c:\TFSDev\PAM\MainTrunk"
nw_path = "\\\\iaai.com/EnterpriseServices/EVM/Staging/IHSData1"
install_path = "\\\\qevm-web02/EVM"
def findDLLS():
        for root, dirs, files in os.walk(install_path): #Get path from CSV file or JSON
                for file in files:
                        if file.endswith(".dll") and not ignoreFile(file): #make the extension configurable (more than one allowed)
                                file_path = os.path.join(root, file)
                                print (".".join ([str (i) for i in get_version_number (file_path)]))
                                print(file_path)
        
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

findDLLS()
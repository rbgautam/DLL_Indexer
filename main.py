import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD, GetFileAttributes
import csv
from datetime import datetime
import time, sys
from tqdm import tqdm
import pyodbc 
from datetime import date
from FindDirToScan import GetScanDirs
#\\iaai.com\tfs-builds\FieldOps
#\\iaai.com\tfs-builds 

search_path_list= []
file_list =[]
output_csv = 'DLL_Catalog'
bar_pos_count = 0
conn = None
def get_settings():
        global bar_pos_count
        connect_to_Sql_Server()
        output_csv_new = init_csv(output_csv)
        
        # write_to_csv(output_csv_new,"","","","",True)
        with open('app.config',mode='r') as csv_file:
                
                csv_reader = csv.DictReader(csv_file)
                line_count = 0
                for row in csv_reader:
                        inspath = row["Install_path"]
                        search_path_list.append(inspath)
                        line_count = line_count +1
        for srcpath in search_path_list:
                print("srcpath",srcpath)
                GetScanDirs(srcpath)
                #Once folderlist created call finddlls for each
                ReadScanDirs(bar_pos_count) 
                
        print('\nTotal files : ' +str(len(file_list)))
        print('\nWriting to DB..')
        catalog_files(file_list,output_csv_new)

       
def ReadScanDirs(bar_pos_count):
        print('read scan')
        with open('scan.csv',mode='r') as csv_file:
                readCSV = csv.reader(csv_file, delimiter=',')
                for row in readCSV:
                        
                        print("row",str(row)[1:-1])
                        findDLLS(str(row)[1:-1],bar_pos_count)

def findDLLS(install_path,bar_pos_count):
        print('fnd dll')
        try:
                print(install_path[3:-1])
                #print(output_csv_new)
                walkDirs(install_path[3:-1],bar_pos_count)
                
        except Exception as error:
                print('Error',error)

def parseDirs(folderpath):
        folders = [f.path for f in os.scandir(folderpath) if f.is_dir()]
        for folder in folders:
                # get dll files from folder path
                files = [f.path for f in os.scandir(folder) if f.name.endswith(".dll")]
                for f_name in files:
                        print(f_name)


def walkDirs(install_path,bar_pos_count):
        
        
        print('walkdir',install_path)
        # install_path ="\\\\iaai.com/tfs-builds/Buyer/AC.APC.Release/APC__2019.11.21_20191112.1"
        # install_path =  '\\\\iaai.com/tfs-builds/Buyer/AC.APC.Release/APC__2019.11.21_20191112.1'
        try:
                for root, dirs, files in os.walk(str(install_path)):
                        #file_count = len(files)
                        print('Root: ', root)
                        for file in files:
                                if file.endswith(".dll") and not ignoreFile(file): #TODO make the extension configurable (more than one allowed)
                                        file_path = os.path.join(root, file)
                                        file_version = ".".join ([str (i) for i in get_version_number (file_path)])
                                        #print (file_version)
                                        #print(file_path)
                                        #write_to_csv(output_csv_new,file,file_version,file_path,install_path,False)
                                        file_tuple = (file,file_version,file_path,install_path)
                                        file_list.append(file_tuple)
                                        
                                update_progress_rotate(bar_pos_count)
                                #time.sleep(0.1)
                                bar_pos_count = bar_pos_count +1
                                if bar_pos_count == 4:
                                        bar_pos_count = 0 
                 
        except Exception as error:
                print(error) 
                        
                                
                        
                        

def catalog_files(file_list,output_csv_new):
        now  = datetime.now()
        counter = 0
        for i in tqdm(range(len(file_list))):
                file_tup = file_list[i]
                #print(file_tup[0])
                #write_to_csv(output_csv_new,file_tup[0],file_tup[1],file_tup[2],file_tup[3],False)
                #file,file_version,file_path,install_path
                insert_into_sql(file_tup[0],file_tup[1],file_tup[2],'',file_tup[3],now)
                time.sleep(0.05)
                counter = counter + 1
        #print(counter)

                        
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
        try:
                info = GetFileVersionInfo (filename, "\\")
                ms = info['FileVersionMS']
                ls = info['FileVersionLS']
                return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
        except Exception as error:
                print(error)
                pass     
   

def write_to_csv(csv_fileName,dll_name, dll_version, dll_path,search_path, Is_Header ):
        with open(csv_fileName, mode='a',newline='') as dll_file:
                fieldnames = ["DLL Name", "DLL Version", "DLL Path","Server Path"]
                writer = csv.writer(dll_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                if(not Is_Header):
                        writer.writerow([dll_name, dll_version, dll_path,search_path])
                if(Is_Header):
                        writer.writerow(fieldnames)

def update_progress(progress):
    barLength = 100 # Modify this to change the length of the progress bar
    status = ""
    if isinstance(progress, int):
        progress = float(progress)
    if not isinstance(progress, float):
        progress = 0
        status = "error: progress var must be float\r\n"
    if progress < 0:
        progress = 0
        status = "Halt...\r\n"
    if progress >= 1:
        progress = 1
        status = "Done...\r\n"
    block = int(round(barLength*progress))
    text = "\rPercent: [{0}] {1}% {2}".format( "#"*block + "-"*(barLength-block), progress*100, status)
    sys.stdout.write(text)
    sys.stdout.flush()

def update_progress_rotate(bar_count):
        bar_pos = ["|","/","--","\\"]
        prog_text =["Indexing in progress .","Indexing in progress ..","Indexing in progress ..."," "]
        text = str(prog_text[bar_count])
        sys.stdout.write("\r "+text+" ")
        sys.stdout.flush()
        
        

def show_progress():
        for i in tqdm(range(10)):
                time.sleep(1)

def connect_to_Sql_Server():
        global conn 
        conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=QCON16DB01;'
                        'Database=ThirdPartydlldetails;'
                        'Trusted_Connection=yes;')
       
              
       
def insert_into_sql( dllName, dllVersion, dllPath, applicationName, applicationServerPath,dateStr):
        cursor = conn.cursor()
        cursor.execute("INSERT INTO dbo.DLLInfo( DLLName, DLLVersion, DLLPath,ApplicationName, ApplicationServerPath, CreateDateTime, UpdateUserDetailId, UpdateDateTime ) VALUES (?,?,?,?,?,?,?,?)",dllName,dllVersion,dllPath,applicationName,applicationServerPath,dateStr,'0',dateStr)
        cursor.commit()

def  show_data_from_sql():
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM dbo.DLLInfo with (nolock)')
        for row in cursor:
                print(row)

get_settings()
#connect_to_Sql_Server()
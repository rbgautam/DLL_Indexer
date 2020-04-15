# DLL_Indexer
Parse through directory and inventories files based on extension
Needs pywin32 to work
pip install pywin32
Progressbar used tqdm
pip install tqdm

pip install pyInstaller
pyinstaller --onefile main.py

Database connection only works upto Pythin 3.7.2
pip install pyodbc 

USE ThirdPartydlldetails
Select * from   dbo.DLLInfo d  with (NOLOCK) ORDER BY CreateDateTime desc

USE CASES
----------
Dll version comapre during release
Check for oldest verion of the dlls
Can be use to scan exe file names
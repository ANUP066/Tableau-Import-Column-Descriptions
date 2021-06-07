from cx_Freeze import setup, Executable 

files = ['mchpLogo.jpg']
  
setup(name = "ColDef" , 
      version = "0.1" , 
      description = "" , 
      options = {'build_exe': {'include_files':files}} ,
      executables = [Executable("importDescriptions.py")]) 
import warnings
from io import StringIO
import pandas as pd
import numpy as np
import warnings
from openpyxl import load_workbook
warnings.filterwarnings('ignore')
import re 
from openpyxl import load_workbook
from django.core.files.storage import FileSystemStorage
from zipfile import ZipFile 
import zipfile
import os

def read_file(file1, file2, file3, file4):

    print("Loading the Zip File")

    input_dir = r'E:\\user_jeff\Jupyter_Notebooks\Project Gillette\file_renamer\static\wsws_file\input'
    filelist = [ f for f in os.listdir(input_dir) if f.endswith(".docx") ]
    print("Zip file Loaded")
    print("Celaring the input folder")
    for f in filelist:
        os.remove(os.path.join(input_dir, f))

    
    #path =  "C:/Users/ansaj/Desktop/DeepSchick/FileRenaming/change_me/data.zip"
    print("Input  Folder Cleared")
    with ZipFile(file1, 'r') as zipObj:
        zipObj.extractall(input_dir)

    
    print("Zip Files Extracted")
#Lets Rename the Files

    filenames = os.listdir(input_dir)
    for file in filenames: 
        filename, file_extension = os.path.splitext(file)
        newname = filename[0:7]
        os.rename(os.path.join(input_dir,file), os.path.join(input_dir, newname +file_extension))

    print("Files Renamed")
# In[46]:

    
    output_zip = zipfile.ZipFile(r'E:\\user_jeff\Jupyter_Notebooks\Project Gillette\file_renamer\static\wsws_file\renamed_files.zip', 'w')
    for folder, subfolders, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.docx'):
                output_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder,file), input_dir), compress_type = zipfile.ZIP_DEFLATED)
    output_zip.close()
    print("All Zipped Up and Ready for Download")
    return []
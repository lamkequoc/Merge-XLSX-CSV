from tkinter.filedialog import askdirectory
from tkinter import messagebox
import os
import glob
import pandas as pd
files_directory = askdirectory(title='Select Folder')
merged_file_name = "merged" #merged file name
files_extension = [".xlsx",".csv"] #files_extension
if files_directory != '': #if merged file directory is selected then:
    current_file = len(glob.glob(os.path.join(files_directory,"*.*")))    
    if current_file <= 1: #Check if there has more than 1 file in directory        
        messagebox.showerror('Error', 'There is NOT enough file to merge!')        
    else:
        hdmsg=messagebox.askquestion('File Header', 'Do the files have a HEADER?', icon='warning') #Check file has header or not
        for ext in files_extension:
            files = os.path.join(files_directory,"*"+ext) #select file extension statement       
            merged_files = os.path.join(files_directory,merged_file_name+ext) #generate full meged file name
                          
            def mergeExcelCsv(): #generate mergeExcelCsv Function
                all_files = glob.glob(files) #Including all files with directory to all_files list   
                #If files has header:            
                if files_extension[0].lower() in str(all_files).lower() and hdmsg == 'yes':
                    df = pd.concat([pd.read_excel(a) for a in all_files])                                                  
                    df.to_excel((merged_files), index=False) #save .xlsx to same location            
                elif files_extension[1].lower() in str(all_files).lower() and hdmsg == 'yes':
                    df = pd.concat([pd.read_csv(a, sep=',', encoding='latin1') for a in all_files])                    
                    df.to_csv((merged_files), index=False) #save .csv to same location     
                #If files without header:
                elif files_extension[0].lower() in str(all_files).lower() and hdmsg == 'no':
                    df = pd.concat([pd.read_excel(a, header=None) for a in all_files],ignore_index=True)                                                  
                    df.to_excel((merged_files), header=False, index=False) #save .xlsx to same location            
                elif files_extension[1].lower() in str(all_files).lower() and hdmsg == 'no':
                    df = pd.concat([pd.read_csv(a, header=None, sep=',', encoding='latin1') for a in all_files], ignore_index=True)                    
                    df.to_csv((merged_files), header=False, index=False) #save .csv to same location            
                else: 
                    pass            
            if os.path.exists(merged_files): #if merged file exist                           
                filename_only = os.path.split(merged_files)[1] #select file name only
                message = filename_only + " already exists, do you want to overwrite it?"
                msgbox=messagebox.askquestion("Files exist", message, icon='warning')
                if msgbox == 'yes':
                    os.remove(merged_files)    
                    messagebox.showinfo('File overwrite', 'This file is overwritten!')                                                    
                    mergeExcelCsv()
                if msgbox == 'no':
                    pass                   
            else:
                mergeExcelCsv()              
else:
    pass

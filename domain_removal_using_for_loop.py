import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile 
from tkinter import filedialog
import os

# intializing the window
window = Tk()
window.title("domain_removal_using_for_loop")
# configuring size of the window 
window.geometry('800x500')

class MyExcelWindow:
    
    def __init__(self,window):
        # widgets 
        self.button_openfile = Button(window, text ='Open Template', command = lambda:self.open_file()) 
        self.button_openfile.place(x=50,y=100)
        self.fileentry= Entry(window,width=100)
        self.fileentry.place(x=150,y=100)
        self.button_remove_domain=Button(window,text="Select Folder",command=self.select_folder_path)
        self.button_remove_domain.place(x=50,y=200)
        self.folderentry= Entry(window,width=100)
        self.folderentry.place(x=150,y=200)        
        self.button_execute=Button(window,text="Execute",command=self.remove_domain)
        self.button_execute.place(x=100,y=300)
        self.label_mailheader=Label(window, text = "Email Header")
        self.label_mailheader.place(x=50,y=250)  
        self.emailentry= Entry(window,width=20)
        self.emailentry.place(x=150,y=250)
        self.emailentry.insert(END, 'Email ')

    # Method to get domain names from excel file
    def open_file(self): 
        file = askopenfile(mode ='r', filetypes =[('Excel Files', '.xls  .xlsx')]) 
        filename=file.name
        filename =filename.replace('/','\\')
        print(filename)
        self.fileentry.delete(0,END)
        self.fileentry.insert(0,filename)   
        

    def select_folder_path(self):
        filename = filedialog.askdirectory(title ='open')
        self.folderpath = filename.replace('/','\\')
        self.folderentry.delete(0,END)
        self.folderentry.insert(0,self.folderpath)

    def remove_domain(self):

        #  read template file
        self.opentemplate_filename=self.fileentry.get()
        if self.opentemplate_filename is not None: 
            self.df_template = pd.read_excel(self.opentemplate_filename)

        #  read  files in folder
        self.folder_name=self.folderentry.get()
        for entry in os.listdir(self.folder_name):
            if os.path.isfile(os.path.join(self.folder_name, entry)):
                file=os.path.join(self.folder_name,entry)
                appended_data = []
                email_header=self.emailentry.get()
                if file.endswith(".xls") or file.endswith(".xlsx") :
                    print(file)
                    df1=pd.read_excel(file)
                    print(df1)
                    for domain in self.df_template['Email']:
                        df_filtered = df1[df1[email_header].str.split("@").str[1]==domain]
                        df1.drop(df1[df1[email_header].str.split("@").str[1]==domain].index, inplace = True)
                        print(df_filtered)
                        if not df_filtered.empty == True:
                            appended_data.append(df_filtered)

                    try:
                        df_removed=pd.concat(appended_data)
                        csvfile=entry.split(".")[0]
                        csvfile=self.folder_name+"\\"+csvfile+" Removed-list"+".csv"
                        print(csvfile)
                        df_removed.to_csv(csvfile, index=False)
                    except ValueError:
                        print("no values to append")
                        pass
                        
                    df1.to_excel(file, index=False)

mywin = MyExcelWindow(window)  
window.mainloop()

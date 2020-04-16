import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile 
from tkinter import filedialog
import os

# intializing the window
window = Tk()
window.title("Remove Domain using Left join pandas")
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
        
    # Method to select folder to remove records with domain names given in above selected excel file 
    def select_folder_path(self):
        filename = filedialog.askdirectory(title ='open')
        self.folderpath = filename.replace('/','\\')
        self.folderentry.delete(0,END)
        self.folderentry.insert(0,self.folderpath)

    # Method to remove domain names from excel files using left join 
    # Results are stored in  excel file and csv file
    # Excel file which contains data where records related to specific domains deleted
    # CSV file with all records which are deleted from previous file 
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
                    df1['domain']=df1[email_header].str.split('@').str[1]
                    df_left = df1.merge(self.df_template.rename({'Email': 'domain_name'},axis=1), left_on='domain',right_on='domain_name', how='left')
                    df_left.to_excel("left.xlsx", index=False)
                    df_valid=df_left[df_left['domain_name'].isnull()]
                    df_valid.drop(['domain','domain_name'],axis=1,inplace=True)  
                    df_invalid=df_left[df_left['domain_name'].notnull()]
                    df_invalid.drop(['domain','domain_name'],axis=1,inplace=True)
                    csvfile=entry.split(".")[0]
                    csvfile=self.folder_name+"\\"+csvfile+" Removed-list"+".csv"
                    print(csvfile)
                    df_invalid.to_csv(csvfile, index=False)
                    df_valid.to_excel(file, index=False)

mywin = MyExcelWindow(window)  
window.mainloop()

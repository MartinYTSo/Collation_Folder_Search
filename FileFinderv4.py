# -*- coding: utf-8 -*-
"""
Created on Thu Apr 28 16:05:14 2022

@author: Martin
"""

from tkinter import *
import docx2txt
import os 
from tkinter import filedialog
import tkinter.messagebox


class FileAppV3:
    def __init__(self):
        self.root=Tk()
        self.root.geometry('1000x150')
        self.root.resizable(False,False)
        self.root.title('Collation Folder Search')
        
        self.mainFrame = Frame(self.root, container="false")
        self.filename = Label(self.mainFrame)
        self.filename.configure(compound="top", text="File Name")

        self.filePthSearchPath=StringVar()
        
        self.filepath = Label(self.mainFrame)
        self.filepath.configure(textvariable=self.filePthSearchPath)

        
        self.keyword = Label(self.mainFrame)
        self.keyword.configure(text="Keyword")

        
        self.keywordEntry = Entry(self.mainFrame)

        
        self.resultsLabel = Label(self.mainFrame)
        self.resultsLabel.configure(text="Results")

        
        self.browseButton = Button(self.mainFrame,command=self.searchFile)
        self.browseButton.configure(text="Browse")

        
        self.searchButton = Button(self.mainFrame)
        self.searchButton.configure(font="TkDefaultFont", text="Search",command=self.find_doc)

        
        self.clearButton = Button(self.mainFrame)
        self.clearButton.configure(text="Clear",command=self.clear)
        
        self.resultsFrame=Frame(self.mainFrame)
        self.resultsFrame.configure(height=400, takefocus=False, width=450)
        self.outerframe=Frame(self.resultsFrame,height=400,width=450)
        self.outerframe.pack(padx=20,pady=20)
        
        self.innerFrame=Frame(self.outerframe)
        self.innerFrame.pack()
        
        self.listbox=Listbox(self.innerFrame,height=5,width=90)
        self.listbox.pack(side='left')
        
        self.vScrollbar=Scrollbar(self.innerFrame,orient=VERTICAL)
        self.vScrollbar.pack(side='right',fill=Y)
        
        self.hScrollbar=Scrollbar(self.outerframe,orient=HORIZONTAL)
        self.hScrollbar.pack(side='bottom',fill=X)
        
        self.hScrollbar.config(command=self.listbox.xview) 
        self.vScrollbar.config(command=self.listbox.yview)
        self.listbox.config(yscrollcommand=self.vScrollbar.set,
                            xscrollcommand=self.hScrollbar.set)

       
        
        self.authorLabel=Label(self.mainFrame)
        self.authorLabel.configure(text='Made By Martin So 2022')
        
        self.mainFrame.configure(cursor="arrow", height="150", width="1000")

        
        #configure frames 
        self.filename.place(anchor="nw", relx="0.01", rely="0.06", x="0", y="0")
        self.filepath.place(anchor="nw", relx="0.07", rely="0.06", x="0", y="0")
        self.keyword.place(anchor="nw", relx="0.01", rely="0.30", x="0", y="0")
        self.keywordEntry.place(anchor="nw", relx="0.07", rely="0.30", x="0", y="0")
        self.resultsLabel.place(anchor="nw", relx="0.55", rely="0.06", x="0", y="0")
        self.browseButton.place(anchor="nw", relx="0.05", rely="0.49", x="0", y="0")
        self.searchButton.place(anchor="nw", relx="0.10", rely="0.49", x="0", y="0")
        self.clearButton.place(anchor="nw", relx="0.15", rely="0.49", x="0", y="0")
        self.authorLabel.place(anchor='nw',relx='0.86',rely='0.85')
        self.mainFrame.place(anchor="nw", x="0", y="0")
        self.resultsFrame.place(anchor="nw", relx="0.25", rely="0.17", y="0")
        
        
        

        
        mainloop()
    def find_doc(self):
        try:
            getFileName= self.filePthSearchPath.get()
    
            getKeyword= str(self.keywordEntry.get())
            stringFormat=''
            os.chdir(getFileName)
        except OSError:
            tkinter.messagebox.showinfo('Error', 'Please enter a valid entry')
        files =[]
        for dirpath, dirnames, filenames in os.walk(getFileName):
            for filename in [f for f in filenames if f.endswith('.docx')]:
                files.append(os.path.join(dirpath,filename))
        getSearch= self.searchAlgo(files,getKeyword) #search file name in keyword
        currentItems=str(list(self.listbox.get(0,END)))
        if currentItems!='': #if listbox is not blank 
            self.listbox.delete(0,END)
            for x in getSearch:
                self.listbox.insert(END,x)
        else:
            for x in getSearch:
                self.listbox.insert(END,x)
                
    def searchAlgo(self,lstofFiles,keyword):
        storage=[]
        for i in range(len(lstofFiles)): 
            text = docx2txt.process(lstofFiles[i]) 
            if keyword.lower() in text.lower(): 
                storage.append(lstofFiles[i])
        cleanStorage=self.cleanList(storage)
        return cleanStorage
    
    
    def clear(self):
        clearResult=''
        self.listbox.delete(0,END)
        self.filePthSearchPath.set(clearResult)
        self.keywordEntry.delete(0,'end')
        
    def searchFile(self):
        getPath=filedialog.askdirectory(parent=self.root,initialdir="/",title='Please select a directory')
        self.filePthSearchPath.set(getPath)
    
    def cleanList(self,lst):
        for x in  range(len(lst)):
            lst[x]=lst[x].replace('/','//')
        return lst
            
        
        

        
        
FileAppV3()
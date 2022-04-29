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

        self.resultsOutputLabeltext=StringVar()
        self.resultsOutputLabel = Label(self.mainFrame)
        self.resultsOutputLabel.configure(textvariable=self.resultsOutputLabeltext)

        
        self.mainFrame.configure(cursor="arrow", height="150", width="1000")

        
        #configure frames 
        self.filename.place(anchor="nw", relx="0.01", rely="0.06", x="0", y="0")
        self.filepath.place(anchor="nw", relx="0.07", rely="0.06", x="0", y="0")
        self.keyword.place(anchor="nw", relx="0.01", rely="0.30", x="0", y="0")
        self.keywordEntry.place(anchor="nw", relx="0.07", rely="0.30", x="0", y="0")
        self.resultsLabel.place(anchor="nw", relx="0.45", rely="0.06", x="0", y="0")
        self.browseButton.place(anchor="nw", relx="0.05", rely="0.49", x="0", y="0")
        self.searchButton.place(anchor="nw", relx="0.10", rely="0.49", x="0", y="0")
        self.clearButton.place(anchor="nw", relx="0.15", rely="0.49", x="0", y="0")
        self.resultsOutputLabel.place(anchor="nw", relx="0.20", rely="0.35", x="0", y="0")
        self.mainFrame.place(anchor="nw", x="0", y="0")
        
        

        
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
        
        getSearch= self.searchAlgo(files,getKeyword)
        for x in getSearch:
            stringFormat+='{}\n'.format(x)
        self.resultsOutputLabeltext.set(stringFormat)
        
    
    def searchAlgo(self,lstofFiles,keyword):
        storage=[]
        for i in range(len(lstofFiles)): 
            text = docx2txt.process(lstofFiles[i]) 
            if keyword.lower() in text.lower(): 
                storage.append(lstofFiles[i])
        return storage
    
    
    def clear(self):
        clearResult=''
        self.resultsOutputLabeltext.set(clearResult)
        self.filePthSearchPath.set(clearResult)
        self.keywordEntry.delete(0,'end')
    def searchFile(self):
        getPath=filedialog.askdirectory(parent=self.root,initialdir="/",title='Please select a directory')
        self.filePthSearchPath.set(getPath)
        

        
        
        
    
FileAppV3()
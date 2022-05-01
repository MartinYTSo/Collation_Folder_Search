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


class FileApp:
    def __init__(self):
        #build app layout
        self.root=Tk()
        self.root.geometry('1000x150')
        self.root.resizable(False,False)
        self.root.title('Collation Folder Search')
        
        self.mainFrame = Frame(self.root, container="false") 
        self.filename = Label(self.mainFrame)
        self.filename.configure(compound="top", text="File Name") #file name label

        self.filePthSearchPath=StringVar()
        
        self.filepath = Label(self.mainFrame)
        self.filepath.configure(textvariable=self.filePthSearchPath) #label for filepath
        
        self.keyword = Label(self.mainFrame)
        self.keyword.configure(text="Keyword") #keyword label
 
        self.keywordEntry = Entry(self.mainFrame) #entry for keyword 
        
        self.resultsLabel = Label(self.mainFrame)
        self.resultsLabel.configure(text="Results") #results label

        self.browseButton = Button(self.mainFrame,command=self.searchFile)
        self.browseButton.configure(text="Browse") #browse button

        self.searchButton = Button(self.mainFrame)
        self.searchButton.configure(font="TkDefaultFont", text="Search",command=self.find_doc) #search button

        self.clearButton = Button(self.mainFrame)
        self.clearButton.configure(text="Clear",command=self.clear) #clear button
        
        self.resultsFrame=Frame(self.mainFrame) #make frame for listbox 
        self.resultsFrame.configure(height=400, takefocus=False, width=450)
        self.outerframe=Frame(self.resultsFrame,height=400,width=450)
        self.outerframe.pack(padx=20,pady=20) #pack frame 
        
        self.innerFrame=Frame(self.outerframe) #innerframe for horizontal scrolling
        self.innerFrame.pack() #pack frame 
        
        self.listbox=Listbox(self.innerFrame,height=5,width=90) #create listbox
        self.listbox.pack(side='left')
        
        self.vScrollbar=Scrollbar(self.innerFrame,orient=VERTICAL) #orient scrollbar to vertical scrolling 
        self.vScrollbar.pack(side='right',fill=Y)
        
        self.hScrollbar=Scrollbar(self.outerframe,orient=HORIZONTAL) #orient scrollbar to horizontal scrolling
        self.hScrollbar.pack(side='bottom',fill=X)
        
        self.hScrollbar.config(command=self.listbox.xview)  
        self.vScrollbar.config(command=self.listbox.yview)
        self.listbox.config(yscrollcommand=self.vScrollbar.set, #allow user to scroll horizontally or vertically
                            xscrollcommand=self.hScrollbar.set)

       
        
        self.authorLabel=Label(self.mainFrame) #main frame 
        
        self.authorLabel.configure(text='Made By Martin So 2022') #add author label
        
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
            getFileName= self.filePthSearchPath.get() #get file path from entry box
            getKeyword= str(self.keywordEntry.get()) #key keyword from entry box
            stringFormat=''
            os.chdir(getFileName) #change file path based on entry
        except OSError:
            tkinter.messagebox.showinfo('Error', 'Please enter a valid entry') #throw error if filename missing
        files =[]
        for dirpath, dirnames, filenames in os.walk(getFileName): #traverse through directory and see if it ends with .docx
            for filename in [f for f in filenames if f.endswith('.docx')]:
                files.append(os.path.join(dirpath,filename)) #append filepath if true 
        getSearch= self.searchAlgo(files,getKeyword) #search file name in keyword
        currentItems=str(list(self.listbox.get(0,END)))  
        if currentItems!='': 
            self.listbox.delete(0,END) #remove current listbox items if there is any already
            for x in getSearch:
                self.listbox.insert(END,x) # return new results if there are current items
        else:
            for x in getSearch:
                self.listbox.insert(END,x)
                
    def searchAlgo(self,lstofFiles,keyword):
        storage=[]
        for i in range(len(lstofFiles)): 
            text = docx2txt.process(lstofFiles[i]) #processing word doc files to notepad text 
            if keyword.lower() in text.lower():
                storage.append(lstofFiles[i]) #append it to a list
        cleanStorage=self.cleanList(storage)
        return cleanStorage
    
    
    def clear(self):
        clearResult='' #remove everything if clear button is passed
        self.listbox.delete(0,END)
        self.filePthSearchPath.set(clearResult)
        self.keywordEntry.delete(0,'end')
        
    def searchFile(self): #ask user for directory
        getPath=filedialog.askdirectory(parent=self.root,initialdir="/",title='Please select a directory')
        self.filePthSearchPath.set(getPath)
    
    def cleanList(self,lst): #format string
        for x in  range(len(lst)):
            lst[x]=lst[x].replace('/','\\')
        return lst
            
        
        

        
        
FileApp()

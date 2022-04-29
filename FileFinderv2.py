# -*- coding: utf-8 -*-
"""
Created on Thu Apr 28 13:07:51 2022

@author: Martin
"""



from tkinter import *
import docx2txt
import os 
from tkinter import filedialog
import tkinter.messagebox

class FileApp:
    def __init__(self):
        self.root= Tk()
        self.root.title('File Search App')
        self.lframe=Frame()
        self.l1=Frame(self.lframe)
        self.l2=Frame(self.lframe)
        self.l3=Frame(self.lframe)
        
        self.rframe=Frame()
        self.r1=Frame(self.rframe)
        self.r2=Frame(self.rframe)
        
        
        #text and entry
        self.filePth=StringVar()
        self.Lab1=Label(self.l1,text='Enter File Path: ')
        self.entry1=Label(self.l1,textvariable=self.filePth)
        self.Lab2=Label(self.l2,text='Enter keyword: ')
        self.entry2=Entry(self.l2,width=20)
        self.button0=Button(self.l3,text='Browse',command=self.searchFile)
        self.button1=Button(self.l3, text='Search',command=self.find_doc)
        self.button2=Button(self.l3,text='Clear',command=self.clear)
        
        self.rlab1=Label(self.r1,text='Results')
        self.getResults=StringVar()
        self.rlab2=Label(self.r2,textvariable=self.getResults)
        
        #pack frames
        self.lframe.pack(side='left')
        self.rframe.pack(side='right')
        self.l1.pack()
        self.l2.pack()
        self.l3.pack()
        self.r1.pack()
        self.r2.pack()
        
        
        #pack labels
        self.Lab1.pack(side='left')
        self.entry1.pack()
        self.Lab2.pack(side='left')
        self.entry2.pack()
        self.button0.pack(side='left')
        self.button1.pack(side='left')
        self.button2.pack()
        self.rlab1.pack()
        self.rlab2.pack()
        
        mainloop()
    def find_doc(self):
        try:
            getFileName= self.filePth.get()
    
            getKeyword= str(self.entry2.get())
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
        self.getResults.set(stringFormat)
        
    
    def searchAlgo(self,lstofFiles,keyword):
        storage=[]
        for i in range(len(lstofFiles)): 
            text = docx2txt.process(lstofFiles[i]) 
            if keyword.lower() in text.lower(): 
                storage.append(lstofFiles[i])
        return storage
    
    
    def clear(self):
        clearResult=''
        self.getResults.set(clearResult)
        self.filePth.set(clearResult)
        self.entry2.delete(0,'end')
    def searchFile(self):
        getPath=filedialog.askdirectory(parent=self.root,initialdir="/",title='Please select a directory')
        self.filePth.set(getPath)
        

        
        
        
    
FileApp()
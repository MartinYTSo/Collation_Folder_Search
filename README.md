
## File Path Finder (Collation Folder Searcher): Retrive the Path of your Word Docs Faster

### How to Run
- Standalone App:
   1. Download `FilePathFinder` folder
   2. Double Click  `app`
   3. Double Click `FilePathFinder.exe`
- With IDE
  1. Download and run FileFinderv4.py 
  
The application will look like this upon opening:

![image](https://user-images.githubusercontent.com/72810148/166083223-720924e9-bc79-48fd-9a75-9e76b60cb45f.png)
### Background 
Imagine if you are working as an administrative assistant in an office and you are managing Word documents. Your boss comes up to you and asks to retrieve a Word document that was dated 5 months back mentioning about a yearly report. You wouldn't want to manually search through 5 months worth of documents in a folder would you? After all, he wants it ASAP.

Introducing the File Path Finder (Collation Folder Searcher). The File Path Finder will allow you to retrieve the path(s) of the file in a matter of seconds.

### Source code
Download `FileFinderv4.py` and run it through the IDE of your choice. Please ensure that the version of Python is Version 3 or later.
### Features
- A `Browse` button to select the folder
- An entry box to put your desired keyword in
- A `Search` button to search the file path(s) where your keywords are located
- A `Clear` button to clear all entries 


Here is an example of the application in practice. I have created a `TestFolder` with random word documents from my past classes from school. I have entered the keyword `exam`. The following results have are listed here as shown: 
![image](https://user-images.githubusercontent.com/72810148/166084303-64a5143b-31c5-40d5-a441-a262db26fd4a.png)

### Dependencies/Requirements
- `Python3.xx`
- [docx2txt](https://github.com/ankushshah89/python-docx2txt) for converting word documents to raw text in order to match with keyword inputted

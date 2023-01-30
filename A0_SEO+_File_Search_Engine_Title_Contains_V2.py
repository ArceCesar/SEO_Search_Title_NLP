# -*- coding: utf-8 -*-
"""
Created on Mon Jan 16 07:40:11 2023
A0_SEO+_File_Search_Engine_Title_Contains_V2
@author: cesar
"""

import datetime
import pathlib
from queue import Queue
from threading import Thread
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askdirectory
import customtkinter
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap import utility
from ttkbootstrap import Window
from ttkthemes import ThemedTk
import subprocess
import os
import clipboard
from tkinter import filedialog
from tkinter import messagebox
import PyPDF2
import docx
import datetime



class FileSearchEngine(ttk.Frame):

    queue = Queue()
    searching = False

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.pack(fill=BOTH, expand=True)

        # application variables
        _path = pathlib.Path().absolute().as_posix()
        self.path_var = ttk.StringVar(value=_path)
        self.term_var = ttk.StringVar(value='pdf')
        self.type_var = ttk.StringVar(value='endswidth')

        # header and labelframe option container
        option_text = "Complete the form to begin your search"
        self.option_lf = ttk.Labelframe(self, text=option_text, padding=15) #self
        self.option_lf.pack(fill=X, expand=YES, anchor=N)

        self.create_path_row()
        self.create_term_row()
        self.create_type_row()
        self.create_results_view()

        self.progressbar = ttk.Progressbar(
            master=self, 
            mode=INDETERMINATE, 
            bootstyle=(STRIPED, INFO) #SUCCESS)
        )
        self.progressbar.pack(fill=X, expand=YES, anchor=S)

    def create_path_row(self):
        """Add path row to labelframe"""
        path_row = ttk.Frame(self.option_lf)
        path_row.pack(fill=X, expand=YES)
        path_lbl = ttk.Label(path_row, text="Path", width=8)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        path_ent = ttk.Entry(path_row, textvariable=self.path_var)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttk.Button(
            master=path_row, 
            text="Browse", 
            command=self.on_browse, 
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5)

    def create_term_row(self):
        """Add term row to labelframe"""
        term_row = ttk.Frame(self.option_lf)
        term_row.pack(fill=X, expand=YES, pady=15)
        term_lbl = ttk.Label(term_row, text="Term", width=8)
        term_lbl.pack(side=LEFT, padx=(15, 0))
        term_ent = ttk.Entry(term_row, textvariable=self.term_var)
        term_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        search_btn = ttk.Button(
            master=term_row, 
            text="Search", 
            command=self.on_search, 
            bootstyle=OUTLINE, 
            width=8
        )
        search_btn.pack(side=LEFT, padx=5)
        term_ent.bind("<Return>", lambda event: search_btn.invoke())
        self.option_lf.bind("<Return>", lambda event: search_btn.invoke())

    def create_type_row(self):
        """Add type row to labelframe"""
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text="Type", width=8)
        type_lbl.pack(side=LEFT, padx=(15, 0))

        contains_opt = ttk.Radiobutton(
            master=type_row, 
            text="Contains", 
            variable=self.type_var, 
            value="contains"
        )
        contains_opt.pack(side=LEFT)

        startswith_opt = ttk.Radiobutton(
            master=type_row, 
            text="StartsWith", 
            variable=self.type_var, 
            value="startswith"
        )
        startswith_opt.pack(side=LEFT, padx=15)

        endswith_opt = ttk.Radiobutton(
            master=type_row, 
            text="EndsWith", 
            variable=self.type_var, 
            value="endswith"
        )
        endswith_opt.pack(side=LEFT)
        endswith_opt.invoke()
        
        global switch_var
        switch_var = customtkinter.StringVar(value="on") 
        def switch_event():
            if switch_var.get() == "darkly":
                app = ttk.Window(themename="darkly")
                app.withdraw()
            elif switch_var.get() == "journal":
                app = ttk.Window(themename="journal")
                app.withdraw()
            else:
                pass
            
            #print("switch toggled, current value:", switch_var.get())
  
        self.switch_1 = customtkinter.CTkSwitch(master=type_row, text="Appearance", command=switch_event,
                                   variable=switch_var, onvalue="journal", offvalue="darkly",
                                   fg_color=('silver','black'))
        self.switch_1.pack(padx=20, side=LEFT) # padx=20, pady=10)
        
        clear_btn = ttk.Button(
            master=type_row, 
            text="Clear", 
            command=self.clear_all, 
            bootstyle=OUTLINE, 
            width=8)
        clear_btn.pack(padx=5, side=RIGHT)
        
        copy_path_button = ttk.Button(
            master=type_row, 
            text="CopyPath", 
            command=self.copy_path, 
            bootstyle=OUTLINE, 
            width=8)
        copy_path_button.pack(padx=5, side=RIGHT)

    def clear_all(self):
        for child in self.resultview.get_children():
            self.resultview.detach(child)
        print("Clear Exec.")
    
    def create_results_view(self):
        """Add result treeview to labelframe"""
        self.resultview = ttk.Treeview(self, 
                                       bootstyle=INFO, 
                                       #columns=[0, 1, 2, 3, 4],
                                       columns=("Name", "Date", "Type", "Size", "Path"),
                                       show=HEADINGS,
                                       selectmode='browse')
        self.resultview.place(x=0,y=180,relwidth=1,relheight=1, height=-210)
        
        #self.resultview.pack(fill=BOTH, expand=YES, pady=10) 
        
        # Creating ScrollBars
        self.scrollx = ttk.Scrollbar(self.resultview, orient="horizontal", command=self.resultview.xview)
        self.scrollx.pack(side ='bottom', fill ='x')
        
        self.scrolly = ttk.Scrollbar(self.resultview, orient="vertical", command=self.resultview.yview)
        self.scrolly.pack(side ='right', fill ='y')
        
        self.resultview["xscrollcommand"] = self.scrollx.set
        self.resultview["yscrollcommand"] = self.scrolly.set

        # setup columns and use `scale_size` to adjust for resolution
        self.resultview.heading(0, text='Name', anchor=W, command=lambda: self.sort_tree("Name")) 
        self.resultview.heading(1, text='Modified', anchor=W, command=lambda: self.sort_tree("Date")) 
        self.resultview.heading(2, text='Type', anchor=E, command=lambda: self.sort_tree("Type")) 
        self.resultview.heading(3, text='Size', anchor=E, command=lambda: self.sort_tree("Size"))
        self.resultview.heading(4, text='Path', anchor=W, command=lambda: self.sort_tree("Path")) 
        self.resultview.bind("<Double-1>", self.on_select)
        self.resultview.column(
            column=0, 
            anchor=W, 
            width=utility.scale_size(self, 250), 
            stretch=False
        )
        self.resultview.column(
            column=1, 
            anchor=W, 
            width=utility.scale_size(self, 140), 
            stretch=False
        )
        self.resultview.column(
            column=2, 
            anchor=E, 
            width=utility.scale_size(self, 50), 
            stretch=False
        )
        self.resultview.column(
            column=3, 
            anchor=E, 
            width=utility.scale_size(self, 60), 
            stretch=False
        )
        self.resultview.column(
            column=4, 
            anchor=W, 
            width=utility.scale_size(self, 200)
        )

    def on_browse(self):
        """Callback for directory browse"""
        path = askdirectory(title="Browse directory")
        if path:
            self.path_var.set(path)

    def on_search(self):
        """Search for a term based on the search type"""
        search_term = self.term_var.get()
        search_path = self.path_var.get()
        search_type = self.type_var.get()

        if search_term == '':
            return     
            
        import nltk
        from nltk.corpus import stopwords
        nltk.download('stopwords')
        nltk.download('punkt')
        from nltk.corpus import stopwords
        from nltk.tokenize import word_tokenize
            
        text_tokens = word_tokenize(search_term)
        keywords = [word for word in text_tokens if not word in stopwords.words()]
            
        print(keywords)
        self.keywords=keywords

        import os
        global keyword
        for keyword in keywords:
            try:

                # start search in another thread to prevent UI from locking
                Thread(target=FileSearchEngine.file_search, 
                          args=(keyword , search_path, search_type), 
                          daemon=True
                          ).start()
 
                self.progressbar.start(10)
                    
                iid = self.resultview.insert(
                                parent='', 
                                index=END)
                
                self.resultview.item(iid, open=True)
                self.after(100, lambda: self.check_queue(iid))
                
            except OSError as e:
                print("Error: %s : %s" % (f, e.strerror))
            

    def sort_tree(self, col): 
        """Sort tree contents when a column header is clicked on."""
        self.resultview.heading(col, command=lambda: self.sort_tree(col))
        l = [(self.resultview.set(k, col), k) for k in self.resultview.get_children('')]
        l.sort(key=lambda t: t[0])
        for index, (val, k) in enumerate(l):
            self.resultview.move(k, '', index) 

    def copy_path(self):
        """Copy the path of the selected file to the clipboard"""
        cur_item = self.resultview.focus()
        self.file_path = self.resultview.item(cur_item, "values")[4]
        clipboard.copy(self.file_path)
        
    def on_select(self, event):
        """Open the selected file when the Path column is double-clicked"""
        cur_item = self.resultview.focus()
        self.file_path = self.resultview.item(cur_item, "values")[4]
        self.normalized_file_path = self.file_path.replace("/", "\\")
        print(self.file_path)
        print(self.normalized_file_path)
        if os.path.exists(self.file_path):
            subprocess.run(["open", self.file_path]) 
            subprocess.Popen(["open", self.file_path])
        else:
            print("The file specified does not exist.")
            
        if os.path.exists(self.normalized_file_path): 
            subprocess.run(["open", self.normalized_file_path]) 
            subprocess.Popen(["open", self.normalized_file_path])
        else:
            print("The file specified does not exist.")

    def check_queue(self, iid):
        global summary
        """Check file queue and print results if not empty"""
        if all([
            FileSearchEngine.searching, 
            not FileSearchEngine.queue.empty()
        ]):
            filename = FileSearchEngine.queue.get()
            self.insert_row(filename, iid)
            self.update_idletasks(),
            self.after(100, lambda: self.check_queue(iid))
            # summ = self.resultview.insert(parent='', index=tk.END, text=summary)
            # self.resultview.selection_set(summ)
            # self.resultview.see(summ)
        elif all([
            not FileSearchEngine.searching,
            not FileSearchEngine.queue.empty()
        ]):
            while not FileSearchEngine.queue.empty():
                filename = FileSearchEngine.queue.get()
                self.insert_row(filename, iid)
            self.update_idletasks()
            self.progressbar.stop()
        elif all([
            FileSearchEngine.searching,
            FileSearchEngine.queue.empty()
        ]):
            self.after(100, lambda: self.check_queue(iid))
        else:
            summ = self.resultview.insert(parent='', index=tk.END, text=summary)
            # self.resultview.selection_set(summ)
            # self.resultview.see(summ)
            self.progressbar.stop()
            # summ = self.resultview.insert(parent='', index=tk.END, text=summary)
            # self.resultview.selection_set(summ)
            # self.resultview.see(summ)  

    def insert_row(self, file, iid):
        """Insert new row in tree search results"""
        try:
            _stats = file.stat()
            _name = file.stem
            _timestamp = datetime.datetime.fromtimestamp(_stats.st_mtime)
            _modified = _timestamp.strftime(r'%m/%d/%Y  %I:%M:%S%p')
            _type = file.suffix.lower()
            _size = FileSearchEngine.convert_size(_stats.st_size)
            _path = file.as_posix()
            iid = self.resultview.insert(
                parent='', 
                index=END, 
                values=(_name, _modified, _type, _size, _path)
            )
            self.resultview.selection_set(iid)
            self.resultview.see(iid)
        except OSError:
            return

        
    @staticmethod
    def file_search(term, search_path, search_type):
        """Recursively search directory for matching files"""
        FileSearchEngine.set_searching(1)
        if search_type == 'contains':
            FileSearchEngine.find_contains(term, search_path)
        elif search_type == 'startswith':
            FileSearchEngine.find_startswith(term, search_path)
        elif search_type == 'endswith':
            FileSearchEngine.find_endswith(term, search_path)

    @staticmethod
    def find_contains(term, search_path):
        global summary
        """Used to search term One by One"""
        matches = 0
        records = 0
        """Find all files that contain the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    records +=1
                    #if all(term.lower() in file.lower() for term in keys):
                    if term.lower() in file.lower():
                        record = pathlib.Path(path) / file
                        FileSearchEngine.queue.put(record)
                        matches += 1
        FileSearchEngine.set_searching(False)
        print('>> There were {:,d} matches of {} out of {:,d} records searched.'.format(matches, term, records))
        summary = ">> There were {:,d} matches of {} out of {:,d} records searched.".format(matches, term, records)

        
    @staticmethod
    def find_startswith(term, search_path):
        global summary
        matches = 0
        records = 0
        """Find all files that start with the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    records +=1
                    if file.startswith(term):
                        record = pathlib.Path(path) / file
                        FileSearchEngine.queue.put(record)
                        matches += 1
        FileSearchEngine.set_searching(False)
        
        print('>> There were {:,d} matches of {} out of {:,d} records searched.'.format(matches, term, records))
        summary = ">> There were {:,d} matches of {} out of {:,d} records searched.".format(matches, term, records)

        
    @staticmethod
    def find_endswith(term, search_path):
        global summary
        matches = 0
        records = 0
        """Find all files that end with the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    records +=1
                    if file.endswith(term):
                        record = pathlib.Path(path) / file
                        FileSearchEngine.queue.put(record)
                        matches += 1
        FileSearchEngine.set_searching(False)

        print('>> There were {:,d} matches of {} out of {:,d} records searched.'.format(matches, term, records))
        summary = ">> There were {:,d} matches of {} out of {:,d} records searched.".format(matches, term, records)

    
    def insert_summary(self):
        global summary
        try:
            summ = self.resultview.insert(parent='', index=tk.END, text=self.summary)
            self.resultview.selection_set(summ)
            self.resultview.see(summ)
            self.resultview.insert(parent='', index=tk.END, text=self.summary) #iid = 
            self.update_idletasks()
            print(self.summary)
        except OSError:
            return
        self.resultview.insert(parent='', index=tk.END, text=summary) 
        
    @staticmethod
    def set_searching(state=False):
        """Set searching status"""
        FileSearchEngine.searching = state

    @staticmethod
    def convert_size(size):
        """Convert bytes to mb or kb depending on scale"""
        kb = size // 1000
        mb = round(kb / 1000, 1)
        if kb > 1000:
            return f'{mb:,.1f} MB'
        else:
            return f'{kb:,d} KB'        


if __name__ == '__main__':
    app = ttk.Window(title="SEO+ File Search Engine - Titre Contains", themename="journal") #"journal") #, "darkly")
    app.geometry("880x520")
    app.iconbitmap('logo2.ico')
    FileSearchEngine(app)
    app.mainloop()

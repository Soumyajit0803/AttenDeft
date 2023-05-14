import tkinter as tk
import tkinter.ttk as ttk
from pickle import dump,load
import tkinter.filedialog as tkf
from os import startfile, rename, getcwd
from subprocess import call
from platform import system
from tkinter import messagebox as mbx
import tkinter.font as tkfont
from tkinter.filedialog import askdirectory as aod
from shutil import copy

class More(tk.Frame):
    '''miscellaneous works'''
    def __init__(self,parent,controller):
        tk.Frame.__init__(self,parent)
        self.option_add('*TCombobox*Font', tkfont.Font(family='Book Antiqua',size=12))
        self.controller = controller
        self.savedir=''# for export copy
        widgets=tk.LabelFrame(self)
        widgets.grid(row=0,column=1,sticky='nse')
        tk.Label(widgets,text='SETTINGS'+' '*61,font=('calibri',20,'italic bold'),fg='#006700').pack(side='top',anchor='w',padx=10)
        tk.Label(widgets,text='Minimum Time limit for attendance: ',font=('Book Antiqua',13,'bold'),justify=tk.LEFT).place(x = 0, y = 50)
        tk.Label(widgets,text='Edit a class file\n(add/delete a student): ',font=('Book Antiqua',13,'bold'),justify=tk.LEFT).place(x = 0, y = 80)
        tk.Button(widgets,text='Click here to know permissible ways of editing.',font=('CaMbRiA',13,'italic underline bold'),borderwidth=0,fg='#006700',justify=tk.LEFT, cursor = 'hand2',command=self.rules).place(x = 0, y = 140)        
        
        self.pic = tk.PhotoImage(file='sidepane2.PNG')
        tk.Label(self,image=self.pic,bd=0).grid(row=0,column=0)
        self.rule = tk.PhotoImage(file='Rule.PNG')

        #min-time
        with open('ChalkBox.AD','rb') as f:
        	default=load(f)['min_time']
        self.time=ttk.Combobox(widgets,width=5,values=(0.50,0.65,0.75,0.80,0.85),state = 'readonly', font=('calibri',15))
        self.time.set(default)
        self.time.place(x=310,y=50)
        tk.Button(widgets,text='Save',command=self.time_update,font=('',12,''),fg='white',bg='#006700').place(x=400,y=50)

        #modify class
        self.Class_modify=ttk.Combobox(widgets,width=15,values=(),postcommand = self.update_combo,state = 'readonly', font=('calibri',15))
        self.Class_modify.place(x=210,y=100)
        self.Class_modify.set('select class')
        tk.Button(widgets,text='Open',command=self.open_class,font=('',12,''),fg='white',bg='#006700').place(x=400,y=100)
        ttk.Separator(widgets,orient='horizontal').place(relx=0,rely=0.60,relwidth=1,relheight=1)

        #Export copy
        tk.Label(widgets,text='Export a copy of\na class file: ',font=('Book Antiqua',13,'bold'),justify=tk.LEFT).place(x = 0, y = 190) #230
        self.Class_copy=ttk.Combobox(widgets,width=15,values=(),postcommand = self.update_combo,state = 'readonly', font=('calibri',15))
        self.Class_copy.set('select class')
        self.Class_copy.place(x=170,y=200)
        self.folder_loc=tk.Label(widgets,text='No folder selected',font=('calibri',13,'italic'),justify=tk.LEFT)
        self.folder_loc.place(x = 170, y = 230)        
        tk.Button(widgets,text='Select folder to\nexport file',fg='white',bg ='#006700' ,font=('',12,''),command = self.choose_folder).place(x=360,y=200)
        tk.Button(widgets, text = 'Export',fg='white',bg ='#006700' ,font=('',12,''),command=self.export).place(x=400,y=255)
        
        #documentation corner
        tk.Label(widgets,text='ABOUT ',font=('calibri',20,'italic'),fg='#006700').place(x = 140, y = 320)
        tk.Label(widgets,text='AttenDeft',font=('calibri',20,''),fg='#006700').place(x = 230, y = 320)
        tk.Button(widgets,text='Software Documentation',command=lambda :self.open_file("AttenDeft_docu.html"),font=('',12,'italic underline bold'),fg='#006700',borderwidth=0, cursor = 'hand2').place(x=150,y=360)
        tk.Button(widgets,text='References',command=lambda :self.open_file("Attendeft_references.html"),font=('',12,'italic underline bold'),fg='#006700',borderwidth=0, cursor = 'hand2').place(x=195,y=390)        
        tk.Button(widgets,text='BACK',font=('',12),bg='#006700',fg='white',command=lambda: controller.show_frame("Main_Page")).place(x=0,y=475)
        
    def time_update(self):
        '''updates min-time'''
        newT=self.time.get()
        with open('ChalkBox.AD','rb+') as f:
            temp=load(f)
            temp['min_time']=float(newT)
            f.seek(0)
            f.truncate()
            dump(temp,f)
        mbx.showinfo('Success','minimum time limit updated successfully')

    def open_class(self):
        '''auto-open a workbook'''
        os = system()
        if self.Class_modify.get() == 'select class':
            mbx.showwarning('Unsufficient data','Class file not selected')                
        else:
            file=f"{self.Class_modify.get()}.xlsx"
            self.open_file(file)

    	
    def update_combo(self):
        '''functionality of combobox'''
        with open('ChalkBox.AD','rb') as f:
            d = load(f)
        self.Class_modify['values']= d.get('classes',())
        self.Class_copy['values']= d.get('classes',())

    def choose_folder(self):
        '''input folder to export copy of workbook'''
        self.savedir= aod()
        self.folder_loc.configure(text = self.savedir)

    def export(self):
        '''export copy of workbook'''
        if self.savedir=='':
            mbx.showwarning('Unsufficient data','Directory not selected')
        elif self.Class_copy.get()=='select class':
            mbx.showwarning('Unsufficient data','Class file not selected')
        else:
            copy(getcwd()+'\\'+self.Class_copy.get()+'.xlsx', self.savedir)
            mbx.showinfo('Success',f"Class {self.Class_copy.get()}'s copy\nexported to {self.savedir}")
    
    def open_file(self,filename):
        '''auto-open the desired file'''
        os=system()
        if os=='Darwin': #macOS
            call(('open',filename))
        elif os=='Windows':
            startfile(filename)
        else:
            call(('xdg-open',filename))
    
    def rules(self):
        '''display the rules of editing'''
        w=tk.Toplevel()
        w.resizable(0,0)
        w.title('Editing Rules')
        w.iconbitmap("progico.ico")
        tk.Label(w,image=self.rule,bd=0).grid(row=0,column=0)
        w.mainloop()


import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkf
from tkinter import messagebox as mbx
from Create_Class import AddClass
from datetime import date as d
from pickle import load,dump
from os import remove,getcwd
             
class Edit_Class(tk.Frame):
    '''for creation/deletion of classes'''
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.pic = tk.PhotoImage(file='sidepane4.PNG')
        tk.Label(self,image=self.pic,bd=0).grid(row=0,column=0)

        widgets=tk.LabelFrame(self)
        widgets.grid(row=0,column=1,sticky='nse')
        tk.Label(widgets,text='CREATE CLASS',font=('calibri',20,'italic bold'),fg='#0662CA').pack(side='top',anchor='w',padx=10)
        
		#The text to write in description
        define="Select your class name and the session to create an Excel workbook for your class. Session is of 1 year.\nYou will require to upload an Excel file containing all student names with roll numbers. Attendance records of this class will be stored by this Excel workbook."
        tk.Label(widgets,text=define,font=('Book Antiqua',13),wraplength=490,justify=tk.LEFT).pack(padx=7,pady=5)
		
        tk.Label(widgets,text='Class Name: ',font=('',12,'bold'),fg='#0662CA').place(x=0,y=175)
        tk.Label(widgets,text='Session: ',font=('',12,'bold'),fg='#0662CA').place(x=0,y=215)
        self.location=tk.Label(widgets,text='No file chosen',font=('',12,'italic'))
        self.location.place(x=360,y=175)
        curyear=d.today().year#for session

        self.class_val=tk.Entry(widgets,width=15,font=('cambria',12,'bold'))       
        self.class_val.place(x=120,y=175)
        self.year=ttk.Combobox(widgets,state='readonly',width=4,font=('cambria',12,'bold'))              
        self.year['values']=[i for i in range(curyear-3,curyear+3)]
        self.year.bind('<<ComboboxSelected>>',self.to_what)
        self.year.set("yyyy")
        self.year.place(x=185,y=215)
        
        self.monthlist=("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC")
        self.month=ttk.Combobox(widgets,state='readonly',width=4,font=('cambria',12,'bold'),values=self.monthlist)
        self.month.bind('<<ComboboxSelected>>',self.to_what)
        self.month.set("mm")        
        self.month.place(x=120,y=215)
        
        self.e_session = tk.Label(widgets,text='to',font=('',12,'bold'),fg='#0662CA')
        self.e_session.place(x=120,y=250)

        tk.Button(widgets,text='UPLOAD\nFILE: ',font=('',12),command=self.getfile,bg='#0662CA',fg='white').place(x=270,y=175)
        tk.Button(widgets,text='SUBMIT',font=('',12),command=self.submitfile,bg='#0662CA',fg='white').place(x=270,y=230)
        
        #DELETE CLASS PORTION
        ttk.Separator(widgets,orient='horizontal').place(relx=0,rely=0.57,relwidth=1,relheight=1)
        tk.Button(widgets,text='BACK',font=('',12),command=lambda: controller.show_frame("Main_Page"),bg='#0662CA',fg='white').place(x=0,y=460)
        
        tk.Label(widgets,text='DELETE CLASS',font=('calibri',20,'italic bold'),fg='#0662CA').place(x=0,y=300)
        tk.Label(widgets,text='Select a class you wish to remove. Please note that this task is irreversible.',font=('Book Antiqua',13),wraplength=490,justify=tk.LEFT).place(x=0,y=345)

        self.Class=ttk.Combobox(widgets,width=15,values=(),state = 'readonly',postcommand=self.update_combo, font=('calibri',15))
        tk.Button(widgets,text='REMOVE',font=('',12),command=self.remove,bg='#0662CA',fg='white').place(x=270,y=400)
        self.Class.set('select class')
        self.Class.place(x=0,y=400)
        
    def submitfile(self):
        '''parse the csv, write data, prompt success msg'''
        if self.class_val.get() =="":
            mbx.showwarning('missing','Class field cannot be empty')

        elif self.year.get()=='' or self.month.get()=="":
            mbx.showwarning('missing','session field cannot be empty')

        elif self.location['text']=='No file chosen':
            mbx.showwarning('missing','file not selected')

        else:

            name = self.class_val.get()+f' {self.year.get()}-{(int(self.year.get())+1)%100}'
            mbx.showinfo('Success','Your class workbook\n has been created!')
            AddClass(name,[int(self.year.get()),int(self.year.get())+1],self.file.name,self.month.get())
            

    def getfile(self):
        '''Returns self.location of file obj'''

        self.file=tkf.askopenfile(mode ='r', filetypes =[('EXCEL files', '*.xlsx')])
        if self.file:
            self.location.config(text=self.file.name.split('/')[-1])

    def to_what(self,event):
        '''shows end month-year of session(12 months in a session)'''
        if self.year.get()!='yyyy' and self.month.get()!='mm':
            curmon=self.monthlist.index(self.month.get())
            self.e_session.configure(text=f'to {self.monthlist[curmon-1]} {int(self.year.get())+ bool(curmon)}')

    def update_combo(self):
        '''functionality of combobox'''
        f = open('ChalkBox.AD','rb')
        d = load(f)
        f.close()
        self.Class['values']= d.get('classes',())

    def remove(self):
        '''delete a class workbook'''
        del_name=self.Class.get()
        if del_name=='select class':
            mbx.showwarning('Empty field','Select a class to remove')
            
        else:
            ask=mbx.askyesno('Behold','Are you sure?')
            if ask:
                with open('ChalkBox.AD','rb+') as f:
                    d = load(f)
                    d['classes'].remove(del_name)
                    f.seek(0)
                    f.truncate()
                    dump(d, f)
                remove(getcwd()+'\\'+self.Class.get()+'.xlsx')
                mbx.showinfo('Success',f'Your class {del_name} has been removed!')
                self.Class.set('select class')
    


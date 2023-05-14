import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox as mbx
from StudentAnalysis import Report
from pickle import load

class Analysis(tk.Frame):
    '''analyse attendance for a class'''
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller                
        self.pic = tk.PhotoImage(file='sidepane1.PNG')
        tk.Label(self,image=self.pic,bd=0).grid(row=0,column=0)

        widgets=tk.LabelFrame(self)
        widgets.grid(row=0,column=1,sticky='ns')
        tk.Label(widgets,text='ANALYSE',font=('calibri',30,'italic bold'),fg='#CC04CA').pack(side='top',anchor='w',padx=10)
        
        define='Here, you can see and analyse how your students are attending in class. You can get to know the overall attendance % of each student. Further you can also see attendance % of students monthwise.'
        tk.Label(widgets,text=define,font=('Book Antiqua',13),wraplength=512,justify=tk.LEFT).pack(padx=7,pady=5)
        tk.Label(widgets,text='Select Class: ',font=('',12,'bold'),fg='#CC04CA').place(x=0,y=165)

        self.Class=ttk.Combobox(widgets,width=18,state = 'readonly',values = (), postcommand=self.update_combo,font=('baskerville old face',15))
        self.Class.set('select class')
        self.Class.place(x=110,y=160)

        tk.Button(widgets,text='GENERATE\nDATA',font=('',12),command = self.showData,bg='#CC04CA',fg='white').place(x=0,y=200)
        tk.Button(widgets,text='BACK',font=('',12),command=lambda: controller.show_frame("Main_Page"),bg='#CC04CA',fg='white').place(x=0,y=475)

        self.floa=tk.PhotoImage(file='analysispic.png')
        tk.Label(widgets,image=self.floa).place(x =120,y=200)
       
    def showData(self):
        '''Opens the report for class'''   
        if self.Class.get()=='select class':
            mbx.showwarning('missing','class field cannot be empty')

        else:
            Report(classfile=self.Class.get()+'.xlsx') 

    def update_combo(self):
        '''functionality of combobox'''
        f = open('ChalkBox.AD','rb')
        d = load(f)
        f.close()
        self.Class['values']= d.get('classes',())
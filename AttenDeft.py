import tkinter as tk
from Attendance import Attendance               
from Analysis import Analysis
from Edit_Class import Edit_Class
from More import More

class Frame_Changer(tk.Tk):
    '''controlling agent'''
    def __init__(self):
        tk.Tk.__init__(self)#calling Tk class

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        
        self.iconbitmap("progico.ico")
        self.title('AttenDeft')
        
        self.geometry('730x510')
        self.resizable(0,0)
        self.frames = {}
        for F in (Main_Page,Attendance,Analysis,Edit_Class,More):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("Main_Page")
        
    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()

class Main_Page(tk.Frame):
    '''main display page'''
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        #controller is the "self" of the 'Frame_Changer' 
        #this will give Frame_Changer explicit CONTROL over this page & other pages
        self.Title=tk.PhotoImage(file='MAINTITLE.png')        
        tk.Label(self,image=self.Title).grid(row=0,column=0)
        
        self.btnbag=tk.LabelFrame(self,bg='khaki')
        self.btnbag.grid(row=1,column=0)
        self.btn_img = {}
        for i in ['Attendance.PNG','More.PNG','Edit_Class.PNG','Analysis.PNG']:            
            self.btn_img[i]=tk.PhotoImage(file = i) #Image object
            
        #Button creation
        tk.Button(self.btnbag,cursor='hand2',borderwidth=5,
                image = self.btn_img['Attendance.PNG'],relief = tk.RAISED,
                command=lambda: controller.show_frame('Attendance')).grid(row=0,column=0,padx=15,pady=10)
        tk.Button(self.btnbag,cursor='hand2',borderwidth=5,
                image = self.btn_img['Edit_Class.PNG'],relief = tk.RAISED,
                command=lambda: controller.show_frame('Edit_Class')).grid(row=0,column=1,padx=15,pady=10)
        tk.Button(self.btnbag,cursor='hand2',borderwidth=5,
                image = self.btn_img['Analysis.PNG'],relief = tk.RAISED,
                command=lambda: controller.show_frame('Analysis')).grid(row=1,column=0,padx=15,pady=10)
        tk.Button(self.btnbag,cursor='hand2',borderwidth=5,
                image = self.btn_img['More.PNG'],relief = tk.RAISED,
                command=lambda: controller.show_frame('More')).grid(row=1,column=1,padx=15,pady=10)

if __name__ == "__main__":
    AttenDeft = Frame_Changer()
    AttenDeft.mainloop()
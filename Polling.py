from Tkinter import *
import tkMessageBox
import ttk


def vp_start_gui():
    from mongo_operation import mongo_oper
    global val, w, root, MDB,v2,v3
    root = Tk()
    v2 = StringVar()
    v3 = StringVar()
    top = Create_Polling(root)
    MDB = mongo_oper()
    init(root, top)
    root.mainloop()

w = None


def init(top, gui, *args, **kwargs):
    global w, top_level, root, MDB
    w = gui
    top_level = top
    root = top

def destroy_window():
    global top_level
    top_level.destroy()
    top_level = None

class Create_Polling:
    def __init__(self, top=None):
        self.combo2 = ''
        self.combo3 = ''
        self._bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        self._fgcolor = '#000000'  # X11 color: 'black'
        self._compcolor = '#d9d9d9' # X11 color: 'gray85'
        self._ana1color = '#d9d9d9' # X11 color: 'gray85'
        self._ana2color = '#d9d9d9' # X11 color: 'gray85'
        self.font3 = "-family Arial -size 9 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"
        self.font4 = "-family Arial -size 10 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"
        self.font5 = "-family Arial -size 11 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.',background=self._bgcolor)
        self.style.configure('.',foreground=self._fgcolor)
        self.style.configure('.',font=self.font3)
        self.style.map('.',background=
            [('selected', self._compcolor), ('active',self._ana2color)])

        top.geometry("538x435+513+136")
        top.title("Create Polling")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")

        self.Label1 = Label(top)
        self.Label1.place(relx=0.09, rely=0.19, height=31, width=137)
        self.Label1.configure(activebackground="#f9f9f9")
        self.Label1.configure(activeforeground="black")
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=self.font3)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(highlightbackground="#d9d9d9")
        self.Label1.configure(highlightcolor="black")
        self.Label1.configure(text='''Enter the Polling Name:''')

        self.Entry1 = Entry(top)
        self.Entry1.place(relx=0.43, rely=0.21, relheight=0.05, relwidth=0.45)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font=self.font4)
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(highlightbackground="#d9d9d9")
        self.Entry1.configure(highlightcolor="black")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(selectbackground="#c4c4c4")
        self.Entry1.configure(selectforeground="black")

        self.Label2 = Label(top)
        self.Label2.place(relx=0.28, rely=0.05, height=41, width=277)
        self.Label2.configure(activebackground="#f9f9f9")
        self.Label2.configure(activeforeground="black")
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font=self.font5)
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(highlightbackground="#d9d9d9")
        self.Label2.configure(highlightcolor="black")
        self.Label2.configure(relief=RIDGE)
        self.Label2.configure(text='''Polling Creator''')

        self.Label3 = Label(top)
        self.Label3.place(relx=0.11, rely=0.3, height=31, width=127)
        self.Label3.configure(activebackground="#f9f9f9")
        self.Label3.configure(activeforeground="black")
        self.Label3.configure(background="#d9d9d9")
        self.Label3.configure(disabledforeground="#a3a3a3")
        self.Label3.configure(font=self.font3)
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(highlightbackground="#d9d9d9")
        self.Label3.configure(highlightcolor="black")
        self.Label3.configure(text='''Track Polling:''')

        self.TCombobox2 = ttk.Combobox(top)
        self.TCombobox2.place(relx=0.43, rely=0.3, relheight=0.05, relwidth=0.27)
        self.value_list = ["POLLING1","POLLING2",]
        self.TCombobox2.configure(values=self.value_list)
        self.TCombobox2.bind('<<ComboboxSelected>>', self.get_combo2_value)
        self.TCombobox2.configure(textvariable=v2)
        self.TCombobox2.configure(takefocus="")

        self.Label4 = Label(top)
        self.Label4.place(relx=0.13, rely=0.57, height=31, width=127)
        self.Label4.configure(activebackground="#f9f9f9")
        self.Label4.configure(activeforeground="black")
        self.Label4.configure(background="#d9d9d9")
        self.Label4.configure(disabledforeground="#a3a3a3")
        self.Label4.configure(font=self.font3)
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(highlightbackground="#d9d9d9")
        self.Label4.configure(highlightcolor="black")
        self.Label4.configure(text='''Polling Text :''')

        self.Button1 = Button(top)
        self.Button1.place(relx=0.43, rely=0.9, height=34, width=106)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(font=self.font3)
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Create Polling''')
        self.Button1.configure(command=self.click_btn_create_polling)

        self.Label5 = Label(top)
        self.Label5.place(relx=0.13, rely=0.41, height=21, width=117)
        self.Label5.configure(background="#d9d9d9")
        self.Label5.configure(disabledforeground="#a3a3a3")
        self.Label5.configure(font=self.font3)
        self.Label5.configure(foreground="#000000")
        self.Label5.configure(text='''Polling for Team:''')
        self.Label5.configure(width=117)

        self.TCombobox3 = ttk.Combobox(top)
        self.TCombobox3.place(relx=0.43, rely=0.41, relheight=0.05, relwidth=0.40)
        self.value_list = self.get_value()
        self.TCombobox3.configure(values=self.value_list)
        self.TCombobox3.bind('<<ComboboxSelected>>', self.get_combo3_value)
        self.TCombobox3.configure(textvariable=v3)
        self.TCombobox3.configure(takefocus="")

        self.Text1 = Text(top)
        self.Text1.place(relx=0.43, rely=0.53, relheight=0.33, relwidth=0.45)
        self.Text1.configure(background="white")
        self.Text1.configure(font=self.font3)
        self.Text1.configure(foreground="black")
        self.Text1.configure(highlightbackground="#d9d9d9")
        self.Text1.configure(highlightcolor="black")
        self.Text1.configure(insertbackground="black")
        self.Text1.configure(selectbackground="#c4c4c4")
        self.Text1.configure(selectforeground="black")
        self.Text1.configure(width=244)
        self.Text1.configure(wrap=WORD)

    def get_value(self):
        from mongo_operation import mongo_oper
        MDB = mongo_oper()
        aaa = MDB.get_distinct_team()
        aaa.sort()
        return aaa

    def get_combo2_value(self, event):
        self.combo2 = v2.get() if v2.get() is not "" else v2.get()

    def get_combo3_value(self, event):
        self.combo3 = v3.get() if v3.get() is not "" else v3.get()

    def click_btn_create_polling(self):
        try:
            from mongo_operation import mongo_oper
            MDB = mongo_oper()
            result = tkMessageBox.askokcancel("Confirmation","Do you really want to create the polling?")
            if result is True:
                if self.combo2 == "POLLING1":
                    MDB.misc_connect.update_one({"FIELD": "POLLING1"}, {"$set": {"POLLING_NAME": str(self.Entry1.get()),
                                                                        "VALUE": str(self.Text1.get("1.0", "end-1c"))}})
                    MDB.em_connect.update({"$and": [{"WO_NAME": str(self.combo3)},
                                                    {"RECORD_ACTIVE_INDICATOR": "Y"}]}, {"$set": {"POLLING1": "TBD"}})
                else:
                    MDB.misc_connect.update_one({"FIELD": "POLLING2"}, {"$set": {"POLLING_NAME": str(self.Entry1.get()),
                                                                        "VALUE": str(self.Text1.get("1.0", "end-1c"))}})
                    MDB.em_connect.update({"$and": [{"WO_NAME": str(self.combo3)},
                                                    {"RECORD_ACTIVE_INDICATOR": "Y"}]}, {"$set": {"POLLING2": "TBD"}})
                tkMessageBox.showinfo("Done","Polling created, The form will close now!!!")
                destroy_window()
            else:
                tkMessageBox.showinfo("Cancel","The polling was not created")
        except Exception, e1:
            print(str(e1))

if __name__ == '__main__':
    vp_start_gui()


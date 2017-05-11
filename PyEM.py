from Tkinter import *
import tkMessageBox


def vp_start_gui():
    from mongo_operation import mongo_oper
    global val, w, root, MDB
    root = Tk()
    top = Welcome_to_PyEM_(root)
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

def click_btn_misc_selection():
    print "YET TO BE CODED!!!"

def click_btn_analysis_selection():
    try:
        import Analysis_Selection
        destroy_window()
        Analysis_Selection.vp_start_gui()
    except Exception, e1:
        tkMessageBox.showerror("Error", str(e1))

def click_btn_refresh_data():
    try:
        from mongo_operation import mongo_oper
        import Analysis_Selection
        MDB.insert_em_rows_into_mongo(MDB.replace_non_em_fields_from_old_record(
            MDB.compare_em_mongo_data(MDB.em_raw_data_to_mongo_data_mapping(MDB.read_em_raw_data()))))
        tkMessageBox.showinfo("Success!!!", "Latest EM Data successfully refreshed to MongoDB")
        destroy_window()
        Analysis_Selection.vp_start_gui()
    except Exception, e1:
        tkMessageBox.showerror("Error", str(e1))

class Welcome_to_PyEM_:
    def __init__(self, top=None):
        self._bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        self._fgcolor = '#000000'  # X11 color: 'black'
        self._compcolor = '#d9d9d9' # X11 color: 'gray85'
        self._ana1color = '#d9d9d9' # X11 color: 'gray85'
        self._ana2color = '#d9d9d9' # X11 color: 'gray85'
        self.font3 = "-family Arial -size 9 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"
        self.font5 = "-family Arial -size 14 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"
        self.font7 = "-family Arial -size 9 -weight bold -slant roman "  \
            "-underline 0 -overstrike 0"

        top.geometry("624x290+499+142")
        top.title("Welcome to PyEM ")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")
        top.resizable(0, 0)



        self.Label1 = Label(top)
        self.Label1.place(relx=0.1, rely=0.07, height=71, width=497)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=self.font5)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(highlightbackground="#ff0000")
        self.Label1.configure(relief=RIDGE)
        self.Label1.configure(text='''Travelers Employee Management Tool - PyEM''')
        self.Label1.configure(width=497)

        self.Button1 = Button(top)
        self.Button1.place(relx=0.14, rely=0.45, height=64, width=106)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(font=self.font7)
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Refresh Data
and
Start the Tool''')
        self.Button1.configure(width=106)
        self.Button1.configure(command=click_btn_refresh_data)


        self.Button2 = Button(top)
        self.Button2.place(relx=0.43, rely=0.45, height=64, width=106)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(font=self.font7)
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''Analysis
with
Existing Data''')
        self.Button2.configure(width=106)
        self.Button2.configure(command=click_btn_analysis_selection)

        self.menubar = Menu(top, font=self.font3, bg=self._bgcolor, fg=self._fgcolor)
        top.configure(menu=self.menubar)

        self.Button3 = Button(top)
        self.Button3.place(relx=0.71, rely=0.45, height=64, width=106)
        self.Button3.configure(activebackground="#d9d9d9")
        self.Button3.configure(activeforeground="#000000")
        self.Button3.configure(background="#d9d9d9")
        self.Button3.configure(disabledforeground="#a3a3a3")
        self.Button3.configure(font=self.font7)
        self.Button3.configure(foreground="#000000")
        self.Button3.configure(highlightbackground="#d9d9d9")
        self.Button3.configure(highlightcolor="black")
        self.Button3.configure(pady="0")
        self.Button3.configure(text='''Miscelleneous
Operations''')
        self.Button3.configure(command=click_btn_misc_selection)


if __name__ == '__main__':
    vp_start_gui()



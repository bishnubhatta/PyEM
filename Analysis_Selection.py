from Tkinter import *
import tkMessageBox


def vp_start_gui():
    from mongo_operation import mongo_oper
    global val, w, root, MDB
    root = Tk()
    top = Analysis_Window(root)
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

def click_emp_btn():
    try:
        import AnalysisEmployee
        destroy_window()
        AnalysisEmployee.vp_start_gui()
    except Exception, e1:
        tkMessageBox.showerror("Error", str(e1))

def click_wo_btn():
    try:
        import AnalysisWO
        destroy_window()
        AnalysisWO.vp_start_gui()
    except Exception, e1:
        tkMessageBox.showerror("Error", str(e1))

class Analysis_Window:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        self._bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        self._fgcolor = '#000000'  # X11 color: 'black'
        self._compcolor = '#d9d9d9' # X11 color: 'gray85'
        self._ana1color = '#d9d9d9' # X11 color: 'gray85'
        self._ana2color = '#d9d9d9' # X11 color: 'gray85'
        self.font5 = "-family Arial -size 12 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"
        self.font6 = "-family Arial -size 10 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"

        top.geometry("585x269+517+135")
        top.title("Analysis Window")
        top.configure(background="#d9d9d9")
        top.resizable(0, 0)


        self.Button1 = Button(top)
        self.Button1.place(relx=0.14, rely=0.41, height=84, width=156)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(font=self.font6)
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Analysis for WO''')
        self.Button1.configure(width=156)
        self.Button1.configure(command=click_wo_btn)

        self.Button2 = Button(top)
        self.Button2.place(relx=0.55, rely=0.41, height=84, width=166)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(font=self.font6)
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''Analysis for Employee''')
        self.Button2.configure(width=166)
        self.Button2.configure(command=click_emp_btn)

        self.Label1 = Label(top)
        self.Label1.place(relx=0.29, rely=0.11, height=21, width=247)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=self.font5)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(relief=RIDGE)
        self.Label1.configure(text='''Analyze Data''')
        self.Label1.configure(width=247)

if __name__ == '__main__':
    vp_start_gui()




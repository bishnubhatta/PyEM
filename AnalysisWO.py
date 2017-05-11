from Tkinter import *
import ttk
import tkMessageBox


def vp_start_gui():
    from mongo_operation import mongo_oper
    global val, w, root, MDB, v1
    root = Tk()
    v1= StringVar()
    top = Analysis_by_WorkOrder(root)
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
    #import Analysis_Selection
    global top_level
    top_level.destroy()
    top_level = None
    #Analysis_Selection.vp_start_gui()



class Analysis_by_WorkOrder:
    def __init__(self, top=None):
        self.combo1=''
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
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
        self.font6 = "-family Arial -size 13 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"
        # self.style = ttk.Style()
        # if sys.platform == "win32":
        #     self.style.theme_use('winnative')
        # self.style.configure('.',background=self._bgcolor)
        # self.style.configure('.',foreground=self._fgcolor)
        # self.style.configure('.',font=self.font3)
        # self.style.map('.',background=
        #     [('selected', self._compcolor), ('active',self._ana2color)])

        top.geometry("546x340+519+123")
        top.title("Analysis by WorkOrder")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")
        top.resizable(0,0)


        self.Label1 = Label(top)
        self.Label1.place(relx=0.26, rely=0.06, height=41, width=257)
        self.Label1.configure(activebackground="#f9f9f9")
        self.Label1.configure(activeforeground="black")
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=self.font5)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(highlightbackground="#d9d9d9")
        self.Label1.configure(highlightcolor="black")
        self.Label1.configure(relief=RIDGE)
        self.Label1.configure(text='''Analysis By WorkOrder''')

        self.Label2 = Label(top)
        self.Label2.place(relx=0.04, rely=0.26, height=21, width=147)
        self.Label2.configure(activebackground="#f9f9f9")
        self.Label2.configure(activeforeground="black")
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font=self.font5)
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(highlightbackground="#d9d9d9")
        self.Label2.configure(highlightcolor="black")
        self.Label2.configure(text='''WorkOrder :''')

        self.TCombobox1 = ttk.Combobox(top)
        self.TCombobox1.place(relx=0.33, rely=0.26, relheight=0.06
                , relwidth=0.41)
        self.TCombobox1.configure(textvariable=v1)
        self.TCombobox1.configure(takefocus="")
        self.value_list = self.get_value()
        self.TCombobox1.configure(values=self.value_list)
        self.TCombobox1.bind('<<ComboboxSelected>>', self.get_combo_value)


        self.Button1 = Button(top)
        self.Button1.place(relx=0.05, rely=0.65, height=24, width=101)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(font=self.font3)
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Buffer Report''')
        self.Button1.configure(width=101)
        self.Button1.configure(command=self.buffer_report)

        self.Button5 = Button(top)
        self.Button5.place(relx=0.05, rely=0.79, height=24, width=106)
        self.Button5.configure(activebackground="#d9d9d9")
        self.Button5.configure(activeforeground="#000000")
        self.Button5.configure(background="#d9d9d9")
        self.Button5.configure(disabledforeground="#a3a3a3")
        self.Button5.configure(font=self.font3)
        self.Button5.configure(foreground="#000000")
        self.Button5.configure(highlightbackground="#d9d9d9")
        self.Button5.configure(highlightcolor="black")
        self.Button5.configure(pady="0")
        self.Button5.configure(text='''Employee Report''')
        self.Button5.configure(command=self.employee_report)

        self.Label3 = Label(top)
        self.Label3.place(relx=0.49, rely=0.35, height=31, width=27)
        self.Label3.configure(activebackground="#f9f9f9")
        self.Label3.configure(activeforeground="black")
        self.Label3.configure(background="#d9d9d9")
        self.Label3.configure(disabledforeground="#a3a3a3")
        self.Label3.configure(font=self.font6)
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(highlightbackground="#d9d9d9")
        self.Label3.configure(highlightcolor="black")
        self.Label3.configure(text='''or''')

        self.Label4 = Label(top)
        self.Label4.place(relx=0.04, rely=0.47, height=31, width=207)
        self.Label4.configure(activebackground="#f9f9f9")
        self.Label4.configure(activeforeground="black")
        self.Label4.configure(background="#d9d9d9")
        self.Label4.configure(disabledforeground="#a3a3a3")
        self.Label4.configure(font=self.font5)
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(highlightbackground="#d9d9d9")
        self.Label4.configure(highlightcolor="black")
        self.Label4.configure(text='''WorkOrder List (, Separated):''')
        self.Label4.configure(width=207)

        self.Entry1 = Entry(top)
        self.Entry1.place(relx=0.48, rely=0.5, relheight=0.06, relwidth=0.39)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font=self.font4)
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(highlightbackground="#d9d9d9")
        self.Entry1.configure(highlightcolor="black")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(selectbackground="#c4c4c4")
        self.Entry1.configure(selectforeground="black")

        self.Button6 = Button(top)
        self.Button6.place(relx=0.27, rely=0.65, height=24, width=101)
        self.Button6.configure(activebackground="#d9d9d9")
        self.Button6.configure(activeforeground="#000000")
        self.Button6.configure(background="#d9d9d9")
        self.Button6.configure(disabledforeground="#a3a3a3")
        self.Button6.configure(font=self.font3)
        self.Button6.configure(foreground="#000000")
        self.Button6.configure(highlightbackground="#d9d9d9")
        self.Button6.configure(highlightcolor="black")
        self.Button6.configure(pady="0")
        self.Button6.configure(text='''Bench Report''')
        self.Button6.configure(command=self.bench_report)

        self.Button7 = Button(top)
        self.Button7.place(relx=0.49, rely=0.65, height=24, width=101)
        self.Button7.configure(activebackground="#d9d9d9")
        self.Button7.configure(activeforeground="#000000")
        self.Button7.configure(background="#d9d9d9")
        self.Button7.configure(disabledforeground="#a3a3a3")
        self.Button7.configure(font=self.font3)
        self.Button7.configure(foreground="#000000")
        self.Button7.configure(highlightbackground="#d9d9d9")
        self.Button7.configure(highlightcolor="black")
        self.Button7.configure(pady="0")
        self.Button7.configure(text='''RRD Report''')
        self.Button7.configure(command = self.rrd_report)

        self.Button8 = Button(top)
        self.Button8.place(relx=0.71, rely=0.65, height=24, width=101)
        self.Button8.configure(activebackground="#d9d9d9")
        self.Button8.configure(activeforeground="#000000")
        self.Button8.configure(background="#d9d9d9")
        self.Button8.configure(disabledforeground="#a3a3a3")
        self.Button8.configure(font=self.font3)
        self.Button8.configure(foreground="#000000")
        self.Button8.configure(highlightbackground="#d9d9d9")
        self.Button8.configure(highlightcolor="black")
        self.Button8.configure(pady="0")
        self.Button8.configure(text='''CI Report''')
        self.Button8.configure(command=self.calc_ci_wo)

        self.Button9 = Button(top)
        self.Button9.place(relx=0.29, rely=0.79, height=24, width=106)
        self.Button9.configure(activebackground="#d9d9d9")
        self.Button9.configure(activeforeground="#000000")
        self.Button9.configure(background="#d9d9d9")
        self.Button9.configure(disabledforeground="#a3a3a3")
        self.Button9.configure(font=self.font3)
        self.Button9.configure(foreground="#000000")
        self.Button9.configure(highlightbackground="#d9d9d9")
        self.Button9.configure(highlightcolor="black")
        self.Button9.configure(pady="0")
        self.Button9.configure(text='''Pyramid Report''')
        self.Button9.configure(command=self.calc_pyramid_wo)

        self.Button2 = Button(top)
        self.Button2.place(relx=0.53, rely=0.79, height=24, width=106)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(font=self.font3)
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''WhatIf Analysis''')
        self.Button2.configure(width=86)
        self.Button2.configure(command=self.wif_ci_analysis)

        self.Button3 = Button(top)
        self.Button3.place(relx=0.75, rely=0.79, height=24, width=106)
        self.Button3.configure(activebackground="#d9d9d9")
        self.Button3.configure(activeforeground="#000000")
        self.Button3.configure(background="#d9d9d9")
        self.Button3.configure(disabledforeground="#a3a3a3")
        self.Button3.configure(font=self.font3)
        self.Button3.configure(foreground="#000000")
        self.Button3.configure(highlightbackground="#d9d9d9")
        self.Button3.configure(highlightcolor="black")
        self.Button3.configure(pady="0")
        self.Button3.configure(text='''Go Back''')
        self.Button3.configure(width=86)
        self.Button3.configure(command=self.go_back)

    def get_value(self):
        from mongo_operation import mongo_oper
        MDB = mongo_oper()
        aaa =  MDB.get_distinct_wo()
        aaa.sort()
        return aaa

    def get_combo_value(self, event):
        self.combo1 = v1.get()

    def get_wo_for_entered_data(self):
        wo = "" if self.Entry1.get() is None else self.Entry1.get()
        wo = wo if self.combo1 == "" else self.combo1
        return wo

    def go_back(self):
        try:
            import Analysis_Selection
            destroy_window()
            Analysis_Selection.vp_start_gui()
        except Exception, e1:
            print(str(e1))

    def wif_ci_analysis(self):
        try:
            import WhatIfCIAnalysis
            destroy_window()
            WhatIfCIAnalysis.vp_start_gui()
        except Exception, e1:
            print(str(e1))

    def calc_pyramid_wo(self):
        wo = self.get_wo_for_entered_data()
        result = MDB.fetch_pyramid_report_by_wo(wo)
        tkMessageBox.showinfo("Pyramid Report", "Report Generated Successfully!!!")

    def calc_ci_wo(self):
        wo = self.get_wo_for_entered_data()
        result = MDB.calculate_ci_for_wo(wo)
        final_display = ''
        for data in result:
            final_display = "CI and Discounted CI for WO " + str(data[0]) + ": " + str(data[1]) +  " and " + str(data[2]) + "\n" + final_display
        tkMessageBox.showinfo("CI for WO", final_display)

    def buffer_report(self):
        wo = self.get_wo_for_entered_data()
        result = MDB.fetch_buffer_report_by_wo(wo)
        tkMessageBox.showinfo("Buffer Report", "Report Generated Successfully!!!")

    def bench_report(self):
        wo = self.get_wo_for_entered_data()
        MDB.fetch_bench_report_by_wo(wo)
        tkMessageBox.showinfo("Bench Report", "Report Generated Successfully!!!")

    def rrd_report(self):
        wo = self.get_wo_for_entered_data()
        result = MDB.rrd_type_report_by_wo(wo)
        tkMessageBox.showinfo("RRD Type Report", "Report Generated Successfully!!!")

    def employee_report(self):
        wo = self.get_wo_for_entered_data()
        result = MDB.view_all_documents_for_WO(wo)
        tkMessageBox.showinfo("Active Employee Report", "Report Generated Successfully!!!")




if __name__ == '__main__':
    vp_start_gui()


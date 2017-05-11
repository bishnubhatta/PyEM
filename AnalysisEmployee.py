from Tkinter import *
import ttk
import tkMessageBox

def vp_start_gui():
    from mongo_operation import mongo_oper
    global val, w, root, MDB, v1,v2,v3
    root = Tk()
    v1 = StringVar()
    v2 = StringVar()
    v3 = StringVar()
    top = Analysis_by_Employee(root)
    MDB = mongo_oper()
    init(root, top)
    root.mainloop()

w = None

def init(top, gui, *args, **kwargs):
    global w, top_level, root, MDB
    w = gui
    top_level = top
    root = top


class Analysis_by_Employee:
    def __init__(self, top=None):
        self.combo1 = ''
        self.textval1 = ''
        self.textval2 = ''

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
        self.font7 = "-family Arial -size 13 -weight bold -slant roman"  \
            " -underline 0 -overstrike 0"

        top.geometry("559x340+519+123")
        top.title("Analysis by Employee")
        top.configure(background="#d9d9d9")
        top.resizable(0, 0)

        self.Label1 = Label(top)
        self.Label1.place(relx=0.25, rely=0.06, height=41, width=257)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=self.font5)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(relief=RIDGE)
        self.Label1.configure(text='''Analysis for Employee''')
        self.Label1.configure(width=257)

        self.Label2 = Label(top)
        self.Label2.place(relx=0.04, rely=0.26, height=21, width=147)
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font=self.font5)
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(text='''Enterprise ID :''')
        self.Label2.configure(width=147)

        self.TCombobox1 = ttk.Combobox(top)
        self.TCombobox1.place(relx=0.32, rely=0.26, relheight=0.06, relwidth=0.4)
        self.TCombobox1.configure(textvariable = v1)
        self.TCombobox1.configure(width=223)
        self.TCombobox1.configure(takefocus="")
        self.value_list = self.get_value()
        #tkMessageBox.showinfo("Values",str(self.value_list))
        self.TCombobox1.configure(values=self.value_list)
        self.TCombobox1.bind('<<ComboboxSelected>>', self.get_combo_value)

        self.Button1 = Button(top)
        self.Button1.place(relx=0.04, rely=0.82, height=24, width=71)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(font=self.font3)
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Show Role''')
        self.Button1.configure(command=self.emp_show_role)

        self.Button2 = Button(top)
        self.Button2.place(relx=0.21, rely=0.82, height=24, width=70)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(font=self.font3)
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''Show LCR''')
        self.Button2.configure(command=self.emp_show_lcr)

        self.Button3 = Button(top)
        self.Button3.place(relx=0.39, rely=0.82, height=24, width=67)
        self.Button3.configure(activebackground="#d9d9d9")
        self.Button3.configure(activeforeground="#000000")
        self.Button3.configure(background="#d9d9d9")
        self.Button3.configure(disabledforeground="#a3a3a3")
        self.Button3.configure(font=self.font3)
        self.Button3.configure(foreground="#000000")
        self.Button3.configure(highlightbackground="#d9d9d9")
        self.Button3.configure(highlightcolor="black")
        self.Button3.configure(pady="0")
        self.Button3.configure(text='''Show CI''')
        self.Button3.configure(width=67)
        self.Button3.configure(command=self.emp_show_ci)

        self.Button4 = Button(top)
        self.Button4.place(relx=0.57, rely=0.82, height=24, width=91)
        self.Button4.configure(activebackground="#d9d9d9")
        self.Button4.configure(activeforeground="#000000")
        self.Button4.configure(background="#d9d9d9")
        self.Button4.configure(disabledforeground="#a3a3a3")
        self.Button4.configure(font=self.font3)
        self.Button4.configure(foreground="#000000")
        self.Button4.configure(highlightbackground="#d9d9d9")
        self.Button4.configure(highlightcolor="black")
        self.Button4.configure(pady="0")
        self.Button4.configure(text='''Show Bill Rate''')
        self.Button4.configure(command=self.emp_show_billrt)

        self.Button5 = Button(top)
        self.Button5.place(relx=0.79, rely=0.82, height=24, width=106)
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
        self.Button5.configure(command=self.emp_gen_rpt)

        self.Label3 = Label(top)
        self.Label3.place(relx=0.48, rely=0.38, height=31, width=27)
        self.Label3.configure(background="#d9d9d9")
        self.Label3.configure(disabledforeground="#a3a3a3")
        self.Label3.configure(font=self.font7)
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(text='''or''')
        self.Label3.configure(width=27)

        self.Label4 = Label(top)
        self.Label4.place(relx=0.07, rely=0.5, height=31, width=97)
        self.Label4.configure(background="#d9d9d9")
        self.Label4.configure(disabledforeground="#a3a3a3")
        self.Label4.configure(font=self.font5)
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(text='''First Name :''')
        self.Label4.configure(width=97)

        self.Entry1 = Entry(top)
        self.Entry1.place(relx=0.32, rely=0.53, relheight=0.06, relwidth=0.38)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font=self.font4)
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(width=214)

        self.Label5 = Label(top)
        self.Label5.place(relx=0.07, rely=0.65, height=24, width=100)
        self.Label5.configure(background="#d9d9d9")
        self.Label5.configure(disabledforeground="#a3a3a3")
        self.Label5.configure(font=self.font5)
        self.Label5.configure(foreground="#000000")
        self.Label5.configure(text='''Last Name :''')
        self.Label5.configure(width=100)

        self.Entry2 = Entry(top)
        self.Entry2.place(relx=0.32, rely=0.65, relheight=0.06, relwidth=0.38)
        self.Entry2.configure(background="white")
        self.Entry2.configure(disabledforeground="#a3a3a3")
        self.Entry2.configure(font=self.font4)
        self.Entry2.configure(foreground="#000000")
        self.Entry2.configure(insertbackground="black")
        self.Entry2.configure(width=214)

    def get_value(self):
        from mongo_operation import mongo_oper
        MDB = mongo_oper()
        r = []
        aaa= MDB.get_distinct_ent_id()
        for data in aaa:
            r.append(data)
        r.sort()
        return r

    def get_combo_value(self,event):
        self.combo1 = v1.get()
        #tkMessageBox.showinfo("Value", self.combo1)

    def get_ent_id_for_entered_data(self):
        fname = "" if self.Entry1.get() is None else self.Entry1.get()
        lname = "" if self.Entry2.get() is None else self.Entry2.get()
        ent_id = MDB.fname_lname_to_entid_mapping(fname,lname) if self.combo1 == "" else self.combo1
        return ent_id


    def emp_show_role(self):
        ent_id = self.get_ent_id_for_entered_data()
        role = MDB.fetch_role_for_entid(ent_id)
        tkMessageBox.showinfo("Role",str(role))

    def emp_show_lcr(self):
        ent_id = self.get_ent_id_for_entered_data()
        #print "ENT ID:" + str(ent_id)
        lcr = MDB.fetch_lcr_for_employeeid(MDB.fetch_gcpid_for_entid(ent_id))
        tkMessageBox.showinfo("LCR", str(lcr))

    def emp_show_billrt(self):
        ent_id = self.get_ent_id_for_entered_data()
        billrt = MDB.fetch_bill_rate(MDB.fetch_role_for_entid(ent_id),MDB.fetch_location_for_entid(ent_id))
        data = MDB.misc_connect.find_one({"FIELD": "BILLING_DISCOUNT"}, {"_id": 0})
        dbillrt = float (billrt * (1 - float(data["VALUE"])))
        tkMessageBox.showinfo("Bill Rate", "Bill Rate : " + str(billrt) + "\n" + "Discounted Bill Rate: " + str(dbillrt))

    def emp_gen_rpt(self):
        ent_id = self.get_ent_id_for_entered_data()
        MDB.view_all_documents_for_entid(ent_id)
        tkMessageBox.showinfo("Employee Report",
                              "Report Generated Successfully!!!")

    def emp_show_ci(self):
        ent_id = self.get_ent_id_for_entered_data()
        tci = MDB.calculate_ci_for_employee(ent_id)
        #print tci
        if tci[0] == 0 or tci[1] == 0:
            tkMessageBox.showinfo("Error","Either LCR data not present or a Buffer :  " + str(ent_id))
        else:
            tkMessageBox.showinfo("CI and Discounted CI",
                              "Original CI : " + str(tci[0]) + "\n" + "Discounted CI: " + str(tci[1]))

if __name__ == '__main__':
    vp_start_gui()


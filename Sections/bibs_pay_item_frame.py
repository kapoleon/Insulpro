import customtkinter as ctk

ctk.set_appearance_mode("system")  # system, dark , light
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class Employee:
    def __init__(self, first, last, vacationdays, passwd):
        self.first = first
        self.last = last
        self.vacationdays = vacationdays
        self.fullname = first + ' ' + last
        self.pwd = passwd
        self.username = first[0] + last


employee0 = Employee('Admin', 'Account', 0, '')
employee1 = Employee("Justin", "Bausinger", 5, '')
employee2 = Employee("Cody", "Billings", 0, '')
employee3 = Employee("Jeremy", "Denton", 5, '')
employee4 = Employee("Glen", "Hayes", 10, '')
employee5 = Employee("Brian", "Maikranz", 5, '')
employee6 = Employee("Mike", "Mato", 5, '')
employee7 = Employee("Kenny", "Ray", 10, '')
employee8 = Employee("Edgar", "Sanchez", 5, '')
employee9 = Employee("Chris", "Waldridge", 10, '')
employee10 = Employee("New", "Employee", 0, '')

class bibs_pay_item_frame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        bibsTitleLabel = ctk.CTkLabel(self,
                                      text="BIBS",
                                      width=200,
                                      height=30,
                                      fg_color="transparent",
                                      font=ctk.CTkFont(size=18, weight="normal"))

        bibsTitleLabel.place(x=450, y=30)

        bibsFullInstallLabel = ctk.CTkLabel(self,
                                            text="Full Install",
                                            width=100,
                                            height=30,
                                            fg_color="transparent",
                                            font=ctk.CTkFont(size=18, weight="normal"))

        bibsFullInstallLabel.place(x=450, y=65)

        self.bibsFullInstall = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.bibsFullInstall.place(x=565, y=65)
        self.bibsFullInstall.insert(0, '0')

        bibsHangInstallLabel = ctk.CTkLabel(self,
                                            text="Hang Netting",
                                            width=100,
                                            height=30,
                                            fg_color="transparent",
                                            font=ctk.CTkFont(size=18, weight="normal"))

        bibsHangInstallLabel.place(x=450, y=100)

        self.bibsHangInstall = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.bibsHangInstall.place(x=565, y=100)
        self.bibsHangInstall.insert(0, '0')

        bibsTackInstallLabel = ctk.CTkLabel(self,
                                            text="Tack Netting",
                                            width=100,
                                            height=30,
                                            fg_color="transparent",
                                            font=ctk.CTkFont(size=18, weight="normal"))

        bibsTackInstallLabel.place(x=450, y=135)

        self.bibsTackInstall = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.bibsTackInstall.place(x=565, y=135)
        self.bibsTackInstall.insert(0, '0')

        bibsBlowInstallLabel = ctk.CTkLabel(self,
                                            text="Walls Blown",
                                            width=100,
                                            height=30,
                                            fg_color="transparent",
                                            font=ctk.CTkFont(size=18, weight="normal"))

        bibsBlowInstallLabel.place(x=450, y=170)

        self.bibsBlowInstall = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.bibsBlowInstall.place(x=565, y=170)
        self.bibsBlowInstall.insert(0, '0')



class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x200")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = bibs_pay_item_frame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.mainloop()

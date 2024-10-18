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

class attic_pay_item_frame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        r19InstalledLabel = ctk.CTkLabel(self,
                                         text="R19 Blow",
                                         width=100,
                                         height=30,
                                         fg_color="transparent",
                                         font=ctk.CTkFont(size=18, weight="normal"))

        r19InstalledLabel.place(x=180, y=65)

        self.r19Installed = ctk.CTkEntry(self,
                                         width=100,
                                         height=30,
                                         font=ctk.CTkFont(size=18, weight="normal"))

        self.r19Installed.place(x=300, y=65, )
        self.r19Installed.insert(0, '0')

        r30InstalledLabel = ctk.CTkLabel(self,
                                         text="R30 Blow",
                                         width=100,
                                         height=30,
                                         fg_color="transparent",
                                         font=ctk.CTkFont(size=18, weight="normal"))

        r30InstalledLabel.place(x=180, y=100)

        self.r30Installed = ctk.CTkEntry(self,
                                         width=100,
                                         height=30,
                                         font=ctk.CTkFont(size=18, weight="normal"))

        self.r30Installed.place(x=300, y=100)
        self.r30Installed.insert(0, '0')

        r38InstalledLabel = ctk.CTkLabel(self,
                                         text="R38/40 Blow",
                                         width=100,
                                         height=30,
                                         fg_color="transparent",
                                         font=ctk.CTkFont(size=18, weight="normal"))

        r38InstalledLabel.place(x=180, y=135)

        self.r38Installed = ctk.CTkEntry(self,
                                         width=100,
                                         height=30,
                                         font=ctk.CTkFont(size=18, weight="normal"))

        self.r38Installed.place(x=300, y=135)
        self.r38Installed.insert(0, '0')

        r49InstalledLabel = ctk.CTkLabel(self,
                                         text="R49/50 Blow",
                                         width=100,
                                         height=30,
                                         fg_color="transparent",
                                         font=ctk.CTkFont(size=18, weight="normal"))

        r49InstalledLabel.place(x=180, y=170)

        self.r49Installed = ctk.CTkEntry(self,
                                         width=100,
                                         height=30,
                                         font=ctk.CTkFont(size=18, weight="normal"))

        self.r49Installed.place(x=300, y=170)
        self.r49Installed.insert(0, '0')

        celluloseInstalledLabel = ctk.CTkLabel(self,
                                               text="Cellulose Blow",
                                               width=100,
                                               height=30,
                                               fg_color="transparent",
                                               font=ctk.CTkFont(size=18, weight="normal"))

        celluloseInstalledLabel.place(x=180, y=205)

        self.celluloseInstalled = ctk.CTkEntry(self,
                                               width=100,
                                               height=30,
                                               font=ctk.CTkFont(size=18, weight="normal"))

        self.celluloseInstalled.place(x=300, y=205)
        self.celluloseInstalled.insert(0, '0')

        soffitInstalledLabel = ctk.CTkLabel(self,
                                            text="Soffits-Existing",
                                            width=100,
                                            height=30,
                                            fg_color="transparent",
                                            font=ctk.CTkFont(size=18, weight="normal"))

        soffitInstalledLabel.place(x=180, y=240)

        self.soffitInstalled = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.soffitInstalled.place(x=300, y=240)
        self.soffitInstalled.insert(0, '0')

        airSealLabel = ctk.CTkLabel(self,
                                    text="Air Seal",
                                    width=100,
                                    height=30,
                                    fg_color="transparent",
                                    font=ctk.CTkFont(size=18, weight="normal"))

        airSealLabel.place(x=180, y=275)

        self.airSeal = ctk.CTkEntry(self,
                                    width=100,
                                    height=30,
                                    font=ctk.CTkFont(size=18, weight="normal"))

        self.airSeal.place(x=300, y=275)
        self.airSeal.insert(0, '0')

        bonusAmountLabel = ctk.CTkLabel(self,
                                        text="Bonus",
                                        width=100,
                                        height=30,
                                        fg_color="transparent",
                                        font=ctk.CTkFont(size=18, weight="normal"))

        bonusAmountLabel.place(x=180, y=310)

        self.bonusAmount = ctk.CTkEntry(self,
                                        width=100,
                                        height=30,
                                        font=ctk.CTkFont(size=18, weight="normal"))

        self.bonusAmount.place(x=300, y=310)
        self.bonusAmount.insert(0, '0')

        otherAmountLabel = ctk.CTkLabel(self,
                                        text="Other Amount",
                                        width=100,
                                        height=30,
                                        fg_color="transparent",
                                        font=ctk.CTkFont(size=18, weight="normal"))

        otherAmountLabel.place(x=180, y=345)

        self.otherAmount = ctk.CTkEntry(self,
                                        width=100,
                                        height=30,
                                        font=ctk.CTkFont(size=18, weight="normal"))

        self.otherAmount.place(x=300, y=345)
        self.otherAmount.insert(0, '0')




class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x200")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = attic_pay_item_frame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.mainloop()

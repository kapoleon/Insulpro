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

class batt_pay_item_frame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)



        battUnder10Label = ctk.CTkLabel(self,
                                        text="Batt > 10ft",
                                        width=100,
                                        height=30,
                                        fg_color="transparent",
                                        font=ctk.CTkFont(size=18, weight="normal"))
        battUnder10Label.place(x=180, y=65)

        self.battUnder10 = ctk.CTkEntry(self,
                                        width=100,
                                        height=30,
                                        font=ctk.CTkFont(size=18, weight="normal"))

        self.battUnder10.place(x=300, y=65)
        self.battUnder10.insert(0, '0')

        battOver10Label = ctk.CTkLabel(self,
                                       text="Batt < 10ft",
                                       width=100,
                                       height=30,
                                       fg_color="transparent",
                                       font=ctk.CTkFont(size=18, weight="normal"))

        battOver10Label.place(x=180, y=100)

        self.battOver10 = ctk.CTkEntry(self,
                                       width=100,
                                       height=30,
                                       font=ctk.CTkFont(size=18, weight="normal"))
        self.battOver10.place(x=300, y=100)
        self.battOver10.insert(0, '0')

        newSoffitLabel = ctk.CTkLabel(self,
                                      text="Soffit-New",
                                      width=100,
                                      height=30,
                                      fg_color="transparent",
                                      font=ctk.CTkFont(size=18, weight="normal"))

        newSoffitLabel.place(x=180, y=135)

        self.newSoffit = ctk.CTkEntry(self,
                                      width=100,
                                      height=30,
                                      font=ctk.CTkFont(size=18, weight="normal"))

        self.newSoffit.place(x=300, y=135)
        self.newSoffit.insert(0, '0')

        caulkFoamLabel = ctk.CTkLabel(self,
                                      text="Caulk & Foam",
                                      width=100,
                                      height=30,
                                      fg_color="transparent",
                                      font=ctk.CTkFont(size=18, weight="normal"))

        caulkFoamLabel.place(x=180, y=170)

        self.caulkFoam = ctk.CTkEntry(self,
                                      width=100,
                                      height=30,
                                      font=ctk.CTkFont(size=18, weight="normal"))

        self.caulkFoam.place(x=300, y=170)
        self.caulkFoam.insert(0, '0')

        bonusLabel = ctk.CTkLabel(self,
                                  text="Bonus",
                                  width=100,
                                  height=30,
                                  fg_color="transparent",
                                  font=ctk.CTkFont(size=18, weight="normal"))

        bonusLabel.place(x=180, y=205)

        self.bonusAmount = ctk.CTkEntry(self,
                                        width=100,
                                        height=30,
                                        font=ctk.CTkFont(size=18, weight="normal"))

        self.bonusAmount.place(x=300, y=205)
        self.bonusAmount.insert(0, '0')

        celluloseLabel = ctk.CTkLabel(self,
                                      text="Cellulose",
                                      width=100,
                                      height=30,
                                      fg_color="transparent",
                                      font=ctk.CTkFont(size=18, weight="normal"))

        celluloseLabel.place(x=180, y=240)

        self.celluloseAmount = ctk.CTkEntry(self,
                                            width=100,
                                            height=30,
                                            font=ctk.CTkFont(size=18, weight="normal"))

        self.celluloseAmount.place(x=300, y=240)
        self.celluloseAmount.insert(0, '0')

        otherAmountLabel = ctk.CTkLabel(self,
                                        text="Other Amount",
                                        width=100,
                                        height=30,
                                        fg_color="transparent",
                                        font=ctk.CTkFont(size=18, weight="normal"))

        otherAmountLabel.place(x=180, y=275)

        self.otherAmount = ctk.CTkEntry(self,
                                        width=100,
                                        height=30,
                                        font=ctk.CTkFont(size=18, weight="normal"))

        self.otherAmount.place(x=300, y=275)
        self.otherAmount.insert(0, '0')


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x200")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = batt_pay_item_frame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.mainloop()

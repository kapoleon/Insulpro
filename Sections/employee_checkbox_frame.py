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

class employee_checkbox_frame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.count = 0

        self.emp1Checkbox_var = ctk.IntVar(value=0)
        self.emp2Checkbox_var = ctk.IntVar(value=0)
        self.emp3Checkbox_var = ctk.IntVar(value=0)
        self.emp4Checkbox_var = ctk.IntVar(value=0)
        self.emp5Checkbox_var = ctk.IntVar(value=0)
        self.emp6Checkbox_var = ctk.IntVar(value=0)
        self.emp7Checkbox_var = ctk.IntVar(value=0)
        self.emp8Checkbox_var = ctk.IntVar(value=0)
        self.emp9Checkbox_var = ctk.IntVar(value=0)
        self.emp10Checkbox_var = ctk.IntVar(value=0)

        emp1Checkbox = ctk.CTkCheckBox(self,
                                       text=employee1.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp1Checkbox_command,
                                       variable=self.emp1Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp1Checkbox.place(x=5, y=30)

        emp2Checkbox = ctk.CTkCheckBox(self,
                                       text=employee2.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp2Checkbox_command,
                                       variable=self.emp2Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp2Checkbox.place(x=5, y=65)

        emp3Checkbox = ctk.CTkCheckBox(self,
                                       text=employee3.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp3Checkbox_command,
                                       variable=self.emp3Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp3Checkbox.place(x=5, y=100)

        emp4Checkbox = ctk.CTkCheckBox(self,
                                       text=employee4.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp4Checkbox_command,
                                       variable=self.emp4Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp4Checkbox.place(x=5, y=135)

        emp5Checkbox = ctk.CTkCheckBox(self,
                                       text=employee5.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp5Checkbox_command,
                                       variable=self.emp5Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp5Checkbox.place(x=5, y=170)

        emp6Checkbox = ctk.CTkCheckBox(self,
                                       text=employee6.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp6Checkbox_command,
                                       variable=self.emp6Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp6Checkbox.place(x=5, y=205)

        emp7Checkbox = ctk.CTkCheckBox(self,
                                       text=employee7.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp7Checkbox_command,
                                       variable=self.emp7Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp7Checkbox.place(x=5, y=240)

        emp8Checkbox = ctk.CTkCheckBox(self,
                                       text=employee8.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp8Checkbox_command,
                                       variable=self.emp8Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp8Checkbox.place(x=5, y=275)

        emp9Checkbox = ctk.CTkCheckBox(self,
                                       text=employee9.fullname,
                                       font=ctk.CTkFont(size=18, weight="normal"),
                                       command=self.emp9Checkbox_command,
                                       variable=self.emp9Checkbox_var,
                                       onvalue=1,
                                       offvalue=0)
        emp9Checkbox.place(x=5, y=310)

        emp10Checkbox = ctk.CTkCheckBox(self,
                                        text=employee10.fullname,
                                        font=ctk.CTkFont(size=18, weight="normal"),
                                        command=self.emp10Checkbox_command,
                                        variable=self.emp10Checkbox_var,
                                        onvalue=1,
                                        offvalue=0)
        emp10Checkbox.place(x=5, y=345)

    def emp1Checkbox_command(self):
        print(employee1.fullname + " selected")
        self.emp1Checkbox_var.set(1)
        self.count += 1

    def emp2Checkbox_command(self):
        print(employee2.fullname + " selected")
        self.emp2Checkbox_var.set(1)
        self.count += 1

    def emp3Checkbox_command(self):
        print(employee3.fullname + " selected")
        self.emp3Checkbox_var.set(1)
        self.count += 1

    def emp4Checkbox_command(self):
        print(employee4.fullname + " selected")
        self.emp4Checkbox_var.set(1)
        self.count += 1

    def emp5Checkbox_command(self):
        print(employee5.fullname + " selected")
        self.emp5Checkbox_var.set(1)
        self.count += 1

    def emp6Checkbox_command(self):
        print(employee6.fullname + " selected")
        self.emp6Checkbox_var.set(1)
        self.count += 1

    def emp7Checkbox_command(self):
        print(employee7.fullname + " selected")
        self.emp7Checkbox_var.set(1)
        self.count += 1

    def emp8Checkbox_command(self):
        print(employee8.fullname + " selected")
        self.emp8Checkbox_var.set(1)
        self.count += 1

    def emp9Checkbox_command(self):
        print(employee9.fullname + " selected")
        self.emp9Checkbox_var.set(1)
        self.count += 1

    def emp10Checkbox_command(self):
        print(employee10.fullname + " selected")
        self.emp10Checkbox_var.set(1)
        self.count += 1


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x200")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = employee_checkbox_frame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.mainloop()

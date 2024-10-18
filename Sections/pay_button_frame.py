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

class pay_button_frame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        calculateButton = ctk.CTkButton(self,
                                        text="Calculate",
                                        width=150,
                                        height=30,
                                        font=ctk.CTkFont(size=17)
                                        )

        calculateButton.place(x=10, y=400)

        resetFormButton = ctk.CTkButton(self,
                                        text="Reset Form",
                                        width=150,
                                        height=30,
                                        font=ctk.CTkFont(size=17))

        resetFormButton.place(x=10, y=435)

        saveFileButton = ctk.CTkButton(self,
                                       text="Save File",
                                       width=150,
                                       height=30,
                                       font=ctk.CTkFont(size=17))

        saveFileButton.place(x=10, y=470)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x200")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = pay_button_frame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.mainloop()

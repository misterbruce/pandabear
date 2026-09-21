import customtkinter
import os
import openpyxl
from tkinter import filedialog

def check_path(path):
    if os.path.exists(path):
        return True
    else:
        return False

def conversion_engine(): #handles the actual conversion for the xlsx_conversion function
    member: str
    work_email: str
    group = "CSR" 
    start_date: str 
    start_time:str 
    end_date: str 
    end_time: str
    theme_color: str 
    cutom_label = None
    unpaid_break = None 
    notes = None
    shared = None

    dictionary_conversion = {
    member: [work_email, group, start_date, start_time, end_date, end_time, theme_color, cutom_label, unpaid_break, notes, shared]
    }




def xlsx_convert(path): #converts the selected schedule into a format that can be imported into Microsoft Teams/Shifts
    if check_path(path):
        conversion_engine(path)
    else:
        return ReferenceError



class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Tammy's Schedule Converter")
        self.geometry("300x300")
        self.resizable(False, False)
        self.grid_columnconfigure((0, 1), weight=1, uniform="equal")

        self.label = customtkinter.CTkLabel(self, text="No Schedule Selected.", wraplength=350)
        self.label.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        upload_buton = customtkinter.CTkButton(self, text="Upload Schedule", width=50, corner_radius=10, command=lambda: self.select_file())
        upload_buton.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        name_mapping = customtkinter.CTkButton(self, text="Employee Email", width=35, corner_radius=20, fg_color="grey")
        name_mapping.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel (xlsx) File",
            filetypes=[("Excel Files", "*.xlsx")],
        )


app = App()
app.mainloop()
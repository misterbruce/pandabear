import os
import json
import customtkinter
from tkinter import filedialog, messagebox
import openpyxl
import xlsx_engine

JSON_DB_PATH = os.path.join(os.path.dirname(__file__), "employee_emails.json")

class EmailEditorWindow(customtkinter.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Manage Employee Emails")
        self.geometry("450x500")
        self.resizable(False, False)
        
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.load_data()
        
        header_lbl = customtkinter.CTkLabel(self, text="Direct JSON Database Editor", font=customtkinter.CTkFont(size=16, weight="bold"))
        header_lbl.grid(row=0, column=0, padx=15, pady=15, sticky="w")
        
        self.scroll_frame = customtkinter.CTkScrollableFrame(self, label_text="Registered Personnel Profiles")
        self.scroll_frame.grid(row=1, column=0, padx=15, pady=10, sticky="nsew")
        self.scroll_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.entries_map = {}
        self.render_rows()
        
        actions_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        actions_frame.grid(row=2, column=0, padx=15, pady=15, sticky="ew")
        actions_frame.grid_columnconfigure((0, 1), weight=1)
        
        add_btn = customtkinter.CTkButton(actions_frame, text="+ Add Row", fg_color="#2caf73", hover_color="#228b5a", command=self.add_blank_row)
        add_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        save_btn = customtkinter.CTkButton(actions_frame, text="Save Updates", command=self.save_data)
        save_btn.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    def load_data(self):
        if os.path.exists(JSON_DB_PATH):
            try:
                with open(JSON_DB_PATH, "r", encoding="utf-8") as f:
                    self.db = json.load(f)
            except Exception:
                self.db = {}
        else:
            self.db = {"Alyssa": "", "Jack": "", "SK": ""}

    def render_rows(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
            
        self.entries_map.clear()
        
        for idx, (name, email) in enumerate(self.db.items()):
            name_input = customtkinter.CTkEntry(self.scroll_frame, placeholder_text="Name")
            name_input.insert(0, name)
            name_input.grid(row=idx, column=0, padx=5, pady=4, sticky="ew")
            
            email_input = customtkinter.CTkEntry(self.scroll_frame, placeholder_text="Work Email")
            email_input.insert(0, email)
            email_input.grid(row=idx, column=1, padx=5, pady=4, sticky="ew")
            
            self.entries_map[idx] = (name_input, email_input)

    def add_blank_row(self):
        idx = len(self.entries_map)
        name_input = customtkinter.CTkEntry(self.scroll_frame, placeholder_text="Name")
        name_input.grid(row=idx, column=0, padx=5, pady=4, sticky="ew")
        
        email_input = customtkinter.CTkEntry(self.scroll_frame, placeholder_text="Work Email")
        email_input.grid(row=idx, column=1, padx=5, pady=4, sticky="ew")
        
        self.entries_map[idx] = (name_input, email_input)

    def save_data(self):
        updated_db = {}
        for idx, (n_widget, e_widget) in self.entries_map.items():
            name_text = n_widget.get().strip()
            email_text = e_widget.get().strip()
            if name_text:
                updated_db[name_text] = email_text
                
        try:
            with open(JSON_DB_PATH, "w", encoding="utf-8") as f:
                json.dump(updated_db, f, indent=2)
            messagebox.showinfo("Success", "Employee database records updated successfully.")
            self.destroy()
        except Exception as err:
            messagebox.showerror("Error", f"Failed writing changes to JSON file database:\n{err}")


class SheetSelectorDialog(customtkinter.CTkToplevel):
    """
    Modal dialog that lets the user pick a sheet/tab from a scrollable list
    of the ACTUAL sheet names in the uploaded workbook, instead of
    free-typing a name that has to match exactly.

    Built on CTkScrollableFrame (not CTkOptionMenu) specifically so it
    stays usable as more sheets are added -- CTkOptionMenu's dropdown is a
    native tkinter.Menu popup with no guaranteed scrolling for long lists.
    """
    def __init__(self, parent, sheet_names):
        super().__init__(parent)
        self.title("Select Sheet/Tab")
        self.geometry("380x420")
        self.minsize(320, 260)

        self.transient(parent)
        self.grab_set()

        self.result = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        label = customtkinter.CTkLabel(
            self,
            text="Select the sheet/tab to convert:",
            font=customtkinter.CTkFont(size=14),
        )
        label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.scroll_frame = customtkinter.CTkScrollableFrame(
            self, label_text=f"{len(sheet_names)} sheet(s) found"
        )
        self.scroll_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        for idx, name in enumerate(sheet_names):
            sheet_btn = customtkinter.CTkButton(
                self.scroll_frame,
                text=name,
                anchor="w",
                fg_color="transparent",
                text_color=("black", "white"),
                hover_color=("#dcdcdc", "#333333"),
                command=lambda n=name: self.on_select(n),
            )
            sheet_btn.grid(row=idx, column=0, padx=5, pady=3, sticky="ew")

        cancel_btn = customtkinter.CTkButton(
            self, text="Cancel", fg_color="grey", hover_color="#5a5a5a",
            command=self.on_cancel,
        )
        cancel_btn.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")

        # Treat closing the window (the "X" button) the same as Cancel,
        # rather than leaving get_selection() waiting forever.
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def on_select(self, name):
        self.result = name
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

    def get_selection(self):
        """Blocks (like CTkInputDialog.get_input()) until the dialog closes,
        then returns the chosen sheet name, or None if cancelled."""
        self.wait_window(self)
        return self.result


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Tammy's Schedule Converter")
        self.geometry("300x300")
        self.resizable(False, False)
        self.grid_columnconfigure((0, 1), weight=1, uniform="equal")

        self.label = customtkinter.CTkLabel(self, text="No Schedule Selected.", wraplength=280)
        self.label.grid(row=0, column=0, columnspan=3, padx=10, pady=25, sticky="ew")
        
        upload_buton = customtkinter.CTkButton(self, text="Upload Schedule", width=50, corner_radius=10, command=self.select_file)
        upload_buton.grid(row=1, column=0, columnspan=3, padx=15, pady=10, sticky="ew")
        
        name_mapping = customtkinter.CTkButton(self, text="Employee Email", width=35, corner_radius=20, fg_color="grey", command=self.open_database_editor)
        name_mapping.grid(row=2, column=0, columnspan=3, padx=15, pady=10, sticky="ew")
        
    def open_database_editor(self):
        EmailEditorWindow(self)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel (xlsx) File",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        
        if not file_path:
            return

        # Read the actual sheet/tab names out of the uploaded workbook so
        # the dropdown below only ever offers real, exact-match options.
        try:
            probe_wb = openpyxl.load_workbook(file_path, read_only=True)
            # Reversed so the most recently added sheet (last tab in the
            # workbook) shows up first, at the top of the list -- that's
            # normally the current/active schedule, so it's the one you
            # want visible without scrolling.
            sheet_names = list(reversed(probe_wb.sheetnames))
            probe_wb.close()
        except Exception as err:
            messagebox.showerror("Error", f"Could not read sheet names from the selected file:\n{err}")
            return

        if not sheet_names:
            messagebox.showerror("Error", "The selected workbook has no sheets/tabs.")
            return

        # 1. SHEET/TAB SELECTION -- dropdown of real sheet names, not free text
        sheet_dialog = SheetSelectorDialog(self, sheet_names)
        target_sheet_name = sheet_dialog.get_selection()

        if not target_sheet_name:
            messagebox.showwarning("Cancelled", "Operation aborted: no sheet/tab selected.")
            return

        # 2. TRIGGER NATIVE SAVE FILE PATH DIALOG MODAL
        save_dest_path = filedialog.asksaveasfilename(
            title="Export Teams Formatted Schedule As",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="Formatted_Teams_Schedules.xlsx"
        )
        
        if not save_dest_path:
            return
            
        self.label.configure(text=f"Processing tab '{target_sheet_name}'... please wait.")
        self.update_idletasks()
        
        try:
            # Pass choice directly to engine function interface parameter 
            success, missing_employees = xlsx_engine.run_conversion(
                file_path, JSON_DB_PATH, save_dest_path, target_sheet_name
            )
            
            if success:
                self.label.configure(text="Conversion Successful!")
                if missing_employees:
                    names_bullet_list = "\n".join([f" • {name}" for name in missing_employees])
                    warning_msg = (
                        f"Processing completed, but {len(missing_employees)} staff names "
                        f"were missing from your JSON database mappings. They have been assigned temporary placeholder emails.\n\n"
                        f"Missing Names:\n{names_bullet_list}\n\n"
                        f"Click 'Employee Email' below to add them to your directory permanently."
                    )
                    messagebox.showwarning("Missing Email Profile Matches Detected", warning_msg)
                else:
                    messagebox.showinfo("Done", f"Tab '{target_sheet_name}' fully synchronized and successfully generated.")
            else:
                self.label.configure(text="Error processing file pipeline.")
                messagebox.showerror("Execution Error", "An issue disrupted the background compilation cycle.")
                
        except ValueError as val_err:
            # Catches sheet mismatch validation trigger safely
            self.label.configure(text="Sheet Name Error.")
            messagebox.showerror("Sheet Mismatch", str(val_err))
        except Exception as err:
            self.label.configure(text="Unexpected Failure.")
            messagebox.showerror("System Error", f"An unexpected issue occurred:\n{err}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
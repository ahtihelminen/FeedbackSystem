import tkinter as tk
from tkinter import filedialog
import systems_update
import areas_update
import hull_update

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Feedback Updater")
        self.master.geometry("640x480")
        self.pack(fill="both", expand=True)
        self.create_widgets()


    def create_widgets(self):
        # Create a button to select a file
        self.select_file_button = tk.Button(self)
        self.select_file_button["text"] = "Select Excel File"
        self.select_file_button["command"] = self.select_file
        self.select_file_button.place(relx=0.25, rely=0.15, anchor="center")
        

        # Create a label to show the selected file
        self.filename_label = tk.Label(self)
        self.filename_label["text"] = "No file selected"
        self.filename_label["font"] = ("Arial", 12)
        self.filename_label["borderwidth"] = 2
        self.filename_label["relief"] = "solid"
        self.filename_label["width"] = 40
        self.filename_label.place(relx=0.65, rely=0.15, anchor="center")

        self.radio_menu_label = tk.Label(self)
        self.radio_menu_label["text"] = "Select mode:"
        self.radio_menu_label["font"] = ("Arial", 12)
        
        self.radio_menu_label.place(relx=0.15, rely=0.25, anchor="w")

        # Option menu to select the mode
        self.radio_var = tk.StringVar()
        self.radio_var.set("init_val")  # set the default value
        
        self.radio_button_area_managers = tk.Radiobutton(self, text="Area managers and outfitting foreman", variable=self.radio_var, value="Area managers and outfitting foreman")
        self.radio_button_area_managers.place(relx=0.15, rely=0.3, anchor="w")

        self.radio_button_comissioning = tk.Radiobutton(self, text="Comissioning", variable=self.radio_var, value="Comissioning")
        self.radio_button_comissioning.place(relx=0.15, rely=0.35, anchor="w")

        self.radio_button_design = tk.Radiobutton(self, text="Design", variable=self.radio_var, value="Design")
        self.radio_button_design.place(relx=0.15, rely=0.4, anchor="w")
        
        self.radio_button_electric = tk.Radiobutton(self, text="Electric", variable=self.radio_var, value="Electric")
        self.radio_button_electric.place(relx=0.15, rely=0.45, anchor="w")
        
        self.radio_button_hull = tk.Radiobutton(self, text="Hull", variable=self.radio_var, value="Hull")
        self.radio_button_hull.place(relx=0.15, rely=0.5, anchor="w")
        
       

        # Create a button to run the feedback update script
        self.update_button = tk.Button(self)
        self.update_button["text"] = "Update Feedback"
        self.update_button["command"] = self.run_script
        self.update_button.place(relx=0.5, rely=0.7, anchor="center")


    def select_file(self):
        # Open a file dialog to select an Excel file
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.filename = filename
            self.filename_label["text"] = filename.split("/")[-1]

    def run_script(self):
        # Run the feedback update script with the selected file
        if hasattr(self, "filename"):
            
            if self.radio_var.get() == "init_val":
                print("Please select a mode!") #label to be mades
                return
            
            elif self.radio_var.get() == "Area managers and outfitting foreman":
                updater = areas_update.AreasUpdate(self.filename, '../databases/feedbackDatabaseTest.json')
            
            elif self.radio_var.get() == "Comissioning":
                updater = systems_update.SystemsUpdate(self.filename, '../databases/feedbackDatabaseTest.json', 'Comissioning')

            elif self.radio_var.get() == "Design":
                updater = systems_update.SystemsUpdate(self.filename, '../databases/feedbackDatabaseTest.json', 'Design')
            
            elif self.radio_var.get() == "Electric":
                updater = systems_update.SystemsUpdate(self.filename, '../databases/feedbackDatabaseTest.json', 'Electric')
            
            elif self.radio_var.get() == "Hull":
                updater = hull_update.HullUpdate(self.filename, '../databases/feedbackDatabaseTest.json')
            
            try:
                updater.main()
                print("Feedback updated successfully!")
                tk.messagebox.showinfo(title='Succes', message='Update finished')
            except TypeError as TE:
                tk.messagebox.showerror(title='Error', message='Incorrect file or mode selected!')
                print("Error in feedback update:", TE)

                

root = tk.Tk()
app = Application(master=root)
app.mainloop()

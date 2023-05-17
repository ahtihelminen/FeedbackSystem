import tkinter as tk
from tkinter import filedialog
import systems_update
import areas_update
import hull_update
import os
from tools import Tools

class Application(tk.Frame, Tools):
    def __init__(self, master=None):
        super().__init__(master)
        self.feedbackDatabasePathRel = '../../databases/feedbackDatabase.json'
        self.master = master
        self.master.title("Feedback Updater")
        self.master.geometry("640x480")
        self.master.resizable(False, False)
        self.pack(fill="both", expand=True)
        self.ship = 'No ship selected'
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

        self.select_mode_label = tk.Label(self)
        self.select_mode_label["text"] = "Select mode:"
        self.select_mode_label["font"] = ("Arial", 12)
        self.select_mode_label.place(relx=0.15, rely=0.25, anchor="w")

        self.ship_label = tk.Label(self)
        self.ship_label["text"] = f"Ship: {self.ship}"
        self.ship_label["font"] = ("Arial", 12)
        self.ship_label["borderwidth"] = 2
        self.ship_label["relief"] = "solid"
        self.ship_label["width"] = 20
        self.ship_label.place(relx=0.55, rely=0.25, anchor="w")

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
        
        # input field to write the ship
        self.ship_input = tk.Entry(self)
        self.ship_input.place(relx=0.55, rely=0.32, anchor="w")

        # button to set the ship
        self.ship_button = tk.Button(self)
        self.ship_button["text"] = "Set ship"
        self.ship_button["command"] = self.set_ship
        self.ship_button.place(relx=0.8, rely=0.32, anchor="w")



        # Create a button to run the feedback update script
        self.update_button = tk.Button(self)
        self.update_button["text"] = "Update Feedback"
        self.update_button["command"] = self.run_script
        self.update_button.place(relx=0.5, rely=0.7, anchor="center")


    def set_ship(self):
        self.ship = self.ship_input.get()
        tk.messagebox.showinfo(title='Succes', message=f'Ship set to "{self.ship}"')
        self.ship_label["text"] = f"Ship: {self.ship}"

    def select_file(self):
        # Open a file dialog to select an Excel file
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.filename = filename
            self.filename_label["text"] = filename.split("/")[-1]

    def run_script(self):
        # Run the feedback update script with the selected file
        if hasattr(self, "filename"):
            
            if self.ship == 'No ship selected':
                tk.messagebox.askquestion(title='confirmation', message='Do you wish to continue without setting a ship?', icon='warning', default='no')
                return

            if self.radio_var.get() == "init_val":
                print("Please select a mode!") #label to be mades
                return
            
            elif self.radio_var.get() == "Area managers and outfitting foreman":
                updater = areas_update.AreasUpdate(self.filename, self.feedbackDatabasePathRel, self.ship)
            
            elif self.radio_var.get() == "Comissioning":
                updater = systems_update.SystemsUpdate(self.filename, self.feedbackDatabasePathRel, 'Comissioning', self.ship)

            elif self.radio_var.get() == "Design":
                updater = systems_update.SystemsUpdate(self.filename, self.feedbackDatabasePathRel, 'Design', self.ship)
            
            elif self.radio_var.get() == "Electric":
                updater = systems_update.SystemsUpdate(self.filename, self.feedbackDatabasePathRel, 'Electric', self.ship)
            
            elif self.radio_var.get() == "Hull":
                updater = hull_update.HullUpdate(self.filename, self.feedbackDatabasePathRel, self.ship)
            
            try:
                updater.main()
                print("Feedback updated successfully!")
                tk.messagebox.showinfo(title='Succes', message='Update finished')
            except TypeError as TE:
                tk.messagebox.showerror(title='Error', message='Incorrect file or mode selected!')
                print("Error in feedback update:", TE)

                
print(os.path.dirname(__file__))
root = tk.Tk()
app = Application(master=root)
app.mainloop()

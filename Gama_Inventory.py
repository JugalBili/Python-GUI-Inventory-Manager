# -*- coding: utf-8 -*-

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import string

# initializng google sheets and authorizing credentials
scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds  = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

# opens sheets 
spreadsheet = client.open("test gama")
sheet = spreadsheet.worksheet("Sheet1")
try: 
    old_sheet = spreadsheet.worksheet("OLD")
    spreadsheet.del_worksheet(old_sheet)
    sheet.duplicate(insert_sheet_index = 1, new_sheet_name = "OLD")
except:
    sheet.duplicate(insert_sheet_index = 1, new_sheet_name = "OLD")
    old_sheet = spreadsheet.worksheet("OLD")

root = Tk()
root.geometry("400x660")

# the following code initializes global variables 
music = sheet.row_values(1)
music[0] = ""

instruments = sheet.col_values(1)
instruments.pop(0)

instrument_values = []
instrument_selected = ""
new_window = ""
change_value = ""

class Events():

    """
    This class is used to contain methods which are called when specific actions take place in requards to user interaction with the GUI

    Methods: 
        box_selected - called when user selects a music piece from the drop down combobox
        treeview_selected - called when an instrument is selected from the displayed treeview list
        value_entered - called when the user inputs a new value after choosing the item from the treview 
        edit_columns - called when the user selects the button to edit the columns (add/remove music pieces in the dropdown combobox)

    """

    def box_selected(self, event):
        """ 
        Method changes the LabelFrame text to the music piece selected, and goes through google sheets and adds values to treeview
        """
        
        global midframe, treeview, instruments, instrument_values
        
        midframe.config(text = combo_box.get())
        
        cell = sheet.find(combo_box.get()) # finds cell of music sheet selected
        instrument_values = sheet.col_values(cell.col) # stores values as list

        instrument_values.pop(0) 
        
        # loops through instrument and instrument_values list and displays them on the treview
        for x in range(0,len(instruments)):
            treeview.set(instruments[x], "Amount", instrument_values[x])
        
        # changes treeview mode so users are able to select data
        treeview.config(selectmode = "browse")
            
    
    def treeview_selected(self, event):
        """
        Method creates a new window for users to enter a new value to replace the current value for the selected item 
        """
        
        global instrument_selected, change_value, new_window
        
        selected_dict = treeview.item(treeview.selection()[0])
        instrument_selected = selected_dict["text"] 
        
        # initializes new window
        new_window = Toplevel(root)
        new_window.geometry("200x100")
        frame = Frame(new_window)
        frame.pack()
        
        #initializing buttons and labels in the new window
        sel_label = Label(frame, text = "You have selected: " + instrument_selected) 
        change_label = Label(frame, text = "Enter Amount: ")
        change_value = Entry(frame, width = 5)
        change_value.insert(0,selected_dict["values"][0])
        
        enter_btn = Button(frame, text = "Enter", command = lambda: self.value_entered(amount = int(change_value.get())))
    
        Label(frame, padx = 2).grid(row = 0, column = 0) # space filler
        sel_label.grid(row = 1, column = 0, columnspan = 2)
        change_label.grid(row = 2, column = 0)
        change_value.grid(row = 2, column = 1)
        enter_btn.grid(row = 3, column = 0, columnspan = 2)
        Label(frame, padx = 2).grid(row = 4, column = 0) # space filler
        
        #binds the "enter" key to call the value_entered method when the entryfield is active (meaning it is selected by the cursor)
        change_value.bind("<Return>", lambda x : self.value_entered(int(change_value.get())))
        
            
    def value_entered(self, amount):
        """
        Method changes the current value displayed in the treeview to the new one entered, and updates the value in the list and google sheets
        """
        
        global midframe, treeview, instruments, instrument_values, instrument_selected, change_value, change_label, new_window, sheet, music, combo_box
        
        # updates the list to the new value   
        index = instruments.index(instrument_selected)
        instrument_values[index] = amount

        # update sthe treeview to the new value  
        treeview.set(instrument_selected, "Amount", amount)
       
        column = music.index(combo_box.get()) + 1
        row = index + 2
        
        # update sthe google sheets
        sheet.update_cell(row, column, amount)
            
        new_window.destroy()
            

    def edit_columns(self):
        """
        Method creates a new window which allows the user to add or remove new/existing music pieces displayed in the dropdown combobox
        """
        
        global music, combo_box, sheet 
        
        # initializes new window
        edit_window = Toplevel(root)
        edit_window.geometry("280x150")
        edit_frame = Frame(edit_window)
        edit_frame.pack()
        
        #initializes buttons and labels in the new window
        edit_label = Label(edit_frame, text = "Enter Music name:   ")
        col_entry = Entry(edit_frame, width = 25)
        add_button = Button(edit_frame, text = "Add Column", command = lambda: add(col_entry.get()))
        del_button = Button(edit_frame, text = "Delete Column", command = lambda: delete(col_entry.get()))
        
        Label(edit_frame, pady = 1).grid(row = 0, column = 0) # space Filler
        edit_label.grid(row = 1, column = 0)
        col_entry.grid(row = 1, column = 1)
        Label(edit_frame, pady = 1).grid(row = 2, column = 0) # space Filler
        add_button.grid(row = 3, column = 0, columnspan = 2)
        del_button.grid(row = 4, column = 0, columnspan = 2)
            
        def add(song):
            """
            Sub-method is called when the user presses the add button. 
            It creates a new music piece column in google sheets and default sets all the instduemnt values to 0.
            It also formats the main music piece to centered and bold, and values to centered in google sheets.
            """
            # does the following only if the entered name is not blank
            if ((song.replace(" ", "")) != ""):
                
                global music, combobox

                # adds a new column to the sheets
                column_num = len(music) + 1
                sheet.update_cell(1, column_num, string.capwords(song.lower()))
                
                # loop whihc makes all the instrument values 0 for the new column
                for x in range(0,len(instruments)):
                    sheet.update_cell(x+2, column_num, 0)
                
                # displays message
                messagebox.showinfo("Edit Column", "Column added sucessfully") 
                edit_window.lift()
                
                music = sheet.row_values(1)
                music[0] = ""
                combo_box.config(value = music)
                
                # formats all the new added cells in google sheets
                cell_location = chr(65+(column_num-1)) + "1:" + chr(65+(column_num-1)) + "1"
                column_range = chr(65+(column_num-1)) + "2:" + chr(65+(column_num-1)) + "23"
                sheet.format(cell_location, {'textFormat': {'bold': True}, "horizontalAlignment": "CENTER"})
                sheet.format(column_range, {"horizontalAlignment": "CENTER", "textFormat": {"fontSize": 11}})
                
                    
            else:
                messagebox.showerror("Edit Column", "Please enter a proper column name to create")
                edit_window.lift()
                
        def delete(song): 
            """
            Sub-method is called when user presses the delete button. 
            It deletes the column of the selected music piece entered on google sheets and stored lists. 
            """
            
            try:
                
                global music, combo_box
                
                # deletes the column on google sheets 
                cell = sheet.find(string.capwords(song.lower()))
                sheet.delete_columns(cell.col)
                
                music = sheet.row_values(1)
                music[0] = ""
                combo_box.config(value = music)
            
            # catches an error if the entered column name does not correspond to any column 
            except: 
                messagebox.showerror("Edit Column", "Please enter valid existing column name")
                edit_window.lift()
                
            else: 
                messagebox.showinfo("Edit Column", "Column deleted sucessfully")
                edit_window.lift()
                
      


# 4 frames to divide window
titleframe = Frame(root)
topframe = Frame(root)
midframe = LabelFrame(root)
botframe = Frame(root)

titleframe.grid(row = 0, column = 0)
topframe.grid(row = 1, column = 0)
midframe.grid(row = 2, column = 0)
botframe.grid(row = 3, column = 0)

event_class = Events()
    
# titleframe includes the main title of the program
title = Label(titleframe, text = "GAMA Inventory", font = ("", 20, "bold"), fg = "#8B0000", width = 24, justify = "center")
title.pack(anchor = "center")
Label(titleframe, pady = 1).pack() # space filler


# topframe includes the combox for users to select sheet music
box_label = Label(topframe, text = "Select music: ")
combo_box = ttk.Combobox(topframe, value = music, state = "readonly")
combo_box.current(0)
edit_button = Button(topframe, text = "Edit Columns", command = event_class.edit_columns) 

box_label.grid(row = 0, column = 0)
Label(topframe, padx = 2).grid(row = 0, column = 1) # space filler
combo_box.grid(row = 0, column = 2)
edit_button.grid(row = 1, column = 1, columnspan = 3)
Label(topframe, pady = 1).grid(row = 2, column = 0) # space Filler


# midframe includes the treeview 
treeview = ttk.Treeview(midframe, height = 22, columns = ("Amount"), selectmode = "none")
treeview.pack()

# adds heading text for each column
treeview.heading("#0", text = "Instrument") #id of first column in #0
treeview.heading("Amount", text = "Amount")
# changes size of column 
treeview.column("#0", width = 150)
treeview.column("Amount", anchor = CENTER, width = 60)

# loop which adds all instruments to column 1 of treeview
for x in range(0,len(instruments)):

    treeview.insert("", x, instruments[x], text = instruments[x])


# botframe includes label which tells users how to edit date

inst_label = Label(botframe, text = "To edit data, select row in list above")
Label(botframe, padx = 2).pack() # space filler
inst_label.pack()



# binds actions to automaticially call upon functions when clicked
combo_box.bind("<<ComboboxSelected>>", event_class.box_selected)
treeview.bind("<<TreeviewSelect>>", event_class.treeview_selected)


root.mainloop()

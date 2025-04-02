"""
* NAME               - Fuzzy Matching Tool v1.4
* AUTHOR             - James Legrand
* DATE LAST MODIFIED - 26/03/2025
"""
# %% Setup

# Base libraries
import sys
import os
import queue
import threading
import random
from datetime import datetime
from typing import Callable, Union
from tkinter import messagebox

# External libraries
import pandas as pd
import rapidfuzz as rf
import customtkinter as ctk
from unidecode import unidecode

# %% Define helper classes

class TextRedirector:
    '''
    A utility class to redirect text output to a CTkinter Textbox widget.
    '''
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        if string.strip():  # Only log non-empty lines
            self.widget.after(0, self.widget.insert, "end", string + "\n") # Insert line between writes
            self.widget.after(0, self.widget.see, "end") # Scroll to latest line
            
    def flush(self):
        pass  # No buffering required


class IntSpinbox(ctk.CTkFrame):
    '''
    A utility class to produce an integer spinbox widget.
    '''
    def __init__(self, *args,
                 width: int = 150,
                 height: int = 32,
                 step_size: int = 1,
                 command: Callable = None,
                 max_entry: int = 100,
                 **kwargs):
        super().__init__(*args, width=width, height=height, **kwargs)

        self.step_size = step_size
        self.command = command
        self.max_entry = max_entry

        self.configure(fg_color=["gray92", "gray14"])

        self.grid_columnconfigure((0, 2), weight=0)
        self.grid_columnconfigure(1, weight=1)

        self.subtract_button = ctk.CTkButton(self, text="-", width=height-6, height=height-6,
                                             command=self.subtract_button_callback)
        self.subtract_button.grid(row=0, column=0, padx=(3, 0), pady=3)

        self.entry = ctk.CTkEntry(self, width=width-(2*height), height=height-6, border_width=0)
        self.entry.grid(row=0, column=1, columnspan=1, padx=3, pady=3, sticky="ew")

        self.add_button = ctk.CTkButton(self, text="+", width=height-6, height=height-6,
                                        command=self.add_button_callback)
        self.add_button.grid(row=0, column=2, padx=(0, 3), pady=3)

        self.entry.insert(0, "80")

    def add_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = int(self.entry.get()) + self.step_size
            if 0 <= value <= self.max_entry:
                self.entry.delete(0, "end")
                self.entry.insert(0, value)
        except ValueError:
            return

    def subtract_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = int(self.entry.get()) - self.step_size
            if 0 <= value <= self.max_entry:
                self.entry.delete(0, "end")
                self.entry.insert(0, value)
        except ValueError:
            return

    def get(self) -> Union[int, None]:
        try:
            return int(self.entry.get())
        except ValueError:
            return None

    def set(self, value: int):
        self.entry.delete(0, "end")
        self.entry.insert(0, str(value))

# %% Define app class

class MatchingTool:
    
    def __init__(self):     
        # Initialise root window 
        self.root = ctk.CTk()
        self.root.resizable(width=False, height=False)
        self.root.title("Fuzzy Matching Tool")
        self.root.geometry("550x600")

        # Initialise variables 
        self.dataset_1_path = ctk.StringVar()                               # Path to dataset 1
        self.dataset_2_path = ctk.StringVar()                               # Path to dataset 2
        self.output_path = ctk.StringVar()                                  # Path to output file
        self.dataset_1_id_col = ctk.StringVar(value='ID Column')            # ID variable for dataset 1
        self.dataset_1_match_col_1 = ctk.StringVar(value='Match Column')    # Match variable 1 for dataset 1
        self.dataset_1_match_col_2 = ctk.StringVar(value='Match Column 2')  # Match variable 2 for dataset 1
        self.dataset_2_id_col = ctk.StringVar(value='ID Column')            # ID variable for dataset 2
        self.dataset_2_match_col_1 = ctk.StringVar(value='Match Column')    # Match variable 1 for dataset 2
        self.dataset_2_match_col_2 = ctk.StringVar(value='Match Column 2')  # Match variable 2 for dataset 2
        self.output_type_var = ctk.IntVar(value=1)                          # 1: All possible combinations, 2: Highest matches only, 3: Matches above threshold
        self.matching_type_var = ctk.IntVar(value=1)                        # 1: Set ratio, 2: Sort ratio, 3: Max(Set ratio, Sort ratio), 4: QRatio
        self.multi_match_var = ctk.IntVar(value=0)                          # Dummy taking value 1 when matching on multiple columns            
        self.score_method_var = ctk.StringVar(value="Score Method")         # Formula for combining scores when matching on multiple columns (maximum, minimum or weighted average)
        self.weight_var = ctk.DoubleVar(value = 0.5)                        # Weight on score 1 when calculating weighted average
        self.dataset_cache = {}                                             # Cache to store dimensions of dataset for display on hover
        self.is_advanced_visible = False                                    # Boolean for displaying advanced options window
        self.theme = "dark"                                                 # Colour scheme (light or dark)
        self.fact_switch_flag = ctk.IntVar(value=0)                         # Flag to toggle if animal facts are displayed on completion 

        # Set theme
        ctk.set_appearance_mode(self.theme)

        # Add a title to the root frame
        self.title = ctk.CTkLabel(self.root, 
                                  text="Fuzzy Matching Tool", 
                                  font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        
        # Add a switch to toggle between light and dark theme to the root frame
        self.theme_switch = ctk.CTkSwitch(self.root, 
                                          text = "Theme", 
                                          command = self.toggle_theme)
        self.theme_switch.place(x=10, y=5, anchor="nw")

        # Add a switch to toggle the visibility of advanced options to the root frame
        self.advanced_switch = ctk.CTkSwitch(self.root, 
                                             text="Advanced Options", 
                                             command=self.toggle_advanced_options)
        self.advanced_switch.place(x=10, y=27, anchor="nw")

        # Create main frame for base display
        self.main_frame = ctk.CTkFrame(self.root, width=600, height=600, fg_color=["gray92", "gray14"])
        self.main_frame.place(x=20, y=50, anchor="nw")
        
        # Create the filled frame for dataset imports
        self.create_dataset_frame(self.main_frame,
                                  1,
                                  self.dataset_1_path,
                                  self.dataset_1_id_col,
                                  self.dataset_1_match_col_1)
        
        self.create_dataset_frame(self.main_frame,
                                  2,
                                  self.dataset_2_path,
                                  self.dataset_2_id_col,
                                  self.dataset_2_match_col_1)
        
        self.output_path_button = ctk.CTkButton(self.main_frame,
                                                 text = "Select Output File", 
                                                 command = lambda: self.browse_file(self.output_path,
                                                                                    self.output_path_button,
                                                                                    is_output = True), 
                                                 width = 200)
        self.output_path_button.pack(pady=10)

        # Create radio buttons to select output type
        self.matching_title = ctk.CTkLabel(self.main_frame, text="Select Output Type", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=5)
        
        # Create a button for computing score for all combinations of strings
        self.output_button_all = ctk.CTkRadioButton(self.main_frame, 
                                                    text="All Possible Combinations",
                                                    variable=self.output_type_var, 
                                                    value=1, 
                                                    radiobutton_width=18, 
                                                    radiobutton_height=18, 
                                                    border_width_checked=5)
        self.output_button_all.pack()
        
        # Create a button for computing only the highest score for each entry
        self.output_button_best = ctk.CTkRadioButton(self.main_frame,
                                                      text="Highest Matches Only", 
                                                      variable=self.output_type_var, 
                                                      value=2, radiobutton_width=18, 
                                                      radiobutton_height=18, 
                                                      border_width_checked=5)
        self.output_button_best.pack()
        
        # Create a frame to store the radio button and spinbox for computing matches above a given threshold.
        self.threshold_frame = ctk.CTkFrame(self.main_frame, fg_color=["gray92", "gray14"])
        self.threshold_frame.pack(pady=0)

        # Create the radio button
        self.output_button_threshold = ctk.CTkRadioButton(self.threshold_frame, 
                                                          text="Matches Above Threshold", 
                                                          variable=self.output_type_var, 
                                                          value=3, 
                                                          radiobutton_width=18, 
                                                          radiobutton_height=18, 
                                                          border_width_checked=5)
        self.output_button_threshold.grid(row=0, column=0, padx=5)

        # Create spinbox to select threshold score
        self.score_threshold_spinbox = IntSpinbox(self.threshold_frame, width=100, step_size=1)
        self.score_threshold_spinbox.grid(row=0, column=1, padx=5, pady=0)

        # Radio buttons for selecting matching type
        self.matching_title = ctk.CTkLabel(self.main_frame,
                                           text="Select Matching Algorithm",
                                           font=ctk.CTkFont(size=13, weight="bold")).pack(pady=5)

        # Create a button to select set ratio
        self.matching_button_set = ctk.CTkRadioButton(self.main_frame, 
                                                      text="Set Ratio", 
                                                      variable=self.matching_type_var, 
                                                      value=1, 
                                                      radiobutton_width = 18, 
                                                      radiobutton_height = 18, 
                                                      border_width_checked=5).pack()
        # Create a button to select sort ratio
        self.matching_button_sort = ctk.CTkRadioButton(self.main_frame, 
                                                       text="Sort Ratio", 
                                                       variable=self.matching_type_var, 
                                                       value=2, radiobutton_width = 18, 
                                                       radiobutton_height = 18, 
                                                       border_width_checked=5).pack()
        
        # Create a button to select the max of set and sort ratio
        self.matching_button_max = ctk.CTkRadioButton(self.main_frame, 
                                                      text="Max of (Set Ratio, Sort Ratio)", 
                                                      variable=self.matching_type_var, 
                                                      value=3, 
                                                      radiobutton_width = 18, 
                                                      radiobutton_height = 18, 
                                                      border_width_checked=5).pack()
        # Create a button to select quick ratio
        self.matching_button_quick = ctk.CTkRadioButton(self.main_frame, 
                                                        text="QRatio", 
                                                        variable=self.matching_type_var, 
                                                        value=4, 
                                                        radiobutton_width = 18, 
                                                        radiobutton_height = 18, 
                                                        border_width_checked=5).pack()
 
        # Create a progress bar to display progress of matching operation
        self.progress_bar = ctk.CTkProgressBar(self.main_frame, width=300)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)
        self.progress_label = ctk.CTkLabel(self.main_frame, text="Progress: 0/0")
        self.progress_label.pack(pady=5)

        # Create a button to begin the matching operation
        self.run_matching_button = ctk.CTkButton(self.main_frame, 
                                                 text="Run Matching", 
                                                 command=self.run_matching, width=200)
        self.run_matching_button.pack(pady=10)

        # Create the advanced frame
        self.advanced_frame = ctk.CTkFrame(self.root, width=500, height=500, fg_color=["gray92", "gray14"])

        # Add a debugging window to the advanced frame
        self.terminal_output_text = ctk.CTkTextbox(self.advanced_frame, width=450, height=250, wrap="word")

        # Redirect print statements (sys.standard out) and errors (sys.standard error) to the debug window
        redirector = TextRedirector(self.terminal_output_text)
        sys.stdout = redirector
        sys.stderr = redirector

        # Create toggle to enable/disable matching on 2 columns
        self.multi_match_switch = ctk.CTkSwitch(self.advanced_frame, 
                                                command = self.toggle_multi_match, 
                                                text="2 Column Match")
        
        # Create dropdowns to set 2nd set of match columns
        self.dataset_1_match_2_dropdown = ctk.CTkOptionMenu(self.advanced_frame, 
                                                            variable= self.dataset_1_match_col_2, 
                                                            values=["Match Column 2"])
        
        self.dataset_2_match_2_dropdown = ctk.CTkOptionMenu(self.advanced_frame, 
                                                            variable= self.dataset_2_match_col_2, 
                                                            values=["Match Column 2"])

        # Create dropdown to select method of combining scores across match sets
        self.combine_score_dropdown = ctk.CTkOptionMenu(self.advanced_frame, 
                                                        variable=self.score_method_var, 
                                                        values=["Maximum", "Minimum", "Weighted Average"])
        
        # Create a labeled slider to select the weight on score 1 when using weighted average
        self.weight_1_slider = ctk.CTkSlider(self.advanced_frame, 
                                             from_=0, 
                                             to = 1, 
                                             variable = self.weight_var, 
                                             width = 100)
        self.slider_label = ctk.CTkLabel(self.advanced_frame, text = "Score 1 Weight:")
        self.slider_value = ctk.CTkLabel(self.advanced_frame, textvariable = self.weight_var)

        # Create switches for quality of life toggles
        self.fact_switch = ctk.CTkSwitch(self.advanced_frame,
                                         text="Fuzzy animal fact",
                                         variable=self.fact_switch_flag)
        
        self.ascii_convert_switch = ctk.CTkSwitch(self.advanced_frame,
                                                 text="Convert to ASCII")
        
        self.clean_switch = ctk.CTkSwitch(self.advanced_frame,
                                          text="Prep for manual checks")

        self.keep_columns_switch = ctk.CTkSwitch(self.advanced_frame,
                                                 text="Keep all columns")


        # Start the event loop
        self.root.mainloop()

    def __del__(self):
        # Reset stdout and stderr to default
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        
    def create_dataset_frame(self,frame, dataset_num, path_variable, id_variable, match_variable):
        '''
        Creates the path button and ID + match column dropdown for a given dataset.
        Automatically fills the dropdowns with variables from the dataset.
        Retrieve cached dataset dimensions to display on hover.

        Parameters
        ----------
        frame
            CTkinter frame in which to place the widgets
        dataset_num
            Dataset number (1 or 2) for text display
        path_variable
            Path to input file
        id_variable
            CTk string var storing the ID variable name
        match_variable
            CTk string var storing the match variable name
        '''
        # Create a frame in which to place the widgets
        dataset_frame = ctk.CTkFrame(frame, fg_color=["gray92", "gray14"])
        dataset_frame.pack(pady=10)

        # Create the button to select the input file
        dataset_button = ctk.CTkButton(dataset_frame, 
                                       text=f"Select Dataset {dataset_num}", 
                                       command=lambda: self.browse_file(path_variable, dataset_button), 
                                       width=200)
        dataset_button.grid(row=0, column=0, padx=5)
    
        # Create the dropdown for the ID column
        id_dropdown = ctk.CTkOptionMenu(dataset_frame, variable=id_variable, values=["ID Column"])
        id_dropdown.grid(row=0, column=1, padx=5)
    
        # Create the dropdown for the primary Match column
        match_dropdown = ctk.CTkOptionMenu(dataset_frame,
                                           variable=match_variable,
                                           values=["Match Column"])
        match_dropdown.grid(row=0, column=2, padx=5)
    
        # Fill dropdowns when dataset is imported
        setattr(self, f"dataset_{dataset_num}_id_dropdown", id_dropdown)
        setattr(self, f"dataset_{dataset_num}_match_1_dropdown", match_dropdown)
    
        # Display dimensions on hover
        dataset_button.bind(
            "<Enter>", 
            lambda event: self.show_dataset_info(path_variable, event, dataset_button)
        )

        # Remove dimensions when hover ends
        dataset_button.bind(
            "<Leave>", 
            lambda event: dataset_button.configure(
                text=os.path.basename(path_variable.get()) 
                if path_variable.get() 
                else f"Select Dataset {dataset_num}"
            )
        )

    def browse_file(self, path, label, is_output=False):
        '''
        Loads the selected dataset and replaces the button label with the filename. 
        Caches dataset dimensions for display later.

        Parameters
        ----------
        path
            Path to input or output file (xlsx, csv or dta)
        label
            Button to update the label of.
        is_output, optional
            Boolean indicating whether the button is the output path, by default False.
            If the file is the output path dimensions are not cached.
        '''
        # Grab file path from the interactive dialog. Set valid filetypes
        file_path = ctk.filedialog.askopenfilename(filetypes=[("Data file", "*.csv"),
                                                                ("Data file", "*.xlsx"),
                                                                ("Data file", "*.dta")])
        if file_path:
            # Set the path to the selected file
            path.set(file_path)
            # Label the button according to the file path
            label.configure(text=os.path.basename(file_path), text_color="white")
            
            # For input datasets only, cache dimensions when imported
            if not is_output:
                if file_path not in self.dataset_cache:   
                    if file_path.endswith('.xlsx'):
                        df = pd.read_excel(file_path)
                    elif file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                    elif file_path.endswith('.dta'):
                        df = pd.read_stata(file_path)
                    columns = df.columns.tolist()

                    self.dataset_cache[file_path] = df.shape 

                    dropdown_id = self.dataset_1_id_dropdown if path == self.dataset_1_path else self.dataset_2_id_dropdown
                    dropdown_match_1 = self.dataset_1_match_1_dropdown if path == self.dataset_1_path else self.dataset_2_match_1_dropdown
                    dropdown_match_2 = self.dataset_1_match_2_dropdown if path == self.dataset_1_path else self.dataset_2_match_2_dropdown

                    # Reset the dropdowns
                    dropdown_id.set("ID Column")
                    dropdown_match_1.set("Match Column")
                    dropdown_match_2.set("Match Column 2")

                    # Replace the dropdowns options
                    dropdown_id.configure(values=columns)
                    dropdown_match_1.configure(values=columns)
                    dropdown_match_2.configure(values=columns)

            self.debug_message(f"Path set: {file_path}")

    def show_dataset_info(self, path_var, event, label):
        '''
        Helper function to display dataset dimensions on hover over button.
        '''
        file_path = path_var.get()
        if file_path:
            if file_path in self.dataset_cache:
                rows, cols = self.dataset_cache[file_path]
                label.configure(text=f"{cols} columns, {rows} rows")

    def toggle_theme(self):
        '''
        Helper function to switch display theme.
        '''
        self.theme = "light" if self.theme == "dark" else "dark"
        ctk.set_appearance_mode(self.theme) 

    def toggle_multi_match(self):
        '''
        Helper function to toggle 2 column matching.
        '''
        # Toggle the multi match variable between 0 and 1
        new_state = 1 - self.multi_match_var.get() 
        self.multi_match_var.set(new_state)

        # If multi match is toggled on, switch matching type to all combos.
        state_config = "disabled" if new_state else "normal"
        self.output_button_best.configure(state=state_config)
        self.output_button_threshold.configure(state=state_config)
        self.output_type_var.set(1)

        self.debug_message(f"Matching on single variable: {bool(new_state)}")

    def toggle_advanced_options(self):
        '''
        Helper function to toggle visibility of the advanced options pane.
        '''
        if self.is_advanced_visible:

            # Shrink the window and forget the frame
            self.root.geometry("550x600")
            self.advanced_frame.place_forget()

        else:

            # Expand the window and place the frame
            self.root.geometry("1050x600")
            self.advanced_frame.place(relx=0.97, y=50,  anchor = "ne")

            # Place all the additional settings
            self.terminal_output_text.grid(row = 0, column = 0, columnspan = 3, pady = 10)
            self.multi_match_switch.grid(row=1, column=0, padx=5, pady=5, sticky="w")
            self.dataset_1_match_2_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="w")
            self.dataset_2_match_2_dropdown.grid(row=1, column=2, padx=5, pady=5, sticky="w")
            self.combine_score_dropdown.grid(row=2, column=0, padx=5, pady=5, sticky="w")
            self.weight_1_slider.grid(row=2, column=1, padx=5, pady=5, columnspan=1)  
            self.slider_label.grid(row=2, column=2, padx=5, pady=5, sticky ="w") 
            self.slider_value.grid(row=2, column=2, padx=5, pady=5, sticky="e") 
            self.fact_switch.grid(row=3, column=1, columnspan=3, pady=5, sticky ="w")
            self.ascii_convert_switch.grid(row=4, column=1, columnspan=3, pady=5, sticky ="w")
            self.clean_switch.grid(row=5, column=1, columnspan=3, pady=5, sticky ="w")
            self.keep_columns_switch.grid(row=6, column=1, columnspan=3, pady=5, sticky ="w")

        # Flip the boolean
        self.is_advanced_visible = not self.is_advanced_visible

    def validate_inputs(self):
        '''
        A function containing all data validation checks that must pass before execution.
        If any fail, displays a popup and breaks the run loop.

        Returns
        -------
            Boolean whether to continue execution
        '''
        # Check if Dataset 1 path is provided
        if not self.dataset_1_path.get():
            self.show_error("Please select Dataset 1.")
            return False
    
        # Check if Dataset 2 path is provided
        if not self.dataset_2_path.get():
            self.show_error("Please select Dataset 2.")
            return False
    
        # Check if Output path is provided
        if not self.output_path.get():
            self.show_error("Please select an output file.")
            return False
    
        # Check if Dataset 1 ID column is selected
        if not self.dataset_1_id_col.get() or self.dataset_1_id_col.get() == "ID Column":
            self.show_error("Please select an ID column for Dataset 1.")
            return False
    
        # Check if Dataset 2 ID column is selected
        if not self.dataset_2_id_col.get() or self.dataset_2_id_col.get() == "ID Column":
            self.show_error("Please select an ID column for Dataset 2.")
            return False
    
        # Check if Dataset 1 Match column is selected
        if not self.dataset_1_match_col_1.get() or self.dataset_1_match_col_1.get() == "Match Column":
            self.show_error("Please select a match column for Dataset 1.")
            return False
    
        # Check if Dataset 2 Match column is selected
        if not self.dataset_2_match_col_1.get() or self.dataset_2_match_col_1.get() == "Match Column":
            self.show_error("Please select a match column for Dataset 2.")
            return False
    
        # Assert no duplicate columns selected
        selected_columns = {
            self.dataset_1_match_col_1.get(),
            self.dataset_1_id_col.get(),
            self.dataset_2_match_col_1.get(),
            self.dataset_2_id_col.get(),
        }
        if len(selected_columns) < 4:
            self.show_error("Please ensure all primary columns are distinct.")
            return False

                
        # Checks for threshold matching:
        if self.output_type_var.get():
            try:
                threshold_value = float(self.score_threshold_spinbox.get())

                # Check threshold is between 1 and 100
                if not (0 <= threshold_value <= 100):
                    self.show_error("Threshold value must be between 0 and 100.")
                    return False
                    
            except (ValueError, TypeError):

                # Check threshold is a valid number
                self.show_error("Please enter a valid number for the threshold.")
                return False
        
        # Checks for multi matching.
        if self.multi_match_var.get():
             
             # Check if Dataset 1 Match column 2 is selected.
             if not self.dataset_1_match_col_2.get() or self.dataset_1_match_col_2.get() == "Match Column 2":
                self.show_error("Please select the second match column for Dataset 1.")
                return False
             
            # Check if Dataset 2 Match column 2 is selected.
             if not self.dataset_2_match_col_2.get() or self.dataset_2_match_col_2.get() == "Match Column 2":
                self.show_error("Please select the second match column for Dataset 2.")
                return False

            # Check if combination method selected.
             if not self.score_method_var.get() or self.score_method_var.get() == "Score Method":
                self.show_error("Please select a method for combining match scores.")
                return False
             
             # Check if additional match columns are distinct from all the others.
             selected_columns.update({
                self.dataset_1_match_col_2.get(),
                self.dataset_2_match_col_2.get(),
            })
             if len(selected_columns) < 6:
                self.show_error("Please ensure all secondary columns are distinct.")
                return False

        return True

    def load_dataset(self, dataset_path, id_col, match_col_1, multi_match, match_col_2 = "" ):
        '''
        Loads dataset from path and returns the columns relevant for matching

        Parameters
        ----------
        dataset_path
            Path to dataset
        id_col
            ID column of the given dataset
        match_col_1
            Primary match column of the given dataset
        multi_match
            Whether multi match is enabled, should another column be read in?
        match_col_2, optional
            Secondary match column of the given dataset if required, by default ""

        Returns
        -------
            The dataframe, the number of rows, and the relevant columns.
            Also contains all other columns not relevant for matching to include if desired.
        '''

        # Read in data
        if dataset_path.endswith('.xlsx'):
            df = pd.read_excel(dataset_path)
        elif dataset_path.endswith('.csv'):
            df = pd.read_csv(dataset_path)
        elif dataset_path.endswith('.dta'):
            df = pd.read_stata(dataset_path)
        else:
            self.show_error(f"Unsupported file format for {dataset_path}.")
            return None, None, None, None

        # Retrieve cached row data
        rows, _ = self.dataset_cache[dataset_path]
        id_column = id_col.get()

        # Convert all match columns to strings
        df[match_col_1] = df[match_col_1].astype(str)
        if multi_match and match_col_2:
            df[match_col_2] = df[match_col_2].astype(str)
        
        # Attempt ascii conversion if toggled
        if self.ascii_convert_switch.get():
            df[match_col_1] = df[match_col_1].apply(unidecode)
            if multi_match and match_col_2:
                df[match_col_2] = df[match_col_2].apply(unidecode)

        # Isid: Check if id_column is unique
        if df[id_column].duplicated().any():
            self.show_error(f"Error: The ID column '{id_column}' contains duplicates.")
            return None, None, None, None

        # Define columns to exclude from 'other'
        exclude_cols = {match_col_1, match_col_2} if multi_match else {match_col_1}
        other_cols = [col for col in df.columns if col not in exclude_cols]

        self.debug_message(f"Dataset {dataset_path} loaded")
        return df, rows, id_column, match_col_1, match_col_2 if multi_match else None, other_cols
        
    def get_scorer(self):
        '''
        Helper function to return the desired matching algorithm.

        Returns
        -------
            RF Scorer object corresponding to the algorithm specified
        '''
        selected_matching_type = self.matching_type_var.get()
        if selected_matching_type == 1:
            return rf.fuzz.token_set_ratio
        elif selected_matching_type == 2:
            return rf.fuzz.token_sort_ratio
        elif selected_matching_type == 3:
            return rf.fuzz.token_ratio
        elif selected_matching_type == 4:
            return rf.fuzz.QRatio

    def setup_tasks(self, dataset_1_rows):
        '''
        Compute the total number of matching operations for the progress bar.

        Parameters
        ----------
        dataset_1_rows
            Rows in dataset 1

        Returns
        -------
            Total number of tasks
            Threshold to update the progress bar (every 5% of the total)
            Desired output type (All combos, threshold, best)
        '''
        selected_output_type = self.output_type_var.get()
        total_tasks = dataset_1_rows
        update_threshold = round(total_tasks, -1) * 0.1

        return total_tasks, update_threshold, selected_output_type

    def update_progress(self, update_threshold, total_tasks):
        '''
        Put progress to the progress queue

        Parameters
        ----------
        update_threshold
            Cutoff for updating progress
        total_tasks
            Total number of tasks to perform
        '''
        # Track progress
        self.current_progress += 1
        # Update progress if at the threshold, or the task has completed
        if self.current_progress % update_threshold == 0 or self.current_progress == round(total_tasks,-1) or self.current_progress == total_tasks:
            # Put progress to the queue
            self.progress_queue.put((total_tasks, self.current_progress))


    def generate_matches(self, selected_output_type, dataset_1_df, dataset_2_df,
                         match_col_1, match_col_2, id_col_1, id_col_2, scorer,
                         total_tasks, update_threshold):
        '''
        Performs the matching operation.
        Updates progress bar after every iteration through dataset 1. 

        Parameters
        ----------
        selected_output_type
            Desired output structure: All combinations, best combinations, above threshold.
        dataset_1_df
            DataFrame of dataset 1
        dataset_2_df
            DataFrame of dataset 2
        match_col_1
            Match column from dataset 1
        match_col_2
            Match column from dataset 2
        id_col_1
            ID column from dataset 1
        id_col_2
            ID column from dataset 2
        scorer
            RapidFuzz scorer object: E.G. QRatio, Set Ratio
        total_tasks
            Total taks for display on progres bar
        update_threshold
            Threshold to update progress bar

        Returns
        -------
            A list of lists containing the row by row matching results.
        '''

        # Initialise list to store output
        data = []
        
        # 1 - All possible combinations
        if selected_output_type == 1: 

            # Loop over all combinations of entries in datsets 1 and 2
            for i in range(len(dataset_1_df)):
                for j in range(len(dataset_2_df)):

                    # Calculate match score of (i,j) and append to data
                    score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])
                    data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[j], 
                                 dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j], score])
                    
                # Update progress
                self.update_progress(update_threshold, total_tasks)
        
        # 2 - Highest matches only
        elif selected_output_type == 2: 

            # Loop over all entries in dataset 1
            for i in range(len(dataset_1_df)):

                # Initialise max score and best match variables to be updated through loop
                # Note: max_score set to -1 so even if all scores are zero all rows from df 1 still appear
                max_score, best_match = -1, None

                # Loop over rows in dataset 2
                for j in range(len(dataset_2_df)):

                    # Calculate match score of (i,j)
                    score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])
                    
                    # Replace max_score and best match index
                    if score > max_score:
                        max_score = score
                        best_match = j
                
                # Append only the best match to results 
                data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[best_match], 
                                dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[best_match], max_score])
                
                # Update progress
                self.update_progress(update_threshold, total_tasks)
        
        # 3 - Matches above threshold 
        elif selected_output_type == 3:

            # Get threshold value
            threshold_value = float(self.score_threshold_spinbox.get())

            # Loop over rows in dataset 1
            for i in range(len(dataset_1_df)):

                # Extract all scores above the specified threshold
                results = rf.process.extract(
                    dataset_1_df[match_col_1].iloc[i],
                    dataset_2_df[match_col_2].tolist(),
                    scorer=scorer,
                    score_cutoff=threshold_value,
                    limit=None
                )
                
                # Extract results and store in the data list
                for result in results:
                    match, score, _ = result
                    matched_row = dataset_2_df[dataset_2_df[match_col_2] == match].iloc[0]
                    data.append([dataset_1_df[id_col_1].iloc[i], matched_row[id_col_2], 
                                 dataset_1_df[match_col_1].iloc[i], matched_row[match_col_2], score])
                    
                # Update progress
                self.update_progress(update_threshold, total_tasks)

        # If any rows are missing a match variable set the score to zero
        for row in data:
            if pd.isna(row[2]) or pd.isna(row[3]):
                row[4] = 0

        # Return a list of lists
        return data
    
    def multi_match(self, selected_output_type, dataset_1_df, dataset_2_df, match_columns_1, match_columns_2, 
                 id_col_1, id_col_2, scorer, total_tasks, update_threshold, combination_method, score_1_weight):
        '''
        Performs the matching operation across multiple columns.
        Aggregates the results using the specified method.

        Parameters
        ----------
        selected_output_type
            Desired output structure
            Takes values: All combinations, best combinations, above threshold
        dataset_1_df
            DataFrame of dataset 1
        dataset_2_df
            DataFrame of dataset 2
        match_col_1
            Match column from dataset 1
        match_col_2
            Match column from dataset 2
        id_col_1
            ID column from dataset 1
        id_col_2
            ID column from dataset 2
        scorer
            RapidFuzz scorer object
            E.G.: QRatio, Set Ratio
        total_tasks
            Total taks for display on progres bar
        update_threshold
            Threshold to update progress bar
        combination_method
            Method to combine scores from multiple matching operations
            Takes values: Maximum, Minimum, Weighted Average
        score_1_weight
            Weight placed on score 1 when the weighted average option is selected

        Returns
        -------
            A list of lists 
        '''
        # Initialise list to store result dataframes in
        results = []
        
        # Loop over entries in match columns
        for index in range(len(match_columns_1)):
            match_col_1 = match_columns_1[index]
            match_col_2 = match_columns_2[index]
            
            # Get list of lists for each set of match columns
            data = self.generate_matches(selected_output_type, dataset_1_df, 
                                         dataset_2_df, match_col_1, match_col_2, 
                                         id_col_1, id_col_2, scorer, total_tasks,
                                         update_threshold)
            
            # Convert to a pandas df for easy merge
            df = pd.DataFrame(data,
                              columns=[id_col_1, id_col_2, match_col_1, match_col_2, f"score_{index+1}"])
            # Append the results to the master 
            results.append(df)
            self.debug_message(f"Matching on column {index+1} completed")
        
        # Intialise the result df with the first set of results
        final_df = results[0]

        # For each other set of results, merge on id columns
        for df in results[1:]: # Currently results is length 2 since only 2 column matching implemented
            final_df = final_df.merge(df, on=[id_col_1, id_col_2], how="outer")
        
        # Extract a list of all the score columns from each match set
        score_columns = [col for col in final_df.columns if col.startswith('score_')]
        
        # Combine the scores from each match into one with the defined method
        if combination_method == 'Maximum':
            # Take the max of all scores, concept equivalent to logical OR
            final_df['score'] = final_df[score_columns].max(axis=1)
        elif combination_method == 'Minimum':
            # Take the min of all scores, concept equivalent to logical AND
            final_df['score'] = final_df[score_columns].min(axis=1)
        elif combination_method == 'Weighted Average':
            # Take the weighted average of scores, using the specified weight on score 1
            final_df['score'] = score_1_weight * final_df['score_1'] + (1- score_1_weight) * final_df['score_2']
        
        # Convert output dataframe back to a list to apply consistent formatting as standard matching
        output_data = final_df.values.tolist()
        return output_data

    def clean_data(self, result_df, id_col_1):
        '''
        Sorts result_df by match score highest to lowest within dataset 1 ID.
        Optionally returns comment and match variables if specified.

        Parameters
        ----------
        result_df
            The dataframe to be cleaned
        id_col_1
            The ID variable from result_df within which to group results

        Returns
        -------
            _description_
        '''
        # Group by the dataset 1 ID variable
        result_df['group_id'] = result_df.groupby([id_col_1], sort=False).ngroup() + 1

        # Sort within group on match score
        result_df.sort_values(by=['group_id', 'Match Score'], ascending=[True, False], inplace=True)
        result_df.drop(columns = ['group_id'], inplace = True)

        # If specified, add columns to aid manual checks
        if self.clean_switch.get() == 1:
            result_df['Valid Match'] = 0
            result_df['Comments'] = ''

        self.debug_message("Output cleaned")
        return result_df  
    
    def save_data(self, result_df):
        '''
        Write the output to the desired file.

        Parameters
        ----------
        result_df
            The dataframe to be output
        '''

        output_file = self.output_path.get()
    
        try:
            if output_file.endswith('.xlsx'):
                result_df.to_excel(output_file, index=False, engine="xlsxwriter")

            elif output_file.endswith('.csv'):
                # Set utf-8-sig to display non unicode characters in excel
                result_df.to_csv(output_file, index=False, encoding='utf-8-sig')

            elif output_file.endswith('.dta'):
                result_df.to_stata(output_file, write_index=False)
    
            # Add an animal fact if fact_switch is set
            animal_fact = "\n\n" + random.choice(facts) if self.fact_switch_flag.get() == 1 else ""
            messagebox.showinfo("Success", "Matching completed and saved to " + output_file + animal_fact)

            self.debug_message(f"Output saved to {output_file}")

        except PermissionError:
            self.show_error("Write permission denied. Is the output file open?")
            self.run_matching_button.configure(state="normal")

    def run_matching(self):
        '''
        Primary function to execute when the run button is clicked.
        Loads, matches, cleans and saves data.
        '''

        if not self.validate_inputs():
            return
        
        multi_match_flag = self.multi_match_switch.get()

        # Load datasets
        dataset_1_df, dataset_1_rows, dataset_1_id_col, dataset_1_match_col_1, dataset_1_match_col_2, dataset_1_other_cols = self.load_dataset(
                    self.dataset_1_path.get(), self.dataset_1_id_col,
                    self.dataset_1_match_col_1.get(), multi_match_flag,
                    self.dataset_1_match_col_2.get()
                )
        dataset_2_df, dataset_2_rows, dataset_2_id_col, dataset_2_match_col_1, dataset_2_match_col_2, dataset_2_other_cols = self.load_dataset(
                    self.dataset_2_path.get(), self.dataset_2_id_col,
                    self.dataset_2_match_col_1.get(), multi_match_flag,
                    self.dataset_2_match_col_2.get()
                )
        
        # If dataset is too large, export to csv
        if dataset_1_rows * dataset_2_rows > 100000 and not self.output_path.get().endswith('.csv'):
            self.show_error("Too much data for this format, please export to a CSV.")
            self.run_matching_button.configure(state="normal")
            return
        
        # Set up tasks and threshold
        total_tasks, update_threshold, selected_output_type = self.setup_tasks(dataset_1_rows)
    
        # Get scorer function
        scorer = self.get_scorer()

        # Setup for matching
        self.current_progress = 0
        # Use a queue here to avoid interfering with the GUI via a worker thread.
        # GUI can only be interacted with via the main thread, so we put the progress to a queue.
        # The main thread then reads from this queue object and updates the progress bar. 
        self.progress_queue = queue.Queue()
        self.debug_message("Progress queue initialised")

        # Disable run button when matching begins
        self.run_matching_button.configure(state="disabled")

        def run_in_thread():
            '''
            Function to execute within the worker thread, performs the actual matching
            '''
            if multi_match_flag:
                data = self.multi_match(
                    selected_output_type, dataset_1_df, dataset_2_df, [dataset_1_match_col_1, dataset_1_match_col_2],
                      [dataset_2_match_col_1, dataset_2_match_col_2], dataset_1_id_col, dataset_2_id_col, scorer, total_tasks*2,
                        update_threshold, self.score_method_var.get(), self.weight_var.get())

                column_list = [dataset_1_id_col,
                               dataset_2_id_col,
                               dataset_1_match_col_1,
                               dataset_2_match_col_1,
                               'Score 1',
                               dataset_1_match_col_2,
                               dataset_2_match_col_2,
                               'Score 2',
                               'Match Score']

            else:
                data = self.generate_matches(
                    selected_output_type, dataset_1_df, dataset_2_df, dataset_1_match_col_1, dataset_2_match_col_1,
                    dataset_1_id_col, dataset_2_id_col, scorer, total_tasks, update_threshold
                )

                column_list = [dataset_1_id_col,
                               dataset_2_id_col,
                               dataset_1_match_col_1,
                               dataset_2_match_col_1,
                               'Match Score']

            # Save result_df, and tell the queue execution has fnished
            self.result_df = pd.DataFrame(data, columns=column_list)
            self.progress_queue.put(("result", None))

        # Initialise the worker thread
        matching_thread = threading.Thread(
            # Use daemon to ensure all threads are killed when the app closes
            target = run_in_thread, daemon = True
        )
        self.debug_message('Worker thread created')

        matching_thread.start()
        self.debug_message('Matching started')

        def check_progress():
            '''
            Monitors progress by retrieving updates from the queue
            '''
            try:
                while True:
                    # Retrieve next update without blocking
                    update = self.progress_queue.get_nowait()

                    if update[0] == "result":
                        # Matching complete
                        on_result_ready()
                        return
                    
                    else:
                        # Update progress if ongoing
                        total, completed = update
                        self.update_progress_bar(total, completed)
            # If no updates available, check again in 100ms
            except queue.Empty:
                self.root.after(100, check_progress)

        def on_result_ready():
            '''
            Clean and save the data when queue reports the thread has finished
            '''
            # Retrieve final df
            self.debug_message('Matching completed')
            result_df = self.result_df

            # Keep additional columns if specified
            if self.keep_columns_switch.get() == 1:
                for df, cols, id_col in [(dataset_1_df, dataset_1_other_cols, dataset_1_id_col),
                                        (dataset_2_df, dataset_2_other_cols, dataset_2_id_col)]:
                    result_df = pd.merge(result_df, df[cols], how='left', on=id_col)
            result_df = self.clean_data(result_df, dataset_1_id_col)
            # Save data
            self.save_data(result_df)
            self.run_matching_button.configure(state="normal")

        check_progress()

    # 4. **Utility Functions**
    def update_progress_bar(self, total_tasks, current_progress):
        '''
        Update progress bar when updates are given.
        '''
        self.progress_bar.set(current_progress / total_tasks)
        self.progress_label.configure(text=f"Progress: {current_progress}/{total_tasks}")
        self.root.update_idletasks()

    def show_error(self, message):
        '''
        Display an error popup with specified message.
        Report error in the debug window also
        '''
        messagebox.showerror("Error", message)
        self.debug_message(message)

    def debug_message(self,message):
        '''
        Display text with timestamp in the debug window
        '''
        c = datetime.now()
        current_time = c.strftime('%H:%M:%S')

        print(current_time + " " + message)


# %% Execute

# Define the list of fun animal facts to include on completion.
facts = [
            "Rabbits don't have pads on their paws, only fur. So if you see a cartoon rabbit with pads on it's paw, completely wrong.",
            "Fuzzy animals are very cute - Chiara", 
            "I know a fact but it's in the back of my head - Owen",
            "Penguins have a gland above their eye that converts saltwater into freshwater.",
            "Sloths are literally too lazy to go looking for a mate, so a female sloth will often sit in a tree and scream until a male hears her and decides to mate with her.",
            "The Western Lowland Gorilla's scientific name is 'gorilla gorilla gorilla'",
            "Opossums have a body temperature so low that is makes them highly resistant to rabies.",
            "Male giraffes will headbutt a female in the bladder until she urinates, then it tastes the pee to help it determine whether or not the female is ovulating.",
            "Platypus' glow teal under a UV light, so perry the Platypus is (kind of) the correct color.", 
            "Cats can jump about seven times their height.",
            "Sea lions are the first non-human animals to be able to keep a beat.",
            "Reindeer eyeballs turn blue in winter to help them see at lower light levels.",
            "Young goats pick up accents from each other.",
            "Horses use facial expressions to communicate with each other.",
            "African buffalo herds display voting behaviour, in which they register their direction of travel preference."
        ]

# Entry point
if __name__ == "__main__":
    MatchingTool()
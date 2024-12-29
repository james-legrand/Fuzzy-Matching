import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import rapidfuzz as rf
import os
import random
import threading
from typing import Callable, Union
from matplotlib import colors as mcolors
import pyarrow as pa
import pyarrow.csv as csv
import queue
import sys
import io

class TextRedirector(io.StringIO):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget

    def write(self, string):
        if string.strip():  # Only log non-empty lines
            self.widget.after(0, self.widget.insert, "end", string + "\n")
            self.widget.after(0, self.widget.see, "end")

    def flush(self):
        pass  # No buffering is required



class IntSpinbox(ctk.CTkFrame):
    def __init__(self, *args,
                 width: int = 150,
                 height: int = 32,
                 step_size: int = 1,
                 command: Callable = None,
                 **kwargs):
        super().__init__(*args, width=width, height=height, **kwargs)

        self.step_size = step_size
        self.command = command

        self.configure(fg_color=["gray92", "gray14"])  # set frame color

        self.grid_columnconfigure((0, 2), weight=0)  # buttons don't expand
        self.grid_columnconfigure(1, weight=1)  # entry expands

        self.subtract_button = ctk.CTkButton(self, text="-", width=height-6, height=height-6,
                                             command=self.subtract_button_callback)
        self.subtract_button.grid(row=0, column=0, padx=(3, 0), pady=3)

        self.entry = ctk.CTkEntry(self, width=width-(2*height), height=height-6, border_width=0)
        self.entry.grid(row=0, column=1, columnspan=1, padx=3, pady=3, sticky="ew")

        self.add_button = ctk.CTkButton(self, text="+", width=height-6, height=height-6,
                                        command=self.add_button_callback)
        self.add_button.grid(row=0, column=2, padx=(0, 3), pady=3)

        # default value
        self.entry.insert(0, "0")

    def add_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = int(self.entry.get()) + self.step_size
            if 0 <= value <= 100:  # Ensure value remains between 0 and 100
                self.entry.delete(0, "end")
                self.entry.insert(0, value)
        except ValueError:
            return

    def subtract_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = int(self.entry.get()) - self.step_size
            if 0 <= value <= 100:  # Ensure value remains between 0 and 100
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

class MatchingTool:
    
    def __init__(self):
        # Initialize main root window first
        self.root = None
        self.is_advanced_visible = False  # Track whether advanced options are visible
        # Call function to set up GUI
        self.setup_gui()

    def __del__(self):
    # Reset stdout to default
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__


    # 1. **Setup GUI**
    def setup_gui(self):
        self.root = ctk.CTk()
        self.root.title("Fuzzy Matching Tool")
        self.root.geometry("550x600")

        # Initialize variables 
        self.dataset_1_path = ctk.StringVar()
        self.dataset_2_path = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.column1 = ctk.StringVar()
        self.column2 = ctk.StringVar()
        self.dataset_1_id_col = ctk.StringVar(value='ID Column')
        self.dataset_1_match_col = ctk.StringVar(value='Match Column')
        self.dataset_2_id_col = ctk.StringVar(value='ID Column')
        self.dataset_2_match_col = ctk.StringVar(value='Match Column')
        self.dataset_1_match_col_2 = ctk.StringVar(value='Match Column 2')
        self.dataset_2_match_col_2 = ctk.StringVar(value='Match Column 2')
        self.output_type_var = ctk.IntVar(value=2)
        self.matching_type_var = ctk.IntVar(value=1)  # Default to 'Set Ratio'
        self.score_method_var = ctk.StringVar(value="Score Method")
        self.weight_var = ctk.DoubleVar(value = 0.5)
        self.dataset_cache = {}
        
        # Set appearance and theme
        ctk.set_appearance_mode("dark")
        #ctk.set_default_color_theme("rbb_theme.json")

        # Main title
        ctk.CTkLabel(self.root, text="Fuzzy Matching Tool", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)

        # Main frame
        self.main_frame = ctk.CTkFrame(self.root, width=600, height=600, fg_color=["gray92", "gray14"])
        self.main_frame.place(x=20, y=50, anchor="nw")

    
        self.create_dataset_frame(self.main_frame, 1, self.dataset_1_path, self.dataset_1_id_col, self.dataset_1_match_col)
        self.create_dataset_frame(self.main_frame, 2, self.dataset_2_path, self.dataset_2_id_col, self.dataset_2_match_col)
        
        # Output file selection button
        self.create_button(self.main_frame, "Select Output File", self.output_path, is_output=True)
    
        # Radio buttons for selecting output type
        ctk.CTkLabel(self.main_frame, text="Select Output Type").pack(pady=5)
        ctk.CTkRadioButton(self.main_frame, text="All Possible Combinations", variable=self.output_type_var, value=1).pack()
        ctk.CTkRadioButton(self.main_frame, text="Highest Matches Only", variable=self.output_type_var, value=2).pack()
    
        threshold_frame = ctk.CTkFrame(self.main_frame, fg_color = ["gray92", "gray14"])
        threshold_frame.pack(pady=5)
    
        threshold_button = ctk.CTkRadioButton(
            threshold_frame, 
            text="Matches Above Threshold", 
            variable=self.output_type_var, 
            value=3
        )
        threshold_button.grid(row=0, column=0, padx=5)
    
        self.score_threshold_spinbox = IntSpinbox(threshold_frame, width=100, step_size=1)
        self.score_threshold_spinbox.grid(row=0, column=1, padx=5, pady=5)
        self.score_threshold_spinbox.set(80)
    
       # Appearance mode switch at the top-left corner
        self.appearance_mode_switch = ctk.CTkSwitch(
            self.root,
            text="Theme",
            command=self.toggle_appearance_mode
        )
        self.appearance_mode_switch.place(x=10, y=5, anchor="nw")  # Adjust to top-left

        # Advanced options toggle button below the appearance mode switch
        self.toggle_advanced_button = ctk.CTkSwitch(
            self.root,
            text="Advanced Options",
            command=self.toggle_advanced_options
        )
        self.toggle_advanced_button.place(x=10, y=30, anchor="nw")  # Slightly below the switch

        # Create the advanced options frame (hidden by default)
        self.advanced_options_frame = ctk.CTkFrame(self.root, width=500, height=500, fg_color=["gray92", "gray14"])
        self.advanced_options_frame.grid_columnconfigure(1, weight=1)  # Make columns flexible
        self.advanced_options_frame.grid_columnconfigure(2, weight=1)
        self.advanced_options_frame.grid_columnconfigure(3, weight=1)
        self.advanced_options_frame.place_forget()  # Hide by default

        # Add debugging window contents to the frame
        self.terminal_output_text = ctk.CTkTextbox(self.advanced_options_frame, width=450, height=250, wrap="word")

        # Redirect stdout and stderr to the text box
        redirector = TextRedirector(self.terminal_output_text)
        sys.stdout = redirector
        sys.stderr = redirector  # Capture errors   

        
        # Radio buttons for selecting matching type
        ctk.CTkLabel(self.main_frame, text="Select Matching Type").pack(pady=5)
        
        ctk.CTkRadioButton(self.main_frame, text="Set Ratio", variable=self.matching_type_var, value=1).pack()
        ctk.CTkRadioButton(self.main_frame, text="Sort Ratio", variable=self.matching_type_var, value=2).pack()
        ctk.CTkRadioButton(self.main_frame, text="Max of (Set Ratio, Sort Ratio)", variable=self.matching_type_var, value=3).pack()
        ctk.CTkRadioButton(self.main_frame, text="QRatio", variable=self.matching_type_var, value=4).pack()
    
        # Progress bar and label
        self.progress_bar = ctk.CTkProgressBar(self.main_frame, width=300)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)
    
        self.progress_label = ctk.CTkLabel(self.main_frame, text="Progress: 0/0")
        self.progress_label.pack(pady=5)

        # Row 1: Radio buttons (3 in a row)
        self.multi_match_switch = ctk.CTkSwitch(self.advanced_options_frame, text="2 Column Match")
        self.dataset_1_match_2_dropdown = ctk.CTkOptionMenu(self.advanced_options_frame, variable= self.dataset_1_match_col_2, values=["Match Column 2"])
        self.dataset_2_match_2_dropdown = ctk.CTkOptionMenu(self.advanced_options_frame, variable= self.dataset_2_match_col_2, values=["Match Column 2"])

        # Row 2: Radio buttons (3 in a row)
        self.combine_score_dropdown = ctk.CTkOptionMenu(self.advanced_options_frame, variable=self.score_method_var, values=["Maxium", "Minimum", "Weighted Average"])
        self.weight_1_slider = ctk.CTkSlider(self.advanced_options_frame, from_=0, to = 1, variable = self.weight_var, width = 200)
        self.slider_label = ctk.CTkLabel(self.advanced_options_frame, text="Score 1 weight")

        # Rows 3-5: Switches
        self.fact_switch = ctk.CTkSwitch(self.advanced_options_frame, text="Fuzzy animal fact")
        self.clean_switch = ctk.CTkSwitch(self.advanced_options_frame, text="Prep for manual checks")
        self.keep_columns_switch = ctk.CTkSwitch(self.advanced_options_frame, text="Keep all columns")



        # Run matching button
        self.run_matching_button = ctk.CTkButton(self.main_frame, text="Run Matching", command=self.run_matching, width=200)
        self.run_matching_button.pack(pady=10)
        
        # Start the GUI loop
        self.root.mainloop()

    def toggle_advanced_options(self):
        if self.is_advanced_visible:
            # Hide advanced options and resize window back
            self.root.geometry("550x600")
            self.advanced_options_frame.place_forget()
        else:
            # Show advanced options and expand window
            self.root.geometry("1050x600")
            self.advanced_options_frame.place(relx=0.97, y=50,  anchor = "ne")  # Adjust placement for advanced options

            self.terminal_output_text.grid(row = 0, column = 0, columnspan = 3, pady = 10)
            self.multi_match_switch.grid(row=1, column=0, padx=5, pady=5, sticky="w")
            self.dataset_1_match_2_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="w")
            self.dataset_2_match_2_dropdown.grid(row=1, column=2, padx=5, pady=5, sticky="w")

            self.combine_score_dropdown.grid(row=2, column=0, padx=5, pady=5, sticky="w")
            self.weight_1_slider.grid(row=2, column=1, padx=5, pady=5, columnspan=2,  sticky="w")  # Slider spans 2 columns
            self.slider_label.grid(row=2, column=2, padx=5, pady=5, sticky="e")  # Place label to the left of the slider



            self.fact_switch.grid(row=3, column=0, columnspan=3, pady=10)
            self.clean_switch.grid(row=4, column=0, columnspan=3, pady=10)
            self.keep_columns_switch.grid(row=5, column=0, columnspan=3, pady=10)

        self.is_advanced_visible = not self.is_advanced_visible

    def create_radio_buttons(self, frame, label_text, variable, options, default_value):
        """Helper function to create a label and radio buttons."""
        ctk.CTkLabel(frame, text=label_text).pack(pady=5)
        variable.set(default_value)
        for text, value in options:
            ctk.CTkRadioButton(frame, text=text, variable=variable, value=value).pack()

        
    def create_button(self, frame, text, var, is_output=False):
        button = ctk.CTkButton(frame, text=text, command=lambda: self.browse_file(var, button, is_output), width=200)
        button.pack(pady=10)
        return button
    
    def create_dataset_frame(self,frame, dataset_num, path_variable, id_variable, match_variable):
        # Create the frame for the dataset
        dataset_frame = ctk.CTkFrame(frame, fg_color=["gray92", "gray14"])
        dataset_frame.pack(pady=10)
    
        # Create the button for selecting the dataset
        dataset_button = ctk.CTkButton(
            dataset_frame,
            text=f"Select Dataset {dataset_num}",
            command=lambda: self.browse_file(path_variable, dataset_button),
            width=200
        )
        dataset_button.grid(row=0, column=0, padx=5)
    
        # Create the dropdown for the ID column
        id_dropdown = ctk.CTkOptionMenu(
            dataset_frame,
            variable=id_variable,
            values=["ID Column"]
        )
        id_dropdown.grid(row=0, column=1, padx=5)
    
        # Create the dropdown for the Match column
        match_dropdown = ctk.CTkOptionMenu(
            dataset_frame,
            variable=match_variable,
            values=["Match Column"]
        )
        match_dropdown.grid(row=0, column=2, padx=5)
    
        # Assign the dropdowns to the instance attributes
        if dataset_num == 1:
            self.dataset_1_id_dropdown = id_dropdown
            self.dataset_1_match_dropdown = match_dropdown
        else:
            self.dataset_2_id_dropdown = id_dropdown
            self.dataset_2_match_dropdown = match_dropdown
    
        # Bind hover events to display dataset info
        dataset_button.bind("<Enter>", lambda event: self.show_dataset_info(path_variable, event, dataset_button))
    
        # Reset button text when mouse leaves the button
        dataset_button.bind("<Leave>", lambda event: dataset_button.configure(
            text=os.path.basename(path_variable.get()) if path_variable.get() else f"Select Dataset {dataset_num}")
        )



    
    def toggle_appearance_mode(self):
        if self.appearance_mode_switch.get():
            ctk.set_appearance_mode("light")  # Switch to light mode
        else:
            ctk.set_appearance_mode("dark")   # Switch to dark mode

    

   ### File Handling
    def browse_file(self, var, label, is_output=False):
        file_path = filedialog.askopenfilename(filetypes=[("Data file", "*.csv"), ("Data file", "*.xlsx"), ("Data file", "*.dta")])
        if file_path:
            var.set(file_path)
            label.configure(text=os.path.basename(file_path), text_color="white")
            
            if not is_output:
                # Cache dataset dimensions when imported
                if file_path not in self.dataset_cache:
                    if file_path.endswith('.xlsx'):
                        df = pd.read_excel(file_path)
                    elif file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                    elif file_path.endswith('.dta'):
                        df = pd.read_stata(file_path)
                    self.dataset_cache[file_path] = df.shape  # Cache the shape (rows, columns)
                
                if var in [self.dataset_1_path, self.dataset_2_path]:
                    self.update_dropdowns(
                        var, 
                        self.dataset_1_id_dropdown if var == self.dataset_1_path else self.dataset_2_id_dropdown, 
                        self.dataset_1_match_dropdown if var == self.dataset_1_path else self.dataset_2_match_dropdown
                    )


    def show_dataset_info(self, var, event, label):
        file_path = var.get()
        if file_path:
            if file_path in self.dataset_cache:
                rows, cols = self.dataset_cache[file_path]  # Retrieve cached dimensions
                label.configure(text=f"{cols} columns, {rows} rows")
    
    def update_dropdowns(self, var, dropdown_id, dropdown_match):
        # Clear the dropdowns
        dropdown_id.set("ID Column")
        dropdown_match.set("Match Column")
    
        if var.get():
            if var.get().endswith('.xlsx'):
                df = pd.read_excel(var.get())
            elif var.get().endswith('.csv'):
                df = pd.read_csv(var.get())
            elif var.get().endswith('.dta'):
                df = pd.read_stata(var.get())
            columns = df.columns.tolist()
            dropdown_id.configure(values=columns)
            dropdown_match.configure(values=columns)
            
    def validate_inputs(self):
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
        if not self.dataset_1_match_col.get() or self.dataset_1_match_col.get() == "Match Column":
            self.show_error("Please select a match column for Dataset 1.")
            return False
    
        # Check if Dataset 2 Match column is selected
        if not self.dataset_2_match_col.get() or self.dataset_2_match_col.get() == "Match Column":
            self.show_error("Please select a match column for Dataset 2.")
            return False
    
        # Check if ID and Match columns are distinct
        if self.dataset_1_match_col.get() == self.dataset_1_id_col.get() or self.dataset_2_match_col.get() == self.dataset_2_id_col.get():
            self.show_error("Please select distinct columns to identify and match on.")
            return False

        # Check if ID and match column names are distinct
        if self.dataset_1_id_col.get() == self.dataset_2_id_col.get() or self.dataset_1_match_col.get() == self.dataset_2_match_col.get():
            self.show_error("Please rename ID or match columns - duplicate column names found.")
            return False
    
        return True
    
    def load_dataset(self, dataset_path, id_col, match_col):
        """Load dataset based on file extension, retrieve cached row data, and select specified columns."""
        # Determine the file extension and load accordingly
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
    
        # Resolve selected columns
        id_column = id_col.get()
        match_column = match_col.get()
        other_cols = [col for col in df.columns if col != match_column]
    
        return df, rows, id_column, match_column, other_cols
    
    def get_scorer(self):
            """Return the appropriate scorer function based on the selected matching type."""
            selected_matching_type = self.matching_type_var.get()
            if selected_matching_type == 1:
                return rf.fuzz.token_set_ratio
            elif selected_matching_type == 2:
                return rf.fuzz.token_sort_ratio
            elif selected_matching_type == 3:
                return rf.fuzz.token_ratio
            elif selected_matching_type == 4:
                return rf.fuzz.QRatio
            else:
                self.show_error("Invalid matching type selected.")
                return None
    
    def setup_tasks(self, dataset_1_rows, dataset_2_rows):
        """Calculate and return total tasks and update threshold based on output type."""
        selected_output_type = self.output_type_var.get()
        total_tasks = dataset_1_rows * dataset_2_rows if selected_output_type == 1 else dataset_1_rows
        update_threshold = round(total_tasks, -1) * 0.1
        return total_tasks, update_threshold, selected_output_type
    

    # 3. **Matching Process Functions**
    def update_progress(self, update_threshold, total_tasks):
        self.current_progress += 1
        if self.current_progress % update_threshold == 0 or self.current_progress == round(total_tasks,-1) or self.current_progress == total_tasks:
            self.progress_queue.put((total_tasks, self.current_progress))
            
    def generate_matches(self, selected_output_type, dataset_1_df, dataset_2_df, match_col_1, match_col_2, 
                     id_col_1, id_col_2, scorer, total_tasks, update_threshold):
        data = []
        
        if selected_output_type == 1:  # All Possible Combinations
            for i in range(len(dataset_1_df)):
                for j in range(len(dataset_2_df)):
                    if pd.isna(dataset_1_df[match_col_1].iloc[i]) or pd.isna(dataset_2_df[match_col_2].iloc[j]):
                        score = 'N/A'
                    else:
                        score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])
                    data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[j], 
                                 dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j], score])
                    self.update_progress(update_threshold, total_tasks)
        
        elif selected_output_type == 2:  # Highest Matches Only
            for i in range(len(dataset_1_df)):
                max_score, best_match = 0, None
                for j in range(len(dataset_2_df)):
                    if pd.isna(dataset_1_df[match_col_1].iloc[i]) or pd.isna(dataset_2_df[match_col_2].iloc[j]):
                        score = 0
                    else:
                        score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])
                    
                    if score > max_score:
                        max_score = score
                        best_match = j
                    
                    data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[j], 
                                 dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j], score])
                
                if best_match is not None:
                    data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[best_match], 
                                 dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[best_match], max_score])
                self.update_progress(update_threshold, total_tasks)
        
        elif selected_output_type == 3:  # Matches above threshold
            # Validate threshold value
            try:
                threshold_value = float(self.score_threshold_spinbox.get())
                if not (0 <= threshold_value <= 100):
                    self.show_error("Threshold value must be between 0 and 100.")
                    self.run_matching_button.configure(state="normal")
                    return False
            except (ValueError, TypeError):
                self.show_error("Please enter a valid number for the threshold.")
                self.run_matching_button.configure(state="normal")
                return False
            
            
            for i in range(len(dataset_1_df)):
                if pd.isna(dataset_1_df[match_col_1].iloc[i]):
                    continue
                
                results = rf.process.extract(
                    dataset_1_df[match_col_1].iloc[i],
                    dataset_2_df[match_col_2].dropna().tolist(),
                    scorer=scorer,
                    score_cutoff=threshold_value,
                    limit=None
                )
                
                for result in results:
                    match, score, _ = result
                    matched_row = dataset_2_df[dataset_2_df[match_col_2] == match].iloc[0]
                    data.append([dataset_1_df[id_col_1].iloc[i], matched_row[id_col_2], 
                                 dataset_1_df[match_col_1].iloc[i], matched_row[match_col_2], score])
                self.update_progress(update_threshold, total_tasks)

        return data

    def clean_data(self, result_df, id_col_1):
        # Ensure columns exist and check for NaN values
        if self.clean_switch.get() == 1:
            result_df['Valid Match'] = 0
            result_df['Comments'] = ''
            result_df['group_id'] = result_df.groupby([id_col_1], sort=False).ngroup() + 1
            result_df['is_highest'] = result_df.groupby('group_id')['Match Score'].transform('max') == result_df['Match Score']
            result_df.sort_values(by=['group_id', 'Match Score'], ascending=[True, False], inplace=True)
            
            # Apply background color formatting for groups
            highlighted_data = []
            current_group_index = -1
            current_group = None
    
            for index, row in result_df.iterrows():
                if row['group_id'] != current_group:
                    current_group_index += 1
                    current_group = row['group_id']
                    
                bg_color = self.get_group_color(current_group_index, row['is_highest'])
                highlighted_data.append({**row.to_dict(), 'Background Color': bg_color})
    
            highlighted_df = pd.DataFrame(highlighted_data)
            return highlighted_df
    
        return result_df  # Return unmodified if no cleaning is needed
    
    def save_data(self, result_df):
        """
        Save the cleaned or non-cleaned dataset to the specified file format.
        """
        output_file = self.output_path.get()
    
        try:
            if self.clean_switch.get() == 0:
                # Save non-cleaned data
                if output_file.endswith('.xlsx'):
                    result_df.to_excel(output_file, index=False, engine="xlsxwriter")
                elif output_file.endswith('.csv'):
                    new_pa_dataframe = pa.Table.from_pandas(result_df)
                    csv.write_csv(new_pa_dataframe, output_file)
                elif output_file.endswith('.dta'):
                    result_df.to_stata(output_file, index=False)
            
            else:
                # Save cleaned data with formatting if clean_switch == 1
                highlighted_df = result_df.drop(columns=['Background Color', 'group_id', 'is_highest'])
                if output_file.endswith('.xlsx'):
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        highlighted_df.to_excel(writer, index=False, sheet_name='Matches')
                        workbook = writer.book
                        worksheet = writer.sheets['Matches']
                        
                        # Apply cell format for background color
                        for i, color in enumerate(result_df['Background Color']):
                            for j in range(len(highlighted_df.columns)):
                                hex_color = mcolors.to_hex(color).replace('#', '')
                                cell_format = workbook.add_format({'bg_color': hex_color})
                                worksheet.write(i + 1, j, highlighted_df.iloc[i, j], cell_format)
                        
                        # Adjust column width
                        for col_num, value in enumerate(highlighted_df.columns):
                            worksheet.set_column(col_num, col_num, 20)
                
                elif output_file.endswith('.csv'):
                    new_pa_dataframe = pa.Table.from_pandas(highlighted_df)
                    csv.write_csv(new_pa_dataframe, output_file)
                
                elif output_file.endswith('.dta'):
                    highlighted_df.to_stata(output_file, index=False)
    
            # Add an animal fact if fact_switch is set
            animal_fact = "\n\n" + random.choice(facts) if self.fact_switch.get() == 1 else ""
            messagebox.showinfo("Success", "Matching completed and saved to " + output_file + animal_fact)
    
        except PermissionError:
            self.show_error("Write permission denied. Please close the output file.")
            self.run_matching_button.configure(state="normal")
    
    def run_matching(self):
        
        if not self.validate_inputs():
            return

        # Load Dataset 1 and Dataset 2 
        dataset_1_df, dataset_1_rows, id_col_1, match_col_1, dataset_1_other_cols = self.load_dataset(
            self.dataset_1_path.get(), self.dataset_1_id_col, self.dataset_1_match_col
        )
        dataset_2_df, dataset_2_rows, id_col_2, match_col_2, dataset_2_other_cols = self.load_dataset(
            self.dataset_2_path.get(), self.dataset_2_id_col, self.dataset_2_match_col
        )
             
        if dataset_1_rows * dataset_2_rows > 100000 and not self.output_path.get().endswith('.csv'):
            self.show_error("Too much data for this format, please export to a CSV.")
            self.run_matching_button.configure(state="normal")
            return
    
        # Set up tasks and threshold
        total_tasks, update_threshold, selected_output_type = self.setup_tasks(dataset_1_rows, dataset_2_rows)
    
        # Get scorer function
        scorer = self.get_scorer()
    
        # Initialize data structures for matching process
        self.current_progress = 0
        self.progress_queue = queue.Queue()
    
        self.run_matching_button.configure(state="disabled")

        def run_in_thread():
            # Run generate_matches and create result_df in a separate thread
            data = self.generate_matches(
                selected_output_type, dataset_1_df, dataset_2_df, match_col_1,
                match_col_2, id_col_1, id_col_2, scorer, total_tasks, update_threshold
            )

            column_list = [id_col_1, id_col_2, match_col_1, match_col_2, 'Match Score']
            self.result_df = pd.DataFrame(data, columns=column_list)

            self.progress_queue.put(("result", None))  # Notify progress checker

        worker_thread = threading.Thread(
            target=run_in_thread, daemon=True
        )
        worker_thread.start()

        def check_progress():
            try:
                while True:
                    update = self.progress_queue.get_nowait()
                    if update[0] == "result":
                        # Matching complete
                        on_result_ready()
                        return
                    else:
                        # Update progress
                        total, completed = update
                        self.update_progress_bar(total, completed)
            except queue.Empty:
                self.root.after(100, check_progress)

        def on_result_ready():
            # This function is triggered after result_df is available
            result_df = self.result_df
            if self.keep_columns_switch.get() == 1:
                for df, cols, id_col in [(dataset_1_df, dataset_1_other_cols, id_col_1),
                                        (dataset_2_df, dataset_2_other_cols, id_col_2)]:
                    result_df = pd.merge(result_df, df[cols], how='left', on=id_col)

            # Clean and save result
            result_df = self.clean_data(result_df, id_col_1)
            self.save_data(result_df)

            # Re-enable button after completion
            self.run_matching_button.configure(state="normal")

        check_progress()


    # 4. **Utility Functions**
    def update_progress_bar(self, total_tasks, current_progress):
        self.progress_bar.set(current_progress / total_tasks)
        self.progress_label.configure(text=f"Progress: {current_progress}/{total_tasks}")
        self.root.update_idletasks()

    def show_error(self, message):
        messagebox.showerror("Error", message)
        print(message)        
        
    def get_group_color(self, group_index, is_highest_score):
        color_name = 'blue' if group_index % 2 == 0 else 'green'
        return group_colors[color_name]['dark'] if is_highest_score else group_colors[color_name]['light']

group_colors = {
        'blue': {'light': "#f6e3de", 'dark': "#f6e3de"},
        'green': {'light': "#def6f5", 'dark': "#def6f5"}
}
facts =["Rabbits don't have pads on their paws, only fur. So if you see a cartoon rabbit with pads on it's paw, completely wrong.",
            "Fuzzy animals are very cute - Chiara", 
            "I know one but it's in the back of my head - Owen",
            "Penguins have a gland above their eye that converts saltwater into freshwater.",
            "Sloths are literally too lazy to go looking for a mate, so a female sloth will often sit in a tree and scream until a male hears her and decides to mate with her.",
            "The Western Lowland Gorilla's scientific name is 'gorilla gorilla gorilla'",
            "Opossums have a body temperature so low that is makes them highly resistant to rabies.",
            "Male giraffes will headbutt a female in the bladder until she urinates, then it tastes the pee to help it determine whether or not the female is ovulating.",
            "Platypus' glow teal under a UV light, so perry the Platypus is actually the correct color.", 
            "Cats can jump about seven times their height.",
            "Sea lions are the first non-human animals to be able to keep a beat.",
            "Reindeer eyeballs turn blue in winter to help them see at lower light levels.",
            "Young goats pick up accents from each other.",
            "Horses use facial expressions to communicate with each other.",
            "African buffalo herds display voting behaviour, in which they register their direction of travel preference."
            ]



if __name__ == "__main__":
    MatchingTool()

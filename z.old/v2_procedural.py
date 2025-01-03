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

ctk.set_appearance_mode("dark")
#ctk.set_default_color_theme("rbb_theme.json")

# Initialize main window
root = ctk.CTk()
root.title("Fuzzy Matching UI")
root.geometry("600x750")

# Global variables
dataset_1_path = ctk.StringVar()
dataset_2_path = ctk.StringVar()
output_path = ctk.StringVar()


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

class IntSpinbox(ctk.CTkFrame):
    def __init__(self, *args,
                 width: int = 100,
                 height: int = 32,
                 step_size: int = 1,
                 command: Callable = None,
                 **kwargs):
        super().__init__(*args, width=width, height=height, **kwargs)

        self.step_size = step_size
        self.command = command

        self.configure(fg_color=("#F5F4EE", "gray14"))  # set frame color

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



dataset_cache = {}

def browse_file(var, label, is_output=False):
    file_path = filedialog.askopenfilename(filetypes=[("Data file", "*.csv"), ("Data file", "*.xlsx"),  ("Data file", "*.dta")])
    if file_path:
        var.set(file_path)
        label.configure(text=os.path.basename(file_path), text_color="white")
        
        if not is_output:
            # Cache dataset dimensions when imported
            if file_path not in dataset_cache:
                if file_path.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                elif file_path.endswith('.dta'):
                    df = pd.read_stata(file_path)
                dataset_cache[file_path] = df.shape  # Cache the shape (rows, columns)
            
            if var in [dataset_1_path, dataset_2_path]:
                update_dropdowns(var, 
                                 dataset_1_id_dropdown if var == dataset_1_path else dataset_2_id_dropdown, 
                                 dataset_1_match_dropdown if var == dataset_1_path else dataset_2_match_dropdown)

def show_dataset_info(var, event, label):
    file_path = var.get()
    if file_path:
        if file_path in dataset_cache:
            rows, cols = dataset_cache[file_path]  # Retrieve cached dimensions
            label.configure(text=f"{cols} columns, {rows} rows")
            
def start_threaded_matching():
    # Start the run_matching function in a separate thread
    threading.Thread(target=run_matching, daemon=True).start()

def update_progress_bar(total_tasks, current_progress):
    progress_bar.set(current_progress / total_tasks)
    progress_label.configure(text=f"Progress: {current_progress}/{total_tasks}")
    root.update_idletasks()

def show_error(message):
    messagebox.showerror("Error", message)
    
def create_button(text, var, is_output=False):
    button = ctk.CTkButton(root, text=text, command=lambda: browse_file(var, button, is_output), width=200)
    button.pack(pady=10)
    return button



dataset_1_id_col = ctk.StringVar(value='ID Column')
dataset_1_match_col = ctk.StringVar(value='Match Column')
dataset_2_id_col = ctk.StringVar(value='ID Column')
dataset_2_match_col = ctk.StringVar(value='Match Column')

def update_dropdowns(var, dropdown_id, dropdown_match):
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

group_colors = {
    'blue': {'light': "#A5D8DD", 'dark': "#A5D8DD"},
    'green': {'light': "#6AB187", 'dark': "#6AB187"}
}

def get_group_color(group_index, is_highest_score):
    color_name = 'blue' if group_index % 2 == 0 else 'green'
    return group_colors[color_name]['dark'] if is_highest_score else group_colors[color_name]['light']

def run_matching():
    
    if not dataset_1_path.get():
        show_error("Please select Dataset 1.")
        return

    if not dataset_2_path.get():
        show_error("Please select Dataset 2.")
        return

    if not output_path.get():
        show_error("Please select an output file.")
        return
    if not dataset_1_id_col.get() or dataset_1_id_col.get() == "ID Column":
       show_error("Please select an ID column for Dataset 1.")
       return
    if not dataset_2_id_col.get() or dataset_2_id_col.get() == "ID Column":
       show_error("Please select an ID column for Dataset 2.")
       return
    if not dataset_1_match_col.get() or dataset_1_match_col.get() == "Match Column":
       show_error("Please select a match column for Dataset 1.")
       return
    if not dataset_2_match_col.get() or dataset_2_match_col.get() == "Match Column":
       show_error("Please select a match column for Dataset 2.")
       return
    if dataset_1_match_col.get() == dataset_1_id_col.get() or dataset_2_match_col.get() == dataset_2_id_col.get():
       show_error("Please select distinct columns to identify and match on.")
       return

    threshold_value = float(score_threshold_spinbox.get())

    # Load dataset 1 based on file extension
    if dataset_1_path.get().endswith('.xlsx'):
        dataset_1_df = pd.read_excel(dataset_1_path.get())
    elif dataset_1_path.get().endswith('.csv'):
        dataset_1_df = pd.read_csv(dataset_1_path.get())
    elif dataset_1_path.get().endswith('.dta'):
        dataset_1_df = pd.read_stata(dataset_1_path.get()) 
    
    # Load dataset 2 based on file extension
    if dataset_2_path.get().endswith('.xlsx'):
        dataset_2_df = pd.read_excel(dataset_2_path.get())
    elif dataset_2_path.get().endswith('.csv'):
        dataset_2_df = pd.read_csv(dataset_2_path.get())
    elif dataset_2_path.get().endswith('.dta'):
        dataset_2_df = pd.read_stata(dataset_2_path.get())
        
    dataset_1_rows, _ = dataset_cache[dataset_1_path.get()]  # Get cached rows for dataset 1
    dataset_2_rows, _ = dataset_cache[dataset_2_path.get()]  # Get cached rows for dataset 2
        
    if dataset_1_rows * dataset_2_rows > 100000 and not output_path.get().endswith('.csv'):
        show_error("Too much data for this format, please export to a CSV.")
        run_matching_button.configure(state="normal")
        return

    # Resolve selected columns using .get()
    id_col_1 = dataset_1_id_col.get()
    match_col_1 = dataset_1_match_col.get()
    id_col_2 = dataset_2_id_col.get()
    match_col_2 = dataset_2_match_col.get()
    

    dataset_1_other_cols = [col for col in dataset_1_df.columns if col not in [match_col_1]]
    dataset_2_other_cols = [col for col in dataset_2_df.columns if col not in [match_col_2]]


    selected_matching_type = matching_type_var.get()
    selected_output_type = output_type_var.get()

    total_tasks = dataset_1_rows * dataset_2_rows if selected_output_type == 1 else dataset_1_rows
    update_threshold = round(total_tasks,-1) * 0.1
    
    data = []
    result_df = None
    current_progress = 0
    
    if selected_matching_type == 1:
        scorer = rf.fuzz.token_set_ratio
    elif selected_matching_type == 2:
        scorer = rf.fuzz.token_sort_ratio
    elif selected_matching_type == 3:
        scorer = rf.fuzz.token_ratio
    elif selected_matching_type ==4:
        scorer = rf.fuzz.WRatio
        
    run_matching_button.configure(state="disabled")

    # Output type 1: All Possible Combinations
    if selected_output_type == 1:
        for i in range(len(dataset_1_df)):
            for j in range(len(dataset_2_df)):
                # Check if any of the selected match columns contain missing values
                if pd.isna(dataset_1_df[match_col_1].iloc[i]) or pd.isna(dataset_2_df[match_col_2].iloc[j]):
                    score = 'N/A'
                else:
                    score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])

                data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[j], 
                             dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j], score])

                # Update progress
                current_progress += 1
                if current_progress % update_threshold == 0 or current_progress == round(total_tasks,-1) or current_progress == total_tasks or current_progress==0:
                    # Use root.after() to update the progress bar from the main thread
                    root.after(0, update_progress_bar, total_tasks, current_progress)

        column_list = [id_col_1, id_col_2, match_col_1, match_col_2, 'Match Score']
        result_df = pd.DataFrame(data, columns=column_list)
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_1_df[dataset_1_other_cols],
            how = 'left',
            on = id_col_1,
            )
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_2_df[dataset_2_other_cols],
            how = 'left',
            on = id_col_2,
            )
        
        

    # Output type 2: Highest Matches Only
    elif selected_output_type == 2:
        for i in range(len(dataset_1_df)):
            max_score, best_match = 0, None
            for j in range(len(dataset_2_df)):
                if pd.isna(dataset_1_df[match_col_1].iloc[i]) or pd.isna(dataset_2_df[match_col_2].iloc[j]):
                    score = 0
                else:
                    score = scorer(dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j])

                data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[j], 
                             dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[j], score])

                if score > max_score:
                    max_score = score
                    best_match = j

            if best_match is not None:
                data.append([dataset_1_df[id_col_1].iloc[i], dataset_2_df[id_col_2].iloc[best_match], 
                             dataset_1_df[match_col_1].iloc[i], dataset_2_df[match_col_2].iloc[best_match], max_score])

            # Update progress
            current_progress += 1
            if current_progress % update_threshold == 0 or current_progress == round(total_tasks,-1) or current_progress == total_tasks or current_progress==0:
                    # Use root.after() to update the progress bar from the main thread
                    root.after(0, update_progress_bar, total_tasks, current_progress)

        column_list = [id_col_1, id_col_2, match_col_1, match_col_2, 'Match Score']
        result_df = pd.DataFrame(data, columns=column_list)
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_1_df[dataset_1_other_cols],
            how = 'left',
            on = id_col_1,
            )
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_2_df[dataset_2_other_cols],
            how = 'left',
            on = id_col_2,
            )
        
    elif selected_output_type == 3:
        for i in range(len(dataset_1_df)):
            if pd.isna(dataset_1_df[match_col_1].iloc[i]):
                score = 'N/A'
    
            # Extract all matches above the score_cutoff threshold
            results = rf.process.extract(
                dataset_1_df[match_col_1].iloc[i],
                dataset_2_df[match_col_2].dropna().tolist(),
                scorer = scorer,
                score_cutoff=threshold_value,  # Apply threshold directly
                limit = None
            )
            
            # Adjust unpacking for RapidFuzz and FuzzyWuzzy
            for result in results:
                match, score, _ = result  # Unpack match, score, and index (ignore index)
                
                # Find the corresponding ID for the matched value
                matched_row = dataset_2_df[dataset_2_df[match_col_2] == match].iloc[0]
                data.append([dataset_1_df[id_col_1].iloc[i], matched_row[id_col_2], 
                             dataset_1_df[match_col_1].iloc[i], matched_row[match_col_2], score])
    
            # Update progress
            current_progress += 1
            if current_progress % update_threshold == 0 or current_progress == round(total_tasks,-1) or current_progress == total_tasks or current_progress==0:
                    # Use root.after() to update the progress bar from the main thread
                    root.after(0, update_progress_bar, total_tasks, current_progress)
    
        column_list = [id_col_1, id_col_2, match_col_1, match_col_2, 'Match Score']
        result_df = pd.DataFrame(data, columns=column_list)
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_1_df[dataset_1_other_cols],
            how = 'left',
            on = id_col_1,
            )
        
        result_df = pd.merge(
            left= result_df,
            right=dataset_2_df[dataset_2_other_cols],
            how = 'left',
            on = id_col_2,
            )



        # Save result to the selected file
    if result_df is not None:
        
        output_file = output_path.get()
        
        if clean_switch.get() == 0:
            try:
                if output_file.endswith('.xlsx'):
                    result_df.to_excel(output_file, index=False, engine="xlsxwriter")
                elif output_file.endswith('.csv'):
                    new_pa_dataframe = pa.Table.from_pandas(result_df)
                    csv.write_csv(new_pa_dataframe, output_file)
                    #result_df.to_csv(output_file, index=False)
                elif output_file.endswith('.dta'):
                    result_df.to_stata(output_file, index=False)
                animal_fact = ""
                if fact_switch.get() == 1:
                    animal_fact = "\n\n" + random.choice(facts)
                    
                messagebox.showinfo("Success", "Matching completed and saved to " + output_file + animal_fact)
                
            except PermissionError:
                show_error("Write permission denied. Please close the output file.")
                run_matching_button.configure(state="normal")
                return
    
        # Ensure the columns exist and check for NaN values
    if clean_switch.get() == 1:
        try:
            # Perform grouping and formatting
            result_df['group_id'] = result_df.groupby([id_col_1], sort=False).ngroup() + 1
            result_df['is_highest'] = result_df.groupby('group_id')['Match Score'].transform(max) == result_df['Match Score']
            #result_df['Valid Match'] = result_df.groupby('group_id')['group_id'].transform('size').eq(1).astype(int)
            result_df['Valid Match'] = 0
            result_df['Comments'] = ''
            result_df.sort_values(by=['group_id', 'Match Score'], ascending=[True, False], inplace=True)
    
            highlighted_data = []
            current_group_index = -1
            current_group = None
    
            for index, row in result_df.iterrows():
                if row['group_id'] != current_group:
                    current_group_index += 1
                    current_group = row['group_id']
    
                bg_color = get_group_color(current_group_index, row['is_highest'])
                highlighted_data.append({**row.to_dict(), 'Background Color': bg_color})
    
            highlighted_df = pd.DataFrame(highlighted_data)
    
            # Save to file
            if output_file.endswith('.xlsx'):
                with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                    highlighted_df.drop(columns=['Background Color', 'group_id', 'is_highest']).to_excel(writer, index=False, sheet_name='Matches')
                    workbook = writer.book
                    worksheet = writer.sheets['Matches']
            
                    # Define cell format for the background color
                    for i, color in enumerate(highlighted_df['Background Color']):
                        for j in range(len(highlighted_df.columns) - 3):
                            hex_color = mcolors.to_hex(color).replace('#', '')
                            cell_format = workbook.add_format({'bg_color': hex_color})
            
                            # Write the value with format applied
                            worksheet.write(i + 1, j, highlighted_df.iloc[i, j], cell_format)
            
                    # Adjust column width
                    for col_num, value in enumerate(highlighted_df.columns):
                        worksheet.set_column(col_num, col_num, 20)
                    

            elif output_file.endswith('.csv'):
                highlighted_df = highlighted_df.drop(columns=['Background Color', 'group_id', 'is_highest'])
                new_pa_dataframe = pa.Table.from_pandas(highlighted_df)
                csv.write_csv(new_pa_dataframe, output_file)
                
            elif output_file.endswith('.dta'):
                highlighted_df = highlighted_df.drop(columns=['Background Color', 'group_id', 'is_highest'])
                highlighted_df.to_stata(output_file, index=False)
    
            animal_fact = ""
    
            if fact_switch.get() == 1:
                animal_fact = "\n\n" + random.choice(facts)
    
            messagebox.showinfo("Success", "Matching completed and saved to " + output_file + animal_fact)
    
        except PermissionError:
            show_error("Write permission denied. Please close the output file.")
            run_matching_button.configure(state="normal")
            return
        
    run_matching_button.configure(state="normal")




# UI Components
ctk.CTkLabel(root, text="Fuzzy Matching Tool", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)

# Create a frame for Dataset 1 and place components in a grid
dataset_1_frame = ctk.CTkFrame(root)
dataset_1_frame.pack(pady=10)

dataset_1_button = ctk.CTkButton(dataset_1_frame, text="Select Dataset 1", command=lambda: browse_file(dataset_1_path, dataset_1_button), width=200)
dataset_1_button.grid(row=0, column=0, padx=5)

dataset_1_id_dropdown = ctk.CTkOptionMenu(dataset_1_frame, variable=dataset_1_id_col, values=["ID Column"])
dataset_1_id_dropdown.grid(row=0, column=1, padx=5)

dataset_1_match_dropdown = ctk.CTkOptionMenu(dataset_1_frame, variable=dataset_1_match_col, values=["Match Column"])
dataset_1_match_dropdown.grid(row=0, column=2, padx=5)

# Create a frame for Dataset 2 and place components in a grid
dataset_2_frame = ctk.CTkFrame(root)
dataset_2_frame.pack(pady=10)

dataset_2_button = ctk.CTkButton(dataset_2_frame, text="Select Dataset 2", command=lambda: browse_file(dataset_2_path, dataset_2_button), width=200)
dataset_2_button.grid(row=0, column=0, padx=5)

dataset_2_id_dropdown = ctk.CTkOptionMenu(dataset_2_frame, variable=dataset_2_id_col, values=["ID Column"])
dataset_2_id_dropdown.grid(row=0, column=1, padx=5)

dataset_2_match_dropdown = ctk.CTkOptionMenu(dataset_2_frame, variable=dataset_2_match_col, values=["Match Column"])
dataset_2_match_dropdown.grid(row=0, column=2, padx=5)


# Bind hover events to display dataset info
dataset_1_button.bind("<Enter>", lambda event: show_dataset_info(dataset_1_path, event, dataset_1_button))
dataset_2_button.bind("<Enter>", lambda event: show_dataset_info(dataset_2_path, event, dataset_2_button))

# Reset button text when mouse leaves the button
dataset_1_button.bind("<Leave>", lambda event: dataset_1_button.configure(text=os.path.basename(dataset_1_path.get()) if dataset_1_path.get() else "Select Dataset 1"))
dataset_2_button.bind("<Leave>", lambda event: dataset_2_button.configure(text=os.path.basename(dataset_2_path.get()) if dataset_2_path.get() else "Select Dataset 2"))

output_button = create_button("Select Output File", output_path, is_output=True)

# Radio buttons for selecting output type
ctk.CTkLabel(root, text="Select Output Type").pack(pady=5)
output_type_var = ctk.IntVar(value=2)  # Default to 'Highest Matches Only'
ctk.CTkRadioButton(root, text="All Possible Combinations", variable=output_type_var, value=1).pack()
ctk.CTkRadioButton(root, text="Highest Matches Only", variable=output_type_var, value=2).pack()

threshold_frame = ctk.CTkFrame(root)
threshold_frame.pack(pady=10)

def validate_input(value):
    if value.isdigit() and 0 <= int(value) <= 100:
        return True
    return False

threshold_button = ctk.CTkRadioButton(threshold_frame, text="Matches Above Threshold", variable=output_type_var, value=3)
threshold_button.grid(row=0, column=0, padx=5)

score_threshold_spinbox = IntSpinbox(threshold_frame, width=90, step_size=1)
score_threshold_spinbox.grid(row=0, column=1, padx=5, pady=5)
score_threshold_spinbox.set(80)


def toggle_appearance_mode():
    if appearance_mode_switch.get():
        ctk.set_appearance_mode("light")  # Switch to light mode
    else:
        ctk.set_appearance_mode("dark")   # Switch to dark mode

# Toggle button positioned at the top-right corner
appearance_mode_switch = ctk.CTkSwitch(root, text="Theme", command=toggle_appearance_mode)
appearance_mode_switch.place(relx=0.95, rely=0.02, anchor="ne")  # Adjust position to top-right


# Radio buttons for selecting matching type
ctk.CTkLabel(root, text="Select Matching Type").pack(pady=5)
matching_type_var = ctk.IntVar(value=1)  # Default to 'Set Ratio'
ctk.CTkRadioButton(root, text="Set Ratio", variable=matching_type_var, value=1).pack()
ctk.CTkRadioButton(root, text="Sort Ratio", variable=matching_type_var, value=2).pack()
ctk.CTkRadioButton(root, text="Max of (Set Ratio, Sort Ratio)", variable=matching_type_var, value=3).pack()
ctk.CTkRadioButton(root, text="WRatio", variable=matching_type_var, value=4).pack()

# Add Progress Bar and Progress Label to UI
progress_bar = ctk.CTkProgressBar(root, width=300)
progress_bar.set(0)
progress_bar.pack(pady=10)

progress_label = ctk.CTkLabel(root, text="Progress: 0/0")
progress_label.pack(pady=5)

fact_switch = ctk.CTkSwitch(root, text="Fuzzy animal fact") 
fact_switch.pack(pady=10)

clean_switch = ctk.CTkSwitch(root, text = "Prep for manual checks")
clean_switch.pack(pady=10)

keep_columns_switch = ctk.CTkSwitch(root, text="Keep all columns")
keep_columns_switch.pack(pady=10)

run_matching_button = ctk.CTkButton(root, text="Run Matching", command=start_threaded_matching, width=200)
run_matching_button.pack(pady=10)

root.mainloop()
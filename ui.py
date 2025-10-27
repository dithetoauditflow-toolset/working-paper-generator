#ui.py
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import time  # Import time module for measuring process duration
import threading  # For running processing in a separate thread
import os

# Import processing functions from respective files
from tp_1 import process_files as process_tp1, process_files_for_all_processing as process_tp1_all
from tp_2 import process_files as process_tp2, process_files_for_all_processing as process_tp2_all
from tp_3 import process_files as process_tp3, process_files_for_all_processing as process_tp3_all
from tp_4 import process_files as process_tp4, process_files_for_all_processing as process_tp4_all
from helper_funcs import get_company_info, create_folder_structure_for_all_working_papers

# Helper function to get template paths
def get_template_paths():
    """
    Retrieves and returns the paths of template files located in the 'TEMPLATES/Working_Papers_Templates' directory.

    This function searches for all `.xlsx` files in the 'TEMPLATES/Working_Papers_Templates' folder, sorts them, and checks 
    that there are at least four templates available. If the folder doesn't exist or doesn't contain 
    enough templates, a FileNotFoundError is raised.

    Returns:
        list: A sorted list of file paths to the template files.

    Raises:
        FileNotFoundError: If the 'TEMPLATES' folder or 'Working_Papers_Templates' subfolder is not found or if there are fewer than 4 template files.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(script_dir, "TEMPLATES", "Working_Papers_Templates")
    
    if not os.path.exists(templates_dir):
        raise FileNotFoundError(f"Working_Papers_Templates folder not found in {os.path.join(script_dir, 'TEMPLATES')}")
    
    templates = sorted(
        [os.path.join(templates_dir, file) for file in os.listdir(templates_dir) if file.endswith(".xlsx")]
    )
    
    if len(templates) < 4:
        raise FileNotFoundError("Not enough template files found. Ensure there are at least 4 templates in the Working_Papers_Templates folder.")
    
    return templates

# Core application logic
def select_file():
    """
    Creates and displays a Tkinter file selection window for selecting a file.

    This function sets up the main GUI window for the Working Paper Populator (WP2) application. 
    The window includes a customizable layout, centered on the screen, with an option for selecting 
    a data file. It also sets up the window's appearance, including setting a background color, 
    favicon, and layout for widgets. Additionally, this function initializes variables and buttons 
    needed for user interaction, as well as the process for resetting the state.

    The user interface includes options for selecting the data file, entering consultant information, 
    and selecting the output directory, while allowing for the reset of the process.

    Returns:
        None
    """
    root = tk.Tk()
    root.title("WP2")
    root.geometry("1000x600")

    # Center the main window on the screen
    root.update_idletasks()  # Ensure all dimensions are calculated
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = 1000
    window_height = 600
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x}+{y-30}")

    # Set window color to white
    root.config(bg="white")

    # Center frame for all content
    center_frame = tk.Frame(root, bg="white")
    center_frame.place(relx=0.5, rely=0.5, anchor="center")

    # Set the favicon (icon.ico)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, "img", "icon.ico")
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)
    else:
        messagebox.showwarning("Warning", "Favicon icon.ico not found. Proceeding without it.")

    # Variables to store state
    selected_file_path = None
    consultant_name = None
    output_directory = None
    template_paths = []
    selected_files_paths = []

    # Button style configuration
    button_style = {
        "bg": "#D8D8D8",
        "width": 45,
        "height": 2,
        "font": ("Helvetica", 10),
        "relief": "flat"
    }

    # Reset the process
    def reset_process():
        """
        Resets the entire process, clearing selections and re-initializing the UI.

        This function clears all selected data, consultant information, output directory,
        and template paths. It also destroys all widgets in the center frame and calls 
        the `initialize_ui` function to set up the UI again for a fresh start.
        """
        nonlocal selected_file_path, consultant_name, output_directory, template_paths
        selected_file_path = None
        consultant_name = None
        output_directory = None
        template_paths = []
        for widget in center_frame.winfo_children():
            widget.destroy()
        initialize_ui()

    # Exit the application
    def close_app():
        """
        Closes the main application window.

        This function terminates the Tkinter main loop and closes the application 
        window when called.
        """
        root.destroy()

    def ask_for_name_multiple():
        """
        Prompts the user to enter the consultant's name, after multiple files have been selected.
        This function asks the user for the consultant's name via a simple dialog.
        If a name is provided, it proceeds to initialize templates; otherwise, an error is shown.
        """
        nonlocal consultant_name
        consultant_name = simpledialog.askstring("Consultant Name", "Enter the Consultant Name:")
        if consultant_name:
            try:
                initialize_templates_multiple()
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "Consultant name is required.")
        
    # Select the data file
    def select_data_file():
        """
        Opens a file dialog to select a data file.

        This function prompts the user to select a data file using a file dialog, 
        and then displays a success message with the selected file's name.
        If no file is selected, an error message is shown.
        """
        nonlocal selected_file_path
        selected_file_path = filedialog.askopenfilename(
            title="Select the Data File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if selected_file_path:
            messagebox.showinfo("Success", f"Data file selected: {os.path.basename(selected_file_path)}")
            ask_for_name()
        else:
            messagebox.showerror("Error", "No data file selected.")

    def initialize_templates_multiple():
        """
        Initializes the templates by fetching the paths to available templates.
        This function retrieves the paths to the template files in the "TEMPLATES" folder
        and prompts the user to select an output directory. If either step fails,
        an error message is displayed.
        """
        nonlocal template_paths, output_directory
        try:
            template_paths = get_template_paths()
            output_directory = filedialog.askdirectory(title="Select Output Directory")
            if not output_directory:
                raise ValueError("Output directory selection is required.")
            update_ui_multiple()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize templates or output directory: {str(e)}")
    
    # Select multiple data files
    def select_multiple_data_files():
        """
        Opens a file dialog to select multiple data files.
        This function prompts the user to select multiple data files using a file dialog,
        and then displays a success message with the selected files' names.
        If no file is selected, an error message is shown.
        """
        nonlocal selected_files_paths
        selected_files_paths = filedialog.askopenfilenames(
            title="Select Multiple Data Files",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if selected_files_paths:
            messagebox.showinfo("Success", f"Data files selected: {[os.path.basename(f) for f in selected_files_paths]}")
            ask_for_name_multiple()
        else:
            messagebox.showerror("Error", "No data files selected.")
        
    # Ask for consultant name
    def ask_for_name():
        """
        Prompts the user to enter the consultant's name.

        This function asks the user for the consultant's name via a simple dialog.
        If a name is provided, it proceeds to initialize templates; otherwise, an error is shown.
        """
        nonlocal consultant_name
        consultant_name = simpledialog.askstring("Consultant Name", "Enter the Consultant Name:")
        if consultant_name:
            try:
                initialize_templates()
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "Consultant name is required.")

    # Detect templates and prompt for output directory
    def initialize_templates():
        """
        Initializes the templates by fetching the paths to available templates.

        This function retrieves the paths to the template files in the "TEMPLATES" folder 
        and prompts the user to select an output directory. If either step fails, 
        an error message is displayed.
        """
        nonlocal template_paths, output_directory
        try:
            template_paths = get_template_paths()
            output_directory = filedialog.askdirectory(title="Select Output Directory")
            if not output_directory:
                raise ValueError("Output directory selection is required.")
            update_ui()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize templates or output directory: {str(e)}")

    def update_ui():
        """
        Updates the user interface with company information and generates options.

        This function retrieves the company information from the selected data file 
        and dynamically generates the UI to display the company details, along with 
        buttons for generating working papers and uploading new files. It adjusts 
        the layout to ensure proper centering and structure.
        """
        # Get company information
        company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(selected_file_path)
        
        # Split the string into a list using the comma as a deliter
        periods_list = periods_claimed.split(",")
        
        # Count the number of Periods.
        number_of_periods = len(periods_list)

        # Clear previous widgets in the center frame
        for widget in center_frame.winfo_children():
            widget.destroy()

        # Center frame configuration for proper centering
        center_frame.grid_rowconfigure(0, weight=1)
        center_frame.grid_columnconfigure(0, weight=1)

        # Create an inner frame to hold content for better centering
        inner_frame = tk.Frame(center_frame, bg="white")
        inner_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        # Title for working papers
        tk.Label(
            inner_frame, text="Working Papers Ready to Generate",
            font=("Helvetica", 16), bg="white", anchor="center"
        ).grid(row=0, column=0, columnspan=4, pady=5, sticky="nsew")

        # Define headings and normal text
        heading_1 = "Company Name: "
        heading_2 = "UIF Ref: "
        heading_3 = "Periods Claimed:\n"
        heading_3_1 = "Number of Periods Claimed: "
        heading_4 = "Total Amount Claimed: "
        heading_5 = "Num of Employees: "

        # Use Text widget for the combined info
        text_widget = tk.Text(inner_frame, font=("Helvetica", 10), wrap="word", bg="white", borderwidth=0, spacing2=2)

        # Create the Scroll bar widget
        scrollbar = ttk.Scrollbar(inner_frame, orient="vertical", command=text_widget.yview)  # ttk for themed scrollbar

        # Configure the Text widget to use the Scrollbar
        text_widget.config(yscrollcommand=scrollbar.set)

        # Grid placement: Ensure proper alignment
        text_widget.grid(row=1, column=0, padx=(5, 0), pady=5, sticky="nsew")  # Right padding to avoid overlap
        scrollbar.grid(row=1, column=1, sticky="ns")  # Place scrollbar in separate column

        # Configure Scrollbar to work with text widget
        scrollbar.config(command=text_widget.yview)

        # Make sure the parent frame expands properly
        inner_frame.grid_columnconfigure(0, weight=1)  # Allow text widget to expand properly
        inner_frame.grid_columnconfigure(1, weight=0)  # Keep scrollbar from expanding

        # Insert heading 1 and heading 2 on the same line
        text_widget.insert("1.0", f"{heading_1}", "bold")  # Insert bold heading 1
        text_widget.insert("end", f"{company_name}  |  ", "normal")  # Insert normal text for company name
        text_widget.insert("end", f"{heading_2}", "bold")  # Insert bold heading 2
        text_widget.insert("end", f"{uif_ref}\n\n", "normal")  # Insert normal text for UIF Ref
        
        # Insert heading 3_1 on its own line
        text_widget.insert("end", f"{heading_3_1}", "bold")  # Insert bold heading 3
        text_widget.insert("end", f"{number_of_periods}   |   ", "normal")  # Insert normal periods claimed
        text_widget.insert("end", f"{heading_3}", "bold")  # Insert bold heading 3
        text_widget.insert("end", f"{periods_claimed}\n\n", "normal")  # Insert normal periods claimed

        # Insert heading 4 and heading 5 on the same line
        text_widget.insert("end", f"{heading_4}", "bold")  # Insert bold heading 4
        text_widget.insert("end", f"R{total_amount_claimed}  |  ", "normal")  # Insert normal total amount
        text_widget.insert("end", f"{heading_5}", "bold")  # Insert bold heading 5
        text_widget.insert("end", f"{number_of_employees}", "normal")  # Insert normal number of employees

        # Configure the tags for bold and normal text
        text_widget.tag_configure("bold", font=("Helvetica", 10, "bold"))
        text_widget.tag_configure("normal", font=("Helvetica", 10))

        # Center-align the entire text content
        text_widget.tag_configure("center", justify="center")
        text_widget.tag_add("center", "1.0", "end")

        # Calculate required height
        num_lines = int(text_widget.index('end-1c').split('.')[0])  # Get the number of lines in the text
        text_widget.config(height=num_lines+1, state="disabled")  # Set the height dynamically and disable editing

        # Add to the grid
        text_widget.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        # Select an Option Label (Centered)
        tk.Label(
            inner_frame, text="Select an Option:",
            font=("Helvetica", 10, "bold"), bg="white", anchor="center"
        ).grid(row=2, column=0, columnspan=4, pady=5, sticky="nsew")

        # Buttons Section (Centered, no stretching)
        buttons = [
            ("Generate TP.1 - Compliance and Existence Testing", lambda: process_working_paper(0, "TP.1", selected_file_path)),
            ("Generate TP.2 - Employment Verification Testing", lambda: process_working_paper(1, "TP.2", selected_file_path)),
            ("Generate TP.3 - Payment Verification Testing", lambda: process_working_paper(2, "TP.3", selected_file_path)),
            ("Generate TP.4 - Confirmation of UIF Contributions", lambda: process_working_paper(3, "TP.4", selected_file_path)),
            ("Generate All Working Papers", lambda: process_all(selected_file_path)),
            ("Upload New File", reset_process),
            ("Close", close_app),
        ]

        for idx, (text, command) in enumerate(buttons):
            tk.Button(inner_frame, text=text, command=command, **button_style).grid(row=3 + idx, column=0, columnspan=4, pady=5, sticky="nsew")

    def create_table(inner_frame, selected_files_paths, get_company_info):
        # Define columns
        columns = ("Company Name", "UIF Ref", "Periods Claimed", "Num Periods", "Total Amount Claimed", "Num Employees")
        
        # Create Treeview widget
        tree = ttk.Treeview(inner_frame, columns=columns, show="headings", height=10)
        
        # Define column properties
        for col in columns:
            tree.heading(col, text=col, command=lambda _col=col: sort_treeview(tree, _col, False))  # Sortable columns
            tree.column(col, anchor="w", width=150, stretch=True)
        
        # Add Scrollbars
        v_scroll = ttk.Scrollbar(inner_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(inner_frame, orient="horizontal", command=tree.xview)
        
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Grid placement
        tree.grid(row=1, column=0, sticky="nsew")
        v_scroll.grid(row=1, column=1, sticky="ns")
        h_scroll.grid(row=2, column=0, sticky="ew")
        
        # Adjust column/row weights for resizing
        inner_frame.grid_columnconfigure(0, weight=1)
        inner_frame.grid_rowconfigure(0, weight=1)
        
        # Insert data into table
        for idx, file_path in enumerate(selected_files_paths):
            company_name, uif_ref, periods_claimed, num_employees, total_amount_claimed = get_company_info(file_path)
            periods_list = periods_claimed.split(",")
            num_periods = len(periods_list)
            
            tree.insert("", "end", values=(company_name, uif_ref, periods_claimed, num_periods, f"R{total_amount_claimed}", num_employees))
        
        return tree

    # Sorting function for columns
    def sort_treeview(tree, col, reverse):
        data = [(tree.set(child, col), child) for child in tree.get_children("")]
        data.sort(reverse=reverse)
        
        for index, (val, child) in enumerate(data):
            tree.move(child, "", index)
        
        tree.heading(col, command=lambda: sort_treeview(tree, col, not reverse))

    
    def update_ui_multiple():
        """
        Updates the user interface with company information for multiple files and generates options.
        This function retrieves the company information from the selected data files and dynamically
        generates the UI to display the company details, along with buttons for generating working
        papers and uploading new files. It adjusts the layout to ensure proper centering and structure.
        """
        # Clear previous widgets in the center frame
        for widget in center_frame.winfo_children():
            widget.destroy()

        # Center frame configuration for proper centering
        center_frame.grid_rowconfigure(0, weight=1)
        center_frame.grid_columnconfigure(0, weight=1)

        # Create an inner frame to hold content for better centering
        inner_frame = tk.Frame(center_frame, bg="white")
        inner_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        # Title for working papers
        tk.Label(
            inner_frame, text="Working Papers Ready to Generate for Multiple Files",
            font=("Helvetica", 16), bg="white", anchor="center"
        ).grid(row=0, column=0, columnspan=4, pady=5, sticky="nsew")

        # Call the function to create and display the table
        tree = create_table(inner_frame, selected_files_paths, get_company_info) 
        
        # Select an Option Label (Centered)
        tk.Label(
            inner_frame, text="Select an Option:",
            font=("Helvetica", 10, "bold"), bg="white", anchor="center"
        ).grid(row=2, column=0, columnspan=4, pady=5, sticky="nsew")
        
        # Buttons Section (Centered, no stretching)
        buttons = [
            ("Generate All Working Papers for All Files", lambda: process_all_multiple(tree)),
            ("Upload New File(s)", reset_process),
            ("Close", close_app),
        ]

        for idx, (text, command) in enumerate(buttons):
            tk.Button(inner_frame, text=text, command=command, **button_style).grid(row=3 + idx, column=0, columnspan=4, pady=5, sticky="nsew")

    # Process a single working paper
    def process_working_paper(index, name, file_path):
        """
        Processes a single working paper for a single file
        """
        if not file_path or not consultant_name or not output_directory:
            messagebox.showerror("Error", "Incomplete setup. Ensure data file, name, and output directory are selected.")
            return

        def processing_logic():
            try:
                start_time = time.time()
                process_functions = [process_tp1, process_tp2, process_tp3, process_tp4]
                output_path = process_functions[index](
                    file_path, template_paths[index], consultant_name, output_directory
                )
                end_time = time.time()
                duration = end_time - start_time
                # Convert duration to minutes, seconds, and milliseconds
                minutes = int(duration // 60)
                seconds = int(duration % 60)
                milliseconds = int((duration % 1) * 1000)

                progress_window.destroy()
                messagebox.showinfo(
                    "Success",
                    f"{name} processed successfully for {os.path.basename(file_path)}!\n"
                    f"Time taken: {minutes} minutes {seconds} seconds {milliseconds} milliseconds"
                )
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("Error", f"Failed to process {name} for {os.path.basename(file_path)}: {str(e)}")
        # Create and center the progress bar window
        progress_window = tk.Toplevel()
        progress_window.title(f"Processing {name} for {os.path.basename(file_path)}")
        progress_window.geometry("400x100")
        # Center the window
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width // 2) - (400 // 2)
        y = (screen_height // 2) - (100 // 2)
        progress_window.geometry(f"400x100+{x}+{y}")
        progress_window.resizable(False, False)
        progress_window.transient()

        tk.Label(progress_window, text=f"Processing {name} for {os.path.basename(file_path)}...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="indeterminate")
        progress_bar.pack(pady=10)
        progress_bar.start()
        threading.Thread(target=processing_logic, daemon=True).start()

    def process_multiple_files(index, name):
        """
        Processes the selected working paper for multiple files
        """
        if not selected_files_paths or not consultant_name or not output_directory:
            messagebox.showerror("Error", "Incomplete setup. Ensure data files, name, and output directory are selected.")
            return

        def processing_logic():
            try:
                messages = []
                total_time = []
                total_start_time = time.time()
                for file_path in selected_files_paths:
                    try:
                        start_time = time.time()
                        process_functions = [process_tp1, process_tp2, process_tp3, process_tp4]
                        output_path = process_functions[index](
                            file_path, template_paths[index], consultant_name, output_directory
                        )
                        end_time = time.time()
                        duration = end_time - start_time
                        # Convert duration to minutes, seconds, and milliseconds
                        minutes = int(duration // 60)
                        seconds = int(duration % 60)
                        milliseconds = int((duration % 1) * 1000)
                        messages.append(
                            f"{name} for {os.path.basename(file_path)}: Processed Successfully!\n"
                            f"Time taken: {minutes} minutes {seconds} seconds {milliseconds} milliseconds\n"
                        )
                    except Exception as e:
                        messages.append(f"{name} for {os.path.basename(file_path)}: Failed, error: {str(e)}")
                total_end_time = time.time()
                total_duration = total_end_time - total_start_time
                # Convert total duration to minutes, seconds, and milliseconds
                total_minutes = int(total_duration // 60)
                total_seconds = int(total_duration % 60)
                total_milliseconds = int((total_duration % 1) * 1000)
                total_time.append(
                    f"{total_minutes}m, {total_seconds}s, {total_milliseconds}ms"
                )
                progress_window.destroy()
                show_results_in_table(messages, total_time)  # Call the function to display results in a table
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        # Create and center the progress bar window
        progress_window = tk.Toplevel()
        progress_window.title(f"Processing {name} for All Files")
        progress_window.geometry("400x100")
        # Center the window
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width // 2) - (400 // 2)
        y = (screen_height // 2) - (100 // 2)
        progress_window.geometry(f"400x100+{x}+{y}")
        progress_window.resizable(False, False)
        progress_window.transient()
        tk.Label(progress_window, text=f"Processing {name} for all files...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="indeterminate")
        progress_bar.pack(pady=10)
        progress_bar.start()
        threading.Thread(target=processing_logic, daemon=True).start()

    def process_all(file_path):
        """
        Processes all working papers for a single file
        """
        if not file_path or not consultant_name or not output_directory:
            messagebox.showerror("Error", "Incomplete setup. Ensure data file, name, and output directory are selected.")
            return

        # Get company information to create folder structure
        company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(file_path)
        
        process_functions_all = [process_tp1_all, process_tp2_all, process_tp3_all, process_tp4_all]
        def processing_logic():
            try:
                messages = []
                total_start_time = time.time()
                
                # Create the folder structure and copy files once
                audit_working_papers_folder = create_folder_structure_for_all_working_papers(
                    output_directory, company_name, uif_ref, file_path, template_paths
                )
                
                for i, name in enumerate(["TP.1", "TP.2", "TP.3", "TP.4"]):
                    try:
                        start_time = time.time()
                        output_path = process_functions_all[i](
                            file_path, template_paths[i], consultant_name, audit_working_papers_folder
                        )
                        end_time = time.time()
                        duration = end_time - start_time
                        # Convert duration to minutes, seconds, and milliseconds
                        minutes = int(duration // 60)
                        seconds = int(duration % 60)
                        milliseconds = int((duration % 1) * 1000)
                        messages.append(
                            f"{name}: Processed Successfully!\n"
                            f"Time taken: {minutes} minutes {seconds} seconds {milliseconds} milliseconds\n"
                        )
                    except Exception as e:
                        messages.append(f"{name}: Failed, error: {str(e)}")
                total_end_time = time.time()
                total_duration = total_end_time - total_start_time
                # Convert total duration to minutes, seconds, and milliseconds
                total_minutes = int(total_duration // 60)
                total_seconds = int(total_duration % 60)
                total_milliseconds = int((total_duration % 1) * 1000)
                messages.append(
                    f"\nTotal Time Taken: {total_minutes} minutes {total_seconds} seconds {total_milliseconds} milliseconds"
                )
                progress_window.destroy()
                messagebox.showinfo("Processing Results", "\n".join(messages))
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

        # Create and center the progress bar window
        progress_window = tk.Toplevel()
        progress_window.title("Processing All Working Papers")
        progress_window.geometry("400x100")
        # Center the window
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width // 2) - (400 // 2)
        y = (screen_height // 2) - (100 // 2)
        progress_window.geometry(f"400x100+{x}+{y}")
        progress_window.resizable(False, False)
        progress_window.transient()
        tk.Label(progress_window, text="Processing all working papers...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="indeterminate")
        progress_bar.pack(pady=10)
        progress_bar.start()
        threading.Thread(target=processing_logic, daemon=True).start()

    def process_all_multiple(tree):
        """
        Processes all working papers for all selected files, updating a single progress bar
        and a results table in real-time.
        """
        if not selected_files_paths or not consultant_name or not output_directory:
            messagebox.showerror("Error", "Incomplete setup. Ensure data files, name, and output directory are selected.")
            return

        def processing_logic():
            try:
                results = []  # Collect results for each file
                total_start_time = time.time()  # Start total processing time
                num_files = len(selected_files_paths)

                for idx, file_path in enumerate(selected_files_paths):
                    file_name = os.path.basename(file_path)
                    company_name = file_name.split("_")[0]  # Adjust based on actual filename structure

                    # Update the progress label dynamically
                    def update_label(company, current, total):
                        progress_label.config(text=f"Processing Data File: {company}.xlsx - ({current}/{total})")

                    root.after(0, update_label, company_name, idx + 1, num_files)

                    try:
                        start_time = time.time()  # Start time for this file
                        tp_results = []
                        # Get company information to create folder structure
                        company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(file_path)
                        
                        # Create the folder structure and copy files once for this file
                        audit_working_papers_folder = create_folder_structure_for_all_working_papers(
                            output_directory, company_name, uif_ref, file_path, template_paths
                        )
                        
                        process_functions_all = [process_tp1_all, process_tp2_all, process_tp3_all, process_tp4_all]
                        for i, name in enumerate(["TP.1", "TP.2", "TP.3", "TP.4"]):
                            try:
                                output_path = process_functions_all[i](file_path, template_paths[i], consultant_name, audit_working_papers_folder)
                                tp_results.append("Success")
                            except Exception as e:
                                tp_results.append(str(e))

                        end_time = time.time()  # End time for this file
                        duration = end_time - start_time
                        minutes = int(duration // 60)
                        seconds = int(duration % 60)
                        milliseconds = int((duration % 1) * 1000)
                        time_taken = f"{minutes}m, {seconds}s, {milliseconds}ms"
                        overall_status = "Success"
                        results.append((file_name, overall_status, *tp_results, time_taken))

                    except Exception as e:
                        overall_status = "Fail"
                        results.append((file_name, overall_status, "Failed", "Failed", "Failed", "Failed", str(e)))

                    finally:
                        if idx + 1 == num_files:  # Stop progress bar when all files are processed
                            def stop_progress():
                                progress_bar.stop()
                                progress_window.destroy()  # Destroy the progress window

                            root.after(0, stop_progress)

                # Calculate total processing time
                total_end_time = time.time()
                total_duration = total_end_time - total_start_time  # Get total time in seconds
                minutes = int(total_duration // 60)
                seconds = int(total_duration % 60)
                milliseconds = int((total_duration % 1) * 1000)
                formatted_total_time = f"{minutes}m, {seconds}s, {milliseconds}ms"

                # Pass results and total duration to results window
                show_results_in_table(results, formatted_total_time)

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")



        # Create and center the progress bar window
        progress_window = tk.Toplevel()
        progress_window.title("Processing All Working Papers for All Files")
        progress_window.geometry("400x100")

        # Center the window
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width // 2) - (400 // 2)
        y = (screen_height // 2) - (100 // 2)
        progress_window.geometry(f"400x100+{x}+{y}")
        progress_window.resizable(False, False)
        progress_window.transient()

        progress_label = tk.Label(progress_window, text="Initializing...")
        progress_label.pack(pady=10)

        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="indeterminate")
        progress_bar.start()
        progress_bar.pack(pady=10)

        threading.Thread(target=processing_logic, daemon=True).start()
    
    def show_results_in_table(data, total_duration):
        """
        Displays processing results in a Toplevel window with a scrollable table,
        along with a heading, description, and total processing duration.
        """
        results_window = tk.Toplevel()
        results_window.title("Processing Results")
        results_window.geometry("800x400")

        # Center the window
        screen_width = results_window.winfo_screenwidth()
        screen_height = results_window.winfo_screenheight()
        x = (screen_width // 2) - (800 // 2)
        y = (screen_height // 2) - (400 // 2)
        results_window.geometry(f"800x400+{x}+{y}")
        results_window.transient()

        # Create a frame for layout (centers content)
        frame = tk.Frame(results_window)
        frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Add a heading and short description
        heading = tk.Label(frame, text="Processing Results Overview", font=("Helvetica", 14, "bold"))
        heading.grid(row=0, column=0, pady=5, sticky="n")

        # Add a heading and short description
        heading = tk.Label(frame, text="Processing Results Overview", font=("Helvetica", 14, "bold"))
        heading.grid(row=1, column=0, pady=5, sticky="n")

        # Display total duration
        total_duration_label = tk.Label(frame, text=f"Total Duration: {total_duration}", font=("Helvetica", 10, "bold"))
        total_duration_label.grid(row=2, column=0, pady=5, sticky="n")

        # Ensure labels and widgets expand properly
        frame.grid_columnconfigure(0, weight=1)


        # Define columns
        columns = ("File Name", "Overall Progress", "TP1", "TP2", "TP3", "TP.4", "Duration")

        # Create Treeview widget
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)

        # Define column properties
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=100, stretch=True)

        # Add Scrollbars
        v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        # Grid placement
        tree.grid(row=3, column=0, sticky="nsew")
        v_scroll.grid(row=3, column=1, sticky="ns")
        h_scroll.grid(row=4, column=0, sticky="ew")

        # Adjust column/row weights for resizing
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        # Insert data into table
        for file_name, overall_progress, tp1, tp2, tp3, tp4, duration in data:
            tree.insert("", "end", values=(file_name, overall_progress, tp1, tp2, tp3, tp4, duration))

        # Add an OK button centered below the table
        ok_button = tk.Button(results_window, text="OK", command=results_window.destroy, **button_style)
        ok_button.grid(row=1, column=0, columnspan=2, pady=10)

        # Ensure the main window expands properly
        results_window.grid_columnconfigure(0, weight=1)
        results_window.grid_rowconfigure(0, weight=1)

    # Create initial UI   
    def initialize_ui():
        """
        Initializes the user interface (UI) for the Working Paper Populator application.

        This function sets up the UI elements for the application, including:
        - Dynamically loading and displaying the application logo (if available).
        - Displaying the software's title and a brief explanation of its purpose.
        - Providing a button for the user to select a data file for processing.
        - Displaying developer information.

        The function ensures that the logo image is resized and placed in the center of the window.
        If the logo image is not found, a warning message is displayed to the user.
        Additionally, the function includes an explanation of the software and provides a button 
        to allow the user to upload a data file for generating populated working papers.

        Returns:
            None
        """
        # Dynamically load logo image (logo.png) and resize it
        img_path = os.path.join(script_dir, "img", "logo.png")
        if os.path.exists(img_path):
            img = tk.PhotoImage(file=img_path)
            img_resized = img.subsample(6, 6)  # Resize the image (adjust the values as needed)
            img_placeholder = tk.Label(center_frame, image=img_resized, bg="white")
            img_placeholder.image = img_resized  # Keep a reference to the image object
            img_placeholder.pack(pady=0)
        else:
            messagebox.showwarning("Warning", "Logo image not found. Proceeding without it.")
        
        # New section explaining the software's purpose and creator
        tk.Label(center_frame, text="WORKING PAPER POPULATOR", font=("Helvetica", 16), bg="white").pack(pady=10)
        
        explanation_text = (
            "Welcome to Working Paper Popullator - WP2, a data processing and Working Paper Generator.\n\n TP.1 | TP.2 | TP.3 | TP.4 \n\nClick the 'Select Data File' button to upload a Data File and generate populated working paper. Working papers will be populated with relevant information automatically, allowing for seamless working paper generation."
        )
        
        explanation_label = tk.Label(center_frame, text=explanation_text, font=("Helvetica", 10), wraplength=350, justify="center", bg="white")
        explanation_label.pack(pady=10)
        
        # Button to select data file
        tk.Button(center_frame, text="Select Data File", **button_style, command=select_data_file).pack(pady=10)
        
        # Button to select multiple data files
        tk.Button(center_frame, text="Select Multiple Data Files", **button_style, command=select_multiple_data_files).pack(pady=10)
        
        developer_text = (
            "Developed by Seward Mupereri for Resumes By Ropa Pty Ltd\nWP2 Version: 1.4"
        )
        
        developer_label = tk.Label(center_frame, text=developer_text, font=("Helvetica", 7), wraplength=350, justify="center", bg="white")
        developer_label.pack(pady=10)
    
    # Start the UI
    initialize_ui()
    root.mainloop()

if __name__ == "__main__":
    select_file()
 
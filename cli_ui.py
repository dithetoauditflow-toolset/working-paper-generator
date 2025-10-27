#cli_ui.py
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import time
import threading
import os
import sys
from datetime import datetime
from pathlib import Path
import shutil

# Import processing functions from respective files
from tp_1 import process_files as process_tp1, process_files_for_all_processing as process_tp1_all
from tp_2 import process_files as process_tp2, process_files_for_all_processing as process_tp2_all
from tp_3 import process_files as process_tp3, process_files_for_all_processing as process_tp3_all
from tp_4 import process_files as process_tp4, process_files_for_all_processing as process_tp4_all
from helper_funcs import get_company_info, create_folder_structure_for_all_working_papers

# ANSI color codes for beautiful CLI output
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    BLUE = '\033[34m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    RED = '\033[31m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'

class CLIInterface:
    def __init__(self):
        self.selected_files = []
        self.consultant_name = ""
        self.output_directory = ""
        self.template_paths = []
        self.processing_results = []
        
    def clear_screen(self):
        """Clear the terminal screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def print_header(self):
        """Print the application header"""
        self.clear_screen()
        print(f"{Colors.HEADER}{'='*80}")
        print(f"{Colors.BOLD}{' '*20}WORKING PAPER GENERATOR (WP2) v1.4")
        print(f"{Colors.BOLD}{' '*25}by Seward Mupereri")
        print(f"{'='*80}{Colors.ENDC}")
        print()
    
    def print_box(self, title, content, color=Colors.OKBLUE):
        """Print content in a beautiful box"""
        print(f"{color}┌─ {title} {'─'*(60-len(title))}┐{Colors.ENDC}")
        for line in content:
            print(f"{color}│ {line:<60} │{Colors.ENDC}")
        print(f"{color}└{'─'*62}┘{Colors.ENDC}")
        print()
    
    def print_status(self, message, status="info"):
        """Print status messages with appropriate colors"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        if status == "success":
            print(f"{Colors.OKGREEN}[{timestamp}] ✓ {message}{Colors.ENDC}")
        elif status == "error":
            print(f"{Colors.FAIL}[{timestamp}] ✗ {message}{Colors.ENDC}")
        elif status == "warning":
            print(f"{Colors.WARNING}[{timestamp}] ⚠ {message}{Colors.ENDC}")
        elif status == "info":
            print(f"{Colors.OKCYAN}[{timestamp}] ℹ {message}{Colors.ENDC}")
        else:
            print(f"{Colors.WHITE}[{timestamp}] {message}{Colors.ENDC}")
    
    def print_progress_bar(self, current, total, width=50):
        """Print a progress bar"""
        progress = current / total
        filled = int(width * progress)
        bar = '█' * filled + '░' * (width - filled)
        percentage = int(progress * 100)
        print(f"\r{Colors.OKGREEN}[{bar}] {percentage}% ({current}/{total}){Colors.ENDC}", end='', flush=True)
        if current == total:
            print()
    
    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def get_template_paths(self):
        """Get template file paths"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            templates_dir = os.path.join(script_dir, "TEMPLATES", "Working_Papers_Templates")
            
            if not os.path.exists(templates_dir):
                raise FileNotFoundError(f"Working_Papers_Templates folder not found")
            
            templates = sorted(
                [os.path.join(templates_dir, file) for file in os.listdir(templates_dir) if file.endswith(".xlsx")]
            )
            
            if len(templates) < 4:
                raise FileNotFoundError("Not enough template files found. Need at least 4 templates.")
            
            return templates
        except Exception as e:
            self.print_status(f"Error loading templates: {str(e)}", "error")
            return []
    
    def select_files_dialog(self):
        """Open Tkinter dialog for file selection"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        files = filedialog.askopenfilenames(
            title="Select Excel Data Files",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        root.destroy()
        return list(files) if files else []
    
    def get_consultant_name_dialog(self):
        """Open Tkinter dialog for consultant name input"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        name = simpledialog.askstring(
            "Consultant Name",
            "Enter the consultant's name:",
            initialvalue=self.consultant_name if self.consultant_name else ""
        )
        
        root.destroy()
        return name if name else ""
    
    def select_output_directory_dialog(self):
        """Open Tkinter dialog for output directory selection"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_directory if self.output_directory else os.getcwd()
        )
        
        root.destroy()
        return directory if directory else ""
    
    def display_file_status(self):
        """Display current file selection status"""
        if not self.selected_files:
            content = ["No files selected"]
        else:
            content = []
            for i, file_path in enumerate(self.selected_files, 1):
                try:
                    file_name = os.path.basename(file_path)
                    file_size = os.path.getsize(file_path)
                    mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    content.append(f"{i}. {file_name}")
                    content.append(f"   Path: {file_path}")
                    content.append(f"   Size: {self.format_file_size(file_size)}")
                    content.append(f"   Modified: {mod_time.strftime('%Y-%m-%d %H:%M:%S')}")
                    if i < len(self.selected_files):
                        content.append("")
                except Exception as e:
                    content.append(f"{i}. Error reading file: {str(e)}")
        
        self.print_box("SELECTED FILES", content, Colors.OKGREEN)
    
    def display_current_status(self):
        """Display current application status"""
        status_items = []
        
        # Files status
        if self.selected_files:
            status_items.append(f"Files: {len(self.selected_files)} selected")
        else:
            status_items.append("Files: None selected")
        
        # Consultant status
        if self.consultant_name:
            status_items.append(f"Consultant: {self.consultant_name}")
        else:
            status_items.append("Consultant: Not specified")
        
        # Output directory status
        if self.output_directory:
            status_items.append(f"Output: {self.output_directory}")
        else:
            status_items.append("Output: Not specified")
        
        # Templates status
        if self.template_paths:
            status_items.append(f"Templates: {len(self.template_paths)} loaded")
        else:
            status_items.append("Templates: Not loaded")
        
        self.print_box("CURRENT STATUS", status_items, Colors.OKCYAN)
    
    
    def display_working_paper_menu(self):
        """Display working paper generation options"""
        wp_options = [
            "1. Generate TP.1 (Compliance and Existence Testing)",
            "2. Generate TP.2 (Employment Verification Testing)",
            "3. Generate TP.3 (Payment Verification)",
            "4. Generate TP.4 (Confirmation of UIF Contributions)",
            "5. Generate All Working Papers (TP.1 to TP.4)",
            "6. Generate Multiple Types (e.g., 1,3,4)",
            "7. Exit Application"
        ]
        
        self.print_box("WORKING PAPER GENERATION", wp_options, Colors.MAGENTA)
    
    def validate_selections(self):
        """Validate that all required selections are made"""
        errors = []
        
        if not self.selected_files:
            errors.append("No files selected")
        
        if not self.consultant_name:
            errors.append("Consultant name not specified")
        
        if not self.output_directory:
            errors.append("Output directory not selected")
        
        if not self.template_paths:
            errors.append("Templates not loaded")
        
        return errors
    
    def process_working_papers(self, wp_type):
        """Process working papers based on type"""
        errors = self.validate_selections()
        if errors:
            self.print_status("Cannot proceed with missing information:", "error")
            for error in errors:
                self.print_status(f"  - {error}", "error")
            return False
        
        self.print_status("Starting working paper generation...", "info")
        
        try:
            start_time = time.time()
            results = []
            
            for i, file_path in enumerate(self.selected_files):
                self.print_status(f"Processing file {i+1}/{len(self.selected_files)}: {os.path.basename(file_path)}", "info")
                
                file_start_time = time.time()
                
                if wp_type == "all":
                    # Process all working papers for this file
                    result = self.process_all_working_papers(file_path)
                else:
                    # Process specific working paper
                    result = self.process_single_working_paper(file_path, wp_type)
                
                file_end_time = time.time()
                processing_time = file_end_time - file_start_time
                
                results.append({
                    'file': os.path.basename(file_path),
                    'status': 'Success' if result else 'Failed',
                    'time': f"{processing_time:.2f}s"
                })
                
                self.print_progress_bar(i+1, len(self.selected_files))
            
            end_time = time.time()
            total_time = end_time - start_time
            
            self.display_results(results, total_time)
            return True
            
        except Exception as e:
            self.print_status(f"Error during processing: {str(e)}", "error")
            return False
    
    def process_single_working_paper(self, file_path, wp_type):
        """Process a single working paper type"""
        try:
            # Get company info for status message
            company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(file_path)
            
            wp_names = {
                "tp1": "TP.1 - Compliance and Existence Testing",
                "tp2": "TP.2 - Employment Verification Testing", 
                "tp3": "TP.3 - Payment Verification",
                "tp4": "TP.4 - Confirmation of UIF Contributions"
            }
            
            self.print_status(f"Processing {wp_names[wp_type]} for {company_name}", "info")
            
            if wp_type == "tp1":
                result = process_tp1(file_path, self.template_paths[0], self.consultant_name, self.output_directory)
            elif wp_type == "tp2":
                result = process_tp2(file_path, self.template_paths[1], self.consultant_name, self.output_directory)
            elif wp_type == "tp3":
                result = process_tp3(file_path, self.template_paths[2], self.consultant_name, self.output_directory)
            elif wp_type == "tp4":
                result = process_tp4(file_path, self.template_paths[3], self.consultant_name, self.output_directory)
            
            if result:
                self.print_status(f"✓ {wp_names[wp_type]} completed for {company_name}", "success")
            
            return result
        except Exception as e:
            self.print_status(f"Error processing {wp_type}: {str(e)}", "error")
            return False
    
    def process_all_working_papers(self, file_path):
        """Process all working papers for a file with structured folder creation"""
        try:
            # Get company information to create folder structure
            company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(file_path)
            
            # Create the folder structure and copy files once
            audit_working_papers_folder = create_folder_structure_for_all_working_papers(
                self.output_directory, company_name, uif_ref, file_path, self.template_paths
            )
            
            self.print_status(f"Created structured folders for: {company_name}", "success")
            
            results = []
            results.append(process_tp1_all(file_path, self.template_paths[0], self.consultant_name, audit_working_papers_folder))
            results.append(process_tp2_all(file_path, self.template_paths[1], self.consultant_name, audit_working_papers_folder))
            results.append(process_tp3_all(file_path, self.template_paths[2], self.consultant_name, audit_working_papers_folder))
            results.append(process_tp4_all(file_path, self.template_paths[3], self.consultant_name, audit_working_papers_folder))
            return all(results)
        except Exception as e:
            self.print_status(f"Error processing all working papers: {str(e)}", "error")
            return False
    
    def process_multiple_working_papers(self, choice_string):
        """Process multiple working paper types based on user input"""
        try:
            # Parse the input (e.g., "1,3,4" or "1 3 4")
            choices = [c.strip() for c in choice_string.replace(',', ' ').split()]
            
            wp_mapping = {
                "1": "tp1",
                "2": "tp2", 
                "3": "tp3",
                "4": "tp4"
            }
            
            valid_choices = []
            for choice in choices:
                if choice in wp_mapping:
                    valid_choices.append(wp_mapping[choice])
                else:
                    self.print_status(f"Invalid choice '{choice}'. Skipping.", "warning")
            
            if not valid_choices:
                self.print_status("No valid working paper types selected.", "error")
                return False
            
            self.print_status(f"Processing {len(valid_choices)} working paper types: {', '.join(valid_choices)}", "info")
            
            # Process each selected type
            for wp_type in valid_choices:
                self.process_working_papers(wp_type)
            
            return True
            
        except Exception as e:
            self.print_status(f"Error processing multiple working papers: {str(e)}", "error")
            return False
    
    def display_results(self, results, total_time):
        """Display processing results in a beautiful table"""
        print()
        self.print_status("Processing completed!", "success")
        
        # Create results table
        table_content = []
        table_content.append("File Name                    Status    Time")
        table_content.append("─" * 50)
        
        for result in results:
            status_color = Colors.OKGREEN if result['status'] == 'Success' else Colors.FAIL
            file_name = result['file'][:25] + "..." if len(result['file']) > 25 else result['file']
            table_content.append(f"{file_name:<28} {status_color}{result['status']:<8}{Colors.ENDC} {result['time']}")
        
        table_content.append("─" * 50)
        table_content.append(f"{'Total Processing Time:':<28} {Colors.BOLD}{total_time:.2f}s{Colors.ENDC}")
        
        self.print_box("PROCESSING RESULTS", table_content, Colors.OKBLUE)
        
        # Show folder structure information
        if self.output_directory:
            folder_info = [
                f"Output Directory: {self.output_directory}",
                "",
                "Each company has been organized into structured folders:",
                "├── {UIF Reference} - {Company Name}/",
                "│   ├── AUDIT REPORTING TEMPLATES/",
                "│   ├── AUDIT WORKING PAPERS/",
                "│   │   ├── TP.1_Compliance and Existence Testing/",
                "│   │   ├── TP.2_Employment Verification Testing/",
                "│   │   ├── TP.3_Payment Verification/",
                "│   │   └── TP.4_Confirmation of UIF Contributions/",
                "│   ├── INFORMATION FROM EMPLOYER/",
                "│   └── UIF DATAFILE/"
            ]
            self.print_box("FOLDER STRUCTURE", folder_info, Colors.OKGREEN)
    
    def reset_selections(self):
        """Reset all selections"""
        self.selected_files = []
        self.consultant_name = ""
        self.output_directory = ""
        self.processing_results = []
        self.print_status("All selections have been reset", "info")
    
    def run(self):
        """Main application loop with automatic guided flow"""
        # Load templates
        self.template_paths = self.get_template_paths()
        if not self.template_paths:
            self.print_status("Failed to load templates. Exiting.", "error")
            return
        
        self.print_status(f"Loaded {len(self.template_paths)} template files", "success")
        
        # Welcome message
        self.print_header()
        welcome_content = [
            "Welcome to the Working Paper Generator!",
            "",
            "This application will guide you through the process of",
            "generating audit working papers automatically.",
            "",
            "You will be prompted for:",
            "1. Excel data files to process",
            "2. Consultant name",
            "3. Output directory for generated files",
            "",
            "Then you can choose which working papers to generate."
        ]
        self.print_box("WELCOME", welcome_content, Colors.OKGREEN)
        
        # Automatic guided flow
        self.guided_input_flow()
        
        # Main processing loop
        while True:
            self.print_header()
            self.display_current_status()
            self.display_working_paper_menu()
            
            try:
                wp_choice = input(f"{Colors.BOLD}Select working paper type (1-7): {Colors.ENDC}").strip()
                
                if wp_choice == "6":
                    # Multiple types selection
                    multi_choice = input(f"{Colors.BOLD}Enter working paper numbers (e.g., 1,3,4): {Colors.ENDC}").strip()
                    self.process_multiple_working_papers(multi_choice)
                elif wp_choice == "7":
                    self.print_status("Thank you for using Working Paper Generator!", "info")
                    break
                else:
                    wp_mapping = {
                        "1": "tp1",
                        "2": "tp2", 
                        "3": "tp3",
                        "4": "tp4",
                        "5": "all"
                    }
                    
                    if wp_choice in wp_mapping:
                        self.process_working_papers(wp_mapping[wp_choice])
                    else:
                        self.print_status("Invalid choice. Please select 1-7.", "error")
                        continue
                
                # Ask if user wants to process more files or exit
                continue_choice = input(f"{Colors.BOLD}Process more working papers? (y/n): {Colors.ENDC}").strip().lower()
                if continue_choice in ['n', 'no']:
                    self.print_status("Thank you for using Working Paper Generator!", "info")
                    break
                    
            except KeyboardInterrupt:
                self.print_status("\nApplication interrupted by user", "warning")
                break
            except Exception as e:
                self.print_status(f"Unexpected error: {str(e)}", "error")
    
    def guided_input_flow(self):
        """Guide user through all required inputs automatically with minimal Enter presses"""
        self.print_header()
        
        # Show all steps overview
        overview_content = [
            "You will be guided through 3 quick steps:",
            "",
            "1. Select Excel data files",
            "2. Enter consultant name", 
            "3. Choose output directory",
            "",
            "Each step will open a dialog automatically.",
            "Complete all steps to proceed to working paper selection."
        ]
        self.print_box("SETUP OVERVIEW", overview_content, Colors.OKGREEN)
        
        # Execute all steps in sequence with minimal interruptions
        self.guide_file_selection()
        self.guide_consultant_name()
        self.guide_output_directory()
        
        # Final confirmation
        self.print_header()
        self.print_status("✓ All setup steps completed successfully!", "success")
        self.display_current_status()
        input(f"{Colors.BOLD}Press Enter to proceed to working paper selection...{Colors.ENDC}")
    
    def guide_file_selection(self):
        """Guide user through file selection"""
        self.print_status("STEP 1: Opening file selection dialog...", "info")
        
        files = self.select_files_dialog()
        if files:
            self.selected_files = files
            self.print_status(f"✓ Selected {len(files)} file(s)", "success")
        else:
            self.print_status("No files selected. Please try again.", "warning")
            self.guide_file_selection()  # Retry
    
    def guide_consultant_name(self):
        """Guide user through consultant name input"""
        self.print_status("STEP 2: Opening consultant name dialog...", "info")
        
        name = self.get_consultant_name_dialog()
        if name:
            self.consultant_name = name
            self.print_status(f"✓ Consultant name set to: {name}", "success")
        else:
            self.print_status("No consultant name entered. Please try again.", "warning")
            self.guide_consultant_name()  # Retry
    
    def guide_output_directory(self):
        """Guide user through output directory selection"""
        self.print_status("STEP 3: Opening output directory dialog...", "info")
        
        directory = self.select_output_directory_dialog()
        if directory:
            self.output_directory = directory
            self.print_status(f"✓ Output directory set to: {directory}", "success")
        else:
            self.print_status("No output directory selected. Please try again.", "warning")
            self.guide_output_directory()  # Retry

def main():
    """Main entry point for CLI interface"""
    cli = CLIInterface()
    cli.run()

if __name__ == "__main__":
    main()

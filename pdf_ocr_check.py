import os
import shutil
import fitz # PyMuPDF
import customtkinter as ctk
from tkinter import filedialog 
import csv 
from datetime import datetime 

# =========================
# APPLICATION CLASS (GUI)
# =========================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configure Window ---
        self.title("PDF OCR CHECK")
        self.geometry("700x650") 
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1) 

        # --- Configuration Defaults ---
        self.SOURCE_DIR = r"C:\PDFs\All_PDFs"
        self.DEST_DIR = r"C:\PDFs\Non_OCR_PDFs"
        self.REPORT_DIR = r"C:\PDFs\Reports" 
        self.MOVE_FILES = ctk.BooleanVar(value=True) 
        self.ENABLE_REPORT = ctk.BooleanVar(value=True) 
        self.MIN_TEXT_CHARS = 50 
        
        # New property to store the last generated report path
        self.last_csv_path = None
        
        # --- Build UI Elements ---
        self.create_widgets()

    # --- New Report Viewer Handler ---
    def view_report(self):
        """Opens the last generated CSV file using the default system program."""
        if self.last_csv_path and os.path.exists(self.last_csv_path):
            try:
                # os.startfile() is the standard way to open a file with the default program on Windows
                os.startfile(self.last_csv_path)
                self.log(f"Opened report: {os.path.basename(self.last_csv_path)}")
            except Exception as e:
                self.log(f"ERROR: Could not open report file. Check file association. Error: {e}")
        else:
            self.log("ERROR: No report file available to view.")

    # --- Directory Selection Methods (Unchanged) ---
    def select_source_dir(self):
        """Opens dialog to select the source directory and updates the entry field."""
        initial_dir = self.source_entry.get() if os.path.isdir(self.source_entry.get()) else self.SOURCE_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, folder_selected)
            self.log(f"Source Directory set to: {folder_selected}")

    def select_dest_dir(self):
        """Opens dialog to select the destination directory and updates the entry field."""
        initial_dir = self.dest_entry.get() if os.path.isdir(self.dest_entry.get()) else self.DEST_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, folder_selected)
            self.log(f"Destination Directory set to: {folder_selected}")

    def select_report_dir(self): 
        """Opens dialog to select the report directory and updates the entry field."""
        initial_dir = self.report_entry.get() if os.path.isdir(self.report_entry.get()) else self.REPORT_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.report_entry.delete(0, "end")
            self.report_entry.insert(0, folder_selected)
            self.log(f"Report Directory set to: {folder_selected}")

    def create_widgets(self):
        
        # 1. Header Frame (Source/Destination)
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1) 

        # 2. Source Directory Label & Input & Button
        ctk.CTkLabel(header_frame, text="Source Directory:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.source_entry = ctk.CTkEntry(header_frame, width=450)
        self.source_entry.insert(0, self.SOURCE_DIR)
        self.source_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.source_button = ctk.CTkButton(header_frame, text="...", width=30, command=self.select_source_dir)
        self.source_button.grid(row=0, column=2, padx=(0, 10), pady=5, sticky="e")

        # 3. Destination Directory Label & Input & Button
        ctk.CTkLabel(header_frame, text="Destination Directory:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.dest_entry = ctk.CTkEntry(header_frame, width=450)
        self.dest_entry.insert(0, self.DEST_DIR)
        self.dest_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        self.dest_button = ctk.CTkButton(header_frame, text="...", width=30, command=self.select_dest_dir)
        self.dest_button.grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")
        
        # ----------------------------------------------------
        # THRESHOLD INPUT FRAME
        # ----------------------------------------------------
        threshold_frame = ctk.CTkFrame(self)
        threshold_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        threshold_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(threshold_frame, text="Min. Text Chars (Threshold):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        self.threshold_entry = ctk.CTkEntry(threshold_frame, width=100)
        self.threshold_entry.insert(0, str(self.MIN_TEXT_CHARS))
        self.threshold_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(threshold_frame, text="Set minimum characters needed to be considered 'OCR-D'").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        
        # ----------------------------------------------------
        # REPORT CONTROL FRAME
        # ----------------------------------------------------
        report_frame = ctk.CTkFrame(self)
        report_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        report_frame.grid_columnconfigure(1, weight=1)
        
        # Report Checkbox
        self.report_check = ctk.CTkCheckBox(
            report_frame, 
            text="Generate CSV Report", 
            variable=self.ENABLE_REPORT
        )
        self.report_check.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # Report Directory Label & Input & Button
        ctk.CTkLabel(report_frame, text="Report Directory:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.report_entry = ctk.CTkEntry(report_frame, width=450)
        self.report_entry.insert(0, self.REPORT_DIR)
        self.report_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        self.report_button = ctk.CTkButton(report_frame, text="...", width=30, command=self.select_report_dir)
        self.report_button.grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")

        # ----------------------------------------------------
        # 4. Control Frame (Run and View Report Buttons)
        # ----------------------------------------------------
        control_frame = ctk.CTkFrame(self)
        control_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        control_frame.grid_columnconfigure(0, weight=1)
        
        # Checkbox for moving files
        self.move_check = ctk.CTkCheckBox(
            control_frame, 
            text=f"Move files to destination directory ({self.DEST_DIR})", 
            variable=self.MOVE_FILES
        )
        self.move_check.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # Run Button
        self.run_button = ctk.CTkButton(control_frame, text="START SCAN", command=self.start_scan)
        self.run_button.grid(row=0, column=2, padx=(5, 10), pady=10, sticky="e") # Moved to column 2

        # NEW: View Report Button
        self.view_report_button = ctk.CTkButton(
            control_frame, 
            text="View Report", 
            command=self.view_report,
            state="disabled" # Disabled until report is generated
        )
        self.view_report_button.grid(row=0, column=1, padx=10, pady=10, sticky="e") # Placed between checkbox and Run

        # 5. Output Log
        ctk.CTkLabel(self, text="Real-time Log:").grid(row=4, column=0, padx=20, pady=(10, 0), sticky="sw")
        self.log_textbox = ctk.CTkTextbox(self, wrap="word")
        self.log_textbox.grid(row=5, column=0, padx=20, pady=(5, 20), sticky="nsew")
        self.log_textbox.insert("end", "Ready to start. Click 'START SCAN'.\n")
        self.log_textbox.configure(state="disabled")

    # --- Logging Utility (Unchanged) ---
    def log(self, message):
        """Adds a message to the text box and forces a GUI update."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"{message}\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
        self.update() 

    # --- Core Logic Integration (Unchanged) ---
    def is_pdf_ocred(self, pdf_path):
        """
        Returns True if PDF contains extractable text based on character count.
        """
        try:
            doc = fitz.open(pdf_path)
            total_text = 0
            for page in doc:
                text = page.get_text("text")
                if text:
                    total_text += len(text)
                    if total_text >= self.MIN_TEXT_CHARS: 
                        doc.close()
                        return True
            doc.close()
            return False
        except Exception as e:
            self.log(f"ERROR reading {pdf_path}: {e}")
            return False

    def start_scan(self):
        """The main function, now run from the GUI button."""
        
        # Reset report path and button state at start
        self.last_csv_path = None
        self.view_report_button.configure(state="disabled") 

        # Disable button and clear log
        self.run_button.configure(state="disabled", text="Scanning...")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        # --- Update Runtime Variables from GUI ---
        try:
            self.MIN_TEXT_CHARS = int(self.threshold_entry.get())
        except ValueError:
            self.log("ERROR: Invalid threshold value. Using default of 50.")
            self.MIN_TEXT_CHARS = 50

        # Update directories and controls
        source_dir = self.source_entry.get()
        dest_dir = self.dest_entry.get()
        report_dir = self.report_entry.get()
        move_files = self.MOVE_FILES.get()
        enable_report = self.ENABLE_REPORT.get()

        self.log(f"Starting scan of: {source_dir}")
        self.log(f"Move files is set to: {move_files}")
        self.log(f"Report generation is set to: {'ON' if enable_report else 'OFF'}")
        self.log(f"OCR Check Threshold: {self.MIN_TEXT_CHARS} characters.")

        # Prepare for CSV logging
        log_data = []
        csv_header = ["Timestamp", "Filename", "Full Path", "OCR Status", "Action", "Final Destination", "Error"]
        
        if move_files:
            os.makedirs(dest_dir, exist_ok=True)
            self.log(f"Destination folder created/checked: {dest_dir}")

        if enable_report:
             os.makedirs(report_dir, exist_ok=True)
             self.log(f"Report folder created/checked: {report_dir}")

        non_ocr_files = []
        file_count = 0
        
        try:
            for root, _, files in os.walk(source_dir):
                for filename in files:
                    if filename.lower().endswith(".pdf"):
                        file_count += 1
                        full_path = os.path.join(root, filename)
                        file_log = {"Timestamp": datetime.now().strftime("%H:%M:%S"), "Filename": filename, "Full Path": full_path, "Error": ""}
                        
                        ocr_status = "OCR-D"
                        action = "Kept"
                        final_dest = "N/A"

                        if not self.is_pdf_ocred(full_path):
                            ocr_status = "NON-OCR"
                            non_ocr_files.append(full_path)
                            self.log(f"--> NON-OCR FOUND: {filename}")
                            
                            if move_files:
                                action = "Moved"
                                dest_path = os.path.join(dest_dir, filename)
                                # Auto-rename if a file with the same name already exists
                                if os.path.exists(dest_path):
                                    name, ext = os.path.splitext(filename)
                                    counter = 1
                                    while os.path.exists(dest_path):
                                        dest_path = os.path.join(dest_dir, f"{name}_{counter}{ext}")
                                        counter += 1
                                final_dest = dest_path
                                try:
                                    shutil.move(full_path, dest_path)
                                    self.log(f"    (Moved to {dest_dir})")
                                except Exception as e:
                                    action = "Error Moving"
                                    file_log["Error"] = str(e)
                                    self.log(f"    (ERROR moving file: {e})")
                            
                        else:
                            self.log(f"OCR-D: {filename}")

                        # Compile log row for CSV
                        file_log["OCR Status"] = ocr_status
                        file_log["Action"] = action
                        file_log["Final Destination"] = final_dest
                        log_data.append(file_log)

        except FileNotFoundError:
            self.log(f"\nFATAL ERROR: Source directory not found: {source_dir}")
            
        except Exception as e:
            self.log(f"\nAN UNEXPECTED ERROR OCCURRED: {e}")
        
        # --- Save CSV Log (Conditional) ---
        if enable_report:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename = os.path.join(report_dir, f"pdf_scan_log_{timestamp}.csv")
            
            try:
                with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=csv_header)
                    writer.writeheader()
                    writer.writerows(log_data)
                
                # Success: Store path and enable button
                self.last_csv_path = csv_filename
                self.view_report_button.configure(state="normal")
                
                self.log("\n=========================")
                self.log(f"CSV log successfully saved to: {csv_filename}")
            except Exception as e:
                self.log(f"\nERROR saving CSV log to {report_dir}: {e}")
                self.last_csv_path = None # Reset path if save failed

        # --- Summary ---
        self.log("SUMMARY")
        self.log("=========================")
        self.log(f"Total files scanned: {file_count}")
        self.log(f"Total non-OCR PDFs found: {len(non_ocr_files)}")
        if move_files:
            self.log(f"Files moved to: {dest_dir}")
        self.log("Scan complete.")
        
        self.run_button.configure(state="normal", text="START SCAN")


if __name__ == "__main__":
    app = App()
    app.mainloop()
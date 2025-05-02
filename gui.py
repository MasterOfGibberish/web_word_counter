import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import sys
import subprocess
import webbrowser
from ttkthemes import ThemedTk

class WebWordCounterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Web Word Counter")
        self.root.geometry("700x500")
        self.root.minsize(700, 500)  # Increased minimum size to ensure all elements are visible
        
        # Set icon if available
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        self.create_widgets()
        self.process = None
        self.configure_styles()
        
    def configure_styles(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 10))
        style.configure("TLabel", font=('Helvetica', 11))
        style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        style.configure("Title.TLabel", font=('Helvetica', 16, 'bold'))
        style.configure("Action.TButton", padding=8, font=('Helvetica', 12, 'bold'))
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Web Word Counter", style="Title.TLabel")
        title_label.pack(pady=(0, 20))
        
        # Top section - URL and action buttons
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # URL input
        url_label = ttk.Label(top_frame, text="Website URL:", width=12)
        url_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(top_frame, textvariable=self.url_var, width=40)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Start button - place it prominently in the top frame
        self.start_button = ttk.Button(top_frame, text="START", 
                                       command=self.start_counting, 
                                       style="Action.TButton", 
                                       width=12)
        self.start_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Output frame
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10 10 10 10")
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Output file
        file_frame = ttk.Frame(output_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        file_label = ttk.Label(file_frame, text="Output Filename:", width=15)
        file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.output_var = tk.StringVar(value="website_text")
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_var, width=40)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_button = ttk.Button(file_frame, text="Browse...", command=self.browse_output)
        browse_button.pack(side=tk.LEFT)
        
        # Output format info
        format_frame = ttk.Frame(output_frame)
        format_frame.pack(fill=tk.X, pady=5)
        
        format_info = ttk.Label(format_frame, 
                               text="Both Excel (.xlsx) and Word (.docx) files will be created automatically")
        format_info.pack(side=tk.LEFT, padx=(25, 0))
        
        # Overwrite option
        overwrite_frame = ttk.Frame(output_frame)
        overwrite_frame.pack(fill=tk.X, pady=5)
        
        self.overwrite_var = tk.BooleanVar(value=False)
        overwrite_check = ttk.Checkbutton(overwrite_frame, text="Overwrite existing files", variable=self.overwrite_var)
        overwrite_check.pack(side=tk.LEFT, padx=(25, 0))
        
        # Button frame (for secondary buttons)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_counting, width=15, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        open_button = ttk.Button(button_frame, text="Open Output Folder", command=self.open_output_folder, width=20)
        open_button.pack(side=tk.RIGHT)
        
        # Status and log
        log_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10 10 10 10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Scrollable text widget for log
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, bg="#f5f5f5")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Set initial log message
        self.log_text.insert(tk.END, "Welcome to Web Word Counter. Enter a URL and click 'START' to begin.\n")
        self.log_text.insert(tk.END, "All pages will be crawled automatically and results will be saved in both Excel and Word formats.\n")
        self.log_text.config(state=tk.DISABLED)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
    def add_log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
        
    def browse_output(self):
        initial_dir = os.path.dirname(os.path.abspath(self.output_var.get())) if os.path.isabs(self.output_var.get()) else os.getcwd()
        
        filename = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            title="Save output as",
            filetypes=[('All files', '*.*')],
            defaultextension=""
        )
        
        if filename:
            # Remove any extension as we'll add our own
            filename = os.path.splitext(filename)[0]
            self.output_var.set(filename)
            
    def open_output_folder(self):
        output_file = self.output_var.get()
        if os.path.exists(output_file + ".xlsx") or os.path.exists(output_file + ".docx"):
            folder_path = os.path.dirname(os.path.abspath(output_file))
        else:
            folder_path = os.getcwd()
            
        if sys.platform == 'win32':
            os.startfile(folder_path)
        elif sys.platform == 'darwin':  # macOS
            subprocess.call(['open', folder_path])
        else:  # Linux
            subprocess.call(['xdg-open', folder_path])
            
    def start_counting(self):
        url = self.url_var.get().strip()
        
        if not url:
            messagebox.showerror("Error", "Please enter a website URL")
            return
            
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
            self.url_var.set(url)
            
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("Processing...")
        
        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Get base output filename
        base_output = self.output_var.get()
        if not base_output:
            base_output = "website_text"
            self.output_var.set(base_output)
        
        # Generate Excel and Word filenames
        excel_output = base_output + ".xlsx"
        word_output = base_output + ".docx"
        
        # Run Excel version first
        self.add_log(f"Starting web crawl for {url}")
        self.add_log("Crawling ALL pages on the website - this may take a while...")
        self.add_log(f"Will save as Excel: {excel_output}")
        self.add_log(f"Will save as Word: {word_output}")
        self.add_log("Processing... please wait\n")
        
        # Build commands for both formats, setting very high limit
        excel_cmd = [
            sys.executable, 
            'main.py',
            url,
            '-l', '9999',  # Very high limit to get all pages
            '-w', '20',    # Longer wait time for better page loading
            '-o', excel_output,
            '--format', 'excel'
        ]
        
        word_cmd = [
            sys.executable, 
            'main.py',
            url,
            '-l', '9999',  # Very high limit to get all pages
            '-w', '20',    # Longer wait time for better page loading
            '-o', word_output,
            '--format', 'docx'
        ]
        
        if self.overwrite_var.get():
            excel_cmd.append('--overwrite')
            word_cmd.append('--overwrite')
        
        # Run in a separate thread
        self.process_thread = threading.Thread(
            target=self.run_process_both_formats, 
            args=(excel_cmd, word_cmd)
        )
        self.process_thread.daemon = True
        self.process_thread.start()
            
    def run_process_both_formats(self, excel_cmd, word_cmd):
        excel_success = False
        content_data = None
        
        try:
            # First run to get content - modified to save content data
            self.add_log("Step 1: Crawling website and creating Excel file...")
            
            # Modify command to include a temp file for storing content
            temp_json_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_content.json")
            if "--save-json" not in excel_cmd:
                excel_cmd.extend(["--save-json", temp_json_file])
                
            self.process = subprocess.Popen(
                excel_cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )
            
            # Read output line by line
            for line in iter(self.process.stdout.readline, ''):
                if line:
                    self.add_log(line.strip())
                    
            self.process.stdout.close()
            self.process.wait()
            
            if self.process.returncode == 0:
                self.add_log("Excel file created successfully!")
                excel_success = True
                
                # Now modify the Word command to use the same content data
                if "--load-json" not in word_cmd and os.path.exists(temp_json_file):
                    word_cmd.extend(["--load-json", temp_json_file])
                
                # Now run Word format using the same data
                self.add_log("\nStep 2: Creating Word document from the same data...")
                self.process = subprocess.Popen(
                    word_cmd, 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.STDOUT,
                    universal_newlines=True,
                    bufsize=1
                )
                
                # Read output line by line
                for line in iter(self.process.stdout.readline, ''):
                    if line:
                        self.add_log(line.strip())
                        
                self.process.stdout.close()
                self.process.wait()
                
                if self.process.returncode == 0:
                    self.add_log("Word document created successfully!")
                    self.status_var.set("Completed")
                    
                    # Clean up temp file
                    if os.path.exists(temp_json_file):
                        try:
                            os.remove(temp_json_file)
                        except:
                            pass
                    
                    # Ask to open the output folder
                    if messagebox.askyesno("Process Complete", "Word counting completed. Open the output folder?"):
                        self.open_output_folder()
                else:
                    self.add_log("Failed to create Word document.")
                    self.status_var.set("Partially completed")
            else:
                self.add_log("Failed to create Excel file.")
                self.status_var.set("Failed")
                
        except Exception as e:
            self.add_log("Error: {}".format(str(e)))
            self.status_var.set("Error")
            
            # Clean up temp file
            if os.path.exists(temp_json_file):
                try:
                    os.remove(temp_json_file)
                except:
                    pass
                    
        finally:
            self.process = None
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            
    def open_output_file(self, filepath):
        try:
            if os.path.exists(filepath):
                if sys.platform == 'win32':
                    os.startfile(filepath)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.call(['open', filepath])
                else:  # Linux
                    subprocess.call(['xdg-open', filepath])
            else:
                self.add_log(f"Error: File not found: {filepath}")
        except Exception as e:
            self.add_log(f"Error opening file: {str(e)}")
            
    def stop_counting(self):
        if self.process:
            self.add_log("Stopping the process...")
            self.process.terminate()
            self.process = None
            self.status_var.set("Stopped")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            

if __name__ == "__main__":
    try:
        # Try to use themed Tk
        root = ThemedTk(theme="arc")
    except:
        # Fall back to standard Tk
        root = tk.Tk()
        
    app = WebWordCounterGUI(root)
    root.mainloop() 
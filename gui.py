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
        self.root.geometry("800x600")
        self.root.minsize(750, 550)
        
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
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Web Word Counter", style="Title.TLabel")
        title_label.pack(pady=(0, 20))
        
        # Input frame
        input_frame = ttk.LabelFrame(main_frame, text="Website Settings", padding="10 10 10 10")
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        # URL input
        url_frame = ttk.Frame(input_frame)
        url_frame.pack(fill=tk.X, pady=5)
        
        url_label = ttk.Label(url_frame, text="Website URL:", width=15)
        url_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=50)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Pages limit
        limit_frame = ttk.Frame(input_frame)
        limit_frame.pack(fill=tk.X, pady=5)
        
        limit_label = ttk.Label(limit_frame, text="Max Pages:", width=15)
        limit_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.limit_var = tk.IntVar(value=10)
        limit_spinner = ttk.Spinbox(limit_frame, from_=1, to=100, textvariable=self.limit_var, width=10)
        limit_spinner.pack(side=tk.LEFT)
        
        # Wait time
        wait_frame = ttk.Frame(input_frame)
        wait_frame.pack(fill=tk.X, pady=5)
        
        wait_label = ttk.Label(wait_frame, text="Wait Time (sec):", width=15)
        wait_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.wait_var = tk.IntVar(value=10)
        wait_spinner = ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.wait_var, width=10)
        wait_spinner.pack(side=tk.LEFT)
        
        # Output frame
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10 10 10 10")
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Output file
        file_frame = ttk.Frame(output_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        file_label = ttk.Label(file_frame, text="Output File:", width=15)
        file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.output_var = tk.StringVar(value="website_text.xlsx")
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_var, width=40)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_button = ttk.Button(file_frame, text="Browse...", command=self.browse_output)
        browse_button.pack(side=tk.LEFT)
        
        # Format selection
        format_frame = ttk.Frame(output_frame)
        format_frame.pack(fill=tk.X, pady=5)
        
        format_label = ttk.Label(format_frame, text="Format:", width=15)
        format_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.format_var = tk.StringVar(value="excel")
        excel_radio = ttk.Radiobutton(format_frame, text="Excel", variable=self.format_var, value="excel")
        excel_radio.pack(side=tk.LEFT, padx=(0, 10))
        
        word_radio = ttk.Radiobutton(format_frame, text="Word", variable=self.format_var, value="docx")
        word_radio.pack(side=tk.LEFT)
        
        # Overwrite option
        overwrite_frame = ttk.Frame(output_frame)
        overwrite_frame.pack(fill=tk.X, pady=5)
        
        self.overwrite_var = tk.BooleanVar(value=False)
        overwrite_check = ttk.Checkbutton(overwrite_frame, text="Overwrite existing file", variable=self.overwrite_var)
        overwrite_check.pack(side=tk.LEFT, padx=(25, 0))
        
        # Status and log
        log_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10 10 10 10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Scrollable text widget for log
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, bg="#f5f5f5")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Set initial log message
        self.log_text.insert(tk.END, "Welcome to Web Word Counter. Enter a URL and click 'Start' to begin.\n")
        self.log_text.config(state=tk.DISABLED)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_counting, width=15)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_counting, width=15, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        open_button = ttk.Button(button_frame, text="Open Output Folder", command=self.open_output_folder, width=20)
        open_button.pack(side=tk.RIGHT)
        
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
        
        file_types = [
            ('Excel files', '*.xlsx'), 
            ('Word documents', '*.docx'),
            ('All files', '*.*')
        ]
        
        filename = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            title="Save output as",
            filetypes=file_types,
            defaultextension=".xlsx" if self.format_var.get() == "excel" else ".docx"
        )
        
        if filename:
            self.output_var.set(filename)
            
    def open_output_folder(self):
        output_file = self.output_var.get()
        if os.path.exists(output_file):
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
        
        # Build command
        cmd = [
            sys.executable, 
            'main.py',
            url,
            '-l', str(self.limit_var.get()),
            '-w', str(self.wait_var.get()),
            '-o', self.output_var.get(),
            '--format', self.format_var.get()
        ]
        
        if self.overwrite_var.get():
            cmd.append('--overwrite')
            
        self.add_log(f"Starting web crawl for {url}")
        self.add_log(f"Max pages: {self.limit_var.get()}")
        self.add_log(f"Output file: {self.output_var.get()}")
        self.add_log(f"Format: {self.format_var.get()}")
        self.add_log("Processing... please wait\n")
        
        # Run in a separate thread
        self.process_thread = threading.Thread(target=self.run_process, args=(cmd,))
        self.process_thread.daemon = True
        self.process_thread.start()
            
    def run_process(self, cmd):
        try:
            self.process = subprocess.Popen(
                cmd, 
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
                self.add_log("\nProcess completed successfully!")
                self.status_var.set("Completed")
                
                output_file = self.output_var.get()
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file) / 1024
                    self.add_log(f"Output file created: {output_file} ({output_size:.1f} KB)")
                    
                    # Ask to open the file
                    if messagebox.askyesno("Process Complete", "Word counting completed. Open the output file?"):
                        self.open_output_file(output_file)
            else:
                self.add_log("\nProcess failed.")
                self.status_var.set("Failed")
                
        except Exception as e:
            self.add_log(f"Error: {str(e)}")
            self.status_var.set("Error")
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
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import threading
import sys
import traceback
import subprocess

class UniversalPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal PDF Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Check available libraries
        self.libraries = self.check_available_libraries()
        
        # Variables
        self.input_files = []
        self.output_format = tk.StringVar(value="docx")
        self.conversion_mode = tk.StringVar(value="single")
        
        self.setup_ui()
        
    def check_available_libraries(self):
        """Check which conversion libraries are available"""
        libraries = {
            'pdf2docx': False,
            'pdf2image': False,
            'pypdf': False,
            'fitz': False
        }
        
        try:
            from pdf2docx import Converter
            libraries['pdf2docx'] = True
        except ImportError:
            pass
            
        try:
            from pdf2image import convert_from_path
            libraries['pdf2image'] = True
        except ImportError:
            pass
            
        try:
            import PyPDF2
            libraries['pypdf'] = True
        except ImportError:
            pass
            
        try:
            import fitz  # PyMuPDF
            libraries['fitz'] = True
        except ImportError:
            pass
            
        return libraries
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Universal PDF Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Library Status
        status_frame = ttk.LabelFrame(main_frame, text="Library Status", padding="5")
        status_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        status_text = "Available: "
        available_libs = [lib for lib, available in self.libraries.items() if available]
        status_text += ", ".join(available_libs) if available_libs else "None"
        
        status_label = ttk.Label(status_frame, text=status_text, foreground="green" if available_libs else "red")
        status_label.pack()
        
        # Format Selection
        format_frame = ttk.LabelFrame(main_frame, text="Conversion Format", padding="10")
        format_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        formats = [
            ("Word Document (.docx)", "docx"),
            ("Images (PNG)", "png"),
            ("Images (JPG)", "jpg"),
            ("Text File (.txt)", "txt"),
            ("PDF (Copy)", "pdf")
        ]
        
        for i, (text, value) in enumerate(formats):
            ttk.Radiobutton(format_frame, text=text, variable=self.output_format, 
                           value=value).grid(row=0, column=i, sticky=tk.W, padx=5)
        
        # Mode Selection
        mode_frame = ttk.LabelFrame(main_frame, text="Conversion Mode", padding="10")
        mode_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Radiobutton(mode_frame, text="Single File", variable=self.conversion_mode, 
                       value="single").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Radiobutton(mode_frame, text="Batch Files", variable=self.conversion_mode, 
                       value="batch").grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Radiobutton(mode_frame, text="Folder", variable=self.conversion_mode, 
                       value="folder").grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(0, weight=1)
        
        # Single file input
        self.single_file_frame = ttk.Frame(file_frame)
        self.single_file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.single_file_frame.columnconfigure(0, weight=1)
        
        self.single_file_path = tk.StringVar()
        single_file_entry = ttk.Entry(self.single_file_frame, textvariable=self.single_file_path, state='readonly')
        single_file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(self.single_file_frame, text="Browse File", 
                  command=self.browse_single_file).grid(row=0, column=1)
        
        # Batch files input
        self.batch_files_frame = ttk.Frame(file_frame)
        self.batch_files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.batch_files_frame.columnconfigure(0, weight=1)
        
        self.batch_files_listbox = tk.Listbox(self.batch_files_frame, height=6)
        self.batch_files_listbox.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        batch_btn_frame = ttk.Frame(self.batch_files_frame)
        batch_btn_frame.grid(row=0, column=2, sticky=(tk.N, tk.S), padx=5)
        
        ttk.Button(batch_btn_frame, text="Add Files", 
                  command=self.add_batch_files).pack(fill=tk.X, pady=2)
        ttk.Button(batch_btn_frame, text="Remove", 
                  command=self.remove_batch_file).pack(fill=tk.X, pady=2)
        ttk.Button(batch_btn_frame, text="Clear All", 
                  command=self.clear_batch_files).pack(fill=tk.X, pady=2)
        
        # Folder input
        self.folder_frame = ttk.Frame(file_frame)
        self.folder_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.folder_frame.columnconfigure(0, weight=1)
        
        self.folder_path = tk.StringVar()
        folder_entry = ttk.Entry(self.folder_frame, textvariable=self.folder_path, state='readonly')
        folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(self.folder_frame, text="Browse Folder", 
                  command=self.browse_folder).grid(row=0, column=1)
        
        # Show/hide frames based on mode
        self.update_mode_display()
        self.conversion_mode.trace('w', self.on_mode_change)
        
        # Output Location
        output_frame = ttk.LabelFrame(main_frame, text="Output Location", padding="10")
        output_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        self.output_path = tk.StringVar(value=str(Path.home() / "Documents" / "Converted_Files"))
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(output_frame, text="Browse", 
                  command=self.browse_output_location).grid(row=0, column=1)
        
        # Options Frame
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="10")
        options_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.option1 = tk.BooleanVar(value=True)
        self.option2 = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(options_frame, text="Preserve formatting", 
                       variable=self.option1).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="High quality images", 
                       variable=self.option2).grid(row=0, column=1, sticky=tk.W)
        
        # Convert Button
        self.convert_btn = ttk.Button(main_frame, text="Start Conversion", 
                                     command=self.start_conversion)
        self.convert_btn.grid(row=7, column=0, columnspan=3, pady=20)
        
        # Progress
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, expand=True)
        
        self.status_label = ttk.Label(main_frame, text="Ready to convert")
        self.status_label.grid(row=9, column=0, columnspan=3)
        
        # Results
        results_frame = ttk.LabelFrame(main_frame, text="Conversion Results", padding="10")
        results_frame.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(10, weight=1)
        
        self.results_text = tk.Text(results_frame, height=8, wrap=tk.WORD)
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
    def on_mode_change(self, *args):
        self.update_mode_display()
        
    def update_mode_display(self):
        mode = self.conversion_mode.get()
        
        # Hide all frames first
        self.single_file_frame.grid_remove()
        self.batch_files_frame.grid_remove()
        self.folder_frame.grid_remove()
        
        # Show selected frame
        if mode == "single":
            self.single_file_frame.grid()
        elif mode == "batch":
            self.batch_files_frame.grid()
        elif mode == "folder":
            self.folder_frame.grid()
    
    def browse_single_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.single_file_path.set(file_path)
    
    def add_batch_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.input_files:
                self.input_files.append(file)
                self.batch_files_listbox.insert(tk.END, os.path.basename(file))
    
    def remove_batch_file(self):
        selection = self.batch_files_listbox.curselection()
        if selection:
            index = selection[0]
            self.input_files.pop(index)
            self.batch_files_listbox.delete(index)
    
    def clear_batch_files(self):
        self.input_files.clear()
        self.batch_files_listbox.delete(0, tk.END)
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with PDF Files")
        if folder:
            self.folder_path.set(folder)
    
    def browse_output_location(self):
        folder = filedialog.askdirectory(title="Select Output Location")
        if folder:
            self.output_path.set(folder)
    
    def get_files_to_convert(self):
        mode = self.conversion_mode.get()
        files = []
        
        if mode == "single":
            if self.single_file_path.get():
                files.append(self.single_file_path.get())
        elif mode == "batch":
            files = self.input_files.copy()
        elif mode == "folder":
            folder = self.folder_path.get()
            if folder and os.path.exists(folder):
                for file in Path(folder).glob("*.pdf"):
                    files.append(str(file))
        
        return files
    
    def start_conversion(self):
        # Debug information
        print("=== DEBUG INFO ===")
        print(f"Available libraries: {self.libraries}")
        print(f"Output format: {self.output_format.get()}")
        print(f"Python version: {sys.version}")
        print("==================")
        
        files = self.get_files_to_convert()
        
        if not files:
            messagebox.showwarning("Warning", "Please select files to convert.")
            return
        
        if not self.output_path.get():
            messagebox.showwarning("Warning", "Please select an output location.")
            return
        
        # Check if required libraries are available
        target_format = self.output_format.get()
        if target_format == "docx" and not self.libraries['pdf2docx']:
            self.install_missing_library("pdf2docx")
            return
        elif target_format in ["png", "jpg"] and not self.libraries['pdf2image']:
            self.install_missing_library("pdf2image")
            return
        
        # Disable convert button during conversion
        self.convert_btn.config(state='disabled')
        self.progress_bar.config(value=0, maximum=len(files))
        
        # Start conversion in thread
        thread = threading.Thread(target=self.convert_files, args=(files,))
        thread.daemon = True
        thread.start()
    
    def install_missing_library(self, library_name):
        result = messagebox.askyesno(
            "Missing Library", 
            f"The '{library_name}' library is required for this conversion.\n\n"
            f"Would you like to install it now?"
        )
        
        if result:
            try:
                # Show installation progress
                self.status_label.config(text=f"Installing {library_name}...")
                
                # Install the library
                if library_name == "pdf2image":
                    # pdf2image requires additional dependencies
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "pdf2image", "pillow"])
                else:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", library_name])
                
                messagebox.showinfo("Success", f"{library_name} installed successfully!\nPlease restart the application.")
                self.status_label.config(text="Library installed - Please restart application")
                
            except Exception as e:
                error_msg = f"Failed to install {library_name}:\n{str(e)}"
                messagebox.showerror("Installation Failed", error_msg)
                self.status_label.config(text="Installation failed")
    
    def convert_files(self, files):
        successful = 0
        failed = 0
        failed_files = []
        
        output_dir = self.output_path.get()
        os.makedirs(output_dir, exist_ok=True)
        target_format = self.output_format.get()
        
        for i, file_path in enumerate(files):
            try:
                self.update_status(f"Converting {i+1}/{len(files)}: {os.path.basename(file_path)}")
                self.update_progress(i + 1)
                
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                output_path = os.path.join(output_dir, f"{file_name}.{target_format}")
                
                # Handle duplicate files
                counter = 1
                original_output = output_path
                while os.path.exists(output_path):
                    output_path = os.path.join(output_dir, f"{file_name}_{counter}.{target_format}")
                    counter += 1
                
                # Perform conversion based on target format
                success = False
                if target_format == "docx":
                    success = self.convert_to_docx(file_path, output_path)
                    # If primary method fails, try alternative
                    if not success and self.libraries['fitz']:
                        self.add_result("Primary method failed, trying alternative...")
                        success = self.convert_to_docx_alternative(file_path, output_path)
                elif target_format in ["png", "jpg"]:
                    success = self.convert_to_image(file_path, output_path, target_format)
                elif target_format == "txt":
                    success = self.convert_to_text(file_path, output_path)
                elif target_format == "pdf":
                    success = self.copy_pdf(file_path, output_path)
                else:
                    success = False
                
                if success:
                    successful += 1
                    self.add_result(f"✓ {os.path.basename(file_path)} → {os.path.basename(output_path)}")
                else:
                    failed += 1
                    failed_files.append(os.path.basename(file_path))
                    self.add_result(f"✗ {os.path.basename(file_path)} - Conversion failed")
                    
            except Exception as e:
                failed += 1
                failed_files.append(os.path.basename(file_path))
                error_details = f"✗ {os.path.basename(file_path)} - Error: {str(e)}"
                self.add_result(error_details)
                print(f"Conversion error details: {traceback.format_exc()}")
        
        # Final update
        self.root.after(0, self.conversion_complete, successful, failed, failed_files)
    
    def convert_to_docx(self, pdf_path, output_path):
        """Convert PDF to Word document with better error handling"""
        try:
            from pdf2docx import Converter
            import traceback
            
            print(f"Converting {pdf_path} to {output_path}")  # Debug info
            
            cv = Converter(pdf_path)
            cv.convert(output_path)
            cv.close()
            print("Conversion successful!")  # Debug info
            return True
        except Exception as e:
            error_msg = f"DOCX conversion error: {str(e)}"
            print(error_msg)
            print(traceback.format_exc())  # This will show the full error trace
            return False
    
    def convert_to_docx_alternative(self, pdf_path, output_path):
        """Alternative method using PyMuPDF for basic text extraction"""
        try:
            if not self.libraries['fitz']:
                return False
                
            import fitz  # PyMuPDF
            
            print(f"Using alternative conversion for {pdf_path}")
            
            doc = fitz.open(pdf_path)
            text_content = []
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text = page.get_text()
                text_content.append(text)
            
            doc.close()
            
            # Create a simple text file as fallback
            if output_path.endswith('.docx'):
                output_path = output_path.replace('.docx', '_alternative.txt')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(text_content))
            
            print("Alternative conversion completed (saved as text)")
            return True
            
        except Exception as e:
            print(f"Alternative conversion also failed: {e}")
            return False
    
    def convert_to_image(self, pdf_path, output_path, format):
        """Convert PDF to images"""
        try:
            from pdf2image import convert_from_path
            
            # Create a directory for multiple pages
            base_name = os.path.splitext(output_path)[0]
            os.makedirs(base_name, exist_ok=True)
            
            # Use higher DPI for better quality if option is selected
            dpi = 300 if self.option2.get() else 200
            
            images = convert_from_path(pdf_path, dpi=dpi, fmt=format.upper())
            
            for i, image in enumerate(images):
                if len(images) == 1:
                    # Single page - use original output path
                    image.save(output_path, format=format.upper())
                else:
                    # Multiple pages - save as page1, page2, etc.
                    page_path = os.path.join(base_name, f"page_{i+1}.{format}")
                    image.save(page_path, format=format.upper())
            
            return True
        except Exception as e:
            print(f"Image conversion error: {e}")
            print(traceback.format_exc())
            return False
    
    def convert_to_text(self, pdf_path, output_path):
        """Convert PDF to text"""
        try:
            if self.libraries['fitz']:
                import fitz
                doc = fitz.open(pdf_path)
                text = ""
                for page in doc:
                    text += page.get_text()
                doc.close()
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                return True
            elif self.libraries['pypdf']:
                # Fallback to PyPDF2
                import PyPDF2
                with open(pdf_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    text = ""
                    for page in reader.pages:
                        text += page.extract_text() + "\n"
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                return True
            else:
                return False
        except Exception as e:
            print(f"Text conversion error: {e}")
            print(traceback.format_exc())
            return False
    
    def copy_pdf(self, pdf_path, output_path):
        """Copy PDF file (useful for batch processing)"""
        try:
            import shutil
            shutil.copy2(pdf_path, output_path)
            return True
        except Exception as e:
            print(f"PDF copy error: {e}")
            return False
    
    def update_status(self, message):
        self.root.after(0, lambda: self.status_label.config(text=message))
    
    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_bar.config(value=value))
    
    def add_result(self, message):
        self.root.after(0, lambda: self.results_text.insert(tk.END, message + "\n"))
        self.root.after(0, lambda: self.results_text.see(tk.END))
    
    def conversion_complete(self, successful, failed, failed_files):
        self.convert_btn.config(state='normal')
        self.update_status(f"Conversion complete: {successful} successful, {failed} failed")
        
        if failed > 0:
            messagebox.showwarning(
                "Conversion Complete with Errors",
                f"Conversion completed!\n\n"
                f"Successful: {successful}\n"
                f"Failed: {failed}\n\n"
                f"Check the results panel for details."
            )
        else:
            messagebox.showinfo(
                "Conversion Complete",
                f"All files converted successfully!\n\n"
                f"Files saved to: {self.output_path.get()}"
            )

def main():
    root = tk.Tk()
    app = UniversalPDFConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()

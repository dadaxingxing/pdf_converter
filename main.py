import tkinter as tk
from tkinter import filedialog
import sys
import os 
import comtypes.client
from docx import Document
from time import sleep
import gc
from threading import Thread

class MyGUI:
    def __init__(self):
        # Setup window
        self.window = tk.Tk()
        self.window.geometry("500x300")
        self.window.title("Mass Convert .docx & .doc to .pdf")

        # Input folder path
        self.inputPath = tk.StringVar()
        self.inputButton = tk.Button(
            self.window, 
            text='Input Folder', 
            font=('Arial', 10), 
            command=lambda: self.select_file(self.inputPath))
        self.inputButton.pack(pady=10)

        # Output folder path
        self.outputPath = tk.StringVar()
        self.outputButton = tk.Button(
            self.window, 
            text='Output Folder', 
            font=('Arial', 10), 
            command=lambda: self.select_file(self.outputPath))
        self.outputButton.pack(padx=10)

        # Create the convert button
        self.convertButton = tk.Button(
            self.window, 
            text='Convert', 
            font=('Arial', 10), 
            command=self.mass_convert)
        self.convertButton.pack(pady=19)

        # Text widget for logging messages
        self.log = tk.Text(self.window, height=10, wrap=tk.WORD, font=("Arial", 10))
        self.log.pack(pady=10, fill=tk.BOTH, expand=True)

        # Start the window
        self.window.mainloop()

    def convert_input_folder(self, in_dir, out_dir):
        """Convert all .doc files into .docx"""
        if not os.path.exists(in_dir) or not os.path.exists(out_dir):
            self.log_message(f'Input Folder: {in_dir} or {out_dir} does not exist')
            return
        
        word = None
        try:
            # Create the Word application instance once
            word = comtypes.client.CreateObject('word.application')

            for filename in os.listdir(in_dir):
                input_file = os.path.join(in_dir, filename)
                output_file = os.path.join(out_dir, filename)
                
                if filename.endswith('.doc') or filename.endswith('.docx'):
                    try:
                        output_file = os.path.splitext(output_file)[0] + '.pdf'
                        if os.path.exists(output_file):
                            os.remove(output_file)

                        doc = word.Documents.Open(input_file)
                        doc.SaveAs(output_file, FileFormat=17)  # FileFormat 17 for PDF
                        doc.Close()
                        
                        sleep(1)
                    except Exception as doc_err:
                        self.log_message(f"Error processing file {filename}: {doc_err}")
            self.log_message('Conversion complete!')

        except Exception as e:
            self.log_message(f"Error during Word application initialization or processing: {e}")
        finally:
            if word:
                try:
                    word.Quit()  # Ensure Word quits cleanly
                    del word
                    gc.collect()
                except Exception as quit_err:
                    self.log_message(f"Error while closing Word application: {quit_err}")
                    
        self.log_message('--------------------')

    def log_message(self, message):
        """Log message to the text widget."""
        self.log.insert(tk.END, message + '\n')
        self.log.see(tk.END)

    def mass_convert(self):
        """Perform Mass Conversion from .docx into .pdf using docx2pdf."""
        in_dir = self.inputPath.get()
        out_dir = self.outputPath.get()

        if not in_dir or not out_dir:
            self.log_message('Please select a directory for both folders.')
            return

        try:
            self.log_message('...')
            self.log_message(f'Converting .docx from {in_dir}')
            self.log_message(f'into .pdf from {out_dir}')
            self.log_message('...')
            sys.stderr = open("consoleoutput.log", "w")

            Thread(target=self.convert_input_folder, args=(in_dir, out_dir), daemon=True).start()
        except Exception as e:
            self.log_message(f'An error occurred during conversion: {e}')

    def select_file(self, file_path_var):
        """Open a directory selection dialog and set the path to StringVar."""
        file_path = filedialog.askdirectory(title='Select a Folder')

        if file_path:
            self.log_message(f'Selected: {file_path}')
            file_path_var.set(file_path)

# Run the GUI application
MyGUI()


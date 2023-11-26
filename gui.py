from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from woo_fun_class import WooFunc
import shutil
import pandas as pd
import os
from datetime import datetime
from tkinter import simpledialog


#pyinstaller --onefile --add-data "C:\Users\Shafir R\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\LocalCache\local-packages\Python39\site-packages\tkdnd\tkdnd;." gui.py



class WooGui:

    def __init__(self) -> None:
        self.root = TkinterDnD.Tk()  # Use TkinterDnD instead of tk.Tk
        self.root.title("Woo Functions")
        self.woo_func = WooFunc()
        
        # Create buttons for downloading and updating stock
        download_button = tk.Button(self.root, text="Download Report", command=self.prompt_save_location)
        download_button.pack()

        # Create a label for file drop for updating stock
        self.file_drop_label = tk.Label(self.root, text="Drop a file here or click to select for stock update", bg="lightgrey", width=40, height=4)
        self.file_drop_label.pack(padx=10, pady=10)

        # Bind the drop event and click event for the label
        self.file_drop_label.drop_target_register(DND_FILES)
        self.file_drop_label.dnd_bind('<<Drop>>', self.drop)
        self.file_drop_label.bind("<Button-1>", lambda e: self.prompt_file_selection())
        
        

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="indeterminate")
        self.progress_bar.pack()

        # Output text area
        self.output_text = tk.Text(self.root, height=10, width=100)
        self.output_text.pack()
        self.output_text.config(state='disabled') 
        
        
        # Drop Zone for PDF Files
        self.pdf_drop_label = tk.Label(self.root, text="Click or drop PDF here for Shipping Label Update", bg="lightgrey", width=40, height=4)
        self.pdf_drop_label.pack(padx=10, pady=10)
        self.pdf_drop_label.drop_target_register(DND_FILES)
        self.pdf_drop_label.dnd_bind('<<Drop>>', self.drop_pdf)

        # Bind click event for the label to prompt for file selection
        self.pdf_drop_label.bind("<Button-1>", self.prompt_pdf_selection)
        self.pdf_drop_label.dnd_bind('<<Drop>>', self.drop_pdf)

        # Button to Choose Save Location
        self.save_button = tk.Button(self.root, text="Choose Save Location and Process Shipping PDF", command=self.choose_save_location)
        self.save_button.pack()
        
        self.golden_fileDrop = tk.Label(self.root, text="Drop a file here or click to select the golden sample (excel file)", bg="lightgrey", width=60, height=4)
        self.golden_fileDrop.pack(padx=10, pady=10)
        
        self.golden_fileDrop.drop_target_register(DND_FILES)
        
        self.golden_fileDrop.dnd_bind('<<Drop>>', self.save_golden_sample_drop)
        self.golden_fileDrop.bind("<Button-1>", self.save_golden_sample_click)
        
        
         
        self.download_stock_update_from_golden = tk.Button(self.root, text="Click to choose process and choose location of restock excel")
        self.download_stock_update_from_golden.pack(padx=10, pady=10)
        self.download_stock_update_from_golden.bind("<Button-1>", lambda e: self.download_stock_update_fromGoldenSample())
        
        
        # Frame for credentials
        self.credential_frame = tk.Frame(self.root)
        self.credential_frame.pack(fill=tk.X, pady=10)

        # Store URL entry
        tk.Label(self.credential_frame, text="Store URL:").pack(side=tk.LEFT)
        self.store_url_entry = tk.Entry(self.credential_frame)
        self.store_url_entry.pack(side=tk.LEFT)

        # Consumer Key entry
        tk.Label(self.credential_frame, text="Consumer Key:").pack(side=tk.LEFT,)
        self.consumer_key_entry = tk.Entry(self.credential_frame, show="*")
        self.consumer_key_entry.pack(side=tk.LEFT)

        # Consumer Secret entry
        tk.Label(self.credential_frame, text="Consumer Secret:").pack(side=tk.LEFT)
        self.consumer_secret_entry = tk.Entry(self.credential_frame, show="*")
        self.consumer_secret_entry.pack(side=tk.LEFT)

        # Buttons frame
        self.buttons_frame = tk.Frame(self.root)
        self.buttons_frame.pack(fill=tk.X)

        self.delete_button = tk.Button(self.buttons_frame, text="Delete Credentials", command=self.delete_credentials)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        self.update_button = tk.Button(self.buttons_frame, text="Update Credentials", command=self.update_credentials)
        self.update_button.pack(side=tk.LEFT, padx=5)

        # Load existing credentials if available
        self.load_credentials()
        
        
        
        
        
        
        

        # Set window size to screen size
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")

        self.root.mainloop()
        
    def save_golden_sample_drop(self,event):
        file_path = event.data
        threading.Thread(target=self.process_golden_sample, args=(file_path, self.report_done_golden_sample_copy), daemon=True).start()
        print(file_path)
        
    def save_golden_sample_click(self,event):
        file_path = filedialog.askopenfilename()
        threading.Thread(target=self.process_golden_sample, args=(file_path, self.report_done_golden_sample_copy), daemon=True).start()
        print(file_path)
    
    def process_golden_sample(self,filepath,callback=None):
        success = False
        if filepath.endswith('.xlsx'):
            try:
                shutil.copyfile(filepath, "golden_sample.xlsx")
                print("copied")
                success = True
                
            except IOError as e:
                print(f"Error occurred while copying file: {e}")
        else:
            print("file not excel")
            
        if callback:
            callback(success)
            
    def report_done_golden_sample_copy(self,success):
        self.progress_bar.stop()
        if success:
            messagebox.showinfo("Success", "Golden sample Updated")
        else:
            messagebox.showinfo("Failed", "Could not update Golden Sample. Maybe wrong file type input") 
            
    def download_stock_update_fromGoldenSample(self):
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

        if not output_path:
            # User cancelled the save operation
            self.root.after(0, self.resetButtonState)
            return

        # Start the progress bar and initiate the process in a new thread
        self.progress_bar.start()
        threading.Thread(target=self.process_stock_update, args=(output_path,), daemon=True).start()

    def resetButtonState(self):
        self.download_stock_update_from_golden.config(relief=tk.RAISED, state=tk.NORMAL)

    
    def process_stock_update(self, output_path):
        file_path = 'product_stock_by_category.xlsx'
        
        golden_sample_path = 'golden.xlsx'
        
        if os.path.exists(golden_sample_path) == False:
            messagebox.showinfo("Failed", "Golden File not exists. Please Upload golden sample")
            self.progress_bar.stop()
            return

        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"File {file_path} has been deleted.")
        else:
            print(f"The file {file_path} does not exist.")

        try:
            woo_instance = self.woo_func
            woo_instance.generate_stock_report(file_path)
            print("Current stock report generated.")
        except Exception as e:
            messagebox.showinfo("Error", f"Failed to generate stock report: {e}")
            self.progress_bar.stop()
            return

        if not os.path.exists(file_path):
            messagebox.showinfo("Failed", "Could not download current stock")
            self.progress_bar.stop()
            return

        current_stock_path = file_path
       

        try:
            # Read the excel files
            current_stock_excel = pd.ExcelFile(current_stock_path)
            golden_sample_excel = pd.ExcelFile(golden_sample_path)

            writer = pd.ExcelWriter(output_path, engine='openpyxl')

            for sheet_name in golden_sample_excel.sheet_names:
                golden_sample_df = golden_sample_excel.parse(sheet_name)
                if sheet_name in current_stock_excel.sheet_names:
                    current_stock_df = current_stock_excel.parse(sheet_name)
                    golden_sample_df.set_index('Product Name', inplace=True)
                    current_stock_df.set_index('Product Name', inplace=True)

                    required_stock = golden_sample_df.subtract(current_stock_df, fill_value=0)
                    required_stock = required_stock.clip(lower=0)
                    required_stock = required_stock[(required_stock.T != 0).any()]

                    if not required_stock.empty:
                        required_stock.to_excel(writer, sheet_name=sheet_name)
                else:
                    golden_sample_df.to_excel(writer, sheet_name=sheet_name)

            writer.close()
            print("Restock list created at:", output_path)
        except Exception as e:
            messagebox.showinfo("Error", f"Failed to create restock list: {e}")
        finally:
            # Stop the progress bar
            self.progress_bar.stop()
            self.root.after(0, self.resetButtonState)
                
        
        

    def update_output_text(self, message):
        self.output_text.config(state='normal')  # Temporarily enable the widget to insert text
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)  # Auto-scroll to the end
        self.output_text.config(state='disabled')  # Disable the widget again

    def prompt_save_location(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.progress_bar.start()
            threading.Thread(target=self.woo_func.generate_stock_report, args=(file_path, self.report_done), daemon=True).start()
            
            if os.path.exists(file_path):
                folder_path = 'all_past_stocks'
                # Check if the folder exists
                if not os.path.exists(folder_path):
                    # If it doesn't exist, create it
                    os.makedirs(folder_path)
                    print(f"Folder '{folder_path}' created.")
                else:
                    print(f"Folder '{folder_path}' already exists.")
                
                # Get the current date and time
                current_datetime = datetime.now()

                # Format the date and time as a string
                #formatted_datetime = current_datetime.strftime("_%Y-%m-%d %H-%M-%S")
                formatted_datetime = current_datetime.strftime("_%d-%m-%Y %H-%M-%S")
                shutil.copyfile(file_path, folder_path + "/current_stock" + formatted_datetime + ".xlsx" )
                #shutil.copyfile(file_path,"current_stock.xlsx" )
                print("report saved")
            

    def report_done(self):
        self.progress_bar.stop()
        messagebox.showinfo("Success", "Report saved successfully!")
    
    
    
    def drop(self, event):
        file_path = event.data

        # Handle the file path format for different platforms
        if file_path.startswith('{'):
            file_path = file_path[1:-1]  # Remove curly braces

        if file_path.startswith('file:///'):
            file_path = file_path[8:]  # Remove 'file:///' prefix

        # Replace forward slashes with backslashes if on Windows
        file_path = file_path.replace('/', '\\')

        if file_path:
            self.process_file(file_path)

    def prompt_file_selection(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        self.progress_bar.start()
        threading.Thread(target=self.woo_func.update_stock_from_excel, args=(file_path, self.update_done_stock, self.update_output_text), daemon=True).start()

    def update_done_stock(self):
        self.progress_bar.stop()
        messagebox.showinfo("Success", "Stock updated successfully!")
        
        
    def drop_pdf(self, event):
        self.pdf_path = event.data
        if self.pdf_path.startswith('{'):
            self.pdf_path = self.pdf_path[1:-1]
        if self.pdf_path.startswith('file:///'):
            self.pdf_path = self.pdf_path[8:]
        self.pdf_path = self.pdf_path.replace('/', '\\')
        messagebox.showinfo("PDF Selected", f"PDF file selected: {self.pdf_path}")
        
    def prompt_pdf_selection(self, event):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if file_path:
            self.pdf_path = file_path
            self.process_pdf()
            #messagebox.showinfo("PDF Selected", f"PDF file selected: {self.pdf_path}")

    def process_shipping_pdf_str(self):
        
        if self.pdf_path.startswith('{'):
            self.pdf_path = self.pdf_path[1:-1]
        if self.pdf_path.startswith('file:///'):
            self.pdf_path = self.pdf_path[8:]
        self.pdf_path = self.pdf_path.replace('/', '\\')
        messagebox.showinfo("PDF Selected", f"PDF file selected: {self.pdf_path}")


    def choose_save_location(self):
        self.save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if self.save_path:
            self.progress_bar.start()
            # Start the thread without passing a function to start()
            threading.Thread(target=self.woo_func.shipping_label_update, args=(self.pdf_path, self.save_path, self.shipping_update_done), daemon=True).start()


    def shipping_update_done(self):
        self.progress_bar.stop()
        messagebox.showinfo("Success", "Shipping label updated successfully!")



    
    def load_credentials(self):
        if os.path.exists('credential'):
            with open('credential', 'r') as file:
                lines = file.readlines()
                self.store_url_entry.insert(0, lines[0].strip())
                self.consumer_key_entry.insert(0, lines[1].strip())
                self.consumer_secret_entry.insert(0, lines[2].strip())            
    

    def update_credentials(self):
        with open('credential', 'w') as file:
            file.write(self.store_url_entry.get() + '\n')
            file.write(self.consumer_key_entry.get() + '\n')
            file.write(self.consumer_secret_entry.get())
        messagebox.showinfo("Update", "Credentials updated successfully.")

    def delete_credentials(self):
        if os.path.exists('credential'):
            os.remove('credential')
        self.store_url_entry.delete(0, tk.END)
        self.consumer_key_entry.delete(0, tk.END)
        self.consumer_secret_entry.delete(0, tk.END)
        messagebox.showinfo("Delete", "Credentials deleted.")
    
    
    



WooGui()

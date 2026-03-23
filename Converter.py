import os
import win32com.client
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

class PsbToPsdConverter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        # UI Setup
        self.title("PSB to PSD Batch Converter")
        self.geometry("400x250")
        self.configure(bg="#1e1e1e")
        
        # Drag and Drop Area
        self.drop_area = tk.Label(
            self, 
            text="Drag & Drop .psb files here",
            bg="#2d2d2d", 
            fg="#ffffff",
            font=("Helvetica", 12)
        )
        self.drop_area.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Register the drop target
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.handle_drop)
        
    def handle_drop(self, event):
        import os

        paths = self.split_dnd_paths(event.data)
        psb_files = []

        for path in paths:
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith('.psb'):
                            full_path = os.path.join(root, file)
                            psb_files.append(full_path)
            
            elif os.path.isfile(path) and path.lower().endswith('.psb'):
                psb_files.append(path)
        
        if not psb_files:
            messagebox.showinfo("Notice", "No .psb files detected in the dropped items or folders.")
            return
        
        self.process_files_sequentially(psb_files)
    def split_dnd_paths(self, data):
        # Helper to clean up file paths from tkinterdnd2
        import re
        raw_paths = re.findall(r'\{.*?\}|\S+', data)
        
        # Strip the curly braces immediately so the .psb check won't fail
        cleaned_paths = [path.strip('{}') for path in raw_paths]
        
        return cleaned_paths

    def process_files_sequentially(self, file_list):
        self.drop_area.config(text="Processing... Please wait.")
        self.update()
        
        try:
            # Connect to Photoshop silently
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            
            for file_path in file_list:
                clean_path = file_path.strip('{}') # Remove braces if any
                
                # Verify file exists
                if not os.path.exists(clean_path):
                    continue
                    
                # 1. Open the PSB file
                doc = ps_app.Open(clean_path)
                
                # 2. Prepare the new PSD path
                new_path = os.path.splitext(clean_path)[0] + ".psd"
                
                # 3. Save as PSD (Photoshop SaveOptions format)
                # 1 is the constant for Photoshop save format in ExtendScript
                psd_save_options = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
                doc.SaveAs(new_path, psd_save_options, True)
                
                # 4. Close document without saving changes to original
                doc.Close(2) # 2 means "Do not save changes"
                
            self.drop_area.config(text="Done! Drag & Drop more files.")
            messagebox.showinfo("Success", f"Successfully converted {len(file_list)} files.")
            
        except Exception as e:
            self.drop_area.config(text="Drag & Drop .psb files here")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    app = PsbToPsdConverter()
    app.mainloop()
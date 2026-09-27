import fitz  # PyMuPDF
import pandas as pd
from tkinter import Tk, filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import os
from openpyxl.styles import Font

def extract_all_annotations(pdf_path):
    doc = fitz.open(pdf_path)
    data = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        annots = page.annots()

        if annots is None:
            continue

        for annot in annots:
            annot_code = annot.type[0]
            annot_type = annot.type[1]
            comment = annot.info.get("content", "").strip()
            page_number = page_num + 1
            highlight_text = ""

            if annot_code == 8:  # Highlight
                quad_points = annot.vertices
                quad_count = int(len(quad_points) / 4)
                highlight_text = ""

                for i in range(quad_count):
                    quad = quad_points[i * 4:(i + 1) * 4]
                    rect = fitz.Quad(quad).rect
                    
                    # Get blocks (text boxes) within the highlight area
                    blocks = page.get_text("blocks", clip=rect)
                    
                    # Extract text from blocks, respecting their vertical order
                    block_texts = []
                    for block in blocks:
                        if block[4]:  # block[4] is the text content
                            block_texts.append(block[4].strip())
                    
                    highlight_text += " " + " ".join(block_texts)

                highlight_text = " ".join(highlight_text.split())  # Clean up extra spaces
                data.append({
                    "Page": page_number,
                    "Annotation Type": "Highlight",
                    "Highlighted Text": highlight_text,
                    "Comment": comment
                })

            elif annot_type == "Text":  # Sticky note
                data.append({
                    "Page": page_number,
                    "Annotation Type": "Comment",
                    "Highlighted Text": "",
                    "Comment": comment
                })

    return data

def get_next_filename(base_name):
    """Find the next available filename with incrementing number"""
    counter = 1
    while True:
        file_name = f"{base_name}_{counter}.xlsx"
        if not os.path.exists(file_name):
            return file_name
        counter += 1

class PDFSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Annotation Extractor - Multi-Folder Selector")
        self.root.geometry("900x600")
        
        self.selected_files = []  # List to store selected PDF paths
        self.current_dir = os.path.expanduser("~")
        
        # Top frame for directory navigation
        nav_frame = ttk.Frame(root)
        nav_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(nav_frame, text="Current Directory:").pack(side=tk.LEFT)
        self.dir_label = ttk.Label(nav_frame, text=self.current_dir, foreground="blue")
        self.dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        ttk.Button(nav_frame, text="Browse Folder", command=self.browse_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="Select All", command=self.select_all).pack(side=tk.LEFT, padx=5)
        
        # Main content frame (two columns)
        content_frame = ttk.Frame(root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left column - Files in current directory
        left_frame = ttk.LabelFrame(content_frame, text="Files in Current Directory")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # Listbox for available PDFs
        scrollbar_left = ttk.Scrollbar(left_frame)
        scrollbar_left.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(left_frame, yscrollcommand=scrollbar_left.set, selectmode=tk.MULTIPLE)
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar_left.config(command=self.file_listbox.yview)
        
        # Right column - Selected files
        right_frame = ttk.LabelFrame(content_frame, text="Selected Files")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Listbox for selected PDFs
        scrollbar_right = ttk.Scrollbar(right_frame)
        scrollbar_right.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.selected_listbox = tk.Listbox(right_frame, yscrollcommand=scrollbar_right.set, selectmode=tk.MULTIPLE)
        self.selected_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar_right.config(command=self.selected_listbox.yview)
        
        # Middle buttons
        middle_frame = ttk.Frame(content_frame)
        middle_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        ttk.Button(middle_frame, text="Add ➜", command=self.add_files).pack(fill=tk.X, pady=5)
        ttk.Button(middle_frame, text="Remove ➜", command=self.remove_files).pack(fill=tk.X, pady=5)
        
        # Bottom frame for action buttons
        button_frame = ttk.Frame(root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Extract Annotations", command=self.extract).pack(side=tk.RIGHT, padx=5)
        
        self.refresh_file_list()
    
    def browse_folder(self):
        """Browse and select a folder"""
        folder = filedialog.askdirectory(initialdir=self.current_dir)
        if folder:
            self.current_dir = folder
            self.dir_label.config(text=self.current_dir)
            self.refresh_file_list()
    
    def select_all(self):
        """Select all PDFs in current directory"""
        self.file_listbox.select_set(0, tk.END)
        self.add_files()
    
    def refresh_file_list(self):
        """Refresh the list of PDFs in current directory"""
        self.file_listbox.delete(0, tk.END)
        
        try:
            files = sorted([f for f in os.listdir(self.current_dir) if f.lower().endswith('.pdf')])
            for file in files:
                self.file_listbox.insert(tk.END, file)
        except PermissionError:
            messagebox.showerror("Error", "Permission denied to access this folder")
    
    def add_files(self):
        """Add selected files from current directory to selected list"""
        selected_indices = self.file_listbox.curselection()
        
        for idx in selected_indices:
            file_name = self.file_listbox.get(idx)
            full_path = os.path.join(self.current_dir, file_name)
            
            # Check if already in selected list
            if full_path not in self.selected_files:
                self.selected_files.append(full_path)
        
        self.refresh_selected_list()
    
    def remove_files(self):
        """Remove selected files from selected list"""
        selected_indices = self.selected_listbox.curselection()
        
        # Remove in reverse order to avoid index shifting
        for idx in reversed(selected_indices):
            self.selected_files.pop(idx)
        
        self.refresh_selected_list()
    
    def refresh_selected_list(self):
        """Refresh the display of selected files"""
        self.selected_listbox.delete(0, tk.END)
        
        for file_path in self.selected_files:
            display_text = os.path.basename(file_path)
            self.selected_listbox.insert(tk.END, display_text)
    
    def clear_all(self):
        """Clear all selected files"""
        if messagebox.askyesno("Confirm", "Clear all selected files?"):
            self.selected_files = []
            self.refresh_selected_list()
    
    def extract(self):
        """Extract annotations from selected PDFs"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select at least one PDF file")
            return
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = get_next_filename(os.path.join(script_dir, "all_pdf_annotations"))
        
        try:
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                processed_count = 0
                sheet_names_used = []
                
                for idx, pdf_file in enumerate(self.selected_files):
                    try:
                        pdf_name = os.path.basename(pdf_file)
                        annotations = extract_all_annotations(pdf_file)
                        
                        # Sanitize sheet name (max 31 chars, no special chars)
                        sheet_name = pdf_name[:31].replace('[', '').replace(']', '').replace(':', '').replace('*', '').replace('?', '').replace('/', '').replace('\\', '')
                        
                        # Handle duplicate sheet names
                        original_sheet_name = sheet_name
                        counter = 1
                        while sheet_name in sheet_names_used:
                            sheet_name = f"{original_sheet_name[:25]}_{counter}"
                            counter += 1
                        
                        sheet_names_used.append(sheet_name)
                        
                        if annotations:
                            df = pd.DataFrame(annotations)
                            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
                            
                            # Add PDF name at top
                            worksheet = writer.sheets[sheet_name]
                            worksheet['A1'] = f"PDF: {pdf_name}"
                            worksheet['A1'].font = Font(bold=True, size=12)
                            
                            # Format header row (now at row 3 instead of 1)
                            for cell in worksheet[3]:
                                cell.font = Font(bold=True)
                            
                            # Auto-adjust column widths
                            for column in worksheet.columns:
                                max_length = 0
                                column_letter = column[0].column_letter
                                for cell in column:
                                    try:
                                        if cell.value:
                                            max_length = max(max_length, len(str(cell.value)))
                                    except:
                                        pass
                                adjusted_width = min(max_length + 2, 50)
                                worksheet.column_dimensions[column_letter].width = adjusted_width
                        else:
                            # Create empty sheet with message
                            df = pd.DataFrame({"Message": ["No annotations found in this PDF"]})
                            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
                            
                            worksheet = writer.sheets[sheet_name]
                            worksheet['A1'] = f"PDF: {pdf_name}"
                            worksheet['A1'].font = Font(bold=True, size=12)
                        
                        processed_count += 1
                    
                    except Exception as file_error:
                        messagebox.showwarning("File Error", f"Error processing {os.path.basename(pdf_file)}:\n{str(file_error)}\n\nContinuing with other files...")
                        continue
            
            messagebox.showinfo("Extraction Complete ✅", f"Successfully processed {processed_count} out of {len(self.selected_files)} PDF(s)\n\nFile saved to:\n{excel_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during extraction:\n{str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = PDFSelector(root)
    root.mainloop()
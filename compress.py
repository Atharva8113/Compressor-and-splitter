import os
import io
import shutil
import tempfile
import subprocess
import numpy as np
import fitz  # PyMuPDF
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk

# Default target size (can be changed via GUI)
DEFAULT_TARGET_MB = 1.9

# ------------------------------------------------------
# UTILITY
# ------------------------------------------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_file_size(path):
    try:
        return os.path.getsize(path)
    except FileNotFoundError:
        return 0

def is_scanned_pdf(pdf):
    """Decide if PDF is scanned based on image density. 
    (Used to recommend default modes, but user overrides)."""
    try:
        doc = fitz.open(pdf)
        img_pages = 0
        for p in doc:
            if p.get_images():
                img_pages += 1
        scanned = img_pages >= (len(doc) * 0.4)
        doc.close()
        return scanned
    except:
        return True

# ------------------------------------------------------
# COMPRESSION METHODS
# ------------------------------------------------------

def compress_standard_gs(pdf_path, output_path):
    """Standard compression using Ghostscript (ebook quality). Good for general use."""
    gs_commands = ["gswin64c", "gswin32c", "gs"]
    
    success = False
    for gs_exe in gs_commands:
        try:
            command = [
                gs_exe,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/ebook", # /screen is lower, /printer is higher
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                pdf_path
            ]
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            ret = subprocess.call(command, startupinfo=startupinfo)
            if ret == 0 and os.path.exists(output_path) and get_file_size(output_path) > 0:
                success = True
                break
        except Exception:
            continue
    return success

def compress_pymupdf_optimize(pdf_path, output_path):
    """Fallback compression: garbage collection and stream deflation. Lossless-ish."""
    try:
        doc = fitz.open(pdf_path)
        doc.save(output_path, deflate=True, garbage=4)
        doc.close()
        return True
    except:
        return False

def compress_binary_bw(pdf_path, output_path):
    """Extreme compression: Rasterize to 1-bit B&W images."""
    try:
        doc = fitz.open(pdf_path)
        new_pdf = fitz.open()

        for page in doc:
            # 200 DPI is a balance. 150 might be too low output for text.
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.convert("L") # Grayscale

            # Threshold
            arr = np.array(img)
            thresh = np.mean(arr) * 0.92
            binary = (arr > thresh) * 255
            img_bw = Image.fromarray(binary.astype(np.uint8), 'L')

            temp = io.BytesIO()
            img_bw.save(temp, format="PNG", optimize=True, compress_level=9)
            img_bytes = temp.getvalue()

            rect = page.rect
            pdf_page = new_pdf.new_page(width=rect.width, height=rect.height)
            pdf_page.insert_image(rect, stream=img_bytes)

        new_pdf.save(output_path, deflate=True, garbage=4)
        new_pdf.close()
        doc.close()
        return True
    except Exception as e:
        print(f"Error in binary compression: {e}")
        return False

# ------------------------------------------------------
# SPLIT LOGIC
# ------------------------------------------------------

def split_pdf(pdf_path, output_folder, base_name, max_size_bytes):
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    outputs = []
    
    start = 0
    part_index = 1

    while start < total_pages:
        part_doc = fitz.open()
        p = start
        
        while p < total_pages:
            temp = fitz.open()
            temp.insert_pdf(doc, from_page=p, to_page=p)
            
            # Always add first page of a chunk
            if part_doc.page_count == 0:
                part_doc.insert_pdf(temp)
                p += 1
                continue

            # Trial save for size check
            trial_doc = fitz.open()
            trial_doc.insert_pdf(part_doc)
            trial_doc.insert_pdf(temp)
            
            trial_path = os.path.join(tempfile.gettempdir(), f"__trial_{base_name}.pdf")
            trial_doc.save(trial_path, deflate=True, garbage=4)
            trial_size = get_file_size(trial_path)
            trial_doc.close()

            if trial_size > max_size_bytes:
                break
            
            part_doc.insert_pdf(temp)
            p += 1

        out_path = os.path.join(output_folder, f"{base_name}_part{part_index}.pdf")
        part_doc.save(out_path, deflate=True, garbage=4)
        outputs.append(out_path)

        start = p
        part_index += 1

    doc.close()
    return outputs

# ------------------------------------------------------
# GUI APPLICATION
# ------------------------------------------------------

class NagarkotCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Nagarkot PDF Tool")
        self.root.geometry("900x700")
        self.root.configure(bg="#ffffff") # White background

        # Variables
        self.pdf_list = []
        self.output_dir_var = tk.StringVar()
        self.action_var = tk.StringVar(value="Split Only")
        self.comp_level_var = tk.StringVar(value="Standard (GS)")
        self.size_var = tk.StringVar(value=str(DEFAULT_TARGET_MB))

        self.setup_styles()
        self.setup_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")  # Valid starting point for customization

        # Colors
        bg_color = "#ffffff"
        fg_color = "#333333"
        accent_color = "#0056b3"
        
        # General configurations
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=fg_color, font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10), padding=6)
        style.configure("TLabelframe", background=bg_color, relief="groove")
        style.configure("TLabelframe.Label", background=bg_color, foreground=accent_color, font=("Arial", 11, "bold"))
        
        # Primary Button (Blue)
        style.configure("Primary.TButton", background=accent_color, foreground="white", borderwidth=0)
        style.map("Primary.TButton", background=[("active", "#004494")])
        
        # Secondary Button (White/Gray)
        style.configure("Secondary.TButton", background="#f0f0f0", foreground="#333333", bordercolor="#cccccc")
        style.map("Secondary.TButton", background=[("active", "#e0e0e0")])

        # Treeview
        style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
        style.map("Treeview", background=[("selected", accent_color)], foreground=[("selected", "white")])
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background="#f0f0f0")

    def setup_ui(self):
        # 1. HEADER
        header_frame = tk.Frame(self.root, bg="white", height=80)
        header_frame.pack(fill="x", padx=20, pady=10)
        
        # Logo
        try:
            # Try loading logo from resource path (works for exe and dev)
            logo_path = resource_path("Nagarkot Logo.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                # Resize to ~50px height
                h_size = 20
                w_size = int((h_size / img.height) * img.width)
                img = img.resize((w_size, h_size), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header_frame, image=self.logo_img, bg="white").pack(side="left", padx=(0, 15))
            else:
                tk.Label(header_frame, text="[LOGO MISSING]", bg="white", fg="red").pack(side="left", padx=(0, 15))
                print("Warning: Nagarkot Logo.png not found.")
        except Exception as e:
            tk.Label(header_frame, text="[LOGO ERROR]", bg="white", fg="red").pack(side="left", padx=(0, 15))
            print(f"Error loading logo: {e}")

        # Title/Subtitle
        title_frame = tk.Frame(header_frame, bg="white")
        title_frame.pack(side="left", fill="y", expand=True)
        tk.Label(title_frame, text="Compressor & Splitter", font=("Helvetica", 18, "bold"), fg="#0056b3", bg="white").pack(anchor="w")
        tk.Label(title_frame, text="Document Processing Utility", font=("Helvetica", 10), fg="#777777", bg="white").pack(anchor="w")

        # 2. MAIN CONTENT
        main_content = tk.Frame(self.root, bg="white")
        main_content.pack(fill="both", expand=True, padx=20, pady=5)

        # Left Column (Controls)
        controls_frame = tk.Frame(main_content, bg="white")
        controls_frame.pack(side="left", fill="y", padx=(0, 10), anchor="n")

        # -> File Selection Section
        fs_frame = ttk.LabelFrame(controls_frame, text="File Selection", padding=15)
        fs_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Button(fs_frame, text="Select PDFs", style="Secondary.TButton", command=self.select_pdfs).pack(fill="x", pady=5)
        
        self.file_count_label = ttk.Label(fs_frame, text="No files selected")
        self.file_count_label.pack(anchor="w")

        # -> Options Section
        opt_frame = ttk.LabelFrame(controls_frame, text="Configuration", padding=15)
        opt_frame.pack(fill="x", pady=(0, 15))

        # Output Folder
        ttk.Label(opt_frame, text="Output Folder:").pack(anchor="w", pady=(5,0))
        out_box = tk.Frame(opt_frame, bg="white")
        out_box.pack(fill="x", pady=5)
        ttk.Entry(out_box, textvariable=self.output_dir_var).pack(side="left", fill="x", expand=True)
        ttk.Button(out_box, text="...", width=3, style="Secondary.TButton", command=self.select_folder).pack(side="right", padx=(5,0))

        # Action
        ttk.Label(opt_frame, text="Action Mode:").pack(anchor="w", pady=(10,0))
        ttk.Combobox(opt_frame, textvariable=self.action_var, state="readonly", 
                     values=["Split Only", "Compress Only", "Compress + Split"]).pack(fill="x", pady=5)

        # Compression Level
        ttk.Label(opt_frame, text="Compression Level:").pack(anchor="w", pady=(10,0))
        ttk.Combobox(opt_frame, textvariable=self.comp_level_var, state="readonly", 
                     values=["Standard (GS)", "Extreme (B&W)"]).pack(fill="x", pady=5)
        
        # Split Limit
        ttk.Label(opt_frame, text="Split Limit (MB):").pack(anchor="w", pady=(10,0))
        ttk.Spinbox(opt_frame, from_=0.5, to=50, increment=0.1, textvariable=self.size_var).pack(fill="x", pady=5)

        # Right Column (Data Table)
        table_frame = ttk.LabelFrame(main_content, text="Preview & Status", padding=10)
        table_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        cols = ("file", "size", "status")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("file", text="File Name")
        self.tree.heading("size", text="Size (KB)")
        self.tree.heading("status", text="Status")
        
        self.tree.column("file", width=250)
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("status", width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 3. FOOTER
        footer_frame = tk.Frame(self.root, bg="#f0f0f0", height=50)
        footer_frame.pack(fill="x", side="bottom")
        
        # Copyright
        tk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", bg="#f0f0f0", fg="#888888", font=("Arial", 9)).pack(side="left", padx=20, pady=15)
        
        # Action Buttons
        btn_frame = tk.Frame(footer_frame, bg="#f0f0f0")
        btn_frame.pack(side="right", padx=20, pady=10)
        
        ttk.Button(btn_frame, text="Clear", style="Secondary.TButton", command=self.clear_selection).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Generate/Run", style="Primary.TButton", command=self.run_process).pack(side="left", padx=5)

    def select_folder(self):
        d = filedialog.askdirectory()
        if d: 
            self.output_dir_var.set(d)

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            for f in files:
                if f not in self.pdf_list:
                    self.pdf_list.append(f)
                    sz = get_file_size(f) / 1024
                    self.tree.insert("", "end", values=(os.path.basename(f), f"{sz:.1f}", "Pending"))
            
            self.file_count_label.config(text=f"{len(self.pdf_list)} file(s) loaded")

    def clear_selection(self):
        self.pdf_list = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.file_count_label.config(text="No files selected")

    def update_status(self, filename_base, msg):
        # Find item in tree
        # This is a bit inefficient (O(N)), but N is small
        for item in self.tree.get_children():
            vals = self.tree.item(item)['values']
            if vals[0] == filename_base:
                self.tree.item(item, values=(vals[0], vals[1], msg))
                self.root.update_idletasks()
                break

    def run_process(self):
        if not self.pdf_list:
            messagebox.showerror("Error", "No files selected.")
            return
        
        output_folder = self.output_dir_var.get()
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        split_mb = float(self.size_var.get())
        act = self.action_var.get()
        clevel = self.comp_level_var.get()

        for pdf_path in self.pdf_list:
            base_name = os.path.basename(pdf_path)
            self.update_status(base_name, "Processing...")
            
            try:
                # Logic copied from process_one_file but tailored for this class
                base = os.path.splitext(base_name)[0]
                temp_comp = os.path.join(tempfile.gettempdir(), base + "_comp_temp.pdf")
                final_output = os.path.join(output_folder, base + "_processed.pdf")

                current_source = pdf_path
                should_compress = (act == "Compress Only" or act == "Compress + Split")

                if should_compress:
                    self.update_status(base_name, f"Compressing ({clevel})...")
                    success = False
                    if clevel == "Extreme (B&W)":
                        success = compress_binary_bw(pdf_path, temp_comp)
                    elif clevel == "Standard (GS)":
                        success = compress_standard_gs(pdf_path, temp_comp)
                        if not success:
                            success = compress_pymupdf_optimize(pdf_path, temp_comp)
                    
                    if success and os.path.exists(temp_comp) and get_file_size(temp_comp) > 0:
                        current_source = temp_comp
                    else:
                        # Compression failed or not effective
                        pass
                
                if "Split" in act:
                    self.update_status(base_name, "Splitting...")
                    target_bytes = int(split_mb * 1024 * 1024)
                    split_pdf(current_source, output_folder, base, target_bytes)
                    self.update_status(base_name, "Done (Split)")
                else:
                    # Save
                    if current_source != pdf_path:
                        shutil.move(current_source, final_output)
                    else:
                        shutil.copy(pdf_path, final_output)
                    self.update_status(base_name, "Done (Saved)")
                    
            except Exception as e:
                print(e)
                self.update_status(base_name, "Error")

        messagebox.showinfo("Success", "All tasks completed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = NagarkotCompressorApp(root)
    root.mainloop()

"""OneNote to Markdown Converter - GUI"""

import customtkinter as ctk
import os
import sys
import subprocess
import threading
from pathlib import Path
from tkinter import filedialog

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def log(msg):
    """Print to console for debugging."""
    print(f"[LOG] {msg}")
    sys.stdout.flush()


class OneNote2MDApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("OneNote to Markdown")
        self.geometry("900x700")
        self.cfg = self.load_config()
        self.selected_files = []
        self.is_exporting = False
        self.setup_ui()
        log("App initialized")

    def load_config(self):
        import json
        config_dir = Path.home() / ".onenote2md"
        config_file = config_dir / "config.json"
        default = {"source_folder": "", "output_dir": "./output"}
        if config_file.exists():
            with open(config_file, 'r') as f:
                cfg = json.load(f)
                return {**default, **cfg}
        return default

    def save_config(self, cfg):
        import json
        config_dir = Path.home() / ".onenote2md"
        config_dir.mkdir(parents=True, exist_ok=True)
        with open(config_dir / "config.json", 'w') as f:
            json.dump(cfg, f)

    def setup_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main, text="OneNote to Markdown", font=("Arial", 24, "bold")).pack(pady=(0, 10))
        ctk.CTkLabel(main, text="Select .one files -> PDF -> Markdown", font=("Arial", 12)).pack(pady=(0, 10))

        sf = ctk.CTkFrame(main)
        sf.pack(fill="x", pady=10)
        ctk.CTkLabel(sf, text="Files:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.source_entry = ctk.CTkEntry(sf, width=400)
        self.source_entry.pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(sf, text="Browse", command=self.browse).pack(side="right", padx=10)

        of = ctk.CTkFrame(main)
        of.pack(fill="x", pady=10)
        ctk.CTkLabel(of, text="Output:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.output_entry = ctk.CTkEntry(of, width=400)
        self.output_entry.insert(0, self.cfg.get("output_dir", "./output"))
        self.output_entry.pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(of, text="Browse", command=self.browse_out).pack(side="right", padx=10)

        ff = ctk.CTkFrame(main)
        ff.pack(fill="both", expand=True, pady=10)
        ctk.CTkLabel(ff, text="Selected:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.files_list = ctk.CTkTextbox(ff)
        self.files_list.pack(fill="both", expand=True, padx=10, pady=5)

        self.progress = ctk.CTkProgressBar(main)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)

        self.status = ctk.CTkLabel(main, text="Ready", font=("Arial", 12))
        self.status.pack(pady=5)

        self.export_btn = ctk.CTkButton(main, text="Convert to Markdown", font=("Arial", 14, "bold"), height=40, command=self.start_export)
        self.export_btn.pack(pady=10)

    def browse(self):
        files = filedialog.askopenfilenames(title="Select Files", filetypes=[("OneNote", "*.one"), ("PDF", "*.pdf")])
        if files:
            self.selected_files = list(files)
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, f"{len(files)} files")
            self.refresh_list()
            log(f"Selected {len(files)} files")

    def browse_out(self):
        folder = filedialog.askdirectory(title="Output Folder")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.cfg["output_dir"] = folder
            self.save_config(self.cfg)
            log(f"Output folder: {folder}")

    def refresh_list(self):
        self.files_list.delete("1.0", "end")
        for f in self.selected_files:
            self.files_list.insert("end", f"{Path(f).name}\n")

    def start_export(self):
        if self.is_exporting or not self.selected_files:
            log("Export skipped: already exporting or no files")
            return
        self.is_exporting = True
        self.export_btn.configure(state="disabled", text="Converting...")
        log("Starting export...")
        threading.Thread(target=self.run_export).start()

    def run_export(self):
        output = self.output_entry.get()
        log(f"Output directory: {output}")
        os.makedirs(output, exist_ok=True)
        total = len(self.selected_files)
        log(f"Total files to process: {total}")
        
        for i, f in enumerate(self.selected_files, 1):
            log(f"Processing file {i}/{total}: {f}")
            self.after(0, lambda p=i/total: self.progress.set(p))
            ext = Path(f).suffix.lower()
            pdf = None
            
            if ext == ".pdf":
                log(f"Direct PDF conversion: {f}")
                self.after(0, lambda n=Path(f).name: self.status.configure(text=f"Converting PDF: {n}"))
                self.convert_pdf(f, output)
            else:
                log(f"Converting .one to PDF: {f}")
                self.after(0, lambda n=Path(f).name: self.status.configure(text=f"Converting .one to PDF: {n}"))
                pdf = self.convert_one_to_pdf(f, output)
                log(f"PDF result: {pdf}")
                if pdf and os.path.exists(pdf):
                    log(f"PDF exists, converting to MD: {pdf}")
                    self.after(0, lambda n=Path(pdf).name: self.status.configure(text=f"Converting PDF to MD: {n}"))
                    self.convert_pdf(pdf, output)
                    log(f"Markdown created from PDF")
                else:
                    log(f"PDF creation failed, using fallback direct extract")
                    self.after(0, lambda n=Path(f).name: self.status.configure(text=f"Fallback extract: {n}"))
                    self.convert_one_direct(f, output)

        log("Export complete!")
        self.after(0, lambda: self.status.configure(text="Done!"))
        self.after(0, lambda: self.export_btn.configure(state="normal", text="Convert to Markdown"))
        self.after(0, lambda: self.progress.set(1))
        self.is_exporting = False

    def convert_one_to_pdf(self, one_file, output_dir):
        log(f"convert_one_to_pdf called with: {one_file}")
        pdf_path = os.path.join(output_dir, Path(one_file).stem + ".pdf")
        log(f"Target PDF path: {pdf_path}")
        
        # Method 1: Try OneNote via COM with proper error handling
        one_abs = os.path.abspath(one_file)
        pdf_abs = os.path.abspath(pdf_path)
        
        # First, try opening OneNote and the file programmatically
        log("Method 1: Starting OneNote and opening file...")
        ps1 = f'''
$one = New-Object -ComObject OneNote.Application
$one.Windows.CurrentWindow.NavigateTo("{one_abs}")
Start-Sleep -Seconds 2
'''
        result1 = subprocess.run(["powershell", "-Command", ps1], capture_output=True, text=True, timeout=30)
        log(f"Method 1 result: {result1.stderr}")
        
        # Method 2: Try using Shell to print
        log("Method 2: Using Shell print command...")
        ps2 = f'''
$shell = New-Object -ComObject Shell.Application
$folder = $shell.NameSpace((Split-Path "{one_abs}"))
$item = $folder.ParseName((Split-Path "{one_abs}" -Leaf))
$item.InvokeVerb("Print")
'''
        result2 = subprocess.run(["powershell", "-Command", ps2], capture_output=True, text=True, timeout=30)
        log(f"Method 2 result: {result2.stderr}")
        
        # Method 3: Direct file path that OneNote can open
        # OneNote requires the file to be accessible - try local copy first
        log("Method 3: Try with Windows API...")
        ps3 = f'''
[System.Runtime.InteropServices.Marshal]::GetActiveObject("OneNote.Application") | ForEach-Object {{
    $_.Windows.CurrentWindow.NavigateTo("{one_abs}")
}}
'''
        result3 = subprocess.run(["powershell", "-Command", ps3], capture_output=True, text=True, timeout=30)
        log(f"Method 3 result: {result3.stderr}")
        
        # Method 4: Use OneNote's built-in converter via command line
        log("Method 4: Check if OneNote is installed...")
        ps4 = '''Get-ItemProperty "HKLM:\\SOFTWARE\\Microsoft\\Office\\*\\OneNote" -ErrorAction SilentlyContinue | Select-Object -First 1'''
        result4 = subprocess.run(["powershell", "-Command", ps4], capture_output=True, text=True, timeout=15)
        log(f"OneNote installation: {result4.stdout}")
        
        # Method 5: Use NirCmd or similar for print-to-PDF
        log("Method 5: Check for print drivers...")
        ps5 = '''Get-Printer | Where-Object {$_.Name -like "*PDF*"}'''
        result5 = subprocess.run(["powershell", "-Command", ps5], capture_output=True, text=True, timeout=15)
        log(f"PDF printers: {result5.stdout}")
        
        log(f"All OneNote COM methods failed. Returning None to use fallback.")
        return None

    def convert_pdf(self, pdf_path, output_dir):
        log(f"convert_pdf called with: {pdf_path}")
        text = ""
        try:
            from pdfminer.high_level import extract_text
            text = extract_text(pdf_path)
            log(f"pdfminer extracted {len(text)} chars")
        except Exception as e:
            log(f"pdfminer failed: {e}")
            try:
                import PyPDF2
                with open(pdf_path, 'rb') as pf:
                    reader = PyPDF2.PdfReader(pf)
                    text = "\n\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
                    log(f"PyPDF2 extracted {len(text)} chars")
            except Exception as e2:
                log(f"PyPDF2 failed: {e2}")
                pass
        if not text:
            text = "[Could not extract]"
            log("No text extracted from PDF")
        
        md = f"# {Path(pdf_path).stem}\n\n" + self.format_text(text)
        out = Path(output_dir) / f"{Path(pdf_path).stem}.md"
        with open(out, 'w', encoding='utf-8') as f:
            f.write(md)
        log(f"Markdown saved to: {out}")

    def format_text(self, text):
        import re
        lines = []
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                lines.append("")
            elif len(line) < 50 and line.isupper():
                lines.append(f"## {line}")
            elif len(line) < 40 and not line.endswith(('.', ',')):
                lines.append(f"### {line}")
            else:
                lines.append(line)
        return '\n'.join(lines)

    def convert_one_direct(self, one_file, output_dir):
        log(f"convert_one_direct called with: {one_file}")
        try:
            with open(one_file, 'rb') as f:
                data = f.read()
            log(f"Read {len(data)} bytes from .one file")
            
            # Try many encodings for Unicode/UTF-8 support
            encodings = [
                'utf-16-le', 'utf-16-be', 'utf-16',
                'utf-8', 'utf-8-sig',
                'cp1252', 'latin-1',
                'iso-8859-1', 'cp1250', 'cp1251',
                'big5', 'shift_jis', 'euc-kr'
            ]
            
            text = None
            used_encoding = None
            for enc in encodings:
                try:
                    text = data.decode(enc, errors='ignore')
                    # Check if it looks like valid text (not mostly nulls or control chars)
                    if len(text) > 100:
                        valid_ratio = sum(1 for c in text[:1000] if c.isprintable() or c in '\n\r\t ') / min(1000, len(text))
                        if valid_ratio > 0.7:
                            used_encoding = enc
                            log(f"Decoded with {enc}: {len(text)} chars, validity: {valid_ratio:.2f}")
                            break
                except:
                    continue
            
            if text is None or len(text) < 100:
                text = data.decode('utf-8', errors='replace')
                used_encoding = 'utf-8-replace'
                log(f"Fallback to {used_encoding}")
            
            # Clean up the extracted text
            lines = []
            for line in text.split('\n'):
                line = line.strip()
                # Keep lines with meaningful content
                if line and len(line) > 1:
                    # Try to clean up garbled XML artifacts
                    line = line.replace('\x00', '')
                    if len(line) > 1:
                        lines.append(line)
            
            text = '\n'.join(lines[:500])  # Limit to 500 lines
            
            md = f"# {Path(one_file).stem}\n\n{text[:15000]}\n"
            out = Path(output_dir) / f"{Path(one_file).stem}.md"
            with open(out, 'w', encoding='utf-8') as f:
                f.write(md)
            log(f"Direct extract saved to: {out} (encoding: {used_encoding})")
        except Exception as e:
            log(f"ERROR in convert_one_direct: {e}")


if __name__ == "__main__":
    log("Starting app...")
    app = OneNote2MDApp()
    app.mainloop()

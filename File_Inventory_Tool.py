import os
import subprocess
import threading
import queue
import re
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Helper 🚀")
        self.root.geometry("860x680")

        self.excel_path = None
        self.ui_queue = queue.Queue()

        # Path variables (can be typed directly by user)
        self.pdf_path_var = tk.StringVar()
        self.file_var = tk.StringVar()
        self.save_var = tk.StringVar()
        self.use_same_folder_var = tk.BooleanVar(value=True)
        self.command_var = tk.StringVar(value='echo Processing {file}')

        self._build_ui()

        # Keep button states in sync with typed paths
        self.pdf_path_var.trace_add("write", self.on_pdf_path_change)
        self.file_var.trace_add("write", lambda *_: self.update_button_states())
        self.save_var.trace_add("write", lambda *_: self.update_button_states())

        self.toggle_save_mode()
        self.update_button_states()
        self.root.after(100, self.process_ui_queue)

    def _build_ui(self):
        lbl_title = tk.Label(self.root, text="PDF Invoice Extractor", font=("Arial", 16, "bold"), pady=10)
        lbl_title.pack()

        # Source PDF row
        frame_pdf = tk.Frame(self.root)
        frame_pdf.pack(pady=6, padx=20, fill="x")
        tk.Label(frame_pdf, text="PDF File Path:", width=16, anchor="w").pack(side="left")
        self.entry_pdf = tk.Entry(frame_pdf, textvariable=self.pdf_path_var)
        self.entry_pdf.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(frame_pdf, text="Browse", command=self.browse_pdf_file, bg="#e1e1e1", width=10).pack(side="right")

        # Save location row
        frame_save = tk.Frame(self.root)
        frame_save.pack(pady=6, padx=20, fill="x")
        tk.Label(frame_save, text="Save Excel As:", width=16, anchor="w").pack(side="left")
        self.entry_save = tk.Entry(frame_save, textvariable=self.save_var)
        self.entry_save.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_browse_save = tk.Button(
            frame_save,
            text="Browse",
            command=self.browse_save_location,
            bg="#e1e1e1",
            width=10,
        )
        self.btn_browse_save.pack(side="right")

        frame_save_mode = tk.Frame(self.root)
        frame_save_mode.pack(pady=(0, 6), padx=20, fill="x")
        tk.Checkbutton(
            frame_save_mode,
            text="Use same folder for save output",
            variable=self.use_same_folder_var,
            command=self.toggle_save_mode,
            anchor="w",
        ).pack(side="left")

        # File path row (command input file)
        frame_file = tk.Frame(self.root)
        frame_file.pack(pady=6, padx=20, fill="x")
        tk.Label(frame_file, text="Command File:", width=16, anchor="w").pack(side="left")
        self.entry_file = tk.Entry(frame_file, textvariable=self.file_var)
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(frame_file, text="Browse", command=self.browse_file, bg="#e1e1e1", width=10).pack(side="right")

        # Action buttons
        frame_actions = tk.Frame(self.root)
        frame_actions.pack(pady=8)

        self.btn_run = tk.Button(
            frame_actions,
            text="▶ Start Extraction",
            command=self.start_extraction_thread,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            width=20,
        )
        self.btn_run.pack(side="left", padx=10)

        self.btn_open = tk.Button(
            frame_actions,
            text="📊 Open Excel",
            command=self.open_excel,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10, "bold"),
            state="disabled",
            width=20,
        )
        self.btn_open.pack(side="left", padx=10)

        # Command runner
        frame_cmd = tk.Frame(self.root)
        frame_cmd.pack(pady=8, padx=20, fill="x")

        tk.Label(frame_cmd, text="Command ({file} will be replaced):", anchor="w").pack(fill="x")
        self.entry_command = tk.Entry(frame_cmd, textvariable=self.command_var)
        self.entry_command.pack(fill="x", pady=(4, 8))

        self.btn_run_cmd = tk.Button(
            frame_cmd,
            text="⚙ Run Command",
            command=self.start_command_thread,
            bg="#FF9800",
            fg="white",
            font=("Arial", 10, "bold"),
            width=20,
        )
        self.btn_run_cmd.pack(anchor="w")

        tk.Label(self.root, text="Process Log:", anchor="w").pack(fill="x", padx=20)
        self.log_area = scrolledtext.ScrolledText(self.root, height=16, state="disabled", font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def post_ui(self, fn, *args, **kwargs):
        self.ui_queue.put((fn, args, kwargs))

    def process_ui_queue(self):
        while True:
            try:
                fn, args, kwargs = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            fn(*args, **kwargs)
        self.root.after(100, self.process_ui_queue)

    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")

    def normalize_path(self, path_text):
        return path_text.strip().strip('"').strip("'")

    def on_pdf_path_change(self, *_):
        self.sync_save_with_source()
        self.update_button_states()

    def sync_save_with_source(self):
        if not self.use_same_folder_var.get():
            return

        pdf_text = self.normalize_path(self.pdf_path_var.get())
        if not pdf_text:
            self.save_var.set("")
            return

        source = Path(pdf_text)
        base_dir = source.parent if source.suffix.lower() == ".pdf" else source
        self.save_var.set(str(base_dir / "PDF_Inventory_GUI.xlsx"))

    def toggle_save_mode(self):
        if self.use_same_folder_var.get():
            self.entry_save.config(state="disabled")
            self.btn_browse_save.config(state="disabled")
            self.sync_save_with_source()
        else:
            self.entry_save.config(state="normal")
            self.btn_browse_save.config(state="normal")

    def browse_pdf_file(self):
        pdf_file = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf *.PDF"), ("All files", "*.*")]
        )
        if pdf_file:
            self.pdf_path_var.set(pdf_file)
            self.log(f"📄 PDF file set: {pdf_file}")

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_var.set(file_path)
            self.log(f"📄 Command file set: {file_path}")

    def browse_save_location(self):
        source_text = self.normalize_path(self.pdf_path_var.get())
        initial_dir = str(Path(source_text).parent) if source_text else None
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialdir=initial_dir,
            initialfile="PDF_Inventory_GUI.xlsx",
            title="Select Excel save location",
        )
        if file_path:
            self.save_var.set(file_path)
            self.log(f"💾 Save location set: {file_path}")

    def update_button_states(self):
        pdf_ok = bool(self.normalize_path(self.pdf_path_var.get()))
        file_ok = bool(self.normalize_path(self.file_var.get()))

        self.btn_run.config(state="normal" if pdf_ok else "disabled")
        self.btn_run_cmd.config(state="normal" if file_ok else "disabled")

    def validate_pdf_file(self):
        pdf_text = self.normalize_path(self.pdf_path_var.get())
        if not pdf_text:
            messagebox.showerror("Error", "Please enter a PDF file path.")
            return None

        pdf_file = Path(pdf_text)
        if not pdf_file.exists() or not pdf_file.is_file():
            messagebox.showerror("Error", "PDF file path is invalid.")
            return None

        if pdf_file.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Selected file must be a PDF.")
            return None

        return pdf_file

    def validate_file(self):
        file_text = self.normalize_path(self.file_var.get())
        if not file_text:
            messagebox.showerror("Error", "Please enter a file path.")
            return None

        selected_file = Path(file_text)
        if not selected_file.exists() or not selected_file.is_file():
            messagebox.showerror("Error", "File path is invalid.")
            return None

        return selected_file

    def resolve_save_path(self, source_pdf):
        source_dir = source_pdf.parent
        if self.use_same_folder_var.get():
            save_path = source_dir / "PDF_Inventory_GUI.xlsx"
            self.post_ui(self.save_var.set, str(save_path))
            return save_path

        save_text = self.normalize_path(self.save_var.get())
        if not save_text:
            save_path = source_dir / "PDF_Inventory_GUI.xlsx"
            self.post_ui(self.save_var.set, str(save_path))
            return save_path

        save_path = Path(save_text)
        if save_path.suffix.lower() != ".xlsx":
            save_path = save_path.with_suffix(".xlsx")
        return save_path

    def start_extraction_thread(self):
        pdf_file = self.validate_pdf_file()
        if pdf_file is None:
            return

        self.btn_run.config(state="disabled")
        self.btn_open.config(state="disabled")
        threading.Thread(target=self.run_extraction, args=(pdf_file,), daemon=True).start()

    def run_extraction(self, pdf_file):
        self.post_ui(self.log, "-" * 40)
        self.post_ui(self.log, f"🚀 Starting extraction from: {pdf_file}")

        try:
            pdf_files = [pdf_file]
            self.post_ui(self.log, "✅ Found 1 PDF file.")

            data = []
            for i, pdf in enumerate(pdf_files, 1):
                parsed = self.parse_filename(pdf.name)
                safe_path = str(pdf).replace('"', '""')
                link = f'=HYPERLINK("{safe_path}", "Open PDF")'

                row = {
                    "S.No": i,
                    "Document Number": parsed["doc_num"],
                    "Company Name": parsed["company"],
                    "Date": parsed["formatted_date"],
                    "File Name": pdf.name,
                    "Link": link,
                    "Size (KB)": round(pdf.stat().st_size / 1024, 2),
                }
                data.append(row)

                if i % 5 == 0 or i == len(pdf_files):
                    self.post_ui(self.log, f"   Processed {i}/{len(pdf_files)}...")

            save_path = self.resolve_save_path(pdf_file)
            if not save_path.parent.exists():
                self.post_ui(self.log, "❌ Save location folder does not exist.")
                self.post_ui(self.btn_run.config, state="normal")
                return

            self.post_ui(self.log, f"💾 Saving Excel to: {save_path}")

            df = pd.DataFrame(data)
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Inventory", index=False)
                ws = writer.sheets["Inventory"]
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 20

            self.excel_path = save_path
            self.post_ui(self.log, "🎉 DONE! Excel file created successfully.")
            self.post_ui(self.btn_run.config, state="normal")
            self.post_ui(self.btn_open.config, state="normal")
            self.post_ui(messagebox.showinfo, "Success", "Extraction Complete!")

        except Exception as e:
            self.post_ui(self.log, f"❌ Error: {str(e)}")
            self.post_ui(self.btn_run.config, state="normal")

    def start_command_thread(self):
        command_text = self.command_var.get().strip()
        if not command_text:
            messagebox.showerror("Error", "Please enter a command.")
            return

        selected_file = self.validate_file()
        if selected_file is None:
            return

        self.btn_run_cmd.config(state="disabled")
        threading.Thread(target=self.run_command, args=(selected_file,), daemon=True).start()

    def run_command(self, selected_file):
        try:
            command_template = self.command_var.get().strip()
            command = command_template.replace("{file}", f'"{selected_file}"')

            source_text = self.normalize_path(self.pdf_path_var.get())
            cwd = None
            if source_text:
                source_path = Path(source_text)
                if source_path.exists():
                    cwd = str(source_path.parent) if source_path.is_file() else str(source_path)

            self.post_ui(self.log, "-" * 40)
            self.post_ui(self.log, f"⚙ Running command: {command}")

            completed = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                cwd=cwd,
            )

            if completed.stdout:
                self.post_ui(self.log, "--- STDOUT ---")
                for line in completed.stdout.strip().splitlines():
                    self.post_ui(self.log, line)

            if completed.stderr:
                self.post_ui(self.log, "--- STDERR ---")
                for line in completed.stderr.strip().splitlines():
                    self.post_ui(self.log, line)

            self.post_ui(self.log, f"Command finished with exit code: {completed.returncode}")

        except Exception as e:
            self.post_ui(self.log, f"❌ Command error: {str(e)}")
        finally:
            self.post_ui(self.btn_run_cmd.config, state="normal")

    def open_excel(self):
        if self.excel_path and self.excel_path.exists():
            os.startfile(self.excel_path)
        else:
            messagebox.showerror("Error", "File not found!")

    def parse_filename(self, filename):
        name = Path(filename).stem
        match = re.match(r"^([^-]+)-(.*)-([\d_.-]+)$", name)

        if match:
            date_str = match.group(3).replace("_", "/").replace("-", "/").replace(".", "/")
            return {
                "doc_num": match.group(1).strip(),
                "company": match.group(2).replace("_", " ").strip(),
                "formatted_date": date_str,
            }

        return {"doc_num": "?", "company": name, "formatted_date": "N/A"}


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

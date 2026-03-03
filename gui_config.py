"""
Thai Certificate Generator - GUI Configuration Tool
Allows users to configure certificate template data and generate certificates.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import subprocess
from pathlib import Path


class CertificateConfigGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Thai Certificate Generator - Configuration")
        self.root.geometry("600x700")
        self.root.resizable(False, False)

        self.config_file = "certificate_config.json"
        self.config = self.load_config()

        self.create_widgets()
        self.load_widgets_from_config()

    def load_config(self):
        """Load configuration from JSON file, or return defaults."""
        defaults = {
            "school_name": "โรงเรียนบ้านโพนแท่น",
            "province_name": "ร้อยเอ็ด",
            "office_name": "สพป. ร้อยเอ็ด เขต ๒",
            "graduated_date": "๓๑",
            "graduated_month": "มีนาคม",
            "graduated_year": "๒๕๖๘",
            "head_teacher_name": "(นางสาวอำพร วรวงษ์)",
            "position_name": "รักษาการในตำแหน่งผู้อำนวยการโรงเรียนบ้านโพนแท่น",
            "excel_file": "name_list.xlsx",
            "template_pdf": "examples.pdf",
            "output_pdf": "output_with_thai_text.pdf",
            "font_file": "DSN-LaiThai.ttf",
            "font_size": "20",
            "remove_background": "True"
        }

        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return {**defaults, **json.load(f)}
            except Exception as e:
                print(f"Error loading config: {e}")

        return defaults

    def save_config_to_file(self):
        """Save current configuration to JSON file."""
        self.config["school_name"] = self.school_name_var.get()
        self.config["province_name"] = self.province_name_var.get()
        self.config["office_name"] = self.office_name_var.get()
        self.config["graduated_date"] = self.graduated_date_var.get()
        self.config["graduated_month"] = self.graduated_month_var.get()
        self.config["graduated_year"] = self.graduated_year_var.get()
        self.config["head_teacher_name"] = self.head_teacher_name_var.get()
        self.config["position_name"] = self.position_name_var.get()
        self.config["excel_file"] = self.excel_file_var.get()
        self.config["template_pdf"] = self.template_pdf_var.get()
        self.config["output_pdf"] = self.output_pdf_var.get()
        self.config["font_file"] = self.font_file_var.get()
        self.config["font_size"] = self.font_size_var.get()
        self.config["remove_background"] = self.remove_background_var.get()

        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration:\n{e}")
            return False

    def load_widgets_from_config(self):
        """Load widget values from configuration."""
        self.school_name_var.set(self.config.get("school_name", ""))
        self.province_name_var.set(self.config.get("province_name", ""))
        self.office_name_var.set(self.config.get("office_name", ""))
        self.graduated_date_var.set(self.config.get("graduated_date", ""))
        self.graduated_month_var.set(self.config.get("graduated_month", ""))
        self.graduated_year_var.set(self.config.get("graduated_year", ""))
        self.head_teacher_name_var.set(self.config.get("head_teacher_name", ""))
        self.position_name_var.set(self.config.get("position_name", ""))
        self.excel_file_var.set(self.config.get("excel_file", ""))
        self.template_pdf_var.set(self.config.get("template_pdf", ""))
        self.output_pdf_var.set(self.config.get("output_pdf", ""))
        self.font_file_var.set(self.config.get("font_file", ""))
        self.font_size_var.set(self.config.get("font_size", "20"))
        self.remove_background_var.set(self.config.get("remove_background", "True"))

    def create_widgets(self):
        """Create all GUI widgets."""

        # Create notebook for tabbed interface
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 1: Certificate Information
        cert_frame = ttk.Frame(notebook, padding="15")
        notebook.add(cert_frame, text="Certificate Information")

        self.create_certificate_fields(cert_frame)

        # Tab 2: File Settings
        file_frame = ttk.Frame(notebook, padding="15")
        notebook.add(file_frame, text="File Settings")

        self.create_file_fields(file_frame)

        # Tab 3: Advanced Settings
        adv_frame = ttk.Frame(notebook, padding="15")
        notebook.add(adv_frame, text="Advanced")

        self.create_advanced_fields(adv_frame)

        # Button frame at bottom
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill='x')

        ttk.Button(button_frame, text="Save Configuration", command=self.save_config).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Generate Certificates", command=self.generate_certificates).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_defaults).pack(side='right', padx=5)

    def create_certificate_fields(self, parent):
        """Create certificate information input fields."""
        self.school_name_var = tk.StringVar()
        self.province_name_var = tk.StringVar()
        self.office_name_var = tk.StringVar()
        self.graduated_date_var = tk.StringVar()
        self.graduated_month_var = tk.StringVar()
        self.graduated_year_var = tk.StringVar()
        self.head_teacher_name_var = tk.StringVar()
        self.position_name_var = tk.StringVar()

        row = 0
        ttk.Label(parent, text="Certificate Text Configuration", font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, pady=(0, 15), sticky='w')

        self.create_input_row(parent, "School Name:", self.school_name_var, row=1)
        self.create_input_row(parent, "Province Name:", self.province_name_var, row=2)
        self.create_input_row(parent, "Office Name:", self.office_name_var, row=3)

        row = 4
        ttk.Label(parent, text="Graduation Date", font=('Arial', 10, 'bold')).grid(row=row, column=0, columnspan=2, pady=(15, 5), sticky='w')
        self.create_input_row(parent, "Date (Thai numerals):", self.graduated_date_var, row=5, width=10)
        self.create_input_row(parent, "Month (Thai text):", self.graduated_month_var, row=6, width=15)
        self.create_input_row(parent, "Year (Thai numerals):", self.graduated_year_var, row=7, width=10)

        row = 8
        ttk.Label(parent, text="Signer Information", font=('Arial', 10, 'bold')).grid(row=row, column=0, columnspan=2, pady=(15, 5), sticky='w')
        self.create_input_row(parent, "Head Teacher Name:", self.head_teacher_name_var, row=9)
        self.create_input_row(parent, "Position:", self.position_name_var, row=10)

    def create_file_fields(self, parent):
        """Create file path input fields."""
        self.excel_file_var = tk.StringVar()
        self.template_pdf_var = tk.StringVar()
        self.output_pdf_var = tk.StringVar()
        self.font_file_var = tk.StringVar()

        ttk.Label(parent, text="File Paths", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky='w')

        self.create_file_input_row(parent, "Excel Data File:", self.excel_file_var, 1,
                                   [("Excel files", "*.xlsx *.xls")])
        self.create_file_input_row(parent, "PDF Template:", self.template_pdf_var, 2,
                                   [("PDF files", "*.pdf")])
        self.create_file_input_row(parent, "Output PDF:", self.output_pdf_var, 3,
                                   [("PDF files", "*.pdf")])
        self.create_file_input_row(parent, "Font File (.ttf):", self.font_file_var, 4,
                                   [("TrueType fonts", "*.ttf")])

    def create_advanced_fields(self, parent):
        """Create advanced settings fields."""
        self.font_size_var = tk.StringVar(value="20")
        self.remove_background_var = tk.StringVar(value="True")

        ttk.Label(parent, text="Advanced Settings", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky='w')

        self.create_input_row(parent, "Font Size:", self.font_size_var, row=1, width=10)

        ttk.Label(parent, text="Remove Background:").grid(row=2, column=0, sticky='w', pady=5)
        ttk.Combobox(parent, textvariable=self.remove_background_var,
                    values=["True", "False"], state="readonly", width=10).grid(row=2, column=1, sticky='w', pady=5)

        ttk.Label(parent, text="(True = use overlay only, False = merge with template PDF)",
                 font=('Arial', 8), foreground='gray').grid(row=3, column=0, columnspan=2, sticky='w')

    def create_input_row(self, parent, label, variable, row, width=40):
        """Helper to create a labeled input row."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=5, padx=(0, 10))
        entry = ttk.Entry(parent, textvariable=variable, width=width)
        entry.grid(row=row, column=1, sticky='w', pady=5)

    def create_file_input_row(self, parent, label, variable, row, filetypes):
        """Helper to create a file input row with browse button."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=5, padx=(0, 10))
        entry = ttk.Entry(parent, textvariable=variable, width=35)
        entry.grid(row=row, column=0, columnspan=1, sticky='e', pady=5, padx=(150, 0))
        ttk.Button(parent, text="Browse...", width=10,
                  command=lambda: self.browse_file(variable, filetypes)).grid(row=row, column=1, sticky='w', pady=5, padx=(5, 0))

    def browse_file(self, variable, filetypes):
        """Open file browser dialog."""
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            variable.set(filename)

    def save_config(self):
        """Save configuration to file."""
        if self.save_config_to_file():
            messagebox.showinfo("Success", "Configuration saved successfully!")

    def reset_defaults(self):
        """Reset all fields to default values."""
        if messagebox.askyesno("Confirm Reset", "Reset all fields to default values?"):
            # Reset config file by deleting it and reloading defaults
            if os.path.exists(self.config_file):
                os.remove(self.config_file)
            self.config = self.load_config()
            self.load_widgets_from_config()
            messagebox.showinfo("Reset Complete", "Configuration reset to defaults.")

    def generate_certificates(self):
        """Generate certificates using current configuration."""
        # Save config first
        if not self.save_config_to_file():
            return

        # Check if required files exist
        excel_file = self.excel_file_var.get()
        template_pdf = self.template_pdf_var.get()
        font_file = self.font_file_var.get()

        missing_files = []
        if not os.path.exists(excel_file):
            missing_files.append(f"Excel file: {excel_file}")
        if not os.path.exists(template_pdf):
            missing_files.append(f"Template PDF: {template_pdf}")
        if not os.path.exists(font_file):
            missing_files.append(f"Font file: {font_file}")

        if missing_files:
            messagebox.showerror("Missing Files",
                               f"The following required files are missing:\n\n" +
                               "\n".join(missing_files) +
                               "\n\nPlease check the file paths and try again.")
            return

        # Run the certificate generation script
        try:
            result = subprocess.run(
                ["python", "main.py"],
                capture_output=True,
                text=True,
                cwd=os.getcwd()
            )

            if result.returncode == 0:
                messagebox.showinfo("Success",
                                  f"Certificates generated successfully!\n\nOutput: {self.output_pdf_var.get()}")
            else:
                messagebox.showerror("Error",
                                   f"Failed to generate certificates:\n\n{result.stderr}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to run generation script:\n\n{e}")


def main():
    root = tk.Tk()
    app = CertificateConfigGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import PyPDF2
import webbrowser  # Import webbrowser to open email client
from datetime import datetime

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Text Extractor")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")

        self.input_pdf_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.data = []

        self.create_widgets()

    def create_widgets(self):
        # Frame untuk input PDF
        frame_input = tk.Frame(self.root, bg="#f0f0f0")
        frame_input.pack(pady=10)

        tk.Label(frame_input, text="Input PDF File:", bg="#f0f0f0").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(frame_input, textvariable=self.input_pdf_var, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(frame_input, text="Browse", command=self.select_input_pdf).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(frame_input, text="Output Directory:", bg="#f0f0f0").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(frame_input, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(frame_input, text="Browse", command=self.select_output_dir).grid(row=1, column=2, padx=10, pady=10)

        # Tombol untuk memulai ekstraksi
        tk.Button(self.root, text="Start Extraction", command=self.start_extraction, bg="#4CAF50", fg="white", font=("Arial", 12)).pack(pady=20)

        # Tombol untuk mengekspor ke Excel
        self.export_button = tk.Button(self.root, text="Export to Excel", command=self.export_to_excel, state=tk.DISABLED)
        self.export_button.pack(pady=10)

        # Treeview untuk menampilkan hasil ekstraksi
        self.tree = ttk.Treeview(self.root, columns=("Full Name", "Email", "Company/University", "Date", "Your Score"), show='headings')
        self.tree.heading("Full Name", text="Full Name")
        self.tree.heading("Email", text="Email")
        self.tree.heading("Company/University", text="Company/University")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Your Score", text="Your Score")
        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)

        # Label untuk informasi aplikasi dengan hyperlink
        self.info_label = tk.Label(self.root, text="Dibuat oleh Ade Ramdani", fg="blue", cursor="hand2", bg="#f0f0f0")
        self.info_label.pack(pady=10)
        self.info_label.bind("<Button-1>", self.open_email)  # Bind click event to open email

    def open_email(self, event):
        webbrowser.open("mailto:mr.aderamdani@gmail.com")  # Open email client

    def select_input_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        self.input_pdf_var.set(file_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        self.output_dir_var.set(dir_path)

    def start_extraction(self):
        input_pdf = self.input_pdf_var.get()
        output_dir = self.output_dir_var.get()

        if not input_pdf or not output_dir:
            messagebox.showerror("Error", "Silakan pilih file PDF dan direktori output.")
            return

        self.extract_text(input_pdf)
        self.export_button.config(state=tk.NORMAL)

    def extract_text(self, input_pdf):
        try:
            with open(input_pdf, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"

                self.parse_text(text)
        except PyPDF2.errors.PdfReadError:
            messagebox.showerror("Error", "Gagal membaca file PDF. Pastikan file PDF valid.")
        except FileNotFoundError:
            messagebox.showerror("Error", "File PDF tidak ditemukan.")


    def parse_text(self, text):
        # Regex pattern to extract required fields (without Time Spent and Result)
        pattern = r"Full Name\s*([^\n]*)\s*Email\s*([^\n]*)\s*Company/University\s*([^\n]*)\s*Date/Time:\s*([^\n]*)\s*Answered:\s*[^\n]*\s*Your Score:\s*([^\n]*)\s*Passing Score:\s*[^\n]*\s*Time Spent:\s*[^\n]*\s*Result\s*[^\n]*"

        matches = re.findall(pattern, text, re.DOTALL)

        for match in matches:
            extracted_data = {
                "Full Name": match[0].strip(),
                "Email": match[1].strip(),
                "Company/University": match[2].strip(),
                "Date": self.format_date(match[3].strip()),  # Format the date
                "Your Score": self.extract_percentage(match[4].strip()),  # Extract percentage from Your Score
            }
            self.data.append(extracted_data)

            # Insert data into the treeview (without Time Spent and Result)
            self.tree.insert("", "end", values=(extracted_data["Full Name"], extracted_data["Email"], extracted_data["Company/University"],
                                                extracted_data["Date"], extracted_data["Your Score"]))

        messagebox.showinfo("Success", f"{len(self.data)} records extracted successfully!")

    def extract_percentage(self, score):
        # Extract the percentage value from the score string
        match = re.search(r'(\d+)%', score)
        if match:
            return match.group(1)  # Return the percentage value as a string
        return score  # Return the original score if no percentage is found

    def format_date(self, date_str):
        # Remove time from the date string
        date_only = date_str.split()[0:3]  # Get the first three parts (day, month, year)
        date_str = " ".join(date_only)  # Join them back into a string

        # Define month mappings for conversion
        month_map_id = {
            "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
            "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
            "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
        }

        month_map_en = {
            "January": "01", "February": "02", "March": "03", "April": "04",
            "May": "05", "June": "06", "July": "07", "August": "08",
            "September": "09", "October": "10", "November": "11", "December": "12"
        }

        # Try to convert the date string to a datetime object
        for month in month_map_id.keys():
            if month in date_str:
                month_number = month_map_id[month]
                day_year = date_str.replace(month, month_number).strip()
                try:
                    date_obj = datetime.strptime(day_year, "%d %m %Y")
                    return date_obj.strftime("%d/%m/%Y")  # Return in DD/MM/YYYY format
                except ValueError:
                    continue

        for month in month_map_en.keys():
            if month in date_str:
                month_number = month_map_en[month]
                day_year = date_str.replace(month, month_number).strip()
                try:
                    date_obj = datetime.strptime(day_year, "%d %m %Y")
                    return date_obj.strftime("%d/%m/%Y")  # Return in DD/MM/YYYY format
                except ValueError:
                    continue

        return date_str  # Return original if no conversion is possible

    def export_to_excel(self):
        df = pd.DataFrame(self.data)
        df = df[["Full Name", "Email", "Company/University", "Date", "Your Score"]]  # Columns for Excel export
        output_dir = self.output_dir_var.get()
        file_path = os.path.join(output_dir, "extract-text.xlsx")  # Set default filename

        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Data exported to {file_path} successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to Excel: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

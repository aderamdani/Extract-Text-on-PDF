import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import PyPDF2
import webbrowser  # Import webbrowser to open email client

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
        self.tree = ttk.Treeview(self.root, columns=("Name", "Email", "Company", "Department", "Job Title", "Date", "Score"), show='headings')
        self.tree.heading("Name", text="Name")
        self.tree.heading("Email", text="Email")
        self.tree.heading("Company", text="Company")
        self.tree.heading("Department", text="Department")
        self.tree.heading("Job Title", text="Job Title")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Score", text="Score")
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
        with open(input_pdf, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"

            self.parse_text(text)

    def parse_text(self, text):
        # Regex pattern to extract required fields
        pattern = r"Name\s*([^\n]*)\s*Email\s*([^\n]*)\s*Company\s*([^\n]*)\s*Department\s*([^\n]*)\s*Job Title\s*([^\n]*)\s*Date/Time:\s*([^\n]*)\s*Answered:\s*([^\n]*)\s*Your Score:\s*([^\n]*)\s*Passing Score:\s*([^\n]*)\s*Time Spent:\s*([^\n]*)\s*Result\s*([^\n]*)"

        matches = re.findall(pattern, text, re.DOTALL)  # Use DOTALL to match across multiple lines

        for match in matches:
            # Extract and transform data
            date_time = match[5].strip()
            date_only = date_time.split(' ')[0] + ' ' + date_time.split(' ')[1] + ' ' + date_time.split(' ')[2]  # Get only the date part

            score_raw = match[7].strip()  # Get the raw score string
            score_value = score_raw.split('(')[-1].split('%')[0].strip()  # Extract the percentage value

            # Convert score to integer and remove decimal part
            if score_value.replace('.', '', 1).isdigit():  # Check if it's a valid number
                score_value = int(float(score_value))  # Convert to float first, then to int to remove decimal
            else:
                score_value = 0  # Default to 0 if not found

            extracted_data = {
                "Name": match[0].strip(),
                "Email": match[1].strip(),
                "Company": match[2].strip(),
                "Department": match[3].strip(),
                "Job Title": match[4].strip(),
                "Date": date_only,  # Use only the date
                "Score": score_value  # Store the extracted score
            }
            self.data.append(extracted_data)

            # Insert data into the treeview
            self.tree.insert("", "end", values=(extracted_data["Name"], extracted_data["Email"], extracted_data["Company"],
                                                extracted_data["Department"], extracted_data["Job Title"],
                                                extracted_data["Date"], extracted_data["Score"]))

        messagebox.showinfo("Success", f"{len(self.data)} records extracted successfully!")

    def export_to_excel(self):
        # Create DataFrame with the desired column order
        df = pd.DataFrame(self.data)
        df = df[["Name", "Email", "Company", "Department", "Job Title", "Date", "Score"]]  # Set the desired column order
        output_dir = self.output_dir_var.get()
        file_path = os.path.join(output_dir, "extract-text.xlsx")  # Set default filename

        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Data exported to {file_path} successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

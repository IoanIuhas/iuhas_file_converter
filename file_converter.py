import tkinter as tk
from tkinter import filedialog
import pandas as pd
from PyPDF2 import PdfReader
from fpdf import FPDF


class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")

        self.input_file_path = ""
        self.output_file_path = ""
        self.output_format = ""
        self.data_type = ""
        self.file_types = [
            ("All", "*"),
            ("Text files", "*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx"),
            ("PDF files", "*.pdf")
        ]

        # Create GUI elements
        self.input_label = tk.Label(root, text="Select a file to convert:")
        self.input_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.output_label = tk.Label(root, text="Select the output format:")
        self.output_format_var = tk.StringVar()
        self.output_format_var.set("txt")  # Default format
        self.output_format_menu = tk.OptionMenu(root, self.output_format_var, "txt", "csv", "xlsx", "pdf")
        self.convert_button = tk.Button(root, text="Convert", command=self.convert_file)
        self.result_label = tk.Label(root, text="")

        # Pack GUI elements
        self.input_label.pack()
        self.input_button.pack()
        self.output_label.pack()
        self.output_format_menu.pack()
        self.convert_button.pack()
        self.result_label.pack()

    def browse_file(self):
        self.input_file_path = filedialog.askopenfilename(filetypes=self.file_types)

    def convert_file(self):
        if not self.input_file_path:
            self.result_label.config(text="Select a file to convert.")
            return

        self.output_format = self.output_format_var.get()

        if self.input_file_path.endswith(f".{self.output_format}"):
            self.result_label.config(text=f"Conversion skipped. Input and output formats are identical")
            return

        self.output_file_path = filedialog.asksaveasfilename(
            defaultextension="." + self.output_format,
            filetypes=[(f"{self.output_format.upper()} files", f"*.{self.output_format}")]
        )

        if not self.output_file_path:
            return

        data = self.read_input_file()
        self.write_to_output_file(data)
        self.result_label.config(text=f"Conversion successful. Saved to: {self.output_file_path}")

    def read_input_file(self):
        if self.input_file_path.endswith(".txt"):
            self.data_type = "text"
            with open(self.input_file_path, 'r') as file:
                data = file.read()
            return data
        elif self.input_file_path.endswith(".csv"):
            self.data_type = "dataframe"
            return pd.read_csv(self.input_file_path, sep=";")
        elif self.input_file_path.endswith(".xlsx"):
            self.data_type = "dataframe"
            return pd.read_excel(self.input_file_path)
        elif self.input_file_path.endswith(".pdf"):
            self.data_type = "text"
            return self.extract_text_from_pdf(self.input_file_path)

    def write_to_output_file(self, data):
        if self.output_format.endswith("txt"):
            if self.data_type == "text":
                with open(self.output_file_path, 'w', encoding='utf-8') as file:
                    file.write(data)
            else:
                data.to_csv(self.output_file_path, index=False)
        elif self.output_format.endswith("csv"):
            if self.data_type == "text":
                lines = data.split(sep="\n")
                data = pd.DataFrame(data=lines, columns=['Text'])
            data.to_csv(self.output_file_path, index=False)
        elif self.output_format.endswith("xlsx"):
            if self.data_type == "text":
                lines = data.split(sep="\n")
                data = pd.DataFrame(data=lines, columns=['Text'])
            data.to_excel(self.output_file_path, index=False)
        elif self.output_format.endswith("pdf"):
            pdf = FPDF()
            pdf.add_page()
            if self.data_type == "text":
                pdf.set_font("Times")
                pdf.write(text=data)
            else:
                pdf.write_html(data.to_html(index=False))
            pdf.output(self.output_file_path)

    def extract_text_from_pdf(self, pdf_path):
        text = ""
        pdf = PdfReader(open(pdf_path, "rb"))
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text += page.extract_text()
        return text

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()

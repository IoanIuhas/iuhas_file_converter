import tkinter as tk
from tkinter import filedialog, ttk
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
        self.output_file_name = ""
        self.data = ""
        self.data_type = ""
        self.file_types = [
            ("All", "*"),
            ("Text files", "*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx"),
            ("PDF files", "*.pdf")
        ]

    # Load TTK theme

        style = ttk.Style(root)
        root.tk.call('lappend', 'auto_path', 'awthemes')
        root.tk.call('package', 'require', 'awdark')
        style.theme_use('awdark')

    # Create GUI elements

        self.frame = ttk.Frame(root, width = 300, height = 330, padding = 20)

        self.input_label_frame = ttk.LabelFrame(self.frame, labelanchor = "n", padding = (0, 5, 0, 5), width = 200, height = 100)
        self.input_label = ttk.Label(self.input_label_frame, text = "Select a file to convert", padding = (0, 0, 0, 5), justify = "center")
        self.input_button = ttk.Button(self.input_label_frame, text = "Browse", command = self.browse_file)

        self.output_label_frame = ttk.LabelFrame(self.frame, labelanchor = "n", padding = (0, 5, 0, 5), width = 200, height = 100)
        self.output_label = ttk.Label(self.output_label_frame, text = "Select the output format", padding = (0, 0, 0, 5), justify = "center")
        self.output_format_var = tk.StringVar()
        self.output_format_menu = ttk.OptionMenu(self.output_label_frame, self.output_format_var,"txt", "txt", "csv", "xlsx", "pdf")
        # self.output_format_var.set("csv")  # Default format

        self.separator_label = ttk.Label(self.frame)
        self.convert_button = ttk.Button(self.frame, text = "Convert", command = self.convert_file)
        self.result_label = ttk.Label(self.frame, text = "", justify = "center", padding = (0, 10, 0, 0))

    # Pack GUI elements

        self.frame.pack()
        self.frame.pack_propagate(0)

        self.input_label_frame.pack()
        self.input_label_frame.pack_propagate(0)
        self.input_label.pack()
        self.input_button.pack()

        self.output_label_frame.pack()
        self.output_label_frame.pack_propagate(0)
        self.output_label.pack()
        self.output_format_menu.pack()

        self.separator_label.pack()
        self.convert_button.pack()
        self.result_label.pack()

    def browse_file(self):
        """
        Set input file path
        """
        self.input_file_path = filedialog.askopenfilename(filetypes = self.file_types)  # Set input file path
        self.result_label.config(text = "")                                             # Reset result label
        self.data = self.read_input_file()                                              # Load file contents into data

    def read_input_file(self):
        """
        Returns the input file contents using different methods based on file type. Also sets the resulting data_type
        """
        if self.input_file_path.endswith(".txt"):
            self.data_type = "text"
            with open(self.input_file_path, 'r') as file:
                data = file.read()
            return data
        elif self.input_file_path.endswith(".csv"):
            self.data_type = "dataframe"
            return pd.read_csv(self.input_file_path)            # sep=";" argument can be used for custom csv delimiters
        elif self.input_file_path.endswith(".xlsx"):
            self.data_type = "dataframe"
            return pd.read_excel(self.input_file_path)
        elif self.input_file_path.endswith(".pdf"):
            self.data_type = "text"
            return self.extract_text_from_pdf(self.input_file_path)

    def extract_text_from_pdf(self, pdf_path):
        text = ""
        pdf = PdfReader(open(pdf_path, "rb"))
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text += page.extract_text()
        return text

    def convert_file(self):
        """
        Converts the input file content to the selected output format
        """
        # Perform file checks
        if not self.input_file_path:
            self.result_label.config(text = "Select a file to convert.")
            return

        self.output_format = self.output_format_var.get()

        if self.input_file_path.endswith(f".{self.output_format}"):
            self.result_label.config(text = f"Conversion skipped!\nInput and output formats are identical.")
            return

        # Set output file path
        self.output_file_path = filedialog.asksaveasfilename(
            defaultextension = "." + self.output_format,
            filetypes = [(f"{self.output_format.upper()} files", f"*.{self.output_format}")]
        )

        if not self.output_file_path:
            return

        self.output_file_name = self.output_file_path[self.output_file_path.rindex('/')+1:]

        # Write resulting file to selected path
        self.write_to_output_file(self.data)
        # Display result message
        self.result_label.config(text = f"Conversion successful\nFile {self.output_file_name} saved!")

    def write_to_output_file(self, data):
        """
        Converts string or dataframe input to selected output format and saves it to the output file path
        :param data:
        """
        # For each output format check the input data_type and use the appropriate methods

        # Txt output
        if self.output_format.endswith("txt"):
            if self.data_type == "text":
                with open(self.output_file_path, 'w', encoding='utf-8') as file:
                    file.write(data)
            else:
                data.to_csv(self.output_file_path, index=False)

        # Csv output, when input data_type is text, loads it into a new DataFrame first
        elif self.output_format.endswith("csv"):
            if self.data_type == "text":
                lines = data.split(sep="\n")
                data = pd.DataFrame(data = lines, columns = [''])
            data.to_csv(self.output_file_path, index=False)

        # Xlsx output, when input data_type is text, loads it into a new DataFrame first
        elif self.output_format.endswith("xlsx"):
            if self.data_type == "text":
                lines = data.split(sep = "\n")
                data = pd.DataFrame(data = lines, columns = [''])
            data.to_excel(self.output_file_path, index = False)

        # Pdf output, when input data_type is dataframe, converts it to html first
        elif self.output_format.endswith("pdf"):
            pdf = FPDF()
            pdf.add_page()
            if self.data_type == "text":
                pdf.set_font("Times")
                pdf.write(text = data)
            else:
                pdf.write_html(data.to_html(index = False))
            pdf.output(self.output_file_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()

from docx import Document
from fpdf import FPDF
import pandas as pd
from docx.shared import Inches

class StatPrint:

    def __init__(self, filename="report", doc_type="word", title="Report"):
        # Initialize with default filename, doc_type, and title
        self.filename = filename
        self.doc_type = doc_type
        self.title = title  # Set the title
        self.content = []

    def add_heading(self, heading): # add a heading to each section, table, or graph
        self.content.append(('heading', heading))

    def add_table(self, data): # can add pandas dataframe or series object
        if isinstance(data, pd.Series):
            df = data.to_frame(name=data.name if data.name else "Value")
            df.insert(0, "Index", df.index)  # Add index as a column
        else:
            df = data  # Assume it's already a DataFrame
        
        # Append the DataFrame to content as a table
        self.content.append(('table', df))

    def add_graph(self, graph, filename="graph.png"):
        graph.savefig(filename, format='png')  # Use the figure object to save the image
        # Add the image to the content as a graph
        self.content.append(('graph', filename))

    def generate_report(self):
        if self.doc_type == 'word':
            self.generate_word_report()
        elif self.doc_type == 'pdf':
            self.generate_pdf_report()

    def generate_word_report(self):
        # Create a new Word document
        doc = Document()
        doc.add_heading(self.title, 0)  # Use custom title instead of default filename

        # Add content (headings, tables, graphs)
        for content_type, content in self.content:
            if content_type == 'heading':
                doc.add_heading(content, level=1)  # Add the heading to the document
            elif content_type == 'table':
                # Add table with borders
                table = doc.add_table(rows=1, cols=len(content.columns))
                table.style = 'Table Grid'  # Set default table style with borders

                # Add headers
                hdr_cells = table.rows[0].cells
                for i, column_name in enumerate(content.columns):
                    hdr_cells[i].text = column_name

                # Add rows
                for row in content.itertuples(index=False, name=None):
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        row_cells[i].text = str(value)

            elif content_type == 'graph':
                # Add graph image to the Word document
                doc.add_paragraph()  # Add a blank line before the image
                doc.add_picture(content, width=Inches(6))  # Adjust width if necessary

        # Save the document as .docx
        doc.save(f"{self.filename}.docx")
        print(f"Report saved as {self.filename}.docx")

    def generate_pdf_report(self):
        # Create a new PDF document
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Set title font and add the title
        pdf.set_font("Arial", size=16, style='B')
        pdf.cell(200, 10, txt=self.title, ln=True, align='C')
        
        # Add content (headings, tables, graphs)
        pdf.set_font("Arial", size=12)
        for content_type, content in self.content:
            if content_type == 'heading':
                pdf.ln(10)  # Add line break
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(200, 10, txt=content, ln=True)
            elif content_type == 'table':
                # Add a table to PDF
                pdf.set_font("Arial", size=10)
                for column in content.columns:
                    pdf.cell(30, 10, column, border=1)
                pdf.ln()
                for row in content.itertuples(index=False, name=None):
                    for value in row:
                        pdf.cell(30, 10, str(value), border=1)
                    pdf.ln()
            elif content_type == 'graph':
                # Add graph image to the PDF document
                pdf.ln(10)  # Add line break
                pdf.image(content, x=None, y=None, w=150)  # Adjust width if necessary

        # Save the PDF document
        pdf.output(f"{self.filename}.pdf")
        print(f"Report saved as {self.filename}.pdf")

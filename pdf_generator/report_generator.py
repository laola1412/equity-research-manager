from weasyprint import HTML

class ReportGenerator:
    def __init__(self, html_template):
        self.html_template = html_template

    def generate_report(self):
        html = HTML(string=self.html_template)
        html.write_pdf(target='report.pdf')
        output_path = 'report.pdf'
        
        # Save the generated PDF to the output path
        with open(output_path, 'wb') as f:
            f.write(pdf.write_pdf())
# from jinja2 import Environment, FileSystemLoader
# import pdfkit

# # Data to be included in the template
# data = {
#     'title': 'My PDF Document',
#     'heading': 'Hello, PDF!',
#     'content': 'This is some content for the PDF file.',
# }

# # Path to the directory containing the HTML template
# template_dir = 'c:\\Users\\meha\\OneDrive\\Desktop\\jinja2'
# file_loader = FileSystemLoader(template_dir)
# env = Environment(loader=file_loader)

# # Load the Jinja2 template
# template = env.get_template('sample.html')

# # Render the template with data
# rendered_html = template.render(data)

# # Configuration for pdfkit (set the path to wkhtmltopdf if not in PATH)
# config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

# # Convert HTML to PDF
# output_pdf = 'sample.pdf'
# pdfkit.from_string(rendered_html, output_pdf, configuration=config)

# print(f'PDF successfully generated: {output_pdf}'
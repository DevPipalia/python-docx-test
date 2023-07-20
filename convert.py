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




import os
from jinja2 import Environment, FileSystemLoader
from docx import Document
from docx.shared import Inches,Pt,Cm
from html.parser import HTMLParser

# Define your mock CV data here (or import it from another file)
mockCVData = [
    {
        "name": "Fiona Wenhan Zhao",
        "email": "wenhanzhao8890@gmail.com",
        "socialMedia": {
            "linkedin": "fiona@linkedin",
            "twitter": "fiona@twitter",
            "github": "fiona@github"
        },
        "professionalExperience": [
            {
                "title": "Founder",
                "company": "UNIQUE BUNNY",
                "location": "Winnipeg, Canada",
                "startDate": "2014-01-01",
                "endDate": "Present",
                "points": [
                    "Founder and GM of the largest chain boutique in Manitoba that specializes in Japanese & Korean beauty and lifestyle products",
                    "Managed 3 brick-n-mortars and online store with $5Mn+ GMV and $1Mn+ annual revenue & $1.2M free cash flow in 2021",
                    "Created an inventory of X+ products ranging from X categories resulting in a YoY revenue growth of X%",
                    "Improved the customer retention rate by X% by supervising 15 store staff and developing customer service training manuals, teaching product features and selling points",
                    "Performed inventory analysis and improved stock-forecasting mechanism by X% by communicating with vendors, couriers, and Canadian Border Services Agency to ensure on-time, complete delivery of products",
                    "Conducted product-mix optimization drives to analyze consumer behavior and accordingly founded X best-selling products",
                    "Collaborated with X+ marketing firms to run online advertising and in-store marketing by allocating a total budget of X$",
                    "Led the digital transformation of the company by designing and launching the official website that has X MAU",
                    "Managed the company’s social media presence across X platforms by actively posting promotions, blogs, and new products; Accumulated 15k+ followers across multiple platforms"
                ]
            },
            {
                "title": "Boarding Advisor",
                "company": "ST. JOHNS - RAVENSCOURT SCHOOL",
                "location": "Winnipeg, Canada",
                "startDate": "2016-01-01",
                "endDate": "2020-12-31",
                "points": [
                    "Designed & executed efficient study programs; Improved student results by X%",
                    "Mentored 30+ international boarding students, providing each student with peer mentorship sessions to help students adjust to the boarding school environment and improve their academic and social performances",
                    "Planned and executed X stimulating programs and activities, connecting students to the Winnipeg community at large and providing students with a deeper understanding of the Canadian culture"
                ]
            },
            {
                "title": "Counter Manager",
                "company": "HUDSON’S BAY COMPANY",
                "location": "Winnipeg, Canada",
                "startDate": "2013-01-01",
                "endDate": "2014-12-31",
                "points": [
                    "Managed the Clarins Paris counter at the Hudson’s Bay Company – Winnipeg flagship, achieving 30% revenue increase",
                    "Awarded as the Top Sales Associate of the Month – Three times",
                    "Created a client & store management SOP that enhanced the customer experience by offering professional consultations to X+ customers; Efforts yielded strong customer satisfaction, earning recognition from Clarins HQ",
                    "Built a clientele of X+ customers by promoting the products on social media platforms"
                ]
            }
        ],
        "skills": [
            "Digital Marketing",
            "Inventory Management",
            "Customer Service",
            "Data Analysis",
            "Social Media Management"
        ],
        "education": [
            {
                "degree": "Bachelor of Arts",
                "major": "Women and Gender Studies",
                "university": "UNIVERSITY OF WINNIPEG",
                "location": "Winnipeg, Canada",
                "gpa":3.0/5.0,
                "startDate":2019,
                "endDate":2022
            }
        ],
        "languages": ["English", "Mandarin"],
        "certifications": [],
        "interests": [
            "Entrepreneurship",
            "Fashion",
            "Blogging"
        ]
    }
]

# Load the Jinja environment and specify the directory containing the template file
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template('sample.html')



# Function to create a Word document from the HTML content
def create_word_document(cv_data,html_content, output_file):
    plain_text = strip_tags(html_content)

    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1.0)  # Adjust the top margin in centimeters
    name_paragraph = document.add_paragraph()
    name_paragraph.add_run(mockCVData[0]['name']).bold = True
    name_paragraph.alignment = 1  # 1 means centered alignment

    contact_info_paragraph = document.add_paragraph()
    contact_info_paragraph.add_run(cv_data['email'])
    social_media = cv_data.get('socialMedia', {})
    if social_media.get('linkedin') or social_media.get('twitter') or social_media.get('github'):
        contact_info_paragraph.add_run(" | ")
        if social_media.get('linkedin'):
            contact_info_paragraph.add_run(social_media['linkedin'])
            if social_media.get('twitter') or social_media.get('github'):
                contact_info_paragraph.add_run(" | ")
        if social_media.get('twitter'):
            contact_info_paragraph.add_run(social_media['twitter'])
            if social_media.get('github'):
                contact_info_paragraph.add_run(" | ")
        if social_media.get('github'):
            contact_info_paragraph.add_run(social_media['github'])
    contact_info_paragraph.alignment = 1








    document.add_paragraph(plain_text, style='BodyText')

    # Save the Word document
    document.save(output_file)

# Rendering the template with the CV data
rendered_cv = template.render(cv_data=mockCVData[0])

# Convert HTML to Word document
output_file_path = 'cv_output13.docx'
create_word_document(mockCVData[0],rendered_cv, output_file_path)

print(f"Word document '{output_file_path}' created successfully.")
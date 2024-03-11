from docx import Document
import re
import textwrap
import os
import tkinter as tk
from tkinter import filedialog

def generate_document(title, link, abstract, count):
    formatted_title = re.sub(' +', ' ', title.strip())
    formatted_title = textwrap.fill(formatted_title, width=60)
    formatted_link = re.sub(' +', ' ', link.strip())
    formatted_link = textwrap.fill(formatted_link, width=60)
    link = f'<a href="{link}">Link</a>'

    document = f'''
    <p style="font-size:2em; line-height:0.5em;"><a href="{formatted_link}" target="_blank">{formatted_title}</a></p>
    <button onclick="toggleAbstract('abstract_{count}')">摘要</button>
    <p id="abstract_{count}" style="display:none; font-size:1em;">{abstract}</p>
    <hr>
    '''

    return document

def parse_word_document(docx_file_path):
    doc = Document(docx_file_path)
    email_content = ""
    for paragraph in doc.paragraphs:
        email_content += paragraph.text + '\n'
    return email_content

def parse_alter_email(email_content):
    papers = []

    email_sections = re.split(r'\\(?!\\)', email_content)
    for section in email_sections:
        if section.startswith('---'):
            continue
        if section.startswith('\n  '):
            modified_section = section.rstrip('\\').strip()
            paper['abstract'] = modified_section
            continue
        lines = section.strip().split('\n')
        paper = {}
        for line in lines:
            if line.startswith('Title:'):
                title = line.split(': ', 1)[1]
                paper['title'] = title
            if line.startswith('( '):
                link = re.findall(r'https?://arxiv\.org/abs/([^,\s]+)', line)
                paper['link'] = 'https://arxiv.org/abs/' + link[0]

        if 'title' in paper:
            papers.append(paper)
        if 'link' in paper:
            index = len(papers) - 1
            papers[index].update(paper)
    return papers

def generate_html_content(papers):
    output = ""
    for i, paper in enumerate(papers):
        title = paper.get("title", "")
        link = paper.get("link", "")
        abstract = paper.get("abstract", "")

        document = generate_document(title, link, abstract, i+1)
        output += document

    html_content = f'''
    <html>
    <head>
    <style>
    p {{
        margin: 0.25em 0;
    }}
    </style>
    <script>
    function toggleAbstract(id) {{
        var abstract = document.getElementById(id);
        if (abstract.style.display === "none") {{
            abstract.style.display = "block";
        }} else {{
            abstract.style.display = "none";
        }}
    }}
    </script>
    </head>
    <body>
    {output}
    </body>
    </html>
    '''

    return html_content

# Read Alter email content from Word document
def select_word_document():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Word Document",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )

    return file_path

def select_html_output():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.asksaveasfilename(
        title="Save HTML Output As",
        defaultextension=".html",
        filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")]
    )

    return file_path

# Read Alter email content from Word document
docx_file_path = select_word_document()

if not docx_file_path:
    print("No file selected. Exiting.")
    exit()

email_content = parse_word_document(docx_file_path)

# Parse Alter email and generate HTML content
papers_data = parse_alter_email(email_content)
html_content = generate_html_content(papers_data)

# Save HTML content to a file
html_file_path = select_html_output()

if not html_file_path:
    print("No HTML file path selected. Exiting.")
    exit()
with open(html_file_path, 'w', encoding='utf-8') as file:
    file.write(html_content)

print(f"Successfully generated HTML file: {html_file_path}")

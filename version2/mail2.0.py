from docx import Document
import re
import textwrap
import os
import tkinter as tk
from tkinter import filedialog
import time

# 生成每篇论文的呈现内容
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

# 从Word文档中解析邮件内容
def parse_word_document(docx_file_path):
    doc = Document(docx_file_path)
    email_content = ""
    for paragraph in doc.paragraphs:
        email_content += paragraph.text + '\n'
    return email_content

# 解析Alter邮件
def parse_alter_email(email_content):
    papers = []

    email_sections = re.split(r'\\(?!\\)', email_content)
    for section in email_sections:
        if section.startswith('---'):
            continue
        if section.startswith('\n  '):
            modified_section = section.rstrip('\\').strip()
            paper['abstract'] = modified_section #解析摘要部分
            continue
        lines = section.strip().split('\n')
        paper = {}
        for line in lines:
            if line.startswith('Title:'):
                title = line.split(': ', 1)[1] 
                paper['title'] = title #解析标题部分
            if line.startswith('( '):
                link = re.findall(r'https?://arxiv\.org/abs/([^,\s]+)', line)
                paper['link'] = 'https://arxiv.org/abs/' + link[0] #解析链接部分

        if 'title' in paper:
            papers.append(paper)
        if 'link' in paper:
            index = len(papers) - 1
            papers[index].update(paper)
    return papers

# 生成HTML内容
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

# 选择输入的Word文档
def select_word_document():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Word Document",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )

    return file_path

# 选择Alert-filter.txt
def select_alert_filter():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Alert-filter.txt",
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        defaultextension=".txt"
    )

    return file_path

# 另存为输出的HTML文件
def select_html_output():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.asksaveasfilename(
        title="Save HTML Output As",
        defaultextension=".html",
        filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")]
    )

    return file_path

# 读取Alert-filter.txt获取排斥词和关注词
def read_alert_filter(alert_filter_path):
    exclude_keywords = []
    include_keywords = []

    if not alert_filter_path:
        return exclude_keywords, include_keywords

    with open(alert_filter_path, 'r', encoding='utf-8') as filter_file:
        for line in filter_file:
            line = line.strip()
            if line.startswith('-'):
                exclude_keywords.append(line[1:])
            elif line.startswith('+'):
                include_keywords.append(line[1:])

    return exclude_keywords, include_keywords

# 根据排斥词和关注词筛选文档
def filter_documents(document_list, exclude_keywords, include_keywords):
    excluded_documents = []
    included_documents = []
    remaining_documents = []

    for document in document_list:
        title = document.get("title", "").strip()

        # Check if title contains any exclude keyword
        if any(keyword.lower() in title.lower() for keyword in exclude_keywords):
            excluded_documents.append(document)
        # Check if title contains any include keyword
        elif any(keyword.lower() in title.lower() for keyword in include_keywords):
            included_documents.append(document)
        else:
            remaining_documents.append(document)

    return excluded_documents, included_documents, remaining_documents

""" 
1.读取邮件,及“筛选词表”
2.执行一期任务,获取文档列表
3.解析“筛选词表”,获取“排斥词”和“关注词”
4.根据“排斥词”,循环筛选排斥文章列表
	（注：以附件 Alert-filter.txt 为例，下同）
以当前“排斥词”检索文章列表，文章题目如含此词(如-Knowledge graph),将该文转移至排斥文章列表（如无，新建）；
	从“当前文档列表中”删除被转移的文章；
	重复此步骤，直至处理完所有“排斥词”。
5.根据“关注词”，循环筛选关注文章
以当前“关注词”检索文章列表，文章题目如含此词(如+Graph Matching),将该文转移至关注文章列表（如无，新建）；
	从“当前文档列表中”删除被转移的文章；
	重复此步骤，直至处理完所有“关注词”。
6.如果用户未设置“筛选词表”,缺省为Alert-filter.txt。如Alert-filter.txt缺失,则跳过本功能。
7.重新组合文档列表
按照“关注文档列表”、“当前文档列表”（剩余部分）、和“排斥文档列表”的顺序重新组合，并生成新的文档列表。
注：各部分之间，可暂插入标题（或分页、或标签）等简单标志，有更好的方式可后续替换。
""" 
# 1.1 读取邮件
docx_file_path = select_word_document()

if not docx_file_path:
    print("No file selected. Exiting.")
    time.sleep(5)
    exit()

# 1.2 读取筛选词表
alert_filter_path = select_alert_filter()

# 1.3（6.1） 缺省为Alert-filter.txt
if not alert_filter_path:
    alert_filter_path = "Alert-filter.txt"

# 1.4（6.2）如Alert-filter.txt缺失,则跳过本功能
if not os.path.exists(alert_filter_path):
    print("Alert-filter.txt not found. Skipping filtering.")
    time.sleep(5)
    exit()

# 2.1 获取文档内容    
email_content = parse_word_document(docx_file_path)

# 2.2 解析邮件内容
papers_data = parse_alter_email(email_content)

# 3. 解析筛选词表并获取排斥词和关注词
exclude_keywords, include_keywords = read_alert_filter(alert_filter_path)

# 4. 根据排斥词筛选文档
# 5. 根据关注词筛选文档
excluded_docs, included_docs, remaining_docs = filter_documents(papers_data, exclude_keywords, include_keywords)

excluded_file_path = "excluded_documents.txt"
included_file_path = "included_documents.txt"
remaining_file_path = "remaining_documents.txt"

with open(excluded_file_path, 'w', encoding='utf-8') as excluded_file:
    for doc in excluded_docs:
        excluded_file.write(f"{doc['title']}\n")

with open(included_file_path, 'w', encoding='utf-8') as included_file:
    for doc in included_docs:
        included_file.write(f"{doc['title']}\n")

with open(remaining_file_path, 'w', encoding='utf-8') as remaining_file:
    for doc in remaining_docs:
        remaining_file.write(f"{doc['title']}\n")

# 7.1 重新组合文档列表并生成新的HTML内容
organized_document_list = included_docs + remaining_docs + excluded_docs
html_content = generate_html_content(organized_document_list)

# 7.2 保存HTML内容
html_file_path = select_html_output()

if not html_file_path:
    print("No HTML file path selected. Exiting.")
    time.sleep(5)
    exit()
with open(html_file_path, 'w', encoding='utf-8') as file:
    file.write(html_content)

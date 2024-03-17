from docx import Document
import re
import textwrap
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# 生成每篇论文的呈现内容
def generate_document(title, link, abstract, count, section):
    if any(char in title for char in ['$']):
        title = f'<script type="text/x-mathjax-config">MathJax.Hub.Config({{ tex2jax: {{ inlineMath: [ ["$", "$"] ] }} }});</script>{title}'

    # 变更标题关键词颜色
    for keyword in exclude_keywords:
        if re.search(r'\b' + re.escape(keyword) + r'\b', title, re.IGNORECASE):
            title = re.sub(re.escape(keyword),f'<span style="font-size: 0.75em; color: green;"><b>{keyword}</b></span>', title,flags=re.IGNORECASE)
    for keyword in include_keywords:
        if re.search(r'\b' + re.escape(keyword) + r'\b', title, re.IGNORECASE):
            title = re.sub(re.escape(keyword),f'<span style="font-size: 1.25em; color: red;"><b>{keyword}</b></span>', title,flags=re.IGNORECASE)
    # 变更摘要关键词颜色
    if abstract:
        for keyword in exclude_keywords:
            if re.search(r'\b' + re.escape(keyword) + r'\b', abstract, re.IGNORECASE):
                abstract = re.sub(re.escape(keyword),f'<span style="text-decoration: underline;"><b>{keyword}</b></span>', abstract,flags=re.IGNORECASE)
        for keyword in include_keywords:
            if re.search(r'\b' + re.escape(keyword) + r'\b', abstract, re.IGNORECASE):
                abstract = re.sub(re.escape(keyword),f'<span style="font-size: 1em; color: red;"><b>{keyword}</b></span>', abstract,flags=re.IGNORECASE)

    formatted_title = re.sub(' +', ' ', title.strip())
    formatted_title = textwrap.fill(formatted_title, width=60)
    formatted_link = re.sub(' +', ' ', link.strip())
    formatted_link = textwrap.fill(formatted_link, width=60)
    link_content = f'<a href="{link}" target="_blank" style="color: blue;">{formatted_link}</a>'
    title_content = f'<span style="color: black;">{formatted_title}</span>'

    if abstract:
        abstract_button = f'<button style="font-size: 0.8em; background-color: orange; color: white; margin-left: 4em;" onclick="toggleAbstract(\'abstract_{section}\')"><b>Abstract</b></button>'
        abstract_content = f'<p id="abstract_{section}" style="display:none; font-size:1em; margin-top: 0.25em; margin-left: 4em;">{abstract}</p>'
    else:
        abstract_button = ""
        abstract_content = f'<p id="abstract_{section}" style="font-size:1em; color: brown; margin-left: 3em;">Abstract content not provided. Please click the link above for more details.</p>'

    document = f'''
    <div class="paper">
    <p style="font-size:1.5em; line-height:1.25em; margin-left: 2em;">[{count}] {title_content}</p>
    <p style="font-size:1em; margin-left: 3em;">arxiv link: {link_content}</p>
    {abstract_button}
    {abstract_content}
    </div>
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

# 解析Arxiv邮件日期
def parse_arxiv_date(email_content):
    email_sections = re.split(r'\n\\\\(?!\\)', email_content)
    middle_sections = [item.split('\n\n') for item in email_sections]
    email_sections = [item for sublist in middle_sections for item in sublist]
    for section in email_sections:
        if section.startswith('---'):
            lines = section.strip().split('\n')
            for line in lines:
                if line.startswith(' received from'):
                    date_pattern = r'\b[A-Z][a-z]+\s+\d{1,2}\s+[A-Z][a-z]+\b'
                    matched_dates = re.findall(date_pattern, line)
                    date = matched_dates[1]
                    return date

# 解析Arxiv邮件
def parse_arxiv_email(email_content):
    papers = []

    email_sections = re.split(r'\n\\\\(?!\\)', email_content)
    middle_sections = [item.split('\n\n') for item in email_sections]
    email_sections = [item for sublist in middle_sections for item in sublist]

    for section in email_sections:
        if section.startswith('---'):
            continue
        if section.startswith('\n  '):
            modified_section = section.rstrip('\\').strip()
            paper['abstract'] = modified_section #解析摘要部分
            continue
        lines = section.strip().split('\n')
        paper = {}
        outflag=0
        for line in lines:
            if line.startswith('Title:'):
                title = line.split(': ', 1)[1]
                paper['title'] = title #解析标题部分 
                title2 = '\0'
                outflag=1
            if outflag==1 and not line.startswith('Title:') and not line.startswith('Author'):
                    title2 += line  # 连接多行标题
                    paper['title'] = title + title2 
            if line.startswith('Author'):
                    break
        for line in lines:
            if line.startswith('( '):
                link = re.findall(r'https?://arxiv\.org/abs/([^,\s]+)', line)
                paper['link'] = 'https://arxiv.org/abs/' + link[0] #解析链接部分

        if 'title' in paper:
            papers.append(paper)
        if 'link' in paper:
            index = len(papers) - 1
            papers[index].update(paper)
    return papers

# 检查标题和摘要中是否包含特殊字符
def check_special_characters(papers_data):
    try:
        charflag=0
        special_characters = ['#']
        for i, paper in enumerate(papers_data):
            title = paper.get("title", "")
            abstract = paper.get("abstract", "")
            if any(char in title for char in special_characters):
                problematic_char = next(char for char in title if char in special_characters)
                tk.messagebox.showinfo(
                    "Notice",
                    f"Special character '{problematic_char}' found in title at index {i + 1}:\n\n{title}\n\nPlease consider removing this line."
                )
                charflag+=1
            elif any(char in abstract for char in special_characters):
                problematic_char = next(char for char in abstract if char in special_characters)
                tk.messagebox.showinfo(
                    "Notice",
                    f"Special character '{problematic_char}' found in abstract at index {i + 1}:\n\n{abstract}\n\nPlease consider removing this part."
                )
                charflag+=1
        if charflag>0:
            exit()
    except Exception as e:
        tk.messagebox.showerror("Error", f"An error occurred: {str(e)}\n\nPlease check the input file and try again.")
        exit()

# 生成HTML内容
def generate_html_content(papers):
    output = ""
    focus_content = ""
    focus_abstract_content = ""
    remaining_content = ""
    excluded_content = ""
    excluded_abstract_content = ""
    focus_count = 1
    focus_abstract_count = 1
    remaining_count = 1
    excluded_count = 1
    excluded_abstract_count = 1
    date = parse_arxiv_date(email_content)

    for i, paper in enumerate(papers):
        title = paper.get("title", "")
        link = paper.get("link", "")
        abstract = paper.get("abstract", "")
        section = i + 1

        if paper in included_docs:
            document = generate_document(title, link, abstract, focus_count, section)
            focus_content += document
            focus_count += 1
        elif paper in included_abstract_docs:
            document = generate_document(title, link, abstract, focus_abstract_count, section)
            focus_abstract_content += document
            focus_abstract_count += 1
        elif paper in remaining_docs:
            document = generate_document(title, link, abstract, remaining_count, section)
            remaining_content += document
            remaining_count += 1
        elif paper in excluded_abstract_docs:
            document = generate_document(title, link, abstract, excluded_abstract_count, section)
            excluded_abstract_content += document
            excluded_abstract_count += 1
        else:
            document = generate_document(title, link, abstract, excluded_count, section)
            excluded_content += document
            excluded_count += 1

    # Add section titles and combine content in the desired order
    if focus_content:
        output += f"<h2 style='color: rgb(139, 0, 18);'>Hot concerns:</h2>\n" + focus_content
    if focus_abstract_content:
        output += f"<h2 style='color: rgb(139, 0, 18);'>Guess you like it:</h2>\n" + focus_abstract_content
    if remaining_content:
        output += f"<h2 style='color: rgb(139, 0, 18);'>Research trends:</h2>\n" + remaining_content
    if excluded_abstract_content:
        output += f"<h2 style='color: rgb(139, 0, 18);'>Read next time:</h2>\n" + excluded_abstract_content
    if excluded_content:
        output += f"<h2 style='color: rgb(139, 0, 18);'>Irrelevant topics:</h2>\n" + excluded_content

    html_content = f'''
    <html>
    <head>
    <style>
    p {{
        margin: 0.25em 0;
    }}
    </style>
    <script type="text/javascript" async
        src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-MML-AM_CHTML">
    </script>
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
    <h1 style="text-align: center; font-size: 1.5em; color: rgb(139, 0, 18);">Arxiv Mail Subscription on {date}</h1>
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
    
    encodings_to_try = ['utf-8', 'latin-1']  # Add more encodings if needed

    for encoding in encodings_to_try:
        try:
            with open(alert_filter_path, 'r', encoding=encoding) as filter_file:
                for line in filter_file:
                    line = line.strip()
                    if line.startswith('-'):
                        exclude_keywords.append(line[1:])
                    elif line.startswith('+'):
                        include_keywords.append(line[1:])
            return exclude_keywords, include_keywords
        except UnicodeDecodeError:
            # 无法解码文件
            tk.messagebox.showerror("Error", f"Unable to decode file using any of the specified encodings: {encodings_to_try}, please check the file and try again.")
            exit()

# 根据排斥词和关注词筛选文档
def filter_documents(document_list, exclude_keywords, include_keywords):
    excluded_documents = []
    included_documents = []
    remaining_documents = []
    excluded_abstract_documents = []
    included_abstract_documents = []

    for document in document_list:
        title = document.get("title", "").strip()

        exclude_match_title = any(re.search(r'\b' + re.escape(keyword.lower()) + r'\b', title.lower()) for keyword in exclude_keywords)
        include_match_title = any(re.search(r'\b' + re.escape(keyword.lower()) + r'\b', title.lower()) for keyword in include_keywords)

        abstract = document.get("abstract", "").strip()

        exclude_match_abstract = any(re.search(r'\b' + re.escape(keyword.lower()) + r'\b', abstract.lower()) for keyword in exclude_keywords)
        include_match_abstract = any(re.search(r'\b' + re.escape(keyword.lower()) + r'\b', abstract.lower()) for keyword in include_keywords)

        if exclude_match_title:
            excluded_documents.append(document)
        elif include_match_title:
            included_documents.append(document)
        elif exclude_match_abstract:
            excluded_abstract_documents.append(document)
        elif include_match_abstract:
            included_abstract_documents.append(document)
        else:
            remaining_documents.append(document)

    return excluded_documents, included_documents, excluded_abstract_documents, included_abstract_documents, remaining_documents

# 1.1 读取邮件
docx_file_path = select_word_document()

if not docx_file_path:
    tk.messagebox.showinfo("Notice", "No file selected. Exiting.")
    exit()

# 1.2 读取筛选词表
alert_filter_path = select_alert_filter()

# 1.3（6.1） 缺省为Alert-filter.txt
if not alert_filter_path:
    alert_filter_path = "Alert-filter.txt"

# 1.4（6.2）如Alert-filter.txt缺失,则跳过本功能
if not os.path.exists(alert_filter_path):
    tk.messagebox.showinfo("Notice", "Alert-filter.txt not found. Skipping filtering.")
    exit()

# 2.1 获取文档内容    
email_content = parse_word_document(docx_file_path)

# 2.2 解析邮件内容
papers_data = parse_arxiv_email(email_content)

# 2.3 检查标题和摘要中是否包含特殊字符
check_special_characters(papers_data)

# 3. 解析筛选词表并获取排斥词和关注词
exclude_keywords, include_keywords = read_alert_filter(alert_filter_path)

# 4. 根据排斥词筛选文档
# 5. 根据关注词筛选文档
excluded_docs, included_docs, excluded_abstract_docs, included_abstract_docs, remaining_docs = filter_documents(papers_data, exclude_keywords, include_keywords)

excluded_file_path = "excluded_documents.txt"
included_file_path = "included_documents.txt"
excluded_abstract_file_path = "excluded_abstract_documents.txt"
included_abstract_file_path = "included_abstract_documents.txt"
remaining_file_path = "remaining_documents.txt"

with open(excluded_file_path, 'w', encoding='utf-8') as excluded_file:
    for doc in excluded_docs:
        excluded_file.write(f"{doc['title']}\n")

with open(excluded_abstract_file_path, 'w', encoding='utf-8') as excluded_abstract_file:
    for doc in excluded_abstract_docs:
        excluded_abstract_file.write(f"{doc['title']}\n")

with open(included_file_path, 'w', encoding='utf-8') as included_file:
    for doc in included_docs:
        included_file.write(f"{doc['title']}\n")

with open(included_abstract_file_path, 'w', encoding='utf-8') as included_abstract_file:
    for doc in included_abstract_docs:
        included_abstract_file.write(f"{doc['title']}\n")

with open(remaining_file_path, 'w', encoding='utf-8') as remaining_file:
    for doc in remaining_docs:
        remaining_file.write(f"{doc['title']}\n")

# 7.1 重新组合文档列表并生成新的HTML内容
organized_document_list = included_docs + included_abstract_docs + remaining_docs + excluded_abstract_docs + excluded_docs
html_content = generate_html_content(organized_document_list)

# 7.2 保存HTML内容
html_file_path = select_html_output()

# 删除临时文件
os.remove(excluded_file_path)
os.remove(included_file_path)
os.remove(excluded_abstract_file_path)
os.remove(included_abstract_file_path)
os.remove(remaining_file_path)

if not html_file_path:
    tk.messagebox.showinfo("Notice", "No HTML file path selected. Exiting.")
    exit()
with open(html_file_path, 'w', encoding='utf-8') as file:
    file.write(html_content)


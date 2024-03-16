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
    formatted_title = re.sub(' +', ' ', title.strip())
    formatted_title = textwrap.fill(formatted_title, width=60)
    formatted_link = re.sub(' +', ' ', link.strip())
    formatted_link = textwrap.fill(formatted_link, width=60)
    link = f'<a href="{link}" target="_blank">{formatted_title}</a>'

    if abstract:
        abstract_button = f'<button style="font-size: 0.8em; background-color: red; color: white; margin-left: 4em;" onclick="toggleAbstract(\'abstract_{section}\')"><b>Abstract</b></button>'
        abstract_content = f'<p id="abstract_{section}" style="display:none; font-size:1em; margin-top: 0.25em; margin-left: 4em;">{abstract}</p>'
    else:
        abstract_button = ""
        abstract_content = f'<p id="abstract_{section}" style="font-size:1em; margin-left: 4em;">Content not provided. For more details about the abstract, please click the link above.</p>'

    document = f'''
    <div class="paper">
    <p style="font-size:1.5em; line-height:1.25em; margin-left: 2em;">[{count}] {link}</p>
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
    remaining_content = ""
    excluded_content = ""
    focus_count = 1
    remaining_count = 1
    excluded_count = 1

    for i, paper in enumerate(papers):
        title = paper.get("title", "")
        link = paper.get("link", "")
        abstract = paper.get("abstract", "")
        section = i + 1

        if paper in included_docs:
            document = generate_document(title, link, abstract, focus_count, section)
            focus_content += document
            focus_count += 1
        elif paper in remaining_docs:
            document = generate_document(title, link, abstract, remaining_count, section)
            remaining_content += document
            remaining_count += 1
        else:
            document = generate_document(title, link, abstract, excluded_count, section)
            excluded_content += document
            excluded_count += 1

    # Add section titles and combine content in the desired order
    if focus_content:
        output += "<h2>重点关注：</h2>\n" + focus_content
    if remaining_content:
        output += "<h2>猜你喜欢：</h2>\n" + remaining_content
    if excluded_content:
        output += "<h2>下次再看：</h2>\n" + excluded_content

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
    tk.messagebox.showinfo("Notice", "No HTML file path selected. Exiting.")
    exit()
with open(html_file_path, 'w', encoding='utf-8') as file:
    file.write(html_content)

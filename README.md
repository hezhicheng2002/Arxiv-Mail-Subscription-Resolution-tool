# Arxiv-Mail-Subscription-Resolution-tool
A desktop tool to retrieve essays you preferred from your subscription to arxiv.org.
- [arxiv](https://arxiv.org/)

## Requirements

Running **the source code** requires anaconda environment and makes sure you have installed the corresponding packages.

Running **the exe program** needs to keep the program under the same directory with the **"_internal"** folder, where all the necessary environments are already embedded there.

You may also go to the directory and encapsulate the source code into exe by running the line below in your anaconda IDE.

```python
pip install pyinstaller
pyintaller python_name.py
```

### Recommand Input format

The input mail file should be formatted as **"*.docx"**, with a version newer than 2003.

The input filtrating words list should be formatted as "*.txt", with **"UTF-8" or "latin-1"** format.

### Recommend Environment

Windows 10, python>=3.10.0, anaconda suggested.

## Getting Started

### arxiv mail preperation

Go to the bottom of the home page at https://arxiv.org/ and click Subscribe, you will find the official tutorials there. Follow the procedure until you successfully received the confirmation letter.

After receiving the mail, copy them into a word document. Here is a limited example, you may copy the whole content directly.

```txt
------------------------------------------------------------------------------
------------------------------------------------------------------------------
Send any comments regarding submissions directly to submitter.
------------------------------------------------------------------------------
Archives at http://arxiv.org/
To unsubscribe, e-mail To: cs@arXiv.org, Subject: cancel
------------------------------------------------------------------------------
```

### filtrating words list

New a txt document and write in the key words you like or dislike. Here is a formatted example.

```txt
-dislike
-no space after the sign
+like
+CAPITAL IS OK
```

**NOTICE:** In order to avoid unnecessary bugs, please add words in a format that you **exactly** wanted, since what you put in here will be highlighted if mentioned in the result html page.

### Installation

Download the zip file or clone using the web URL.

	git clone https://github.com/He52/Arxiv-Mail-Subscription-Resolution-tool.git

## Usage

Choose the mail word document and then the filtrated words list txt file, and select a path to save the output html file. You may click the html file and see the result shown in the default browser.

The part that you favorate are shown at the very beginning, subsequently are the middle part and the unfavored part.

Click the title will guide you to the arxiv page, and if abstract content is provided, you may click the highlighted button to see more.

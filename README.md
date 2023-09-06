# Docx Splitter 

Using `lxml` to split docx file according to its `TOC`(table of contents) in one travesal.

Some docx files are too large and hard to read, like the original `text.docx` which has almost 1300 pages. This script want to split it into some small docx files according its `TOC`.

This script only work well with the docx file which using `w:hyperlink` as the label of `TOC`. Like `text.docx` in this project. And it has been modified to avoid some sensitive information.  

## Install 
```bash
pip install lxml
```

## Usage
```bash
python DocHandle.py
```

This script will make a foler and split the text.docx into some small docx files to this folder.

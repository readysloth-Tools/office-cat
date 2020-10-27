
import re
import zipfile
import tempfile
import argparse

import xml.dom.minidom as minidom

from pathlib import Path

def get_contents(path_to_doc: Path, docx: bool) -> minidom.Document:
    with zipfile.ZipFile(str(path_to_doc), 'r') as doc:
        content_file_path = 'word/document.xml' if docx else 'content.xml'
        with doc.open(content_file_path) as content:
            xml = content.read()
            dom = minidom.parseString(xml)
            dom.normalize()
            return dom

def print_content(path_to_doc: Path,
                  docx: bool,
                  only_text: bool = False,
                  join_spans: bool = False):
    dom = get_contents(path_to_doc, docx)
    pretty_xml_as_string = dom.toprettyxml()

    if not only_text:
        print(pretty_xml_as_string)
        return

    if docx:
        for text in dom.getElementsByTagName('w:t'):
            for line in text.childNodes:
                print(line.nodeValue)
    else:
        for text in dom.getElementsByTagName('text:p'):
            for span in text.childNodes:
                if span.nodeValue:
                    print(span.nodeValue)
                for line in span.childNodes:
                    if join_spans:
                        print(line.nodeValue, end=' ')
                    else:
                        print(line.nodeValue)


parser = argparse.ArgumentParser(description='Program to "cat" .docx and .odt files')
parser.add_argument('-t', '--text-only', action='store_true', help='print only text')
parser.add_argument('-j', '--join-spans', action='store_true', help='join spans (odt only)')
parser.add_argument('PATH', nargs='+', help='path to document file')

args = parser.parse_args()

for path in args.PATH:
    path_to_doc = Path(path)
    extension = path_to_doc.suffix[1:]
    
    print_content(path_to_doc,
                  extension == 'docx',
                  args.text_only,
                  args.join_spans)

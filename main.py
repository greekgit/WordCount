from docx import Document

from os import chdir

from pathlib import Path


def word_count(filename):
    total = 0
    docfile = open(filename, "rb")
    document = Document(docfile)
    for p in document.paragraphs:
        total += len(p.text.split())
    docfile.close()
    return total


path = Path()
target = 'C:\\Users\paul_\\iCloudDrive\\Personal\\Book\\Fourth Draft'

chdir(target)

grand_total = 0
file_count = 0
wc = 0
file_list = path.glob('*.docx')
for f in file_list:
    if f.name[0] != '~':
        wc = word_count(f)
        print(f"  {f.name} : {wc}")
    grand_total += wc
    file_count += 1
print(f"\n Total for directory {target} is {grand_total} words in {file_count} documents")

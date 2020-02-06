from docx import Document
from docx.oxml.ns import qn
import re

docxName = '杨晓雨1月（6701-6750）的副本.docx'

document = Document(docxName)
length = len(document.paragraphs)
for index in range(length):
    if re.match('\d{4}', document.paragraphs[index].text) is not None:  # 是startpara
        print(document.paragraphs[index].text)
        document.paragraphs[index].text = re.sub('\s(?=[āáǎàōóǒòēéěèīíǐìūúǔùǖǘǚǜüa-z])', '', document.paragraphs[index].text)
        for run in document.paragraphs[index].runs:
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.name = 'Times New Roman'
document.save(docxName)
print('finish')
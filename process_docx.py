from docx import Document
from docx.oxml.ns import qn
import re

class Pinyin:
    def __init__(self, pinyin):
        self.content = pinyin  # real pinyin
        tableSort = str.maketrans('āáǎàōóǒòēéěèīíǐìūúǔùǖǘǚǜü', 'aaaaooooeeeeiiiiuuuuuuuuu')
        self.contentSort = pinyin.translate(tableSort)  # pinyin for sort
        tableForSame = str.maketrans('āáǎàōóǒòēéěèīíǐìūúǔùǖǘǚǜü', 'abcdabcdabcdabcdabcdabcda')
        self.contentForSame = pinyin.translate(tableForSame)  # pinyin for same letter sort

    def __lt__(self, other):
        if self.contentSort < other.contentSort:
            return True
        elif self.contentSort > other.contentSort:
            return False
        else:
            return self.contentForSame < other.contentForSame


def fromParatoPara(parato, parafrom):
    for run in parafrom.runs:
        r = parato.add_run(run.text)
        r.font.name = run.font.name
        if r._element.rPr.rFonts is not None:
            r._element.rPr.rFonts.set(qn('w:eastAsia'), run.font.name)
        r.bold = run.bold


document = Document('李安涛1月（6601-6650）.docx')
document2 = Document('test.docx')
startParas = []  # 每个词的第一段

# 把doc2已有词条加入列表(至少手动添加一个1234zzzz)
for paragraph in document2.paragraphs:
    if re.match('\d{4}', paragraph.text) is not None:
        startParas.append(paragraph)

count = 0
for paragraph in document.paragraphs:
    if re.match('\d{4}', paragraph.text) is not None:  # 是startpara
        ParagraphPY = Pinyin(paragraph.runs[4].text)  # 生成Pinyin并比较,找插入的位置
        indexStartPara = 0
        for startPara in startParas:
            startParaPY = Pinyin(startPara.runs[4].text)
            if ParagraphPY < startParaPY:
                length = len(document2.paragraphs)
                for order in range(length):
                    if document2.paragraphs[order].text == startPara.text:
                        paraBase = document2.paragraphs[order] # 在paraBase前插入新词条
                        break
                break
            indexStartPara += 1
        paraInsert = paraBase.insert_paragraph_before()
        fromParatoPara(paraInsert, paragraph)
        startParas.insert(indexStartPara, paragraph)
    else:  # 不是startpara
        paraInsert = paraBase.insert_paragraph_before()
        fromParatoPara(paraInsert, paragraph)
document2.save('test.docx')
print('finish')
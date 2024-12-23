# This is a Python script used to fill up a template docx file.
import sys

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def set_font(cell, bold=True, center_aligned=True, font_name='FangSong', font_size=12):
    for paragraph in cell.paragraphs:
        if center_aligned:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold


def generate_docx(template_file_name, person_info, paragraphs):
    (student, course, time) = person_info

    outfile = (student.replace('\n', '') + '-'
               + course.replace('\n', '') + '-'
               + (time + "").split(' ')[0] + '.docx')
    print("outfile name: " + outfile)

    docx = Document(template_file_name)
    docx.save(outfile)
    docx = Document(outfile)

    table = docx.tables[0]

    table.cell(0, 1).text = "学生：" + student
    set_font(table.cell(0, 1))

    table.cell(0, 2).text = "科目：" + course
    set_font(table.cell(0, 2))

    table.cell(1, 0).text = "时间：" + time
    set_font(table.cell(1, 0), bold=False, center_aligned=False)

    # remove the old content
    table.cell(2, 0).paragraphs.clear()
    for paragraph in paragraphs:
        table.cell(2, 0).add_paragraph(paragraph)
    set_font(table.cell(2, 0), bold=False, center_aligned=False, font_size=11)

    docx.save(outfile)
    print("updated")


def parse_input(person_info, paragraphs):
    filename = args[1]
    with open(filename) as f:
        paragraph = ''

        while True:
            line = f.readline()
            if not line:
                paragraphs.append(paragraph)
                break
            # process(line)
            if len(person_info) < 3:
                person_info.append(line)
            else:
                if len(line) == 0:
                    if paragraph != '':
                        paragraphs.append(paragraph)
                    paragraph = ''
                else:
                    paragraph += line


if __name__ == '__main__':
    args = sys.argv
    print(args)

    if len(args) < 2:
        print("usage: python main.py input_file")
        exit(1)

    person_info_list = []
    paragraph_list = []
    parse_input(person_info_list, paragraph_list)

    generate_docx("demo.docx", person_info_list, paragraph_list)

    print("done")

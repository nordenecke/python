# -*- coding: utf-8 -*-
"""
Created on Fri Jul 28 19:54:04 2017

@author: norden
"""

import os

from docx import Document

import winreg

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,\
                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
    return winreg.QueryValueEx(key, "Desktop")[0]


def print_all(module_):
  modulelist = dir(module_)
  length = len(modulelist)
  for i in range(0,length,1):
    print(getattr(module_,modulelist[i]))

def main():
    get_docx("card.docx")
    put_docx("card_output.docx")

def get_docx(filename):
    desktop_path=get_desktop()
    if True==os.path.exists(desktop_path+"\\"+filename):
        doc=Document(desktop_path+"\\"+filename)
    tabs = doc.tables
    for t in tabs:
        print("table.alignment=%s" % t.alignment)
        print("table.autofit=%s" % t.autofit)
        print("table.style=%s" % t.style)
        print("table.table_direction=%s" % t.table_direction)

        rows=t.rows
#        print("rows.height=%d"%rows.height)
        for r in rows:
#            print("row.height=%d"%r._tr.Height_val)
#            print(r._tr)
#            print("row.height_rule=%s"%r.height_rule)
            cs=r.cells
            for c in cs:
#                par=c.paragraphs
#                print("cell.paragraphs=%s"%par[0].runs[0].font.size)
                print("cell.width=%d"%c.width)

def put_docx(filename):
    output_column_number=3
    output_row_number=5

    #打开文档
    document = Document()
    #加入不同等级的标题
#    document.add_heading(u'MS WORD写入测试',0)
#    document.add_heading(u'一级标题',1)
#    document.add_heading(u'二级标题',2)
    #添加文本
#    paragraph = document.add_paragraph(u'我们在做文本测试！')
    #设置字号
#    run = paragraph.add_run(u'设置字号、')
#    run.font.size = Pt(24)

    #设置字体
#    run = paragraph.add_run('Set Font,')
#    run.font.name = 'Consolas'

    #设置中文字体
#    run = paragraph.add_run(u'设置中文字体、')
#    run.font.name=u'宋体'
#    r = run._element
#    r.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    #设置斜体
#    run = paragraph.add_run(u'斜体、')
#    run.italic = True

    #设置粗体
#    run = paragraph.add_run(u'粗体').bold = True

    #增加引用
#    document.add_paragraph('Intense quote', style='Intense Quote')

    #增加无序列表
#    document.add_paragraph(
#        u'无序列表元素1', style='List Bullet'
#    )
#    document.add_paragraph(
#        u'无序列表元素2', style='List Bullet'
#    )
    #增加有序列表
#    document.add_paragraph(
#        u'有序列表元素1', style='List Number'
#    )
#    document.add_paragraph(
#        u'有序列表元素2', style='List Number'
#    )
    #增加图像（此处用到图像image.bmp，请自行添加脚本所在目录中）
#    document.add_picture('image.bmp', width=Inches(1.25))

    page_item_number=output_column_number*output_row_number
    total_input_item_number=15
    unused_item_number=total_input_item_number
    current_item=0

    while unused_item_number>0:
        #增加表格
        table1 = document.add_table(rows=output_row_number, cols=output_column_number, style='Table Grid')
        table1.autofit = False
        page_first_item=current_item
        for i in range(min(unused_item_number,page_item_number)):
            first_page_content_item="test"
            hdr_cells = table1.rows[int(i/output_column_number)].cells
            hdr_cells[i%output_column_number].text = first_page_content_item
#            print(hdr_cells[i%output_column_number].width)
#            hdr_cells[i%output_column_number].width=1828800*10
            current_item+=1

        #增加分页
        document.add_page_break()

        #增加表格
        table2 = document.add_table(rows=output_row_number, cols=output_column_number, style='Table Grid')
        table2.autofit = False
        current_item=page_first_item
        for i in range(min(unused_item_number,page_item_number)):
            second_page_content_item="test"
            hdr_cells = table2.rows[int(i/output_column_number)].cells
            hdr_cells[output_column_number-i%output_column_number-1].text = second_page_content_item
            current_item+=1

        if unused_item_number -page_item_number<=0:
            break
        else:
            document.add_page_break()
            unused_item_number-=min(page_item_number,unused_item_number)

    #保存文件
    desktop_path=get_desktop()
    if os.path.exists(desktop_path+"\\"+filename):
        os.remove(desktop_path+"\\"+filename)
    document.save(desktop_path+"\\"+filename)

if __name__=="__main__":
    main()
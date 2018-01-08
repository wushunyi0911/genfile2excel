# -*- coding: utf-8 -*-
import os
import xlwt

def genfilelinks(location='.', prenum=0, splitchar='.'):
    filelinks = []
    for x in [x for x in os.listdir(location) if os.path.isfile(os.path.join(location,x))]:
        filelinks += [''.join(x[prenum::].split(splitchar)[0:-1]),os.path.join(location,x),x.split(splitchar)[-1]]
    for x in [x for x in os.listdir(location) if os.path.isdir(os.path.join(location,x))]:
        filelinks += genfilelinks(os.path.join(location,x),prenum,splitchar)
    return filelinks

def genlinks2excel(filename, location='.', prenum=0, splitchar='.'):
    filelinks = genfilelinks(location, prenum, splitchar)
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet1 = workbook.add_sheet('sheet1',cell_overwrite_ok=True)
    #为样式创建字体
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = False
    #设置样式的字体
    style.font = font
    sheet1.write(0, 0, '链接', style)
    sheet1.write(0, 1, '文件名', style)
    sheet1.write(0, 2, '后缀名', style)
    sheet1.write(0, 3, '地址', style)
    k = 0
    while 3*k+2 < len(filelinks):
        link = 'HYPERLINK("%s";"%s")' % (filelinks[3*k+1], filelinks[3*k])
        sheet1.write(k+1, 0, xlwt.Formula(link),style)
        sheet1.write(k+1, 1, '%s' % filelinks[3*k],style)
        sheet1.write(k+1, 2, '%s' % filelinks[3*k+2],style)
        sheet1.write(k+1, 3, '%s' % filelinks[3*k+1],style)
        k += 1
    workbook.save(filename)

genlinks2excel('test.xls','.',0,'.')

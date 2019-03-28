#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/1/2'
# God helps who help themselves!

import InterfacePptModule as pptHandle
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE #MSO_SHAPE # 添加形状时使用
import InterfaceExcelModule as xlsHandle

if __name__ == '__main__':
    print("This is main function!")

    # 测试生成的额PPT
    '''
    # 新建PPT: 自动生成报告时,使用type6每次添加空白页就可以了
    # type = 0: PPT第一页的类型, ppt标题页;
    # type=1: 带title和文本框的PPT页;
    # type=5: 带title的PPT页;
    # type=6: 完全空白页;
    newPpt = pptHandle.createNewPptFile()
    pptHandle.createNewPage(newPpt, 0)
    pptHandle.createNewPage(newPpt, 1)
    pptHandle.createNewPage(newPpt, 2)
    pptHandle.createNewPage(newPpt, 3)
    pptHandle.createNewPage(newPpt, 4)
    pptHandle.createNewPage(newPpt, 5)
    pptHandle.createNewPage(newPpt, 6)
    pptHandle.createNewPage(newPpt, 7)
    pptHandle.createNewPage(newPpt, 6)
    print("当前PPT页数为:", pptHandle.getSildesNum(newPpt))
    page1 = pptHandle.getOneSilde(newPpt,6)
    print("当前PPT的shapes个数为:", pptHandle.getPageShapesNum(page1))

    # 向page1(第7页)中输出文字:
    newTitle = pptHandle.addTitleBox(page1, "添加标题栏")
    newTextBox = pptHandle.addTextBox(page1,"添加文本框,设置字体颜色!\n这是第二段")
    pptHandle.addNewPara2TextBox(newTextBox,'添加新段落') # 25,True)
    pptHandle.setTextFont(newTextBox,20,True)

    # 向page2(第5页)中添加图片和形状:
    page2 = pptHandle.getOneSilde(newPpt, 5)
    pptHandle.addTitleBox(page2, "添加新的图片和形状")
    pptHandle.addPicture(page2, 'D:/侧碰变形.png')
    pptHandle.addPicture(page2, 'D:/侧碰变形.png',False)
    addShape = pptHandle.addShape(page2, '形状\n标题')
    pptHandle.setShapeFont(addShape,25,True)

    # 向page3中(第8页)添加表格测试:
    page3 = pptHandle.getOneSilde(newPpt,2)
    pptHandle.addTitleBox(page3, "添加新的表格")
    tableData = [ ['111', '222', '哈哈哈'], ['45', '67', '2999'] ]
    table = pptHandle.addTables(page3,tableData)
    pptHandle.setTableOneColumnWidth(table, 0, 2)
    pptHandle.setTableOneColumnWidth(table, 1, 2.5)
    pptHandle.setTableOneColumnWidth(table, 2, 2.5)
    pptHandle.setTableRowHeight(table,0.5)
    pptHandle.setTableAllCellFormat(table)

    # 向page4中(新加页)添加图标测试:
    pptHandle.createNewPage(newPpt, 6)
    page3 = pptHandle.getOneSilde(newPpt, 8)
    pptHandle.addChartBarDemo(page3)
    pptHandle.addChartPieDemo(page3)
    pptHandle.addChartPieDemo2(page3)
    pptHandle.addChartPieDemo3(page3)

    pptHandle.saveAs(newPpt,"D:/python_test.pptx")
    '''

    '''
    # 测试生成EXCEL
    newExcel = xlsHandle.createNewExcelFile()
    print(xlsHandle.getSheetNames(newExcel))
    newSheet = xlsHandle.createNewSheet(newExcel,u'测试sheet')
    print("SHEET1: 行数: %d, 列数： %d" % (xlsHandle.getRowCount(newSheet), xlsHandle.getCloumnCount(newSheet)) )
    for num1 in range(100):
        for num2 in range(1000):
            xlsHandle.setCellValue(newSheet,num2+1,num1+1,'行:'+str(num2+1)+' 列:'+str(num1+1))
    print("SHEET2: 行数: %d, 列数： %d" % (xlsHandle.getRowCount(newSheet), xlsHandle.getCloumnCount(newSheet)))
    print(xlsHandle.getSheetNum(newExcel))

    xlsHandle.saveAs(newExcel,"D:/python_excel.xlsx")
    '''

    print("Application finished!")
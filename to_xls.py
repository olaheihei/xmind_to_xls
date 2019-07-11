# -*- coding: utf:8-*-
# author：olaheihei

import xlwt, xlrd
from xmindparser import xmind_to_dict, xmind_to_xml, xmind_to_json
import pysnooper



class xmind_to_xx(object):
    '''
    xmind_to_xx('路径','app.xmind','app_v7.9')
    .xmind格式转xls、dict、json、xml方法
    xmind_to_xx.data_dict()
    xmind_to_xx.to_xml
    xmind_to_xx.to_json
    xmind_to_xx.to_excel
    '''
    def __init__(self, xmind_path, xmind_file, name):
        self.xmind_file_path = xmind_path + xmind_file
        self.data_dict = xmind_to_dict(self.xmind_file_path)
        self.xls_path = xmind_path + 'xls\\' + name + '.xls'
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet(self.data_dict[0]['topic']['title'])
        # self.workbook.save(self.xls_path)
        self.ex_row = 0


    def to_xml(self): 
        #转xml并存储
        return xmind_to_dict(self.xmind_file_path)

    def to_json(self):
        #转json格式并存储
        return xmind_to_json(self.xmind_file_path)

    def save_xls(self):
        self.workbook.save(self.xls_path)


    # @pysnooper.snoop()
    def to_excel(self, topics, ex_column = 0):
        #dict转xls
        self.write_excel(topics, self.ex_row, ex_column)
        if 'topics' in topics:

            for topic in list(topics['topics']):
                title_str = topic.get('title')
                self.to_excel(topic, ex_column + 1)
                if 'topics' not in topic and title_str is not None:
                    self.ex_row = self.ex_row + 1



    def write_excel(self, title, row, ex_column):
        #写入判断
        title_str = title['title']
        if title_str is not None:
            self.do_write_excel(title_str, row, ex_column)


    def do_write_excel(self, text, row, column):
        #执行写入
        self.worksheet.write(row, column, text)


class style_excel(object):
    '''
    xls格式设置样式
    style_excel('目录','app_v7.9.xls','app_v7.9')
    calculate 计算合并单元格坐标
    merge_excel 执行合并

    '''

    def __init__(self, xls_path, xls_file, name):
        self.style_workbook = xlwt.Workbook()   
        self.style_worksheet = self.style_workbook.add_sheet(name)
        path = xls_path + xls_file
        self.name = name
        self.data = xlrd.open_workbook(path).sheets()[0]
        self.row, self.column = self.data.nrows, self.data.ncols
        self.style_xls_path = xls_path + 'style_xml\\' + name + '.xls'
        self.set_style()
        self.set_level()

    def end_col(self, row, next_cols):
        #是否为最后一列
        return 'text' in str(next_cols[row])


    def do_merge_excel(self, num_start, num_end, column, text):
        #合并单元并写入内容
        self.style_worksheet.col(column).width =5000 #设置宽度
        self.style_worksheet.write_merge(num_start+1, num_end+1, column, column, text, self.style)


    def save_style_excel(self, path = 0):
        #保存
        if path == 0:
            path = self.style_xls_path
        self.style_workbook.save(path)

    def set_style(self):
        #设置样式
        self.style = xlwt.XFStyle()
        self.style1 = xlwt.XFStyle()
        am = xlwt.Alignment()   #对齐格式
        am1 = xlwt.Alignment()   #对齐格式
        font1 = xlwt.Font()  # 为样式创建字体
        font1.bold = True
        borders = xlwt.Borders()    #边框
        pattern = xlwt.Pattern()    #背景颜色
        pattern1 = xlwt.Pattern()    #背景颜色
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1
        pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern1.pattern_fore_colour = 22
        #设置边框
        borders.left = xlwt.Borders.THIN  
        borders.right = xlwt.Borders.THIN  
        borders.top = xlwt.Borders.THIN  
        borders.bottom = xlwt.Borders.THIN  
        am.vert = 0x01      # 设置水平居中
        am1.vert = 0x01     # 设置水平居中
        am1.horz = 0x02     # 设置垂直居中
        self.style.alignment = am
        self.style.borders = borders
        self.style1.alignment = am1
        self.style1.borders = borders
        self.style1.pattern = pattern1
        self.style.pattern = pattern
        self.style1.font = font1

    def set_level(self):
        #写入顶格样式、及内容
        for i in range(self.column):
            text = 'level' + str(i+1)
            self.style_worksheet.write(0,i,text,self.style1)

    # @pysnooper.snoop()
    def calculate(self):
        #计算
        self.do_merge_excel(0, self.row-1, 0, self.name)
        d = []
        for i in range(1,self.column):
            s, r, text = 0, 0, None
            l = [] 
            for o in self.data.col(i):
                if r == self.row-1:
                    try:
                        if 'text' in str(o):
                            if self.end_col(s, self.data.col(i+1)):
                                l.append([s,r-1,i,text])                            
                            else:
                                l.append([s,s,i,text])
                            text = str(o)[6:-1]
                            l.append([r,r,i,text])            
                        elif self.end_col(s, self.data.col(i+1)):
                            l.append([s,r,i,text])                            
                        else:
                            l.append([s,s,i,text])
                    except Exception as e:
                        l.append([s,s,i,text])
                elif 'text' in str(o):
                    if text == None:
                        text = str(o)[6:-1]
                    else:
                        try:
                            if self.end_col(s, self.data.col(i+1)):
                                l.append([s,r-1,i,text])
                            else:
                                l.append([s,s,i,text])
                            text = str(o)[6:-1]
                        except Exception as e:
                            text = str(o)[6:-1]
                            l.append([s,s,i,text])
                    s = r
                r += 1
            d.append(l)
        return d

    # @pysnooper.snoop()
    def merge_excel(self, d):
        #合并计算
        for i in d:
            for merge_row in i:
                if merge_row[1] != merge_row[0]:
                    r = 1
                    k = None
                    for o in range(merge_row[0],merge_row[1]):
                        try:
                            for c in range(0,merge_row[2]):
                                if k != None:
                                    break
                                if 'text' in str(self.data.col(c)[o+1]) and self.end_col(merge_row[0], self.data.col(merge_row[2]+1)):
                                    self.do_merge_excel(merge_row[0],merge_row[0]+(r-1),merge_row[2],merge_row[3])
                                    k = 1
                                elif merge_row[0] + r == merge_row[1] and c >= merge_row[2]-1 and o >= merge_row[1]-1:
                                    self.do_merge_excel(merge_row[0],merge_row[1],merge_row[2],merge_row[3])
                                    k = 1
                        except Exception as es:
                            pass
                        r += 1
                else:
                    self.do_merge_excel(merge_row[0],merge_row[1],merge_row[2],merge_row[3])



#转xls
a = xmind_to_xx('', 'demo.xmind', 'demo')
a.to_excel(a.data_dict[0]['topic'])
a.save_xls()

b = style_excel('xls/', 'demo.xls', a.data_dict[0]['topic']['title'])
b.merge_excel(b.calculate())
b.save_style_excel('xls/demo.xls')




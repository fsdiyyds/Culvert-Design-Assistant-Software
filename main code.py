# -*- coding: utf-8 -*-

"""
@Time    : 2020/9/22 17:26
@Author  : fenglei
@FileName: huiqiao.py
@Software: PyCharm

"""

import xlrd
from math import pi, cos, tan, sin, sqrt, atan
import datetime
from sys import exit as sy_exit
import pythoncom
from win32com.client import VARIANT
from win32com.client import Dispatch
from win32api import MessageBox
from win32con import MB_OK

# window = tk.Tk()  # 第1步，实例化object，建立窗口window
# # 第2步，给窗口的可视化起名字
# window.title('箱形桥设计     -------桥二所-XXX')
# # 第3步，设定窗口的大小(长 * 宽)
# window.geometry('600x400')  # 这里的乘是小x
# # 第4步，在图形界面上设定标签
# l = tk.Label(window, text='你好！请选择AutoCAD版本', bg='pink', font=('黑体', 12), width=75,
#              height=2).place(x=0, y=12)
# def select():
#     global version_cad
#     version_cad = comboxlist_plt.get()
#
#
# def button2_click():
#     global ju
#     ju = Fals
# comvalue_plt = tk.StringVar()
# comboxlist_plt = ttk.Combobox(window, textvariable=comvalue_plt, width=40)
# comboxlist_plt["values"] = ['AutoCAD-2014', 'AutoCAD-2016', 'AutoCAD-2018', 'AutoCAD-2020', '敬请期待...']
# comboxlist_plt.current(0)
# comboxlist_plt.bind("<<ComboboxSelected>>", select)
# comboxlist_plt.place(x=180, y=150)
# button2 = tk.Button(window, text="确定", command=button2_click)
# button2.place(x=490, y=145)

dict_cad_id = {'1': 19.1,
               '2': 20,
               '3': 22,
               '4': 23.1,
               '5': 24
               }

print("***********************************************************\n\n注意：本程序为桥二所设计箱形桥绘图程序，支持的CAD版本及对应的版本代码如下\n\n"
      "          AutoCAD2014------1\n"
      "          AutoCAD2016------2\n"
      "          AutoCAD2018------3\n"
      "          AutoCAD2020------4\n"
      "          AutoCAD2021------5\n")
version_cad = input("输入版本代码（数字1-4）然后回车运行程序")
prog_id = 'AutoCAD.Application.' + str(dict_cad_id[version_cad])
try:
    wincad = Dispatch(prog_id)  # 19.1
    wincad.Visible = 1
    try:
        doc = wincad.ActiveDocument
    except:
        wincad.Documents.Add()
        doc = wincad.ActiveDocument
    doc.Utility.Prompt("Hello! Autocad from pywin32com.\n")
    msp = doc.ModelSpace
    Layer_kuang = doc.Layers.Add("图框")
    Layer_biaoge = doc.Layers.Add("biaoge")
    Layer_jiegou = doc.Layers.Add("结构线")
    Layer_biaozhu = doc.Layers.Add("biaozhu")
    Layer_midline = doc.Layers.Add("中心线")
    Layer_elevation = doc.Layers.Add("高程标注")
except:
    MessageBox(0, "绘图失败！尝试重新打开运行！", "警告", MB_OK)
    sy_exit(0)

try:
    te_style = doc.TextStyles.Add("SMFS")
    te_style.FontFile = "txtwt.shx"
    te_style.BigFontFile = "smfs.shx"
    te_style.Width = 0.8
    te_style.Update
except:
    print('')


def vtpnt(x, y, z=0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtfloat(lst):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def vtint(lst):
    """列表转化为整数"""
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, lst)


def vtvariant(lst):
    """列表转化为变体"""
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, lst)


def add_text_mid(y1, y3, text, height=4.5):
    l = len(text) * 3.5
    y14 = [(y1[0] + y3[0]) / 2 - l / 2, y1[1]]
    y15 = [(y1[0] + y3[0]) / 2 + l / 2, y1[1]]
    y_text = vtpnt(y14[0], y14[1] + 2)
    doc.ActiveLayer = Layer_biaozhu
    textobj1 = msp.AddText(text, y_text, height)
    textobj1.StyleName = "SMFS"
    textobj1.ScaleFactor = 0.8
    return textobj1


def cal_l_length(lst1, lst2):
    l1 = lst1[1] - lst2[1]
    l2 = lst1[0] - lst2[0]
    l = (l1 ** 2 + l2 ** 2) ** 0.5
    return l


def cal_kjl_length(lst1, lst2):
    l1 = lst1[1] - lst2[1]
    l2 = lst1[0] - lst2[0]
    l3 = lst1[2] - lst2[2]
    l = (l1 ** 2 + l2 ** 2 + l3 ** 2) ** 0.5
    return l


def cal_s_triangle(lst1, lst2, lst3):
    a = cal_l_length(lst1, lst2) * scale / 100
    b = cal_l_length(lst2, lst3) * scale / 100
    c = cal_l_length(lst3, lst1) * scale / 100
    p = (a + b + c) / 2
    s_tri = (p * (p - a) * (p - b) * (p - c)) ** 0.5
    return s_tri


def cal_kjs_triangle(lst1, lst2, lst3):
    a = cal_kjl_length(lst1, lst2) * scale / 100
    b = cal_kjl_length(lst2, lst3) * scale / 100
    c = cal_kjl_length(lst3, lst1) * scale / 100
    p = (a + b + c) / 2
    s_tri = (p * (p - a) * (p - b) * (p - c)) ** 0.5
    return s_tri


def hui_tukuang(l):
    doc.ActiveLayer = Layer_kuang
    ClrNum = 7
    Layer_kuang.color = ClrNum
    Layer_kuang.Linetype = "ByBlock"
    poly_huizhi([[0, 0], [l, 0], [l, 297], [0, 297], [0, 0]])
    nei = poly_huizhi([[20, 5], [l - 5, 5], [l - 5, 293], [20, 293], [20, 5]])
    nei.ConstantWidth = 0.5
    p1 = vtpnt(l - 60, 315)
    text_tc = msp.AddText("图长为" + str(round(l, 1)), p1, 12)
    text_tc.Stylename = "SMFS"
    p2 = [l - 5 - 180, 5]
    p3 = [l - 5 - 180, 5 + 30]
    p4 = [l - 5, 5 + 30]
    poly1 = poly_huizhi([p2, p3, p4])
    poly1.ConstantWidth = 0.4

    p5 = [p2[0], p2[1] + 7.5]
    p6 = [p2[0], p2[1] + 15]
    p7 = [p2[0], p2[1] + 22.5]

    p51 = [p2[0] + 50, p2[1] + 7.5]
    p61 = [p2[0] + 50, p2[1] + 15]
    p71 = [p2[0] + 50, p2[1] + 22.5]
    p21 = [p2[0] + 50, p2[1]]

    p8 = [p2[0] + 25, p2[1]]
    p9 = [p2[0] + 25, p3[1]]
    p10 = [p2[0] + 50, p2[1]]
    p11 = [p2[0] + 50, p3[1]]
    line_huizhi(p5, p51)
    line_huizhi(p6, p61)
    line_huizhi(p7, p71)
    line_huizhi(p8, p9)
    line_huizhi(p10, p11)

    p52 = [p2[0] + 140, p2[1] + 7.5]
    p62 = [p2[0] + 140, p2[1] + 15]
    p72 = [p2[0] + 140, p2[1] + 22.5]
    p22 = [p2[0] + 140, p2[1]]
    p32 = [p3[0] + 140, p3[1]]

    p23 = [p2[0] + 155, p2[1]]
    p33 = [p3[0] + 155, p3[1]]
    p53 = [p2[0] + 155, p2[1] + 7.5]
    p63 = [p2[0] + 155, p2[1] + 15]
    p73 = [p2[0] + 155, p2[1] + 22.5]

    p54 = [p2[0] + 180, p2[1] + 7.5]
    p64 = [p2[0] + 180, p2[1] + 15]
    p74 = [p2[0] + 180, p2[1] + 22.5]

    p55 = [p2[0] + 25, p2[1] + 7.5]
    p65 = [p2[0] + 25, p2[1] + 15]
    p75 = [p2[0] + 25, p2[1] + 22.5]
    p25 = [p2[0] + 25, p2[1]]

    line_huizhi(p22, p32)
    line_huizhi(p53, p33)
    line_huizhi(p52, p54)
    line_huizhi(p62, p64)
    line_huizhi(p72, p74)

    def add_text(text, lst0):
        t1 = msp.AddText(text, vtpnt(lst0[0] + 5.23, lst0[1] + 1.4), 4.5)
        t1.ScaleFactor = 0.8
        t1.Stylename = "SMFS"
        return t1

    date = str(datetime.datetime.now().year) + "." + str(datetime.datetime.now().month)

    add_text_mid(p7, p75, "设计者", height=4.5)
    add_text_mid(p6, p65, "复核者", height=4.5)
    add_text_mid(p5, p55, "审核者", height=4.5)
    add_text_mid(p2, p25, "审定者", height=4.5)

    add_text_mid(p72, p73, "图 号", height=4.5)
    add_text_mid(p62, p63, "比例尺", height=4.5)
    add_text_mid(p52, p53, "日 期", height=4.5)

    add_text("XXXX", [p73[0] - 1.65, p73[1]])
    add_text("如  图", [p63[0] - 1.65, p63[1]])
    add_text(date, [p53[0] - 1.65, p53[1]])

    word1 = add_text("中铁第一勘察设计院集团有限公司", [p21[0] + 12.056, p21[1] + 22.22])
    word1.Height = 5.5

    title2 = str(licheng_qianzhui) + str(licheng) + " 1-" + str(l0 * scale / 100) + "m箱形桥全桥总布置图"

    add_text_mid(p21, p22, title2, height=4.5)
    # add_text(title2, [p21[0] + 8.58, p21[1] + 2.29])

    add_text("XXXXXXXXXXXXX工程", [p21[0] + 20.09, p21[1] + 16.3])
    add_text("施工图", [p21[0] + 37.64, p21[1] + 9.91])


class pou():
    Count = 0

    def __init__(self, hyi, scale, theta=0):
        cos_the = cos(pi * theta / 180)
        self.theta = theta
        self.slope = pou_slope
        self.hyi = hyi / scale
        self.left_l = pou_left_l / cos_the / scale
        self.right_l = pou_right_l / cos_the / scale
        self.jidi_h = 150 / scale
        self.dingkuan = pou_dingkuan / cos_the / scale
        self.d_bian = pou_jinbian / cos_the / scale
        self.ding_gao = 25 / scale
        self.di_chang = self.hyi * self.slope / cos_the


def line_hidden(lineobj, scale=0.5):
    # 用于给线段虚线，从而体现遮挡的效果
    try:
        doc.Linetypes.Load("HIDDEN", "acadiso.lin")
    except:
        pass
    lineobj.Linetype = "HIDDEN"
    lineobj.LinetypeScale = scale
    lineobj.Update
    return lineobj


def hui_lujian(lst1, lst2):
    doc.ActiveLayer = Layer_biaoge
    line_huizhi(lst1, lst2)
    p = vtpnt(min(lst2[0], lst1[0]), lst2[1] + biao_ju / 4)
    textobj1 = msp.AddText("设计路肩线", p, 2.5)
    textobj1.StyleName = "SMFS"

    doc.ActiveLayer = Layer_jiegou


def hui_guimian(lst1):
    doc.ActiveLayer = Layer_biaoge
    p = vtpnt(lst1[0], lst1[1] + biao_ju / 4)
    textobj1 = msp.AddText("设计轨面线", p, 3)
    textobj1.StyleName = "SMFS"
    doc.ActiveLayer = Layer_jiegou


def line_white(lineobj):
    # 用于给线段变成白色，改变图层
    lineobj.Layer = "0"
    lineobj.Update


def line_white_hidden(lineobj):
    # 用于给线段变成白色虚线，改变图层
    try:
        doc.Linetypes.Load("HIDDEN", "acadiso.lin")
    except:
        pass
    lineobj.Linetype = "HIDDEN"
    lineobj.LinetypeScale = 1.5
    lineobj.Layer = "0"
    lineobj.Update


def hatch_po(polyobj):
    ptnName, ptnType, bAss = "ANSI31", 0, True
    outerLoop = []
    outerLoop.append(polyobj)
    outerLoop = vtobj(outerLoop)
    hatchObj = msp.AddHatch(ptnType, ptnName, bAss)
    hatchObj.AppendOuterLoop(outerLoop)
    hatchObj.Evaluate()  # 进行填充计算，使图案吻合于边界
    hatchObj.PatternScale = 0.4  # 设置填充图案比例


def poly_huizhi(mubiao1):
    x = []
    for i in mubiao1:
        if i not in x and len(i) == 2:
            i.append(0)
            x.append(i)
        else:
            x.append(i)
    mubiao1 = [j for i in x for j in i]
    polyobj = msp.AddPolyLine(vtfloat(mubiao1))
    return polyobj


def line_huizhi(lst1, lst2):
    p1 = vtpnt(lst1[0], lst1[1])
    p2 = vtpnt(lst2[0], lst2[1])
    lineobj = msp.Addline(p1, p2)
    return lineobj


def add_mid_line(lst1, lst2):
    p1 = vtpnt(lst1[0], lst1[1])
    p2 = vtpnt(lst2[0], lst2[1])
    doc.ActiveLayer = Layer_midline
    Layer_midline.color = 6
    try:
        doc.Linetypes.Load("CENTER", "acadiso.lin")
    except:
        pass
    Layer_midline.Linetype = "CENTER"
    lineobj = msp.Addline(p1, p2)
    lineobj.LinetypeScale = 0.25
    doc.ActiveLayer = Layer_jiegou
    return lineobj


def add_licheng(li, lst0, lst2):
    li = float(li.split('+')[0]) * 1000 + float(li.split('+')[1]) + licheng_fx * (lst2[0] - lst0[0]) * scale / 100
    li_km = licheng_qianzhui + '%d+%.2f' % (li // 1000, li % 1000)
    p1 = vtpnt(lst2[0], lst2[1])
    p2 = vtpnt(lst2[0], lst2[1] + 110 / scale)
    p3 = vtpnt(lst2[0] + 200 / scale, lst2[1] + 110 / scale + 60 / scale)
    p4 = vtpnt(lst2[0] - 22.6 / scale, lst2[1] + 110 / scale + 4.7 / scale)
    doc.ActiveLayer = Layer_elevation
    msp.Addline(p1, p2)
    msp.Addline(p2, p3)
    textobj = msp.AddText(li_km, p4, 2.5)
    textobj.Rotation = atan(6 / 20)
    textobj.StyleName = "SMFS"

    doc.ActiveLayer = Layer_jiegou


def add_licheng_name(lst1, lst2, name_l, name_r):
    doc.ActiveLayer = Layer_biaozhu
    p1 = [lst1[0] - 8 * biao_ju, lst1[1]]
    p2 = [lst1[0] - 15 * biao_ju, lst1[1]]
    p3 = [lst1[0] - 13 * biao_ju, lst1[1] + 1.2 * biao_ju]
    p4 = vtpnt(lst1[0] - 12 * biao_ju, lst1[1] + 0.5 * biao_ju)
    poly1 = poly_huizhi([p1, p2, p3])
    t1 = msp.AddText(name_l, p4, 3)
    t1.StyleName = "SMFS"
    p5 = [lst2[0] + 8 * biao_ju, lst2[1]]
    p6 = [lst2[0] + 16 * biao_ju, lst2[1]]
    p7 = [lst2[0] + 14 * biao_ju, lst2[1] + 1.2 * biao_ju]
    p8 = vtpnt(lst2[0] + 10 * biao_ju, lst2[1] + 0.5 * biao_ju)
    poly2 = poly_huizhi([p5, p6, p7])
    t2 = msp.AddText(name_r, p8, 3)
    t2.StyleName = "SMFS"
    doc.ActiveLayer = Layer_jiegou


def add_name(y1, y3, text, bili=""):
    l = len(text) * 3.5

    y14 = [(y1[0] + y3[0]) / 2 - l / 2, y1[1] - 14]
    y15 = [(y1[0] + y3[0]) / 2 + l / 2, y1[1] - 14]
    y_text = vtpnt(y14[0], y14[1] + 2)
    doc.ActiveLayer = Layer_biaozhu
    textobj1 = msp.AddText(text, y_text, 4.5)
    textobj1.StyleName = "SMFS"
    textobj1.ScaleFactor = 0.8
    y16 = [(y1[0] + y3[0]) / 2 - l / 2 + 0.575, y1[1] - 14.575]
    y17 = [(y1[0] + y3[0]) / 2 + l / 2 - 0.575, y1[1] - 14.575]
    line_huizhi(y16, y17)
    if bili != "":
        y_text2 = vtpnt(y14[0] + l / 2 - 5.5, y14[1] - 4.35)
        textobj2 = msp.AddText(bili, y_text2, 3)
        textobj2.StyleName = "SMFS"
        textobj2.ScaleFactor = 0.8
    doc.ActiveLayer = Layer_jiegou
    poly_huizhi([y14, y15])
    return textobj1


def add_jiemian_num(lst1, lst2, te_str, fangx):
    doc.ActiveLayer = Layer_biaozhu
    if fangx == 1:  # 水平方向
        p1 = [lst1[0] - biao_ju, lst1[1] + biao_ju]
        p2 = [lst1[0] - 2.5 * biao_ju, lst1[1] + biao_ju]
        p3 = [lst1[0] - 2.5 * biao_ju, lst1[1] - 0.5 * biao_ju]
        p4 = vtpnt(lst1[0] - 4.1 * biao_ju, lst1[1])
        poly1 = poly_huizhi([p1, p2, p3])
        poly1.ConstantWidth = 0.4
        tex_1 = msp.AddText(te_str, p4, 3)
        p5 = [lst2[0] + biao_ju, lst1[1] + biao_ju]
        p6 = [lst2[0] + 2.5 * biao_ju, lst1[1] + biao_ju]
        p7 = [lst2[0] + 2.5 * biao_ju, lst1[1] - 0.5 * biao_ju]
        p8 = vtpnt(lst2[0] + 2.8 * biao_ju, lst1[1])
        poly2 = poly_huizhi([p5, p6, p7])
        poly2.ConstantWidth = 0.3
        tex_2 = msp.AddText(te_str, p8, 3)

    if fangx == 0:  # 垂直方向
        p1 = [lst1[0], lst1[1] + biao_ju]
        p2 = [lst1[0], lst1[1] + 2.5 * biao_ju]
        p3 = [lst1[0] - 1.5 * biao_ju, lst1[1] + 2.5 * biao_ju]
        p4 = vtpnt(lst1[0] - 2.2 * biao_ju, lst1[1] + 3.0 * biao_ju)
        poly1 = poly_huizhi([p1, p2, p3])
        poly1.ConstantWidth = 0.4
        tex_1 = msp.AddText(te_str, p4, 3)
        p5 = [lst1[0], lst2[1] - biao_ju]
        p6 = [lst1[0], lst2[1] - 2.5 * biao_ju]
        p7 = [lst1[0] - 1.5 * biao_ju, lst2[1] - 2.5 * biao_ju]
        p8 = vtpnt(lst1[0] - 2.2 * biao_ju, lst2[1] - 3.7 * biao_ju)
        poly2 = poly_huizhi([p5, p6, p7])
        poly2.ConstantWidth = 0.3
        tex_2 = msp.AddText(te_str, p8, 3)
    doc.ActiveLayer = Layer_jiegou
    tex_1.StyleName = "SMFS"
    tex_2.StyleName = "SMFS"
    total_lst = [poly1, poly2, tex_1, tex_2]
    return total_lst


def add_elevation(elevation0, lst0, lst2):
    real_ele = round(elevation0 + (lst2[1] - lst0[1]) * scale / 100, 2)
    p0 = vtpnt(lst2[0], lst2[1])
    l1 = biao_ju * tan(pi * 30 / 180)
    p1 = vtpnt(lst2[0] - l1, lst2[1] + biao_ju)
    p2 = vtpnt(lst2[0] + l1, lst2[1] + biao_ju)
    doc.ActiveLayer = Layer_elevation
    Layer_elevation.color = 2
    msp.Addline(p0, p2)
    msp.Addline(p1, p2)
    msp.Addline(p0, p1)

    p3 = vtpnt(lst2[0] - 6 * biao_ju, lst2[1] + 0.6 * biao_ju)
    textobj = msp.AddText(str(real_ele), p3, 2.5)
    textobj.StyleName = "SMFS"
    doc.ActiveLayer = Layer_jiegou
    return textobj


def mirror_zuoyou(obj_list, plst1, plst2):
    mirror_obj = []
    p1 = vtpnt(plst1[0], plst1[1])
    p2 = vtpnt(plst2[0], plst2[1])
    for obj in obj_list:
        try:
            mirror = obj.Mirror(p1, p2)
            mirror_obj.append(mirror)
        except:
            pass
    return mirror_obj


def line_biaozhu(x_1, x_2, x_3, theta=0, bu_ping=False):
    doc.ActiveLayer = Layer_biaozhu
    ClrNum = 7
    Layer_biaozhu.color = ClrNum

    # ExtLine1Point = vtpnt(x_1[0], x_1[1])
    # ExtLine2Point = vtpnt(x_2[0], x_2[1])
    # TextPosition = vtpnt((x_3[0]), x_3[1])
    if bu_ping:
        rotAngle = pi / 2
        if theta != 0:
            text = str(abs(round((x_2[1] - x_1[1]) * scale * cos(pi * theta / 180)))) + '/cos' + str(
                abs(theta)) + '°'
        else:
            text = str(abs(round((x_2[1] - x_1[1]) * scale * cos(pi * theta / 180))))
        if x_2[1] - x_1[1] != 0:
            if x_3[0] >= max(x_1[0], x_2[0]):
                ExtLine1Point = vtpnt(x_3[0] - 2 * biao_ju, x_1[1])
                ExtLine2Point = vtpnt(x_3[0] - 2 * biao_ju, x_2[1])
            else:
                ExtLine1Point = vtpnt(x_3[0] + 2 * biao_ju, x_1[1])
                ExtLine2Point = vtpnt(x_3[0] + 2 * biao_ju, x_2[1])

            TextPosition = vtpnt((x_3[0]), x_3[1])
            Dim_liang = msp.AddDimRotated(ExtLine1Point, ExtLine2Point, TextPosition, rotAngle)
    else:
        rotAngle = 0.0
        if theta != 0:
            text = str(abs(round((x_2[0] - x_1[0]) * scale * cos(pi * theta / 180)))) + '/cos' + str(
                abs(theta)) + '°'
        else:
            text = str(abs(round((x_2[0] - x_1[0]) * scale * cos(pi * theta / 180))))
        if x_2[0] - x_1[0] != 0:
            if x_3[1] > max(x_1[1], x_2[1]):
                ExtLine1Point = vtpnt(x_1[0], x_3[1] - 1.1 * biao_ju)
                ExtLine2Point = vtpnt(x_2[0], x_3[1] - 1.1 * biao_ju)
            else:
                ExtLine1Point = vtpnt(x_1[0], x_3[1] + 2 * biao_ju)
                ExtLine2Point = vtpnt(x_2[0], x_3[1] + 2 * biao_ju)
            TextPosition = vtpnt((x_3[0]), x_3[1])
            Dim_liang = msp.AddDimRotated(ExtLine1Point, ExtLine2Point, TextPosition, rotAngle)
    # Dim_liang = msp.AddDimRotated(ExtLine1Point, ExtLine2Point, TextPosition, rotAngle)
    # AddDimAligned(ExtLine1Point, ExtLine2Point, TextPosition)
    try:
        Dim_liang.TextOverride = text
        Dim_liang.TextInside = True
        Dim_liang.Fit
        Dim_liang.TextStyle = "SMFS"
        Dim_liang.TextHeight = 25 / scale
        doc.ActiveLayer = Layer_jiegou
        return Dim_liang
    except:
        pass
        doc.ActiveLayer = Layer_jiegou


def hui_xiang(x0, y0, l0, h0, d_le, d1, d2, d3, theta):
    doc.ActiveLayer = Layer_jiegou
    ClrNum = 1
    Layer_jiegou.color = ClrNum
    Layer_jiegou.Linetype = "ByBlock"
    cos_t = cos(pi * theta / 180)
    x1 = [x0, y0]
    x2 = [x0, y0 + h0 + d1 + d2]
    x3 = [x0 + (2 * d_le + l0) / cos_t, y0 + h0 + d1 + d2]
    x4 = [x0 + (2 * d_le + l0) / cos_t, y0]
    poly_huizhi([x4, x1, x2, x3, x4])

    xj1 = [x1[0] - xiang_jichu_l / cos_t, x1[1]]
    xj2 = [x1[0] - xiang_jichu_l / cos_t, x1[1] - xiang_jichu_h]
    xj3 = [x4[0] + xiang_jichu_l / cos_t, x4[1] - xiang_jichu_h]
    xj4 = [x4[0] + xiang_jichu_l / cos_t, x4[1]]
    poly_huizhi([xj4, xj1, xj2, xj3, xj4])
    b4 = [x1[0], x1[1] + biao_ju / 2]
    line_biaozhu(xj1, x1, b4, theta=theta, bu_ping=False)
    line_biaozhu(xj4, x4, b4, theta=theta, bu_ping=False)

    x5 = [x0, y0 + h0 + d1 + d2 + d3 + (l0 * x_slope / 2 / cos_t)]
    x6 = [x0 + (2 * d_le + l0) / cos_t, y0 + h0 + d1 + d2 + d3 - (l0 * x_slope / 2 / cos_t)]
    xgui = [x5[0] + (x6[0] - x5[0]) / 4, x5[1]]
    line_hidden(line_huizhi(x5, x6))

    x7 = [x0 + (l_lei_down + d_le) / cos_t, y0 + d2]
    x8 = [x0 + d_le / cos_t, y0 + d2 + h_lei_down]
    x9 = [x0 + d_le / cos_t, y0 + h0 - h_lei_up + d2]
    x10 = [x0 + (l_lei_up + d_le) / cos_t, y0 + h0 + d2]
    x11 = [x0 + (l0 + d_le - l_lei_up) / cos_t, y0 + h0 + d2]
    x12 = [x0 + (l0 + d_le) / cos_t, y0 + h0 + d2 - h_lei_up]
    x13 = [x0 + (l0 + d_le) / cos_t, y0 + d2 + h_lei_down]
    x14 = [x0 + (l0 + d_le - l_lei_down) / cos_t, y0 + d2]
    x15 = [x5[0], x5[1] - x_guilu]
    x16 = [x6[0], x6[1] - x_guilu]
    x17 = [x5[0] - d_le, x5[1] - x_guilu + d_le * x_slope]
    x18 = [x6[0] + d_le, x6[1] - x_guilu - d_le * x_slope]
    hui_lujian(x15, x17)
    hui_lujian(x16, x18)
    hui_guimian(xgui)

    poly_huizhi([x7, x8, x9, x10, x11, x12, x13, x14, x7])

    b1 = [x8[0], x8[1] + 6 * biao_ju]
    line_biaozhu(x8, x12, b1, theta=theta, bu_ping=False)
    line_biaozhu(x1, x8, b1, theta=theta, bu_ping=False)
    line_biaozhu(x12, x4, b1, theta=theta, bu_ping=False)

    b2 = [x9[0], x9[1] - 2 * biao_ju]
    line_biaozhu(x9, x10, b2, theta=theta, bu_ping=False)
    line_biaozhu(x9, x10, [x10[0] + 2 * biao_ju, x10[1]], theta=0, bu_ping=True)
    line_biaozhu(x8, x7, [x8[0], x8[1] + 2 * biao_ju], theta=theta, bu_ping=False)
    line_biaozhu(x8, x7, [x7[0] + 2 * biao_ju, x7[1]], theta=0, bu_ping=True)
    line_biaozhu(x2, x5, [x5[0] + 2 * biao_ju, x5[1]], theta=0, bu_ping=True)
    line_biaozhu(x3, x6, [x6[0], x6[1]], theta=0, bu_ping=True)

    x_mid = (x2[0] + x3[0]) / 2
    x1_5 = [x_mid, x2[1]]
    x2_5 = [x_mid, x10[1]]
    x3_5 = [x_mid, x7[1]]
    x4_5 = [x_mid, x1[1]]
    x5_5 = [x_mid, xj2[1]]
    add_mid_line([x1_5[0], x1_5[1] + 2 * biao_ju], [x4_5[0], x4_5[1] - 2 * biao_ju])

    b3 = [x1_5[0] + 4 * biao_ju, x1_5[1]]
    line_biaozhu(x1_5, x2_5, b3, theta=0, bu_ping=True)
    line_biaozhu(x2_5, x3_5, b3, theta=0, bu_ping=True)
    line_biaozhu(x3_5, x4_5, b3, theta=0, bu_ping=True)
    line_biaozhu(x5, x15, [x15[0] - 4 * biao_ju, x15[1]], theta=0, bu_ping=True)
    line_biaozhu(x6, x16, [x16[0] + 4 * biao_ju, x16[1]], theta=0, bu_ping=True)
    line_biaozhu(x4_5, x5_5, b3, theta=0, bu_ping=True)

    x0_5 = [x1_5[0], x1_5[1] + d3]
    for x in [x0_5, x5, x6, x1_5, x2_5, x3_5, x4_5, x5_5]:
        add_elevation(elevation0, x0_5, x)

    for x in [x0_5, x5, x6]:
        add_licheng(licheng, x0_5, x)
    tuchang = abs(x4[0] - x1[0])
    add_name([x1[0], xj1[1] - 2.2 * biao_ju], [x3[0], x3[1] - 2.0 * biao_ju], "箱身正面", "（1:100）")

    j2 = [(x1[0] + x4[0]) / 2, xj2[1] + 0.9 * biao_ju]
    j1 = [(x5[0] + x6[0]) / 2, x5[1] + 5 * biao_ju]
    add_jiemian_num(j1, j2, "Ⅰ", 0)

    add_licheng_name(x5, x6, licheng_name_l, licheng_name_r)

    return tuchang


def limian_yiqiang(x0, y0, l0, d_le, d2, theta, afa, right):
    global pou3, pou5
    cos_t = cos(pi * theta / 180)
    tan_t = tan(pi * theta / 180)
    tan_a = tan(pi * afa / 180)
    daierta = yanshen * tan_t
    xishu = -1
    tan_a = sin(pi * (-afa) / 180)
    if theta > 0:
        if right:
            tan_a = tan(pi * (theta) / 180)
        else:
            tan_a = tan(pi * (-afa) / 180)
    if theta < 0:
        if right:
            tan_a = tan(pi * (afa) / 180)
        else:
            tan_a = tan(pi * (-theta) / 180)

    if not right:
        pou3, pou5 = pou2, pou4
        daierta = -daierta
        if theta > 0:
            pian_afa = ((pou3.hyi - pou5.hyi) * 1.5 + pou5.dingkuan) * (-tan_a)
        else:
            pian_afa = ((pou3.hyi - pou5.hyi) * 1.5 + pou5.dingkuan) * (-tan_a)
    if right:
        x0 = x0 + (l0 + 2 * d_le) / cos_t
        xishu = 1
        if theta > 0:
            pian_afa = ((pou3.hyi - pou5.hyi) * 1.5 + pou5.dingkuan) * tan_a
        else:
            pian_afa = ((pou3.hyi - pou5.hyi) * 1.5 + pou5.dingkuan) * tan_a

    p1 = [x0 - xishu * d_le / cos_t, y0 + d2]
    p2 = [x0 - xishu * d_le / cos_t - xishu * pou5.right_l, y0 + d2]
    p3 = [x0 - xishu * (d_le / cos_t + pou5.right_l), y0 + d2 - pou3.jidi_h]
    p4 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi + xishu * pou5.right_l,
          y0 + d2 - pou3.jidi_h]
    p5 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi + xishu * pou5.right_l,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h]
    p6 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h]
    p7 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    p8 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    p9 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa, y0 + d2 - pou3.jidi_h + pou5.jidi_h]
    p23 = [x0 - xishu * d_le / cos_t + xishu * daierta, y0 + d2]
    p92 = [p9[0] - xishu * pou5.right_l, p23[1]]
    p93 = [p9[0] - xishu * pou5.right_l, p23[1] - pou5.jidi_h]
    if right:
        if p9[0] < p2[0]:
            p2 = p92
            p3 = p93
        if theta < 0:
            line_biaozhu(p9, p2, [p3[0], p3[1] - 2 * biao_ju], theta=0, bu_ping=False)

        if theta >= 0:
            line_biaozhu(p1, p2, [p3[0], p3[1] - 2 * biao_ju], theta=0, bu_ping=False)
    else:
        if p92[0] > p2[0]:
            p2 = p92
            p3 = p93
        if theta < 0:
            line_biaozhu(p1, p2, [p3[0], p3[1] - 2 * biao_ju], theta=0, bu_ping=False)

        if theta >= 0:
            line_biaozhu(p9, p2, [p3[0], p3[1] - 2 * biao_ju], theta=0, bu_ping=False)

    p10 = [x0 - xishu * d_le / cos_t, y0 + d2]
    p11 = [x0 - xishu * d_le / cos_t, y0 + d2 + pou3.hyi]
    p12 = [x0 - xishu * d_le / cos_t - xishu * pou3.d_bian, y0 + d2 + pou3.hyi]
    p13 = [x0 - xishu * d_le / cos_t - xishu * pou3.d_bian,
           y0 + d2 + pou3.hyi + pou3.ding_gao - pou3.d_bian]
    p14 = [x0 - xishu * d_le / cos_t, y0 + d2 + pou3.hyi + pou3.ding_gao]
    p15 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan, y0 + d2 + pou3.hyi + pou3.ding_gao]
    p16 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan, y0 + d2 + pou3.hyi]

    poly_huizhi([p1, p2, p3, p4, p5, p6, p7, p8, p9])
    # limian_right = [p1, p2, p3, p4, p5, p6, p7, p8, p9]
    xu_gai = poly_huizhi([p10, p11, p12, p13, p14, p15, p16, p11])

    if not right:
        if theta > 0:
            line_hidden(xu_gai)
    else:
        if theta < 0:
            line_hidden(xu_gai)

    p17 = [x0 - xishu * d_le / cos_t + xishu * daierta, y0 + d2 + pou3.hyi]
    p18 = [x0 - xishu * d_le / cos_t - xishu * pou3.d_bian + xishu * daierta, y0 + d2 + pou3.hyi]
    p19 = [x0 - xishu * d_le / cos_t - xishu * pou3.d_bian + xishu * daierta,
           y0 + d2 + pou3.hyi + pou3.ding_gao - pou3.d_bian]
    p20 = [x0 - xishu * d_le / cos_t + xishu * daierta, y0 + d2 + pou3.hyi + pou3.ding_gao]
    p21 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan + xishu * daierta,
           y0 + d2 + pou3.hyi + pou3.ding_gao]
    p22 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan + xishu * daierta, y0 + d2 + pou3.hyi]
    poly_huizhi([p18, p19, p20, p21])

    h1 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    h2 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi + xishu * pou5.d_bian,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    h3 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi + xishu * pou5.d_bian,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi + pou5.ding_gao - pou5.d_bian]
    h4 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou5.hyi,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi + pou5.ding_gao]
    h5 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi + pou5.ding_gao]
    h6 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa - xishu * pou5.d_bian,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi + pou5.ding_gao - pou5.d_bian]
    h7 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa - xishu * pou5.d_bian,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    h8 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi]
    h9 = [x0 - xishu * d_le / cos_t + xishu * daierta + pian_afa + xishu * pou3.dingkuan,
          y0 + d2 - pou3.jidi_h + pou5.jidi_h + pou5.hyi + pou5.ding_gao]

    x1 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan + xishu * pou3.di_chang, y0 + d2]
    x2 = [x0 - xishu * d_le / cos_t + xishu * pou3.dingkuan + xishu * pou3.di_chang + xishu * daierta,
          y0 + d2]

    if right:
        if x1[0] > p5[0]:
            x3 = [x1[0] + pou3.right_l, x1[1]]
            x4 = [x1[0] + pou3.right_l, x1[1] - pou3.jidi_h]
            line_huizhi(p5, x3)
            line_huizhi(x3, x4)
            line_huizhi(x4, p4)

    poly_huizhi([h1, h2, h3, h4, h5, h6, h7, h8])
    line_huizhi(p6, p9)
    line_huizhi(p1, p9)
    line_huizhi(h6, h3)
    line_huizhi(h7, p18)
    line_huizhi(h6, p19)
    line_huizhi(h5, p20)
    line_huizhi(h9, p21)
    line_huizhi(p20, p23)
    line_huizhi(p18, p22)
    line1 = line_huizhi(x1, p16)
    line2 = line_huizhi(x2, p22)
    line_hidden(line1)
    line_hidden(line2)
    line_w1 = line_huizhi(p13, p19)
    line_w2 = line_huizhi(p15, p20)
    # line_white(line_w1)
    # line_white(line_w2)
    line_wh1 = line_huizhi(p21, p22)
    line_wh2 = line_huizhi(p22, p16)

    b1 = [p3[0], p3[1] - 2 * biao_ju]

    line_biaozhu(p5, p6, b1, theta=0, bu_ping=False)
    line_biaozhu(p6, p9, b1, theta=0, bu_ping=False)
    line_biaozhu(p23, p9, b1, theta=0, bu_ping=False)

    b2 = [p4[0] + xishu * 2 * biao_ju, p4[1]]
    line_biaozhu(p4, p5, b2, theta=0, bu_ping=True)
    line_biaozhu(p5, p21, b2, theta=0, bu_ping=True)

    b3 = [p5[0], p5[1]]
    line_biaozhu(p6, p7, b3, theta=0, bu_ping=True)
    line_biaozhu(p7, h4, b3, theta=0, bu_ping=True)
    line_biaozhu(h4, p22, b3, theta=0, bu_ping=True)
    line_biaozhu(p22, p21, b3, theta=0, bu_ping=True)
    tuchang = abs(p5[0] - p1[0])
    return tuchang


def cemian_huizhi(x00, y00, w0, h0, h_lei_up, h_lei_down, d1, d2, d3, theta):
    diju = pou2.jidi_h - d2
    x1 = [x00, y00]
    x2 = [x00 + w0, y00]
    x3 = [x00 + w0, y00 + diju]
    x4 = [x00, y00 + diju]
    x5 = [x00, y00 + diju + d2]
    x6 = [x00 + w0, y00 + diju + d2]
    x7 = [x00 + w0, y00 + diju + d2 + h_lei_down]
    x8 = [x00, y00 + diju + d2 + h_lei_down]
    x9 = [x00, y00 + diju + d2 + h0 - h_lei_up]
    x10 = [x00 + w0, y00 + diju + d2 + h0 - h_lei_up]
    x11 = [x00 + w0, y00 + diju + d2 + h0]
    x12 = [x00, y00 + diju + d2 + h0]
    x13 = [x00, y00 + diju + d2 + h0 + d1]
    x14 = [x00 + w0, y00 + diju + d2 + h0 + d1]
    x0_5 = [x00 + w0 / 2, x14[1] + d3]
    x15 = [x13[0] - yanshen - 3 / scale, x0_5[1] - x_guilu]
    x16 = [x15[0] - d_le, x15[1]]
    x17 = [x14[0] + yanshen + 3 / scale, x0_5[1] - x_guilu]
    x18 = [x17[0] + d_le, x17[1]]
    hui_lujian(x15, x16)
    hui_lujian(x17, x18)

    x1_5 = [x00 + w0 / 2, x1[1] - 2 * biao_ju]
    x3_5 = [x00 + w0 / 2, x3[1]]
    x5_5 = [x00 + w0 / 2, x5[1]]
    x11_5 = [x00 + w0 / 2, x11[1]]
    x13_5 = [x00 + w0 / 2, x4[1] - xiang_jichu_h]  # 需要减去桥底下地基的厚度
    xl13_5 = [x00, x4[1] - xiang_jichu_h]
    xr13_5 = [x00 + w0, x4[1] - xiang_jichu_h]
    poly_huizhi([x4, x3, xr13_5, xl13_5, x4])
    for x in [x0_5, x3_5, x5_5, x11_5, x13_5]:
        add_elevation(elevation0, x0_5, x)
    add_mid_line(x1_5, [x00 + w0 / 2, x13[1] + d3 + 2 * biao_ju])
    line_hidden(line_huizhi([x0_5[0] - w0 / 2, x0_5[1]], [x0_5[0] + w0 / 2, x0_5[1]]))
    hui_guimian([x0_5[0] + w0 / 4, x0_5[1]])

    line_huizhi(x7, x8)
    line_huizhi(x9, x10)
    ding2 = [x00 + w0, y00 + diju + d2 + h0 + d1, 0, x00, y00 + diju + d2 + h0 + d1, 0,
             x00,
             y00, 0]
    msp.AddPolyLine(vtfloat(ding2))
    ding1 = [x00 + w0, y00 + diju, 0, x00, y00 + diju, 0, x00, y00 + diju + d2, 0, x00 + w0, y00 + diju + d2, 0,
             x00 + w0,
             y00 + diju, 0]
    poly1 = msp.AddPolyLine(vtfloat(ding1))
    hatch_po(poly1)
    ding3 = [x00 + w0, y00 + diju + d2 + h0, 0, x00, y00 + diju + d2 + h0, 0, x00, y00 + diju + d2 + h0 + d1, 0,
             x00 + w0,
             y00 + diju + d2 + h0 + d1, 0, x00 + w0, y00 + diju + d2 + h0, 0]
    poly2 = msp.AddPolyLine(vtfloat(ding3))
    hatch_po(poly2)

    y1 = [x00 - fenxi, y00]
    y2 = [x00 - pou4.right_l - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi, y00]
    y3 = [x00 - pou4.right_l - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi, y00 + pou4.jidi_h]
    y4 = [x00 - yanshen - fenxi, y00 + pou4.jidi_h]
    y5 = [x00 - fenxi, y00 + pou4.jidi_h]
    y6 = [x00 - fenxi, y00 + pou4.jidi_h + pou2.hyi]
    y7 = [x00 - fenxi - yanshen, y00 + pou4.jidi_h + pou2.hyi]
    y8 = [x00 - fenxi - yanshen, y00 + pou4.jidi_h]
    y9 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi, y00 + pou4.jidi_h]
    y10 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi, y00 + pou4.jidi_h + pou4.hyi]
    y11 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - yanshen - fenxi, y00 + pou4.jidi_h + pou4.hyi]
    y12 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - pou4.d_bian,
           y00 + pou4.jidi_h + pou4.hyi]
    y13 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - pou4.d_bian,
           y00 + pou4.jidi_h + pou4.hyi + pou4.ding_gao - pou4.d_bian]
    y14 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi,
           y00 + pou4.jidi_h + pou4.hyi + pou4.ding_gao]
    y15 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - yanshen - fenxi, y00 + pou4.jidi_h + pou4.hyi + pou4.ding_gao]
    y16 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - yanshen - fenxi,
           y00 + pou4.jidi_h + pou4.hyi + pou4.ding_gao - pou4.d_bian]
    y17 = [x00 - fenxi - yanshen, y00 + pou4.jidi_h + pou2.hyi + pou4.ding_gao - pou4.d_bian]
    y18 = [x00 - fenxi - yanshen, y00 + pou4.jidi_h + pou2.hyi + pou4.ding_gao]
    y19 = [x00 - fenxi, y00 + pou4.jidi_h + pou2.hyi + pou4.ding_gao]
    y20 = [x00 - fenxi, y00 + pou4.jidi_h + pou2.hyi + pou4.ding_gao - pou4.d_bian]
    y21 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + pou4.hyi * pou4.slope + pou4.dingkuan,
           y00 + pou4.jidi_h]

    z1 = y9
    z2 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w, y00 + pou4.jidi_h]
    z3 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w, y00 + pou4.jidi_h - zhi_l]
    z4 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi, y00 + pou4.jidi_h - zhi_l]
    k1 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w / 2 - kang_w / 2,
          y00 + pou4.jidi_h]
    k2 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w / 2 + kang_w / 2,
          y00 + pou4.jidi_h]
    k3 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w / 2 + kang_w / 2,
          y00 + pou4.jidi_h - kang_l]
    k4 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi + zhi_w / 2 - kang_w / 2,
          y00 + pou4.jidi_h - kang_l]

    zp1 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi,
           y00 + pou4.jidi_h + pou4.hyi - pou4.dingkuan / 1.15]
    zp2 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi,
           y00 + pou4.jidi_h + pou4.hyi - pou4.dingkuan / 1.15 - zhuiti1]
    zp3 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi,
           y00 + pou4.jidi_h + pou4.hyi - pou4.dingkuan / 1.15 - zhuiti1 - zhuiti2]
    zp4 = [
        x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - (pou4.hyi - pou4.dingkuan / 1.15) * 1.5,
        y00 + pou4.jidi_h]
    zp5 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - (
            pou4.hyi - pou4.dingkuan / 1.15 - zhuiti1) * 1.5, y00 + pou4.jidi_h]
    zp6 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - (
            pou4.hyi - pou4.dingkuan / 1.15 - zhuiti1 - zhuiti2) * 1.5, y00 + pou4.jidi_h]
    zp7 = [
        x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - (
                pou4.hyi - pou4.dingkuan / 1.15) * 1.5 - 20 / scale,
        y00 + pou4.jidi_h - zhuiti_shen]
    zp8 = [x00 - (pou2.hyi - pou4.hyi) * 1.5 - pou4.dingkuan - yanshen - fenxi - (
            pou4.hyi - pou4.dingkuan / 1.15 - zhuiti1 - zhuiti2) * 1.5 - 20 / scale, y00 + pou4.jidi_h - zhuiti_shen]
    zhuiti = poly_huizhi([zp1, zp4, zp5, zp2, zp3, zp6, zp5, zp6, zp8, zp7, zp4])

    x_line1 = line_hidden(line_huizhi(y11, y21))
    polyline1 = poly_huizhi([y1, y2, y3, y4, y5, y19, y6, y7, y8, y5, y1])
    polyline2 = poly_huizhi(
        [y10, y9, y4, y7, y11, y12, y13, y14, y15, y16, y13, y16, y11, y16, y17, y7, y18, y19, y20, y17, y18, y15])

    zhicheng_liang = line_hidden(poly_huizhi([z1, z2, z3, z4, z1]))
    kanghua_zhuang = poly_huizhi([k1, k2, k3, k4, k1])

    b1 = [x8[0], x8[1] + 5 * biao_ju]
    bz1 = line_biaozhu(x5_5, x8, b1, bu_ping=False)
    bz2 = line_biaozhu(y1, x8, [x8[0], x8[1] + 5 * biao_ju], bu_ping=False)
    bz3 = line_biaozhu(y1, y4, b1, bu_ping=False)
    bz4 = line_biaozhu(y4, y11, b1, bu_ping=False)
    bz5 = line_biaozhu(y11, y10, b1, bu_ping=False)

    bz6 = line_biaozhu(y9, y21, [y21[0], y21[1] + biao_ju], bu_ping=False)
    bz7 = line_biaozhu(y9, y3, [y21[0], y21[1] + biao_ju], bu_ping=False)

    bz8 = line_biaozhu(y2, y3, [y2[0] - biao_ju, y2[1]], bu_ping=True)
    line_biaozhu(y3, y10, [y2[0] - biao_ju, y2[1]], bu_ping=True)
    line_biaozhu(y10, y14, [y2[0] - biao_ju, y2[1]], bu_ping=True)
    line_biaozhu(y14, y7, [y2[0] - biao_ju, y2[1]], bu_ping=True)
    line_biaozhu(y7, y18, [y2[0] - biao_ju, y2[1]], bu_ping=True)
    bz11 = line_biaozhu(y3, y18, [y2[0] - 2.5 * biao_ju, y2[1]], bu_ping=True)
    bz12 = line_biaozhu(z2, z3, [z2[0] + 2 * biao_ju, z2[1]], bu_ping=True)
    bz13 = line_biaozhu(z3, z4, [z2[0], z2[1] - 2 * biao_ju], bu_ping=False)
    bz14 = line_biaozhu(k3, k4, [k4[0], k4[1] - 2 * biao_ju], bu_ping=False)
    bz15 = line_biaozhu(k2, k4, [k2[0] + 3 * biao_ju, k2[1]], bu_ping=True)

    add_name(x1, x2, "Ⅰ--Ⅰ剖面", "（1:100）")
    m1 = [x00 + w0 / 2, y00]
    m2 = [x00 + w0 / 2, y00 + h0]
    mirror_zuoyou([polyline1, polyline2, x_line1, zhicheng_liang, kanghua_zhuang, zhuiti], m1, m2)
    tuchang = 2 * abs(x1_5[0] - zp4[0])
    length1 = abs(x1[0] - zp4[0])

    return tuchang, length1


def pingmian_huizhi(x00, y00, l0, w0, d_lef, theta, afa):
    cos_t = cos(pi * theta / 180)
    sin_t = sin(pi * theta / 180)
    tan_t = tan(pi * theta / 180)
    tan_a = tan(pi * afa / 180)
    cos_a = cos(pi * afa / 180)
    x0 = [x00, y00]
    x1 = [x00 - fenxi / cos_t - pou2.right_l, y00]
    x2 = [x00 - fenxi / cos_t - pou2.right_l - d_lef / cos_t, y00]
    x3 = [x00 - fenxi / cos_t - pou2.right_l - (d_lef + jichu_xbian) / cos_t, y00]
    x4 = [x00 - fenxi / cos_t - pou2.right_l - (d_lef + jichu_xbian) / cos_t - w0 * tan_t / 2, y00 + w0 / 2]
    x5 = [x00 - fenxi / cos_t - pou2.right_l - d_lef / cos_t - w0 * tan_t / 2, y00 + w0 / 2]
    x6 = [x00 - fenxi / cos_t - pou2.right_l - w0 * tan_t / 2, y00 + w0 / 2]
    x7 = [x00 - fenxi / cos_t - pou2.right_l - w0 * tan_t / 2 + l0 / cos_t, y00 + w0 / 2]
    x8 = [x00 - fenxi / cos_t - pou2.right_l - w0 * tan_t / 2 + (l0 + d_lef) / cos_t, y00 + w0 / 2]
    x9 = [x00 - fenxi / cos_t - pou2.right_l - w0 * tan_t / 2 + (l0 + d_lef + jichu_xbian) / cos_t, y00 + w0 / 2]
    x10 = [x00 - fenxi / cos_t - pou2.right_l + (l0 + d_le + jichu_xbian) / cos_t, y00]
    x11 = [x00 - fenxi / cos_t - pou2.right_l + (l0 + d_lef) / cos_t, y00]
    x12 = [x00 - fenxi / cos_t - pou2.right_l + l0 / cos_t, y00]
    x13 = [x00 - fenxi / cos_t - pou2.right_l + l0 / cos_t - pou3.right_l - fenxi / cos_t, y00]
    x14 = [x00 - fenxi / cos_t - pou2.right_l + l0 / cos_t - pou3.right_l - fenxi / cos_t + (jin_bian + fenxi) * tan_t,
           y00 - jin_bian]
    x15 = [x00 + (jin_bian + fenxi) * tan_t, y00 - jin_bian]
    x16 = [x4[0] - 10, x4[1]]
    x17 = [x9[0] + 10, x9[1]]
    add_name(x15, x14, "1/2平面")

    xmid_low = [x1[0] + l0 / 2 / cos_t, x1[1]]
    xmid_low1 = [x6[0] + l0 / 2 / cos_t, x1[1]]
    xmid_up = [x6[0] + l0 / 2 / cos_t, x6[1]]
    text_pos = [(xmid_low[0] + xmid_up[0]) / 2, (xmid_low[1] + xmid_up[1]) / 2]

    x22 = [xmid_low[0], xmid_low[1] - jin_bian]
    line_biaozhu(x22, xmid_low, xmid_low, theta=0, bu_ping=True)
    line_biaozhu(xmid_up, xmid_low, xmid_low, theta=0, bu_ping=True)

    line_huizhi(xmid_up, [xmid_low1[0], xmid_low1[1] + (xmid_low[0] - xmid_low1[0])])
    p1 = vtpnt(xmid_up[0], xmid_up[1])
    p2 = vtpnt(xmid_low[0], xmid_low[1])
    p3 = vtpnt(xmid_low1[0], xmid_low1[1])
    p4 = vtpnt(text_pos[0], text_pos[1])
    if theta != 0:
        msp.AddDimAngular(p1, p2, p3, p4)
    add_mid_line(xmid_low, xmid_up)

    b1 = [x6[0], x6[1] + biao_ju]
    biaozhu1 = [x4, x5, x6, x7, x8, x9]
    for i in range(len(biaozhu1) - 1):
        line_biaozhu(biaozhu1[i], biaozhu1[i + 1], b1, theta)

    xiangti = poly_huizhi([x4, x3, x2, x5, x2, x1, x6, x1, x0, x15, x14, x13, x0, x12, x7, x12, x11, x8, x11, x10, x9])
    line_hidden(line_huizhi(x16, x17), scale=1.5)

    def yiqiang(left):
        text1, text2 = "Ⅱ", "Ⅳ"
        global pou2, pou4
        tan_a = tan(pi * (-afa) / 180)
        if theta > 0:
            if not left:
                tan_a = tan(pi * (theta) / 180)
            else:
                tan_a = tan(pi * (-afa) / 180)
        if theta < 0:
            if not left:
                tan_a = tan(pi * (afa) / 180)
            else:
                tan_a = tan(pi * (-theta) / 180)
        if not left:
            pou2, pou4 = pou3_1, pou5_1
            text1, text2 = "Ⅲ", "v"
        z1 = [x00 + fenxi * (tan_t - cos_t), y00 - fenxi]
        z2 = [z1[0] + yanshen * tan_t, z1[1] - yanshen]
        z7 = [z2[0] - pou2.right_l - pou2.dingkuan - pou2.di_chang - pou2.left_l, z2[1]]
        z8 = [z1[0] - pou2.right_l - pou2.dingkuan - pou2.di_chang - pou2.left_l, y00 - fenxi]
        z9 = [z1[0] - pou2.right_l - pou2.dingkuan - pou2.di_chang, y00 - fenxi]
        z20 = [z2[0] - pou2.right_l - pou2.dingkuan - pou2.di_chang, z2[1]]
        z10 = [z1[0] - pou2.right_l - pou2.dingkuan, y00 - fenxi]
        z21 = [z2[0] - pou2.right_l - pou2.dingkuan, z2[1]]
        z11 = [z1[0] - pou2.right_l, y00 - fenxi]
        z12 = [z1[0] - pou2.right_l + pou2.d_bian, y00 - fenxi]
        z22 = [z2[0] - pou2.right_l, z2[1]]
        z13 = [z2[0] - pou2.right_l + pou2.d_bian, z2[1]]
        Y2 = [x00 + fenxi * (tan_t - cos_t) - yanshen * tan_t, y00 - fenxi - yanshen]
        Y13 = [x00 + fenxi * (tan_t - cos_t) - yanshen * tan_t - pou2.right_l + pou2.d_bian, y00 - fenxi - yanshen]
        Y22 = [x00 + fenxi * (tan_t - cos_t) - yanshen * tan_t - pou2.right_l, y00 - fenxi - yanshen]
        Y21 = [x00 + fenxi * (tan_t - cos_t) - pou2.right_l - pou2.dingkuan - yanshen * tan_t, y00 - fenxi - yanshen]
        Y20 = [x00 + fenxi * (tan_t - cos_t) - pou2.right_l - pou2.dingkuan - pou2.di_chang - yanshen * tan_t,
               y00 - fenxi - yanshen]
        Y7 = [x00 + fenxi * (
                tan_t - cos_t) - pou2.right_l - pou2.dingkuan - pou2.di_chang - pou2.left_l - yanshen * tan_t,
              y00 - fenxi - yanshen]

        z3 = [z2[0] - ((pou2.hyi - pou4.hyi) * 1.5 + pou4.right_l + pou4.dingkuan) * tan_a,
              z2[1] - ((pou2.hyi - pou4.hyi) * 1.5 + pou4.right_l + pou4.dingkuan)]

        z24 = [z22[0] - (pou4.dingkuan + (pou2.hyi - pou4.hyi) * 1.5) * tan_a,
               z22[1] - (pou4.dingkuan + (pou2.hyi - pou4.hyi) * 1.5)]
        z26 = [z24[0] - pou4.hyi, z24[1]]
        z17 = [z26[0], z26[1] + pou4.dingkuan]
        z16 = [z17[0] - pou4.d_bian, z17[1]]
        z23 = [z13[0] - (pou2.hyi - pou4.hyi) * 1.5 * tan_a, z16[1]]

        z15 = [z26[0] - pou4.d_bian, z26[1] - pou4.d_bian]
        z14 = [z23[0] - (pou4.dingkuan + pou4.d_bian) * tan_a, z15[1]]
        z4 = [z26[0] - pou4.right_l, z26[1] - pou4.right_l]
        z5 = [z4[0], z4[1] + pou4.right_l + pou4.left_l + pou4.dingkuan + pou4.hyi * pou4.slope]
        z18 = [z17[0], z17[1] + pou4.hyi * pou4.slope]
        z19 = [z21[0] - ((pou2.hyi - pou4.hyi) * 1.5 - pou4.hyi * pou4.slope) * tan_a - 1.264 * pou4.hyi * pou4.slope,
               z18[1]]
        z25 = [z23[0] - pou2.dingkuan - pou2.d_bian, z23[1]]
        z6 = [(z5[1] - z7[1]) * (z19[0] - z20[0]) / (z19[1] - z20[1]) + z7[0], z5[1]]
        zk1 = [z2[0] - (abs(z23[1] - z2[1]) - zhi_w / 2) * tan_a, z23[1] + zhi_w / 2]
        zk2 = [z2[0] - (abs(z23[1] - z2[1]) + zhi_w / 2) * tan_a, z23[1] - zhi_w / 2]
        zk1_2 = [(zk1[0] + zk2[0]) / 2, (zk1[1] + zk2[1]) / 2]
        zk3 = [zk1_2[0] + zhi_length, zk1[1]]
        zk4 = [zk1_2[0] + zhi_length, zk2[1]]
        zk5 = [zk4[0], zk4[1] - (kang_w - zhi_w) / 2]
        zk6 = [zk5[0] + kang_k, zk5[1]]
        zk7 = [zk6[0], zk6[1] + kang_w]
        zk8 = [zk7[0] - kang_k, zk7[1]]
        zkb1 = line_biaozhu(zk1_2, zk3, [zk3[0], zk3[1] + biao_ju], theta=0, bu_ping=False)
        zkb2 = line_biaozhu(zk8, zk7, [zk7[0], zk7[1] + biao_ju], theta=0, bu_ping=False)
        zkb3 = line_biaozhu(zk6, zk7, [zk7[0] + biao_ju, zk7[1]], theta=0, bu_ping=True)

        b2 = [z8[0], z8[1] + biao_ju]
        b2_1 = [z8[0], z8[1] + 2.3 * biao_ju]
        bz1 = line_biaozhu(z8, z9, b2_1, theta)
        bz2 = line_biaozhu(z9, z10, b2, theta)
        bz3 = line_biaozhu(z10, z11, b2_1, theta)
        bz4 = line_biaozhu(z11, z1, b2, theta)

        b3 = [z26[0], z4[1] - 2 * biao_ju]
        bz5 = line_biaozhu(z4, z26, b3)
        bz6 = line_biaozhu(z26, z24, b3)
        bz7 = line_biaozhu(z24, z22, b3)

        b4 = [z4[0] - 2 * biao_ju, z4[1]]
        bz8 = line_biaozhu(z4, z26, b4, theta=0, bu_ping=True)
        bz9 = line_biaozhu(z17, z18, b4, theta=0, bu_ping=True)
        bz10 = line_biaozhu(z26, z17, b4, theta=0, bu_ping=True)
        bz11 = line_biaozhu(z18, z5, b4, theta=0, bu_ping=True)

        b5 = [z3[0], z3[1] - 4 * biao_ju]
        bz12 = line_biaozhu(z3, z2, b5, theta=0, bu_ping=False)

        b6 = [z1[0], z1[1]]
        bz13 = line_biaozhu(z1, z2, b6, theta=0, bu_ping=True)
        bz14 = line_biaozhu(z2, z23, b6, theta=0, bu_ping=True)
        bz15 = line_biaozhu(z23, z3, b6, theta=0, bu_ping=True)

        bz16 = line_biaozhu(z7, z6, z6, theta=0, bu_ping=False)
        bz17 = line_biaozhu(z19, z20, z19, theta=0, bu_ping=False)

        jiemian1 = add_jiemian_num(z7, z2, text1, 1)
        jiemian2 = add_jiemian_num([z18[0], z6[1]], [z17[0], z4[1] - 1.5 * biao_ju], text2, 0)

        biaozhu_delete = [bz1, bz2, bz3, bz4, bz5, bz6, bz7, bz8, bz9, bz10, bz11, bz12, bz13, bz14, bz15, bz16, bz17,
                          zkb1, zkb2, zkb3] + jiemian1 + jiemian2

        zuoyi = poly_huizhi([z7, z8, z9, z10, z9, z20, z21, z13, z21, z10, z11, z22, z11, z12, z12, z13, z12, z1, z2])

        zuoyi1 = poly_huizhi(
            [z13, z23, z14, z24, z22, z24, z26, z15, z14, z15, z16, z23, z17, z26, z18, z19, z20, z19, z25, z21])
        zuoyi2 = poly_huizhi([z2, zk2, zk1, zk3, zk4, zk5, zk6, zk7, zk8, zk4, zk2, z3, z4, z5, z6, z7])
        if not left:
            m1 = [x1[0] + l0 / 2 / cos_t + fenxi * tan_t, z2[1]]
            m2 = [x1[0] + l0 / 2 / cos_t + fenxi * tan_t, z2[1] + 5]
            m3 = [x1[0] + l0 / 2 / cos_t + (yanshen + fenxi) * tan_t, z2[1]]
            m4 = [x1[0] + l0 / 2 / cos_t + (yanshen + fenxi) * tan_t, z2[1] + 5]
            zuoyi11 = poly_huizhi(
                [Y7, z8, z9, z10, z9, Y20, Y21, Y13, Y21, z10, z11, Y22, z11, z12, z12, Y13, z12, z1, Y2])
            mirror_zuoyou([zuoyi11, bz1, bz2, bz3, bz4], m1, m2)
            mirror_zuoyou([zuoyi1, zuoyi2, bz5, bz6, bz7, bz8, bz9, bz10, bz11,
                           bz12, bz13, bz14, bz15, bz16, bz17] + jiemian1 + jiemian2, m3, m4)
            for biaozhu in biaozhu_delete:
                try:
                    biaozhu.Delete()
                except:
                    pass

            zuoyi.Delete()
            zuoyi11.Delete()
            zuoyi1.Delete()
            zuoyi2.Delete()

        length = abs(z1[0] - z4[0])
        s_qipa = cal_s_triangle(z2, z6, z7) + cal_s_triangle(z2, z3, z6) + abs(z1[0] - z8[0]) * abs(z2[1] - z1[1]) * \
                 scale / 100 * scale / 100 + cal_s_triangle(z3, z4, z6) + cal_s_triangle(z4, z5, z6)
        # 开始空间点定义
        zkj10 = [z10[0], z10[1], pou2.hyi + pou2.jidi_h]
        zkj9 = [z9[0], z9[1], pou2.jidi_h]
        zkj20 = [z20[0], z20[1], pou2.jidi_h]
        zkj21 = [z21[0], z21[1], pou2.hyi + pou2.jidi_h]
        zkj19 = [z19[0], z19[1], pou4.jidi_h]
        zkj18 = [z18[0], z18[1], pou4.jidi_h]
        zkj17 = [z17[0], z17[1], pou4.hyi + pou4.jidi_h]
        zkj25 = [z25[0], z25[1], pou4.hyi + pou4.jidi_h]

        zkj22 = [z22[0], z22[1], pou2.hyi + pou2.jidi_h]
        zkj24 = [z24[0], z24[1], pou4.hyi + pou4.jidi_h]
        zkj22_1 = [z22[0], z22[1], pou2.jidi_h]
        zkj24_1 = [z24[0], z24[1], pou4.jidi_h]

        s_humian_1 = cal_kjs_triangle(zkj10, zkj9, zkj20) + cal_kjs_triangle(zkj20, zkj10, zkj21) + \
                     cal_kjs_triangle(zkj20, zkj21, zkj19) + cal_kjs_triangle(zkj19, zkj21, zkj25) + \
                     cal_kjs_triangle(zkj19, zkj25, zkj17) + cal_kjs_triangle(zkj19, zkj17, zkj18) + \
                     cal_kjs_triangle(zkj22, zkj22_1, zkj24) + cal_kjs_triangle(zkj24, zkj22_1, zkj24_1)
        s1 = (pou2.dingkuan + pou2.dingkuan + pou2.slope * pou2.hyi) / cos_t * pou2.hyi / 2 * scale * scale / 10000
        s2 = pou2.hyi * 1.5 * scale / 100
        s5 = pou4.hyi * pou4.hyi * scale * scale / 10000
        s6 = (pou4.dingkuan + pou4.dingkuan + pou4.slope * pou4.hyi) / cos_t * pou4.hyi / 2 * scale * scale / 10000
        s_humian = s_humian_1 + s1 + s2 + s5 + s6
        # 结束空间点定义

        return length, s_qipa, s_humian

    length1, s_jichu_left, s_humian_left = yiqiang(left=True)
    length2, s_jichu_right, s_humian_right = yiqiang(left=False)
    s_jichu = s_jichu_left + s_jichu_right
    s_humian = s_humian_left + s_humian_right
    tuchang = abs(x13[0] - x0[0]) + length1 + length2
    return tuchang, s_jichu, s_humian


def cal_v_zhuiti(d, h):
    if h > 8:
        n = 1.25
        m = 1.75
    else:
        n = 1.0
        m = 1.5
    a = 0
    b = 0
    A1 = h * n
    B1 = h * m
    a1 = A1 - d
    b1 = B1 - d
    A = a + d
    B = b + d
    v1 = 1 / 24 * pi * h * (2 * (A * B + A1 * B1) + A * B1 + A1 * B)
    v2 = 1 / 24 * pi * h * (2 * (a * b + a1 * b1) + a * b1 + a1 * b)
    L = sqrt(4 * A1 * B1 * pi ** 2 + 15 * (A1 - B1) ** 2) * (
            1 + (4 / sqrt(15) - 1) * ((A1 - B1) / A1) ** 9) / 4
    v_puqi = v1 - v2
    return v_puqi, v2, L


def huizhi_shuliangbiao(ping_l, x00, y00):
    dictObj = doc.Dictionaries.Item("acad_tablestyle")
    keyName = "fengltableStyle"
    className = "AcDbTableStyle"
    customObj = dictObj.AddObject(keyName, className)
    customObj.Name = "fengltableStyle"
    customObj.Description = "New Style for My Tables"
    # customObj.FlowDirection = "acTableBottomToTop"
    customObj.HorzCellMargin = 0.22
    customObj.BitFlags = 1
    customObj.SetAlignment(1, 5)
    customObj.SetAlignment(2, 5)
    customObj.SetAlignment(3, 5)
    customObj.SetTextHeight(1, 2.5)
    customObj.SetTextHeight(2, 2.5)
    customObj.SetTextHeight(4, 2.5)
    customObj.SetTextStyle(1, "SMFS")
    customObj.SetTextStyle(2, "SMFS")
    customObj.SetTextStyle(4, "SMFS")

    doc.ActiveLayer = Layer_biaoge
    Layer_biaoge.color = 7
    InsertionPoint = vtpnt(x00, y00)
    NumRows = 62
    NumColumns = 5
    RowHeight = 1

    ColWidth = 20
    table = msp.AddTable(InsertionPoint, NumRows, NumColumns, RowHeight, ColWidth)
    table.StyleName = "fengltableStyle"
    table.Height = 250
    table.SetColumnWidth(2, 40)
    table.UnmergeCells(0, 0, 0, 4)

    bridge_excel = xlrd.open_workbook('表格合并.xls').sheet_by_name("Sheet1")
    rows = bridge_excel.nrows
    hang1 = bridge_excel.col_values(0, 1, rows + 1)
    hang2 = bridge_excel.col_values(1, 1, rows + 1)
    lie1 = bridge_excel.col_values(2, 1, rows + 1)
    lie2 = bridge_excel.col_values(3, 1, rows + 1)
    bridge_excel2 = xlrd.open_workbook('表格合并.xls').sheet_by_name("Sheet2")
    table_x = bridge_excel2.col_values(0, 1, bridge_excel2.nrows + 1)
    table_y = bridge_excel2.col_values(1, 1, bridge_excel2.nrows + 1)
    table_text = bridge_excel2.col_values(2, 1, bridge_excel2.nrows + 1)

    for i in range(len(hang1)):
        try:
            table.MergeCells(hang1[i], hang2[i], lie1[i], lie2[i])
        except:
            pass
    for j in range(len(table_x)):
        table.SetText(table_x[j], table_y[j], str(table_text[j]))

    v_jichu = round(ping_l[1] * 2 * jidi_h / 100, 2)
    s_sui_rock = round(ping_l[1] * 2 * 0.1, 2)
    g_humian = round(ping_l[2] * 2 * 11.84 + 9.8 * 2 * (pou4.hyi + pou5.hyi) * scale / 100, 2)

    v_jiangqs = 2 * round(cal_v_zhuiti(0.35, pou4.hyi * scale / 100)[0] + cal_v_zhuiti(0.35, pou5.hyi * scale / 100)[0],
                          2)
    v_suishi = 2 * round(cal_v_zhuiti(0.1, pou4.hyi * scale / 100)[0] + cal_v_zhuiti(0.1, pou5.hyi * scale / 100)[0], 2)
    v_hangtian = ((2 * round(
        cal_v_zhuiti(0.35, pou4.hyi * scale / 100)[1] + cal_v_zhuiti(0.35, pou5.hyi * scale / 100)[1],
        2)) // 10 + 1) * 10
    v_waji = ((0.5 * 1 * 2 * round(
        cal_v_zhuiti(0.35, pou4.hyi * scale / 100)[2] + cal_v_zhuiti(0.35, pou5.hyi * scale / 100)[2],
        2)) // 10 + 1) * 10

    table.SetText(36, 4, str(s_sui_rock))
    table.SetText(35, 4, str(v_jichu))
    table.SetText(34, 4, str(g_humian))
    table.SetText(48, 4, str(v_jiangqs))
    table.SetText(49, 4, str(v_suishi))
    table.SetText(50, 4, str(v_hangtian))
    table.SetText(51, 4, str(v_waji))
    table.Update
    # try:
    #     doc.SelectionSets.Item("SS1").Delete()
    # except:
    #     print("Delete selection failed")
    # slt = doc.SelectionSets.Add("SS1")
    # filterType = [0]  # 定义过滤类型
    # filterData = ["LWPOLYLINE"]  # 设置过滤参数
    # filterType = vtint(filterType)  # 数据类型转化
    # filterData = vtvariant(filterData)  # 数据类型转化
    # slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
    result = vtvariant([])
    doc.ActiveLayer = Layer_jiegou
    x1 = [x00, y00]
    x2 = [x00, y00 - 250]
    x3 = [x00 + 120, y00 - 250]
    x4 = [x00 + 120, y00]

    x5 = [x1[0], x1[1] + 21]
    x6 = [x4[0], x4[1] + 21]
    t_name = add_name(x5, x6, "主 要 工 程 数 量 表")
    t_name.Height = 5
    poly1 = poly_huizhi([x1, x2, x3, x4, x1])
    result = poly1.Explode
    charu_fuzhu(x4[0] + 20, x4[1])


# slope hyi left_l right_l jidi_h dingkuan d_bian ding_gao di_chang theta
def hui_pou(x00, y00, pob, text):
    y1 = [x00, y00]
    y2 = [x00, y00 + pob.jidi_h]
    y3 = [x00 + pob.left_l + pob.right_l + pob.di_chang + pob.dingkuan, y00]
    y4 = [y3[0], y2[1]]
    y5 = [y4[0] - pob.right_l, y4[1]]
    y6 = [y5[0], y5[1] + pob.hyi]
    y7 = [y6[0] + pob.d_bian, y6[1]]
    y8 = [y7[0], y7[1] + pob.ding_gao - pob.d_bian]
    y9 = [y6[0], y6[1] + pob.ding_gao]
    y10 = [y9[0] - pob.dingkuan, y9[1]]
    y11 = [y10[0], y10[1] - pob.d_bian]
    y12 = [y10[0], y10[1] - pob.ding_gao]
    y13 = [y2[0] + pob.left_l, y2[1]]
    y18 = vtpnt((y12[0] + y13[0]) / 2 - 0.1 * biao_ju, (y12[1] + y13[1]) / 2)
    poly_huizhi([y1, y2, y1, y3, y4, y5, y6, y7, y8, y11, y8, y9, y10, y11, y12, y6, y12, y13, y2, y5])

    b1 = [y2[0], y2[1] + pob.ding_gao]
    line_biaozhu(y2, y13, b1, pob.theta, False)
    line_biaozhu(y13, y5, b1, pob.theta, False)
    line_biaozhu(y5, y4, b1, pob.theta, False)

    line_biaozhu(y10, y9, [y9[0], y9[1] + pob.ding_gao], pob.theta, False)
    line_biaozhu(y9, y8, y8, pob.theta, False)

    b2 = [y3[0] + pob.dingkuan, y3[1]]
    line_biaozhu(y3, y4, b2, 0, True)
    line_biaozhu(y4, y7, b2, 0, True)
    line_biaozhu(y7, y9, b2, 0, True)

    y14 = [(y1[0] + y3[0]) / 2 - 14, y1[1] - 14]
    y15 = [(y1[0] + y3[0]) / 2 + 14, y1[1] - 14]
    y_text = vtpnt(y14[0] + 2, y14[1] + 2)
    doc.ActiveLayer = Layer_biaozhu
    textobj1 = msp.AddText(text, y_text, 4.5)
    textobj1.StyleName = "SMFS"
    textobj1.ScaleFactor = 0.8
    text_pou = "1:" + str(pou_slope)
    textobj2 = msp.AddText(text_pou, y18, 2.5)
    textobj2.StyleName = "SMFS"
    textobj2.ScaleFactor = 0.8
    textobj2.Rotation = 1.18682

    y16 = [(y1[0] + y3[0]) / 2 - 14 + 0.575, y1[1] - 14.575]
    y17 = [(y1[0] + y3[0]) / 2 + 14 - 0.575, y1[1] - 14.575]
    line_huizhi(y16, y17)

    doc.ActiveLayer = Layer_jiegou
    poly1 = poly_huizhi([y14, y15])
    tuchang = 1.5 * abs(y3[0] - y1[0])
    return tuchang


def charu_fuzhu(x, y):
    doc.ActiveLayer = Layer_elevation
    InsertionPoint = vtpnt(x, y)
    if theta == 0:
        jiajiao = "正交"
    else:
        jiajiao = "斜交%d°" % int(abs(theta))

    Text1 = "附注：\n1、本图尺寸除里程高程以米计及注明者外，余均以厘米计。\n2、本桥为立交兼排洪而设，位于直线上，箱形桥按" + jiajiao + "设置，结构高度%sm设计,出入口均设八字翼墙。\n" % str(
        float(h0) * scale / 100)
    text2 = "3、本桥基底位于粉土层中，其承载力不足，基底应采用1:2砂石换填至砂砾土层中（σ=350kPa），1：2砂石" \
            "换填需分层夯实，分层厚度0.25m，每层至少夯打三遍，夯打前适当加水，夯打需达到中密，即相对密度D≥0.67，数量已计列。\n4、为防止翼墙滑" \
            "移，每处翼墙基础设支撑梁及抗滑桩，如图所示，数量已计列。\n5、桥址处地下水位为8.8m。无侵蚀性，本桥箱身采用C35混凝土，翼墙、基础采用" \
            "C30混凝土。采用现浇法施工。\n6、本桥挡砟墙、人行道、桥面防排水及沉降缝、护面钢筋等具体构造和钢筋布置详见相关图纸。\n7、箱形桥箱身" \
            "两侧采用涂聚氨酯防水涂料处理，箱顶设2%横坡，箱顶及防水层均为水平，最上层由C40纤维混凝土调坡，图中未示，数量已计列。\n8、箱身背面基" \
            "坑采用C25混凝土回填至原地面，数量已计列。\n9、在路基两侧边坡上设M10浆砌片石检查台阶，参照《桥梁综合参考图》施工，数量已计列。\n" \
            "10、本桥铺砌防磨层填高后，净空不足5.0m，在出入口道路上各设限高架1处。\n11、施工注意事项：\n(1)施工单位在施工前应根据施工图对路基" \
            "宽度、标高、边坡率、线间距、结构尺寸进行校核，如发现与设计不符，应及时通知设计单位研究解决。\n(2)施工前应探明地下及地上管线位置和产" \
            "权单位，并和产权单位签订有关协议进行迁改和做好保护措施，严禁盲目施工，必要时进行物理探测。\n(3)基坑开挖完成后，应核实地基持力层的承" \
            "载力。特别是当出现软硬不均时，应及时通知相关单位作处理。\n(4)箱形桥施工应严格按照相关规程、规范及标准图要求，施工用水及混凝土粗细骨" \
            "料均按要求检验，不得采用侵蚀性水及有侵蚀的骨料。混凝土材料应满足相关规范和规定的各项指标要求，施工工艺严格按相关规范和规定的要求办理" \
            "，使结构满足耐久性使用要求。\n(5)混凝土碱含量应符合《铁路混凝土工程预防碱-骨料反应技术条件》(TB/T3054)的规定。\n(6)箱形桥两侧路" \
            "堤填碎石土必须对称均匀填筑，并用小型平板振动机压实，禁止单侧填级配碎石造成偏压。\n(7)施工时采取可行的排水措施，加强地表及基坑排水，" \
            "工作坑开挖后，及时浇筑基础，严禁地表水浸泡基坑而影响地基承载力，箱形桥施工完毕后，应及时用不透水土层夯实回填高出原地面，以利排水。\n" \
            "(8)施工过程要注意环境保护，对破坏的植被应予恢复。\n12、本箱型桥上下游需进行顺沟，以保证上下游排水顺畅，数量已计列。\n13、桥上人行" \
            "道见本册相关图纸。\n14、本桥上下游两侧采用堤坝铺砌，上游长15m，下游长20m，共70m，铺砌末端挖沟与原沟顺接。\n15、箱内回填20cm厚混" \
            "凝土防磨层，施工请注意。\n16、本图未尽事宜按有关规范规定办理。\n"
    Text = Text1 + text2
    Width = 200
    Mtext = msp.AddMText(InsertionPoint, Width, Text)
    Mtext.Height = 4.5
    Mtext.LineSpacingFactor = 0.853
    Mtext.StyleName = "SMFS"

    doc.ActiveLayer = Layer_jiegou


# bridge_excel = xlrd.open_workbook('表格合并.xls').sheet_by_name("Sheet1")
# rows = bridge_excel.nrows
# hang1 = bridge_excel.col_values(0, 1, rows + 1)
# hang2 = bridge_excel.col_values(1, 1, rows + 1)
# lie1 = bridge_excel.col_values(2, 1, rows + 1)
# lie2 = bridge_excel.col_values(3, 1, rows + 1)

data = xlrd.open_workbook('表格合并.xls').sheet_by_name("箱形桥输入数据")
data_input = data.col_values(2, 1, data.nrows + 1)
elevation0, x_slope = [float(i) for i in data_input[0:2]]
theta = int(data_input[4])
afa = int(data_input[28])
pou_slope = data_input[27]
d3, x_guilu = [float(100 * i) for i in data_input[2:4]]
w0, l0, d1, d2, l_lei_up, h_lei_up, l_lei_down, h_lei_down, h0, d_le, thk_pu, d4, jichu_xbian = [float(100 * i) for i in
                                                                                                 data_input[5:18]]
jidi_h, pou2_h, pou4_h, pou3_h, pou5_h, pou_dingkuan, pou_jinbian, pou_left_l, pou_right_l = \
    [float(100 * i) for i in data_input[18:27]]


licheng_qianzhui = str(data.col_values(4, 1, data.nrows + 1)[0])
licheng = str(data.col_values(5, 1, data.nrows + 1)[0])
licheng_fx = int(data.col_values(4, 1, data.nrows + 1)[1])
licheng_name_l = str(data.col_values(4, 1, data.nrows + 1)[2])
licheng_name_r = str(data.col_values(4, 1, data.nrows + 1)[3])



scale = 10

l0, h0, d_le, d1, d2, d3, w0, fenxi, h_lei_up, l_lei_up, h_lei_down, l_lei_down, jichu_xbian, x_guilu, pou_dinggao = \
    [i / scale for i in [l0, h0, d_le, d1, d2, d3, w0, 3, h_lei_up, l_lei_up, h_lei_down, l_lei_down, jichu_xbian,
                         x_guilu, 25]]
yanshen = float(data.col_values(4, 1, data.nrows + 1)[6]) * 100 / scale

if pou2_h == 0 and pou3_h == 0:
    y_gui1 = d3 + (l0 * x_slope / 2 / cos(pi * theta / 180))
    y_gui2 = d3 - (l0 * x_slope / 2 / cos(pi * theta / 180))
    pou2_h = 100 / scale * (y_gui1 + h0 + d1 - x_guilu - pou_dinggao + 0.05)
    pou3_h = 100 / scale * (y_gui2 + h0 + d1 - x_guilu - pou_dinggao + 0.05)

pou2 = pou(pou2_h, scale, theta)
pou4 = pou(pou4_h, scale)
pou3 = pou(pou3_h, scale, theta)
pou5 = pou(pou5_h, scale)
pou2.jidi_h = jidi_h / scale
pou3.jidi_h = jidi_h / scale
pou4.jidi_h = jidi_h / scale
pou5.jidi_h = jidi_h / scale

pou3_1 = pou(100, 2, 0)
pou5_1 = pou(100, 2, 0)
pou3_1, pou5_1 = pou3, pou5

right = True
xiang_jichu_h = 50 / scale
xiang_jichu_l = 20 / scale
jin_bian = 100 / scale  # 箱形桥的襟边默认为100

kang_w = 140 / scale  # 抗滑桩在侧面图中的宽度
kang_l = 300 / scale  # 抗滑桩在侧面图中的深度
kang_k = 100 / scale
zhi_w = 80 / scale  # 支撑梁的宽度为80
zhi_length = 250 / scale
zhi_l = 100 / scale  # 支撑梁的高度为100
zhuiti1 = 42 / scale  # 椎体堆砌的上部分高
zhuiti2 = 12 / scale  # 椎体堆砌的下部分高
zhuiti_shen = 100 / scale
biao_ju = 30 / scale  # 标注线的距离


# l0 = l0 / 2


def main():
    xiang_l = hui_xiang(980 / scale, 1720 / scale, l0, h0, d_le, d1, d2, d3, theta)  # 前两个是定位点，后面是
    print("1---------------立面图箱体绘制完成！")
    yiliang_l = limian_yiqiang(980 / scale, 1720 / scale, l0, d_le, d2, theta, afa, right)
    limian_yiqiang(980 / scale, 1720 / scale, l0, d_le, d2, theta, afa, False)
    print("2---------------立面图翼墙绘制完成！")
    cel_l = cemian_huizhi(1400 / scale, 500 / scale, w0, h0, h_lei_up, h_lei_down, d1, d2, d3, theta)
    print("3---------------侧视图绘制完成！")
    ping_l = pingmian_huizhi(cel_l[0] + 1550 / scale, 1040 / scale, l0, w0, d_le, theta, afa)
    print("4---------------平面图绘制完成！")
    huizhi_shuliangbiao(ping_l, cel_l[0] + 1600 / scale + ping_l[0], 265)
    print("5---------------数量表计算并绘制完成！")

    t2 = hui_pou(980 / scale + xiang_l + 1.99 * yiliang_l, 1720 / scale, pou3, "Ⅱ--Ⅱ 剖 面")
    t3 = hui_pou(980 / scale + xiang_l + 1.99 * yiliang_l + t2, 1720 / scale, pou2, "Ⅲ--Ⅲ 剖 面")
    t4 = hui_pou(980 / scale + xiang_l + 1.99 * yiliang_l + t2 + t3, 1720 / scale, pou4, "Ⅳ--Ⅳ 剖 面")
    t5 = hui_pou(980 / scale + xiang_l + 1.99 * yiliang_l + t2 + t3 + t4, 1720 / scale, pou5, "Ⅴ--Ⅴ 剖 面")
    print("6---------------翼墙剖面绘制完成！")
    hui_tukuang(24.7 + cel_l[0] + 1600 / scale + ping_l[0] + 320 + 20)
    print("7---------------图长计算完成！")

    wincad.ZoomExtents

    MessageBox(0, "绘图成功！谢谢使用！", "提醒", MB_OK)


main()

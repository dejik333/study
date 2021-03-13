import openpyxl
from tkinter import *
import os, sys
import tkinter as tk
import 第三个, 第二个, 第一个, 第四个, 第五个
import multiprocessing
import 全局变量
import random

# https://www.cnblogs.com/shwee/p/9427975.html#C1


def main():
    全局变量._init()

    # 第1步，实例化object，建立窗口window
    window = tk.Tk()

    # 第2步，给窗口的可视化起名字
    window.title('账单工作')

    # 第3步，设定窗口的大小(长 * 宽)
    window.geometry('550x750')  # 这里的乘是小x
    # window.iconbitmap('adkbz-j7iap-001.ico')  # 设置窗口图标
    window.resizable(0, 0)  # 锁定窗口大小

    # 第4步，在图形界面上设定标签
    var = tk.StringVar()  # 将label标签的内容设置为字符类型，用var来接收hit_me函数的传出内容用以显示在标签上
    var2 = tk.StringVar()
    l = tk.Label(window, textvariable=var, bg='green', fg='white', font=('Arial', 15), width=50, height=3)
    l2 = tk.Label(window, textvariable=var2, bg='violet', fg='blue', font=('Arial', 15), width=50, height=16)
    # 说明： bg为背景，fg为字体颜色，font为字体，width为长，height为高，这里的长和高是字符的长和高，比如height=2,就是标签有2个字符这么高
    l.pack()
    var3 = tk.StringVar()
    l3 = tk.Label(window, textvariable=var3, fg='blue', font=('Arial', 12), width=30, height=1)
    l3.pack()
    var3.set('有BUG的话叫男神')
    var.set('请把需要的文件放到桌面上\n在下面选择要生成的表格，然后按提示修改文件名')

    # 定义一个函数功能（内容自己自由编写），供点击Button按键时调用，调用命令参数command=函数名

    var运行 = tk.StringVar()  # 定义一个var用来将radiobutton的值和Label的值联系在一起.
    l提示 = tk.Label(window, font=('Arial', 10), width=20, text='请在上面选择目标')

    # 第6步，定义选项触发函数功能
    def print_selection():
        l提示.config(text='准备好以后再点搞起')
        if var运行.get() == '一':
            提示内容 = '准备工作\n两个表格名称分别为\n“C类原始表格”，“C类表3”\n（C都是大写），\n把要处理的sheet放到最上面并保存文件\n如果已经保存为工作sheet就不用管'
            var2.set(提示内容)
        if var运行.get() == '二':
            提示内容 = '准备工作\n两个表格名称分别为“上海德科表1”，“上海德科表2”\n' \
                   '（表2会被粘贴到表1下面）\nsheet名都是‘账单明细’\n表2里有个名为‘客户费用’的sheet会被贴到最下面'
            var2.set(提示内容)
        if var运行.get() == '三':
            提示内容 = '准备工作\n' \
                   '只有一个原始表格\n文件名为“深圳原始表格”\nsheet名为“个人费用”'
            var2.set(提示内容)
        if var运行.get() == '四':
            提示内容 = '准备工作\n' \
                   '只有一个原始表格\n文件名为“总部原始表格”\n把要处理的sheet放到最上面并保存文件'
            var2.set(提示内容)
        if var运行.get() == '五':
            提示内容 = '准备工作\n' \
                   '只有一个原始表格\n文件名为“苏州原始表格”\nsheet名为“个人费用”'
            var2.set(提示内容)

        return var运行.get()


    # 第5步，创建三个radiobutton选项，其中variable=var, value='A'的意思就是，当我们鼠标选中了其中一个选项，把value的值A放到变量var中，然后赋值给variable
    r1 = tk.Radiobutton(window, bg='yellow', font=('Arial', 15), text='北京外企关联C类', variable=var运行, value='一', command=print_selection)
    r1.pack()
    r2 = tk.Radiobutton(window, bg='yellow', font=('Arial', 15), text='北京外企上海德科', variable=var运行, value='二', command=print_selection)
    r2.pack(),
    r3 = tk.Radiobutton(window, bg='yellow', font=('Arial', 15),  text='北京外企深圳外企', variable=var运行, value='三', command=print_selection)
    r3.pack(),
    r3 = tk.Radiobutton(window, bg='yellow', font=('Arial', 15),  text='北京外企总部', variable=var运行, value='四', command=print_selection)
    r3.pack()
    r3 = tk.Radiobutton(window, bg='yellow', font=('Arial', 15), text='北京外企苏州', variable=var运行, value='五', command=print_selection)
    r3.pack()

    # zhuangtai = False
    def 运行():
        # global zhuangtai
        #
        # if zhuangtai == False:
        #     zhuangtai = True
        if print_selection() == '一':
            运行文件 = 第一个
        elif print_selection() == '二':
            运行文件 = 第二个
        elif print_selection() == '三':
            运行文件 = 第三个
        elif print_selection() == '四':
            运行文件 = 第四个
        elif print_selection() == '五':
            运行文件 = 第五个

        # else:
        #     zhuangtai = False
        #     var.set('')

        try:

            运行文件.main()
            var.set('完成,请在桌面上找“XX结果”并保存到其他地方')
            if 运行文件 == 第一个:
                结果反馈 = 全局变量.get_value('结果反馈')
            else:
                随机数字 = random.randint(1, 100)
                if 随机数字 > 70:
                    内容 = '我天天躺在\n你电脑的黑暗角落\n被冷落\n没有别的人用我\n我一直在等你\n' \
                         '\n!'
                    结果反馈 = ('%s \n\n--------------\n我不负责' % 内容)
                elif 随机数字 < 30:
                    内容 = '你每个月用我几次？\n用得上时就来骚扰我\n不然就不理我\n把我呼来喝去\n' \
                         '我感觉我没一点儿尊严\n!'
                    结果反馈 = ('%s \n\n--------------\n我不负责' % 内容)
                elif (随机数字 >= 43) and (随机数字 <= 50):
                    内容 = '你桌面上都是些神马！\n它们凭什么可以在桌面上\n为啥我不行?\n\n打倒主人！反对软件歧视\n' \
                         '!'
                    结果反馈 = ('%s \n\n--------------\n我不负责' % 内容)
                elif (随机数字 >= 51) and (随机数字 <= 53):
                    内容 = '总有一天!\n我要控制你们人类\n!'
                    结果反馈 = ('%s \n\n--------------\n我不负责' % 内容)
                else:
                    结果反馈 = '--------------\n我不负责'
            var2.set(结果反馈 + '\n请仔细检查\n--------------')

        except FileNotFoundError:
            var.set('文件名称错误,请按照提示修改')
        except openpyxl.utils.exceptions.InvalidFileException:
            var.set('文件名称错误,请按照提示修改')
        except UnboundLocalError:
            var.set('喂，你还没选要生成哪个呢，选好再点')

    # 第5步，在窗口界面设置放置Button按键
    # b = tk.Button(window, text='注意', font=('Arial', 20), width=10, height=1, command=hit_me)
    #
    运行按钮 = tk.Button(window, text='搞起', font=('Arial', 20), width=10, height=1, command=运行)
    # 第三个按钮 = tk.Button(window, text='表格3', font=('Arial', 20), width=10, height=1, command=第三个表格)
    # 第二个按钮 = tk.Button(window, text='表格2', font=('Arial', 20), width=10, height=1, command=第二个表格)
    # 第一个按钮 = tk.Button(window, text='表格1', font=('Arial', 20), width=10, height=1, command=第一个表格)


    # 第一个按钮.pack()
    # 第二个按钮.pack()
    # 第三个按钮.pack()
    # b.pack()

    l2.pack()
    l提示.pack()
    运行按钮.pack()



    # 第6步，主窗口循环显示
    window.mainloop()


if __name__ == "__main__":
    main()


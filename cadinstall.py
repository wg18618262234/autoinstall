# coding=utf-8
# !/usr/bin/python
import time
import win32api
import win32gui
import zipfile
import os
from win32api import keybd_event
from win32gui import FindWindow, SetForegroundWindow, FindWindowEx, SendMessage, EnumChildWindows, GetWindowText, \
    GetClassName

import rarfile
import win32con


def ini():
    # 语言代码
    # https://msdn.microsoft.com/en-us/library/cc233982.aspx
    LID = {0x0804: "Chinese (Simplified) (People's Republic of China)",
           0x0409: 'English (United States)'}

    # 获取前景窗口句柄
    hwnd = win32gui.GetForegroundWindow()

    # 获取前景窗口标题
    title = win32gui.GetWindowText(hwnd)
    print('当前窗口：' + title)

    # 获取键盘布局列表
    im_list = win32api.GetKeyboardLayoutList()
    im_list = list(map(hex, im_list))
    print(im_list)

    # 设置键盘布局为英文
    result = win32api.SendMessage(
        hwnd,
        win32con.WM_INPUTLANGCHANGEREQUEST,
        0,
        0x0409)
    if result == 0:
        print('设置英文键盘成功！')


def open_app(app_dir):
    '''
    打开app
    :param app_dir:
    :return:
    '''

    if not app_dir:
        return

    os.startfile(app_dir)


def un_zip(file_name):
    """unzip zip file不好用"""
    zip_file = zipfile.ZipFile(file_name)
    if os.path.isdir(file_name + "_files"):
        pass
    else:
        os.mkdir(file_name + "_files")
    for names in zip_file.namelist():
        zip_file.extract(names, file_name + "_files/")
    zip_file.close()


def un_rar(file_name):
    """unrar zip file不好用"""
    rar = rarfile.RarFile(file_name)
    if os.path.isdir(file_name + "_files"):
        pass
    else:
        os.mkdir(file_name + "_files")
    os.chdir(file_name + "_files")
    rar.extractall()
    rar.close()


def is64windows():
    '''判断系统是否64位，是return ture,否return flase'''
    return 'PROGRAMFILES(X86)' in os.environ


def tar():
    '''
    AutoCAD 2010解压安装
    :return:
    '''
    lable = 'AutoCAD 2010'
    while True:
        # 根据类名及标题名查询句柄
        w = FindWindow(None, lable)
        if w > 0:
            break
    if w > 0:
        # 将软件窗口置于最前
        SetForegroundWindow(w)
        # 更改解压路径
        e = FindWindowEx(w, None, 'ComboBox', None)
        SendMessage(e, win32con.WM_SETTEXT, None, r'D:\Autodesk\AutoCAD_2010_Simplified_Chinese_MLD_WIN_64bit')
        # 开始解压
        b = FindWindowEx(w, None, 'Button', 'Install')
        SendMessage(b, win32con.BM_CLICK, None, None)
        print('%x,%x,%x' % (w, e, b))


def net_install():
    '''
    .net 3.5安装
    :return:
    '''
    while True:
        w = FindWindow(None, 'Windows 功能')
        if w > 0:
            break
    SetForegroundWindow(w)
    print(w)
    while True:
        b = FindWindowEx(w, None, 'Button', None)
        if b > 0:
            break
    print('%x,%x' % (w, b))


def show_window_attr(hWnd):
    '''
    显示窗口的属性
    :return:
    '''
    if not hWnd:
        return

    # 中文系统默认title是gb2312的编码
    title = GetWindowText(hWnd)
    clsname = GetClassName(hWnd)

    print('窗口句柄:%x ' % (hWnd))
    print('窗口标题:%s' % (title))
    print('窗口类名:%s' % (clsname))
    print()


def show_windows(hWndList):
    for h in hWndList:
        show_window_attr(h)


def demo_child_windows(parent):
    '''
    演示如何列出所有的子窗口
    :return:
    '''
    if not parent:
        return

    hWndChildList = []
    EnumChildWindows(parent, lambda hWnd, param: param.append(hWnd), hWndChildList)
    show_windows(hWndChildList)


def cad_install():
    '''
    cad安装
    :return:
    '''
    time.sleep(5)
    while True:
        w = FindWindow(None, 'AutoCAD 2010')
        if w > 0:
            break
        time.sleep(5)
    # 将软件窗口置于最前
    SetForegroundWindow(w)
    time.sleep(1)
    # 快捷键alt+i进入安装页面
    keybd_event(18, 0, 0, 0)  # Alt
    keybd_event(73, 0, 0, 0)  # i
    keybd_event(73, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
    # 点击下一步
    while True:
        w = FindWindow(None, 'AutoCAD 2010')
        if w > 0:
            break
        time.sleep(0.5)
    # 快捷键alt+n进入安装页面
    keybd_event(18, 0, 0, 0)  # Alt
    keybd_event(78, 0, 0, 0)  # n
    keybd_event(78, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(15)
    # 选中我接受
    while True:
        while True:
            w = FindWindow(None, 'AutoCAD 2010')
            if w > 0:
                break
            time.sleep(0.5)
        try:
            c = FindWindowEx(w, None, 'Button', '我接受(&A)')
            if c > 0:
                break
            time.sleep(0.5)
        except Exception as e:
            print(e)
    SendMessage(c, win32con.BM_CLICK, None, None)
    # 点击下一步
    while True:
        d = FindWindowEx(w, None, 'Button', '下一步(&N)')
        if d > 0:
            break
        time.sleep(0.5)
    SendMessage(d, win32con.BM_CLICK, None, None)
    # time.sleep(0.5)
    # 解压注册机   不好用，直接读取单文件
    # zip = r'D:\cad\AutoCAD.2010注册机.zip'
    # un_zip(zip)
    # 打开注册机
    # if is64windows():
    #     open_app('xf-a2010-64.exe')
    # else:
    #     open_app('xf-a2010-32.exe')
    # 填写产品用户信息
    while True:
        w = FindWindow(None, 'AutoCAD 2010')
        if w > 0:
            break
        time.sleep(0.5)
    sn1 = '356'
    sn2 = '72378422'
    key = '001B1'
    last_name = 'auto'
    first_name = 'installer'
    org = 'autoinstaller'
    e = FindWindowEx(w, None, 'Edit', None)
    SendMessage(e, win32con.WM_SETTEXT, None, sn1)
    keybd_event(9, 0, 0, 0)
    keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    f = FindWindowEx(w, e, 'Edit', None)
    SendMessage(f, win32con.WM_SETTEXT, None, sn2)
    keybd_event(9, 0, 0, 0)
    keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    g = FindWindowEx(w, f, 'Edit', None)
    SendMessage(g, win32con.WM_SETTEXT, None, key)
    keybd_event(9, 0, 0, 0)
    keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    h = FindWindowEx(w, g, 'Edit', None)
    SendMessage(h, win32con.WM_SETTEXT, None, last_name)
    keybd_event(9, 0, 0, 0)
    keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    i = FindWindowEx(w, h, 'Edit', None)
    SendMessage(i, win32con.WM_SETTEXT, None, first_name)
    keybd_event(9, 0, 0, 0)
    keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    j = FindWindowEx(w, i, 'Edit', None)
    SendMessage(j, win32con.WM_SETTEXT, None, org)
    # 点击下一步
    k = FindWindowEx(w, None, 'Button', '下一步(&N)')
    SendMessage(k, win32con.BM_CLICK, None, None)
    # 点击配置
    while True:
        w = FindWindow(None, 'AutoCAD 2010')
        if w > 0:
            break
        time.sleep(0.5)
    l = FindWindowEx(w, None, 'Button', '安装(&N)')
    SendMessage(l, win32con.BM_CLICK, None, None)

    # demo_child_windows(w)


if __name__ == '__main__':
    try:
        ini()
        # 解压安装
        # app = r'D:\cad\AutoCAD_2010__64位.exe'
        app = r'AutoCAD_2010__64位.exe'
        open_app(app)
        tar()
        #
        # net安装
        # net = r'dotNetFx35setup.exe'
        # open_app(net)
        # net_install()

        # cad安装
        cad_install()
    except Exception as e:
        import traceback
        import logging

        logging.basicConfig(level=logging.DEBUG,  # 控制台打印的日志级别
                            filename='new.log',
                            filemode='a',  ##模式，有w和a，w就是写模式，每次都会重新写日志，覆盖之前的日志
                            # a是追加模式，默认如果不写的话，就是追加模式
                            format=
                            '%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'
                            # 日志格式
                            )
        logging.error(f'Exception:{e}')
        traceback.print_exc()
        a = input('')

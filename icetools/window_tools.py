# -*- coding = utf-8 -*-
# @Time：2022/3/11 20:05
# @Author：kksk43
# @File：window_tools.py
# @Software：PyCharm
import math
import sys
import time

import numpy as np
import win32api
import win32con
import win32gui
import win32process
from PyQt5.QtWidgets import QApplication
import pyautogui

# 整个窗口的宽、高、句柄、app、screen
from win32clipboard import OpenClipboard, SetClipboardData, CloseClipboard

full_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
full_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
cap_full_hwnd = win32gui.GetDesktopWindow()
app = QApplication(sys.argv)
full_screen = QApplication.primaryScreen()


def fast_capture_full(x, y, w, h):
    """
    全屏快速截图

    :param x: 截图左上方横坐标（相对于屏幕的绝对坐标）
    :param y: 截图左上方纵坐标（相对于屏幕的绝对坐标）
    :param w: 截图宽度
    :param h: 截图盖度
    :return: 截得的图像，ndarray形式
    """
    img = full_screen.grabWindow(cap_full_hwnd, x, y, w, h).toImage()
    ptr = img.constBits()
    ptr.setsize(img.byteCount())
    return np.array(ptr).reshape(img.height(), img.width(), 4)[..., :-1][..., ::-1]

# def fast_capture_full(x, y, w, h):
#     return np.array(pyautogui.screenshot())[y:y+h, x:x+w]


class WindowTool:
    def __init__(self, class_name: str = None, window_name: str = None, init_hwnd=-1, x=-1, y=-1):
        """
        窗口工具，设计为操纵某个窗口行为的类，里面有很多方法都是针对指定窗口的常用操作

        初始化有4种方式（优先级由高到低依次介绍）：

        1、直接使用句柄初始化，指定目标窗口的句柄

        2、使用窗口的类名和窗口名进行初始化，通过类名和窗口名找到句柄

        3、指定屏幕上出现的窗口内的任一坐标，通过该坐标可以确定该窗口的句柄

        4、不指定任何参数，默认句柄是整个屏幕

        :param class_name: 指定窗口的类名
        :param window_name: 指定窗口的名字
        :param init_hwnd: 指定窗口的句柄
        :param x: 窗口左上方在当前屏幕的横坐标（绝对坐标）
        :param y: 窗口左上方在当前屏幕的纵坐标（绝对坐标）
        """
        if init_hwnd > 0:
            self.init_hwnd = init_hwnd
        elif window_name is not None:
            self.init_hwnd = win32gui.FindWindow(class_name, window_name)
            self.cap_hwnd = self.init_hwnd
            self.opt_hwnd = self.init_hwnd
        elif x >= 0 and y >= 0:
            self.init_hwnd = win32gui.WindowFromPoint((x, y))
            self.cap_hwnd = self.init_hwnd
            self.opt_hwnd = self.init_hwnd
        else:
            self.init_hwnd = win32gui.GetDesktopWindow()
            self.cap_hwnd = self.init_hwnd
            self.opt_hwnd = -1
        self.app = app
        self.screen = full_screen

    def __str__(self):
        return "\ninit_hwnd:%d cap_hwnd:%d opt_hwnd:%d app:%s screen:%s\n" % (
            self.init_hwnd, self.cap_hwnd, self.opt_hwnd, self.app, self.screen)

    def _get_option_screen(self, hwnd_, paras):
        ret_list, text_len_lim = paras
        s = win32gui.GetWindowText(hwnd_)
        if len(s) > text_len_lim:
            ret_list.append(hwnd_)
        return 1

    def get_child_screen(self, parent_hwnd=None, window_text_lim=3):
        """
        搜寻某个窗口的所有子窗口的名字

        :param parent_hwnd: 父窗口句柄
        :param window_text_lim: 子窗口名字长度的限制
        :return: 所有一级子窗口的名字
        """
        if parent_hwnd is None:
            parent_hwnd = self.init_hwnd

        hwnd_list = []
        win32gui.EnumChildWindows(parent_hwnd, self._get_option_screen, (hwnd_list, window_text_lim))
        return hwnd_list

    def set_cap_hwnd(self, cap_hwnd: int):
        """
        设置截屏相关的窗口句柄

        :param cap_hwnd: 截屏相关的窗口句柄，非必要，若不指定则使用实例中的默认截图句柄
        :return:
        """
        self.cap_hwnd = cap_hwnd

    def set_opt_hwnd(self, opt_hwnd: int):
        """
        设置操作相关的窗口句柄

        :param opt_hwnd: 要操作的窗口的句柄，非必要，不指定则采用实例默认的操作句柄
        :return:
        """
        self.opt_hwnd = opt_hwnd

    def set_foreground_window(self, cap_hwnd=-1):
        """
        将本窗口（有截图句柄指定）放置到所有窗口的顶部，显示出来

        :param cap_hwnd: 截图句柄，非必要，不指定则采用实例默认的截图句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd

        hForeWnd = win32gui.GetForegroundWindow()
        FormThreadID = win32api.GetCurrentThreadId()
        CWndThreadID = win32process.GetWindowThreadProcessId(hForeWnd)
        win32process.AttachThreadInput(CWndThreadID[0], FormThreadID, True)
        win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNORMAL)
        win32gui.SetForegroundWindow(cap_hwnd)
        win32process.AttachThreadInput(CWndThreadID[0], FormThreadID, False)

    def put_window_to_bottom(self, cap_hwnd=-1):
        """
        将本窗口（有截图句柄指定）放置到所有窗口的底部

        :param cap_hwnd: 截图句柄，非必要，不指定则采用实例默认的截图句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd

        x1, y1, x2, y2 = win32gui.GetWindowRect(cap_hwnd)
        win32gui.SetWindowPos(cap_hwnd, win32con.HWND_BOTTOM, x1, y1, x2 - x1, y2 - y1, win32con.SWP_SHOWWINDOW)

    def get_window_rect(self, mode=0, cap_hwnd=-1):
        """
        获取窗口矩形的参数。

        返回形式有两种：

        1、返回左上角(x1,y1)坐标和窗口的宽w高h

        2、返回左上角(x1,y1)和右下角(x2,y2)的坐标

        :param mode: 希望返回的形式
        :param cap_hwnd: 指定的窗口句柄，其实也是截图句柄，非必要，不指定则采用实例默认的截图句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd

        win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNOACTIVATE)
        x1, y1, x2, y2 = win32gui.GetWindowRect(cap_hwnd)

        if mode == 1:
            return x1, y1, x2, y2
        return x1, y1, x2 - x1, y2 - y1

    def set_window_rect(self, x=0, y=0, w=1920, h=1030, mode=0, cap_hwnd=-1):
        """
        设置窗口大小，包括最小化窗口和最大化窗口

        :param x: 窗口左上方横坐标
        :param y: 窗口左上方横坐标
        :param w: 窗口宽度
        :param h: 窗口高度
        :param mode: 调节模式。0是自定义大小（默认），1是最小化窗口，2是恢复之前自定义的窗口大小，3是最大化窗口，其它数字无效均认为是2
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd

        if mode == 0:
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.SetWindowPos(cap_hwnd, win32con.HWND_TOP, x, y, w, h, win32con.SWP_SHOWWINDOW)
        elif mode == 1:
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWMINIMIZED)
        elif mode == 2:
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNORMAL)
        elif mode == 3:
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWMAXIMIZED)
        else:
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.ShowWindow(cap_hwnd, win32con.SW_SHOWDEFAULT)

    def fast_capture(self, x: int, y: int, w: int, h: int, cap_hwnd=-1):
        """
        快速截屏，使用实例内部默认的截图句柄或者参数指定截图句柄

        :param x: 截图的左上角横坐标（相对于窗口左上角）
        :param y: 截图的左上角纵坐标（相对于窗口左上角）
        :param w: 截图的宽度
        :param h: 截图的高度
        :param cap_hwnd: 截图句柄，非必要，不指定则使用实例默认的截图句柄
        :return: 返回截得的图形，格式为ndarray
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd

        img = self.screen.grabWindow(cap_hwnd, x, y, w, h).toImage()
        ptr = img.constBits()
        ptr.setsize(img.byteCount())
        return np.array(ptr).reshape(img.height(), img.width(), 4)[..., :-1][..., ::-1]

    # def fast_capture(self, x: int, y: int, w: int, h: int, cap_hwnd=-1):
    #     if cap_hwnd <= 0:
    #         return np.array(pyautogui.screenshot())[y:y+h, x:x+w]

    #     img = self.screen.grabWindow(cap_hwnd, x, y, w, h).toImage()
    #     ptr = img.constBits()
    #     ptr.setsize(img.byteCount())
    #     return np.array(ptr).reshape(img.height(), img.width(), 4)

    def key_down(self, vkey, opt_hwnd=-1):
        """
        在操作句柄对应的窗口上，按住键盘上的某个按键不放

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        win32gui.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)

    def key_up(self, vkey, opt_hwnd=-1):
        """
        在操作句柄对应的窗口上，松开键盘上的某个按键

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        win32gui.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)

    def press_key(self, vkey, times=1, interval=0.01, opt_hwnd=-1):
        """
        在操作句柄对应的窗口上，按times次键盘上的某个按键

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param times: 按的次数
        :param interval: 连续两次之间的间隔
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        while times > 1:
            times -= 1
            win32gui.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
            win32gui.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)
            time.sleep(interval)
        win32gui.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
        win32gui.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)

    def type_write(self, string: str, interval=0.05, opt_hwnd=-1):
        """
        在操作句柄对应的窗口上，键入一串字符

        :param string: 要键入的字符串，可以包括中文
        :param interval: 字符串中，相邻两个字符的输入时间间隔
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        if interval > 0:
            for ch in string:
                win32gui.SendMessage(opt_hwnd, win32con.WM_CHAR, ord(ch), 0)
                time.sleep(interval)
        else:
            for ch in string:
                win32gui.SendMessage(opt_hwnd, win32con.WM_CHAR, ord(ch), 0)

    def set_clipboard_text(self, text: str):
        """
        将字符串导入剪切板中

        :param text: 要导入的字符串
        :return:
        """
        OpenClipboard()
        SetClipboardData(win32con.CF_UNICODETEXT, text)
        CloseClipboard()

    def paste(self, opt_hwnd=-1):
        """
        将剪切板中的字符串粘贴到操作句柄对应的窗口中

        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        win32api.SendMessage(opt_hwnd, win32con.WM_CLEAR, 0, 0)
        win32api.SendMessage(opt_hwnd, win32con.WM_PASTE, 0, 0)

    def mouse_left_down(self, x: int, y: int, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口中，在指定位置（相对于该窗口）按下鼠标左键

        :param x: 按下位置的横坐标（相对于窗口左上角）
        :param y: 按下位置的纵坐标（相对于窗口左上角）
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        client_pos = win32gui.ScreenToClient(opt_hwnd, (winx1 + x, winy1 + y))
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(opt_hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)

    def mouse_left_up(self, x=0, y=0, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口中，在指定位置（相对于该窗口）松开鼠标左键

        :param x: 松开位置的横坐标（相对于窗口左上角）
        :param y: 松开位置的纵坐标（相对于窗口左上角）
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        client_pos = win32gui.ScreenToClient(opt_hwnd, (winx1 + x, winy1 + y))
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(opt_hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)

    def mouse_right_down(self, x: int, y: int, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口中，在指定位置（相对于该窗口）按下鼠标右键

        :param x: 点击位置的横坐标（相对于窗口左上角）
        :param y: 点击位置的纵坐标（相对于窗口左上角）
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        client_pos = win32gui.ScreenToClient(opt_hwnd, (winx1 + x, winy1 + y))
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(opt_hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32api.SendMessage(opt_hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, tmp)

    def mouse_right_up(self, x=0, y=0, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口中，在指定位置（相对于该窗口）松开鼠标右键

        :param x: 松开位置的横坐标（相对于窗口左上角）
        :param y: 松开位置的纵坐标（相对于窗口左上角）
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        client_pos = win32gui.ScreenToClient(opt_hwnd, (winx1 + x, winy1 + y))
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(opt_hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32api.SendMessage(opt_hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, tmp)

    def click(self, x, y, times=1, interval=0.01, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口中，在指定位置（相对于该窗口）连续点击鼠标左键tiems次

        :param x: 松开位置的横坐标（相对于窗口左上角）
        :param y: 松开位置的纵坐标（相对于窗口左上角）
        :param times: 连续按下的次数
        :param interval: 连续按下之间的间隔时间
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        if opt_hwnd != -1:
            client_pos = win32gui.ScreenToClient(opt_hwnd, (winx1 + x, winy1 + y))
            tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
            win32gui.SendMessage(opt_hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
            while times > 1:
                times -= 1
                win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
                win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)
                time.sleep(interval)
            win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
            win32api.SendMessage(opt_hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)
        else:
            win32api.SetCursorPos((x, y))
            while times > 1:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
                time.sleep(0.005)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
                time.sleep(interval)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            time.sleep(0.005)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

    def mouse_move(self, from_pos: tuple, to_pos: tuple, duration=0.1, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口上，将鼠标从起始位置移动到终点位置

        :param from_pos: 起始位置（相对窗口左上方），格式为(x,y)
        :param to_pos: 终点位置（相对窗口左上方），格式为(x,y)
        :param duration: 移动的间隔时间
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        times = math.ceil(duration / 0.01) if duration > 0 else 1
        delta_x = (to_pos[0] - from_pos[0]) / times
        delta_y = (to_pos[1] - from_pos[1]) / times

        winx1, winy1, winx2, winy2 = win32gui.GetWindowRect(cap_hwnd)
        clint_x, clint_y = win32gui.ScreenToClient(opt_hwnd, (winx1 + from_pos[0], winy1 + from_pos[1]))

        x, y = (clint_x, clint_y)
        win32api.SendMessage(opt_hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, (y << 16 | x))
        while times > 0:
            times -= 1
            time.sleep(0.009)
            x += delta_x
            y += delta_y
            win32api.SendMessage(opt_hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON,
                                 (math.floor(y) << 16 | math.floor(x)))
        win32api.SendMessage(opt_hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON,
                             ((clint_y + to_pos[1] - from_pos[1]) << 16 | (clint_x + to_pos[0] - from_pos[0])))

    def mouse_drag(self, from_pos: tuple, to_pos: tuple, duration=0.1, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作句柄指定的窗口上，将鼠标从起始位置拖拽到终点位置

        :param from_pos: 起始位置（相对窗口左上方），格式为(x,y)
        :param to_pos: 终点位置（相对窗口左上方），格式为(x,y)
        :param duration: 拖拽的间隔时间
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        self.mouse_left_down(from_pos[0], from_pos[1], cap_hwnd, opt_hwnd)
        time.sleep(0.2)
        self.mouse_move(from_pos, to_pos, duration, cap_hwnd, opt_hwnd)
        time.sleep(0.2)
        self.mouse_left_up(to_pos[0], to_pos[1], cap_hwnd, opt_hwnd)

    def scroll(self, x: int, y: int, val=120, cap_hwnd=-1, opt_hwnd=-1):
        """
        在操作窗口置顶的窗口上，在指定位置滚动鼠标滚轮混动一定距离

        滚动距离为正时滚轮倾向于远离你（一般页面会出现上面的内容）

        滚动距离为负时滚轮倾向于靠近你（一般页面会出现下面的内容）

        :param x: 滚动时鼠标位置的横坐标（相对于窗口左上方）
        :param y: 滚动时鼠标位置的纵坐标（相对于窗口左上方）
        :param val: 滚动的距离
        :param cap_hwnd: 截图句柄，非必要，若不指定则使用实例中的默认截图句柄
        :param opt_hwnd: 操作句柄，非必要，若不指定则使用实例中的默认操作句柄
        :return:
        """
        cap_hwnd = cap_hwnd if cap_hwnd > 0 else self.cap_hwnd
        opt_hwnd = opt_hwnd if opt_hwnd > 0 else self.opt_hwnd

        winx1, winy1, _, _ = win32gui.GetWindowRect(cap_hwnd)
        win_rx, win_ry = (winx1 + x, winy1 + y)
        clint_x, clint_y = win32gui.ScreenToClient(opt_hwnd, (win_rx, win_ry))

        win32api.SendMessage(opt_hwnd, win32con.WM_MOUSEMOVE, 0, (clint_y << 16 | clint_x))
        win32api.SendMessage(opt_hwnd, win32con.WM_MOUSEWHEEL, val << 16, (win_ry << 16 | win_rx))

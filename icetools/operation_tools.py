# -*- coding = utf-8 -*-
# @Time：2022/12/30 17:58
# @Author：kksk43
# @File：operation_tools.py
# @Software：PyCharm

import time

import numpy as np
import pyautogui
import pyperclip
import win32api
import win32con
# from cnocr import CnOcr

from icetools.window_tools import fast_capture_full, WindowTool

import cv2


class OperationTool:
    threshold = 0.8
    # ocr = CnOcr()

    def __init__(self):
        super().__init__()

    @staticmethod
    def tp_match(src_img, temp_img, threshold, mode=0, mask=None):
        """
        模版匹配，固定采用TM_CCOEFF_NORMED方法

        :param src_img: 源图像（待检测图像）
        :param temp_img: 模版图像
        :param threshold: 模版匹配相似度阈值
        :param mode: 匹配模式，0则为单一匹配，1则返回所有满足要求的区域
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 返回 单一坐标 或 多个坐标（包括匹配度） 或 None
        """
        h, w = temp_img.shape[:2]
        res = cv2.matchTemplate(cv2.cvtColor(src_img, cv2.COLOR_BGR2GRAY),
                                cv2.cvtColor(temp_img, cv2.COLOR_BGR2GRAY), cv2.TM_CCOEFF_NORMED, mask=mask)
        if mode == 0:
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            # print(max_val)
            if max_val >= threshold:
                return max_loc, (max_loc[0] + w, max_loc[1] + h)
        elif mode == 1:
            loc = np.where(res >= threshold)
            return [((pt[1], pt[0]), (pt[1] + w, pt[0] + h), res[pt]) for pt in zip(*loc)]
        return None

    @staticmethod
    def is_exist(temp_img, tpmatch_threshold=-1, app_win: WindowTool = None, area=(0, 0, 1920, 1030), src_img=None,
                 is_rgb=False, rgb_threshold=0.9875, interval=0.05, time_lim=-1, sleep_time=0.0,
                 exception_throw=False, mask=None) -> tuple:
        """
        检查模版图像是否在屏幕/某个窗口/给定图像的某个区域出现，若出现则返回出现时的位置，若没有出现则返回None

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不是等待time_lim后再检测

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param temp_img: 模版图片，必须是ndarray类型
        :param tpmatch_threshold: 模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param area: 检测区域，以检测窗口左上方为原点，格式为(x, y, w, h)，设置好区域可有效减小检测花费的时间
        :param src_img: 待查找图片，非必要，若有且time_lim<0则取消截屏，若没有则从窗口中截图获得
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param interval: 配合time_lim使用，当time_lim>0且当前待查找图片中没有模版图片时，第二次截图检测的间隔时间
        :param time_lim: 在time_lim这个时间段内屏幕是否出现过模版图片
        :param sleep_time: 成功执行后的睡眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 返回模版图像出现的位置（相对于截图或给定的图片），若不存在（找不到）则返回None
        """
        if tpmatch_threshold < 0:
            tpmatch_threshold = OperationTool.threshold
        if is_rgb and mask is not None:
            temp_img = cv2.bitwise_and(temp_img, temp_img, mask=mask)

        img_h, img_w = temp_img.shape[:2]
        x, y, w, h = area
        if src_img is None:
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
        time_lim += time.time()
        while True:
            pos = OperationTool.tp_match(src_img, temp_img, tpmatch_threshold, 0, mask)
            if is_rgb and pos:
                src_img = src_img[pos[0][1]:pos[1][1], pos[0][0]:pos[1][0]]
                if mask is not None:
                    src_img = cv2.bitwise_and(src_img, src_img, mask=mask)
                # cv2.imshow("img", src_img)
                # cv2.waitKey(0)
                # cv2.destroyAllWindows()

                delta = max((np.sum(np.abs(temp_img[:, :, 0].astype(int) - src_img[:, :, 0].astype(int))),
                             np.sum(np.abs(temp_img[:, :, 1].astype(int) - src_img[:, :, 1].astype(int))),
                             np.sum(np.abs(temp_img[:, :, 2].astype(int) - src_img[:, :, 2].astype(int))))) / (
                                img_w * img_h)
                # print(delta, 128 * (1 - rgb_threshold))
                if delta < 128 * (1 - rgb_threshold):
                    break
                pos = None
            if time.time() >= time_lim or pos:
                break
            time.sleep(interval)
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)

        if exception_throw and not pos:
            raise Exception("Time Limit Exceeded!")
        time.sleep(sleep_time)
        # print(pos[0][0], pos[0][1], pos[1][0]-pos[0][0], pos[1][1]-pos[0][1])
        return pos

    @staticmethod
    def first_exist(temp_img_list, interval=0.5, time_lim=999999999, area=(0, 0, 1920, 1030),
                    tpmatch_threshold=-1, app_win: WindowTool = None,
                    is_rgb=False, rgb_threshold=0.9875, sleep_time=0.0, exception_throw=False,
                    mask_list: list = None) -> int:
        """
        根据给出的一些模版图像，找到在屏幕/某个窗口/给定图像上某个区域内首个出现的模版图像

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不是等待time_lim后再检测

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param temp_img_list: 模版图片序列，必须是ndarray类型
        :param interval: 配合time_lim使用，当time_lim>0且当前待查找图片中没有出现任何模版图片时，第二次截图检测的间隔时间
        :param time_lim: 在time_lim这个时间段内屏幕是否出现过任何的模版图片。
        :param area: 检测区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param tpmatch_threshold: 模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param sleep_time: 成功执行后的睡眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask_list: 蒙版mask序列，对应temp_img_list中的图像，若该列表长度小于temp_img_list，默认后边的图像没有mask
        :return: 第一个检测到存在的模版的索引号，即在temp_img_list中的位置
        """
        if mask_list is None:
            mask_list = []
        mask_list += [None] * (len(temp_img_list) - len(mask_list))
        x, y, w, h = area
        time_lim += time.time()
        while True:
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
            for idx, temp_img in enumerate(temp_img_list):
                if OperationTool.is_exist(temp_img, tpmatch_threshold, app_win, (0, 0, w, h), src_img,
                                          is_rgb, rgb_threshold, mask=mask_list[idx]):
                    return idx
            if time.time() >= time_lim:
                break
            time.sleep(interval)
        if exception_throw:
            raise Exception("Time Limit Exceeded!")
        time.sleep(sleep_time)
        return -1

    @staticmethod
    def key_down(vkey, opt_hwnd=-1, sleep_time=0.0):
        """
        按下vkey对应的按键

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param opt_hwnd: 操作窗口的句柄，非必要
        :param sleep_time: 完成操作后的休眠时间
        :return:
        """
        if opt_hwnd > 0:
            win32api.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
        else:
            win32api.keybd_event(vkey, 0, 0, 0)
        time.sleep(sleep_time)

    @staticmethod
    def key_up(vkey, opt_hwnd=-1, sleep_time=0.0):
        """
        松开vkey对应的按键

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param opt_hwnd: 操作窗口的句柄，非必要
        :param sleep_time: 完成操作后的休眠时间
        :return:
        """
        if opt_hwnd > 0:
            win32api.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)
        else:
            win32api.keybd_event(vkey, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(sleep_time)

    @staticmethod
    def press_key(vkey, opt_hwnd=-1, times=1, interval=0.05, sleep_time=0.0):
        """
        按下vkey对应的按键times次

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey: 虚拟按键
        :param opt_hwnd: 操作窗口的句柄，非必要
        :param times: 按键按下的次数
        :param interval: 相邻两次按下的时间间隔
        :param sleep_time: 执行完操作后的休息时间
        :return:
        """
        if opt_hwnd > 0:
            while times > 1:
                times -= 1
                win32api.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
                win32api.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)
                time.sleep(interval)
            win32api.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
            win32api.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)
        else:
            while times > 1:
                times -= 1
                win32api.keybd_event(vkey, 0, 0, 0)
                win32api.keybd_event(vkey, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(interval)
            win32api.keybd_event(vkey, 0, 0, 0)
            win32api.keybd_event(vkey, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(sleep_time)

    @staticmethod
    def hotkey(vkey_list: list, opt_hwnd=-1, interval=0.05, sleep_time=0.0):
        """
        组合键，将一堆vkey按顺序塞到vkey_list中，然后按顺序逐个按下按键再释放

        vkey是虚拟按键可用win32con.VK_xxx指示，与物理按键对应关系可见：https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

        :param vkey_list: 虚拟按键集合，将按顺序按下（包括按住和松开）
        :param opt_hwnd: 操作窗口对应的句柄，非必要
        :param interval: 相邻两个虚拟按键按住的时间间隔，松开时与此无关（一瞬间全部松开）
        :param sleep_time: 执行完毕后的睡眠时间
        :return:
        """
        vkey_rlist = vkey_list[::-1]
        if opt_hwnd > 0:
            for vkey in vkey_list:
                win32api.SendMessage(opt_hwnd, win32con.WM_KEYDOWN, vkey, 0)
                time.sleep(interval)
            for vkey in vkey_rlist:
                win32api.SendMessage(opt_hwnd, win32con.WM_KEYUP, vkey, 0)
        else:
            for vkey in vkey_list:
                win32api.keybd_event(vkey, 0, 0, 0)
                time.sleep(interval)
            for vkey in vkey_rlist:
                win32api.keybd_event(vkey, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(sleep_time)

    @staticmethod
    def paste_text(text: str, opt_hwnd=-1, sleep_time=0.0):
        """
        将指定文本以粘贴方式写出来

        :param text: 指定的文本内容
        :param opt_hwnd: 操作窗口的句柄，非必要
        :param sleep_time: 完成操作后的休眠时间
        :return:
        """
        pyperclip.copy(text)
        pyperclip.paste()
        time.sleep(0.2)
        if opt_hwnd > 0:
            win32api.SendMessage(opt_hwnd, win32con.WM_CLEAR, 0, 0)
            win32api.SendMessage(opt_hwnd, win32con.WM_PASTE, 0, 0)
        else:
            pyautogui.hotkey('Ctrl', 'v')
        time.sleep(sleep_time)

    @staticmethod
    def click(temp_img=None, times=1, interval=0.2, relative_pos=(0, 0), tpmatch_threshold=-1,
              area=(0, 0, 1920, 1030), app_win: WindowTool = None, sleep_time=0.0, mask=None) -> bool:
        """
        完成在屏幕/窗口的指定位置上进行一定次数的点击操作

        最后点击的位置是area的左上角位置+temp_img图像的中心偏移位置+relative_pos自定义偏移

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param interval: 两次点击以上时，两次点击之间的间隔时间
        :param temp_img: 模版图像，非必要，给出后点击的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param times: 点击次数
        :param relative_pos: 自定义的点击偏移位置
        :param tpmatch_threshold: 模版匹配的阈值
        :param area: 识别temp_img的限制区域，非必要，若未指定则area定义为(0, 0, 1920, 1030)
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param sleep_time: 执行完点击后的睡眠时间
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 是否成功执行点击动作，True是，False否
        """
        if tpmatch_threshold < 0:
            tpmatch_threshold = OperationTool.threshold

        x, y, w, h = area
        pos = relative_pos
        if temp_img is not None:
            ori_pos = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                temp_img, tpmatch_threshold, 0, mask)
            if not ori_pos:
                return False

            pos = (x + (ori_pos[0][0] + ori_pos[1][0]) // 2 + relative_pos[0],
                   y + (ori_pos[0][1] + ori_pos[1][1]) // 2 + relative_pos[1])

        if app_win:
            app_win.click(pos[0], pos[1], times, interval)
        elif times == 1:
            pyautogui.click(pos[0], pos[1])
        elif times == 2:
            pyautogui.doubleClick(pos[0], pos[1])
        elif times > 2:
            for i in range(times - 1):
                pyautogui.click(pos[0], pos[1])
                time.sleep(interval)
            pyautogui.click(pos[0], pos[1])
        time.sleep(sleep_time)
        return True

    @staticmethod
    def click_until_showup(show_img, click_img=None, interval=0.5, time_lim=999999999,
                           show_threshold=-1, app_win=None, show_area=(0, 0, 1920, 1030),
                           is_rgb=False, rgb_threshold=0.9875,
                           double_click=False, relative_pos=(0, 0), click_threshold=-1,
                           click_area=(0, 0, 1920, 1030), sleep_time=0, exception_throw=False,
                           show_img_mask=None, click_img_mask=None):
        """
        点击指定位置直到show_img图像出现。第一次点击前会先检测一遍，若已出现则不点击直接结束

        若time_lim＜=0，那么该函数只会检测一次，不会点击

        若time_lim＜=interval，那么该函数会检测两次，中间点击一次

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param show_img: 判定出现的图像
        :param click_img: 点击的模版图像，即click函数的temp_img
        :param interval: 两轮“判断-点击”之间的时间间隔
        :param time_lim: 在time_lim这个时间段内屏幕是否出现过模版图片
        :param show_threshold: show_img的模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param show_area: 检测show_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param double_click: 单击还是双击，规定一轮的点击只能为一次单击或双击
        :param relative_pos: 自定义的点击偏移位置
        :param click_threshold: click_img的模版匹配的阈值
        :param click_area: 检测click_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param sleep_time: 执行完整个操作后，最后休息的时间，若时间限制内未出现则提前结束
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param show_img_mask: show_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :param click_img_mask: click_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 时间限制内，目标图像是否出现
        """
        times = 2 if double_click else 1
        time_lim += time.time()
        while not OperationTool.is_exist(show_img, show_threshold, app_win, show_area, None,
                                         is_rgb, rgb_threshold, mask=show_img_mask):
            if time.time() >= time_lim:
                if exception_throw:
                    raise Exception("Time Limit Exceeded!")
                return False
            OperationTool.click(click_img, times, 0.2, relative_pos, click_threshold, click_area, app_win,
                                0, click_img_mask)
            time.sleep(interval)
        time.sleep(sleep_time)
        return True

    @staticmethod
    def click_until_disappear(disappear_img, click_img=None, interval=0.5, time_lim=999999999,
                              show_threshold=-1, app_win=None, show_area=(0, 0, 1920, 1030),
                              is_rgb=False, rgb_threshold=0.9875,
                              double_click=False, relative_pos=(0, 0), click_threshold=-1,
                              click_area=(0, 0, 1920, 1030), sleep_time=0, exception_throw=False,
                              show_img_mask=None, click_img_mask=None):
        """
        点击指定位置直到disappear_img图像消失。第一次点击前会先检测一遍，若已消失则不点击直接结束

        若time_lim＜=0，那么该函数只会检测一次，不会点击

        若time_lim＜=interval，那么该函数会检测两次，中间点击一次

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param disappear_img: 判定消失的图像
        :param click_img: 点击的模版图像
        :param interval: 两轮“判断-点击”之间的时间间隔
        :param time_lim: 在time_lim这个时间段内屏幕中的模版图片是否消失
        :param show_threshold: show_img的模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param show_area: 检测disappear_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param double_click: 单击还是双击，规定一轮的点击只能为一次单击或双击
        :param relative_pos: 自定义的点击偏移位置
        :param click_threshold: click_img的模版匹配的阈值
        :param click_area: 检测click_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param sleep_time: 执行完整个操作后，最后休息的时间，若时间限制内未消失则提前结束
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param show_img_mask: show_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :param click_img_mask: click_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return:
        """
        times = 2 if double_click else 1
        time_lim += time.time()
        while OperationTool.is_exist(disappear_img, show_threshold, app_win, show_area, None,
                                     is_rgb, rgb_threshold, mask=show_img_mask):
            if time.time() >= time_lim:
                if exception_throw:
                    raise Exception("Time Limit Exceeded!")
                return False
            OperationTool.click(click_img, times, 0.2, relative_pos, click_threshold, click_area, app_win, 0,
                                click_img_mask)
            time.sleep(interval)
        time.sleep(sleep_time)
        return True

    @staticmethod
    def wait_until_showup(show_img=None, interval=0.5, time_lim=999999999,
                          show_threshold=-1, app_win=None, show_area=(0, 0, 1920, 1030),
                          is_rgb=False, rgb_threshold=0.9875, sleep_time=0, exception_throw=False, mask=None):
        """
        等待直到某个模版图像出现在屏幕/窗口上，或者干等sleep_time

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不是等待time_lim后再检测

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param show_img: 判定出现的图像
        :param interval: 两次相邻检测的间隔时间
        :param time_lim: 在time_lim这个时间段内屏幕是否出现模版图片
        :param show_threshold: show_img的模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param show_area: 检测show_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param sleep_time: 执行完全部操作后的休眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 干等或者时间截止时都不存在则返回False，否则返回True
        """
        if show_img is None:
            time.sleep(sleep_time)
            return False

        return OperationTool.is_exist(show_img, show_threshold, app_win, show_area, None,
                                      is_rgb, rgb_threshold, interval, time_lim,
                                      sleep_time, exception_throw, mask) is not None

    @staticmethod
    def wait_until_disappear(disappear_img=None, interval=0.5, time_lim=999999999,
                             show_threshold=-1, app_win=None, show_area=(0, 0, 1920, 1030),
                             is_rgb=False, rgb_threshold=0.9875, sleep_time=0, exception_throw=False, mask=None):
        """
        等待直到某个模版图像不存在于在屏幕/窗口上，或者干等sleep_time

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不是等待time_lim后再检测

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param disappear_img: 判定消失的图像
        :param interval: 两次相邻检测的间隔时间
        :param time_lim: 在time_lim这个时间段内屏幕的模版图片是否消失
        :param show_threshold: show_img的模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param show_area: 检测disappear_img的限制区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param is_rgb: 是否以彩色图片方式查找，非必要
        :param rgb_threshold: 彩色图片寻找的阈值
        :param sleep_time: 执行完全部操作后的休眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 干等或者时间截止时都一直存在则返回False，否则返回True
        """
        if disappear_img is None:
            time.sleep(sleep_time)
            return False

        if show_threshold < 0:
            show_threshold = OperationTool.threshold

        if is_rgb and mask is not None:
            disappear_img = cv2.bitwise_and(disappear_img, disappear_img, mask=mask)

        img_h, img_w = disappear_img.shape[:2]
        x, y, w, h = show_area
        time_lim += time.time()
        while True:
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
            # cv2.imshow("img", src_img)
            # cv2.waitKey(0)
            # cv2.destroyAllWindows()
            pos = OperationTool.tp_match(src_img, disappear_img, show_threshold, 0, mask)
            if is_rgb and pos:
                src_img = src_img[pos[0][1]:pos[1][1], pos[0][0]:pos[1][0]]
                if mask is not None:
                    src_img = cv2.bitwise_and(src_img, src_img, mask=mask)

                delta = max((np.sum(np.abs(disappear_img[:, :, 0].astype(int) - src_img[:, :, 0].astype(int))),
                             np.sum(np.abs(disappear_img[:, :, 1].astype(int) - src_img[:, :, 1].astype(int))),
                             np.sum(np.abs(disappear_img[:, :, 2].astype(int) - src_img[:, :, 2].astype(int))))) / (
                                img_w * img_h)
                # print(delta, 128 * (1 - rgb_threshold))
                if delta > 128 * (1 - rgb_threshold):
                    pos = None
            if time.time() >= time_lim or not pos:
                break
            time.sleep(interval)

        if pos:
            if exception_throw:
                raise Exception("Time Limit Exceeded!")
            return False
        time.sleep(sleep_time)
        # print(pos[0][0], pos[0][1], pos[1][0]-pos[0][0], pos[1][1]-pos[0][1])
        return True

    @staticmethod
    def get_pos(temp_img, tpmatch_threshold=-1, app_win: WindowTool = None, area=(0, 0, 1920, 1030),
                interval=0.1, time_lim=-1, sleep_time=0.0, exception_throw=False, mask=None):
        """
        返回指定图像在屏幕的位置（可以指定一段时间检测），只使用做灰度图的匹配，若想对彩色图更加精准的辨别，可以使用is_exist方法

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不是等待time_lim后再检测

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param temp_img: 模版图片，必须是ndarray类型
        :param tpmatch_threshold: 模版匹配的阈值
        :param app_win: 检测窗口的WindowTool实例
        :param area: 检测区域，以检测窗口左上方为原点，格式为(x, y, w, h)
        :param interval: 配合time_lim使用，当time_lim>0且当前待查找图片中没有模版图片时，第二次截图检测的间隔时间
        :param time_lim: 在time_lim这个时间段内屏幕是否出现过模版图片。
        :param sleep_time: 成功执行后的睡眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 若检测到模版图像则返回图像中心位置，若不存在则返回None
        """
        if tpmatch_threshold < 0:
            tpmatch_threshold = OperationTool.threshold

        x, y, w, h = area
        time_lim += time.time()
        while True:
            res = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                temp_img, tpmatch_threshold, 0, mask)
            if time.time() >= time_lim or res is not None:
                break
            time.sleep(interval)
        if res is None:
            if exception_throw:
                raise Exception("Time Limit Exceeded!")
            return None
        time.sleep(sleep_time)
        return x + (res[0][0] + res[1][0]) // 2, y + (res[0][1] + res[1][1]) // 2

    @staticmethod
    def move_to(to_area=(0, 0, 1920, 1030), to_img=None, to_rpos=(0, 0), to_threshold=-1,
                from_area=(0, 0, 1920, 1030), from_img=None, from_rpos=(0, 0), from_threshold=-1,
                from_cursor=False, app_win: WindowTool = None, duration=0.1, sleep_time=0.0,
                to_img_mask=None, from_img_mask=None) -> bool:
        """
        将光标从起点位置移动到终点位置，其中起点位置默认为空，不使用时把鼠标当前位置作为起点位置

        起点位置是：from_area的左上角位置+from_img图像的中心位置+from_rpos自定义偏移，也可以直接使用鼠标当前位置

        终点位置是：to_area的左上角位置+to_img图像的中心位置+to_rpos自定义偏移

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param to_area: 识别to_img的限制区域，非必要，若未指定则to_area定义为(0, 0, 1920, 1030)
        :param to_img: 关于终点的模版图像，非必要，给出后终点的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param to_rpos: 终点的自定义偏移位置
        :param to_threshold: to_img图像的模版匹配的阈值，不指定则使用默认阈值
        :param from_area: 识别from_img的限制区域，非必要，若未指定则from_area定义为(0, 0, 1920, 1030)
        :param from_img: 关于终点的模版图像，非必要，给出后起点的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param from_rpos: 起点的自定义偏移位置
        :param from_threshold: from_img图像的模版匹配的阈值，不指定则使用默认阈值
        :param from_cursor: 起点位置是否指定为鼠标当前位置
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param duration: 移动完成的过程时间
        :param sleep_time: 执行完移动后的睡眠时间
        :param to_img_mask: to_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :param from_img_mask: from_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 完整执行完返回True，若检测不到模版图像则返回False
        """
        if from_img is not None and not from_cursor:
            if from_threshold < 0:
                from_threshold = OperationTool.threshold

            x, y, w, h = from_area
            res = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                from_img, from_threshold, 0, from_img_mask)
            if res is None:
                return False

            from_rpos = (x + (res[0][0] + res[1][0]) // 2 + from_rpos[0],
                         y + (res[0][1] + res[1][1]) // 2 + from_rpos[1])

        if to_img is not None:
            if to_threshold < 0:
                to_threshold = OperationTool.threshold

            x, y, w, h = to_area
            res = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                to_img, to_threshold, 0, to_img_mask)
            if res is None:
                return False

            to_rpos = (x + (res[0][0] + res[1][0]) // 2 + to_rpos[0],
                       y + (res[0][1] + res[1][1]) // 2 + to_rpos[1])

        if app_win:
            if from_cursor:
                from_rpos = win32api.GetCursorPos()
            app_win.mouse_move(from_rpos, to_rpos, duration)
        else:
            if not from_cursor:
                win32api.SetCursorPos(from_rpos)
            pyautogui.moveTo(to_rpos[0], to_rpos[1], duration=duration)
        time.sleep(sleep_time)

        return True

    @staticmethod
    def scroll_until_showup(temp_img=None, stay_pos: tuple = None, distance=300, by_drag=False,
                            interval=0.5, time_lim=999999999, app_win: WindowTool = None,
                            area=(0, 0, 1920, 1030), tpmatch_threshold=-1, sleep_time=0.0, exception_throw=False,
                            mask=None) -> bool:
        """
        将鼠标滚轮在指定位置上滚动指定距离

        若time_lim＜=interval或time_lim＜=0，那么该函数只会检测一次且是立刻检测而不会滚动一次后再检测

        一般有两种方式：

        1、给出temp_img模版图像，意味着不断执行滚动时出现模版图像后提前停止

        2、没有给出模版图像，意味着只简单执行一次滚动

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param temp_img: 模版图像，非必要，若给出则当检测到该模版时滚动停止
        :param stay_pos: 滑动滚轮时，光标停留的位置，非必要，未指定时使用当前鼠标位置
        :param distance: 每次滚动的距离，为正时滚轮倾向于远离你，为负时滚轮倾向于靠近你
        :param by_drag: 是否通过拖拽方式实现
        :param interval: 相邻两次滚动的间隔时间
        :param time_lim: 整个滑动过程的限时，初值小于等于0时认为不限时，默认为999999999
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param area: 别temp_img的限制区域，非必要，若未指定则area定义为(0, 0, 1920, 1030)
        :param tpmatch_threshold: 模版匹配的阈值
        :param sleep_time: 执行完移动后的睡眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param mask: 使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 若只是滚动不检测测恒定返回True，若滚动直到检测出模版图像，则返回值表示有没有出现模版图像
        """

        if not stay_pos:
            stay_pos = win32api.GetCursorPos()
        to_pos = (stay_pos[0], stay_pos[1] + distance)
        if temp_img is None:
            if app_win:
                app_win.mouse_drag(stay_pos, to_pos, 0.2) if by_drag else app_win.scroll(stay_pos[0], stay_pos[1],
                                                                                         distance)
            else:
                if by_drag:
                    win32api.SetCursorPos(stay_pos)
                    pyautogui.dragTo(to_pos[0], to_pos[1], 0.2, button='left')
                else:
                    pyautogui.scroll(distance)
            time.sleep(sleep_time)
            return True

        if tpmatch_threshold < 0:
            tpmatch_threshold = OperationTool.threshold

        x, y, w, h = area
        time_lim += time.time()
        while OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                temp_img, tpmatch_threshold, 0, mask) is None:
            if app_win:
                app_win.mouse_drag(stay_pos, to_pos, 0.2) if by_drag else app_win.scroll(stay_pos[0], stay_pos[1],
                                                                                         distance)
            else:
                if by_drag:
                    win32api.SetCursorPos(stay_pos)
                    pyautogui.dragTo(to_pos[0], to_pos[1], 0.2, button='left')
                else:
                    pyautogui.scroll(distance)
            if time.time() >= time_lim:
                if exception_throw:
                    raise Exception("Time Limit Exceeded!")
                return False
            time.sleep(interval)
        time.sleep(sleep_time)
        return True

    @staticmethod
    def drag_to(to_area=(0, 0, 1920, 1030), to_img=None, to_rpos=(0, 0), to_threshold=-1,
                from_area=(0, 0, 1920, 1030), from_img=None, from_rpos=(0, 0), from_threshold=-1,
                from_cursor=False, app_win: WindowTool = None, duration=0.1, sleep_time=0.0,
                to_img_mask=None, from_img_mask=None) -> bool:
        """
        将光标从起点位置拖动到终点位置，其中起点位置默认为空，不使用时把鼠标当前位置作为起点位置

        起点位置是：from_area的左上角位置+from_img图像的中心位置+from_rpos自定义偏移，也可以直接使用鼠标当前位置

        终点位置是：to_area的左上角位置+to_img图像的中心位置+to_rpos自定义偏移

        对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param to_area: 识别to_img的限制区域，非必要，若未指定则to_area定义为(0, 0, 1920, 1030)
        :param to_img: 关于终点的模版图像，非必要，给出后终点的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param to_rpos: 终点的自定义偏移位置
        :param to_threshold: to_img图像的模版匹配的阈值，不指定则使用默认阈值
        :param from_area: 识别from_img的限制区域，非必要，若未指定则from_area定义为(0, 0, 1920, 1030)
        :param from_img: 关于终点的模版图像，非必要，给出后起点的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param from_rpos: 起点的自定义偏移位置
        :param from_threshold: from_img图像的模版匹配的阈值，不指定则使用默认阈值
        :param from_cursor: 起点位置是否指定为鼠标当前位置
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param duration: 拖动完成的过程时间
        :param sleep_time: 执行完拖动后的睡眠时间
        :param to_img_mask: to_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :param from_img_mask: from_img使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return: 完整执行完返回True，若检测不到模版图像则返回False
        """
        if from_img is not None and not from_cursor:
            if from_threshold < 0:
                from_threshold = OperationTool.threshold

            x, y, w, h = from_area
            res = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                from_img, from_threshold, 0, from_img_mask)
            if res is None:
                return False

            from_rpos = (x + (res[0][0] + res[1][0]) // 2 + from_rpos[0],
                         y + (res[0][1] + res[1][1]) // 2 + from_rpos[1])

        if to_img is not None:
            if to_threshold < 0:
                to_threshold = OperationTool.threshold

            x, y, w, h = to_area
            res = OperationTool.tp_match(
                app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h),
                to_img, to_threshold, 0, to_img_mask)
            if res is None:
                return False

            to_rpos = (x + (res[0][0] + res[1][0]) // 2 + to_rpos[0],
                       y + (res[0][1] + res[1][1]) // 2 + to_rpos[1])

        if app_win:
            if from_cursor:
                from_rpos = win32api.GetCursorPos()
            app_win.mouse_drag(from_rpos, to_rpos, duration)
        else:
            if not from_cursor:
                win32api.SetCursorPos(from_rpos)
            pyautogui.dragTo(to_rpos[0], to_rpos[1], duration=duration, button='left')
        time.sleep(sleep_time)

        return True

    @staticmethod
    def rgb_reject(temp_img, left_top: tuple = (0, 0), src_img=None,
                   app_win: WindowTool = None, rgb_threshold=0.9875, sleep_time=0.0):
        """
        使用彩色图像的哈密顿距离，判断模版图像与源图像在一定的相似阈值下是否被拒绝

        优先使用src_img作为源图像，若src_img为空，则选择截图作为源图像

        若采用截图作为源图像，则截图的左上角坐标由left_top指定，截图的宽高与temp_img保持一致

        :param temp_img: 模版图像
        :param left_top: 截图的左上角坐标（使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标）
        :param src_img: 是否给出原图片，非必要，若不指定则选择截图作为源图像
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param rgb_threshold: 彩色图片相似阈值
        :param sleep_time: 执行完后的睡眠时间
        :return: 若被拒绝则返回源图像，若接受则返回None
        """
        h, w, _ = temp_img.shape
        if src_img is None:
            x, y = left_top
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)

        delta = max((np.sum(np.abs(temp_img[:, :, 0].astype(int) - src_img[:, :, 0].astype(int))),
                     np.sum(np.abs(temp_img[:, :, 1].astype(int) - src_img[:, :, 1].astype(int))),
                     np.sum(np.abs(temp_img[:, :, 2].astype(int) - src_img[:, :, 2].astype(int))))) / (w * h)
        # print(delta)
        time.sleep(sleep_time)

        return src_img if delta > 128 * (1 - rgb_threshold) else None

    @staticmethod
    def click_until_changes(change_area, change_times, change_threshold=-1, is_rgb=True, rgb_threshold=0.9875,
                            click_area=(0, 0, 1920, 1030), click_img=None, click_rpos=(0, 0), click_threshold=-1,
                            interval=0.5, time_lim=999999999, app_win: WindowTool = None, sleep_time=0,
                            exception_throw=False, click_img_mask=None):
        """
        点击某个位置直到某个区域发生一定次数的变化

        点击的位置是click_area的左上角位置+click_img图像的中心偏移位置+click_rpos自定义偏移

        此外对于所有的area的定位，使用app_win则是相对于窗口左上角的坐标，否则是屏幕左上角的绝对坐标

        :param change_area: 检测变化的区域
        :param change_times: 要求变化的次数
        :param change_threshold: 检测变化区域图像的模版匹配阈值
        :param is_rgb: 是否使用色彩欧氏距离检测
        :param rgb_threshold: 采用色彩欧氏距离检测的阈值
        :param click_area: 点击的区域，非必要，未指定时默认为(0, 0, 1920, 1030)
        :param click_img: 点击区域内部的模板图像，非必要，给出后点击的位置与模版图像的中心有关，若未指定则图片的中心位置定义为(0,0)
        :param click_rpos: 自定义的点击偏移位置
        :param click_threshold: click_img图像的模版匹配阈值
        :param interval: 每次点击的时间间隔
        :param time_lim: 整个滑动过程的限时，初值小于等于0时认为不限时，默认为999999999
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param sleep_time: 执行完目的操作后的睡眠时间
        :param exception_throw: 若有设置限时（time_lim），超时后是否抛出异常
        :param click_img_mask: click_img_mask使用mask蒙版，蒙版中0/非0代表不考虑/考虑相应像素点
        :return:
        """
        if change_threshold < 0:
            change_threshold = OperationTool.threshold
        if click_threshold < 0:
            click_threshold = OperationTool.threshold

        x, y, w, h = change_area
        time_lim += time.time()
        temp_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
        while change_times > 0:
            change_times -= 1
            OperationTool.click(click_img, 1, 0.2, click_rpos, click_threshold,
                                click_area, app_win, 0, click_img_mask)
            time.sleep(interval)
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
            while OperationTool.tp_match(src_img, temp_img, change_threshold, 0) is not None and \
                    (not is_rgb or OperationTool.rgb_reject(
                        temp_img=temp_img,
                        src_img=src_img,
                        rgb_threshold=rgb_threshold) is None):
                if time.time() >= time_lim:
                    if exception_throw:
                        raise Exception("Time Limit Exceeded!")
                    return False
                OperationTool.click(click_img, 1, 0.2, click_rpos, click_threshold,
                                    click_area, app_win, 0, click_img_mask)
                time.sleep(interval)
                src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
            temp_img = src_img
        time.sleep(sleep_time)
        return True

    @staticmethod
    def detect_number_inrow(num_temps: list, number_area=(0, 0, 1920, 1030), src_img=None, app_win: WindowTool = None,
                            tpmatch_threshold=0.92, rgb_threshold=0.87, mask_list: list = None, cancel_scale=0.75):
        """
        该方法使用模版匹配识别一行内的数字串，要先手动把目标各数字的图像模版分别保存下来，按0~9的顺序排好

        :param num_temps: 各数字的模版图像，不同样式以及不同大小的数字图像不能通用
        :param number_area: 检测的区域，即数字串出现的大概位置（相对于窗口左上角），格式为(x,y,w,h)
        :param src_img: 检测的图像，默认为空，若该值非空，则直接使用该图像而不采取截屏的方式（即忽略ch_area参数）
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param tpmatch_threshold: 模版匹配的阈值
        :param rgb_threshold: 颜色图像差异的阈值
        :param mask_list: 蒙版mask序列，对应num_temps中的图像，若该列表长度小于num_temps，默认后边的图像没有mask
        :param cancel_scale: 模版匹配会得到重叠区域，该参数指明消去重叠区域的横向占比
        取值范围是0~1，值越大取消重叠的程度越高（甚至可能消掉原本相邻的数字）
        :return: 指定范围内的数字串
        """
        if mask_list is None:
            mask_list = []
        mask_list += [None] * (len(num_temps) - len(mask_list))

        if src_img is None:
            x, y, w, h = number_area
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
        selected_num = []
        for idx, num_temp in enumerate(num_temps):
            if mask_list[idx] is not None:
                num_temp = cv2.bitwise_and(num_temp, num_temp, mask=mask_list[idx])
            num_temp_h, num_temp_w = num_temp.shape[:2]
            cancel_width = num_temp_w * cancel_scale

            pos_list = OperationTool.tp_match(src_img, num_temp, tpmatch_threshold, 1, mask_list[idx])
            pos_list.append(((9999, 0), (0, 0), 0))
            pos_list = sorted(pos_list, key=lambda tpos: tpos[0][0])
            ac_left = max_loc = 0
            for i, pos in enumerate(pos_list):
                if pos[0][0] - pos_list[ac_left][0][0] <= cancel_width:
                    max_loc = max_loc if pos_list[max_loc][2] > pos[2] else i
                    continue
                # print(num)
                # print(pos_list[max_loc][2])
                pos = pos_list[max_loc]
                ac_left = max_loc = i

                # cv2.imshow("img", src_img)
                # cv2.waitKey(0)
                # cv2.destroyAllWindows()
                detect_img = src_img[pos[0][1]:pos[1][1], pos[0][0]:pos[1][0]]
                if mask_list[idx] is not None:
                    detect_img = cv2.bitwise_and(detect_img, detect_img, mask=mask_list[idx])
                # cv2.imshow("img", detect_img)
                # cv2.waitKey(0)
                # cv2.destroyAllWindows()

                delta = max((np.sum(np.abs(num_temp[:, :, 0].astype(int) - detect_img[:, :, 0].astype(int))),
                             np.sum(np.abs(num_temp[:, :, 1].astype(int) - detect_img[:, :, 1].astype(int))),
                             np.sum(np.abs(num_temp[:, :, 2].astype(int) - detect_img[:, :, 2].astype(int))))) / (
                                num_temp_w * num_temp_h)
                # print(delta, 128 * (1 - rgb_threshold))
                # print(pos[0][0])
                if delta < 128 * (1 - rgb_threshold):
                    selected_num.append((pos[0][0], pos[0][1], idx))
        selected_num = sorted(selected_num, key=lambda pos_x: pos_x[0])
        return "".join(str(pos_x[2]) for pos_x in selected_num)

    @staticmethod
    def detect_character_inrow(ch_temps: list, ch_str: str, ch_area=(0, 0, 1920, 1030), src_img=None,
                               app_win: WindowTool = None, tpmatch_threshold=0.92,
                               rgb_threshold=0.87, mask_list: list = None, cancel_scale=0.75):
        """
        该方法使用模版匹配识别字符串，要先手动把目标各字符的图像模版保存下来，并且保证ch_temps、ch_str、mask_list内容顺序对应好

        :param ch_temps: 所有待检测字符的模版图像，不同样式以及不同大小的字符图像不能通用，并且保证模版图像的顺序和ch_str中对应字符的顺序一致
        :param ch_str: 所有待检测字符组成的字符串
        :param ch_area: 检测的区域，即字符串出现的大概位置（相对于窗口左上角），格式为(x,y,w,h)
        :param src_img: 检测的图像，默认为空，若该值非空，则直接使用该图像而不采取截屏的方式（即忽略ch_area参数）
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param tpmatch_threshold: 模版匹配的阈值
        :param rgb_threshold: 颜色图像差异的阈值
        :param mask_list: 蒙版mask序列，对应ch_temps中的图像，若该列表长度小于ch_temps，默认后边的图像没有mask
        :param cancel_scale: 模版匹配会得到重叠区域，该参数指明消去重叠区域的横向占比
        取值范围是0~1，值越大取消重叠的程度越高（甚至可能消掉原本相邻的数字）
        :return: 指定范围内的字符串
        """
        assert len(ch_temps) == len(ch_str)

        if mask_list is None:
            mask_list = []
        mask_list += [None] * (len(ch_temps) - len(mask_list))

        if src_img is None:
            x, y, w, h = ch_area
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
        selected_num = []
        for idx, num_temp in enumerate(ch_temps):
            if mask_list[idx] is not None:
                num_temp = cv2.bitwise_and(num_temp, num_temp, mask=mask_list[idx])
            num_temp_h, num_temp_w = num_temp.shape[:2]
            cancel_width = num_temp_w * cancel_scale

            pos_list = OperationTool.tp_match(src_img, num_temp, tpmatch_threshold, 1, mask_list[idx])
            pos_list.append(((9999, 0), (0, 0), 0))
            pos_list = sorted(pos_list, key=lambda tpos: tpos[0][0])
            ac_left = max_loc = 0
            for i, pos in enumerate(pos_list):
                if pos[0][0] - pos_list[ac_left][0][0] <= cancel_width:
                    max_loc = max_loc if pos_list[max_loc][2] > pos[2] else i
                    continue
                # print(num)
                # print(pos_list[max_loc][2])
                pos = pos_list[max_loc]
                ac_left = max_loc = i

                # cv2.imshow("img", src_img)
                # cv2.waitKey(0)
                # cv2.destroyAllWindows()
                detect_img = src_img[pos[0][1]:pos[1][1], pos[0][0]:pos[1][0]]
                if mask_list[idx] is not None:
                    detect_img = cv2.bitwise_and(detect_img, detect_img, mask=mask_list[idx])
                # cv2.imshow("img", detect_img)
                # cv2.waitKey(0)
                # cv2.destroyAllWindows()

                delta = max((np.sum(np.abs(num_temp[:, :, 0].astype(int) - detect_img[:, :, 0].astype(int))),
                             np.sum(np.abs(num_temp[:, :, 1].astype(int) - detect_img[:, :, 1].astype(int))),
                             np.sum(np.abs(num_temp[:, :, 2].astype(int) - detect_img[:, :, 2].astype(int))))) / (
                                num_temp_w * num_temp_h)
                # print(delta, 128 * (1 - rgb_threshold))
                # print(pos[0][0])
                if delta < 128 * (1 - rgb_threshold):
                    selected_num.append((pos[0][0], pos[0][1], ch_str[idx]))
        selected_num = sorted(selected_num, key=lambda pos_x: pos_x[0])
        return "".join(pos_x[2] for pos_x in selected_num)

    @staticmethod
    def ocr_inrow(area=(0, 0, 1920, 1030), app_win: WindowTool = None, src_img=None):
        """
        使用开源的OCR工具包，识别单行的简体中文、繁体中文、英文、数字

        要么指定图像识别，要么指定区域识别，图像的优先级更高（即给定图像后不会识别区域）

        ↓该OCR工具包的文档↓

        https://cnocr.readthedocs.io/zh/latest/usage

        :param area: 待识别的区域（相对于窗口左上角），格式为(x,y,w,h)
        :param app_win: WindowTool实例，用来实现发送点击消息，解放鼠标指针
        :param src_img: 待识别的图像，默认为空，非空则直接检测该图像而不采用截图方式（即忽略area参数）
        :return: text是识别的内容，score是确信度
        """
        if src_img is None:
            x, y, w, h = area
            src_img = app_win.fast_capture(x, y, w, h) if app_win else fast_capture_full(x, y, w, h)
        return OperationTool.ocr.ocr_for_single_line(src_img[:, :, :3])

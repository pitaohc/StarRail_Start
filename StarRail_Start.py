import os
import subprocess
import sys
import time
import configparser
import traceback

import cv2
import numpy as np
import pyautogui
import win32com.client
import win32con
import win32gui
from PIL import ImageGrab

file_path = ""
windows_title = ""

def getColorRate(img, threshold_brightness, greater):
    """
    计算屏幕白色像素比例
    :param img: 图片
    :param threshold_brightness: 亮度阈值
    :param greater: 是否大于阈值
    :return: 超过或低于阈值的像素比例
    """
    # 计算直方图
    histogram = cv2.calcHist([img], [0], None, [256], [0, 256])
    print(histogram.shape)
    # 统计亮度低于的像素数量
    pixels_below_threshold = 0
    if (greater):
        pixels_below_threshold = np.sum(histogram[threshold_brightness:])
    else:
        pixels_below_threshold = np.sum(histogram[:threshold_brightness])
    total_pixels = img.shape[0] * img.shape[1]
    return pixels_below_threshold / total_pixels


def getExecutionFile(file_path: str):
    # 获得file_path的后缀名
    type = os.path.splitext(file_path)[1]
    # 将type转换为小写
    type = type.lower()
    if(type == ".exe"):
        return file_path
    elif(type == ".lnk"):
        # 解析快捷方式获取安装路径
        shell = win32com.client.Dispatch("WScript.Shell")
        install_dir = shell.CreateShortCut(file_path)
        # install_dir = install_dir.TargetPath
        install_dir = install_dir.TargetPath.replace('launcher.exe', '')
        # 拼接游戏exe路径
        game_exe = os.path.join(install_dir, 'Game', 'StarRail.exe')
        return game_exe


if __name__ == '__main__':
    # 隐藏自身控制台窗口

    try:
        # 读取当前目录下的配置文件
        config = configparser.ConfigParser(allow_no_value=True)
        config.read('config.ini', encoding='utf-8')
        file_path = config.get('luncher', 'file_path')  # 获取快捷方式或可执行程序路径
        windows_title = config.get('luncher', 'windows_title')  # 获取窗口标题
        greater = config.getboolean('luncher', 'greater')  # 获取是否大于阈值
        threshold_brightness = config.getint('luncher', 'threshold_brightness')  # 获取亮度阈值
        threshold_rate = config.getfloat('luncher', 'threshold_rate')  # 获取比例阈值
        wait_time = config.getint("luncher", "wait_time") # 获取等待时间
        transition_steps = config.getint('luncher', 'transition_steps') # 获取过渡步数
        hidden = config.getboolean("luncher","hidden") # 获取是否隐藏窗口
        game_exe = getExecutionFile(file_path) # 获取游戏exe路径
    except configparser.Error:
        print("读取配置文件错误")
        print(traceback.format_exc())
        sys.exit(1)

    if hidden:
        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_HIDE)

    print("配置文件读取完毕!\n"
          "目标文件路径: {}\n"
          "窗口标题: {}\n"
          "是否大于阈值: {}\n"
          "亮度阈值: {}\n"
          "比例阈值: {}\n"
          "可执行程序路径：{}\n".format(file_path, windows_title, greater, threshold_brightness, threshold_rate,game_exe))
    # 获取屏幕分辨率
    screen_width, screen_height = pyautogui.size()
    print(f"屏幕分辨率: {screen_width}x{screen_height}")
    while True:
        # 检查是否有星穹铁道进程
        if os.system('tasklist | findstr StarRail.exe') == 0:
            print("星穹铁道已启动!")
            break

        # 截图
        screenshot = cv2.cvtColor(np.array(ImageGrab.grab(bbox=(0, 0, screen_width, screen_height))), cv2.COLOR_BGR2RGB)

        # 计算屏幕白色像素比例
        white_percentage = getColorRate(screenshot, threshold_brightness, greater)
        print(f"屏幕含穹量{white_percentage * 100}%")

        # 判断是否满足启动条件
        if white_percentage >= threshold_rate:

            # 创建过渡用的图片
            white_image = np.full((screen_height, screen_width, 3), 0, dtype=np.uint8)
            # 创建过渡窗口并置顶
            cv2.namedWindow('Transition', cv2.WND_PROP_FULLSCREEN)
            cv2.setWindowProperty('Transition', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
            transition_window = pyautogui.getWindowsWithTitle("Transition")[0]
            pyautogui.moveTo(transition_window.left, transition_window.top)
            cv2.imshow('Transition', screenshot)
            hwnd = win32gui.FindWindow(None, "Transition")
            CVRECT = cv2.getWindowImageRect("Transition")
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, screen_width, screen_height,
                                  win32con.SWP_SHOWWINDOW)

            try:
                subprocess.Popen(game_exe)
            except OSError:
                print("打开应用程序错误")
                print(traceback.format_exc())
                sys.exit(2)

            # 进行过渡并在过渡窗口上显示
            for step in range(transition_steps):
                alpha = (step + 1) / transition_steps
                blended_image = cv2.addWeighted(screenshot, 1 - alpha, white_image, alpha, 0)
                cv2.imshow('Transition', blended_image)
                cv2.waitKey(10)

            # 枚举窗口,找到名称包含 windows_title 的窗口
            while True:
                windows = pyautogui.getWindowsWithTitle(windows_title)
                print("枚举窗口")
                if windows:
                    # 将过渡窗口置于非最高层
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, screen_width, screen_height,
                                          win32con.SWP_SHOWWINDOW)
                    time.sleep(wait_time)
                    # 将 windows_title 置于最高层
                    hwnd = win32gui.FindWindow("UnityWndClass", windows_title)
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, screen_width, screen_height,
                                          win32con.SWP_SHOWWINDOW)
                    # genshin_window = windows[0]
                    # pyautogui.moveTo(genshin_window.left, genshin_window.top)
                    # pyautogui.click(genshin_window.left, genshin_window.top)
                    print("星穹铁道 启动!")

                    # 过渡完毕，偷偷删除过渡窗口
                    time.sleep(wait_time)
                    cv2.destroyAllWindows()
                    break

                time.sleep(1)
            sys.exit(0)
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的企业微信截图方法测试
使用最可靠的屏幕截图方式
"""

import time
import win32gui
import win32con
from PIL import ImageGrab
from wxauto.utils.win32 import FindWindow

def activate_window(hwnd):
    """激活窗口"""
    try:
        # 先恢复窗口（如果最小化）
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        
        # 显示窗口
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        time.sleep(0.2)
        
        # 置于前台
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.2)
        
        # 确保在最上层
        win32gui.BringWindowToTop(hwnd)
        time.sleep(0.5)
        
        return True
    except Exception as e:
        print(f"激活窗口失败: {e}")
        return False

def capture_wecom_screen(window_title, region_ratio=(0.05, 0.15, 0.95, 0.75)):
    """使用屏幕截图方式截取企业微信消息区域"""
    try:
        # 查找窗口
        hwnd = FindWindow(name=window_title)
        if not hwnd:
            print(f"未找到窗口: {window_title}")
            return None
        
        print(f"找到窗口: {window_title}, 句柄: {hwnd}")
        
        # 激活窗口
        print("激活窗口...")
        if not activate_window(hwnd):
            return None
        
        # 获取窗口在屏幕上的位置
        rect = win32gui.GetWindowRect(hwnd)
        window_width = rect[2] - rect[0]
        window_height = rect[3] - rect[1]
        
        print(f"窗口位置: {rect}")
        print(f"窗口大小: {window_width}x{window_height}")
        
        # 计算消息区域在屏幕上的绝对坐标
        left_offset = int(window_width * region_ratio[0])
        top_offset = int(window_height * region_ratio[1])
        right_offset = int(window_width * region_ratio[2])
        bottom_offset = int(window_height * region_ratio[3])
        
        screen_left = rect[0] + left_offset
        screen_top = rect[1] + top_offset
        screen_right = rect[0] + right_offset
        screen_bottom = rect[1] + bottom_offset
        
        crop_width = screen_right - screen_left
        crop_height = screen_bottom - screen_top
        
        print(f"截图区域: ({screen_left},{screen_top}) 到 ({screen_right},{screen_bottom})")
        print(f"截图大小: {crop_width}x{crop_height}")
        
        # 使用PIL截取屏幕区域
        print("开始屏幕截图...")
        screen_img = ImageGrab.grab(bbox=(screen_left, screen_top, screen_right, screen_bottom))
        
        print(f"截图成功: {screen_img.size}")
        return screen_img
        
    except Exception as e:
        print(f"截图失败: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    window_title = "无人机AI助教"
    
    print(f"开始测试截图: {window_title}")
    
    # 截图
    image = capture_wecom_screen(window_title)
    
    if image:
        # 保存截图
        image.save("test_capture.png")
        print("截图已保存为: test_capture.png")
        print(f"图像尺寸: {image.size}")
    else:
        print("截图失败")

if __name__ == "__main__":
    main()
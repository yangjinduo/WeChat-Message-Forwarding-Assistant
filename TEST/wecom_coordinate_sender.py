#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信坐标发送消息测试
通过坐标点击和键盘输入的方式发送消息
"""

from wxauto.utils.win32 import FindWindow
import win32gui
import win32api
import win32con
import time

def send_message_by_coordinates(window_title, message):
    """通过坐标点击发送消息到企业微信"""
    print(f"尝试向 {window_title} 发送消息: {message}")
    
    # 查找窗口
    hwnd = FindWindow(name=window_title)
    if not hwnd:
        print(f"❌ 未找到窗口: {window_title}")
        return False
    
    # 获取窗口位置和大小
    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    
    print(f"窗口信息:")
    print(f"  位置: {rect}")
    print(f"  大小: {width} x {height}")
    
    # 激活窗口
    win32gui.SetForegroundWindow(hwnd)
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.2)
    
    # 计算输入区域的大概位置（通常在窗口底部）
    input_area_x = rect[0] + width // 2  # 水平中心
    input_area_y = rect[3] - 80  # 距离底部80像素（输入框通常在这里）
    
    print(f"预估输入区域坐标: ({input_area_x}, {input_area_y})")
    
    # 点击输入区域
    print("点击输入区域...")
    win32api.SetCursorPos((input_area_x, input_area_y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.2)
    
    # 方法1: 直接键盘输入（仅限ASCII字符）
    print("尝试直接键盘输入...")
    try:
        for char in message:
            if char == ' ':
                win32api.keybd_event(win32con.VK_SPACE, 0, 0, 0)
                win32api.keybd_event(win32con.VK_SPACE, 0, win32con.KEYEVENTF_KEYUP, 0)
            elif char.isascii() and char.isalnum():
                if char.isupper():
                    # 大写字母需要按住Shift
                    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
                    win32api.keybd_event(ord(char), 0, 0, 0)
                    win32api.keybd_event(ord(char), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
                else:
                    vk_code = ord(char.upper())
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.02)
        
        # 发送回车
        time.sleep(0.3)
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        print("✅ 通过键盘输入发送成功")
        return True
        
    except Exception as e:
        print(f"键盘输入失败: {e}")
    
    # 方法2: 剪贴板粘贴
    print("尝试剪贴板粘贴...")
    try:
        import win32clipboard
        
        # 放入剪贴板
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(message)
        win32clipboard.CloseClipboard()
        
        # 再次点击确保焦点
        win32api.SetCursorPos((input_area_x, input_area_y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.1)
        
        # Ctrl+V 粘贴
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(0.3)
        
        # 回车发送
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        print("✅ 通过剪贴板粘贴发送成功")
        return True
        
    except Exception as e:
        print(f"剪贴板粘贴失败: {e}")
    
    return False

def test_multiple_positions(window_title):
    """测试多个可能的输入区域位置"""
    print(f"测试多个可能的输入位置...")
    
    hwnd = FindWindow(name=window_title)
    if not hwnd:
        print(f"❌ 未找到窗口: {window_title}")
        return
    
    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    
    # 激活窗口
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    
    # 测试多个可能的输入区域位置
    test_positions = [
        # (描述, x偏移, y偏移)
        ("底部中心", width // 2, height - 50),
        ("底部中心偏上", width // 2, height - 80),
        ("底部中心偏下", width // 2, height - 30),
        ("底部左侧", width // 4, height - 50),
        ("底部右侧", width * 3 // 4, height - 50),
        ("中间偏下", width // 2, height * 3 // 4),
    ]
    
    for desc, x_offset, y_offset in test_positions:
        test_x = rect[0] + x_offset
        test_y = rect[1] + y_offset
        
        print(f"\n测试位置: {desc} ({test_x}, {test_y})")
        
        # 点击测试位置
        win32api.SetCursorPos((test_x, test_y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)
        
        # 输入测试文本
        test_text = f"test{desc}"
        win32api.keybd_event(ord('T'), 0, 0, 0)
        win32api.keybd_event(ord('T'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(ord('E'), 0, 0, 0)
        win32api.keybd_event(ord('E'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(ord('S'), 0, 0, 0)
        win32api.keybd_event(ord('S'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(ord('T'), 0, 0, 0)
        win32api.keybd_event(ord('T'), 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(1)
        
        print(f"已在 {desc} 位置输入测试文本")

def interactive_position_test(window_title):
    """交互式位置测试"""
    print(f"交互式位置测试")
    print("使用方向键调整鼠标位置，空格键测试点击，回车键确认发送")
    
    hwnd = FindWindow(name=window_title)
    if not hwnd:
        print(f"❌ 未找到窗口: {window_title}")
        return
    
    # 激活窗口
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    
    rect = win32gui.GetWindowRect(hwnd)
    current_x = rect[0] + (rect[2] - rect[0]) // 2
    current_y = rect[3] - 80
    
    print(f"初始位置: ({current_x}, {current_y})")
    print("控制说明:")
    print("  方向键: 移动鼠标位置")
    print("  空格键: 测试点击该位置")
    print("  回车键: 在该位置发送测试消息")
    print("  ESC键: 退出")
    
    win32api.SetCursorPos((current_x, current_y))
    
    step = 10  # 每次移动的像素数
    
    while True:
        key = input("\n请按键 (↑↓←→ 移动, 空格测试, 回车发送, q退出): ").strip().lower()
        
        if key == 'q' or key == 'quit':
            break
        elif key == 'w' or key == 'up':
            current_y -= step
        elif key == 's' or key == 'down':
            current_y += step
        elif key == 'a' or key == 'left':
            current_x -= step
        elif key == 'd' or key == 'right':
            current_x += step
        elif key == ' ' or key == 'space':
            print(f"测试点击位置: ({current_x}, {current_y})")
            win32api.SetCursorPos((current_x, current_y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            continue
        elif key == '' or key == 'enter':
            print(f"在位置 ({current_x}, {current_y}) 发送测试消息")
            if send_message_by_coordinates_at_position(hwnd, "测试消息", current_x, current_y):
                print("✅ 发送成功！")
            else:
                print("❌ 发送失败")
            continue
        else:
            print("无效按键")
            continue
        
        # 更新鼠标位置
        win32api.SetCursorPos((current_x, current_y))
        print(f"当前位置: ({current_x}, {current_y})")

def send_message_by_coordinates_at_position(hwnd, message, x, y):
    """在指定坐标位置发送消息"""
    try:
        # 点击指定位置
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.2)
        
        # 使用剪贴板粘贴
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(message)
        win32clipboard.CloseClipboard()
        
        # Ctrl+V
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(0.3)
        
        # 回车
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        return True
    except Exception as e:
        print(f"发送失败: {e}")
        return False

def main():
    """主函数"""
    print("企业微信坐标发送消息测试工具")
    print("=" * 50)
    
    window_title = input("请输入窗口标题 (默认: 无人机AI助教): ").strip() or "无人机AI助教"
    
    print("\n选择测试方式:")
    print("1. 发送测试消息")
    print("2. 测试多个位置")
    print("3. 交互式位置测试")
    
    choice = input("请选择 (1-3): ").strip()
    
    if choice == "1":
        message = input("请输入要发送的消息: ").strip() or "测试消息"
        send_message_by_coordinates(window_title, message)
    elif choice == "2":
        test_multiple_positions(window_title)
    elif choice == "3":
        interactive_position_test(window_title)
    else:
        print("无效选择")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")
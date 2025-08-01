#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信窗口检测脚本
用于检测企业微信的窗口类名和结构
"""

import win32gui
import win32api
import win32con
from wxauto.utils.win32 import GetAllWindows, GetText

def detect_wecom_windows():
    """检测企业微信相关窗口"""
    print("正在检测企业微信窗口...")
    print("=" * 50)
    
    all_windows = GetAllWindows()
    wecom_windows = []
    
    # 查找可能的企业微信窗口
    for hwnd, class_name, window_title in all_windows:
        # 检查窗口标题或类名中是否包含企业微信相关关键词
        if any(keyword in window_title.lower() for keyword in ['企业微信', 'wecom', 'work', 'wework']) or \
           any(keyword in class_name.lower() for keyword in ['wecom', 'work', 'wework']):
            wecom_windows.append((hwnd, class_name, window_title))
            print(f"发现可能的企业微信窗口:")
            print(f"  句柄: {hwnd}")
            print(f"  类名: {class_name}")
            print(f"  标题: {window_title}")
            print("-" * 30)
    
    # 如果没有找到明确的企业微信窗口，列出所有可能相关的窗口
    if not wecom_windows:
        print("未找到明确的企业微信窗口，列出所有可能相关的窗口:")
        for hwnd, class_name, window_title in all_windows:
            if window_title and (
                '微信' in window_title or 
                'WeChat' in window_title or
                'WeCom' in window_title or
                'Work' in window_title
            ):
                print(f"  句柄: {hwnd}, 类名: {class_name}, 标题: {window_title}")
    
    return wecom_windows

def analyze_window_structure(hwnd):
    """分析窗口结构"""
    print(f"\n分析窗口结构 (句柄: {hwnd}):")
    print("=" * 50)
    
    def enum_child_windows(parent_hwnd, level=0):
        """递归枚举子窗口"""
        indent = "  " * level
        
        def enum_callback(child_hwnd, param):
            try:
                class_name = win32gui.GetClassName(child_hwnd)
                window_text = GetText(child_hwnd)
                rect = win32gui.GetWindowRect(child_hwnd)
                
                print(f"{indent}├─ 句柄: {child_hwnd}")
                print(f"{indent}   类名: {class_name}")
                print(f"{indent}   文本: {window_text[:50]}{'...' if len(window_text) > 50 else ''}")
                print(f"{indent}   位置: {rect}")
                
                # 递归枚举子窗口（限制深度避免过深）
                if level < 3:
                    enum_child_windows(child_hwnd, level + 1)
                    
            except Exception as e:
                print(f"{indent}   错误: {e}")
            
            return True
        
        try:
            win32gui.EnumChildWindows(parent_hwnd, enum_callback, None)
        except Exception as e:
            print(f"{indent}枚举子窗口失败: {e}")
    
    try:
        # 获取主窗口信息
        class_name = win32gui.GetClassName(hwnd)
        window_text = GetText(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        
        print(f"主窗口信息:")
        print(f"  类名: {class_name}")
        print(f"  标题: {window_text}")
        print(f"  位置: {rect}")
        print(f"\n子窗口结构:")
        
        # 枚举子窗口
        enum_child_windows(hwnd)
        
    except Exception as e:
        print(f"分析窗口结构失败: {e}")

if __name__ == "__main__":
    print("企业微信窗口检测工具")
    print("请确保企业微信已经启动并登录")
    print()
    
    # 检测企业微信窗口
    wecom_windows = detect_wecom_windows()
    
    if wecom_windows:
        print(f"\n找到 {len(wecom_windows)} 个可能的企业微信窗口")
        
        # 分析第一个窗口的结构
        if wecom_windows:
            hwnd, class_name, window_title = wecom_windows[0]
            analyze_window_structure(hwnd)
    else:
        print("\n未找到企业微信窗口")
        print("请检查:")
        print("1. 企业微信是否已启动")
        print("2. 企业微信是否已登录")
        print("3. 企业微信版本是否支持")
    
    input("\n按回车键退出...")
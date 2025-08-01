#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试企业微信独立聊天窗口操作
验证是否可以通过UIAutomation操作企业微信的独立聊天窗口
"""

import win32gui
from wxauto.utils.win32 import FindWindow, GetAllWindows, GetText
from wxauto import uiautomation as uia
import time

def find_wecom_chat_windows():
    """查找企业微信聊天窗口"""
    print("查找企业微信独立聊天窗口...")
    print("=" * 60)
    
    all_windows = GetAllWindows()
    chat_windows = []
    
    # 查找可能的企业微信聊天窗口
    possible_classes = ['ChatWnd', 'WeComChatWnd', 'WorkWeChatChatWnd', 'WeWorkChatWnd']
    
    for hwnd, class_name, window_title in all_windows:
        # 检查是否是企业微信相关的聊天窗口
        if any(chat_class in class_name for chat_class in possible_classes) or \
           (window_title and window_title not in ['', '企业微信'] and \
            any(parent_class in win32gui.GetClassName(win32gui.GetParent(hwnd) or hwnd) \
                for parent_class in ['WeWorkWindow', 'WeCom'])):
            chat_windows.append((hwnd, class_name, window_title))
            print(f"发现聊天窗口:")
            print(f"  句柄: {hwnd}")
            print(f"  类名: {class_name}")
            print(f"  标题: {window_title}")
            print("-" * 30)
    
    return chat_windows

def test_chat_window_uia(hwnd, window_title):
    """测试聊天窗口的UIAutomation访问"""
    print(f"\n测试聊天窗口UIAutomation访问: {window_title}")
    print("=" * 60)
    
    try:
        control = uia.ControlFromHandle(hwnd)
        if not control:
            print("✗ 无法创建UIAutomation控件")
            return False
            
        print("✓ UIAutomation控件创建成功")
        
        # 测试基本属性
        try:
            print(f"  名称: {control.Name}")
            print(f"  类名: {control.ClassName}")
            print(f"  类型: {control.ControlTypeName}")
            print(f"  边界: {control.BoundingRectangle}")
            print(f"  可见: {control.IsEnabled}")
        except Exception as e:
            print(f"✗ 基本属性访问失败: {e}")
            
        # 测试子控件访问
        try:
            children = control.GetChildren()
            print(f"✓ 子控件数量: {len(children)}")
            
            if len(children) > 0:
                print("  子控件详情:")
                for i, child in enumerate(children[:5]):  # 只显示前5个
                    try:
                        print(f"    [{i}] {child.ClassName} - {child.Name} - {child.ControlTypeName}")
                        print(f"        边界: {child.BoundingRectangle}")
                        
                        # 进一步检查是否包含聊天相关的控件
                        grandchildren = child.GetChildren()
                        if len(grandchildren) > 0:
                            print(f"        子元素: {len(grandchildren)}个")
                            for j, grandchild in enumerate(grandchildren[:3]):
                                print(f"          [{j}] {grandchild.ClassName} - {grandchild.Name}")
                                
                    except Exception as e:
                        print(f"    [{i}] 访问失败: {e}")
                        
                return True
            else:
                print("⚠️  没有找到子控件")
                return False
                
        except Exception as e:
            print(f"✗ 子控件访问失败: {e}")
            return False
            
    except Exception as e:
        print(f"✗ UIAutomation访问失败: {e}")
        return False

def test_send_message_simulation(hwnd, window_title):
    """测试模拟发送消息"""
    print(f"\n测试消息发送模拟: {window_title}")
    print("=" * 60)
    
    try:
        control = uia.ControlFromHandle(hwnd)
        
        # 查找输入框
        def find_input_control(parent_control, depth=0):
            if depth > 3:  # 限制搜索深度
                return None
                
            try:
                children = parent_control.GetChildren()
                for child in children:
                    # 检查是否是编辑框或可输入的控件
                    if (child.ControlTypeName in ['EditControl', 'DocumentControl'] or 
                        'Edit' in child.ClassName or 'Input' in child.ClassName):
                        print(f"找到可能的输入控件: {child.ClassName} - {child.Name}")
                        return child
                    
                    # 递归搜索
                    result = find_input_control(child, depth + 1)
                    if result:
                        return result
            except:
                pass
            return None
        
        input_control = find_input_control(control)
        if input_control:
            print("✓ 找到输入控件")
            print("  注意: 实际发送测试需要手动确认，这里只是检测")
            return True
        else:
            print("✗ 未找到输入控件")
            return False
            
    except Exception as e:
        print(f"✗ 输入控件搜索失败: {e}")
        return False

def interactive_test():
    """交互式测试"""
    print("\n交互式测试模式")
    print("=" * 60)
    print("请在企业微信中打开一个聊天窗口（如：无人机AI助教）")
    print("然后按回车继续测试...")
    input()
    
    # 重新查找窗口
    chat_windows = find_wecom_chat_windows()
    
    if not chat_windows:
        print("未找到聊天窗口，请确保已打开企业微信独立聊天窗口")
        return
    
    print(f"找到 {len(chat_windows)} 个聊天窗口")
    
    for i, (hwnd, class_name, window_title) in enumerate(chat_windows):
        print(f"\n[{i+1}] 测试窗口: {window_title}")
        
        # 基本UIAutomation测试
        uia_success = test_chat_window_uia(hwnd, window_title)
        
        # 发送消息模拟测试
        if uia_success:
            send_success = test_send_message_simulation(hwnd, window_title)
            
            if send_success:
                print(f"✅ 窗口 '{window_title}' 可以进行UIAutomation操作")
            else:
                print(f"⚠️  窗口 '{window_title}' 部分功能可用")
        else:
            print(f"❌ 窗口 '{window_title}' 无法进行UIAutomation操作")
        
        print()

def main():
    """主函数"""
    print("企业微信独立聊天窗口测试工具")
    print("请确保以管理员权限运行此脚本")
    print()
    
    # 首先查找现有的聊天窗口
    chat_windows = find_wecom_chat_windows()
    
    if chat_windows:
        print(f"找到 {len(chat_windows)} 个可能的聊天窗口")
        
        # 测试每个窗口
        for hwnd, class_name, window_title in chat_windows:
            test_chat_window_uia(hwnd, window_title)
    else:
        print("未找到现有的聊天窗口")
    
    # 交互式测试
    interactive_test()
    
    print("\n测试完成！")
    print("如果独立聊天窗口可以进行UIAutomation操作，")
    print("那么消息转发助手就可以正常工作。")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")

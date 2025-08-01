#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接测试已找到的企业微信聊天窗口
"""

from wxauto.utils.win32 import FindWindow
from wxauto import uiautomation as uia
import win32gui

def test_wecom_chat_window():
    """测试企业微信聊天窗口"""
    print("测试企业微信聊天窗口 - 无人机AI助教")
    print("=" * 80)
    
    # 直接查找我们知道存在的窗口
    hwnd = FindWindow(name="无人机AI助教")
    if not hwnd:
        print("❌ 未找到聊天窗口")
        print("请确保'无人机AI助教'聊天窗口仍然打开")
        return False
    
    class_name = win32gui.GetClassName(hwnd)
    print(f"✅ 找到聊天窗口:")
    print(f"  句柄: {hwnd}")
    print(f"  类名: {class_name}")
    print(f"  标题: 无人机AI助教")
    
    # 检查窗口状态
    is_visible = win32gui.IsWindowVisible(hwnd)
    is_enabled = win32gui.IsWindowEnabled(hwnd)
    rect = win32gui.GetWindowRect(hwnd)
    print(f"  可见: {is_visible}")
    print(f"  启用: {is_enabled}")
    print(f"  位置: {rect}")
    
    print(f"\n开始UIAutomation测试...")
    print("-" * 50)
    
    try:
        # 创建UIAutomation控件
        control = uia.ControlFromHandle(hwnd)
        if not control:
            print("❌ 无法创建UIAutomation控件")
            return False
        
        print("✅ UIAutomation控件创建成功")
        
        # 测试基本属性
        try:
            print(f"  控件名称: {control.Name}")
            print(f"  控件类名: {control.ClassName}")
            print(f"  控件类型: {control.ControlTypeName}")
            bounds = control.BoundingRectangle
            print(f"  控件边界: {bounds}")
            print(f"  控件启用: {control.IsEnabled}")
            
            # 检查边界是否有效
            if bounds.width() > 0 and bounds.height() > 0:
                print("✅ 控件边界正常")
            else:
                print("⚠️  控件边界异常")
                
        except Exception as e:
            print(f"❌ 基本属性测试失败: {e}")
            return False
        
        # 测试子控件
        print(f"\n测试子控件...")
        try:
            children = control.GetChildren()
            print(f"✅ 子控件数量: {len(children)}")
            
            if len(children) == 0:
                print("⚠️  没有找到子控件")
                print("这可能意味着:")
                print("  1. 窗口使用了自定义渲染")
                print("  2. 需要特殊的访问方法")
                print("  3. 权限不足")
                return False
            else:
                print("✅ 找到子控件，分析结构:")
                
                for i, child in enumerate(children):
                    try:
                        child_name = child.Name
                        child_class = child.ClassName
                        child_type = child.ControlTypeName
                        child_bounds = child.BoundingRectangle
                        
                        print(f"  子控件[{i}]:")
                        print(f"    名称: {child_name}")
                        print(f"    类名: {child_class}")
                        print(f"    类型: {child_type}")
                        print(f"    边界: {child_bounds}")
                        
                        # 查找可能的输入框
                        if child_type in ['EditControl', 'DocumentControl'] or 'Edit' in child_class:
                            print(f"    🎯 可能的输入控件!")
                        
                        # 递归查看子控件的子控件
                        try:
                            grandchildren = child.GetChildren()
                            if len(grandchildren) > 0:
                                print(f"    包含 {len(grandchildren)} 个子元素")
                                for j, gc in enumerate(grandchildren[:3]):  # 最多显示3个
                                    gc_type = gc.ControlTypeName
                                    gc_class = gc.ClassName
                                    gc_name = gc.Name
                                    print(f"      [{j}] {gc_type} - {gc_class} - {gc_name}")
                                    
                                    # 查找输入相关控件
                                    if gc_type in ['EditControl', 'DocumentControl'] or 'Edit' in gc_class:
                                        print(f"        🎯 发现输入控件!")
                        except:
                            pass
                        
                        print()
                        
                    except Exception as e:
                        print(f"  子控件[{i}] 分析失败: {e}")
                
                return True
                
        except Exception as e:
            print(f"❌ 子控件测试失败: {e}")
            return False
            
    except Exception as e:
        print(f"❌ UIAutomation测试失败: {e}")
        return False

def test_message_sending():
    """测试消息发送功能"""
    print("\n测试消息发送功能")
    print("=" * 50)
    
    hwnd = FindWindow(name="无人机AI助教")
    if not hwnd:
        print("❌ 未找到聊天窗口")
        return False
    
    try:
        control = uia.ControlFromHandle(hwnd)
        
        # 查找输入框的递归函数
        def find_input_controls(parent, depth=0, max_depth=5):
            input_controls = []
            if depth > max_depth:
                return input_controls
            
            try:
                children = parent.GetChildren()
                for child in children:
                    child_type = child.ControlTypeName
                    child_class = child.ClassName
                    
                    # 检查是否是输入相关控件
                    if (child_type in ['EditControl', 'DocumentControl', 'TextControl'] or
                        'Edit' in child_class or 'Input' in child_class or 'Text' in child_class):
                        input_controls.append((child, child_type, child_class, child.Name))
                    
                    # 递归搜索
                    input_controls.extend(find_input_controls(child, depth + 1, max_depth))
                    
            except:
                pass
            
            return input_controls
        
        # 查找所有可能的输入控件
        input_controls = find_input_controls(control)
        
        print(f"找到 {len(input_controls)} 个可能的输入控件:")
        for i, (ctrl, ctrl_type, ctrl_class, ctrl_name) in enumerate(input_controls):
            print(f"  输入控件[{i}]: {ctrl_type} - {ctrl_class} - {ctrl_name}")
            print(f"    边界: {ctrl.BoundingRectangle}")
            print(f"    启用: {ctrl.IsEnabled}")
        
        if len(input_controls) > 0:
            print("✅ 找到输入控件，理论上可以发送消息")
            return True
        else:
            print("❌ 未找到输入控件")
            return False
            
    except Exception as e:
        print(f"❌ 消息发送测试失败: {e}")
        return False

def main():
    """主函数"""
    print("企业微信聊天窗口功能测试")
    print("基于已找到的窗口: WwStandaloneConversationWnd")
    print()
    
    # 测试基本UIAutomation功能
    basic_success = test_wecom_chat_window()
    
    if basic_success:
        # 测试消息发送功能
        send_success = test_message_sending()
        
        print("\n" + "=" * 80)
        print("📋 测试结果总结")
        print("=" * 80)
        
        if send_success:
            print("🎉 完全成功!")
            print("✅ 企业微信独立聊天窗口完全支持UIAutomation操作")
            print("✅ 可以找到输入控件")
            print("✅ 消息转发助手应该可以正常工作")
            print()
            print("🚀 下一步: 可以运行主程序 wechat_message_forwarder.py")
        else:
            print("⚠️  部分成功")
            print("✅ 可以访问聊天窗口")
            print("❌ 无法找到输入控件")
            print("⚠️  消息转发可能需要使用其他方法")
    else:
        print("\n" + "=" * 80)
        print("❌ 测试失败")
        print("❌ 企业微信聊天窗口无法进行UIAutomation操作")
        print("💡 建议使用图像识别或坐标点击的方式")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")